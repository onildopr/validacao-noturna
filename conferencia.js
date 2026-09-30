// Pull periódico do banco (rede de segurança caso o realtime caia)
const SYNC_INTERVAL_MS = 15000;
// Espera após a última bipagem antes de enviar ao banco
const CLOUD_SAVE_DEBOUNCE_MS = 1200;

// Prefixo de chaves no localStorage (separa por operação e por dia)
const STORAGE_KEY_PREFIX = 'conferencia.routes.v3';

const ConferenciaApp = {
  routes: new Map(),     // routeId -> routeObject (somente do dia selecionado)
  currentRouteId: null,
  viaCsv: false,
  operationCode: null, // ex: ERD1
  deviceId: null,
  cloudEnabled: true,
  cloudDirty: false,
  cloudSaving: false,
  cloudLastSaveAt: 0,
  cloudSaveTimer: null,
  cloudMutationSeq: 0,         // incrementa a cada mudança local (detecta bipagens durante o save)
  syncTimer: null,
  lastLocalMutationAt: 0,
  cloudPullGraceMs: 2000,
  workDay: null,               // YYYY-MM-DD
  lastEvents: [],              // log simples de bipagens (últimos eventos)
  deletedRoutes: new Map(),    // routeId -> ts (epoch ms)
  revivedRoutes: new Map(),    // routeId -> ts (epoch ms) (desfaz exclusão)

  // ===== Realtime sync (Supabase) =====
  rtChannel: null,
  rtBound: { op: null, day: null },
  lastRemoteUpdatedAt: null,
  lastPushedSnapshotHash: '',
  lastSavedLocalSnapshotHash: '',

  // =======================
  // Carretas (Placa -> Rotas QR)
  // =======================
  carretas: {
    currentPlateKey: null,
    plates: new Map(),        // plateKey -> {raw, license_plate, carrier_name, vehicle_type_description, routes:Set(routeKey), tsFirst, tsLast}
    routeToPlate: new Map(),  // routeKey -> plateKey
    routesRaw: new Map(),     // routeKey -> rawText
    routesJson: new Map(),    // routeKey -> jsonText (para export)
    routesTs: new Map(),      // routeKey -> tsScan
  },

  // ===== Lock da UI de seleção de rotas =====
  routeUiLockUntil: 0,
  isRouteDropdownOpen: false,
  lastRoutesSignature: '',

  // Escapa texto antes de inserir em HTML (IDs, clusters, QRs, nomes vindos de fora)
  escHtml(s) {
    return String(s ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  },

  lockRouteUi(ms = 2500) {
    this.routeUiLockUntil = Date.now() + ms;
  },

  isRouteUiLocked() {
    return this.isRouteDropdownOpen || Date.now() < (this.routeUiLockUntil || 0);
  },

  // =======================
  // Util data/strings
  // =======================

  // Normaliza texto de cluster/assignment para comparação (sem depender de acentos/lixo do leitor)
  normalizeCluster(v) {
    return String(v ?? '')
      .trim()
      .toUpperCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^\w\-]+/g, '');
  },

  pad2(n) { return String(n).padStart(2, '0'); },

  todayLocalISO() {
    const fmt = new Intl.DateTimeFormat('en-CA', {
      timeZone: 'America/Porto_Velho',
      year: 'numeric',
      month: '2-digit',
      day: '2-digit'
    });
    return fmt.format(new Date());
  },

  monthKeyFromDay(dayISO) {
    return String(dayISO || '').slice(0, 7);
  },

  normalizeCaretKey(k) {
    const key = String(k || '').trim().toLowerCase();
    const noAcc = key.normalize('NFD').replace(/[\u0300-\u036f]/g, '');

    if (noAcc === 'assignment' || noAcc === 'assigment' || noAcc === 'asssignment') return 'assignment';
    if (noAcc === 'license_plate') return 'license_plate';
    if (noAcc === 'carrier_name') return 'carrier_name';
    if (noAcc === 'carrier_id') return 'carrier_id';
    if (noAcc === 'vehicle_type_description') return 'vehicle_type_description';
    if (noAcc === 'container_id') return 'container_id';
    if (noAcc === 'facility_id') return 'facility_id';
    if (noAcc === 'id') return 'id';

    return noAcc;
  },

  parseCaretKV(raw) {
    const cleaned = String(raw || '').replace(/\r/g, '').trim();
    const first = cleaned.split('\n')[0].trim();

    const kv = {};
    const tokens = first.split(',').map(t => t.trim()).filter(Boolean);

    for (const tok of tokens) {
      const m = tok.match(/^(\^?)([^\\^]+?)\^Ç\^?(.+?)\^?$/);
      if (!m) continue;

      const rawKey = m[2];
      let val = m[3];
      const key = this.normalizeCaretKey(rawKey);

      val = String(val)
        .replace(/^\^+|\^+$/g, '')
        .replace(/[{}]/g, '')
        .trim();

      kv[key] = val;
    }

    return kv;
  },

  storageKeyForDay(dayISO) {
    const op = this.getOperationCode() || 'NOOP';
    return `${STORAGE_KEY_PREFIX}.${op}.${dayISO}`;
  },

  // =======================
  // Operação (ERD1, ERD2...) e Device
  // =======================
  getDeviceId() {
    if (this.deviceId) return this.deviceId;
    const k = 'conf_device_id.v1';
    let v = localStorage.getItem(k);
    if (!v) {
      v = (crypto.randomUUID && crypto.randomUUID()) || `${Date.now()}-${Math.random().toString(16).slice(2)}`;
      localStorage.setItem(k, v);
    }
    this.deviceId = v;
    return v;
  },

  setOperationCode(code) {
    const norm = String(code || '').trim().toUpperCase();
    if (!norm) return;
    localStorage.setItem('conf_operation_code.v1', norm);
    this.operationCode = norm;
    const $badge = $('#op-badge');
    if ($badge.length) $badge.text(norm);
  },

  getOperationCode() {
    if (this.operationCode) return this.operationCode;
    const v = localStorage.getItem('conf_operation_code.v1');
    this.operationCode = v ? String(v).toUpperCase() : null;
    return this.operationCode;
  },

  // =======================
  // Supabase: client ÚNICO
  // =======================
  getSb() {
    if (window.__confSbClient) return window.__confSbClient;
    if (window.sbClient) {
      window.__confSbClient = window.sbClient;
      return window.__confSbClient;
    }
    if (!window.supabase || !window.SB_URL || !window.SB_ANON) return null;

    const url = window.SB_URL;
    const key = window.SB_ANON;

    window.__confSbClient = window.supabase.createClient(url, key, {
      auth: {
        persistSession: false,
        autoRefreshToken: false,
        detectSessionInUrl: false,
      },
      global: {
        headers: {
          apikey: key,
          Authorization: `Bearer ${key}`,
        },
      },
      realtime: {
        params: {
          eventsPerSecond: 2,
        },
      },
    });

    return window.__confSbClient;
  },

  // =======================
  // Admin
  // =======================
  async adminUpsertOperation(code, name, active = true) {
    const sb = this.getSb();
    if (!sb) throw new Error('Supabase client não encontrado.');

    const op = {
      code: String(code || '').trim().toUpperCase(),
      name: String(name || '').trim() || null,
      active: !!active,
    };

    if (!/^[A-Z]{3}\d$/.test(op.code)) {
      throw new Error('Código inválido. Use 3 letras e 1 número (ex.: ERD1).');
    }

    const { error } = await sb.from('operations').upsert(op, { onConflict: 'code' });
    if (error) throw error;
  },

  async adminLoadOperations(includeInactive = true) {
    const sb = this.getSb();
    if (!sb) throw new Error('Supabase client não encontrado.');
    let q = sb.from('operations').select('code,name,active,created_at').order('code', { ascending: true });
    if (!includeInactive) q = q.eq('active', true);
    const { data, error } = await q;
    if (error) throw error;
    return data || [];
  },

  // =======================
  // Realtime
  // =======================
  async stopRealtimeSync() {
    try {
      const sb = this.getSb();
      if (sb && this.rtChannel) {
        await sb.removeChannel(this.rtChannel);
      }
    } catch (e) {
      console.warn('Falha ao parar realtime:', e);
    } finally {
      this.rtChannel = null;
      this.rtBound = { op: null, day: null };
    }
  },

  async startRealtimeSync(dayISO) {
    const sb = this.getSb();
    const op = this.getOperationCode();
    if (!sb || !op || !dayISO) return;

    if (this.rtChannel && this.rtBound && this.rtBound.op === op && this.rtBound.day === dayISO) return;

    await this.stopRealtimeSync();

    this.rtChannel = sb
      .channel('routes_state_live')
      .on(
        'postgres_changes',
        { event: '*', schema: 'public', table: 'routes_state', filter: `operation_code=eq.${op}` },
        async (payload) => {
          try {
            const row = payload && payload.new;
            if (!row || !row.data) return;

            if (row.operation_code !== op) return;
            if (row.day !== dayISO) return;
            if (row.device_id && row.device_id === this.getDeviceId()) return;

            // A mesclagem é por união, então é seguro aplicar mesmo com mudanças locais pendentes
            this.lastRemoteUpdatedAt = row.updated_at || this.lastRemoteUpdatedAt;
            this.mergeSnapshotIntoLocal(row.data);
            this.saveToStorage(dayISO, { syncCloud: false });

            // Se o local tem algo a mais que o remoto, reenvia a união
            const remoteHash = this.computeSnapshotHash(row.data);
            const localHash = this.computeSnapshotHash(this.buildDaySnapshotObject());
            if (localHash !== remoteHash) this.markCloudDirty(dayISO, op);
            else if (!this.cloudDirty) this.lastPushedSnapshotHash = remoteHash;

            this.refreshAfterRemoteMerge();
            this.setStatus(`Realtime • atualizado • ${op} • ${dayISO}`, "info");
          } catch (e) {
            console.warn('Falha ao aplicar realtime payload:', e);
          }
        }
      )
      .subscribe((status) => {
        if (status === 'SUBSCRIBED') {
          this.setStatus(`Realtime ON • ${op} • ${dayISO}`, 'success');
        } else if (status === 'CHANNEL_ERROR' || status === 'TIMED_OUT') {
          this.setStatus(`Realtime instável • ${op} • ${dayISO}`, 'warning');
        } else if (status === 'CLOSED') {
          // O pull periódico (startPeriodicSync) cobre o período sem realtime
          this.setStatus(`Realtime CLOSED • ${op} • ${dayISO}`, 'warning');
        }
        try { console.log('[Realtime]', status, { op, dayISO }); } catch {}
      });

    this.rtBound = { op, day: dayISO };
  },

  buildDaySnapshotObject() {
    const obj = {};
    for (const [routeId, r] of this.routes.entries()) {
      obj[routeId] = this.serializeRoute(r);
    }
    obj.__meta = {
      deletedRoutes: Object.fromEntries(this.deletedRoutes || new Map()),
      revivedRoutes: Object.fromEntries(this.revivedRoutes || new Map()),
      carretas: this.serializeCarretas()
    };
    return obj;
  },

  // =======================
  // Carretas: persistência (vai junto no snapshot do dia)
  // =======================
  resetCarretas() {
    this.carretas = {
      currentPlateKey: null,
      plates: new Map(),
      routeToPlate: new Map(),
      routesRaw: new Map(),
      routesJson: new Map(),
      routesTs: new Map(),
    };
  },

  serializeCarretas() {
    const c = this.carretas;
    const plates = {};
    for (const [k, p] of c.plates.entries()) {
      plates[k] = Object.assign({}, p, { routes: Array.from(p.routes || []) });
    }
    return {
      plates,
      routesRaw: Object.fromEntries(c.routesRaw),
      routesJson: Object.fromEntries(c.routesJson),
      routesTs: Object.fromEntries(c.routesTs),
    };
  },

  mergeCarretas(ser) {
    if (!ser || typeof ser !== 'object') return;
    const c = this.carretas;

    for (const [k, p] of Object.entries(ser.plates || {})) {
      const remoteRoutes = Array.isArray(p.routes) ? p.routes : [];
      const cur = c.plates.get(k);
      if (!cur) {
        c.plates.set(k, Object.assign({}, p, { routes: new Set(remoteRoutes) }));
      } else {
        remoteRoutes.forEach(rk => cur.routes.add(rk));
        if (Number(p.tsLast || 0) > Number(cur.tsLast || 0)) {
          for (const f of ['raw', 'jsonText', 'carrier_name', 'vehicle_type_description', 'tsLast', 'tsScan']) {
            if (p[f]) cur[f] = p[f];
          }
        }
      }
      remoteRoutes.forEach(rk => c.routeToPlate.set(rk, k));
    }

    // Desempate determinístico para que todos os aparelhos terminem com o mesmo valor
    const pick = (map, rk, v) => {
      const cur = map.get(rk);
      if (cur == null || String(v) > String(cur)) map.set(rk, v);
    };
    for (const [rk, v] of Object.entries(ser.routesRaw || {})) pick(c.routesRaw, rk, v);
    for (const [rk, v] of Object.entries(ser.routesJson || {})) pick(c.routesJson, rk, v);
    for (const [rk, v] of Object.entries(ser.routesTs || {})) {
      if (Number(v) > Number(c.routesTs.get(rk) || 0)) c.routesTs.set(rk, Number(v));
    }

    // "Excluir bipagem desta placa" precisa vencer a união: remove rotas bipadas antes da limpeza
    for (const [k, p] of Object.entries(ser.plates || {})) {
      const cur = c.plates.get(k);
      if (!cur) continue;
      cur.clearedAt = Math.max(Number(cur.clearedAt || 0), Number(p.clearedAt || 0));
      if (!cur.clearedAt) continue;
      for (const rk of Array.from(cur.routes)) {
        if (Number(c.routesTs.get(rk) || 0) <= cur.clearedAt) {
          cur.routes.delete(rk);
          if (c.routeToPlate.get(rk) === k) c.routeToPlate.delete(rk);
        }
      }
    }
  },

  // Hash independente de ordem (Sets/Maps viram arrays/objetos em ordem de inserção,
  // que muda entre aparelhos). Sem isso dois aparelhos ficariam reenviando o mesmo estado.
  computeSnapshotHash(snapshotObj) {
    const canon = (v) => {
      if (Array.isArray(v)) return v.map(canon).sort((a, b) => (JSON.stringify(a) < JSON.stringify(b) ? -1 : 1));
      if (v && typeof v === 'object') {
        const out = {};
        for (const k of Object.keys(v).sort()) out[k] = canon(v[k]);
        return out;
      }
      return v;
    };
    try {
      return JSON.stringify(canon(snapshotObj || {}));
    } catch {
      return `${Date.now()}`;
    }
  },

  async supaLoadDaySnapshot(operationCode, dayISO) {
    const sb = this.getSb();
    if (!sb) throw new Error('Supabase client não encontrado (window.sbClient).');
    const { data, error } = await sb
      .from('routes_state')
      .select('data,updated_at,device_id')
      .eq('operation_code', operationCode)
      .eq('day', dayISO)
      .maybeSingle();
    if (error) throw error;
    if (!data) return null;
    return data;
  },

  async supaSaveDaySnapshot(operationCode, dayISO, snapshotObj) {
    const sb = this.getSb();
    if (!sb) throw new Error('Supabase client não encontrado (window.sbClient).');
    const payload = {
      operation_code: operationCode,
      day: dayISO,
      data: snapshotObj,
      updated_at: new Date().toISOString(),
      device_id: this.getDeviceId()
    };
    const { error } = await sb
      .from('routes_state')
      .upsert(payload, { onConflict: 'operation_code,day' });
    if (error) throw error;
  },

  markCloudDirty(targetDay = this.workDay || this.todayLocalISO(), targetOp = this.getOperationCode()) {
    this.cloudDirty = true;
    this.cloudMutationSeq++;
    this.lastLocalMutationAt = Date.now();

    this.pendingCloudSave = {
      day: targetDay,
      op: targetOp
    };

    this.scheduleCloudSave(CLOUD_SAVE_DEBOUNCE_MS);
  },

  scheduleCloudSave(delayMs) {
    if (this.cloudSaveTimer) clearTimeout(this.cloudSaveTimer);
    this.cloudSaveTimer = setTimeout(() => {
      this.cloudSaveTimer = null;
      const pending = this.pendingCloudSave || {};
      this.flushCloudSave(pending.op, pending.day).catch(e => {
        console.warn('Falha ao salvar no Supabase (vai tentar de novo):', e);
        this.setStatus('Falha ao salvar no banco (Supabase). Mantido no cache local, tentando de novo...', 'warning');
        this.scheduleCloudSave(5000);
      });
    }, delayMs);
  },

  async flushCloudSave(forceOp, forceDay) {
    if (!this.cloudEnabled) return;
    if (!this.cloudDirty) return;
    if (this.cloudSaving) {
      // Já tem um save em andamento: tenta de novo quando ele terminar
      this.scheduleCloudSave(500);
      return;
    }

    const op = forceOp || this.getOperationCode();
    const day = forceDay || this.workDay || this.todayLocalISO();
    if (!op || !day) return;

    // Só envia se o estado em memória ainda for desse dia/operação
    if (op !== this.getOperationCode() || day !== (this.workDay || this.todayLocalISO())) return;

    this.cloudSaving = true;
    const seqAtStart = this.cloudMutationSeq;
    try {
      // Ler → mesclar → gravar: incorpora o que outros aparelhos salvaram antes de sobrescrever
      const remote = await this.supaLoadDaySnapshot(op, day);
      if (remote && remote.data) {
        this.mergeSnapshotIntoLocal(remote.data);
        this.saveToStorage(day, { syncCloud: false });
      }

      const snapshot = this.buildDaySnapshotObject();
      const snapshotHash = this.computeSnapshotHash(snapshot);
      const remoteHash = remote && remote.data ? this.computeSnapshotHash(remote.data) : '';

      if (snapshotHash !== remoteHash) {
        await this.supaSaveDaySnapshot(op, day, snapshot);
        this.lastRemoteUpdatedAt = new Date().toISOString();
      }

      this.cloudLastSaveAt = Date.now();
      this.lastPushedSnapshotHash = snapshotHash;

      // Se alguém bipou durante o save, continua sujo e agenda outro envio
      if (this.cloudMutationSeq !== seqAtStart) {
        this.scheduleCloudSave(300);
      } else {
        this.cloudDirty = false;
        this.pendingCloudSave = null;
        this.dirty = false;
        $('#dirty-flag').addClass('d-none');
        this.setStatus(`Salvo no banco • ${op} • ${day}`, 'success');
      }

      if (remote && remote.data) this.refreshAfterRemoteMerge();
    } finally {
      this.cloudSaving = false;
    }
  },

  // Re-renderiza a tela depois de mesclar dados vindos do banco
  refreshAfterRemoteMerge() {
    if (this._rtUiTimer) clearTimeout(this._rtUiTimer);
    this._rtUiTimer = setTimeout(() => {
      try {
        const keepRouteId = this.currentRouteId;
        if (!this.isRouteUiLocked()) {
          this.renderRoutesSelects();
        } else if (keepRouteId && this.routes.has(String(keepRouteId))) {
          this.currentRouteId = String(keepRouteId);
        }
        this.refreshUIFromCurrent();
        this.renderAcompanhamento();
        if (!$('#carreta-interface').hasClass('d-none')) this.renderCarretaUI();
      } catch (e) {
        console.warn('Falha ao renderizar após sync:', e);
      }
    }, 250);
  },

  // Converte duplicados antigos (que viraram array por bug de merge) em número
  dupCount(v) {
    if (Array.isArray(v)) return v.reduce((m, x) => Math.max(m, Number(x) || 0), 0);
    return Number(v) || 0;
  },

  mergeSnapshotIntoLocal(snapshotObj) {
    if (!snapshotObj || typeof snapshotObj !== 'object') return;
    if (!this.deletedRoutes) this.deletedRoutes = new Map();
    if (!this.revivedRoutes) this.revivedRoutes = new Map();

    const remoteDel = snapshotObj.__meta?.deletedRoutes || {};
    for (const [rid, ts] of Object.entries(remoteDel)) {
      const id = String(rid);
      const t = Number(ts || 0);
      if (t > Number(this.deletedRoutes.get(id) || 0)) this.deletedRoutes.set(id, t);
    }

    const remoteRev = snapshotObj.__meta?.revivedRoutes || {};
    for (const [rid, ts] of Object.entries(remoteRev)) {
      const id = String(rid);
      const t = Number(ts || 0);
      if (t > Number(this.revivedRoutes.get(id) || 0)) this.revivedRoutes.set(id, t);
    }

    // Exclusão vs. reimportação: vence o mais recente
    for (const [id, delTs] of Array.from(this.deletedRoutes.entries())) {
      if (Number(this.revivedRoutes.get(id) || 0) > Number(delTs)) this.deletedRoutes.delete(id);
    }

    for (const rid of this.deletedRoutes.keys()) {
      this.routes.delete(String(rid));
    }

    this.mergeCarretas(snapshotObj.__meta?.carretas);

    for (const [routeId, ser] of Object.entries(snapshotObj)) {
      const id = String(routeId);
      if (id === '__meta') continue;
      if (this.deletedRoutes.has(id)) continue;

      const existing = this.routes.get(id);

      if (!existing) {
        this.routes.set(id, this.deserializeRoute(id, ser));
        continue;
      }

      const tmp = this.deserializeRoute(id, ser);

      tmp.ids.forEach(v => existing.ids.add(v));
      tmp.conferidos.forEach(v => existing.conferidos.add(v));
      tmp.foraDeRota.forEach(v => existing.foraDeRota.add(v));

      // Campos de texto: pega o preenchido; se ambos preenchidos e diferentes, escolha determinística
      for (const f of ['cluster', 'destinationFacilityId', 'destinationFacilityName']) {
        const a = String(existing[f] || '');
        const b = String(tmp[f] || '');
        if (!a || (b && b > a)) existing[f] = b || a;
      }

      // Vínculo placa/rota: vence o mais recente
      if (Number(tmp.plateUpdatedAt || 0) > Number(existing.plateUpdatedAt || 0)) {
        for (const f of ['plateKey', 'plateRaw', 'plateLicense', 'routeQrKey', 'routeQrRaw', 'plateScanTs', 'routeQrScanTs', 'plateUpdatedAt']) {
          existing[f] = tmp[f];
        }
      }

      for (const [k, v] of tmp.timestamps.entries()) {
        const cur = existing.timestamps.get(k);
        if (!cur || Number(v) > Number(cur)) existing.timestamps.set(k, v);
      }
      for (const [k, v] of tmp.duplicados.entries()) {
        existing.duplicados.set(k, Math.max(this.dupCount(existing.duplicados.get(k)), this.dupCount(v)));
      }

      // Um ID conferido não pode continuar em "fora de rota" na mesma rota
      for (const c of existing.conferidos) existing.foraDeRota.delete(c);

      existing.totalInicial = Math.max(Number(existing.totalInicial || 0), Number(tmp.totalInicial || 0), existing.ids.size);
      existing.faltantes = new Set(existing.ids);
      for (const c of existing.conferidos) existing.faltantes.delete(c);
    }

    // A união traria de volta "fora de rota" já resolvidos em outro aparelho; limpa aqui
    this.globalCleanupForaDeRotaForConferidos();
  },

  async syncFromSupabaseForDay(dayISO) {
    const op = this.getOperationCode();
    if (!op) return;
    if (this.cloudSaving) return;

    try {
      const row = await this.supaLoadDaySnapshot(op, dayISO);
      if (!row || !row.data) {
        // Banco ainda sem esse dia, mas há dados locais (ex.: bipado offline): envia
        if (this.routes.size && op === this.getOperationCode() && dayISO === this.workDay) {
          this.markCloudDirty(dayISO, op);
        }
        return;
      }

      // O dia/operação pode ter mudado enquanto a consulta estava em andamento
      if (op !== this.getOperationCode() || dayISO !== this.workDay) return;

      const remoteHash = this.computeSnapshotHash(row.data);
      if (remoteHash === this.lastPushedSnapshotHash) return;

      this.lastRemoteUpdatedAt = row.updated_at || this.lastRemoteUpdatedAt;
      this.mergeSnapshotIntoLocal(row.data);
      this.saveToStorage(dayISO, { syncCloud: false });

      // Se o local tem algo que o banco não tem (ex.: bipado offline), envia
      const localHash = this.computeSnapshotHash(this.buildDaySnapshotObject());
      if (localHash !== remoteHash) {
        this.markCloudDirty(dayISO, op);
      } else {
        this.lastPushedSnapshotHash = remoteHash;
      }

      this.refreshAfterRemoteMerge();
      this.setStatus(`Sincronizado do banco • ${op} • ${dayISO}`, 'info');
    } catch (e) {
      if (e && (e.code === '42P01' || /routes_state/i.test(String(e.message || '')))) {
        this.setStatus('Tabela routes_state não existe no Supabase. Crie a tabela para sincronizar.', 'danger');
        console.warn('Crie no Supabase:', this.getRoutesStateCreateSQL());
        return;
      }
      console.warn('Falha ao sincronizar do Supabase:', e);
    }
  },

  startPeriodicSync() {
    if (this.syncTimer) clearInterval(this.syncTimer);
    this.syncTimer = setInterval(() => {
      if (!this.workDay || !this.getOperationCode()) return;
      if (this.cloudDirty) {
        // Save travado (ex.: falha de rede sem retry agendado)? força nova tentativa
        if (!this.cloudSaveTimer && !this.cloudSaving) this.scheduleCloudSave(0);
        return;
      }
      this.syncFromSupabaseForDay(this.workDay);
    }, SYNC_INTERVAL_MS);
  },

  getRoutesStateCreateSQL() {
    return `create table if not exists public.routes_state (
  operation_code text not null references public.operations(code) on delete restrict,
  day date not null,
  data jsonb not null,
  updated_at timestamptz not null default now(),
  device_id text,
  primary key (operation_code, day)
);

alter table public.routes_state enable row level security;

create policy "routes_state_select_all"
  on public.routes_state for select
  using (true);

create policy "routes_state_upsert_all"
  on public.routes_state for insert
  with check (true);

create policy "routes_state_update_all"
  on public.routes_state for update
  using (true)
  with check (true);`;
  },

  async ensureOperationSelected() {
    const sb = this.getSb();
    if (!sb) return;

    const { data, error } = await sb.from('operations').select('code,name,active').eq('active', true).order('code', { ascending: true });
    if (error) {
      console.warn('Falha ao carregar operações:', error);
      return;
    }

    const $sel = $('#op-select');
    if ($sel.length) {
      $sel.empty();
      (data || []).forEach(op => {
        const label = op.name ? `${op.code} — ${op.name}` : op.code;
        $sel.append(`<option value="${this.escHtml(op.code)}">${this.escHtml(label)}</option>`);
      });
    }

    const current = this.getOperationCode();
    const exists = (data || []).some(o => o.code === current);

    if ($sel.length && current && exists) {
      $sel.val(current);
    }

    if (!current || !exists) {
      $('#modal-operation').modal({ backdrop: 'static', keyboard: false });
    } else {
      this.setOperationCode(current);
    }
  },

  setStatus(txt, kind = 'muted') {
    const $s = $('#sync-status');
    $s.removeClass('text-muted text-success text-danger text-warning text-info');
    $s.addClass(`text-${kind}`);
    $s.text(txt);
  },

  markDirty(reason = '') {
    this.dirty = true;
    this.markCloudDirty();
    const msg = reason ? `pendente salvar (${reason})` : 'pendente salvar';
    this.setStatus(msg, 'warning');
    $('#dirty-flag').removeClass('d-none');
  },

  markClean() {
    this.dirty = false;
    $('#dirty-flag').addClass('d-none');
    this.setStatus('sincronizado', 'success');
  },

  // =======================
  // Busca/Auditoria no banco
  // =======================
  getRoutesMap() {
    if (this.routes instanceof Map) return this.routes;
    const m = new Map();
    const obj = this.routes || {};
    Object.keys(obj).forEach(k => m.set(String(k), obj[k]));
    this.routes = m;
    return m;
  },

  getLocalStatusForId(idRaw) {
    const id = String(idRaw || '').trim();
    if (!id) return null;

    for (const r of this.routes.values()) {
      const rid = String(r.routeId || '');
      const cl = r.cluster ? String(r.cluster).trim() : '';
      const xpt = (r.destinationFacilityId != null && r.destinationFacilityId !== '') ? String(r.destinationFacilityId) : '';

      if (r.conferidos?.has(id)) return { where: 'local', status: 'conferido', route_id: rid, cluster: cl, xpt };
      if (r.foraDeRota?.has(id)) return { where: 'local', status: 'fora', route_id: rid, cluster: cl, xpt };
      if (r.duplicados?.has(id)) return { where: 'local', status: 'duplicado', route_id: rid, cluster: cl, xpt };
      if (r.faltantes?.has(id)) return { where: 'local', status: 'faltante', route_id: rid, cluster: cl, xpt };
      if (r.ids?.has(id)) return { where: 'local', status: 'pertence', route_id: rid, cluster: cl, xpt };
    }

    return null;
  },

  async searchIdsFull(idsRaw, opts = {}) {
    const ids = Array.isArray(idsRaw) ? idsRaw.map(String) : this.parseIdsList(idsRaw);
    if (!ids.length) return { ids: [], rows: [], summary: [] };

    const rows = await this.searchScanEventsByIds(ids, opts);

    const byId = new Map();
    for (const id of ids) {
      byId.set(id, { id, events: [], ops: new Set(), last: null, local: null });
    }

    for (const r of rows) {
      const pid = String(r.package_id ?? '');
      if (!byId.has(pid)) continue;
      const ref = byId.get(pid);
      ref.events.push(r);
      if (r.operation_code) ref.ops.add(String(r.operation_code));
    }

    const summary = [];
    for (const id of ids) {
      const ref = byId.get(id);
      const last = (ref.events && ref.events.length) ? ref.events[0] : null;
      const local = !last ? this.getLocalStatusForId(id) : null;

      ref.last = last;
      ref.local = local;

      summary.push({
        id,
        last_seen_at: last ? last.scanned_at : null,
        last_operation: last ? last.operation_code : null,
        last_day: last ? last.day : null,
        last_route_id: last ? last.route_id : null,
        last_cluster: last ? last.cluster : null,
        last_xpt: last ? last.xpt : null,
        last_result: last ? last.result : null,
        operations: Array.from(ref.ops),
        local_status: local ? local.status : null,
        local_route_id: local ? local.route_id : null,
        local_cluster: local ? local.cluster : null,
        local_xpt: local ? local.xpt : null,
        has_db_history: !!last
      });
    }

    return { ids, rows, summary };
  },

  getRouteMeta(routeId) {
    const rid = routeId != null ? String(routeId) : '';
    const routesMap = this.getRoutesMap();
    const r = routesMap.get(rid);
    if (!r) return { routeId: rid || null, cluster: null, xpt: null };
    return {
      routeId: String(r.routeId || rid || '') || null,
      cluster: r.cluster ? String(r.cluster) : null,
      xpt: (r.destinationFacilityId != null && r.destinationFacilityId !== '') ? String(r.destinationFacilityId) : null
    };
  },

  enrichEventForCloud(evt) {
    const e = Object.assign({}, evt || {});
    const meta = this.getRouteMeta(e.currentRouteId);
    e._meta = meta;
    return e;
  },

  async logScanEventToCloud(evt) {
    try {
      const sb = this.getSb();
      if (!sb) return;

      const op = this.getOperationCode();
      const dayISO = this.workDay || this.todayLocalISO();
      if (!op || !dayISO) return;

      const e = this.enrichEventForCloud(evt);
      const code = this.normalizarCodigo(e.code);
      if (!code) return;
      if (!/^\d+$/.test(String(code))) return;

      const payload = {
        operation_code: String(op),
        day: String(dayISO),
        package_id: String(code),
        scanned_at: new Date(Number(e.ts || Date.now())).toISOString(),
        result: String(e.type || '').toLowerCase() || null,
        route_id: e._meta?.routeId || null,
        cluster: e._meta?.cluster || null,
        xpt: e._meta?.xpt || null
      };

      const { error } = await sb.from('scan_events').insert(payload);
      if (error) console.warn('Falha ao gravar scan_events:', error);
    } catch (err) {
      console.warn('Falha ao gravar scan_events (catch):', err);
    }
  },

  parseIdsList(raw) {
    const txt = String(raw || '');
    // Aceita qualquer ID numérico (não só o padrão de 11 dígitos da bipagem)
    const ids = txt.split(/[;,\s]+/g)
      .map(p => this.normalizarCodigo(p) || String(p).replace(/\D+/g, ''))
      .filter(p => p && p.length >= 5);
    return Array.from(new Set(ids));
  },

  async searchScanEventsByIds(idsRaw, opts = {}) {
    const sb = this.getSb();
    if (!sb) throw new Error('Supabase client não encontrado.');

    const ids = Array.isArray(idsRaw) ? idsRaw.map(String) : this.parseIdsList(idsRaw);
    if (!ids.length) return [];

    const op = (opts.operation_code ? String(opts.operation_code) : '').trim().toUpperCase();
    const dayFrom = opts.day_from ? String(opts.day_from) : null;
    const dayTo = opts.day_to ? String(opts.day_to) : null;

    const BATCH = 200;
    const out = [];

    for (let i = 0; i < ids.length; i += BATCH) {
      const batch = ids.slice(i, i + BATCH);

      let q = sb.from('scan_events')
        .select('package_id,operation_code,day,scanned_at,route_id,cluster,xpt,result')
        .in('package_id', batch)
        .order('scanned_at', { ascending: false });

      if (op) q = q.eq('operation_code', op);
      if (dayFrom) q = q.gte('day', dayFrom);
      if (dayTo) q = q.lte('day', dayTo);

      const { data, error } = await q;
      if (error) throw error;
      (data || []).forEach(r => out.push(r));
    }

    out.sort((a, b) => String(b.scanned_at).localeCompare(String(a.scanned_at)));
    return out;
  },

  renderDbSearchSummary(summaryRows) {
    const $tb = $('#db-search-results');
    const $wrap = $('#db-search-results-wrap');
    if (!$tb.length) return;

    $tb.empty();

    (summaryRows || []).forEach(r => {
      const dt = r.last_seen_at ? new Date(r.last_seen_at) : null;
      const day = r.last_day || (dt ? dt.toISOString().slice(0, 10) : '');
      const hhmm = dt ? String(dt.getHours()).padStart(2, '0') + ':' + String(dt.getMinutes()).padStart(2, '0') : '';

      const ops = (r.operations && r.operations.length) ? r.operations.join(',') : '';
      const status = r.has_db_history
        ? (r.last_result || '')
        : (r.local_status ? `local:${r.local_status}` : 'sem histórico');

      const routeId = r.has_db_history ? (r.last_route_id ?? '') : (r.local_route_id ?? '');
      const cluster = r.has_db_history ? (r.last_cluster ?? '') : (r.local_cluster ?? '');
      const xpt = r.has_db_history ? (r.last_xpt ?? '') : (r.local_xpt ?? '');

      const esc = (x) => this.escHtml(x);
      $tb.append(`
        <tr>
          <td>${esc(r.id)}</td>
          <td>${esc(r.has_db_history ? (r.last_operation ?? '') : '')}</td>
          <td>${esc(day)}</td>
          <td>${hhmm}</td>
          <td>${esc(routeId)}</td>
          <td>${esc(cluster)}</td>
          <td>${esc(xpt)}</td>
          <td>${esc(status)} ${ops ? `<small class="text-muted">(${esc(ops)})</small>` : ''}</td>
        </tr>
      `);
    });

    if ($wrap.length) $wrap.removeClass('d-none');
  },

  pushEvent(evt) {
    const e = this.enrichEventForCloud(evt);
    this.lastEvents.unshift(e);
    if (this.lastEvents.length > 80) this.lastEvents.length = 80;
    this.renderAcompanhamento();
    this.logScanEventToCloud(e);
  },

  renderAcompanhamento() {
    if (this.dirty) $('#dirty-flag').removeClass('d-none');
    else $('#dirty-flag').addClass('d-none');

    const mapa = new Map();

    for (const r of this.routes.values()) {
      const cluster = (r.cluster && String(r.cluster).trim()) || '(sem cluster)';
      const total = (r.totalInicial || r.ids.size || 0);
      const conf = (r.conferidos ? r.conferidos.size : 0);

      if (!mapa.has(cluster)) {
        mapa.set(cluster, {
          conferidos: 0,
          total: 0,
          precisaRevalidar: false
        });
      }

      const acc = mapa.get(cluster);
      acc.conferidos += conf;
      acc.total += total;

      if (conf < total) {
        acc.precisaRevalidar = true;
      }
    }

    const resumo = Array.from(mapa.entries())
      .sort((a, b) => String(a[0]).localeCompare(String(b[0])))
      .map(([cluster, v]) => {
        const ok = (v.total > 0 && v.conferidos === v.total) ? ' ✅' : '';
        return `CLUSTER ${this.escHtml(cluster)}: ${v.conferidos}/${v.total}${ok}`;
      })
      .join('<br>');

    $('#acompanhamento-resumo')
      .css({
        'max-height': '200px',
        'overflow-y': 'auto',
        'overflow-x': 'hidden'
      })
      .html(resumo || '<span class="text-muted">sem clusters</span>');

    const textoClusters = Array.from(mapa.entries())
      .sort((a, b) => String(a[0]).localeCompare(String(b[0])))
      .map(([cluster, v]) => {
        return `CLUSTER ${cluster}: ${v.precisaRevalidar ? 'REVALIDAR' : 'CONTAR'}`;
      })
      .join('\n');

    $('#clusters-copiavel').val(textoClusters);

    const items = this.lastEvents.slice(0, 30).map(ev => {
      const d = new Date(ev.ts);
      const hh = String(d.getHours()).padStart(2, '0');
      const mm = String(d.getMinutes()).padStart(2, '0');
      const ss = String(d.getSeconds()).padStart(2, '0');
      const base = `${hh}:${mm}:${ss} • ${this.escHtml(ev.code)}`;

      if (ev.type === 'fora') {
        const rr = ev.correctRouteId ? this.routes.get(String(ev.correctRouteId)) : null;
        const cl = rr && rr.cluster ? String(rr.cluster).trim() : '';
        const extra = (cl ? ` <small class="text-muted">(CLUSTER ${this.escHtml(cl)})</small>` : ' <small class="text-muted">(cluster desconhecido)</small>');
        return `<li class="list-group-item list-group-item-warning">${base}${extra}</li>`;
      }
      if (ev.type === 'dup') {
        return `<li class="list-group-item list-group-item-secondary">${base} (duplicado)</li>`;
      }
      return `<li class="list-group-item list-group-item-success">${base} (ok)</li>`;
    });

    $('#acompanhamento-log').html(items.join('') || '<li class="list-group-item text-muted">sem eventos</li>');
  },

  // =======================
  // Relatório Noturno
  // =======================
  escapeHtml(s) {
    return this.escHtml(s);
  },

  locateIdInOtherRoutes(id, excludeRouteId) {
    const ex = String(excludeRouteId ?? '');
    for (const r of this.routes.values()) {
      if (String(r.routeId) === ex) continue;
      if (r.conferidos && r.conferidos.has(id)) {
        return { where: 'conferido', routeId: String(r.routeId), cluster: String(r.cluster || '').trim() };
      }
    }
    for (const r of this.routes.values()) {
      if (String(r.routeId) === ex) continue;
      if (r.foraDeRota && r.foraDeRota.has(id)) {
        return { where: 'fora', routeId: String(r.routeId), cluster: String(r.cluster || '').trim() };
      }
    }
    for (const r of this.routes.values()) {
      if (String(r.routeId) === ex) continue;
      if ((r.ids && r.ids.has(id)) || (r.faltantes && r.faltantes.has(id))) {
        return { where: 'pertence', routeId: String(r.routeId), cluster: String(r.cluster || '').trim() };
      }
    }
    return null;
  },

  buildNightReportHtml() {
    const esc = (x) => this.escHtml(x);

    const routesSorted = Array.from(this.routes.values()).sort((a, b) => {
      const ca = String(a.cluster || '').trim();
      const cb = String(b.cluster || '').trim();
      const byC = ca.localeCompare(cb);
      if (byC) return byC;
      return String(a.routeId).localeCompare(String(b.routeId));
    });

    const items = routesSorted.map(r => {
      const total = Number(r.totalInicial || r.ids.size || 0);
      const conf = Number(r.conferidos?.size || 0);
      const falt = Number(r.faltantes?.size || 0);
      const ok = total > 0 ? (conf === total) : (falt === 0);
      return {
        routeId: String(r.routeId),
        cluster: String(r.cluster || '').trim(),
        total, conf, falt,
        ok,
        route: r
      };
    });

    const completas = items.filter(x => x.ok);
    const incompletas = items.filter(x => !x.ok);

    const headerLine = (it) => {
      const c = it.cluster ? `CLUSTER ${esc(it.cluster)}` : 'CLUSTER (vazio)';
      const perc = it.total ? Math.floor((it.conf / it.total) * 100) : 0;
      const badge = it.ok ? `<span class="badge badge-success ml-2">100%</span>` : `<span class="badge badge-warning ml-2">${esc(it.falt)} falt.</span>`;
      return `${c} <small class="text-muted">• Rota ${esc(it.routeId)} • ${esc(it.conf)}/${esc(it.total)} (${perc}%)</small>${badge}`;
    };

    const listCompletas = completas.length
      ? completas.map(it => `<li class="list-group-item d-flex justify-content-between align-items-center">${headerLine(it)}</li>`).join('')
      : `<li class="list-group-item text-muted">Nenhuma rota 100%.</li>`;

    const blocosIncompletas = incompletas.length
      ? incompletas.map((it, idx) => {
          const collapseId = `nr_${esc(it.routeId)}_${idx}`;
          const r = it.route;
          const faltantes = Array.from(r.faltantes || []);
          faltantes.sort((a, b) => String(a).localeCompare(String(b)));

          const faltHtml = faltantes.map(id => {
            const loc = this.locateIdInOtherRoutes(id, r.routeId);
            if (!loc) return `<li class="list-group-item list-group-item-danger">${esc(id)}</li>`;
            const whereTxt = (loc.where === 'conferido')
              ? 'conferido'
              : (loc.where === 'fora')
                ? 'bipado (fora de rota)'
                : 'pertence à rota';
            const cl = loc.cluster ? `CLUSTER ${esc(loc.cluster)}` : 'CLUSTER (vazio)';
            return `<li class="list-group-item list-group-item-warning">
                      ${esc(id)}
                      <small class="text-muted ml-2">→ ${whereTxt} em ${cl} (Rota ${esc(loc.routeId)})</small>
                    </li>`;
          }).join('') || `<li class="list-group-item text-muted">Sem faltantes</li>`;

          return `
            <div class="card mb-2">
              <div class="card-header p-2">
                <button class="btn btn-link p-0" type="button" data-toggle="collapse" data-target="#${collapseId}">
                  ${headerLine(it)}
                </button>
              </div>
              <div id="${collapseId}" class="collapse">
                <ul class="list-group list-group-flush">
                  ${faltHtml}
                </ul>
              </div>
            </div>
          `;
        }).join('')
      : `<div class="text-muted">Nenhuma rota com faltantes.</div>`;

    return `
      <div class="mb-2">
        <div class="small text-muted">Relatório do dia ${esc(this.workDay || this.todayLocalISO())}</div>
        <div class="small text-muted">Mostra todas as rotas, rotas 100% e faltantes (indicando onde foram vistos em outras rotas).</div>
      </div>

      <div class="mb-2">
        <span class="badge badge-success">100%: ${completas.length}</span>
        <span class="badge badge-warning ml-1">Com faltantes: ${incompletas.length}</span>
        <span class="badge badge-light ml-1">Total: ${routesSorted.length}</span>
      </div>

      <div class="mt-2">
        <div class="font-weight-bold mb-1">Rotas 100%</div>
        <ul class="list-group mb-3">
          ${listCompletas}
        </ul>

        <div class="font-weight-bold mb-1">Rotas com faltantes</div>
        ${blocosIncompletas}
      </div>
    `;
  },

  showNightReport() {
    const html = this.buildNightReportHtml();
    $('#night-report').html(html).removeClass('d-none');
    $('#acompanhamento-resumo').addClass('d-none');
    $('#acompanhamento-log').closest('div').addClass('d-none');
    $('#finish-night-btn').addClass('d-none');
    $('#night-report-close').removeClass('d-none');
  },

  hideNightReport() {
    $('#night-report').addClass('d-none').html('');
    $('#acompanhamento-resumo').removeClass('d-none');
    $('#acompanhamento-log').closest('div').removeClass('d-none');
    $('#finish-night-btn').removeClass('d-none');
    $('#night-report-close').addClass('d-none');
  },

  // =======================
  // Modelo de rota
  // =======================
  makeEmptyRoute(routeId) {
    return {
      routeId: String(routeId),
      cluster: '',
      destinationFacilityId: '',
      destinationFacilityName: '',

      timestamps: new Map(),
      ids: new Set(),
      faltantes: new Set(),
      conferidos: new Set(),
      foraDeRota: new Set(),
      duplicados: new Map(),

      totalInicial: 0,

      plateKey: '',
      plateRaw: '',
      plateLicense: '',
      routeQrKey: '',
      routeQrRaw: '',

      plateScanTs: 0,
      routeQrScanTs: 0,
      plateUpdatedAt: 0
    };
  },

  get current() {
    if (!this.currentRouteId) return null;
    return this.routes.get(String(this.currentRouteId)) || null;
  },

  // =======================
  // Persistência local
  // =======================
  loadFromStorage(dayISO) {
    try {
      const key = this.storageKeyForDay(dayISO);
      const raw = localStorage.getItem(key);
      if (!raw) return;

      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed !== 'object') return;

      this.deletedRoutes = new Map(Object.entries(parsed.__meta?.deletedRoutes || {}).map(([k, v]) => [String(k), Number(v || 0)]));
      this.revivedRoutes = new Map(Object.entries(parsed.__meta?.revivedRoutes || {}).map(([k, v]) => [String(k), Number(v || 0)]));

      for (const [rid, rts] of this.revivedRoutes.entries()) {
        const delTs = Number(this.deletedRoutes.get(rid) || 0);
        if (Number(rts || 0) > delTs) this.deletedRoutes.delete(rid);
      }

      this.routes.clear();
      this.currentRouteId = null;

      for (const [routeId, r] of Object.entries(parsed)) {
        if (routeId === '__meta') continue;
        if (this.deletedRoutes.has(String(routeId))) continue;
        this.routes.set(String(routeId), this.deserializeRoute(routeId, r));
      }

      this.mergeCarretas(parsed.__meta?.carretas);
      this.lastSavedLocalSnapshotHash = this.computeSnapshotHash(this.buildDaySnapshotObject());
    } catch (e) {
      console.warn('Falha ao carregar storage:', e);
    }
  },

  saveToStorage(dayISO, opts = {}) {
    const { syncCloud = true } = opts;

    try {
      const key = this.storageKeyForDay(dayISO);
      const obj = this.buildDaySnapshotObject();

      const hash = this.computeSnapshotHash(obj);
      if (hash !== this.lastSavedLocalSnapshotHash) {
        localStorage.setItem(key, JSON.stringify(obj));
        this.lastSavedLocalSnapshotHash = hash;
      }

      if (syncCloud && hash !== this.lastPushedSnapshotHash) {
        this.markCloudDirty(dayISO, this.getOperationCode());
      }
    } catch (e) {
      console.warn('Falha ao salvar storage:', e);
    }
  },

  serializeRoute(r) {
    return {
      routeId: r.routeId,
      cluster: r.cluster,
      destinationFacilityId: r.destinationFacilityId,
      destinationFacilityName: r.destinationFacilityName,
      totalInicial: r.totalInicial,

      plateKey: r.plateKey || '',
      plateRaw: r.plateRaw || '',
      plateLicense: r.plateLicense || '',
      routeQrKey: r.routeQrKey || '',
      routeQrRaw: r.routeQrRaw || '',

      plateScanTs: Number(r.plateScanTs || 0),
      routeQrScanTs: Number(r.routeQrScanTs || 0),
      plateUpdatedAt: Number(r.plateUpdatedAt || 0),

      ids: Array.from(r.ids),
      faltantes: Array.from(r.faltantes),
      conferidos: Array.from(r.conferidos),
      foraDeRota: Array.from(r.foraDeRota),

      timestamps: Object.fromEntries(r.timestamps),
      duplicados: Object.fromEntries(r.duplicados),
    };
  },

  deserializeRoute(routeId, r) {
    const route = this.makeEmptyRoute(routeId);
    r = r || {};

    route.cluster = r.cluster || '';
    route.destinationFacilityId = r.destinationFacilityId || '';
    route.destinationFacilityName = r.destinationFacilityName || '';
    route.totalInicial = Number(r.totalInicial || 0);

    route.plateKey = r.plateKey || '';
    route.plateRaw = r.plateRaw || '';
    route.plateLicense = r.plateLicense || '';
    route.routeQrKey = r.routeQrKey || '';
    route.routeQrRaw = r.routeQrRaw || '';

    route.plateScanTs = Number(r.plateScanTs || 0);
    route.routeQrScanTs = Number(r.routeQrScanTs || 0);
    route.plateUpdatedAt = Number(r.plateUpdatedAt || r.routeQrScanTs || 0);

    (r.ids || []).forEach(id => route.ids.add(id));
    (r.conferidos || []).forEach(id => route.conferidos.add(id));
    (r.foraDeRota || []).forEach(id => route.foraDeRota.add(id));

    route.timestamps = new Map(Object.entries(r.timestamps || {}));
    route.duplicados = new Map(Object.entries(r.duplicados || {}).map(([k, v]) => [k, this.dupCount(v)]));

    // Faltantes sempre derivados de ids - conferidos (evita lista desatualizada)
    if (route.ids.size) {
      route.faltantes = new Set(route.ids);
      for (const c of route.conferidos) route.faltantes.delete(c);
    } else {
      route.faltantes = new Set(r.faltantes || []);
    }

    return route;
  },

  // Remove "fora de rota" de IDs que depois foram conferidos na rota certa.
  // Mantém o alerta quando o "fora" aconteceu DEPOIS da conferência (pacote foi parar na gaiola errada).
  globalCleanupForaDeRotaForConferidos() {
    const conferidoTs = new Map(); // id -> ts da conferência
    for (const r of this.routes.values()) {
      for (const id of r.conferidos) {
        const ts = Number(r.timestamps.get(id) || 0);
        conferidoTs.set(id, Math.max(conferidoTs.get(id) || 0, ts));
      }
    }
    if (!conferidoTs.size) return;

    for (const r of this.routes.values()) {
      for (const id of Array.from(r.foraDeRota)) {
        if (!conferidoTs.has(id)) continue;
        if (r.conferidos.has(id)) {
          r.foraDeRota.delete(id);
          continue;
        }
        const foraTs = Number(r.timestamps.get(id) || 0);
        if (foraTs <= conferidoTs.get(id)) {
          r.foraDeRota.delete(id);
          r.duplicados.delete(id);
        }
      }
    }
  },

  // =======================
  // Troca de dia
  // =======================
  // Envia ao banco o que estiver pendente (usado antes de trocar dia/operação)
  async flushPendingNow() {
    if (this.cloudSaveTimer) {
      clearTimeout(this.cloudSaveTimer);
      this.cloudSaveTimer = null;
    }
    for (let i = 0; i < 50 && this.cloudSaving; i++) {
      await new Promise(res => setTimeout(res, 100));
    }
    if (!this.cloudDirty) return;
    try {
      const pending = this.pendingCloudSave || {};
      await this.flushCloudSave(pending.op, pending.day);
    } catch (e) {
      console.warn("Falha ao enviar pendências antes da troca (ficam no cache local):", e);
    }
    if (this.cloudSaveTimer) {
      clearTimeout(this.cloudSaveTimer);
      this.cloudSaveTimer = null;
    }
  },

  async applyWorkDay(dayISO) {
    await this.flushPendingNow();
    this.cloudDirty = false;
    this.pendingCloudSave = null;

    this.routes.clear();
    this.resetCarretas();
    this.currentRouteId = null;
    this.lastRoutesSignature = '';

    this.deletedRoutes = new Map();
    this.revivedRoutes = new Map();
    this.lastEvents = [];
    this.lastRemoteUpdatedAt = null;
    this.lastPushedSnapshotHash = '';
    this.lastSavedLocalSnapshotHash = '';

    this.workDay = dayISO;
    $('#work-day').val(dayISO);

    const op = this.getOperationCode();
    if (op) $('#op-badge').text(op);

    this.loadFromStorage(dayISO);

    if (op) {
      await this.syncFromSupabaseForDay(dayISO);
      await this.startRealtimeSync(dayISO);
      this.startPeriodicSync();
    }

    this.renderRoutesSelects();
    this.refreshUIFromCurrent();
    this.renderAcompanhamento();

    if (op) this.setStatus(`dia carregado • ${op} • ${dayISO}`, 'success');
    else this.setStatus('dia carregado (sem operação)', 'warning');
  },

  resetForOperationChange() {
    this.stopRealtimeSync();
    this.routes.clear();
    this.resetCarretas();
    this.currentRouteId = null;
    this.viaCsv = false;
    this.lastRoutesSignature = '';

    try {
      $('#saved-routes').html('<option value="">(Nenhuma selecionada)</option>');
      $('#saved-routes-inapp').empty();
      $('#carreta-routes').empty();
      $('#fora-rota-list').empty();
      $('#route-title').text('');
      $('#cluster-title').text('');
      $('#destination-facility-title').text('');
      $('#destination-facility-name').text('');
      $('#extracted-total').text('0');
      $('#verified-total').text('0');
    } catch {}

    $('#global-interface').addClass('d-none');
    $('#carreta-interface').addClass('d-none');
    $('#manual-interface').addClass('d-none');
    $('#initial-interface').removeClass('d-none');
  },

  computeSnapshotStats(snapshotObj) {
    const routes = snapshotObj && typeof snapshotObj === 'object' ? Object.values(snapshotObj) : [];
    let routesCount = 0, totalIds = 0, conferidos = 0, faltantes = 0, fora = 0;

    for (const r of routes) {
      if (!r || typeof r !== 'object') continue;
      routesCount += 1;
      const idsArr = Array.isArray(r.ids) ? r.ids : [];
      const confArr = Array.isArray(r.conferidos) ? r.conferidos : [];
      const faltArr = Array.isArray(r.faltantes) ? r.faltantes : [];
      const foraArr = Array.isArray(r.foraDeRota) ? r.foraDeRota : [];
      totalIds += idsArr.length;
      conferidos += confArr.length;
      faltantes += faltArr.length;
      fora += foraArr.length;
    }
    return { routesCount, totalIds, conferidos, faltantes, fora };
  },

  async loadGlobalProgress(dayISO) {
    const sb = this.getSb();
    if (!sb) throw new Error('Supabase client não encontrado (window.sbClient).');

    const { data: ops, error: e1 } = await sb
      .from('operations')
      .select('code,name,active')
      .eq('active', true)
      .order('code', { ascending: true });
    if (e1) throw e1;

    const tasks = (ops || []).map(async (o) => {
      const code = String(o.code || '').toUpperCase();
      const row = await this.supaLoadDaySnapshot(code, dayISO).catch(() => null);
      const stats = this.computeSnapshotStats(row && row.data ? row.data : {});
      return {
        code,
        name: o.name || '',
        stats,
        updated_at: row ? row.updated_at : null,
        device_id: row ? row.device_id : null
      };
    });

    return Promise.all(tasks);
  },

  renderGlobalProgress(items, dayISO) {
    $('#global-day-label').text(dayISO);
    const $tb = $('#global-ops-tbody');
    const $log = $('#global-log');
    $tb.empty();
    $log.empty();

    const sorted = (items || []).slice().sort((a, b) => (a.code || '').localeCompare(b.code || ''));

    for (const it of sorted) {
      const u = it.updated_at ? new Date(it.updated_at).toLocaleString('pt-BR') : '-';
      $tb.append(`
        <tr>
          <td><strong>${this.escHtml(it.code)}</strong>${it.name ? ` <span class="text-muted small">(${this.escapeHtml(it.name)})</span>` : ''}</td>
          <td>${it.stats.routesCount}</td>
          <td>${it.stats.totalIds}</td>
          <td>${it.stats.conferidos}</td>
          <td>${it.stats.faltantes}</td>
          <td>${it.stats.fora}</td>
          <td>${u}</td>
        </tr>
      `);

      $log.append(`
        <li class="list-group-item d-flex justify-content-between align-items-center">
          <span><strong>${this.escHtml(it.code)}</strong> • ${it.stats.conferidos}/${it.stats.totalIds} conferidos • ${it.stats.faltantes} faltantes</span>
          <span class="badge badge-light">${u}</span>
        </li>
      `);
    }

    if (!sorted.length) {
      $tb.append('<tr><td colspan="7" class="text-muted">Nenhuma operação ativa encontrada.</td></tr>');
    }
  },

  // =======================
  // UI / Rotas
  // =======================
  setCurrentRoute(routeId) {
    console.log('[setCurrentRoute] versão', '2026-02-22-1');

    const id = String(routeId);
    if (!this.routes.has(id)) {
      alert('Rota não encontrada.');
      return;
    }

    this.currentRouteId = id;
    this.renderRoutesSelects();
    this.refreshUIFromCurrent();
    this.renderAcompanhamento();
  },

  renderRoutesSelects() {
    const $sel1 = $('#saved-routes');
    const $sel2 = $('#saved-routes-inapp');

    const routesSorted = Array.from(this.routes.values())
      .sort((a, b) => String(a.routeId).localeCompare(String(b.routeId)));

    const makeLabel = (r) => {
      const parts = [];
      if (r.cluster) parts.push(`CLUSTER ${r.cluster}`);
      if (r.destinationFacilityId) parts.push(`XPT ${r.destinationFacilityId}`);
      return parts.join(' • ') || `(sem dados) • Rota ${r.routeId}`;
    };

    const newCache = routesSorted.map(r => ({
      routeId: String(r.routeId),
      label: makeLabel(r),
      clusterKey: this.normalizeCluster(r.cluster || ''),
      labelKey: this.normalizeCluster(makeLabel(r)),
    }));

    const sig = newCache
      .map(x => `${x.routeId}|${x.clusterKey}|${x.label}`)
      .join('||');

    const changed = sig !== this.lastRoutesSignature;

    if (changed) {
      this.lastRoutesSignature = sig;
      this._routesDropdownCache = newCache;

      $sel1.html(
        ['<option value="">(Nenhuma selecionada)</option>']
          .concat(this._routesDropdownCache.map(x => `<option value="${this.escHtml(x.routeId)}">${this.escHtml(x.label)}</option>`))
          .join('')
      );

      this.applyRouteDropdownFilter($('#route-search').val() || '');
    }

    if (this.currentRouteId) {
      $sel1.val(String(this.currentRouteId));

      const filteredText = $('#route-search').val() || '';
      const q = this.normalizeCluster(String(filteredText).trim());

      const filtered = !q
        ? (this._routesDropdownCache || [])
        : (this._routesDropdownCache || []).filter(x =>
            (x.clusterKey && x.clusterKey.includes(q)) ||
            (x.labelKey && x.labelKey.includes(q)) ||
            String(x.routeId).includes(String(filteredText).trim())
          );

      const existsInFiltered = filtered.some(x => x.routeId === String(this.currentRouteId));
      if (existsInFiltered) {
        $sel2.val(String(this.currentRouteId));
      }
    }
  },

  applyRouteDropdownFilter(filterText) {
    const $sel2 = $('#saved-routes-inapp');

    const list = Array.isArray(this._routesDropdownCache) ? this._routesDropdownCache : [];
    const q = this.normalizeCluster(String(filterText || '').trim());

    const filtered = !q
      ? list
      : list.filter(x =>
          (x.clusterKey && x.clusterKey.includes(q)) ||
          (x.labelKey && x.labelKey.includes(q)) ||
          String(x.routeId).includes(String(filterText || '').trim())
        );

    const options = ['<option value="">(Selecione)</option>']
      .concat(filtered.map(x => `<option value="${this.escHtml(x.routeId)}">${this.escHtml(x.label)}</option>`))
      .join('');

    $sel2.html(options);

    if (this.currentRouteId) {
      const exists = filtered.some(x => x.routeId === String(this.currentRouteId));
      if (exists) $sel2.val(String(this.currentRouteId));
    }
  },

  refreshUIFromCurrent() {
    const r = this.current;
    if (!r) {
      $('#route-title').html('');
      $('#cluster-title').html('');
      $('#destination-facility-title').html('');
      $('#destination-facility-name').html('');
      $('#extracted-total').text('0');
      $('#verified-total').text('0');
      $('#progress-bar').css('width', '0%').text('0%');
      $('#conferidos-list, #faltantes-list, #fora-rota-list, #duplicados-list').html('');
      return;
    }

    const esc = (x) => this.escHtml(x);
    $('#route-title').html(`ROTA: <strong>${esc(r.routeId)}</strong>`);
    $('#cluster-title').html(r.cluster ? `CLUSTER: <strong>${esc(r.cluster)}</strong>` : '');
    $('#destination-facility-title').html(r.destinationFacilityId ? `<strong>XPT:</strong> ${esc(r.destinationFacilityId)}` : '');
    $('#destination-facility-name').html(r.destinationFacilityName ? `<strong>DESTINO:</strong> ${esc(r.destinationFacilityName)}` : '');

    $('#extracted-total').text(r.totalInicial || r.ids.size);
    $('#verified-total').text(r.conferidos.size);

    this.atualizarListas();
  },

  atualizarProgresso() {
    const r = this.current;
    if (!r) return;

    const total = r.totalInicial || (r.ids.size || (r.conferidos.size + r.faltantes.size));
    const perc = total ? (r.conferidos.size / total) * 100 : 0;

    $('#progress-bar').css('width', perc + '%').text(Math.floor(perc) + '%');
  },

  atualizarListas() {
    const r = this.current;
    if (!r) return;

    $('#conferidos-list').html(
      `<h6>Conferidos (<span class='badge badge-success'>${r.conferidos.size}</span>)</h6>` +
      Array.from(r.conferidos).map(id => `<li class='list-group-item list-group-item-success'>${this.escHtml(id)}</li>`).join('')
    );

    $('#faltantes-list').html(
      `<h6>Faltantes (<span class='badge badge-danger'>${r.faltantes.size}</span>)</h6>` +
      Array.from(r.faltantes).map(id => `<li class='list-group-item list-group-item-danger'>${this.escHtml(id)}</li>`).join('')
    );

    $('#fora-rota-list').html(
      `<h6>Fora de Rota (<span class='badge badge-warning'>${r.foraDeRota.size}</span>)</h6>` +
      Array.from(r.foraDeRota).map(id => {
        const correct = this.findCorrectRouteForId(id);
        if (correct && String(correct) !== String(this.currentRouteId)) {
          const rr = this.routes.get(String(correct));
          const cl = rr && rr.cluster ? String(rr.cluster).trim() : '';
          const extra = cl
            ? ` <small class="text-muted">(CLUSTER ${this.escHtml(cl)})</small>`
            : ` <small class="text-muted">(cluster desconhecido)</small>`;
          return `<li class='list-group-item list-group-item-warning'>${this.escHtml(id)}${extra}</li>`;
        }
        return `<li class='list-group-item list-group-item-warning'>${this.escHtml(id)}</li>`;
      }).join('')
    );

    $('#duplicados-list').html(
      `<h6>Duplicados (<span class='badge badge-secondary'>${r.duplicados.size}</span>)</h6>` +
      Array.from(r.duplicados.entries())
        .map(([id, count]) => `<li class='list-group-item list-group-item-secondary'>${this.escHtml(id)} <span class="badge badge-dark ml-2">${this.dupCount(count)}x</span></li>`)
        .join('')
    );

    $('#verified-total').text(r.conferidos.size);
    this.atualizarProgresso();
  },

  deleteRoute(routeId) {
    if (!routeId) return;
    const rid = String(routeId);

    if (!this.deletedRoutes) this.deletedRoutes = new Map();
    const now = Date.now();
    this.deletedRoutes.set(rid, now);
    if (this.revivedRoutes) this.revivedRoutes.delete(rid);

    this.routes.delete(rid);
    if (this.currentRouteId === rid) this.currentRouteId = null;

    this.lastRoutesSignature = '';
    this.saveToStorage(this.workDay);
    this.markDirty('excluir rota');

    this.renderRoutesSelects();
    this.refreshUIFromCurrent();
    this.renderAcompanhamento();
  },

  clearAllRoutes() {
    if (!this.deletedRoutes) this.deletedRoutes = new Map();
    const now = Date.now();

    for (const rid of this.routes.keys()) {
      this.deletedRoutes.set(String(rid), now);
      if (this.revivedRoutes) this.revivedRoutes.delete(String(rid));
    }

    this.routes.clear();
    this.currentRouteId = null;
    this.lastRoutesSignature = '';

    this.saveToStorage(this.workDay, { syncCloud: true });
    this.markDirty('limpar dia');

    this.renderRoutesSelects();
    this.refreshUIFromCurrent();
    this.renderAcompanhamento();
  },

  // =======================
  // Normalização / Som
  // =======================
  normalizarCodigo(raw) {
    if (!raw) return null;
    let s = String(raw).trim().replace(/[\u0000-\u001F\u007F-\u009F]/g, '');

    let m = s.match(/(4\d{10})/);
    if (m) return m[1];

    m = s.replace(/\D/g, '').match(/(\d{11,})/);
    if (m) return m[1].slice(0, 11);

    return null;
  },

  // =======================
  // Leitura inteligente
  // =======================
  parseScanPayload(raw) {
    const cleaned = String(raw || '').trim();
    if (!cleaned) return { kind: 'empty' };

    const firstLine = cleaned.split(/\?\n/)[0].trim();
    if (firstLine.includes('^') && firstLine.includes('Ç')) {
      const kv = this.parseCaretKV(firstLine);

      if (kv.license_plate) {
        const plateKey = String(kv.license_plate).trim().toUpperCase();
        const plateObj = {
          id: kv.id || '',
          carrier_id: kv.carrier_id || '',
          carrier_name: kv.carrier_name || '',
          license_plate: plateKey,
          vehicle_type_description: kv.vehicle_type_description || ''
        };
        const jsonText = JSON.stringify(plateObj);
        return {
          kind: 'plate',
          plateKey,
          plate: {
            raw: firstLine,
            jsonText,
            license_plate: plateKey,
            carrier_name: plateObj.carrier_name || '',
            vehicle_type_description: plateObj.vehicle_type_description || '',
            carrier_id: plateObj.carrier_id || '',
            id: plateObj.id || ''
          }
        };
      }

      if (kv.container_id || kv.assignment) {
        let assignment = kv.assignment ? String(kv.assignment).trim() : '';
        assignment = this.normalizeCluster(assignment);

        const obj = {
          container_id: kv.container_id ? Number(kv.container_id) : undefined,
          facility_id: kv.facility_id || '',
          assignment: assignment
        };
        Object.keys(obj).forEach(k => obj[k] === undefined && delete obj[k]);

        const routeKey = assignment
          ? `assignment:${assignment}`
          : (kv.container_id ? `container:${kv.container_id}` : `caret:${firstLine}`);

        const jsonText = JSON.stringify(obj);

        return {
          kind: 'routeqr',
          routeKey,
          routeIdCandidate: '',
          route: { raw: firstLine, obj, jsonText }
        };
      }
    }

    if (cleaned.startsWith('{') && cleaned.endsWith('}')) {
      try {
        const obj = JSON.parse(cleaned);

        if (obj && typeof obj === 'object' && obj.license_plate) {
          const plateKey = String(obj.license_plate || '').trim().toUpperCase();
          return {
            kind: 'plate',
            plateKey,
            plate: {
              raw: cleaned,
              license_plate: plateKey,
              carrier_name: obj.carrier_name || '',
              vehicle_type_description: obj.vehicle_type_description || '',
              carrier_id: obj.carrier_id || '',
              vehicle_type_id: obj.vehicle_type_id || '',
              id: obj.id || ''
            }
          };
        }

        if (obj && typeof obj === 'object' && (obj.container_id || obj.assignment || obj.routeId || obj.route_id)) {
          const candidate = obj.routeId || obj.route_id || obj.container_id || obj.assignment;
          const routeIdCandidate = candidate != null ? String(candidate) : '';
          const routeKey = (obj.container_id != null)
            ? `container:${obj.container_id}`
            : (obj.assignment != null)
              ? `assignment:${obj.assignment}`
              : (routeIdCandidate ? `route:${routeIdCandidate}` : `json:${cleaned}`);

          return {
            kind: 'routeqr',
            routeKey,
            routeIdCandidate,
            route: { raw: cleaned, obj }
          };
        }
      } catch (e) {}
    }

    const shipmentId = this.normalizarCodigo(cleaned);
    if (shipmentId) return { kind: 'shipment', shipmentId };

    const plateLike = cleaned.toUpperCase().replace(/[^A-Z0-9]/g, '');
    if (/^[A-Z]{3}\d[A-Z]\d{2}$/.test(plateLike) || /^[A-Z]{3}\d{4}$/.test(plateLike)) {
      return {
        kind: 'plate',
        plateKey: plateLike,
        plate: { raw: cleaned, license_plate: plateLike, carrier_name: '', vehicle_type_description: '' }
      };
    }

    const digits = cleaned.replace(/\D/g, '');
    if (digits.length >= 4) {
      const routeIdCandidate = digits;
      return {
        kind: 'routeqr',
        routeKey: `route:${routeIdCandidate}`,
        routeIdCandidate,
        route: { raw: cleaned, obj: null }
      };
    }

    return { kind: 'unknown' };
  },

  ensurePlate(plateInfo) {
    const now = Date.now();
    const key = String(plateInfo.license_plate || '').trim().toUpperCase();
    if (!key) return null;

    if (!this.carretas.plates.has(key)) {
      this.carretas.plates.set(key, {
        raw: plateInfo.raw || '',
        jsonText: plateInfo.jsonText || '',
        tsScan: now,
        license_plate: key,
        carrier_name: plateInfo.carrier_name || '',
        vehicle_type_description: plateInfo.vehicle_type_description || '',
        routes: new Set(),
        tsFirst: now,
        tsLast: now
      });
    } else {
      const p = this.carretas.plates.get(key);
      p.tsLast = now;
      if (plateInfo.raw) p.raw = plateInfo.raw;
      if (plateInfo.jsonText) p.jsonText = plateInfo.jsonText;
      p.tsScan = now;
      if (plateInfo.carrier_name) p.carrier_name = plateInfo.carrier_name;
      if (plateInfo.vehicle_type_description) p.vehicle_type_description = plateInfo.vehicle_type_description;
    }
    return key;
  },

  vincularRouteQrNaPlaca(routeKey, routeRaw, plateKey, routeIdCandidate = '') {
    if (!routeKey || !plateKey) return false;

    const plate = this.carretas.plates.get(plateKey);
    if (!plate) return false;

    plate.routes.add(routeKey);
    this.carretas.routeToPlate.set(routeKey, plateKey);

    const rawStr = (typeof routeRaw === 'string')
      ? routeRaw
      : ((routeRaw && routeRaw.raw) ? String(routeRaw.raw) : '');

    const jsonText = (routeRaw && typeof routeRaw === 'object' && routeRaw.jsonText)
      ? String(routeRaw.jsonText)
      : '';

    if (rawStr) this.carretas.routesRaw.set(routeKey, rawStr);
    if (jsonText) this.carretas.routesJson.set(routeKey, jsonText);
    this.carretas.routesTs.set(routeKey, Date.now());

    const assignMatch = String(routeKey).match(/^assignment:(.+)$/);
    const clusterCandidate = this.normalizeCluster(assignMatch?.[1] || '');

    let linked = 0;

    if (clusterCandidate) {
      for (const r of this.routes.values()) {
        const c = this.normalizeCluster(r.cluster);
        if (c && clusterCandidate && c === clusterCandidate) {
          r.plateKey = plateKey;
          r.plateRaw = plate.jsonText || plate.raw || '';
          r.plateLicense = plate.license_plate || plateKey;
          r.routeQrKey = routeKey;
          r.routeQrRaw = (typeof routeRaw === 'string' ? routeRaw : (routeRaw && routeRaw.raw) ? routeRaw.raw : '') || '';
          const _json = (routeRaw && routeRaw.jsonText) ? routeRaw.jsonText : '';
          if (_json) r.routeQrRaw = _json;
          r.plateScanTs = Number(plate.tsLast || Date.now());
          r.routeQrScanTs = Date.now();
          r.plateUpdatedAt = Date.now();
          linked++;
        }
      }
    }

    if (!linked) {
      const candidateIds = [];
      if (routeIdCandidate) candidateIds.push(String(routeIdCandidate));
      const m = String(routeKey).match(/^route:(.+)$/);
      if (m && m[1]) candidateIds.push(String(m[1]));

      for (const cid of candidateIds) {
        if (this.routes.has(String(cid))) {
          const r = this.routes.get(String(cid));
          r.plateKey = plateKey;
          r.plateRaw = plate.jsonText || plate.raw || '';
          r.plateLicense = plate.license_plate || plateKey;
          r.routeQrKey = routeKey;
          r.routeQrRaw = (typeof routeRaw === 'string' ? routeRaw : (routeRaw && routeRaw.raw) ? routeRaw.raw : '') || '';
          const _json = (routeRaw && routeRaw.jsonText) ? routeRaw.jsonText : '';
          if (_json) r.routeQrRaw = _json;
          r.plateScanTs = Number(plate.tsLast || Date.now());
          r.routeQrScanTs = Date.now();
          r.plateUpdatedAt = Date.now();
          linked++;
        }
      }
    }

    this.saveToStorage(this.workDay);
    this.markDirty('carreta');
    return true;
  },

  checkLinksForCurrentPlate() {
    const plateKey = this.carretas.currentPlateKey;
    if (!plateKey) return alert('Nenhuma placa ativa.');

    const p = this.carretas.plates.get(plateKey);
    if (!p) return alert('Placa ativa não encontrada na memória.');

    const routes = Array.from(p.routes || []);
    routes.sort((a, b) => String(a).localeCompare(String(b)));

    const clustersImportados = new Set(
      Array.from(this.routes.values()).map(r => this.normalizeCluster(r.cluster)).filter(Boolean)
    );

    const detalhes = routes.map(rk => {
      const m = String(rk).match(/^assignment:(.+)$/);
      const cl = this.normalizeCluster(m?.[1] || '');
      const ok = cl && clustersImportados.has(cl);
      return `- ${rk}  => cluster: ${cl || '(vazio)'}  ${ok ? '[OK]' : '[NÃO ENCONTRADO NAS ROTAS IMPORTADAS]'}`;
    });

    const msg =
      `PLACA ATIVA: ${plateKey}\n` +
      `ROTAS VINCULADAS: ${routes.length}\n\n` +
      (detalhes.length ? detalhes.join('\n') : '(nenhuma rota vinculada)\n') +
      `\n\nObs: [OK] significa que existe rota importada com cluster igual ao do QR.`;

    alert(msg);
  },

  clearBipagemForPlate(plateKeyRaw) {
    const plateKey = String(plateKeyRaw || '').trim().toUpperCase();
    if (!plateKey) return alert('Informe uma placa válida.');

    const p = this.carretas.plates.get(plateKey);
    if (!p) return alert('Essa placa não está carregada/vinculada.');

    const routeKeys = Array.from(p.routes || []);
    for (const rk of routeKeys) {
      this.carretas.routeToPlate.delete(rk);
      this.carretas.routesRaw.delete(rk);
      this.carretas.routesJson.delete(rk);
      this.carretas.routesTs.delete(rk);
    }

    p.routes = new Set();
    p.clearedAt = Date.now();

    for (const r of this.routes.values()) {
      if ((r.plateKey || '').toUpperCase() === plateKey) {
        r.plateKey = '';
        r.plateRaw = '';
        r.plateLicense = '';
        r.routeQrKey = '';
        r.routeQrRaw = '';
        r.plateScanTs = 0;
        r.routeQrScanTs = 0;
        r.plateUpdatedAt = Date.now();
      }
    }

    this.saveToStorage(this.workDay);
    this.markDirty('excluir bipagem placa');
    this.renderCarretaUI();
    this.renderPatioGeral();
    this.renderAcompanhamento();

    alert(`Bipagem/vínculos removidos para a placa ${plateKey}.`);
  },

  renderCarretaUI() {
    const plateKey = this.carretas.currentPlateKey;
    const $cur = $('#carreta-current');
    const $list = $('#carreta-routes');
    const $sum = $('#carreta-summary');

    if (!plateKey) {
      $cur.html('<span class="text-muted">Nenhuma placa ativa</span>');
      $list.html('<li class="list-group-item text-muted">bipe uma placa para começar</li>');
      $sum.text('');
      this.renderCarretaProgress();
      return;
    }

    const p = this.carretas.plates.get(plateKey);
    if (!p) return;

    const meta = [];
    const esc = (x) => this.escHtml(x);
    meta.push(`<strong>${esc(p.license_plate)}</strong>`);
    if (p.vehicle_type_description) meta.push(`<span class="text-muted">(${esc(p.vehicle_type_description)})</span>`);
    if (p.carrier_name) meta.push(`<span class="text-muted">• ${esc(p.carrier_name)}</span>`);
    $cur.html(meta.join(' '));

    const routesArr = Array.from(p.routes);
    routesArr.sort((a, b) => String(a).localeCompare(String(b)));

    $list.html(
      routesArr.map(rk => `<li class="list-group-item">${this.escHtml(rk)}</li>`).join('') ||
      '<li class="list-group-item text-muted">sem rotas nessa placa</li>'
    );

    $sum.text(`${routesArr.length} rota(s) vinculada(s)`);
    this.renderCarretaProgress();
  },

  renderCarretaProgress() {
    const $label = $('#carreta-progress-label');
    const $pct = $('#carreta-progress-percent');
    const $bar = $('#carreta-progress-bar');
    const $missing = $('#carreta-missing-list');
    const $extra = $('#carreta-extra-list');

    if (!$label.length || !$pct.length || !$bar.length) return;

    const plateKey = this.carretas.currentPlateKey;

    const setUi = (done, total) => {
      const pctVal = total > 0 ? Math.round((done / total) * 100) : 0;
      $label.text(`${done}/${total}`);
      $pct.text(`${pctVal}%`);
      $bar.css('width', `${pctVal}%`);
      $bar.attr('aria-valuenow', String(pctVal));
    };

    const expected = Array.from(this.routes.keys()).map(String);
    expected.sort((a, b) => (Number(a) - Number(b)) || String(a).localeCompare(String(b)));

    if (!plateKey) {
      setUi(0, expected.length);
      $missing.html('<li class="list-group-item text-muted">bipe uma placa para ver o acompanhamento</li>');
      $extra.html('<li class="list-group-item text-muted">—</li>');
      return;
    }

    const plate = this.carretas.plates.get(plateKey);

    const linked = expected.filter(routeId => {
      const r = this.routes.get(String(routeId));
      return r && String(r.plateKey || '') === String(plateKey);
    });

    const missing = expected.filter(routeId => !linked.includes(routeId));

    const clustersImportados = new Set(
      Array.from(this.routes.values()).map(r => this.normalizeCluster(r.cluster)).filter(Boolean)
    );

    const plateQrRoutes = plate ? Array.from(plate.routes || []) : [];
    plateQrRoutes.sort((a, b) => String(a).localeCompare(String(b)));

    const extra = plateQrRoutes.filter(rk => {
      const m = String(rk).match(/^assignment:(.+)$/);
      const cl = this.normalizeCluster(m?.[1] || '');
      return cl && !clustersImportados.has(cl);
    });

    setUi(linked.length, expected.length);

    const fmtExpected = (routeId) => {
      const r = this.routes.get(String(routeId));
      const cl = r && r.cluster ? String(r.cluster).trim() : '';
      const fac = r && r.destinationFacilityName ? String(r.destinationFacilityName).trim() : '';
      const parts = [`${routeId}`];
      if (cl) parts.push(`CLUSTER ${cl}`);
      if (fac) parts.push(fac);
      return parts.join(' • ');
    };

    const fmtExtra = (rk) => {
      const m = String(rk).match(/^assignment:(.+)$/);
      return m ? m[1] : rk;
    };

    $missing.html(
      missing.map(routeId => `<li class="list-group-item">${this.escHtml(fmtExpected(routeId))}</li>`).join('') ||
      '<li class="list-group-item text-muted">nada faltando 🎉</li>'
    );

    $extra.html(
      extra.map(rk => `<li class="list-group-item">${this.escHtml(fmtExtra(rk))}</li>`).join('') ||
      '<li class="list-group-item text-muted">—</li>'
    );
  },

  renderPatioGeral() {
    // placeholder seguro caso não exista implementação específica
  },

  playAlertSound() {
    try {
      const audio = new Audio('mixkit-alarm-tone-996-_1_.mp3');
      audio.play().catch(() => {});
    } catch {}
  },

  // =======================
  // Fora de rota inteligente
  // =======================
  findCorrectRouteForId(id) {
    for (const [rid, r] of this.routes.entries()) {
      if (r.ids && r.ids.has(id)) return String(rid);
    }
    for (const [rid, r] of this.routes.entries()) {
      if (r.faltantes && r.faltantes.has(id)) return String(rid);
    }
    for (const [rid, r] of this.routes.entries()) {
      if (r.conferidos && r.conferidos.has(id)) return String(rid);
    }
    return null;
  },

  cleanupIdFromOtherRoutes(id, targetRouteId) {
    const target = String(targetRouteId);

    for (const [rid, r] of this.routes.entries()) {
      if (String(rid) === target) continue;

      let changed = false;

      if (r.foraDeRota && r.foraDeRota.has(id)) {
        r.foraDeRota.delete(id);
        changed = true;
      }

      if (r.duplicados && r.duplicados.has(id)) {
        r.duplicados.delete(id);
        changed = true;
      }

      if (changed) {
        const stillRelevant =
          (r.conferidos && r.conferidos.has(id)) ||
          (r.faltantes && r.faltantes.has(id)) ||
          (r.ids && r.ids.has(id)) ||
          (r.foraDeRota && r.foraDeRota.has(id)) ||
          (r.duplicados && r.duplicados.has(id));

        if (!stillRelevant && r.timestamps) r.timestamps.delete(id);
      }
    }
  },

  // =======================
  // Conferência
  // =======================
  conferirId(codigo) {
    const r = this.current;
    if (!r || !codigo) return;

    const now = Date.now();

    const correctRouteId = this.findCorrectRouteForId(codigo);
    const isCorrectHere = correctRouteId && String(correctRouteId) === String(this.currentRouteId);

    if (isCorrectHere) {
      if (!r.conferidos.has(codigo)) {
        if (r.faltantes && r.faltantes.has(codigo)) r.faltantes.delete(codigo);
        if (r.foraDeRota && r.foraDeRota.has(codigo)) r.foraDeRota.delete(codigo);
        r.conferidos.add(codigo);
        r.timestamps.set(codigo, now);

        this.cleanupIdFromOtherRoutes(codigo, this.currentRouteId);

        $('#barcode-input').val('').focus();
        this.pushEvent({ ts: now, type: 'ok', code: codigo, currentRouteId: this.currentRouteId, correctRouteId });
        this.markDirty('bipagem');

        this.saveToStorage(this.workDay);
        this.atualizarListas();
        return;
      }

      this.cleanupIdFromOtherRoutes(codigo, this.currentRouteId);
    }

    if (r.conferidos.has(codigo) || r.foraDeRota.has(codigo)) {
      const count = this.dupCount(r.duplicados.get(codigo)) || 1;
      r.duplicados.set(codigo, count + 1);
      r.timestamps.set(codigo, now);

      if (!this.viaCsv) this.playAlertSound();
      $('#barcode-input').val('').focus();

      this.pushEvent({ ts: now, type: 'dup', code: codigo, currentRouteId: this.currentRouteId, correctRouteId });
      this.markDirty('bipagem');

      this.saveToStorage(this.workDay);
      this.atualizarListas();
      return;
    }

    if (r.faltantes.has(codigo)) {
      r.faltantes.delete(codigo);
      r.conferidos.add(codigo);
      r.timestamps.set(codigo, now);

      this.cleanupIdFromOtherRoutes(codigo, this.currentRouteId);

      $('#barcode-input').val('').focus();
      this.pushEvent({ ts: now, type: 'ok', code: codigo, currentRouteId: this.currentRouteId, correctRouteId });
      this.markDirty('bipagem');

      this.saveToStorage(this.workDay);
      this.atualizarListas();
      return;
    }

    r.foraDeRota.add(codigo);
    r.timestamps.set(codigo, now);
    if (!this.viaCsv) this.playAlertSound();

    $('#barcode-input').val('').focus();
    this.pushEvent({ ts: now, type: 'fora', code: codigo, currentRouteId: this.currentRouteId, correctRouteId });
    this.markDirty('bipagem');

    this.saveToStorage(this.workDay);
    this.atualizarListas();
  },

  // =======================
  // Importação HTML
  // =======================
  importRoutesFromHtml(rawHtml) {
    const html = String(rawHtml || '').replace(/<[^>]+>/g, ' ');

    const idxs = [];
    for (const m of html.matchAll(/"routeId":(\d+)/g)) idxs.push(m.index);

    if (!idxs.length) {
      alert('Não encontrei nenhum "routeId" no HTML.');
      return 0;
    }

    const blocks = [];
    for (let i = 0; i < idxs.length; i++) {
      const start = idxs[i];
      const end = i + 1 < idxs.length ? idxs[i + 1] : html.length;
      blocks.push(html.slice(start, end));
    }

    let imported = 0;

    for (const block of blocks) {
      const routeMatch = /"routeId":(\d+)/.exec(block);
      if (!routeMatch) continue;

      const routeId = String(routeMatch[1]);

      if (this.deletedRoutes?.has(routeId)) {
        this.deletedRoutes.delete(routeId);
        if (!this.revivedRoutes) this.revivedRoutes = new Map();
        this.revivedRoutes.set(routeId, Date.now());
      }

      const route = this.routes.get(routeId) || this.makeEmptyRoute(routeId);

      const clusterMatch = /"cluster":"([^"]+)"/.exec(block);
      if (clusterMatch) route.cluster = this.normalizeCluster(clusterMatch[1]);

      const facMatch = /"destinationFacilityId":"([^"]+)","name":"([^"]+)"/.exec(block);
      if (facMatch) {
        route.destinationFacilityId = facMatch[1];
        route.destinationFacilityName = facMatch[2];
      }

      const idsExtraidos = new Set();
      const regexId = /"id":\s*(\d{11})/g;
      let mId;
      while ((mId = regexId.exec(block)) !== null) {
        const shipmentId = mId[1];
        if (/^4\d{10}$/.test(shipmentId)) idsExtraidos.add(shipmentId);
      }

      if (!idsExtraidos.size) continue;

      for (const id of idsExtraidos) {
        route.ids.add(id);
        if (!route.conferidos.has(id)) route.faltantes.add(id);
      }

      route.totalInicial = route.ids.size;
      this.routes.set(routeId, route);
      imported++;
    }

    this.lastRoutesSignature = '';
    this.saveToStorage(this.workDay);
    this.markDirty('import HTML');

    this.renderRoutesSelects();
    this.renderAcompanhamento();

    if (imported) this.currentRouteId = String(this.routes.keys().next().value);
    this.refreshUIFromCurrent();
    this.atualizarListas();

    return imported;
  },

  // =======================
  // Export helpers
  // =======================
  getIdsForExportByTimestamp(r) {
    if (!r) return [];
    const set = new Set([
      ...Array.from(r.conferidos || []),
      ...Array.from(r.foraDeRota || []),
      ...Array.from((r.duplicados || new Map()).keys())
    ]);
    const ids = Array.from(set);

    ids.sort((a, b) => {
      const ta = r.timestamps?.get(a) ? Number(r.timestamps.get(a)) : 0;
      const tb = r.timestamps?.get(b) ? Number(r.timestamps.get(b)) : 0;
      return (ta - tb) || String(a).localeCompare(String(b));
    });
    return ids;
  },

  csvEscape(v) {
    const s = (v == null) ? '' : String(v);
    return '"' + s.replace(/"/g, '""') + '"';
  },

  buildScannerCsvHeader() {
    return '"date","time","time_zone","format","text","notes","favorite","date_utc","time_utc","metadata"';
  },

  buildScannerCsvRow(dt, format, text, metadata = '') {
    const d = (dt instanceof Date) ? dt : new Date(Number(dt || Date.now()));
    const pad2 = n => String(n).padStart(2, '0');

    const date = `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
    const time = `${pad2(d.getHours())}:${pad2(d.getMinutes())}:${pad2(d.getSeconds())}`;

    const iso = d.toISOString();
    const dateUtc = iso.slice(0, 10);
    const timeUtc = iso.split('T')[1].split('.')[0];

    const tzLabel = 'Horário Padrão do Amazonas';

    const esc = (v) => '"' + String(v ?? '').replace(/"/g, '""') + '"';

    return [
      esc(date),
      esc(time),
      esc(tzLabel),
      esc(format || 'QR Code'),
      esc(text || ''),
      esc(''),
      esc('0'),
      esc(dateUtc),
      esc(timeUtc),
      esc(metadata || '')
    ].join(',');
  },

  buildScannerCsvLinesForRoute(r) {
    const lines = [];
    if (!r) return lines;

    const ids = this.getIdsForExportByTimestamp(r);

    const firstIdTs = ids.length ? (r.timestamps?.get(ids[0]) || Date.now()) : Date.now();
    const plateTs = r.plateScanTs || firstIdTs;
    const routeTs = r.routeQrScanTs || (plateTs ? (Number(plateTs) + 1) : firstIdTs);

    let plateText = '';
    if (r.plateKey) {
      const p = this.carretas?.plates?.get(r.plateKey);
      if (p && p.jsonText) {
        plateText = String(p.jsonText);
      } else {
        plateText = JSON.stringify({
          id: (p && p.id) ? Number(p.id) : undefined,
          carrier_id: (p && p.carrier_id) ? Number(p.carrier_id) : undefined,
          carrier_name: (p && p.carrier_name) ? String(p.carrier_name) : undefined,
          license_plate: String((p && p.license_plate) || r.plateLicense || r.plateKey),
          vehicle_type_description: (p && p.vehicle_type_description) ? String(p.vehicle_type_description) : undefined,
          vehicle_type_id: (p && p.vehicle_type_id) ? Number(p.vehicle_type_id) : undefined,
          tracking_provider_ids: []
        }, (k, v) => (v === undefined ? undefined : v));
        plateText = plateText.replace(/,\s*"(?:id|carrier_id|carrier_name|vehicle_type_description|vehicle_type_id)"\s*:\s*null/g, '');
      }
    }

    let routeText = '';
    if (r.routeQrKey) {
      const routeJson = this.carretas?.routesJson?.get(r.routeQrKey);
      if (routeJson) {
        routeText = String(routeJson);
      } else if (r.routeQrRaw && String(r.routeQrRaw).trim().startsWith('{')) {
        routeText = String(r.routeQrRaw).trim();
      } else {
        routeText = JSON.stringify({
          container_id: (r.container_id != null) ? Number(r.container_id) : undefined,
          facility_id: (r.destinationFacilityId || ''),
          assignment: String(r.routeQrKey).replace(/^assignment:/, '')
        }, (k, v) => (v === undefined ? undefined : v));
      }
    }

    if (plateText) lines.push(this.buildScannerCsvRow(plateTs, 'QR Code', plateText, ''));
    if (routeText) lines.push(this.buildScannerCsvRow(routeTs, 'QR Code', routeText, ''));

    for (const id of ids) {
      const ts = r.timestamps?.get(id) || Date.now();
      const payload = JSON.stringify({ id: String(id), t: 'lm' });
      lines.push(this.buildScannerCsvRow(ts, 'QR Code', payload, ''));
    }

    return lines;
  },

  exportRotaAtualCsvComPlacaERota() {
    const r = this.current;
    if (!r) return alert('Nenhuma rota selecionada.');

    if (!r.plateKey || !r.routeQrKey) {
      return alert('Esta rota ainda não está vinculada a uma PLACA e a um QR de ROTA. Use a tela da CARRETA primeiro.');
    }

    const lines = [];
    lines.push(this.buildScannerCsvHeader());

    const body = this.buildScannerCsvLinesForRoute(r);
    if (!body.length) return alert('Nenhum registro para exportar.');

    lines.push(...body);

    const csv = lines.join('\r\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);

    const cluster = (r.cluster || 'semCluster').replace(/[^\w\-]+/g, '_');
    link.download = `RECEBIMENTO_${this.workDay || this.todayLocalISO()}_${cluster}_ROTA_${r.routeId}_PLACA.csv`;
    link.click();
  },

  exportTodasRotasCsvComPlacaERota() {
    if (!this.routes || this.routes.size === 0) return alert('Não há rotas salvas para exportar.');

    const plateGroups = new Map();
    for (const r of this.routes.values()) {
      if (!r.plateKey || !r.routeQrKey) continue;
      if (!plateGroups.has(r.plateKey)) plateGroups.set(r.plateKey, []);
      plateGroups.get(r.plateKey).push(r);
    }

    if (!plateGroups.size) {
      return alert('Nenhuma rota está vinculada a PLACA/QR de rota. Use a tela da CARRETA primeiro.');
    }

    const lines = [];
    lines.push(this.buildScannerCsvHeader());

    const plateKeys = Array.from(plateGroups.keys()).sort((a, b) => String(a).localeCompare(String(b)));

    for (const plateKey of plateKeys) {
      const routesArr = plateGroups.get(plateKey) || [];
      routesArr.sort((a, b) => {
        const ca = String(a.cluster || '').localeCompare(String(b.cluster || ''));
        if (ca !== 0) return ca;
        return String(a.routeId || '').localeCompare(String(b.routeId || ''));
      });

      for (const r of routesArr) {
        const body = this.buildScannerCsvLinesForRoute(r);
        if (body.length) lines.push(...body);
      }
    }

    if (lines.length <= 1) return alert('Nenhum registro para exportar.');

    const csv = lines.join('\r\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);

    const now = new Date();
    const stamp = `${now.getFullYear()}-${this.pad2(now.getMonth() + 1)}-${this.pad2(now.getDate())}_${this.pad2(now.getHours())}${this.pad2(now.getMinutes())}`;
    link.download = `RECEBIMENTO_${this.workDay || this.todayLocalISO()}_PLACAS_ROTAS_${stamp}.csv`;
    link.click();
  },

  exportMapaCarretasCsv() {
    if (!this.carretas.plates || this.carretas.plates.size === 0) {
      alert('Nenhuma placa/rota vinculada ainda.');
      return;
    }

    const header = 'plate,carrier,vehicle_type,route_qr_key,route_qr_raw';
    const linhas = [];

    for (const [plateKey, p] of this.carretas.plates.entries()) {
      const carrier = (p.carrier_name || '').replace(/,/g, ' ');
      const vt = (p.vehicle_type_description || '').replace(/,/g, ' ');
      for (const rk of Array.from(p.routes)) {
        const raw = (this.carretas.routesRaw.get(rk) || '').replace(/\r?\n/g, ' ');
        const rawEsc = `"${String(raw).replace(/"/g, '""')}"`;
        linhas.push(`${plateKey},${carrier},${vt},${rk},${rawEsc}`);
      }
    }

    const conteudo = [header, ...linhas].join('\r\n');
    const blob = new Blob([conteudo], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);

    const now = new Date();
    const stamp = `${now.getFullYear()}-${this.pad2(now.getMonth() + 1)}-${this.pad2(now.getDate())}_${this.pad2(now.getHours())}${this.pad2(now.getMinutes())}`;
    link.download = `mapa_carretas_${this.workDay || this.todayLocalISO()}_${stamp}.csv`;
    link.click();
  },

  exportRotaAtualCsvPadrao() {
    const r = this.current;
    if (!r) {
      alert('Nenhuma rota selecionada.');
      return;
    }

    const all = [
      ...Array.from(r.conferidos),
      ...Array.from(r.foraDeRota),
      ...Array.from(r.duplicados.keys())
    ];

    if (all.length === 0) {
      alert('Nenhum ID para exportar.');
      return;
    }

    const parseDateSafe = (value) => {
      if (!value) return new Date();
      if (value instanceof Date) return value;
      if (typeof value === 'number') return new Date(value);
      if (typeof value === 'string') {
        if (/^\d{4}-\d{2}-\d{2}T/.test(value)) {
          const d = new Date(value);
          if (!isNaN(d.getTime())) return d;
        }
        const m = value.match(/^(\d{2})\/(\d{2})\/(\d{4})[ T](\d{2}):(\d{2})(?::(\d{2}))?$/);
        if (m) {
          const [, dd, mm, yyyy, HH, MM, SS = '00'] = m;
          const iso = `${yyyy}-${mm}-${dd}T${HH}:${MM}:${SS}`;
          const d = new Date(iso);
          if (!isNaN(d.getTime())) return d;
        }
        if (/^\d{13}$/.test(value)) return new Date(Number(value));
        const d = new Date(value);
        if (!isNaN(d.getTime())) return d;
      }
      return new Date();
    };

    const zona = 'Horário Padrão de Brasília';
    const header = 'date,time,time_zone,format,text,notes,favorite,date_utc,time_utc,metadata,duplicates';

    const linhas = all.map(id => {
      const lidaEm = parseDateSafe(r.timestamps.get(id));
      const pad2 = n => String(n).padStart(2, '0');
      const date = `${lidaEm.getFullYear()}-${pad2(lidaEm.getMonth() + 1)}-${pad2(lidaEm.getDate())}`;
      const time = `${pad2(lidaEm.getHours())}:${pad2(lidaEm.getMinutes())}:${pad2(lidaEm.getSeconds())}`;

      const dateUtc = lidaEm.toISOString().slice(0, 10);
      const timeUtc = lidaEm.toISOString().split('T')[1].split('.')[0];
      const dupCount = r.duplicados.get(id) ? (Number(r.duplicados.get(id)) - 1) : 0;

      return `${date},${time},${zona},Code 128,${id},,0,${dateUtc},${timeUtc},,${dupCount}`;
    });

    const conteudo = [header, ...linhas].join('\r\n');
    const blob = new Blob([conteudo], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);

    const cluster = (r.cluster || 'semCluster').replace(/[^\w\-]+/g, '_');
    const rota = (r.routeId || 'semRota').replace(/[^\w\-]+/g, '_');

    link.download = `${cluster}_${rota}_padrao.csv`;
    link.click();
  },

  exportTodasRotasXlsx() {
    if (typeof XLSX === 'undefined') {
      alert('Biblioteca XLSX não carregou. Verifique o script do SheetJS no HTML.');
      return;
    }
    if (!this.routes || this.routes.size === 0) {
      alert('Não há rotas salvas para exportar.');
      return;
    }

    const routesSorted = Array.from(this.routes.values())
      .sort((a, b) => String(a.routeId).localeCompare(String(b.routeId)));

    const cols = routesSorted.map((r) => {
      const routeId = String(r.routeId || '');
      const cluster = String(r.cluster || '').trim();
      const header = cluster ? `${routeId}-${cluster}` : routeId;

      const ids = Array.from(r.conferidos || []);

      ids.sort((x, y) => {
        const tx = r.timestamps?.get(x) ? Number(r.timestamps.get(x)) : 0;
        const ty = r.timestamps?.get(y) ? Number(r.timestamps.get(y)) : 0;
        return (tx - ty) || String(x).localeCompare(String(y));
      });

      return { header, ids };
    });

    const maxLen = cols.reduce((m, c) => Math.max(m, c.ids.length), 0);

    const aoa = [];
    aoa.push(cols.map(c => c.header || 'ROTA'));
    for (let i = 0; i < maxLen; i++) {
      aoa.push(cols.map(c => c.ids[i] || ''));
    }

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(aoa);

    ws['!freeze'] = { xSplit: 0, ySplit: 1 };
    ws['!cols'] = cols.map(() => ({ wch: 18 }));

    XLSX.utils.book_append_sheet(wb, ws, 'Bipagens');

    const now = new Date();
    const stamp = `${now.getFullYear()}-${this.pad2(now.getMonth() + 1)}-${this.pad2(now.getDate())}_${this.pad2(now.getHours())}${this.pad2(now.getMinutes())}`;

    XLSX.writeFile(wb, `bipagens_todas_rotas_${this.workDay || this.todayLocalISO()}_${stamp}.xlsx`);
  }
};

// =======================
// Eventos / Boot
// =======================
$(document).ready(async () => {
  $(document).on('click', '#db-search-open', () => {
    $('#initial-interface').addClass('d-none');
    $('#db-search-interface').removeClass('d-none');
    $('#db-search-results-wrap').addClass('d-none');
  });

  $(document).on('click', '#db-search-back', () => {
    $('#db-search-interface').addClass('d-none');
    $('#db-search-results-wrap').addClass('d-none');
    $('#initial-interface').removeClass('d-none');
  });

  // Status da busca: no index vai para o painel lateral, no search.html para #db-search-status
  const searchStatus = (txt, kind) => {
    ConferenciaApp.setStatus(txt, kind);
    $('#db-search-status').text(txt);
  };

  const runDbSearch = async () => {
    try {
      const rawIds = ($('#db-ids').val() || '').trim();
      if (!rawIds) { alert('Informe pelo menos um ID.'); return; }

      const dayFrom = ($('#db-day-from').val() || '').trim() || undefined;
      const dayTo = ($('#db-day-to').val() || '').trim() || undefined;
      const op = ($('#db-op-filter').val() || $('#db-op-code').val() || '').trim() || undefined;

      searchStatus('Buscando histórico no banco...', 'info');

      const res = await ConferenciaApp.searchIdsFull(rawIds, {
        operation_code: op,
        day_from: dayFrom,
        day_to: dayTo
      });

      if (!res.ids.length) {
        searchStatus('Nenhum ID válido informado.', 'warning');
        return;
      }

      ConferenciaApp.renderDbSearchSummary(res.summary);
      searchStatus(`Busca concluída • ${res.summary.length} ID(s)`, 'success');
    } catch (e) {
      console.error(e);
      searchStatus('Erro ao buscar histórico.', 'danger');
      alert('Erro na busca: ' + (e?.message || e));
    }
  };

  $(document).on('click', '#db-search-btn', runDbSearch);

  $(document).on('keydown', '#db-ids', (e) => {
    if (e.ctrlKey && e.key === 'Enter') runDbSearch();
  });

  $(document).on('click', '#db-clear-btn', () => {
    $('#db-ids').val('');
    $('#db-search-results').empty();
    $('#db-search-results-wrap').addClass('d-none');
    $('#db-search-status').text('—');
  });

  // search.html: página só de busca, não carrega rotas/realtime
  if ($('#db-search-page').length) return;

  const today = ConferenciaApp.todayLocalISO();
  $('#work-day').val(today);

  await ConferenciaApp.ensureOperationSelected();

  if (ConferenciaApp.getOperationCode()) {
    await ConferenciaApp.applyWorkDay(today);
  }
});

// Confirmar operação escolhida
$(document).on('click', '#btn-op-confirm', async () => {
  const code = String($('#op-select').val() || '').trim().toUpperCase();
  if (!code) return;
  await ConferenciaApp.flushPendingNow();
  ConferenciaApp.setOperationCode(code);
  ConferenciaApp.resetForOperationChange();
  $('#modal-operation').modal('hide');

  const day = $('#work-day').val() || ConferenciaApp.todayLocalISO();
  await ConferenciaApp.applyWorkDay(day);
});

// Trocar operação
$(document).on('click', '#btn-change-op', async () => {
  await ConferenciaApp.ensureOperationSelected();
  $('#modal-operation').modal('show');
});

// Filtro de rotas
$(document).on('focus mousedown keydown input', '#saved-routes-inapp, #saved-routes, #route-search', () => {
  ConferenciaApp.lockRouteUi(3000);
});

$(document).on('focus', '#saved-routes-inapp', () => {
  ConferenciaApp.isRouteDropdownOpen = true;
  ConferenciaApp.lockRouteUi(3000);
});

$(document).on('blur change', '#saved-routes-inapp', () => {
  ConferenciaApp.isRouteDropdownOpen = false;
  ConferenciaApp.lockRouteUi(800);
});

$(document).on('input', '#route-search', (e) => {
  ConferenciaApp.lockRouteUi(3000);
  ConferenciaApp.applyRouteDropdownFilter(e.target.value);
});

// Troca de dia
$(document).on('change', '#work-day', async (e) => {
  const day = e.target.value;
  if (!day) return;
  await ConferenciaApp.applyWorkDay(day);
});

// Relatório noturno
$(document).on('click', '#finish-night-btn', () => {
  try {
    ConferenciaApp.showNightReport();
    const el = document.querySelector('#night-report');
    if (el) el.scrollIntoView({ block: 'start', behavior: 'smooth' });
  } catch (e) {
    console.error(e);
    alert('Falha ao gerar o relatório noturno.');
  }
});

$(document).on('click', '#finish-btn', () => {
  try {
    ConferenciaApp.showNightReport();
    const el = document.querySelector('#night-report');
    if (el) el.scrollIntoView({ block: 'start', behavior: 'smooth' });
  } catch (e) {
    console.error(e);
    alert('Falha ao gerar o relatório noturno.');
  }
});

$(document).on('click', '#night-report-close', () => {
  ConferenciaApp.hideNightReport();
});

// Importar HTML
$('#extract-btn').click(() => {
  const raw = $('#html-input').val();
  if (!raw.trim()) return alert('Cole o HTML antes de importar.');

  const qtd = ConferenciaApp.importRoutesFromHtml(raw);
  if (!qtd) return alert('Nenhuma rota importada. Confira se o HTML está completo.');

  $('#html-input').val('');
});

// Carregar rota
$('#load-route').click(() => {
  const id = $('#saved-routes').val();
  if (!id) return alert('Selecione uma rota salva.');

  ConferenciaApp.setCurrentRoute(id);

  $('#initial-interface').addClass('d-none');
  $('#manual-interface').addClass('d-none');
  $('#conference-interface').removeClass('d-none');
  $('#barcode-input').focus();
});

// Excluir rota
$('#delete-route').click(() => {
  const id = $('#saved-routes').val();
  if (!id) return alert('Selecione uma rota para excluir.');
  ConferenciaApp.deleteRoute(id);
});

// Limpar todas
$('#clear-all-routes').click(() => {
  const ok1 = confirm(
    'ATENÇÃO: isso vai APAGAR TODAS as rotas do DIA selecionado.\n\n' +
    'Quer continuar?'
  );
  if (!ok1) return;

  const day = ConferenciaApp.workDay || $('#work-day').val() || '(dia desconhecido)';
  const typed = prompt(
    `CONFIRMAÇÃO FINAL\n\n` +
    `Para apagar TUDO do dia ${day}, digite exatamente:\n` +
    `APAGAR\n\n` +
    `(Qualquer outra coisa cancela)`
  );

  if (typed !== 'APAGAR') {
    alert('Ação cancelada. Nada foi apagado.');
    return;
  }

  ConferenciaApp.clearAllRoutes();
  alert(`Tudo do dia ${day} foi removido.`);
});

// Trocar rota
$('#switch-route').click(() => {
  const id = $('#saved-routes-inapp').val();
  if (!id) return;
  ConferenciaApp.setCurrentRoute(id);
  $('#barcode-input').focus();
});

// Manual
$('#manual-btn').click(() => {
  $('#initial-interface').addClass('d-none');
  $('#manual-interface').removeClass('d-none');
});

$('#submit-manual').click(() => {
  try {
    const routeId = ($('#manual-routeid').val() || '').trim();
    if (!routeId) return alert('Informe o RouteId.');

    const cluster = ($('#manual-cluster').val() || '').trim();
    const brutos = ($('#manual-input').val() || '').split(/[\s,;]+/).map(x => x.trim()).filter(Boolean);
    // Mesmo formato da bipagem (11 dígitos), senão o ID nunca bateria na conferência
    const manualIds = Array.from(new Set(brutos.map(x => ConferenciaApp.normalizarCodigo(x)).filter(Boolean)));
    const ignorados = brutos.filter(x => !ConferenciaApp.normalizarCodigo(x)).length;

    if (!manualIds.length) return alert('Nenhum ID válido inserido (esperado: 11 dígitos começando com 4).');
    if (ignorados && !confirm(`${ignorados} valor(es) não parecem IDs válidos e serão ignorados. Continuar?`)) return;

    const route = ConferenciaApp.routes.get(String(routeId)) || ConferenciaApp.makeEmptyRoute(routeId);
    route.cluster = cluster || route.cluster;

    for (const id of manualIds) {
      route.ids.add(id);
      if (!route.conferidos.has(id)) route.faltantes.add(id);
    }

    route.totalInicial = route.ids.size;
    ConferenciaApp.routes.set(String(routeId), route);

    ConferenciaApp.lastRoutesSignature = '';
    ConferenciaApp.saveToStorage(ConferenciaApp.workDay);
    ConferenciaApp.markDirty('inserção manual');

    ConferenciaApp.renderRoutesSelects();

    alert(`Rota ${routeId} salva com ${route.totalInicial} ID(s).`);

    $('#manual-interface').addClass('d-none');
    $('#initial-interface').removeClass('d-none');
  } catch (e) {
    console.error(e);
    alert('Erro ao processar IDs manuais.');
  }
});

// Leitura do barcode
$('#barcode-input').on('keypress', (e) => {
  if (e.which === 13) {
    ConferenciaApp.viaCsv = false;

    const raw = $('#barcode-input').val();
    const id = ConferenciaApp.normalizarCodigo(raw);

    if (!id) {
      $('#barcode-input').val('').focus();
      return;
    }

    ConferenciaApp.conferirId(id);
  }
});

// Checar CSV
$('#check-csv').click(() => {
  const r = ConferenciaApp.current;
  if (!r) return alert('Selecione uma rota antes.');

  const fileInput = document.getElementById('csv-input');
  if (fileInput.files.length === 0) return alert('Selecione um arquivo CSV.');

  ConferenciaApp.viaCsv = true;

  const file = fileInput.files[0];
  const reader = new FileReader();

  reader.onload = e => {
    const csvText = e.target.result;
    const linhas = csvText.split(/\r?\n/);
    if (!linhas.length) return alert('Arquivo CSV vazio.');

    const header = linhas[0].split(',');
    const textCol = header.findIndex(h => /(text|texto|id)/i.test(h));
    if (textCol === -1) return alert('Coluna apropriada não encontrada (text/texto/id).');

    for (let i = 1; i < linhas.length; i++) {
      if (!linhas[i].trim()) continue;
      const cols = linhas[i].split(',');
      if (cols.length <= textCol) continue;

      let campo = cols[textCol].trim().replace(/^"|"$/g, '').replace(/""/g, '"');
      const id = ConferenciaApp.normalizarCodigo(campo);
      if (id) ConferenciaApp.conferirId(id);
    }

    ConferenciaApp.viaCsv = false;
    $('#barcode-input').focus();
  };

  reader.readAsText(file, 'UTF-8');
});

// Exports
$(document).on('click', '#export-csv-rota-atual', () => {
  ConferenciaApp.exportRotaAtualCsvPadrao();
});

$(document).on('click', '#export-xlsx-todas-rotas', () => {
  ConferenciaApp.exportTodasRotasXlsx();
});

$('#back-btn').click(() => {
  $('#conference-interface').addClass('d-none');
  $('#manual-interface').addClass('d-none');
  $('#initial-interface').removeClass('d-none');

  $('#barcode-input').val('');
  $('#html-input').focus();
});

// Carretas
$(document).on('click', '#carreta-btn', () => {
  $('#initial-interface').addClass('d-none');
  $('#conference-interface').addClass('d-none');
  $('#manual-interface').addClass('d-none');
  $('#carreta-interface').removeClass('d-none');

  ConferenciaApp.renderCarretaUI();
  $('#carreta-input').val('').focus();
});

$(document).on('click', '#carreta-back-btn', () => {
  $('#carreta-interface').addClass('d-none');
  $('#initial-interface').removeClass('d-none');
  $('#carreta-input').val('');
  $('#html-input').focus();
});

$(document).on('click', '#carreta-clear-current', () => {
  ConferenciaApp.carretas.currentPlateKey = null;
  ConferenciaApp.renderCarretaUI();
  $('#carreta-input').val('').focus();
});

$(document).on('click', '#patio-refresh', function() {
  ConferenciaApp.renderPatioGeral();
});

$(document).on('click', '#carreta-refresh-progress', () => {
  ConferenciaApp.renderCarretaProgress();
});

const processCarretaScan = (rawValue) => {
  const raw = String(rawValue || '').trim();
  if (!raw) return;

  const parsed = ConferenciaApp.parseScanPayload(raw);

  if (parsed.kind === 'plate') {
    const key = ConferenciaApp.ensurePlate(parsed.plate);
    ConferenciaApp.carretas.currentPlateKey = key;
    ConferenciaApp.renderCarretaUI();
    return;
  }

  if (parsed.kind === 'routeqr') {
    const pk = ConferenciaApp.carretas.currentPlateKey;
    if (!pk) {
      alert('Bipe uma PLACA primeiro.');
      return;
    }
    ConferenciaApp.vincularRouteQrNaPlaca(
      parsed.routeKey,
      {
        raw: (parsed.route && parsed.route.raw) ? parsed.route.raw : raw,
        jsonText: (parsed.route && parsed.route.jsonText) ? parsed.route.jsonText : ((parsed.route && parsed.route.obj) ? JSON.stringify(parsed.route.obj) : '')
      },
      pk,
      parsed.routeIdCandidate || ''
    );
    ConferenciaApp.renderCarretaUI();
    return;
  }

  if (parsed.kind === 'shipment') {
    alert('Aqui é a tela da CARRETA. Bipe a PLACA e os QRs das ROTAS (assignment/container).');
    return;
  }

  alert('QR não reconhecido. Bipe uma PLACA (JSON com license_plate) ou um QR de ROTA (JSON com assignment/container_id).');
};

$(document).on('keydown', '#carreta-input', (e) => {
  if (e.key === 'Enter' || e.which === 13) {
    e.preventDefault();
    const raw = $('#carreta-input').val();
    $('#carreta-input').val('');
    processCarretaScan(raw);
  }
});

$(document).on('click', '#carreta-check-links', () => {
  ConferenciaApp.checkLinksForCurrentPlate();
});

$(document).on('click', '#carreta-clear-bipagem-plate', () => {
  const pk = ConferenciaApp.carretas.currentPlateKey;
  if (!pk) return alert('Nenhuma placa ativa.');

  const ok = confirm(`Tem certeza que deseja EXCLUIR a bipagem/vínculos da placa ${pk}?`);
  if (!ok) return;

  ConferenciaApp.clearBipagemForPlate(pk);
});

$(document).on('paste', '#carreta-input', (e) => {
  const pasted = (e.originalEvent && e.originalEvent.clipboardData)
    ? e.originalEvent.clipboardData.getData('text')
    : '';
  setTimeout(() => {
    const raw = $('#carreta-input').val() || pasted;
    $('#carreta-input').val('');
    processCarretaScan(raw);
  }, 0);
});

// Exports novos
$(document).on('click', '#export-csv-rota-atual-placa', () => {
  ConferenciaApp.exportRotaAtualCsvComPlacaERota();
});

$(document).on('click', '#export-csv-todas-rotas-placa', () => {
  ConferenciaApp.exportTodasRotasCsvComPlacaERota();
});

$(document).on('click', '#export-csv-mapa-carretas', () => {
  ConferenciaApp.exportMapaCarretasCsv();
});

// Admin UI
$(document).on('click', '#btn-admin-open', async () => {
  $('#modal-admin').modal('show');
  await refreshAdminOps();
});

// Atalho para o Admin a partir do modal de operação (útil quando ainda não há operações)
$(document).on('click', '#btn-op-admin', () => {
  $('#modal-operation').one('hidden.bs.modal', () => $('#btn-admin-open').trigger('click'));
  $('#modal-operation').modal('hide');
});

// Ao fechar o Admin, recarrega a lista de operações (reabre a seleção se ainda não houver operação válida)
$(document).on('hidden.bs.modal', '#modal-admin', () => {
  ConferenciaApp.ensureOperationSelected();
});

async function refreshAdminOps() {
  try {
    const ops = await ConferenciaApp.adminLoadOperations(true);
    const $tbody = $('#admin-ops-tbody');
    if (!$tbody.length) return;
    $tbody.empty();
    ops.forEach(o => {
      const act = o.active ? 'SIM' : 'NÃO';
      const name = o.name || '';
      $tbody.append(`<tr><td>${ConferenciaApp.escHtml(o.code)}</td><td>${ConferenciaApp.escHtml(name)}</td><td>${act}</td></tr>`);
    });
  } catch (e) {
    console.warn(e);
  }
}

$(document).on('click', '#btn-admin-save-op', async () => {
  const code = $('#admin-op-code').val();
  const name = $('#admin-op-name').val();
  const active = $('#admin-op-active').is(':checked');
  try {
    await ConferenciaApp.adminUpsertOperation(code, name, active);
    await refreshAdminOps();
    alert('Operação salva.');
  } catch (e) {
    console.error(e);
    alert('Erro ao salvar operação: ' + (e.message || e));
  }
});

// Acompanhamento geral
$(document).on('click', '#btn-global-acomp', async () => {
  try {
    const day = $('#work-day').val() || ConferenciaApp.todayLocalISO();

    $('#initial-interface, #carreta-interface, #manual-interface, #conference-interface, #db-search-interface').addClass('d-none');
    $('#global-interface').removeClass('d-none');

    ConferenciaApp.setStatus(`Carregando acompanhamento geral • ${day}`, 'info');
    const items = await ConferenciaApp.loadGlobalProgress(day);
    ConferenciaApp.renderGlobalProgress(items, day);
    ConferenciaApp.setStatus(`Acompanhamento geral carregado • ${day}`, 'success');
  } catch (e) {
    console.warn(e);
    ConferenciaApp.setStatus('Falha ao carregar acompanhamento geral (ver console).', 'danger');
  }
});

$(document).on('click', '#global-refresh', async () => {
  try {
    const day = $('#work-day').val() || ConferenciaApp.todayLocalISO();
    const items = await ConferenciaApp.loadGlobalProgress(day);
    ConferenciaApp.renderGlobalProgress(items, day);
  } catch (e) {
    console.warn(e);
    ConferenciaApp.setStatus('Falha ao atualizar acompanhamento geral.', 'danger');
  }
});

$(document).on('click', '#global-back', () => {
  $('#global-interface').addClass('d-none');
  $('#initial-interface').removeClass('d-none');
});

// encerra realtime ao sair
window.addEventListener('beforeunload', () => {
  try { ConferenciaApp.stopRealtimeSync(); } catch {}
});