// Checagem periódica do banco (rede de segurança caso o realtime caia). Consultas baratas.
const SYNC_INTERVAL_MS = 30000;
// Espera após a última mudança nas rotas (importar/placa/excluir) antes de enviar ao banco
const DEFS_SAVE_DEBOUNCE_MS = 1500;
// Espera após a última bipagem antes de enviar o lote de bipagens
const EVENT_FLUSH_DEBOUNCE_MS = 800;

// Prefixo de chaves no localStorage (separa por operação e por dia)
const STORAGE_KEY_PREFIX = 'conferencia.v4';

const ConferenciaApp = {
  routes: new Map(),     // routeId -> routeObject (somente do dia selecionado)
  currentRouteId: null,
  viaCsv: false,
  operationCode: null, // ex: ERD1
  deviceId: null,
  cloudEnabled: true,
  syncTimer: null,
  periodicBusy: false,
  workDay: null,               // YYYY-MM-DD
  lastEvents: [],              // log simples de bipagens (últimos eventos)
  deletedRoutes: new Map(),    // routeId -> ts (epoch ms)
  revivedRoutes: new Map(),    // routeId -> ts (epoch ms) (desfaz exclusão)

  // ===== Definições das rotas (routes_state) =====
  defsDirty: false,
  defsSaving: false,
  defsSaveTimer: null,
  defsMutationSeq: 0,          // incrementa a cada mudança (detecta mudança durante o save)
  defsRemoteUpdatedAt: null,   // updated_at da última versão do banco que já incorporamos
  lastPushedDefsHash: '',
  lastSavedLocalDefsHash: '',

  // ===== Bipagens (scan_events) =====
  events: new Map(),           // client_id -> {cid, pkg, route, ts, dev, sv(1 = já está no banco), res}
  eventQueue: new Set(),       // client_ids ainda não enviados ao banco
  eventsSending: false,
  eventFlushTimer: null,
  eventsPersistTimer: null,
  maxServerId: 0,              // maior scan_events.id já recebido
  lastAppliedEv: null,
  lastOwnTs: 0,
  _foraOrDupIds: new Set(),    // IDs que já apareceram em fora de rota/duplicados (atalho de desempenho)

  // ===== Realtime =====
  rtChannel: null,
  rtBound: { op: null, day: null },

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
  // Sincronização com o Supabase
  // -----------------------
  // Dois tipos de dado, para gastar o mínimo de banco:
  //  - routes_state: DEFINIÇÕES das rotas (IDs importados, cluster, placas, exclusões).
  //    Muda poucas vezes por noite (importação, placa, exclusão).
  //  - scan_events: uma linha pequena por BIPAGEM. Conferidos / faltantes / fora de rota /
  //    duplicados são recalculados localmente a partir dos eventos, em ordem de horário.
  // =======================

  // ----- Realtime -----
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

  isCurrentDayOp(op, dayISO) {
    return op === this.getOperationCode() && dayISO === this.workDay;
  },

  async startRealtimeSync(dayISO) {
    const sb = this.getSb();
    const op = this.getOperationCode();
    if (!sb || !op || !dayISO) return;

    if (this.rtChannel && this.rtBound && this.rtBound.op === op && this.rtBound.day === dayISO) return;

    await this.stopRealtimeSync();

    this.rtChannel = sb
      .channel(`conferencia_${op}`)
      .on(
        'postgres_changes',
        { event: 'INSERT', schema: 'public', table: 'scan_events', filter: `operation_code=eq.${op}` },
        (payload) => {
          try {
            const row = payload && payload.new;
            if (!row || row.operation_code !== op || String(row.day) !== dayISO) return;
            if (!this.isCurrentDayOp(op, dayISO)) return;
            this.ingestServerEvents([row]);
          } catch (e) {
            console.warn('Falha ao aplicar bipagem do realtime:', e);
          }
        }
      )
      .on(
        'postgres_changes',
        { event: '*', schema: 'public', table: 'routes_state', filter: `operation_code=eq.${op}` },
        (payload) => {
          try {
            const row = payload && payload.new;
            if (!row || row.operation_code !== op || String(row.day) !== dayISO) return;
            if (!this.isCurrentDayOp(op, dayISO)) return;
            if (row.device_id && row.device_id === this.getDeviceId()) return;

            // Linhas grandes chegam sem o conteúdo pelo realtime: nesse caso busca no banco
            if (!row.data) {
              this.pullDefs(dayISO, { force: true }).catch(e => console.warn('Falha ao buscar rotas:', e));
              return;
            }
            this.defsRemoteUpdatedAt = row.updated_at || this.defsRemoteUpdatedAt;
            this.applyRemoteDefs(row.data, dayISO);
          } catch (e) {
            console.warn('Falha ao aplicar rotas do realtime:', e);
          }
        }
      )
      .subscribe((status) => {
        if (status === 'SUBSCRIBED') {
          this.setStatus(`Realtime ON • ${op} • ${dayISO}`, 'success');
        } else if (status === 'CHANNEL_ERROR' || status === 'TIMED_OUT') {
          this.setStatus(`Realtime instável • ${op} • ${dayISO}`, 'warning');
        } else if (status === 'CLOSED') {
          // A checagem periódica (startPeriodicSync) cobre o período sem realtime
          this.setStatus(`Realtime CLOSED • ${op} • ${dayISO}`, 'warning');
        }
        try { console.log('[Realtime]', status, { op, dayISO }); } catch {}
      });

    this.rtBound = { op, day: dayISO };
  },

  // =======================
  // Carretas: persistência (vai junto nas definições do dia)
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

  // =======================
  // Definições das rotas (routes_state)
  // =======================
  buildDefsSnapshot() {
    const routes = {};
    for (const [routeId, r] of this.routes.entries()) {
      routes[routeId] = this.serializeRouteDef(r);
    }
    return {
      v: 2,
      routes,
      meta: {
        deletedRoutes: Object.fromEntries(this.deletedRoutes || new Map()),
        revivedRoutes: Object.fromEntries(this.revivedRoutes || new Map()),
        carretas: this.serializeCarretas()
      }
    };
  },

  // Aceita o formato novo {v:2, routes, meta} e o antigo (rotas na raiz + __meta)
  splitDefsSnapshot(data) {
    if (!data || typeof data !== 'object') return { routes: {}, meta: {} };
    if (data.v === 2) return { routes: data.routes || {}, meta: data.meta || {} };
    const routes = {};
    for (const [k, v] of Object.entries(data)) if (k !== '__meta') routes[k] = v;
    return { routes, meta: data.__meta || {} };
  },

  // Junta definições remotas com as locais (união, com regras determinísticas de desempate)
  mergeDefsIntoLocal(data) {
    this.invalidateIdIndex();
    const { routes, meta } = this.splitDefsSnapshot(data);
    if (!this.deletedRoutes) this.deletedRoutes = new Map();
    if (!this.revivedRoutes) this.revivedRoutes = new Map();

    for (const [rid, ts] of Object.entries(meta.deletedRoutes || {})) {
      const id = String(rid);
      const t = Number(ts || 0);
      if (t > Number(this.deletedRoutes.get(id) || 0)) this.deletedRoutes.set(id, t);
    }
    for (const [rid, ts] of Object.entries(meta.revivedRoutes || {})) {
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

    this.mergeCarretas(meta.carretas);

    for (const [routeId, ser] of Object.entries(routes)) {
      const id = String(routeId);
      if (this.deletedRoutes.has(id)) continue;

      const tmp = this.deserializeRouteDef(id, ser);
      const existing = this.routes.get(id);
      if (!existing) {
        this.routes.set(id, tmp);
        continue;
      }

      tmp.ids.forEach(v => existing.ids.add(v));

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

      existing.resetAt = Math.max(Number(existing.resetAt || 0), Number(tmp.resetAt || 0));
      existing.totalInicial = Math.max(Number(existing.totalInicial || 0), Number(tmp.totalInicial || 0), existing.ids.size);
    }
  },

  // Aplica definições recebidas de outro aparelho
  applyRemoteDefs(data, dayISO) {
    this.mergeDefsIntoLocal(data);
    this.rebuildScanState();
    this.saveLocalDefs();

    // Se o local tem algo que o banco não tem, reenvia a união
    const localHash = this.computeSnapshotHash(this.buildDefsSnapshot());
    const remoteHash = this.computeSnapshotHash(data);
    if (localHash !== remoteHash) this.markDefsDirty();
    else if (!this.defsDirty) this.lastPushedDefsHash = remoteHash;

    this.refreshAfterRemoteMerge();
    this.setStatus(`Rotas atualizadas • ${this.getOperationCode()} • ${dayISO}`, 'info');
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
    return data || null;
  },

  // Consulta barata (~100 bytes): só a data da última gravação das definições
  async fetchDefsUpdatedAt(operationCode, dayISO) {
    const sb = this.getSb();
    if (!sb) throw new Error('Supabase client não encontrado (window.sbClient).');
    const { data, error } = await sb
      .from('routes_state')
      .select('updated_at')
      .eq('operation_code', operationCode)
      .eq('day', dayISO)
      .maybeSingle();
    if (error) throw error;
    return data ? data.updated_at : null;
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
    const { data, error } = await sb
      .from('routes_state')
      .upsert(payload, { onConflict: 'operation_code,day' })
      .select('updated_at')
      .single();
    if (error) throw error;
    return data ? data.updated_at : null;
  },

  markDefsDirty() {
    this.defsDirty = true;
    this.defsMutationSeq++;
    this.scheduleDefsSave(DEFS_SAVE_DEBOUNCE_MS);
    this.updatePendingFlag();
  },

  scheduleDefsSave(delayMs) {
    if (this.defsSaveTimer) clearTimeout(this.defsSaveTimer);
    this.defsSaveTimer = setTimeout(() => {
      this.defsSaveTimer = null;
      this.flushDefsSave().catch(e => {
        console.warn('Falha ao salvar rotas no Supabase (vai tentar de novo):', e);
        this.setStatus('Falha ao salvar rotas no banco. Mantido no cache local, tentando de novo...', 'warning');
        this.scheduleDefsSave(5000);
      });
    }, delayMs);
  },

  async flushDefsSave() {
    if (!this.cloudEnabled || !this.defsDirty) return;
    if (this.defsSaving) {
      this.scheduleDefsSave(500);
      return;
    }

    const op = this.getOperationCode();
    const day = this.workDay;
    if (!op || !day) return;

    this.defsSaving = true;
    const seqAtStart = this.defsMutationSeq;
    try {
      // Só baixa as definições remotas se outro aparelho gravou depois da última vez que vimos
      const remoteTs = await this.fetchDefsUpdatedAt(op, day);
      if (!this.isCurrentDayOp(op, day)) return;

      let remoteHash = null;
      if (remoteTs && remoteTs !== this.defsRemoteUpdatedAt) {
        const row = await this.supaLoadDaySnapshot(op, day);
        if (!this.isCurrentDayOp(op, day)) return;
        if (row && row.data) {
          this.mergeDefsIntoLocal(row.data);
          this.rebuildScanState();
          this.saveLocalDefs();
          this.refreshAfterRemoteMerge();
          remoteHash = this.computeSnapshotHash(row.data);
        }
        this.defsRemoteUpdatedAt = remoteTs;
      }

      const snapshot = this.buildDefsSnapshot();
      const hash = this.computeSnapshotHash(snapshot);
      const remoteUnchanged = remoteTs && remoteTs === this.defsRemoteUpdatedAt;

      if (hash === remoteHash || (remoteUnchanged && hash === this.lastPushedDefsHash)) {
        this.lastPushedDefsHash = hash;
      } else {
        const savedAt = await this.supaSaveDaySnapshot(op, day, snapshot);
        this.defsRemoteUpdatedAt = savedAt || this.defsRemoteUpdatedAt;
        this.lastPushedDefsHash = hash;
      }

      // Se mudou algo durante o save, continua pendente e agenda outro envio
      if (this.defsMutationSeq !== seqAtStart) {
        this.scheduleDefsSave(300);
      } else {
        this.defsDirty = false;
        this.setStatus(`Rotas salvas no banco • ${op} • ${day}`, 'success');
      }
    } finally {
      this.defsSaving = false;
      this.updatePendingFlag();
    }
  },

  // Busca as definições no banco (só baixa o conteúdo se mudou)
  async pullDefs(dayISO, { force = false } = {}) {
    const op = this.getOperationCode();
    if (!op || !dayISO) return;

    const remoteTs = await this.fetchDefsUpdatedAt(op, dayISO);
    if (!this.isCurrentDayOp(op, dayISO)) return;

    if (!remoteTs) {
      // Banco ainda sem esse dia, mas há rotas locais: envia
      if (this.routes.size) this.markDefsDirty();
      return;
    }
    if (!force && remoteTs === this.defsRemoteUpdatedAt) return;

    const row = await this.supaLoadDaySnapshot(op, dayISO);
    if (!row || !row.data || !this.isCurrentDayOp(op, dayISO)) return;

    this.defsRemoteUpdatedAt = row.updated_at || remoteTs;
    this.applyRemoteDefs(row.data, dayISO);
  },

  // =======================
  // Bipagens (scan_events)
  // =======================
  newId() {
    if (crypto.randomUUID) return crypto.randomUUID();
    return `${Date.now().toString(16)}-${Math.random().toString(16).slice(2)}-${Math.random().toString(16).slice(2)}`;
  },

  compareEvents(a, b) {
    return (a.ts - b.ts) || (a.cid < b.cid ? -1 : a.cid > b.cid ? 1 : 0);
  },

  // Aplica UMA bipagem sobre o estado das rotas (mesma regra da conferência manual)
  applyScanEvent(ev) {
    const r = this.routes.get(String(ev.route));
    if (!r) { ev.res = null; return null; }
    // Bipagens anteriores a uma reimportação da rota não contam
    if (ev.ts < Number(r.resetAt || 0)) { ev.res = null; return null; }

    const codigo = ev.pkg;
    const now = ev.ts;
    const correctRouteId = this.findCorrectRouteForId(codigo);
    const isCorrectHere = correctRouteId && String(correctRouteId) === String(r.routeId);

    if (isCorrectHere) {
      if (!r.conferidos.has(codigo)) {
        r.faltantes.delete(codigo);
        r.foraDeRota.delete(codigo);
        r.conferidos.add(codigo);
        r.timestamps.set(codigo, now);
        this.cleanupIdFromOtherRoutes(codigo, r.routeId);
        ev.res = 'ok';
        return ev.res;
      }
      this.cleanupIdFromOtherRoutes(codigo, r.routeId);
    }

    if (r.conferidos.has(codigo) || r.foraDeRota.has(codigo)) {
      r.duplicados.set(codigo, this.dupCount(r.duplicados.get(codigo) || 1) + 1);
      this._foraOrDupIds.add(codigo);
      r.timestamps.set(codigo, now);
      ev.res = 'dup';
      return ev.res;
    }

    if (r.faltantes.has(codigo)) {
      r.faltantes.delete(codigo);
      r.conferidos.add(codigo);
      r.timestamps.set(codigo, now);
      this.cleanupIdFromOtherRoutes(codigo, r.routeId);
      ev.res = 'ok';
      return ev.res;
    }

    r.foraDeRota.add(codigo);
    this._foraOrDupIds.add(codigo);
    r.timestamps.set(codigo, now);
    ev.res = 'fora';
    return ev.res;
  },

  // Recalcula conferidos/faltantes/fora/duplicados do zero a partir de todos os eventos
  rebuildScanState() {
    this.invalidateIdIndex();
    this._foraOrDupIds = new Set();
    for (const r of this.routes.values()) {
      r.conferidos = new Set();
      r.foraDeRota = new Set();
      r.duplicados = new Map();
      r.timestamps = new Map();
      r.faltantes = new Set(r.ids);
    }
    const list = Array.from(this.events.values()).sort((a, b) => this.compareEvents(a, b));
    for (const ev of list) this.applyScanEvent(ev);
    this.lastAppliedEv = list.length ? list[list.length - 1] : null;
  },

  // Adiciona eventos; aplica em sequência ou recalcula tudo se chegou algum "no passado"
  addEvents(evs) {
    const novos = [];
    for (const ev of evs) {
      const cur = this.events.get(ev.cid);
      if (cur) {
        if (ev.sv) cur.sv = 1;
        continue;
      }
      this.events.set(ev.cid, ev);
      novos.push(ev);
    }
    if (!novos.length) return novos;

    novos.sort((a, b) => this.compareEvents(a, b));
    const last = this.lastAppliedEv;
    if (last && this.compareEvents(novos[0], last) < 0) {
      this.rebuildScanState();
    } else {
      for (const ev of novos) this.applyScanEvent(ev);
      this.lastAppliedEv = novos[novos.length - 1];
    }
    return novos;
  },

  serverRowToEvent(row) {
    return {
      cid: row.client_id ? String(row.client_id) : `srv-${row.id}`,
      pkg: String(row.package_id),
      route: String(row.route_id ?? ''),
      ts: Date.parse(row.scanned_at) || 0,
      dev: row.device_id || '',
      sv: 1
    };
  },

  // Eventos vindos do banco (carga inicial, checagem periódica ou realtime)
  ingestServerEvents(rows) {
    if (!rows || !rows.length) return;
    for (const row of rows) {
      if (Number(row.id) > this.maxServerId) this.maxServerId = Number(row.id);
    }
    const novos = this.addEvents(rows.map(r => this.serverRowToEvent(r)));
    this.persistEventsSoon();
    if (novos.length) this.refreshAfterRemoteMerge();
  },

  eventToRow(ev, op, day) {
    const meta = this.getRouteMeta(ev.route);
    return {
      client_id: ev.cid,
      device_id: ev.dev || this.getDeviceId(),
      operation_code: op,
      day,
      package_id: ev.pkg,
      route_id: ev.route || null,
      cluster: meta.cluster || null,
      xpt: meta.xpt || null,
      result: ev.res || null,
      scanned_at: new Date(ev.ts).toISOString()
    };
  },

  scheduleEventFlush(delayMs = EVENT_FLUSH_DEBOUNCE_MS) {
    if (this.eventFlushTimer) clearTimeout(this.eventFlushTimer);
    this.eventFlushTimer = setTimeout(() => {
      this.eventFlushTimer = null;
      this.flushEventQueue();
    }, delayMs);
  },

  // Envia as bipagens pendentes em lotes. client_id único => reenvio nunca duplica.
  async flushEventQueue() {
    if (!this.cloudEnabled || !this.eventQueue.size) return;
    if (this.eventsSending) {
      this.scheduleEventFlush(500);
      return;
    }
    const sb = this.getSb();
    const op = this.getOperationCode();
    const day = this.workDay;
    if (!sb || !op || !day) return;

    this.eventsSending = true;
    try {
      while (this.eventQueue.size && this.isCurrentDayOp(op, day)) {
        const cids = Array.from(this.eventQueue).slice(0, 500);
        const rows = cids.map(cid => this.events.get(cid)).filter(Boolean).map(ev => this.eventToRow(ev, op, day));

        if (rows.length) {
          const { error } = await sb
            .from('scan_events')
            .upsert(rows, { onConflict: 'client_id', ignoreDuplicates: true });
          if (error) throw error;
        }

        for (const cid of cids) {
          this.eventQueue.delete(cid);
          const ev = this.events.get(cid);
          if (ev) ev.sv = 1;
        }
        this.persistEventsSoon();
      }
      if (!this.eventQueue.size) this.setStatus(`Bipagens salvas no banco • ${op} • ${day}`, 'success');
    } catch (e) {
      console.warn('Falha ao enviar bipagens (vai tentar de novo):', e);
      this.setStatus(`Sem conexão com o banco • ${this.eventQueue.size} bipagem(ns) guardada(s) no aparelho`, 'warning');
      this.scheduleEventFlush(5000);
    } finally {
      this.eventsSending = false;
      this.updatePendingFlag();
    }
  },

  // Baixa bipagens do dia. Incremental (id > último visto) por padrão; paginado de 1000 em 1000.
  async pullEvents(dayISO, { full = false } = {}) {
    const sb = this.getSb();
    const op = this.getOperationCode();
    if (!sb || !op || !dayISO) return;

    let after = full ? 0 : this.maxServerId;
    let rows = [];
    for (;;) {
      const { data, error } = await sb
        .from('scan_events')
        .select('id,client_id,device_id,package_id,route_id,scanned_at')
        .eq('operation_code', op)
        .eq('day', dayISO)
        .gt('id', after)
        .order('id', { ascending: true })
        .limit(1000);
      if (error) throw error;
      if (!this.isCurrentDayOp(op, dayISO)) return;
      if (!data || !data.length) break;
      rows = rows.concat(data);
      after = data[data.length - 1].id;
      if (data.length < 1000) break;
    }
    if (rows.length) this.ingestServerEvents(rows);
  },

  // Consulta barata: só a contagem de bipagens do dia no banco (sem baixar linhas)
  async countServerEvents(op, dayISO) {
    const sb = this.getSb();
    const { count, error } = await sb
      .from('scan_events')
      .select('id', { count: 'exact', head: true })
      .eq('operation_code', op)
      .eq('day', dayISO);
    if (error) throw error;
    return count || 0;
  },

  knownServerEventsCount() {
    let n = 0;
    for (const ev of this.events.values()) if (ev.sv) n++;
    return n;
  },

  // Checagem periódica: só baixa algo se o banco tiver mudado
  async periodicSyncTick() {
    const op = this.getOperationCode();
    const day = this.workDay;
    if (!op || !day || this.periodicBusy) return;
    if (typeof document !== 'undefined' && document.hidden) return;

    this.periodicBusy = true;
    try {
      if (this.eventQueue.size && !this.eventsSending && !this.eventFlushTimer) this.scheduleEventFlush(0);
      if (this.defsDirty && !this.defsSaving && !this.defsSaveTimer) this.scheduleDefsSave(0);

      if (!this.defsDirty) await this.pullDefs(day);

      const count = await this.countServerEvents(op, day);
      if (count > this.knownServerEventsCount()) {
        await this.pullEvents(day);
        // Ainda faltando (ex.: inserção concorrente com id menor): recarrega o dia inteiro
        if (count > this.knownServerEventsCount()) await this.pullEvents(day, { full: true });
      }
    } catch (e) {
      console.warn('Falha na checagem periódica:', e);
    } finally {
      this.periodicBusy = false;
    }
  },

  startPeriodicSync() {
    if (this.syncTimer) clearInterval(this.syncTimer);
    this.syncTimer = setInterval(() => this.periodicSyncTick(), SYNC_INTERVAL_MS);
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

  dupCount(v) {
    if (Array.isArray(v)) return v.reduce((m, x) => Math.max(m, Number(x) || 0), 0);
    return Number(v) || 0;
  },

  updatePendingFlag() {
    this.dirty = !!(this.defsDirty || (this.eventQueue && this.eventQueue.size));
    if (this.dirty) $('#dirty-flag').removeClass('d-none');
    else $('#dirty-flag').addClass('d-none');
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

  // Definições das rotas mudaram (importação, exclusão, placa...): agenda envio ao banco
  markDirty(reason = '') {
    this.markDefsDirty();
    const msg = reason ? `pendente salvar (${reason})` : 'pendente salvar';
    this.setStatus(msg, 'warning');
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

  // Log local das últimas bipagens (painel "Acompanhamento")
  pushEvent(evt) {
    this.lastEvents.unshift(Object.assign({}, evt));
    if (this.lastEvents.length > 80) this.lastEvents.length = 80;
    this.renderAcompanhamento();
  },

  renderAcompanhamento() {
    this.updatePendingFlag();

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
      plateUpdatedAt: 0,

      resetAt: 0 // bipagens anteriores a este horário não contam (rota reimportada após exclusão)
    };
  },

  get current() {
    if (!this.currentRouteId) return null;
    return this.routes.get(String(this.currentRouteId)) || null;
  },

  // =======================
  // Persistência local (funciona offline; o banco é sincronizado depois)
  // =======================
  loadLocal(dayISO) {
    const key = this.storageKeyForDay(dayISO);
    try {
      const raw = localStorage.getItem(key);
      if (raw) this.mergeDefsIntoLocal(JSON.parse(raw));
    } catch (e) {
      console.warn('Falha ao carregar rotas do cache local:', e);
    }

    try {
      const raw = localStorage.getItem(`${key}.events`);
      if (raw) {
        const saved = JSON.parse(raw);
        for (const [cid, pkg, route, ts, dev, sv] of (saved.events || [])) {
          this.events.set(cid, { cid, pkg, route, ts: Number(ts), dev, sv: sv ? 1 : 0 });
        }
        (saved.queue || []).forEach(cid => { if (this.events.has(cid)) this.eventQueue.add(cid); });
        this.maxServerId = Number(saved.maxServerId || 0);
      }
    } catch (e) {
      console.warn('Falha ao carregar bipagens do cache local:', e);
    }

    this.rebuildScanState();
    this.lastSavedLocalDefsHash = this.computeSnapshotHash(this.buildDefsSnapshot());
  },

  saveLocalDefs() {
    if (!this.workDay) return '';
    try {
      const obj = this.buildDefsSnapshot();
      const hash = this.computeSnapshotHash(obj);
      if (hash !== this.lastSavedLocalDefsHash) {
        localStorage.setItem(this.storageKeyForDay(this.workDay), JSON.stringify(obj));
        this.lastSavedLocalDefsHash = hash;
      }
      return hash;
    } catch (e) {
      console.warn('Falha ao salvar rotas no cache local:', e);
      return '';
    }
  },

  // Chamado após mudar definições (importar, excluir, placa...): salva local e agenda envio
  saveToStorage(dayISO, opts = {}) {
    // Definições mudaram: reavalia as bipagens (ex.: um "fora de rota" pode virar conferido)
    this.rebuildScanState();
    const { syncCloud = true } = opts;
    const hash = this.saveLocalDefs();
    if (syncCloud && hash !== this.lastPushedDefsHash) this.markDefsDirty();
  },

  persistEventsNow() {
    if (this.eventsPersistTimer) {
      clearTimeout(this.eventsPersistTimer);
      this.eventsPersistTimer = null;
    }
    if (!this.workDay || !this.getOperationCode()) return;
    try {
      const events = Array.from(this.events.values()).map(ev => [ev.cid, ev.pkg, ev.route, ev.ts, ev.dev, ev.sv ? 1 : 0]);
      localStorage.setItem(`${this.storageKeyForDay(this.workDay)}.events`, JSON.stringify({
        events,
        queue: Array.from(this.eventQueue),
        maxServerId: this.maxServerId
      }));
    } catch (e) {
      console.warn('Falha ao salvar bipagens no cache local:', e);
    }
  },

  persistEventsSoon() {
    if (this.eventsPersistTimer) return;
    this.eventsPersistTimer = setTimeout(() => this.persistEventsNow(), 300);
  },

  serializeRouteDef(r) {
    return {
      routeId: r.routeId,
      cluster: r.cluster,
      destinationFacilityId: r.destinationFacilityId,
      destinationFacilityName: r.destinationFacilityName,
      totalInicial: r.totalInicial,
      ids: Array.from(r.ids),
      resetAt: Number(r.resetAt || 0),

      plateKey: r.plateKey || '',
      plateRaw: r.plateRaw || '',
      plateLicense: r.plateLicense || '',
      routeQrKey: r.routeQrKey || '',
      routeQrRaw: r.routeQrRaw || '',
      plateScanTs: Number(r.plateScanTs || 0),
      routeQrScanTs: Number(r.routeQrScanTs || 0),
      plateUpdatedAt: Number(r.plateUpdatedAt || 0),
    };
  },

  deserializeRouteDef(routeId, r) {
    const route = this.makeEmptyRoute(routeId);
    r = r || {};

    route.cluster = r.cluster || '';
    route.destinationFacilityId = r.destinationFacilityId || '';
    route.destinationFacilityName = r.destinationFacilityName || '';
    (r.ids || []).forEach(id => route.ids.add(String(id)));
    route.totalInicial = Math.max(Number(r.totalInicial || 0), route.ids.size);
    route.resetAt = Number(r.resetAt || 0);

    route.plateKey = r.plateKey || '';
    route.plateRaw = r.plateRaw || '';
    route.plateLicense = r.plateLicense || '';
    route.routeQrKey = r.routeQrKey || '';
    route.routeQrRaw = r.routeQrRaw || '';
    route.plateScanTs = Number(r.plateScanTs || 0);
    route.routeQrScanTs = Number(r.routeQrScanTs || 0);
    route.plateUpdatedAt = Number(r.plateUpdatedAt || r.routeQrScanTs || 0);

    route.faltantes = new Set(route.ids);
    return route;
  },

  // =======================
  // Troca de dia / operação
  // =======================
  // Envia ao banco o que estiver pendente (usado antes de trocar dia/operação)
  async flushPendingNow() {
    for (const t of ['defsSaveTimer', 'eventFlushTimer']) {
      if (this[t]) {
        clearTimeout(this[t]);
        this[t] = null;
      }
    }
    for (let i = 0; i < 50 && (this.defsSaving || this.eventsSending); i++) {
      await new Promise(res => setTimeout(res, 100));
    }
    try {
      if (this.defsDirty) await this.flushDefsSave();
    } catch (e) {
      console.warn('Falha ao enviar rotas antes da troca (ficam no cache local):', e);
    }
    await this.flushEventQueue();
    for (const t of ['defsSaveTimer', 'eventFlushTimer']) {
      if (this[t]) {
        clearTimeout(this[t]);
        this[t] = null;
      }
    }
    this.persistEventsNow();
    this.saveLocalDefs();
  },

  resetDayState() {
    this.invalidateIdIndex();
    this.routes.clear();
    this.resetCarretas();
    this.currentRouteId = null;
    this.lastRoutesSignature = '';

    this.deletedRoutes = new Map();
    this.revivedRoutes = new Map();
    this.lastEvents = [];

    this.events = new Map();
    this.eventQueue = new Set();
    this.maxServerId = 0;
    this.lastAppliedEv = null;

    this.defsDirty = false;
    this.defsRemoteUpdatedAt = null;
    this.lastPushedDefsHash = '';
    this.lastSavedLocalDefsHash = '';
  },

  async applyWorkDay(dayISO) {
    await this.flushPendingNow();
    this.resetDayState();

    this.workDay = dayISO;
    $('#work-day').val(dayISO);

    const op = this.getOperationCode();
    if (op) $('#op-badge').text(op);

    this.loadLocal(dayISO);
    this.renderRoutesSelects();
    this.refreshUIFromCurrent();
    this.renderAcompanhamento();

    if (op) {
      // Realtime primeiro, para não perder nada que chegue durante a carga
      await this.startRealtimeSync(dayISO);
      try {
        await this.pullDefs(dayISO, { force: true });
        await this.pullEvents(dayISO);
      } catch (e) {
        console.warn('Falha ao carregar o dia do banco (seguindo com o cache local):', e);
        this.setStatus('Sem conexão com o banco • usando dados do aparelho', 'warning');
      }
      this.startPeriodicSync();
      if (this.eventQueue.size) this.scheduleEventFlush(0);
    }

    this.renderRoutesSelects();
    this.refreshUIFromCurrent();
    this.renderAcompanhamento();
    this.updatePendingFlag();

    if (op) this.setStatus(`dia carregado • ${op} • ${dayISO}`, 'success');
    else this.setStatus('dia carregado (sem operação)', 'warning');
  },

  resetForOperationChange() {
    this.stopRealtimeSync();
    this.resetDayState();
    this.viaCsv = false;

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

  // Acompanhamento geral: calculado no próprio banco (função day_progress), sem baixar os dados
  async loadGlobalProgress(dayISO) {
    const sb = this.getSb();
    if (!sb) throw new Error('Supabase client não encontrado (window.sbClient).');

    const { data, error } = await sb.rpc('day_progress', { p_day: dayISO });
    if (error) throw error;

    return (data || []).map(o => {
      const totalIds = Number(o.total_ids || 0);
      const conferidos = Number(o.conferidos || 0);
      return {
        code: String(o.operation_code || '').toUpperCase(),
        name: o.name || '',
        stats: {
          routesCount: Number(o.routes || 0),
          totalIds,
          conferidos,
          faltantes: Math.max(0, totalIds - conferidos),
          fora: Number(o.fora || 0)
        },
        updated_at: o.updated_at || null
      };
    });
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
  // Rota "dona" do ID (primeira rota que tem o ID na lista importada).
  // Usa um índice ID -> rota, invalidado sempre que as definições mudam.
  findCorrectRouteForId(id) {
    if (!this._idIndex) {
      this._idIndex = new Map();
      for (const [rid, r] of this.routes.entries()) {
        for (const x of r.ids) if (!this._idIndex.has(x)) this._idIndex.set(x, String(rid));
      }
    }
    return this._idIndex.get(id) || null;
  },

  invalidateIdIndex() {
    this._idIndex = null;
  },

  cleanupIdFromOtherRoutes(id, targetRouteId) {
    // Caminho rápido: ID nunca esteve em fora de rota/duplicados => nada para limpar
    if (this._foraOrDupIds && !this._foraOrDupIds.has(id)) return;
    const target = String(targetRouteId);

    for (const [rid, r] of this.routes.entries()) {
      // Caminho rápido: na grande maioria das rotas o ID não está em fora/duplicados
      if (!r.foraDeRota.has(id) && !r.duplicados.has(id)) continue;
      if (rid === target) continue;

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

    // Horário estritamente crescente por aparelho (desempate estável na ordenação)
    let ts = Date.now();
    if (ts <= this.lastOwnTs) ts = this.lastOwnTs + 1;
    this.lastOwnTs = ts;

    const ev = {
      cid: this.newId(),
      pkg: String(codigo),
      route: String(r.routeId),
      ts,
      dev: this.getDeviceId(),
      sv: 0
    };

    this.addEvents([ev]);
    this.eventQueue.add(ev.cid);
    this.persistEventsSoon();
    this.scheduleEventFlush();
    this.updatePendingFlag();

    const tipo = ev.res || 'ok';
    if (tipo !== 'ok' && !this.viaCsv) this.playAlertSound();

    $('#barcode-input').val('').focus();
    this.pushEvent({
      ts,
      type: tipo,
      code: codigo,
      currentRouteId: this.currentRouteId,
      correctRouteId: this.findCorrectRouteForId(codigo)
    });
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

      let revivedAt = 0;
      if (this.deletedRoutes?.has(routeId)) {
        this.deletedRoutes.delete(routeId);
        if (!this.revivedRoutes) this.revivedRoutes = new Map();
        revivedAt = Date.now();
        this.revivedRoutes.set(routeId, revivedAt);
      }

      const route = this.routes.get(routeId) || this.makeEmptyRoute(routeId);
      // Rota excluída e importada de novo começa zerada (bipagens antigas não contam)
      if (revivedAt) route.resetAt = revivedAt;

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
  try { ConferenciaApp.persistEventsNow(); } catch {}
  try { ConferenciaApp.stopRealtimeSync(); } catch {}
});

// Ao voltar para a aba, sincroniza na hora (a checagem periódica pausa com a aba oculta)
document.addEventListener('visibilitychange', () => {
  if (!document.hidden) ConferenciaApp.periodicSyncTick();
});