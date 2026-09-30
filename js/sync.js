// Sincronização com o Supabase (routes_state + scan_events), realtime e cache local.

Object.assign(ConferenciaApp, {
  // Sincronização com o Supabase
  // -----------------------
  // Dois tipos de dado, para gastar o mínimo de banco:
  //  - routes_state: DEFINIÇÕES das rotas (IDs importados, cluster, placas, exclusões).
  //    Muda poucas vezes por noite (importação, placa, exclusão).
  //  - scan_events: uma linha pequena por BIPAGEM. Conferidos / faltantes / fora de rota /
  //    duplicados são recalculados localmente a partir dos eventos, em ordem de horário.

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
        this.cloudOffline = true;
        this.updatePendingFlag();
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
        this.cloudOffline = false;
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
      this.cloudOffline = false;
      if (!this.eventQueue.size) this.setStatus(`Bipagens salvas no banco • ${op} • ${day}`, 'success');
    } catch (e) {
      console.warn('Falha ao enviar bipagens (vai tentar de novo):', e);
      this.cloudOffline = true;
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

  // Definições das rotas mudaram (importação, exclusão, placa...): agenda envio ao banco
  markDirty(reason = '') {
    this.markDefsDirty();
    const msg = reason ? `pendente salvar (${reason})` : 'pendente salvar';
    this.setStatus(msg, 'warning');
  },

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
});
