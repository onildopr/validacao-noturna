// Regras da conferência: aplica bipagens (ok / fora de rota / duplicado) e recalcula o estado.

Object.assign(ConferenciaApp, {
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
});
