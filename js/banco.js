// Consultas ao banco: admin de operações, acompanhamento geral e busca de IDs.

Object.assign(ConferenciaApp, {
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

  async loadOperationsRows(includeInactive) {
    const sb = this.getSb();
    if (!sb) throw new Error('Supabase client não encontrado.');
    let q = sb.from('operations').select('code,name,active,created_at').order('code', { ascending: true });
    if (!includeInactive) q = q.eq('active', true);
    const { data, error } = await q;
    if (error) throw error;
    return data || [];
  },

  async adminLoadOperations(includeInactive = true) {
    return this.loadOperationsRows(includeInactive);
  },

  // ===== PIN (único para todas as operações) =====
  // O hash do PIN fica numa tabela que a chave pública não lê; a conferência é feita
  // no banco pelas funções pin_enabled() e check_pin(). Definir/trocar o PIN: só por SQL.

  // true = existe PIN cadastrado; false = sem PIN (ou SQL do PIN ainda não aplicado)
  async pinEnabled() {
    const sb = this.getSb();
    if (!sb) throw new Error('Supabase client não encontrado.');
    const { data, error } = await sb.rpc('pin_enabled');
    if (error) {
      // Função não existe (SQL do PIN não aplicado): segue sem PIN
      if (error.code === 'PGRST202' || error.code === '42883') return false;
      throw error;
    }
    return !!data;
  },

  async verifyPin(pin) {
    const sb = this.getSb();
    if (!sb) throw new Error('Supabase client não encontrado.');
    const { data, error } = await sb.rpc('check_pin', { p_pin: String(pin || '').trim() });
    if (error) throw error;
    return data === true;
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
});
