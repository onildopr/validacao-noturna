const { jsPDF } = window.jspdf || {};

const STORAGE_KEY_PREFIX = 'conferencia.routes.daycache.v1'; // cache local por dia
const SYNC_INTERVAL_MS = 20000;
const AUTO_SAVE_INTERVAL_MS = 10000; // 2 min


const ConferenciaApp = {
  routes: new Map(),     // routeId -> routeObject (somente do dia selecionado)
  currentRouteId: null,
  viaCsv: false,
  operationCode: null, // ex: ERD1
  adminEnabled: false, // true quando admin estiver logado
  deviceId: null,
  cloudEnabled: true,
  cloudDirty: false,
  cloudSaving: false,
  cloudLastSaveAt: 0,
  cloudSaveTimer: null,
  workDay: null,               // YYYY-MM-DD
  lastEvents: [],            // log simples de bipagens (últimos eventos)

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

  // =======================
  // Util data/strings
  // =======================

  // Normaliza texto de cluster/assignment para comparação (sem depender de acentos/lixo do leitor)
  normalizeCluster(v) {
    return String(v ?? '')
      .trim()
      .toUpperCase()
      // remove acentos
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      // mantém letras/números/_/- e remove o resto (inclusive { } ^ espaços e caracteres estranhos)
      .replace(/[^\w\-]+/g, '');
  },

  pad2(n) { return String(n).padStart(2, '0'); },

  todayLocalISO() {
    // Sempre calcula o "dia" no fuso de Rondônia, mesmo se o PC estiver com outro fuso
    const fmt = new Intl.DateTimeFormat('en-CA', {
      timeZone: 'America/Porto_Velho',
      year: 'numeric',
      month: '2-digit',
      day: '2-digit'
    });
    return fmt.format(new Date()); // YYYY-MM-DD
  },

  monthKeyFromDay(dayISO) {
    // "2026-01"
    return String(dayISO || '').slice(0, 7);
  },

  normalizeCaretKey(k) {
    const key = String(k || '').trim().toLowerCase();

    // remove acentos
    const noAcc = key.normalize('NFD').replace(/[\u0300-\u036f]/g, '');

    // mapeia variações comuns
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
    // pega só a parte antes de "Placa" / "Rota" (se existir)
    const cleaned = String(raw || '').replace(/\r/g, '').trim();
    const first = cleaned.split('\n')[0].trim();

    const kv = {};

    // separa por vírgula e tenta pegar pares "chave^Çvalor"
    const tokens = first.split(',').map(t => t.trim()).filter(Boolean);

    for (const tok of tokens) {
      // exemplos de token:
      // ^license_plate^Ç^SHM6F53^
      // âssignment^Ç^J11_AM5^{
      // ^container_id^Ç10356375
      const m = tok.match(/^(\^?)([^\\^]+?)\^Ç\^?(.+?)\^?$/);
      if (!m) continue;

      const rawKey = m[2];
      let val = m[3];

      const key = this.normalizeCaretKey(rawKey);

      // limpa valor
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
    // Atualiza badge no topo, se existir
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
  // Supabase (snapshot por operação+dia)
  // Tabela esperada: routes_state(operation_code text, day date, data jsonb, updated_at timestamptz, device_id text)
  // =======================
  getSb() {
    // tenta usar client já criado no HTML; se não existir, tenta criar com config global (SB_URL/SB_ANON)
    if (window.sbClient) return window.sbClient;
    if (window.supabase && window.SB_URL && window.SB_ANON) {
      try {
        window.sbClient = window.supabase.createClient(window.SB_URL, window.SB_ANON);
        return window.sbClient;
      } catch (e) {
        console.warn('Falha ao criar Supabase client automaticamente:', e);
      }
    }
    return null;
  },
  async isAdmin() {
    try {
      const sb = this.getSb();
      if (!sb) return false;
      const { data, error } = await sb.auth.getUser();
      if (error) return false;
      return !!data?.user; // admin = usuário autenticado
    } catch { return false; }
  },

  setAdminEnabled(flag) {
    this.adminEnabled = !!flag;
    try {
      const disabled = !this.adminEnabled;
      $('#btn-delete-route').prop('disabled', disabled);
      $('#btn-clear-routes').prop('disabled', disabled);
    } catch {}
  },

  requireAdminSync(actionLabel = 'esta ação') {
    if (!this.adminEnabled) {
      this.setStatus(`Acesso negado: somente o admin pode executar ${actionLabel}.`, 'danger');
      alert('Acesso negado: somente o admin pode executar esta ação.');
      return false;
    }
    return true;
  },


  buildDaySnapshotObject() {
    const obj = {};
    for (const [routeId, r] of this.routes.entries()) {
      obj[routeId] = this.serializeRoute(r);
    }
    return obj;
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

  markCloudDirty() {
    this.cloudDirty = true;
    // Debounce: não salva a cada bipagem
    if (this.cloudSaveTimer) clearTimeout(this.cloudSaveTimer);
    this.cloudSaveTimer = setTimeout(() => {
      this.flushCloudSave().catch(e => {
        console.warn('Falha ao salvar no Supabase (vai tentar de novo depois):', e);
        this.setStatus('Falha ao salvar no banco (Supabase). Mantido no cache local.', 'warning');
      });
    }, 1200);
  },

  async flushCloudSave() {
    if (!this.cloudEnabled) return;
    if (this.cloudSaving) return;
    if (!this.cloudDirty) return;

    const op = this.getOperationCode();
    const day = this.workDay || this.todayLocalISO();
    if (!op || !day) return;

    this.cloudSaving = true;
    try {
      const snapshot = this.buildDaySnapshotObject();
      await this.supaSaveDaySnapshot(op, day, snapshot);
      this.cloudDirty = false;
      this.cloudLastSaveAt = Date.now();
      this.setStatus(`Salvo no banco • ${op} • ${day}`, 'success');
    } finally {
      this.cloudSaving = false;
    }
  },

  mergeSnapshotIntoLocal(snapshotObj) {
    // snapshotObj: { [routeId]: serializedRoute }
    if (!snapshotObj || typeof snapshotObj !== 'object') return;

    for (const [routeId, ser] of Object.entries(snapshotObj)) {
      const id = String(routeId);
      const existing = this.routes.get(id);

      if (!existing) {
        const r = this.deserializeRoute(id, ser);
        // garante faltantes
        if (!r.faltantes.size && r.ids.size) {
          r.faltantes = new Set(r.ids);
          for (const c of r.conferidos) r.faltantes.delete(c);
        }
        this.routes.set(id, r);
        continue;
      }

      // merge: união dos conjuntos e mapas, priorizando timestamps mais recentes
      const tmp = this.deserializeRoute(id, ser);

      tmp.ids.forEach(v => existing.ids.add(v));
      tmp.conferidos.forEach(v => existing.conferidos.add(v));
      tmp.foraDeRota.forEach(v => existing.foraDeRota.add(v));
      tmp.faltantes.forEach(v => existing.faltantes.add(v));

      for (const [k, v] of tmp.timestamps.entries()) {
        const cur = existing.timestamps.get(k);
        if (!cur || String(v) > String(cur)) existing.timestamps.set(k, v);
      }
      for (const [k, v] of tmp.duplicados.entries()) {
        const cur = existing.duplicados.get(k);
        if (!cur) existing.duplicados.set(k, v);
        else existing.duplicados.set(k, Array.from(new Set([].concat(cur, v))));
      }

      // recomputa faltantes
      if (existing.ids.size) {
        existing.faltantes = new Set(existing.ids);
        for (const c of existing.conferidos) existing.faltantes.delete(c);
      }
    }
  },

  async syncFromSupabaseForDay(dayISO) {
    const op = this.getOperationCode();
    if (!op) return;

    try {
      const row = await this.supaLoadDaySnapshot(op, dayISO);
      if (!row || !row.data) return;

      // Mescla do banco -> local
      this.mergeSnapshotIntoLocal(row.data);

      // Persistir merge local
      this.saveToStorage(dayISO);

      this.setStatus(`Sincronizado do banco • ${op} • ${dayISO}`, 'info');
    } catch (e) {
      // Se a tabela não existir, orientar com SQL
      if (e && (e.code === '42P01' || /routes_state/i.test(String(e.message || '')))) {
        this.setStatus('Tabela routes_state não existe no Supabase. Crie a tabela para sincronizar.', 'danger');
        console.warn('Crie no Supabase:', this.getRoutesStateCreateSQL());
        return;
      }
      console.warn('Falha ao sincronizar do Supabase:', e);
    }
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
    // Carrega lista de operações ativas e força seleção se não houver
    const sb = this.getSb();
    if (!sb) return;

    const { data, error } = await sb.from('operations').select('code,name,active').eq('active', true).order('code', { ascending: true });
    if (error) {
      console.warn('Falha ao carregar operações:', error);
      return;
    }

    // Preenche select
    const $sel = $('#op-select');
    if ($sel.length) {
      $sel.empty();
      (data || []).forEach(op => {
        const label = op.name ? `${op.code} — ${op.name}` : op.code;
        $sel.append(`<option value="${op.code}">${label}</option>`);
      });
    }

    const current = this.getOperationCode();
    const exists = (data || []).some(o => o.code === current);

    if (!current || !exists) {
      // abre modal
      $('#modal-operation').modal({ backdrop: 'static', keyboard: false });
    } else {
      this.setOperationCode(current);
    }
  },

// =======================
// Admin (cadastrar operações) - requer login apenas para o admin
// =======================
async adminSignIn(email, password) {
  const sb = this.getSb();
  if (!sb) throw new Error('Supabase client não encontrado.');
  const { data, error } = await sb.auth.signInWithPassword({ email, password });
  if (error) throw error;
    this.setAdminEnabled(true);
    return data;
},

  async adminSignOut() {
    const sb = this.getSb();
    if (!sb) throw new Error('Supabase client não encontrado.');
    const { error } = await sb.auth.signOut();
    if (error) throw error;
    this.setAdminEnabled(false);
  },


async adminUpsertOperation(code, name, active = true) {
  const sb = this.getSb();
  if (!sb) throw new Error('Supabase client não encontrado.');
  const op = {
    code: String(code || '').trim().toUpperCase(),
    name: String(name || '').trim() || null,
    active: !!active
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



  setStatus(txt, kind='muted') {
    const $s = $('#sync-status');
    $s.removeClass('text-muted text-success text-danger text-warning');
    $s.addClass(`text-${kind}`);
    $s.text(txt);
  },


  // =======================
  // =======================
  markDirty(reason = '') {
    this.dirty = true;
    // Supabase: marca para salvar snapshot
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

  pushEvent(evt) {
    // evt: {ts, type: 'ok'|'fora'|'dup', code, currentRouteId, correctRouteId}
    this.lastEvents.unshift(evt);
    if (this.lastEvents.length > 80) this.lastEvents.length = 80;
    this.renderAcompanhamento();
  },

  renderAcompanhamento() {
    // badge pendente
    if (this.dirty) $('#dirty-flag').removeClass('d-none');
    else $('#dirty-flag').addClass('d-none');

    // === Resumo por CLUSTER (consolidado) ===
    const mapa = new Map(); // cluster -> { conferidos, total }
    for (const r of this.routes.values()) {
      const key = (r.cluster && String(r.cluster).trim()) ? String(r.cluster).trim() : '(sem cluster)';
      const total = (r.totalInicial || r.ids.size || 0);
      const conf = (r.conferidos ? r.conferidos.size : 0);

      if (!mapa.has(key)) mapa.set(key, { conferidos: 0, total: 0 });
      const acc = mapa.get(key);
      acc.conferidos += conf;
      acc.total += total;
    }

    const resumo = Array.from(mapa.entries())
      .sort((a, b) => String(a[0]).localeCompare(String(b[0])))
      .slice(0, 14)
      .map(([cluster, v]) => {
        const ok = (v.total > 0 && v.conferidos === v.total) ? ' ✅' : '';
        return `CLUSTER ${cluster}: ${v.conferidos}/${v.total}${ok}`;
      })
      .join('<br>');

    $('#acompanhamento-resumo').html(
      resumo || '<span class="text-muted">sem clusters</span>'
    );

    // === Log de bipagens ===
    const items = this.lastEvents.slice(0, 30).map(ev => {
      const d = new Date(ev.ts);
      const hh = String(d.getHours()).padStart(2,'0');
      const mm = String(d.getMinutes()).padStart(2,'0');
      const ss = String(d.getSeconds()).padStart(2,'0');
      const base = `${hh}:${mm}:${ss} • ${ev.code}`;

      if (ev.type === 'fora') {
        const rr = ev.correctRouteId ? this.routes.get(String(ev.correctRouteId)) : null;
        const cl = rr && rr.cluster ? String(rr.cluster).trim() : '';
        const extra = (cl ? ` <small class="text-muted">(CLUSTER ${cl})</small>` : ' <small class="text-muted">(cluster desconhecido)</small>');
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
  // Relatório Noturno (sidebar)
  // =======================
  escHtml(s) {
    return String(s ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  },

  locateIdInOtherRoutes(id, excludeRouteId) {
    const ex = String(excludeRouteId ?? '');
    // prioridade: conferido > fora de rota > pertence (ids/faltantes)
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
      const badge = it.ok ? `<span class="badge badge-success ml-2">100%</span>` : `<span class="badge badge-warning ml-2">${esc(it.falt) } falt.</span>`;
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
          faltantes.sort((a,b)=>String(a).localeCompare(String(b)));

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

      timestamps: new Map(), // id -> epoch (ms)
      ids: new Set(),
      faltantes: new Set(),
      conferidos: new Set(),
      foraDeRota: new Set(),
      duplicados: new Map(), // id -> count

      totalInicial: 0,

      // vínculo Carreta/QR
      plateKey: '',
      plateRaw: '',
      plateLicense: '',
      routeQrKey: '',
      routeQrRaw: '',

      // timestamps de bipagem (para export CSV estilo scanner)
      plateScanTs: 0,
      routeQrScanTs: 0
    };
  },

  get current() {
    if (!this.currentRouteId) return null;
    return this.routes.get(String(this.currentRouteId)) || null;
  },

  // =======================
  // Persistência local (cache por dia)
  // =======================
  loadFromStorage(dayISO) {
    try {
      
      // sempre limpa estado do dia anterior, mesmo se não houver cache
const key = this.storageKeyForDay(dayISO);
      const raw = localStorage.getItem(key);
      if (!raw) return;

      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed !== 'object') return;

      this.routes.clear();
      this.currentRouteId = null;

      for (const [routeId, r] of Object.entries(parsed)) {
        const route = this.makeEmptyRoute(routeId);

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

    
    route.plateScanTs = Number(r.plateScanTs || 0);
    route.routeQrScanTs = Number(r.routeQrScanTs || 0);

    

        (r.ids || []).forEach(id => route.ids.add(id));
        (r.conferidos || []).forEach(id => route.conferidos.add(id));
        (r.foraDeRota || []).forEach(id => route.foraDeRota.add(id));
        route.faltantes = new Set(r.faltantes || []);

        route.timestamps = new Map(Object.entries(r.timestamps || {}).map(([k, v]) => [k, v]));
        route.duplicados = new Map(Object.entries(r.duplicados || {}).map(([k, v]) => [k, v]));

        if (!route.faltantes.size && route.ids.size) {
          route.faltantes = new Set(route.ids);
          for (const c of route.conferidos) route.faltantes.delete(c);
        }

        this.routes.set(String(routeId), route);
      }
    } catch (e) {
      console.warn('Falha ao carregar storage:', e);
    }
  },

  saveToStorage(dayISO) {
    try {
      const key = this.storageKeyForDay(dayISO);
      const obj = {};
      for (const [routeId, r] of this.routes.entries()) {
        obj[routeId] = this.serializeRoute(r);
      }
      localStorage.setItem(key, JSON.stringify(obj));
      this.markCloudDirty();
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

    (r.ids || []).forEach(id => route.ids.add(id));
    (r.conferidos || []).forEach(id => route.conferidos.add(id));
    (r.foraDeRota || []).forEach(id => route.foraDeRota.add(id));
    route.faltantes = new Set(r.faltantes || []);

    route.timestamps = new Map(Object.entries(r.timestamps || {}).map(([k, v]) => [k, v]));
    route.duplicados = new Map(Object.entries(r.duplicados || {}).map(([k, v]) => [k, v]));

    if (!route.faltantes.size && route.ids.size) {
      route.faltantes = new Set(route.ids);
      for (const c of route.conferidos) route.faltantes.delete(c);
    }

    return route;
  },

  globalCleanupForaDeRotaForConferidos() {
    // mapa: id -> rota onde está conferido (pode existir mais de uma, mas tratamos como "válido" em qualquer)
    const conferidosIds = new Set();
    for (const r of this.routes.values()) {
      for (const id of r.conferidos) conferidosIds.add(id);
    }

    if (!conferidosIds.size) return;

    for (const r of this.routes.values()) {
      for (const id of conferidosIds) {
        if (r.conferidos.has(id)) {
          // se conferido aqui, garante que não esteja fora de rota aqui
          if (r.foraDeRota.has(id)) r.foraDeRota.delete(id);
        } else {
          // se não é conferido aqui, não deve ficar preso como fora de rota aqui se ele foi validado em outra rota
          if (r.foraDeRota.has(id)) r.foraDeRota.delete(id);
          if (r.duplicados.has(id)) r.duplicados.delete(id);
        }
      }
    }
  },

  // =======================
  // Troca de dia
  // =======================
  
async applyWorkDay(dayISO) {
  this.workDay = dayISO;
  $('#work-day').val(dayISO);

  // garante operação selecionada (modal se necessário)
  await this.ensureOperationSelected();

  const op = this.getOperationCode();
  if (op) $('#op-badge').text(op);

  // 1) sempre começa pelo cache local do dia
  // limpa estado antes de carregar novo dia
  this.routes.clear();
  this.currentRouteId = null;
  this.loadFromStorage(dayISO);

  // 2) tenta sincronizar do banco (Supabase) e mesclar no local
  if (op) {
    await this.syncFromSupabaseForDay(dayISO);
  }

  // 3) render
  this.renderRoutesSelects();
  this.refreshUIFromCurrent();
  this.renderAcompanhamento();

  if (op) this.setStatus(`dia carregado • ${op} • ${dayISO}`, 'success');
  else this.setStatus('dia carregado (sem operação)', 'warning');
},



  
  resetForOperationChange() {
    // limpa estado e força voltar para a tela inicial
    this.routes.clear();
    this.currentRouteId = null;
    this.viaCsv = false;

    // limpa selects/listas visuais mais comuns
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

    // alterna interfaces
    $('#global-interface').addClass('d-none');
    $('#carreta-interface').addClass('d-none');
    $('#manual-interface').addClass('d-none');
    $('#initial-interface').removeClass('d-none');
  },

  computeSnapshotStats(snapshotObj) {
    // snapshotObj: { routeId: serializedRoute, ... }
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

    // lista operações ativas
    const { data: ops, error: e1 } = await sb
      .from('operations')
      .select('code,name,active')
      .eq('active', true)
      .order('code', { ascending: true });
    if (e1) throw e1;

    // para cada operação, pega snapshot do dia (se existir)
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

    const sorted = (items || []).slice().sort((a,b)=> (a.code||'').localeCompare(b.code||''));

    for (const it of sorted) {
      const u = it.updated_at ? new Date(it.updated_at).toLocaleString('pt-BR') : '-';
      $tb.append(`
        <tr>
          <td><strong>${it.code}</strong>${it.name ? ` <span class="text-muted small">(${this.escapeHtml(it.name)})</span>` : ''}</td>
          <td>${it.stats.routesCount}</td>
          <td>${it.stats.totalIds}</td>
          <td>${it.stats.conferidos}</td>
          <td>${it.stats.faltantes}</td>
          <td>${it.stats.fora}</td>
          <td>${u}</td>
        </tr>
      `);

      // log simplificado
      $log.append(`
        <li class="list-group-item d-flex justify-content-between align-items-center">
          <span><strong>${it.code}</strong> • ${it.stats.conferidos}/${it.stats.totalIds} conferidos • ${it.stats.faltantes} faltantes</span>
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
    const id = String(routeId);
    if (!this.routes.has(id)) {
      alert('Rota não encontrada.');
      return;
    }
    this.currentRouteId = id;
    this.renderRoutesSelects();
    this.refreshUIFromCurrent();
    this.saveToStorage(this.workDay);
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

    // cache para filtro (somente o dropdown "in-app")
    this._routesDropdownCache = routesSorted.map(r => ({
      routeId: String(r.routeId),
      label: makeLabel(r),
      clusterKey: this.normalizeCluster(r.cluster || ''),
      labelKey: this.normalizeCluster(makeLabel(r)),
    }));

    // Select da tela inicial (sem filtro)
    $sel1.html(
      ['<option value="">(Nenhuma selecionada)</option>']
        .concat(this._routesDropdownCache.map(x => `<option value="${x.routeId}">${x.label}</option>`))
        .join('')
    );

    // Select "Trocar rota" (com filtro por cluster)
    this.applyRouteDropdownFilter($('#route-search').val() || '');

    if (this.currentRouteId) {
      $sel1.val(this.currentRouteId);
      // o $sel2 é definido dentro do filtro (se estiver visível na lista filtrada)
    }
  },

  applyRouteDropdownFilter(filterText) {
    const $sel2 = $('#saved-routes-inapp');

    // se ainda não carregou rotas, só zera
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
      .concat(filtered.map(x => `<option value="${x.routeId}">${x.label}</option>`))
      .join('');

    $sel2.html(options);

    // preserva seleção atual, se ainda estiver na lista filtrada
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

    $('#route-title').html(`ROTA: <strong>${r.routeId}</strong>`);
    $('#cluster-title').html(r.cluster ? `CLUSTER: <strong>${r.cluster}</strong>` : '');
    $('#destination-facility-title').html(r.destinationFacilityId ? `<strong>XPT:</strong> ${r.destinationFacilityId}` : '');
    $('#destination-facility-name').html(r.destinationFacilityName ? `<strong>DESTINO:</strong> ${r.destinationFacilityName}` : '');

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
      Array.from(r.conferidos).map(id => `<li class='list-group-item list-group-item-success'>${id}</li>`).join('')
    );

    $('#faltantes-list').html(
      `<h6>Faltantes (<span class='badge badge-danger'>${r.faltantes.size}</span>)</h6>` +
      Array.from(r.faltantes).map(id => `<li class='list-group-item list-group-item-danger'>${id}</li>`).join('')
    );

    $('#fora-rota-list').html(
      `<h6>Fora de Rota (<span class='badge badge-warning'>${r.foraDeRota.size}</span>)</h6>` +
      Array.from(r.foraDeRota).map(id => {
        const correct = this.findCorrectRouteForId(id);
        if (correct && String(correct) !== String(this.currentRouteId)) {
          const rr = this.routes.get(String(correct));
          const cl = rr && rr.cluster ? String(rr.cluster).trim() : '';
          const extra = cl
            ? ` <small class="text-muted">(CLUSTER ${cl})</small>`
            : ` <small class="text-muted">(cluster desconhecido)</small>`;
          return `<li class='list-group-item list-group-item-warning'>${id}${extra}</li>`;
        }
        return `<li class='list-group-item list-group-item-warning'>${id}</li>`;
      }).join('')
    );
        
    $('#duplicados-list').html(
      `<h6>Duplicados (<span class='badge badge-secondary'>${r.duplicados.size}</span>)</h6>` +
      Array.from(r.duplicados.entries())
        .map(([id, count]) => `<li class='list-group-item list-group-item-secondary'>${id} <span class="badge badge-dark ml-2">${count}x</span></li>`)
        .join('')
    );

    $('#verified-total').text(r.conferidos.size);
    this.atualizarProgresso();
  },
  deleteRoute(routeId) {
    if (!this.requireAdminSync("excluir rota")) return;
    if (!routeId) return;
    this.routes.delete(String(routeId));
    if (this.currentRouteId === String(routeId)) this.currentRouteId = null;

    // cache local (rápido)
    this.saveToStorage(this.workDay);
    this.markDirty('excluir rota');

    this.renderRoutesSelects();
    this.refreshUIFromCurrent();
    this.renderAcompanhamento();
  },
  clearAllRoutes() {
    if (!this.requireAdminSync("limpar rotas do dia")) return;
    this.routes.clear();
    this.currentRouteId = null;

    // remove cache do dia
    localStorage.removeItem(this.storageKeyForDay(this.workDay));
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
  // Leitura inteligente (Placa/QR Rota/ID)
  // =======================
  parseScanPayload(raw) {
    const cleaned = String(raw || '').trim();
    if (!cleaned) return { kind: 'empty' };

    // 0) QR em formato “caret” (ex.: ^license_plate^Ç^SHW8J07^, ...)
// Esse formato vem do leitor; transformamos em KV tolerante (aceita "âssignment", "asssignment", etc).
const firstLine = cleaned.split(/\?\n/)[0].trim();
if (firstLine.includes('^') && firstLine.includes('Ç')) {
  const kv = this.parseCaretKV(firstLine);

  // QR PLACA (carreta)
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

  // QR ROTA (container / assignment)
  if (kv.container_id || kv.assignment) {
    let assignment = kv.assignment ? String(kv.assignment).trim() : '';
    assignment = this.normalizeCluster(assignment); // normaliza (remove lixo e deixa comparável ao cluster)

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


  // 1) tenta JSON (QR da placa / QR da rota)
    if (cleaned.startsWith('{') && cleaned.endsWith('}')) {
      try {
        const obj = JSON.parse(cleaned);

        // QR PLACA (carreta)
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

        // QR ROTA (container/assignment/etc)
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
      } catch (e) {
        // cai para parse normal
      }
    }

    // 2) tenta ID de envio (padrão do sistema atual)
    const shipmentId = this.normalizarCodigo(cleaned);
    if (shipmentId) return { kind: 'shipment', shipmentId };

    // 3) fallback: placa em texto
    const plateLike = cleaned.toUpperCase().replace(/[^A-Z0-9]/g, '');
    if (/^[A-Z]{3}\d[A-Z]\d{2}$/.test(plateLike) || /^[A-Z]{3}\d{4}$/.test(plateLike)) {
      return {
        kind: 'plate',
        plateKey: plateLike,
        plate: { raw: cleaned, license_plate: plateLike, carrier_name: '', vehicle_type_description: '' }
      };
    }

    // 4) rota numérica simples
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
    if (!this.carretas.routesTs.get(routeKey)) this.carretas.routesTs.set(routeKey, Date.now());

    // tenta vincular também ao objeto de rota (se existir no sistema)
    // IMPORTANTE: no seu cenário o QR da rota NÃO traz routeId; ele traz "assignment" (ex.: J9_AM3), que corresponde ao CLUSTER.
    // Então vinculamos por cluster (e o dia já é isolado pelo cache do sistema).
    const assignMatch = String(routeKey).match(/^assignment:(.+)$/);
    const clusterCandidate = this.normalizeCluster(assignMatch?.[1] || '');

    let linked = 0;

    // 1) Principal: vincular por CLUSTER = assignment
    if (clusterCandidate) {
      for (const r of this.routes.values()) {
        const c = this.normalizeCluster(r.cluster);
        if (c && clusterCandidate && c === clusterCandidate) {
          r.plateKey = plateKey;
          r.plateRaw = plate.jsonText || plate.raw || '';
          r.plateLicense = plate.license_plate || plateKey;
          r.routeQrKey = routeKey;
          r.routeQrRaw = (typeof routeRaw === 'string' ? routeRaw : (routeRaw && routeRaw.raw) ? routeRaw.raw : '') || '';
          // se houver jsonText no parse, usa no export
          const _json = (routeRaw && routeRaw.jsonText) ? routeRaw.jsonText : '';
          if (_json) r.routeQrRaw = _json;
          
          r.plateScanTs = Number(plate.tsLast || Date.now());
          r.routeQrScanTs = Date.now();
linked++;
        }
      }
    }

    // 2) Fallback: se por algum motivo vier um routeId em outro formato, tenta achar diretamente pelo Map
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
          // se houver jsonText no parse, usa no export
          const _json = (routeRaw && routeRaw.jsonText) ? routeRaw.jsonText : '';
          if (_json) r.routeQrRaw = _json;
          
          r.plateScanTs = Number(plate.tsLast || Date.now());
          r.routeQrScanTs = Date.now();
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
  routes.sort((a,b)=>String(a).localeCompare(String(b)));

  // quais clusters importados existem
  const clustersImportados = new Set(
    Array.from(this.routes.values()).map(r => this.normalizeCluster(r.cluster)).filter(Boolean)
  );

  // checa para cada routeKey se o cluster existe nas rotas importadas
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

    // 1) remove vínculo das rotas QR na estrutura carretas
    const routeKeys = Array.from(p.routes || []);
    for (const rk of routeKeys) {
      this.carretas.routeToPlate.delete(rk);
      this.carretas.routesRaw.delete(rk);
      this.carretas.routesJson.delete(rk);
      this.carretas.routesTs.delete(rk);
    }

    // zera lista de rotas na placa
    p.routes = new Set();

    // 2) limpa também nas rotas do sistema (conferência)
    for (const r of this.routes.values()) {
      if ((r.plateKey || '').toUpperCase() === plateKey) {
        r.plateKey = '';
        r.plateRaw = '';
        r.plateLicense = '';
        r.routeQrKey = '';
        r.routeQrRaw = '';
        r.plateScanTs = 0;
        r.routeQrScanTs = 0;
      }
    }

    // se a placa ativa era essa, mantém ativa mas sem rotas
    if (this.carretas.currentPlateKey === plateKey) {
      this.carretas.currentPlateKey = plateKey;
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
    meta.push(`<strong>${p.license_plate}</strong>`);
    if (p.vehicle_type_description) meta.push(`<span class="text-muted">(${p.vehicle_type_description})</span>`);
    if (p.carrier_name) meta.push(`<span class="text-muted">• ${p.carrier_name}</span>`);
    $cur.html(meta.join(' '));

    const routesArr = Array.from(p.routes);
    routesArr.sort((a, b) => String(a).localeCompare(String(b)));

    $list.html(
      routesArr.map(rk => `<li class="list-group-item">${rk}</li>`).join('') ||
      '<li class="list-group-item text-muted">sem rotas nessa placa</li>'
    );

    $sum.text(`${routesArr.length} rota(s) vinculada(s)`);
      this.renderCarretaProgress();
  },


  renderCarretaProgress() {
    // Atualiza o card "Acompanhamento da Placa" (esperadas no dia × bipadas na placa ativa)
    const $label = $('#carreta-progress-label');
    const $pct = $('#carreta-progress-percent');
    const $bar = $('#carreta-progress-bar');
    const $missing = $('#carreta-missing-list');
    const $extra = $('#carreta-extra-list');

    // Se o HTML não tiver esses elementos (caso futuro), não faz nada.
    if (!$label.length || !$pct.length || !$bar.length) return;

    const plateKey = this.carretas.currentPlateKey;

    const setUi = (done, total) => {
      const pctVal = total > 0 ? Math.round((done / total) * 100) : 0;
      $label.text(`${done}/${total}`);
      $pct.text(`${pctVal}%`);
      $bar.css('width', `${pctVal}%`);
      $bar.attr('aria-valuenow', String(pctVal));
    };

    // ✅ Rotas "esperadas" = rotas IMPORTADAS (routeId) no dia.
    // Isso permite ver pendências mesmo que você ainda não tenha bipado QR de rota.
    const expected = Array.from(this.routes.keys()).map(String);
    expected.sort((a, b) => (Number(a) - Number(b)) || String(a).localeCompare(String(b)));

    if (!plateKey) {
      setUi(0, expected.length);
      $missing.html('<li class="list-group-item text-muted">bipe uma placa para ver o acompanhamento</li>');
      $extra.html('<li class="list-group-item text-muted">—</li>');
      return;
    }

    const plate = this.carretas.plates.get(plateKey);

    // Rotas importadas já VINCULADAS a essa placa (por cluster via QR, ou por ligação direta)
    const linked = expected.filter(routeId => {
      const r = this.routes.get(String(routeId));
      return r && String(r.plateKey || '') === String(plateKey);
    });

    const missing = expected.filter(routeId => !linked.includes(routeId));

    // "Extra" aqui vira: QRs de rota bipados nessa placa que NÃO batem com nenhuma rota importada.
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
      missing.map(routeId => `<li class="list-group-item">${fmtExpected(routeId)}</li>`).join('') ||
      '<li class="list-group-item text-muted">nada faltando 🎉</li>'
    );

    $extra.html(
      extra.map(rk => `<li class="list-group-item">${fmtExtra(rk)}</li>`).join('') ||
      '<li class="list-group-item text-muted">—</li>'
    );
  },


  playAlertSound() {
    try {
      const audio = new Audio('mixkit-alarm-tone-996-_1_.mp3');
      audio.play().catch(() => {});
    } catch {}
  },

  // =======================
  // Fora de rota inteligente (global)
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
      this.cleanupIdFromOtherRoutes(codigo, this.currentRouteId);
    }

    // duplicata se já conferido OU já fora de rota nessa rota
    if (r.conferidos.has(codigo) || r.foraDeRota.has(codigo)) {
      const count = r.duplicados.get(codigo) || 1;
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

      // se entrou como conferido aqui, remove fora de rota em outras rotas
      this.cleanupIdFromOtherRoutes(codigo, this.currentRouteId);

      $('#barcode-input').val('').focus();
      this.pushEvent({ ts: now, type: 'ok', code: codigo, currentRouteId: this.currentRouteId, correctRouteId });
      this.markDirty('bipagem');

      this.saveToStorage(this.workDay);
      this.atualizarListas();
      return;
    }

    // fora de rota
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
  // Importação HTML: várias rotas
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
      const route = this.routes.get(routeId) || this.makeEmptyRoute(routeId);

      const clusterMatch = /"cluster":"([^"]+)"/.exec(block);
      if (clusterMatch) route.cluster = this.normalizeCluster(clusterMatch[1]);

      const facMatch = /"destinationFacilityId":"([^"]+)","name":"([^"]+)"/.exec(block);
      if (facMatch) {
        route.destinationFacilityId = facMatch[1];
        route.destinationFacilityName = facMatch[2];
      }

      const idsExtraidos = new Set();

      // Extrai IDs de envio de forma tolerante (sem depender de receiver_id).
      // Padrão: "id": 11 dígitos (na prática começam com 4...)
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

    this.saveToStorage(this.workDay);

    this.markDirty('import HTML');

    this.renderRoutesSelects();
    this.renderAcompanhamento();
    return imported;
  },

  // =======================
  // EXPORT: rota atual CSV (FORMATO ANTIGO PADRÃO)
  // =======================

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

// =======================
  // EXPORT: CSV com PLACA + QR ROTA (novo)
  // =======================
  // CSV estilo "scanner" (igual exemplo)
  // Cada linha é um evento com colunas: date,time,time_zone,format,text,notes,favorite,date_utc,time_utc,metadata
  // =======================
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

    const date = `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}`;
    const time = `${pad2(d.getHours())}:${pad2(d.getMinutes())}:${pad2(d.getSeconds())}`;

    // UTC
    const iso = d.toISOString();
    const dateUtc = iso.slice(0, 10);
    const timeUtc = iso.split('T')[1].split('.')[0];

    const tzLabel = 'Horário Padrão do Amazonas';

    // IMPORTANT: o validador espera tudo entre aspas e separador vírgula
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

  // Monta as linhas (scanner) de uma rota: PLACA -> QR ROTA -> IDs (por timestamp)
    // Monta as linhas (scanner) de uma rota: PLACA -> QR ROTA -> IDs (por timestamp)
  // Formato validado:
  // - linha de PLACA: QR Code com JSON contendo "license_plate"
  // - linha de ROTA: QR Code com JSON contendo container_id/facility_id/assignment quando houver
  // - linhas de IDs: QR Code com JSON {"id":"<id>","t":"lm"}
  buildScannerCsvLinesForRoute(r) {
    const lines = [];
    if (!r) return lines;

    const ids = this.getIdsForExportByTimestamp(r);

    const firstIdTs = ids.length ? (r.timestamps?.get(ids[0]) || Date.now()) : Date.now();
    const plateTs = r.plateScanTs || firstIdTs;
    const routeTs = r.routeQrScanTs || (plateTs ? (Number(plateTs) + 1) : firstIdTs);

    // ===== PLACA (JSON deve conter license_plate) =====
    let plateText = '';
    if (r.plateKey) {
      const p = this.carretas?.plates?.get(r.plateKey);
      if (p && p.jsonText) {
        plateText = String(p.jsonText);
      } else {
        // fallback mínimo (mas compatível com o validador)
        plateText = JSON.stringify({
          id: (p && p.id) ? Number(p.id) : undefined,
          carrier_id: (p && p.carrier_id) ? Number(p.carrier_id) : undefined,
          carrier_name: (p && p.carrier_name) ? String(p.carrier_name) : undefined,
          license_plate: String((p && p.license_plate) || r.plateLicense || r.plateKey),
          vehicle_type_description: (p && p.vehicle_type_description) ? String(p.vehicle_type_description) : undefined,
          vehicle_type_id: (p && p.vehicle_type_id) ? Number(p.vehicle_type_id) : undefined,
          tracking_provider_ids: []
        }, (k, v) => (v === undefined ? undefined : v));
        // remove undefined
        plateText = plateText.replace(/,\s*"(?:id|carrier_id|carrier_name|vehicle_type_description|vehicle_type_id)"\s*:\s*null/g, '');
      }
    }

    // ===== QR DA ROTA (JSON) =====
    let routeText = '';
    if (r.routeQrKey) {
      // tenta pegar json salvo na estrutura carretas
      const routeJson = this.carretas?.routesJson?.get(r.routeQrKey);
      if (routeJson) {
        routeText = String(routeJson);
      } else if (r.routeQrRaw && String(r.routeQrRaw).trim().startsWith('{')) {
        routeText = String(r.routeQrRaw).trim();
      } else {
        // fallback: monta JSON com assignment
        routeText = JSON.stringify({
          container_id: (r.container_id != null) ? Number(r.container_id) : undefined,
          facility_id: (r.destinationFacilityId || ''),
          assignment: String(r.routeQrKey).replace(/^assignment:/, '')
        }, (k, v) => (v === undefined ? undefined : v));
      }
    }

    if (plateText) lines.push(this.buildScannerCsvRow(plateTs, 'QR Code', plateText, ''));
    if (routeText) lines.push(this.buildScannerCsvRow(routeTs, 'QR Code', routeText, ''));

    // ===== IDs: QR Code + JSON com id =====
    for (const id of ids) {
      const ts = r.timestamps?.get(id) || Date.now();
      const payload = JSON.stringify({ id: String(id), t: 'lm' });
      lines.push(this.buildScannerCsvRow(ts, 'QR Code', payload, ''));
    }

    return lines;
  },

  // - 1 arquivo por rota (rota atual)
  // - 1 arquivo com todas as rotas
  // - 1 arquivo "mapa" (placa -> rotas qrs)
  // =======================
  
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

    // Agrupa por PLACA
    const plateGroups = new Map(); // plateKey -> [routeObj]
    for (const r of this.routes.values()) {
      if (!r.plateKey || !r.routeQrKey) continue; // só rotas vinculadas
      if (!plateGroups.has(r.plateKey)) plateGroups.set(r.plateKey, []);
      plateGroups.get(r.plateKey).push(r);
    }

    if (!plateGroups.size) {
      return alert('Nenhuma rota está vinculada a PLACA/QR de rota. Use a tela da CARRETA primeiro.');
    }

    const lines = [];
    lines.push(this.buildScannerCsvHeader());

    const plateKeys = Array.from(plateGroups.keys()).sort((a,b)=>String(a).localeCompare(String(b)));

    for (const plateKey of plateKeys) {
      const routesArr = plateGroups.get(plateKey) || [];
      routesArr.sort((a,b) => {
        const ca = String(a.cluster||'').localeCompare(String(b.cluster||''));
        if (ca !== 0) return ca;
        return String(a.routeId||'').localeCompare(String(b.routeId||''));
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
    const stamp = `${now.getFullYear()}-${this.pad2(now.getMonth()+1)}-${this.pad2(now.getDate())}_${this.pad2(now.getHours())}${this.pad2(now.getMinutes())}`;
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
        const rawEsc = `"${String(raw).replace(/"/g,'""')}"`;
        linhas.push(`${plateKey},${carrier},${vt},${rk},${rawEsc}`);
      }
    }

    const conteudo = [header, ...linhas].join('\r\n');
    const blob = new Blob([conteudo], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);

    const now = new Date();
    const stamp = `${now.getFullYear()}-${this.pad2(now.getMonth()+1)}-${this.pad2(now.getDate())}_${this.pad2(now.getHours())}${this.pad2(now.getMinutes())}`;
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
      const date = `${lidaEm.getFullYear()}-${pad2(lidaEm.getMonth()+1)}-${pad2(lidaEm.getDate())}`;
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

  // =======================
  // EXPORT: todas as rotas XLSX
  // 1 aba ("Bipagens"), 1 coluna por rota, só IDs conferidos
  // =======================
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
    const stamp = `${now.getFullYear()}-${this.pad2(now.getMonth()+1)}-${this.pad2(now.getDate())}_${this.pad2(now.getHours())}${this.pad2(now.getMinutes())}`;

    XLSX.writeFile(wb, `bipagens_todas_rotas_${this.workDay || this.todayLocalISO()}_${stamp}.xlsx`);
  }
};

// =======================
// Eventos / Boot
// =======================

$(document).ready(async () => {
  const today = ConferenciaApp.todayLocalISO();
  $('#work-day').val(today);

  // força seleção de operação ao entrar
  await ConferenciaApp.ensureOperationSelected();

  // detecta se admin já está logado (persistência do supabase)
  try {
    const sb = ConferenciaApp.getSb();
    if (sb) {
      const { data } = await sb.auth.getUser();
      ConferenciaApp.setAdminEnabled(!!data?.user);
      sb.auth.onAuthStateChange((_event, session) => {
        ConferenciaApp.setAdminEnabled(!!session?.user);
      });
    }
  } catch {}

  // se já tiver operação, carrega o dia
  if (ConferenciaApp.getOperationCode()) {
    await ConferenciaApp.applyWorkDay(today);
  }
});

// Confirmar operação escolhida
$(document).on('click', '#btn-op-confirm', async () => {
  const code = String($('#op-select').val() || '').trim().toUpperCase();
  if (!code) return;
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


// Filtro de rotas (pesquisa por cluster)
$(document).on('input', '#route-search', (e) => {
  ConferenciaApp.applyRouteDropdownFilter(e.target.value);
});

// Troca de dia
$(document).on('change', '#work-day', async (e) => {
  const day = e.target.value;
  if (!day) return;
  await ConferenciaApp.applyWorkDay(day);
});

// =======================
// Relatório noturno (sidebar)
// =======================
$(document).on('click', '#finish-night-btn', () => {
  try {
    ConferenciaApp.showNightReport();
    // garante que o conteúdo do relatório fique visível
    const el = document.querySelector('#night-report');
    if (el) el.scrollIntoView({ block: 'start', behavior: 'smooth' });
  } catch (e) {
    console.error(e);
    alert('Falha ao gerar o relatório noturno.');
  }
});

// Botão vermelho da tela principal também pode abrir o relatório noturno
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

  // ✅ limpa o campo após importar
  $('#html-input').val('');

  alert(`${qtd} rota(s) importada(s) e salva(s)! Agora selecione e clique em "Carregar rota".`);
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
  // 1ª confirmação (sim/não)
  const ok1 = confirm(
    'ATENÇÃO: isso vai APAGAR TODAS as rotas do DIA selecionado.\n\n' +
    'Quer continuar?'
  );
  if (!ok1) return;

  // 2ª confirmação (digitação)
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
    const manualIds = $('#manual-input').val().split(/[\s,;]+/).map(x => x.trim()).filter(Boolean);

    if (!manualIds.length) return alert('Nenhum ID válido inserido.');

    const route = ConferenciaApp.routes.get(String(routeId)) || ConferenciaApp.makeEmptyRoute(routeId);
    route.cluster = cluster || route.cluster;

    for (const id of manualIds) {
      route.ids.add(id);
      if (!route.conferidos.has(id)) route.faltantes.add(id);
    }

    route.totalInicial = route.ids.size;
    ConferenciaApp.routes.set(String(routeId), route);

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

// Leitura do barcode (ENTER)
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

// Checar CSV (importa bipagens)
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

// Sem bind no finalizar
$('#back-btn').click(() => {
  // volta para a tela inicial sem perder conexão do arquivo
  $('#conference-interface').addClass('d-none');
  $('#manual-interface').addClass('d-none');
  $('#initial-interface').removeClass('d-none');

  // opcional: limpa campo de leitura e foca
  $('#barcode-input').val('');
  $('#html-input').focus();
});


// =======================
// Carretas: abrir/voltar + leitura
// =======================
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



// leitura na tela de carretas (scanner costuma enviar ENTER; alguns enviam via keydown)
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

  // Se cair aqui, pode ser ID de envio; nessa tela a gente só avisa para não misturar
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

$(document).on('keypress', '#carreta-input', (e) => {
  if (e.which === 13) {
    const raw = $('#carreta-input').val();
    $('#carreta-input').val('');
    processCarretaScan(raw);
  }
});

// suporte a colar (alguns leitores “digitam” e você cola do CSV)
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

// Exports novos (CSV com placa/rota)
$(document).on('click', '#export-csv-rota-atual-placa', () => {
  ConferenciaApp.exportRotaAtualCsvComPlacaERota();
});
$(document).on('click', '#export-csv-todas-rotas-placa', () => {
  ConferenciaApp.exportTodasRotasCsvComPlacaERota();
});
$(document).on('click', '#export-csv-mapa-carretas', () => {
  ConferenciaApp.exportMapaCarretasCsv();
});


// =======================
// Admin UI handlers
// =======================
$(document).on('click', '#btn-admin-open', async () => {
  $('#modal-admin').modal('show');
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
      $tbody.append(`<tr><td>${o.code}</td><td>${name}</td><td>${act}</td></tr>`);
    });
  } catch (e) {
    console.warn(e);
  }
}

$(document).on('click', '#btn-admin-login', async () => {
  const email = $('#admin-email').val();
  const pass = $('#admin-pass').val();
  try {
    await ConferenciaApp.adminSignIn(email, pass);
    $('#admin-status').text('Logado.');
    $('#admin-panel').removeClass('d-none');
    await refreshAdminOps();
  } catch (e) {
    console.error(e);
    alert('Falha no login do admin: ' + (e.message || e));
  }
});

$(document).on('click', '#btn-admin-logout', async () => {
  try {
    await ConferenciaApp.adminSignOut();
    $('#admin-status').text('Deslogado.');
    $('#admin-panel').addClass('d-none');
  } catch (e) {
    console.error(e);
  }
});

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




// Abrir acompanhamento geral (todas operações)
$(document).on('click', '#btn-global-acomp', async () => {
  try {
    const day = $('#work-day').val() || ConferenciaApp.todayLocalISO();
    // troca telas
    $('#initial-interface').addClass('d-none');
    $('#carreta-interface').addClass('d-none');
    $('#manual-interface').addClass('d-none');
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

