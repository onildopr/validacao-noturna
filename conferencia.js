const { jsPDF } = window.jspdf || {};

const STORAGE_KEY = 'conferencia.routes.v1';

const ConferenciaApp = {
  routes: new Map(),     // routeId -> routeObject
  currentRouteId: null,
  viaCsv: false,

  // Cloud
  cloudEnabled: false,
  supabase: null,
  deviceId: null,
  pollHandle: null,

  // =======================
  // Init
  // =======================
  async init() {
    this.deviceId = this.getOrCreateDeviceId();

    // local (fallback/offline)
    this.loadFromStorage();

    // cloud
    await this.initSupabaseIfConfigured();

    if (this.cloudEnabled) {
      await this.loadRoutesFromCloud();
    }

    this.renderRoutesSelects();
  },

  getOrCreateDeviceId() {
    const key = 'conferencia.device_id.v1';
    let v = localStorage.getItem(key);
    if (!v) {
      v = (crypto?.randomUUID?.() || `${Date.now()}_${Math.random().toString(16).slice(2)}`);
      localStorage.setItem(key, v);
    }
    return v;
  },

  async initSupabaseIfConfigured() {
    try {
      const cfg = window.APP_CONFIG || {};
      const SUPABASE_URL = cfg.SUPABASE_URL || window.SUPABASE_URL || '';
      const SUPABASE_KEY = cfg.SUPABASE_KEY || window.SUPABASE_KEY || '';

      if (!SUPABASE_URL || !SUPABASE_KEY || !window.supabase?.createClient) {
        this.cloudEnabled = false;
        return;
      }

      this.supabase = window.supabase.createClient(SUPABASE_URL, SUPABASE_KEY);
      this.cloudEnabled = true;
    } catch (e) {
      console.warn('Supabase init falhou:', e);
      this.cloudEnabled = false;
    }
  },

  startPollingCurrentRoute(ms = 3000) {
    this.stopPolling();
    this.pollHandle = setInterval(async () => {
      if (!this.cloudEnabled) return;
      if (!this.currentRouteId) return;
      await this.syncCurrentRouteFromCloud();
    }, ms);
  },

  stopPolling() {
    if (this.pollHandle) {
      clearInterval(this.pollHandle);
      this.pollHandle = null;
    }
  },

  // =======================
  // Persistência LOCAL
  // =======================
  loadFromStorage() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return;

      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed !== 'object') return;

      this.routes.clear();

      for (const [routeId, r] of Object.entries(parsed)) {
        const route = this.makeEmptyRoute(routeId);

        route.cluster = r.cluster || '';
        route.destinationFacilityId = r.destinationFacilityId || '';
        route.destinationFacilityName = r.destinationFacilityName || '';
        route.totalInicial = Number(r.totalInicial || 0);

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

  saveToStorage() {
    try {
      const obj = {};
      for (const [routeId, r] of this.routes.entries()) {
        obj[routeId] = {
          routeId: r.routeId,
          cluster: r.cluster,
          destinationFacilityId: r.destinationFacilityId,
          destinationFacilityName: r.destinationFacilityName,
          totalInicial: r.totalInicial,

          ids: Array.from(r.ids),
          faltantes: Array.from(r.faltantes),
          conferidos: Array.from(r.conferidos),
          foraDeRota: Array.from(r.foraDeRota),

          timestamps: Object.fromEntries(r.timestamps),
          duplicados: Object.fromEntries(r.duplicados),
        };
      }
      localStorage.setItem(STORAGE_KEY, JSON.stringify(obj));
    } catch (e) {
      console.warn('Falha ao salvar storage:', e);
    }
  },

  deleteRoute(routeId) {
    if (!routeId) return;
    const id = String(routeId);
    this.routes.delete(id);
    if (this.currentRouteId === id) this.currentRouteId = null;
    this.saveToStorage();
    this.renderRoutesSelects();

    if (this.cloudEnabled) this.deleteRouteFromCloud(id).catch(() => {});
  },

  clearAllRoutes() {
    this.routes.clear();
    this.currentRouteId = null;
    localStorage.removeItem(STORAGE_KEY);
    this.renderRoutesSelects();
    alert('Rotas locais removidas. (Cloud não foi apagado automaticamente)');
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
      ids: new Set(),        // IDs esperados (route_items)
      faltantes: new Set(),
      conferidos: new Set(), // derivado
      foraDeRota: new Set(), // derivado
      duplicados: new Map(), // id -> count

      totalInicial: 0
    };
  },

  get current() {
    if (!this.currentRouteId) return null;
    return this.routes.get(String(this.currentRouteId)) || null;
  },

  async setCurrentRoute(routeId) {
    const id = String(routeId);
    if (!this.routes.has(id)) {
      alert('Rota não encontrada.');
      return;
    }
    this.currentRouteId = id;
    this.renderRoutesSelects();

    if (this.cloudEnabled) {
      await this.syncCurrentRouteFromCloud();
      this.startPollingCurrentRoute(3000);
    } else {
      this.refreshUIFromCurrent();
    }

    this.saveToStorage();
  },

  // =======================
  // UI
  // =======================
  renderRoutesSelects() {
    const $sel1 = $('#saved-routes');
    const $sel2 = $('#saved-routes-inapp');

    const routesSorted = Array.from(this.routes.values())
      .sort((a, b) => a.routeId.localeCompare(b.routeId));

    const makeLabel = (r) => {
      const parts = [];
      parts.push(`ROTA ${r.routeId}`);
      if (r.cluster) parts.push(`CLUSTER ${r.cluster}`);
      if (r.destinationFacilityId) parts.push(`XPT ${r.destinationFacilityId}`);
      return parts.join(' • ');
    };

    $sel1.html(
      ['<option value="">(Nenhuma selecionada)</option>']
        .concat(routesSorted.map(r => `<option value="${r.routeId}">${makeLabel(r)}</option>`))
        .join('')
    );

    $sel2.html(routesSorted.map(r => `<option value="${r.routeId}">${makeLabel(r)}</option>`).join(''));

    if (this.currentRouteId) {
      $sel1.val(this.currentRouteId);
      $sel2.val(this.currentRouteId);
    }
  },

  refreshUIFromCurrent() {
    const r = this.current;
    if (!r) return;

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
      Array.from(r.foraDeRota).map(id => `<li class='list-group-item list-group-item-warning'>${id}</li>`).join('')
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

  playAlertSound() {
    try {
      const audio = new Audio('mixkit-alarm-tone-996-_1_.mp3');
      audio.play().catch(() => {});
    } catch {}
  },

  // =======================
  // Cloud helpers
  // =======================
  async loadRoutesFromCloud() {
    const { data, error } = await this.supabase
      .from('routes')
      .select('route_id, cluster, destination_facility_id, destination_facility_name, updated_at');

    if (error) {
      console.warn('Cloud routes load error:', error);
      return;
    }

    const { data: items, error: errItems } = await this.supabase
      .from('route_items')
      .select('route_id, shipment_id');

    if (errItems) {
      console.warn('Cloud route_items load error:', errItems);
      return;
    }

    this.routes.clear();

    const itemsByRoute = new Map();
    for (const it of (items || [])) {
      const rid = String(it.route_id);
      if (!itemsByRoute.has(rid)) itemsByRoute.set(rid, []);
      itemsByRoute.get(rid).push(String(it.shipment_id));
    }

    for (const r of (data || [])) {
      const rid = String(r.route_id);
      const route = this.makeEmptyRoute(rid);
      route.cluster = r.cluster || '';
      route.destinationFacilityId = r.destination_facility_id || '';
      route.destinationFacilityName = r.destination_facility_name || '';

      const list = itemsByRoute.get(rid) || [];
      for (const id of list) route.ids.add(id);

      route.totalInicial = route.ids.size;
      route.faltantes = new Set(route.ids);

      this.routes.set(rid, route);
    }

    this.saveToStorage();
  },

  async upsertRouteMetaToCloud(route) {
    if (!this.cloudEnabled) return;

    const payload = {
      route_id: String(route.routeId),
      cluster: route.cluster || null,
      destination_facility_id: route.destinationFacilityId || null,
      destination_facility_name: route.destinationFacilityName || null,
      updated_at: new Date().toISOString()
    };

    const { error } = await this.supabase
      .from('routes')
      .upsert(payload, { onConflict: 'route_id' });

    if (error) console.warn('upsert routes error:', error);
  },

  // ✅ CORRIGIDO: UP SERT em vez de insert (evita duplicados)
  async upsertRouteItemsToCloud(routeId, idsArray) {
    if (!this.cloudEnabled) return;
    if (!idsArray?.length) return;

    const rows = idsArray.map(id => ({
      route_id: String(routeId),
      shipment_id: String(id)
    }));

    // upsert com onConflict no PK composto
    const { error } = await this.supabase
      .from('route_items')
      .upsert(rows, { onConflict: 'route_id,shipment_id', ignoreDuplicates: true });

    if (error) console.warn('upsert route_items error:', error);
  },

  async deleteRouteFromCloud(routeId) {
    if (!this.cloudEnabled) return;
    await this.supabase.from('routes').delete().eq('route_id', String(routeId));
  },

  async findCorrectRouteForIdCloud(id) {
    if (!this.cloudEnabled) return null;

    const { data, error } = await this.supabase
      .from('route_items')
      .select('route_id')
      .eq('shipment_id', String(id))
      .limit(1);

    if (error) return null;
    if (!data || !data.length) return null;
    return String(data[0].route_id);
  },

  async deleteScansFromOtherRoutesCloud(id, correctRouteId) {
    if (!this.cloudEnabled) return;

    await this.supabase
      .from('scans')
      .delete()
      .eq('shipment_id', String(id))
      .neq('route_id', String(correctRouteId));
  },

  async insertScanCloud(routeId, id) {
    if (!this.cloudEnabled) return;

    const payload = {
      route_id: String(routeId),
      shipment_id: String(id),
      scanned_at: new Date().toISOString(),
      device_id: this.deviceId
    };

    const { error } = await this.supabase.from('scans').insert(payload);
    if (error) console.warn('insert scan error:', error);
  },

  async syncCurrentRouteFromCloud() {
    const r = this.current;
    if (!r) return;

    const { data: items, error: errItems } = await this.supabase
      .from('route_items')
      .select('shipment_id')
      .eq('route_id', String(r.routeId));

    if (errItems) return;

    r.ids = new Set((items || []).map(x => String(x.shipment_id)));
    r.totalInicial = r.ids.size;

    const { data: scans, error: errScans } = await this.supabase
      .from('scans')
      .select('shipment_id, scanned_at')
      .eq('route_id', String(r.routeId))
      .order('scanned_at', { ascending: true });

    if (errScans) return;

    const countMap = new Map();
    const tsMap = new Map();

    for (const s of (scans || [])) {
      const id = String(s.shipment_id);
      countMap.set(id, (countMap.get(id) || 0) + 1);

      const t = Date.parse(s.scanned_at);
      if (!tsMap.has(id) || t > tsMap.get(id)) tsMap.set(id, t);
    }

    r.timestamps = tsMap;
    r.duplicados = new Map();
    r.conferidos = new Set();
    r.foraDeRota = new Set();

    for (const [id, cnt] of countMap.entries()) {
      if (cnt > 1) r.duplicados.set(id, cnt);

      if (r.ids.has(id)) r.conferidos.add(id);
      else r.foraDeRota.add(id);
    }

    r.faltantes = new Set(r.ids);
    for (const id of r.conferidos) r.faltantes.delete(id);

    this.refreshUIFromCurrent();
    this.saveToStorage();
  },

  // =======================
  // Conferência
  // =======================
  async conferirId(codigo) {
    const r = this.current;
    if (!r || !codigo) return;

    let correctRouteId = null;
    if (this.cloudEnabled) {
      correctRouteId = await this.findCorrectRouteForIdCloud(codigo);
    }

    const isCorrectHere = correctRouteId && String(correctRouteId) === String(this.currentRouteId);

    if (isCorrectHere && this.cloudEnabled) {
      await this.deleteScansFromOtherRoutesCloud(codigo, this.currentRouteId);
    }

    if (this.cloudEnabled) {
      await this.insertScanCloud(this.currentRouteId, codigo);
      await this.syncCurrentRouteFromCloud();
      $('#barcode-input').val('').focus();
      return;
    }

    // (se quiser manter modo offline, pode colar aqui a lógica local antiga)
  },

  // =======================
  // Importação HTML: várias rotas
  // =======================
  async importRoutesFromHtml(rawHtml) {
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
      if (clusterMatch) route.cluster = clusterMatch[1];

      const facMatch = /"destinationFacilityId":"([^"]+)","name":"([^"]+)"/.exec(block);
      if (facMatch) {
        route.destinationFacilityId = facMatch[1];
        route.destinationFacilityName = facMatch[2];
      }

      const regexEnvio = /"id":(4\d{10})[\s\S]*?"receiver_id":"([^"]+)"/g;
      let match;
      const idsExtraidos = new Set();

      while ((match = regexEnvio.exec(block)) !== null) {
        const shipmentId = match[1];
        const receiverId = match[2];
        if (!receiverId.includes('_')) idsExtraidos.add(shipmentId);
      }

      if (!idsExtraidos.size) continue;

      for (const id of idsExtraidos) {
        route.ids.add(id);
        route.faltantes.add(id);
      }

      route.totalInicial = route.ids.size;
      this.routes.set(routeId, route);
      imported++;

      if (this.cloudEnabled) {
        await this.upsertRouteMetaToCloud(route);
        await this.upsertRouteItemsToCloud(routeId, Array.from(idsExtraidos)); // ✅ upsert robusto
      }
    }

    this.saveToStorage();
    this.renderRoutesSelects();

    if (this.cloudEnabled) {
      await this.loadRoutesFromCloud();
      this.renderRoutesSelects();
      this.saveToStorage();
    }

    return imported;
  },

  // =======================
  // EXPORT: rota atual CSV (PADRÃO ANTIGO)
  // =======================
  exportRotaAtualCsv() {
    const r = this.current;
    if (!r) return alert('Nenhuma rota selecionada.');

    const all = [
      ...Array.from(r.conferidos || []),
      ...Array.from(r.foraDeRota || []),
      ...Array.from((r.duplicados && r.duplicados.keys()) ? r.duplicados.keys() : [])
    ];

    const uniq = Array.from(new Set(all));
    if (uniq.length === 0) return alert('Nenhum ID para exportar.');

    const parseDateSafe = (value) => {
      if (!value) return new Date();
      if (value instanceof Date) return value;
      if (typeof value === 'number') return new Date(value);
      const d = new Date(value);
      if (!isNaN(d.getTime())) return d;
      return new Date();
    };

    const pad2 = (n) => String(n).padStart(2, '0');
    const zona = 'Horário Padrão de Brasília';
    const header = 'date,time,time_zone,format,text,notes,favorite,date_utc,time_utc,metadata,duplicates';

    uniq.sort((a, b) => {
      const ta = r.timestamps?.get(a) ? Number(r.timestamps.get(a)) : 0;
      const tb = r.timestamps?.get(b) ? Number(r.timestamps.get(b)) : 0;
      return (ta - tb) || String(a).localeCompare(String(b));
    });

    const linhas = uniq.map((id) => {
      const lidaEm = parseDateSafe(r.timestamps?.get(id));
      const date = `${lidaEm.getFullYear()}-${pad2(lidaEm.getMonth() + 1)}-${pad2(lidaEm.getDate())}`;
      const time = `${pad2(lidaEm.getHours())}:${pad2(lidaEm.getMinutes())}:${pad2(lidaEm.getSeconds())}`;

      const iso = lidaEm.toISOString();
      const dateUtc = iso.slice(0, 10);
      const timeUtc = iso.split('T')[1].split('.')[0];

      const totalReads = r.duplicados?.get(id) || 0;
      const dupCount = totalReads ? Math.max(0, totalReads - 1) : 0;

      return `${date},${time},${zona},Code 128,${id},,0,${dateUtc},${timeUtc},,${dupCount}`;
    });

    const conteudo = [header, ...linhas].join('\r\n');
    const blob = new Blob([conteudo], { type: 'text/csv;charset=utf-8;' });

    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);

    const cluster = (r.cluster || 'semCluster').trim() || 'semCluster';
    const rota = (r.routeId || 'semRota').trim() || 'semRota';
    link.download = `${cluster}_${rota}_padrao.csv`;

    link.click();
  },

  // =======================
  // EXPORT: todas as rotas XLSX (1 coluna por rota)
  // =======================
  exportTodasRotasXlsx() {
    if (typeof XLSX === 'undefined') return alert('Biblioteca XLSX não carregou.');
    if (!this.routes || this.routes.size === 0) return alert('Não há rotas salvas para exportar.');

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

    const pad2 = (n) => String(n).padStart(2, '0');
    const now = new Date();
    const stamp = `${now.getFullYear()}-${pad2(now.getMonth() + 1)}-${pad2(now.getDate())}_${pad2(now.getHours())}${pad2(now.getMinutes())}`;

    XLSX.writeFile(wb, `bipagens_todas_rotas_${stamp}.xlsx`);
  }
};

// =======================
// Eventos
// =======================
$(document).ready(async () => {
  await ConferenciaApp.init();
  ConferenciaApp.renderRoutesSelects();
});

$('#extract-btn').click(async () => {
  const raw = $('#html-input').val();
  if (!raw.trim()) return alert('Cole o HTML antes de importar.');

  const qtd = await ConferenciaApp.importRoutesFromHtml(raw);
  if (!qtd) return alert('Nenhuma rota importada. Confira se o HTML está completo.');

  alert(`${qtd} rota(s) importada(s) e salva(s)! Agora selecione e clique em "Carregar rota".`);
});

$('#load-route').click(async () => {
  const id = $('#saved-routes').val();
  if (!id) return alert('Selecione uma rota salva.');

  await ConferenciaApp.setCurrentRoute(id);

  $('#initial-interface').addClass('d-none');
  $('#manual-interface').addClass('d-none');
  $('#conference-interface').removeClass('d-none');
  $('#barcode-input').focus();
});

$('#switch-route').click(async () => {
  const id = $('#saved-routes-inapp').val();
  if (!id) return;
  await ConferenciaApp.setCurrentRoute(id);
  $('#barcode-input').focus();
});

$('#barcode-input').keypress(async (e) => {
  if (e.which === 13) {
    ConferenciaApp.viaCsv = false;
    const raw = $('#barcode-input').val();
    const id = ConferenciaApp.normalizarCodigo(raw);

    if (!id) {
      $('#barcode-input').val('').focus();
      return;
    }

    await ConferenciaApp.conferirId(id);
  }
});

// Exportações
$(document).on('click', '#export-csv-rota-atual', () => ConferenciaApp.exportRotaAtualCsv());
$(document).on('click', '#export-xlsx-todas-rotas', () => ConferenciaApp.exportTodasRotasXlsx());

// Finalizar sem ação
$(document).off('click', '#finish-btn');
$('#finish-btn').off('click');
$(document).on('click', '#finish-btn', (e) => {
  e.preventDefault();
  e.stopImmediatePropagation();
  return false;
});

// Voltar
$('#back-btn').click(() => location.reload());
