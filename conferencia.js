/* global $, XLSX */

const { jsPDF } = window.jspdf || {};

const STORAGE_KEY = 'conferencia.routes.v1';

const ConferenciaApp = {
  routes: new Map(),     // routeId -> routeObject
  currentRouteId: null,
  viaCsv: false,

  // =======================
  // Persistência (LOCAL)
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

        (r.ids || []).forEach(id => route.ids.add(String(id)));
        (r.conferidos || []).forEach(id => route.conferidos.add(String(id)));
        (r.foraDeRota || []).forEach(id => route.foraDeRota.add(String(id)));
        route.faltantes = new Set((r.faltantes || []).map(String));

        route.timestamps = new Map(Object.entries(r.timestamps || {}).map(([k, v]) => [String(k), v]));
        route.duplicados = new Map(Object.entries(r.duplicados || {}).map(([k, v]) => [String(k), v]));

        // reconstrói faltantes se vier vazio
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
  },

  clearAllRoutes() {
    this.routes.clear();
    this.currentRouteId = null;
    localStorage.removeItem(STORAGE_KEY);
    this.renderRoutesSelects();
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
      ids: new Set(),        // importados
      faltantes: new Set(),  // ids ainda não conferidos
      conferidos: new Set(), // bipados corretos
      foraDeRota: new Set(), // bipados fora
      duplicados: new Map(), // id -> count (>=2 significa duplicado)

      totalInicial: 0
    };
  },

  get current() {
    if (!this.currentRouteId) return null;
    return this.routes.get(String(this.currentRouteId)) || null;
  },

  setCurrentRoute(routeId) {
    const id = String(routeId);
    if (!this.routes.has(id)) {
      alert('Rota não encontrada.');
      return;
    }
    this.currentRouteId = id;
    this.renderRoutesSelects();
    this.refreshUIFromCurrent();
    this.saveToStorage();
  },

  // =======================
  // UI
  // =======================
  renderRoutesSelects() {
    const $sel1 = $('#saved-routes');
    const $sel2 = $('#saved-routes-inapp');

    const routesSorted = Array.from(this.routes.values())
      .sort((a, b) => String(a.routeId).localeCompare(String(b.routeId)));

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
  // Fora de rota inteligente (global)
  // =======================
  findCorrectRouteForId(id) {
    const needle = String(id);

    for (const [rid, r] of this.routes.entries()) {
      if (r.ids && r.ids.has(needle)) return String(rid);
    }
    for (const [rid, r] of this.routes.entries()) {
      if (r.faltantes && r.faltantes.has(needle)) return String(rid);
    }
    for (const [rid, r] of this.routes.entries()) {
      if (r.conferidos && r.conferidos.has(needle)) return String(rid);
    }
    return null;
  },

  cleanupIdFromOtherRoutes(id, targetRouteId) {
    const needle = String(id);
    const target = String(targetRouteId);

    for (const [rid, r] of this.routes.entries()) {
      if (String(rid) === target) continue;

      let changed = false;

      if (r.foraDeRota && r.foraDeRota.has(needle)) {
        r.foraDeRota.delete(needle);
        changed = true;
      }

      if (r.duplicados && r.duplicados.has(needle)) {
        r.duplicados.delete(needle);
        changed = true;
      }

      if (changed) {
        const stillRelevant =
          (r.conferidos && r.conferidos.has(needle)) ||
          (r.faltantes && r.faltantes.has(needle)) ||
          (r.ids && r.ids.has(needle)) ||
          (r.foraDeRota && r.foraDeRota.has(needle)) ||
          (r.duplicados && r.duplicados.has(needle));

        if (!stillRelevant && r.timestamps) r.timestamps.delete(needle);
      }
    }
  },

  // =======================
  // Conferência
  // =======================
  conferirId(codigo) {
    const r = this.current;
    if (!r || !codigo) return;

    const id = String(codigo);
    const now = Date.now();

    const correctRouteId = this.findCorrectRouteForId(id);
    const isCorrectHere = correctRouteId && String(correctRouteId) === String(this.currentRouteId);

    // se for o ID correto nesta rota, remove de outras rotas onde estava fora
    if (isCorrectHere) {
      this.cleanupIdFromOtherRoutes(id, this.currentRouteId);
    }

    // duplicata (já conferido ou já fora de rota)
    if (r.conferidos.has(id) || r.foraDeRota.has(id)) {
      const count = r.duplicados.get(id) || 1;
      r.duplicados.set(id, count + 1);
      r.timestamps.set(id, now);

      if (!this.viaCsv) this.playAlertSound();

      $('#barcode-input').val('').focus();
      this.saveToStorage();
      this.atualizarListas();
      return;
    }

    // conferência normal
    if (r.faltantes.has(id)) {
      r.faltantes.delete(id);
      r.conferidos.add(id);
      r.timestamps.set(id, now);

      $('#barcode-input').val('').focus();
      this.saveToStorage();
      this.atualizarListas();
      return;
    }

    // fora de rota
    r.foraDeRota.add(id);
    r.timestamps.set(id, now);
    if (!this.viaCsv) this.playAlertSound();

    $('#barcode-input').val('').focus();
    this.saveToStorage();
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
        const shipmentId = String(match[1]);
        const receiverId = String(match[2] || '');
        if (!receiverId.includes('_')) idsExtraidos.add(shipmentId);
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

    this.saveToStorage();
    this.renderRoutesSelects();
    return imported;
  },

  // =======================
  // EXPORT: rota atual CSV (modo local)
  // =======================
  exportRotaAtualCsv() {
    const r = this.current;
    if (!r) {
      alert('Nenhuma rota selecionada.');
      return;
    }

    const rows = [];
    rows.push(['ID', 'STATUS', 'DUPLICADO_X']);

    const idsOrdenados = Array.from(r.ids).sort((a, b) => String(a).localeCompare(String(b)));
    for (const id of idsOrdenados) {
      const dup = r.duplicados.get(id) || 0;
      const status = r.conferidos.has(id) ? 'CONFERIDO' : 'FALTANTE';
      rows.push([id, status, dup ? `${dup}x` : '']);
    }

    const foraOrdenados = Array.from(r.foraDeRota).sort((a, b) => String(a).localeCompare(String(b)));
    for (const id of foraOrdenados) {
      const dup = r.duplicados.get(id) || 0;
      const status = dup > 0 ? 'DUPLICADO' : 'FORA_DE_ROTA';
      rows.push([id, status, dup ? `${dup}x` : '']);
    }

    const esc = (v) => {
      const s = String(v ?? '');
      if (/[",\n\r;]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
      return s;
    };

    const csv = rows.map(line => line.map(esc).join(',')).join('\r\n');

    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);

    const pad2 = (n) => String(n).padStart(2, '0');
    const now = new Date();
    const stamp = `${now.getFullYear()}-${pad2(now.getMonth()+1)}-${pad2(now.getDate())}_${pad2(now.getHours())}${pad2(now.getMinutes())}`;

    link.download = `rota_${r.routeId}_validacao_${stamp}.csv`;
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
      // ordena por timestamp e depois ID
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
    const stamp = `${now.getFullYear()}-${pad2(now.getMonth()+1)}-${pad2(now.getDate())}_${pad2(now.getHours())}${pad2(now.getMinutes())}`;

    XLSX.writeFile(wb, `bipagens_todas_rotas_${stamp}.xlsx`);
  }
};

// =======================
// Eventos
// =======================
$(document).ready(() => {
  console.log('JS carregou', new Date().toISOString());
  ConferenciaApp.loadFromStorage();
  ConferenciaApp.renderRoutesSelects();
});

$(document).on('click', '#extract-btn', () => {
  const raw = $('#html-input').val();
  if (!raw.trim()) return alert('Cole o HTML antes de importar.');

  const qtd = ConferenciaApp.importRoutesFromHtml(raw);
  if (!qtd) return alert('Nenhuma rota importada. Confira se o HTML está completo.');

  alert(`${qtd} rota(s) importada(s) e salva(s)! Agora selecione e clique em "Carregar rota".`);
});

$(document).on('click', '#load-route', () => {
  const id = $('#saved-routes').val();
  if (!id) return alert('Selecione uma rota salva.');

  ConferenciaApp.setCurrentRoute(id);

  $('#initial-interface').addClass('d-none');
  $('#manual-interface').addClass('d-none');
  $('#conference-interface').removeClass('d-none');

  $('#barcode-input').focus();
});

$(document).on('click', '#delete-route', () => {
  const id = $('#saved-routes').val();
  if (!id) return alert('Selecione uma rota para excluir.');
  ConferenciaApp.deleteRoute(id);
});

$(document).on('click', '#clear-all-routes', () => {
  ConferenciaApp.clearAllRoutes();
  alert('Todas as rotas foram removidas.');
});

$(document).on('click', '#switch-route', () => {
  const id = $('#saved-routes-inapp').val();
  if (!id) return;
  ConferenciaApp.setCurrentRoute(id);
  $('#barcode-input').focus();
});

$(document).on('click', '#manual-btn', () => {
  $('#initial-interface').addClass('d-none');
  $('#manual-interface').removeClass('d-none');
});

$(document).on('click', '#submit-manual', () => {
  try {
    const routeId = ($('#manual-routeid').val() || '').trim();
    if (!routeId) return alert('Informe o RouteId.');

    const cluster = ($('#manual-cluster').val() || '').trim();
    const manualIds = $('#manual-input').val().split(/[\s,;]+/).map(x => x.trim()).filter(Boolean);

    if (!manualIds.length) return alert('Nenhum ID válido inserido.');

    const route = ConferenciaApp.routes.get(String(routeId)) || ConferenciaApp.makeEmptyRoute(routeId);
    route.cluster = cluster || route.cluster;

    for (const id of manualIds) {
      const sid = String(id);
      route.ids.add(sid);
      if (!route.conferidos.has(sid)) route.faltantes.add(sid);
    }

    route.totalInicial = route.ids.size;
    ConferenciaApp.routes.set(String(routeId), route);

    ConferenciaApp.saveToStorage();
    ConferenciaApp.renderRoutesSelects();

    alert(`Rota ${routeId} salva com ${route.totalInicial} ID(s).`);

    $('#manual-interface').addClass('d-none');
    $('#initial-interface').removeClass('d-none');
  } catch (e) {
    console.error(e);
    alert('Erro ao processar IDs manuais.');
  }
});

// Enter no input (robusto)
$(document).on('keydown', '#barcode-input', (e) => {
  if (e.key === 'Enter' || e.which === 13) {
    e.preventDefault();

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

$(document).on('click', '#check-csv', () => {
  const r = ConferenciaApp.current;
  if (!r) return alert('Selecione uma rota antes.');

  const fileInput = document.getElementById('csv-input');
  if (!fileInput || fileInput.files.length === 0) return alert('Selecione um arquivo CSV.');

  ConferenciaApp.viaCsv = true;

  const file = fileInput.files[0];
  const reader = new FileReader();

  reader.onload = e => {
    const csvText = e.target.result;
    const linhas = String(csvText || '').split(/\r?\n/);
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

// Exportações
$(document).on('click', '#export-csv-rota-atual', () => {
  ConferenciaApp.exportRotaAtualCsv();
});

$(document).on('click', '#export-xlsx-todas-rotas', () => {
  ConferenciaApp.exportTodasRotasXlsx();
});

// Finalizar sem download
$(document).on('click', '#finish-btn', () => {
  alert('Finalizar não exporta automaticamente. Use os botões de exportação.');
});

$(document).on('click', '#back-btn', () => location.reload());
