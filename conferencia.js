const { jsPDF } = window.jspdf || {};

const STORAGE_KEY_PREFIX = 'conferencia.routes.daycache.v1'; // cache local por dia
const SYNC_INTERVAL_MS = 20000;

const ConferenciaApp = {
  routes: new Map(),     // routeId -> routeObject (somente do dia selecionado)
  currentRouteId: null,
  viaCsv: false,

  // Sync mensal/dia
  workDay: null,               // YYYY-MM-DD
  monthFileHandle: null,       // FileSystemFileHandle
  monthData: null,             // JSON do mês carregado
  lastCloudReadAt: 0,
  lastCloudWriteAt: 0,
  syncing: false,
  pendingWrite: false,

  // =======================
  // Util data/strings
  // =======================
  pad2(n) { return String(n).padStart(2, '0'); },

  todayLocalISO() {
    const d = new Date();
    return `${d.getFullYear()}-${this.pad2(d.getMonth()+1)}-${this.pad2(d.getDate())}`;
  },

  monthKeyFromDay(dayISO) {
    // "2026-01"
    return String(dayISO || '').slice(0, 7);
  },

  storageKeyForDay(dayISO) {
    return `${STORAGE_KEY_PREFIX}.${dayISO}`;
  },

  setStatus(txt, kind='muted') {
    const $s = $('#sync-status');
    $s.removeClass('text-muted text-success text-danger text-warning');
    $s.addClass(`text-${kind}`);
    $s.text(txt);
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

      totalInicial: 0
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

  // =======================
  // Cloud file (Drive) - JSON mensal
  // =======================
  async pickMonthFileHandle() {
    if (!window.showOpenFilePicker) {
      alert('Seu navegador não suporta File System Access API. Use Chrome atualizado.');
      return;
    }
    try {
      const [handle] = await window.showOpenFilePicker({
        types: [{
          description: 'Arquivo JSON de conferência do mês',
          accept: { 'application/json': ['.json'] }
        }],
        multiple: false
      });

      this.monthFileHandle = handle;
      $('#month-file-name').text(handle?.name || '(arquivo selecionado)');
      this.setStatus('arquivo selecionado', 'success');

      // Carrega o mês e aplica o dia atual
      await this.loadMonthFile();
      await this.applyWorkDay(this.workDay || this.todayLocalISO());
    } catch (e) {
      console.warn(e);
      this.setStatus('seleção cancelada', 'muted');
    }
  },

  async loadMonthFile() {
    if (!this.monthFileHandle) return;

    this.setStatus('lendo arquivo...', 'warning');
    try {
      const file = await this.monthFileHandle.getFile();
      const text = await file.text();

      let json;
      if (!text.trim()) {
        // arquivo vazio => inicializa
        json = null;
      } else {
        json = JSON.parse(text);
      }

      if (!json || typeof json !== 'object') {
        json = { version: 1, month: this.monthKeyFromDay(this.workDay || this.todayLocalISO()), days: {} };
      }

      if (!json.days || typeof json.days !== 'object') json.days = {};
      if (!json.version) json.version = 1;

      this.monthData = json;
      this.lastCloudReadAt = Date.now();
      this.setStatus('arquivo carregado', 'success');
    } catch (e) {
      console.error('Erro ao ler JSON mensal:', e);
      this.setStatus('erro ao ler arquivo', 'danger');
    }
  },

  async writeMonthFile() {
    if (!this.monthFileHandle || !this.monthData) return;

    // evita escrita concorrente
    if (this.syncing) {
      this.pendingWrite = true;
      return;
    }

    this.syncing = true;
    this.setStatus('salvando...', 'warning');

    try {
      const writable = await this.monthFileHandle.createWritable();
      const out = JSON.stringify(this.monthData, null, 2);
      await writable.write(out);
      await writable.close();
      this.lastCloudWriteAt = Date.now();
      this.setStatus('sincronizado', 'success');
    } catch (e) {
      console.error('Erro ao salvar JSON mensal:', e);
      this.setStatus('erro ao salvar', 'danger');
    } finally {
      this.syncing = false;
      if (this.pendingWrite) {
        this.pendingWrite = false;
        // roda mais uma vez
        setTimeout(() => this.writeMonthFile(), 50);
      }
    }
  },

  ensureMonthMetaForDay(dayISO) {
    if (!this.monthData) {
      this.monthData = { version: 1, month: this.monthKeyFromDay(dayISO), days: {} };
    }
    const mk = this.monthKeyFromDay(dayISO);
    this.monthData.month = this.monthData.month || mk;
    if (!this.monthData.days) this.monthData.days = {};
    if (!this.monthData.days[dayISO]) {
      this.monthData.days[dayISO] = { routes: {} };
    }
    if (!this.monthData.days[dayISO].routes) this.monthData.days[dayISO].routes = {};
  },

  // Carrega rotas do dia a partir do JSON mensal
  loadDayFromMonthData(dayISO) {
    this.routes.clear();
    this.currentRouteId = null;

    if (!this.monthData || !this.monthData.days || !this.monthData.days[dayISO]) return;

    const routesObj = this.monthData.days[dayISO].routes || {};
    for (const [routeId, r] of Object.entries(routesObj)) {
      const route = this.deserializeRoute(routeId, r);
      this.routes.set(String(routeId), route);
    }
  },

  // Grava rotas do dia no JSON mensal
  writeDayToMonthData(dayISO) {
    this.ensureMonthMetaForDay(dayISO);
    const routesObj = {};
    for (const [routeId, r] of this.routes.entries()) {
      routesObj[String(routeId)] = this.serializeRoute(r);
    }
    this.monthData.days[dayISO].routes = routesObj;
  },

  // Merge: cloud(day) + local(day) => local(day) atualizado
  mergeCloudDayIntoLocal(dayISO) {
    if (!this.monthData?.days?.[dayISO]?.routes) return;

    const cloudRoutes = this.monthData.days[dayISO].routes;
    const allRouteIds = new Set([
      ...Object.keys(cloudRoutes),
      ...Array.from(this.routes.keys())
    ]);

    for (const rid of allRouteIds) {
      const local = this.routes.get(String(rid));
      const cloud = cloudRoutes[String(rid)];

      if (!local && cloud) {
        this.routes.set(String(rid), this.deserializeRoute(rid, cloud));
        continue;
      }
      if (local && !cloud) continue;
      if (!local || !cloud) continue;

      // Merge campos básicos
      local.cluster = local.cluster || cloud.cluster || '';
      local.destinationFacilityId = local.destinationFacilityId || cloud.destinationFacilityId || '';
      local.destinationFacilityName = local.destinationFacilityName || cloud.destinationFacilityName || '';
      local.totalInicial = Math.max(Number(local.totalInicial || 0), Number(cloud.totalInicial || 0));

      // Merge Sets
      const cloudIds = new Set(cloud.ids || []);
      const cloudFalt = new Set(cloud.faltantes || []);
      const cloudConf = new Set(cloud.conferidos || []);
      const cloudFora = new Set(cloud.foraDeRota || []);

      for (const id of cloudIds) local.ids.add(id);
      for (const id of cloudFalt) local.faltantes.add(id);
      for (const id of cloudConf) local.conferidos.add(id);
      for (const id of cloudFora) local.foraDeRota.add(id);

      // Merge Maps (timestamps = maior, duplicados = maior)
      const cloudTs = cloud.timestamps || {};
      for (const [id, ts] of Object.entries(cloudTs)) {
        const cur = local.timestamps.get(id) || 0;
        const val = Number(ts || 0);
        if (val > cur) local.timestamps.set(id, val);
      }

      const cloudDup = cloud.duplicados || {};
      for (const [id, cnt] of Object.entries(cloudDup)) {
        const cur = Number(local.duplicados.get(id) || 0);
        const val = Number(cnt || 0);
        if (val > cur) local.duplicados.set(id, val);
      }

      // Recalcula faltantes coerentes
      for (const id of local.conferidos) {
        local.faltantes.delete(id);
      }
    }

    // Regra global: se um ID está conferido na rota correta, remove "fora de rota" das outras
    this.globalCleanupForaDeRotaForConferidos();
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

  async syncTick() {
    if (!this.monthFileHandle) return;          // sem arquivo => offline
    if (!this.workDay) return;

    // lê o arquivo, faz merge, escreve de volta
    await this.loadMonthFile();
    this.mergeCloudDayIntoLocal(this.workDay);

    // grava local -> cloud(day)
    this.writeDayToMonthData(this.workDay);
    await this.writeMonthFile();

    // salva cache local
    this.saveToStorage(this.workDay);

    // atualiza selects/UI
    this.renderRoutesSelects();
    this.refreshUIFromCurrent();
  },

  startSyncLoop() {
    setInterval(() => {
      // não trava UI: roda e ignora erro
      this.syncTick().catch((e) => console.warn('syncTick error', e));
    }, SYNC_INTERVAL_MS);
  },

  // =======================
  // Troca de dia
  // =======================
  async applyWorkDay(dayISO) {
    this.workDay = dayISO;
    $('#work-day').val(dayISO);

    // carrega do arquivo mensal (se houver)
    if (this.monthFileHandle) {
      await this.loadMonthFile();
      this.ensureMonthMetaForDay(dayISO);
      this.loadDayFromMonthData(dayISO);
      this.saveToStorage(dayISO); // cache
      this.renderRoutesSelects();
      this.refreshUIFromCurrent();
      this.setStatus('dia carregado', 'success');
      return;
    }

    // sem arquivo => apenas cache local
    this.loadFromStorage(dayISO);
    this.renderRoutesSelects();
    this.refreshUIFromCurrent();
    this.setStatus('offline (sem arquivo)', 'muted');
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
  },

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

  deleteRoute(routeId) {
    if (!routeId) return;
    this.routes.delete(String(routeId));
    if (this.currentRouteId === String(routeId)) this.currentRouteId = null;

    this.saveToStorage(this.workDay);
    if (this.monthFileHandle && this.monthData) {
      this.writeDayToMonthData(this.workDay);
      this.writeMonthFile().catch(() => {});
    }

    this.renderRoutesSelects();
    this.refreshUIFromCurrent();
  },

  clearAllRoutes() {
    this.routes.clear();
    this.currentRouteId = null;

    localStorage.removeItem(this.storageKeyForDay(this.workDay));

    if (this.monthFileHandle && this.monthData) {
      this.ensureMonthMetaForDay(this.workDay);
      this.monthData.days[this.workDay].routes = {};
      this.writeMonthFile().catch(() => {});
    }

    this.renderRoutesSelects();
    this.refreshUIFromCurrent();
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
      this.saveToStorage(this.workDay);
      this.atualizarListas();
      return;
    }

    // fora de rota
    r.foraDeRota.add(codigo);
    r.timestamps.set(codigo, now);
    if (!this.viaCsv) this.playAlertSound();

    $('#barcode-input').val('').focus();
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
        if (!route.conferidos.has(id)) route.faltantes.add(id);
      }

      route.totalInicial = route.ids.size;
      this.routes.set(routeId, route);
      imported++;
    }

    this.saveToStorage(this.workDay);

    // grava no arquivo mensal também
    if (this.monthFileHandle && this.monthData) {
      this.writeDayToMonthData(this.workDay);
      this.writeMonthFile().catch(() => {});
    }

    this.renderRoutesSelects();
    return imported;
  },

  // =======================
  // EXPORT: rota atual CSV (FORMATO ANTIGO PADRÃO)
  // =======================
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
  // set day default
  const today = ConferenciaApp.todayLocalISO();
  $('#work-day').val(today);
  await ConferenciaApp.applyWorkDay(today);

  ConferenciaApp.startSyncLoop();
});

// Selecionar arquivo mensal (Drive)
$(document).on('click', '#btn-pick-month-file', () => {
  ConferenciaApp.pickMonthFileHandle();
});

// Troca de dia
$(document).on('change', '#work-day', async (e) => {
  const day = e.target.value;
  if (!day) return;
  await ConferenciaApp.applyWorkDay(day);
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
  ConferenciaApp.clearAllRoutes();
  alert('Todas as rotas do dia foram removidas.');
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

    if (ConferenciaApp.monthFileHandle && ConferenciaApp.monthData) {
      ConferenciaApp.writeDayToMonthData(ConferenciaApp.workDay);
      ConferenciaApp.writeMonthFile().catch(() => {});
    }

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

