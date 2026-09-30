// Núcleo: constantes, estado do app e utilitários (datas, IDs, Supabase, HTML).

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
  opPins: new Map(),           // código da operação -> hash do PIN (vazio = sem PIN)
  pinOkFor: null,              // operação cujo PIN foi digitado
  pinOkUntil: 0,               // PIN digitado vale por alguns minutos
  cloudOffline: false,         // true após falha de envio ao banco (mostra aviso de pendências)
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

  // ===== Carretas (placa -> rotas QR) =====
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

  // SHA-256 em JS puro (crypto.subtle não existe fora de HTTPS; o app também roda como arquivo local)
  sha256Hex(msg) {
    const bytes = Array.from(new TextEncoder().encode(String(msg)));
    const K = [
      0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
      0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
      0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
      0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
      0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
      0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
      0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
      0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2
    ];
    const H = [0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19];
    const bitLen = bytes.length * 8;
    bytes.push(0x80);
    while (bytes.length % 64 !== 56) bytes.push(0);
    for (let i = 7; i >= 0; i--) bytes.push(i >= 4 ? 0 : (bitLen >>> (i * 8)) & 0xff);

    const rotr = (x, n) => (x >>> n) | (x << (32 - n));
    const w = new Array(64);
    for (let off = 0; off < bytes.length; off += 64) {
      for (let i = 0; i < 16; i++) {
        w[i] = (bytes[off + i * 4] << 24) | (bytes[off + i * 4 + 1] << 16) | (bytes[off + i * 4 + 2] << 8) | bytes[off + i * 4 + 3];
      }
      for (let i = 16; i < 64; i++) {
        const s0 = rotr(w[i - 15], 7) ^ rotr(w[i - 15], 18) ^ (w[i - 15] >>> 3);
        const s1 = rotr(w[i - 2], 17) ^ rotr(w[i - 2], 19) ^ (w[i - 2] >>> 10);
        w[i] = (w[i - 16] + s0 + w[i - 7] + s1) | 0;
      }
      let [a, b, c, d, e, f, g, h] = H;
      for (let i = 0; i < 64; i++) {
        const t1 = (h + (rotr(e, 6) ^ rotr(e, 11) ^ rotr(e, 25)) + ((e & f) ^ (~e & g)) + K[i] + w[i]) | 0;
        const t2 = ((rotr(a, 2) ^ rotr(a, 13) ^ rotr(a, 22)) + ((a & b) ^ (a & c) ^ (b & c))) | 0;
        h = g; g = f; f = e; e = (d + t1) | 0; d = c; c = b; b = a; a = (t1 + t2) | 0;
      }
      H[0] = (H[0] + a) | 0; H[1] = (H[1] + b) | 0; H[2] = (H[2] + c) | 0; H[3] = (H[3] + d) | 0;
      H[4] = (H[4] + e) | 0; H[5] = (H[5] + f) | 0; H[6] = (H[6] + g) | 0; H[7] = (H[7] + h) | 0;
    }
    return H.map(x => (x >>> 0).toString(16).padStart(8, '0')).join('');
  },

  // Hash do PIN de uma operação (o PIN nunca é salvo em texto)
  pinHash(opCode, pin) {
    return this.sha256Hex(`conferencia:${String(opCode || '').toUpperCase()}:${String(pin || '').trim()}`);
  },

  // PIN confere? Operação sem PIN cadastrado => liberado
  checkPin(opCode, pin) {
    const h = this.opPins.get(String(opCode || '').toUpperCase());
    if (!h) return true;
    return this.pinHash(opCode, pin) === h;
  },

  lockRouteUi(ms = 2500) {
    this.routeUiLockUntil = Date.now() + ms;
  },

  isRouteUiLocked() {
    return this.isRouteDropdownOpen || Date.now() < (this.routeUiLockUntil || 0);
  },

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

  storageKeyForDay(dayISO) {
    const op = this.getOperationCode() || 'NOOP';
    return `${STORAGE_KEY_PREFIX}.${op}.${dayISO}`;
  },

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

  newId() {
    if (crypto.randomUUID) return crypto.randomUUID();
    return `${Date.now().toString(16)}-${Math.random().toString(16).slice(2)}-${Math.random().toString(16).slice(2)}`;
  },

  dupCount(v) {
    if (Array.isArray(v)) return v.reduce((m, x) => Math.max(m, Number(x) || 0), 0);
    return Number(v) || 0;
  },

  setStatus(txt, kind = 'muted') {
    const $s = $('#sync-status');
    $s.removeClass('text-muted text-success text-danger text-warning text-info');
    $s.addClass(`text-${kind}`);
    $s.text(txt);
  },

  getRoutesMap() {
    if (this.routes instanceof Map) return this.routes;
    const m = new Map();
    const obj = this.routes || {};
    Object.keys(obj).forEach(k => m.set(String(k), obj[k]));
    this.routes = m;
    return m;
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

  escapeHtml(s) {
    return this.escHtml(s);
  },

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
};
