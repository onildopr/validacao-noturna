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
