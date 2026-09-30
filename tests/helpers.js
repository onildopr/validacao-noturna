// Utilitários dos testes: carrega os módulos do app num contexto isolado (sem navegador),
// com um jQuery "de mentira" e um Supabase falso em memória compartilhado entre aparelhos.
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const MODULES = ['core', 'regras', 'sync', 'banco', 'importacao', 'exportacao', 'ui', 'eventos'];
const SRC = MODULES
  .map(m => fs.readFileSync(path.join(__dirname, '..', 'js', `${m}.js`), 'utf8'))
  .join('\n;\n') + '\n;this.App = ConferenciaApp;';

// jQuery mínimo: qualquer chamada devolve o próprio objeto; .val() sem argumento devolve ''
function fakeJQuery() {
  const make = () => {
    const o = new Proxy(function () {}, {
      get: (t, k) => {
        if (k === 'length') return 0;
        if (k === 'val') return (...a) => (a.length ? o : '');
        if (k === 'hasClass') return () => true;
        if (k === 'is') return () => false;
        return () => o;
      },
      apply: () => o,
    });
    return o;
  };
  return new Proxy(make, { apply: () => make(), get: () => () => make() });
}

// ===== Supabase falso =====
function createServer() {
  return {
    tables: { operations: [], routes_state: [], scan_events: [] },
    nextId: 1,
    offline: new Set(),   // nomes de aparelhos sem conexão
    bytesOut: 0,          // bytes "baixados" pelos aparelhos (para medir economia)
    calls: [],            // log de consultas: {dev, table, op, cols}
  };
}

function makeSb(server, dev) {
  const pick = (row, cols) => {
    if (!cols || cols === '*') return Object.assign({}, row);
    const out = {};
    for (const c of cols.split(',').map(x => x.trim())) out[c] = row[c] === undefined ? null : row[c];
    return out;
  };
  const clone = (x) => JSON.parse(JSON.stringify(x));

  const from = (table) => {
    const st = { table, op: 'select', cols: '*', filters: [], order: null, limit: null, count: false, head: false, rows: null, opts: {}, single: false, maybe: false, returning: null };
    const b = {
      select(cols, opts) {
        if (st.op === 'upsert') st.returning = cols; else st.cols = cols;
        if (opts && opts.count) { st.count = true; st.head = !!opts.head; }
        return b;
      },
      eq(k, v) { st.filters.push(r => String(r[k]) === String(v)); return b; },
      gt(k, v) { st.filters.push(r => r[k] > v); return b; },
      gte(k, v) { st.filters.push(r => String(r[k]) >= String(v)); return b; },
      lte(k, v) { st.filters.push(r => String(r[k]) <= String(v)); return b; },
      in(k, arr) { const s = new Set(arr.map(String)); st.filters.push(r => s.has(String(r[k]))); return b; },
      order(k, o) { st.order = [k, !o || o.ascending !== false]; return b; },
      limit(n) { st.limit = n; return b; },
      maybeSingle() { st.maybe = true; return b; },
      single() { st.single = true; return b; },
      upsert(rows, opts) { st.op = 'upsert'; st.rows = Array.isArray(rows) ? rows : [rows]; st.opts = opts || {}; return b; },
      insert(rows) { st.op = 'upsert'; st.rows = Array.isArray(rows) ? rows : [rows]; st.opts = {}; return b; },
      then(res, rej) { return Promise.resolve().then(run).then(res, rej); },
    };

    const run = () => {
      server.calls.push({ dev, table, op: st.op, cols: st.cols, count: st.count });
      if (server.offline.has(dev)) return { data: null, error: { message: 'offline (teste)' } };
      const T = server.tables[table];

      if (st.op === 'upsert') {
        const keys = String(st.opts.onConflict || (table === 'operations' ? 'code' : 'id')).split(',');
        const saved = [];
        for (const r of st.rows) {
          const i = T.findIndex(x => keys.every(k => x[k] != null && String(x[k]) === String(r[k])));
          if (i >= 0) {
            if (st.opts.ignoreDuplicates) continue;
            T[i] = Object.assign({}, T[i], clone(r));
            if (table === 'routes_state') T[i].updated_at = new Date(Date.now() + server.nextId++).toISOString();
            saved.push(T[i]);
          } else {
            const row = Object.assign(table === 'scan_events' ? { id: server.nextId++ } : {}, clone(r));
            if (table === 'routes_state') row.updated_at = new Date(Date.now() + server.nextId++).toISOString();
            T.push(row);
            saved.push(row);
          }
        }
        if (!st.returning) return { data: null, error: null };
        const data = saved.map(r => pick(r, st.returning));
        server.bytesOut += JSON.stringify(data).length;
        return { data: st.single ? data[0] || null : data, error: null };
      }

      let rows = T.filter(r => st.filters.every(f => f(r)));
      if (st.count && st.head) {
        server.bytesOut += 20;
        return { count: rows.length, data: null, error: null };
      }
      if (st.order) {
        const [k, asc] = st.order;
        rows = rows.slice().sort((a, b) => (a[k] < b[k] ? -1 : a[k] > b[k] ? 1 : 0) * (asc ? 1 : -1));
      }
      if (st.limit) rows = rows.slice(0, st.limit);
      let data = rows.map(r => clone(pick(r, st.cols)));
      if (st.maybe || st.single) data = data[0] || null;
      server.bytesOut += JSON.stringify(data).length;
      return { data, error: null };
    };
    return b;
  };

  return {
    from,
    channel: () => ({ on() { return this; }, subscribe() { return this; } }),
    removeChannel: async () => {},
    rpc: async () => ({ data: [], error: null }),
  };
}

// Cria um "aparelho": uma instância independente do app ligada ao servidor falso
function createDevice(name, server, { store = {}, op = 'ERD1' } = {}) {
  let uuid = 0;
  const ctx = {
    console: { log() {}, warn() {}, error() {} },
    setTimeout: () => 0, clearTimeout() {}, setInterval: () => 0, clearInterval() {},
    localStorage: { getItem: k => (k in store ? store[k] : null), setItem: (k, v) => { store[k] = String(v); } },
    crypto: { randomUUID: () => `${name}-${++uuid}-${Math.random().toString(16).slice(2, 8)}` },
    TextEncoder, Intl, Date, JSON, Math, Map, Set, Promise, Array, Object, String, Number, Blob: function () {}, URL: {},
    window: { addEventListener() {} },
    document: { addEventListener() {}, hidden: false, createElement: () => ({ click() {} }) },
    $: fakeJQuery(),
    alert() {}, confirm: () => true, prompt: () => null,
  };
  vm.createContext(ctx);
  vm.runInContext(SRC, ctx);
  const app = ctx.App;
  app.operationCode = op;
  app.deviceId = name;
  if (server) {
    const sb = makeSb(server, name);
    app.getSb = () => sb;
  }
  app.__store = store;
  return app;
}

// Monta uma rota direto na memória (atalho para os testes)
function addRoute(app, routeId, cluster, ids) {
  const r = app.makeEmptyRoute(routeId);
  r.cluster = cluster;
  ids.forEach(id => { r.ids.add(id); r.faltantes.add(id); });
  r.totalInicial = ids.length;
  app.routes.set(String(routeId), r);
  app.saveToStorage(app.workDay);
  return r;
}

// Estado comparável entre aparelhos
function snapshot(app) {
  return JSON.stringify(Array.from(app.routes.values())
    .sort((a, b) => (a.routeId < b.routeId ? -1 : 1))
    .map(r => ({
      id: r.routeId,
      conferidos: [...r.conferidos].sort(),
      faltantes: [...r.faltantes].sort(),
      fora: [...r.foraDeRota].sort(),
      dup: [...r.duplicados.entries()].sort(),
    })));
}

// Bipa um código numa rota. Espera o relógio andar 1 ms para cada bipagem ter horário próprio
// (na vida real duas bipagens no mesmo milissegundo em aparelhos diferentes são raríssimas).
function bipar(app, routeId, code) {
  const t = Date.now();
  while (Date.now() === t) { /* espera 1 ms */ }
  app.currentRouteId = String(routeId);
  app.conferirId(code);
}

// Converte objetos criados dentro do contexto isolado em objetos comuns (para deepEqual)
const plain = (x) => JSON.parse(JSON.stringify(x));

const DAY = '2026-09-30';

module.exports = { createServer, createDevice, addRoute, snapshot, bipar, plain, DAY };
