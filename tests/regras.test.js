// Regras da conferência: ok / fora de rota / duplicado, e recálculo a partir das bipagens.
const test = require('node:test');
const assert = require('node:assert/strict');
const { createDevice, addRoute, bipar, snapshot, DAY } = require('./helpers');

function novoAparelho() {
  const app = createDevice('T', null);
  app.workDay = DAY;
  app.markDefsDirty = () => {};
  app.scheduleEventFlush = () => {};
  addRoute(app, '1', 'J11', ['40000000001', '40000000002']);
  addRoute(app, '2', 'J12', ['40000000003']);
  return app;
}

test('bipar ID da própria rota: conferido e sai dos faltantes', () => {
  const app = novoAparelho();
  bipar(app, '1', '40000000001');
  const r = app.routes.get('1');
  assert.ok(r.conferidos.has('40000000001'));
  assert.ok(!r.faltantes.has('40000000001'));
  assert.equal(r.faltantes.size, 1);
});

test('bipar ID de outra rota: fora de rota', () => {
  const app = novoAparelho();
  bipar(app, '1', '40000000003');
  assert.ok(app.routes.get('1').foraDeRota.has('40000000003'));
  assert.equal(app.findCorrectRouteForId('40000000003'), '2');
});

test('bipar ID desconhecido: fora de rota', () => {
  const app = novoAparelho();
  bipar(app, '1', '40000000999');
  assert.ok(app.routes.get('1').foraDeRota.has('40000000999'));
});

test('bipar duas vezes: duplicado conta 2, 3...', () => {
  const app = novoAparelho();
  bipar(app, '1', '40000000001');
  bipar(app, '1', '40000000001');
  assert.equal(app.routes.get('1').duplicados.get('40000000001'), 2);
  bipar(app, '1', '40000000001');
  assert.equal(app.routes.get('1').duplicados.get('40000000001'), 3);
});

test('fora de rota some quando o pacote é conferido na rota certa', () => {
  const app = novoAparelho();
  bipar(app, '1', '40000000003');
  bipar(app, '2', '40000000003');
  assert.ok(!app.routes.get('1').foraDeRota.has('40000000003'));
  assert.ok(app.routes.get('2').conferidos.has('40000000003'));
});

test('fora de rota DEPOIS da conferência continua como alerta', () => {
  const app = novoAparelho();
  bipar(app, '2', '40000000003');
  bipar(app, '1', '40000000003');
  assert.ok(app.routes.get('1').foraDeRota.has('40000000003'));
});

test('recalcular do zero dá o mesmo resultado que ir aplicando', () => {
  const app = novoAparelho();
  ['40000000001', '40000000003', '40000000003', '40000000002'].forEach(c => bipar(app, '1', c));
  bipar(app, '2', '40000000003');
  const antes = snapshot(app);
  app.rebuildScanState();
  assert.equal(snapshot(app), antes);
});

test('ordem de chegada não importa: bipagem "atrasada" é encaixada pelo horário', () => {
  const a = novoAparelho();
  const b = novoAparelho();
  const evs = [
    { cid: 'x1', pkg: '40000000003', route: '1', ts: 1000, dev: 'A', sv: 1 },
    { cid: 'x2', pkg: '40000000003', route: '2', ts: 2000, dev: 'B', sv: 1 },
    { cid: 'x3', pkg: '40000000001', route: '1', ts: 3000, dev: 'A', sv: 1 },
  ];
  a.addEvents(evs);
  b.addEvents([evs[2]]);
  b.addEvents([evs[1]]);
  b.addEvents([evs[0]]);
  assert.equal(snapshot(a), snapshot(b));
});

test('rota excluída e reimportada ignora as bipagens antigas', () => {
  const app = novoAparelho();
  bipar(app, '2', '40000000003');
  app.deleteRoute('2');
  const rep = app.importRoutesFromHtml('"routeId":2 "cluster":"J12" "id": 40000000003');
  assert.equal(rep.importadas, 1);
  assert.equal(app.routes.get('2').conferidos.size, 0);
  assert.equal(app.routes.get('2').faltantes.size, 1);
});

test('rota importada depois da bipagem: aponta a rota certa do pacote', () => {
  const app = novoAparelho();
  bipar(app, '1', '40000000077');
  assert.equal(app.findCorrectRouteForId('40000000077'), null);
  addRoute(app, '77', 'K7', ['40000000077']);
  assert.equal(app.findCorrectRouteForId('40000000077'), '77');
  assert.ok(app.routes.get('1').foraDeRota.has('40000000077'));
});

test('normalizarCodigo aceita o ID dentro de texto do leitor', () => {
  const app = novoAparelho();
  assert.equal(app.normalizarCodigo('40000000001'), '40000000001');
  assert.equal(app.normalizarCodigo(' {"id":"40000000001","t":"lm"} '), '40000000001');
  assert.equal(app.normalizarCodigo('abc'), null);
});
