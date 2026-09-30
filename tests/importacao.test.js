// Importação de rotas pelo HTML e leitura de QR (placa / rota).
const test = require('node:test');
const assert = require('node:assert/strict');
const { createDevice, plain, DAY } = require('./helpers');

function app() {
  const a = createDevice('T', null);
  a.workDay = DAY;
  a.markDefsDirty = () => {};
  return a;
}

const HTML = `
  <div>{"routeId":101,"cluster":"J11_AM5","destinationFacilityId":"XPT1","name":"Destino Um",
    "shipments":[{"id": 40000000001},{"id": 40000000002},{"id": 30000000009}]}</div>
  <div>{"routeId":102,"cluster":"J12_AM5","shipments":[{"id": 40000000003}]}</div>
  <div>{"routeId":103,"cluster":"J13_AM5","shipments":[]}</div>
  <div>{"routeId":104,"shipments":[{"id": 40000000004},{"id": 40000000001}]}</div>
`;

test('extrai rotas, cluster, destino e só IDs de pacote válidos', () => {
  const rotas = app().parseRoutesFromHtml(HTML);
  const r101 = rotas.find(r => r.routeId === '101');
  assert.equal(rotas.length, 4);
  assert.equal(r101.cluster, 'J11_AM5');
  assert.equal(r101.destinationFacilityId, 'XPT1');
  assert.equal(r101.destinationFacilityName, 'Destino Um');
  assert.deepEqual(plain([...r101.ids].sort()), ['40000000001', '40000000002']);
});

test('relatório avisa rota sem IDs, rota sem cluster e pacote em duas rotas', () => {
  const a = app();
  const rep = a.importRoutesFromHtml(HTML);
  assert.equal(rep.encontradas, 4);
  assert.equal(rep.importadas, 3);
  assert.deepEqual(plain(rep.semIds), ['103']);
  assert.deepEqual(plain(rep.semCluster), ['104']);
  assert.equal(rep.idsEmOutraRota.length, 1);
  assert.equal(rep.idsEmOutraRota[0].id, '40000000001');

  const txt = a.formatImportReport(rep);
  assert.match(txt, /IGNORADA/);
  assert.match(txt, /sem CLUSTER/);
  assert.match(txt, /mais de uma rota/);
});

test('HTML sem routeId: mensagem clara', () => {
  const a = app();
  const rep = a.importRoutesFromHtml('<html>nada aqui</html>');
  assert.equal(rep.importadas, 0);
  assert.match(a.formatImportReport(rep), /Não encontrei nenhum "routeId"/);
});

test('reimportar rota existente conta os pacotes novos', () => {
  const a = app();
  a.importRoutesFromHtml('"routeId":1 "cluster":"A" "id": 40000000001');
  const rep = a.importRoutesFromHtml('"routeId":1 "cluster":"A" "id": 40000000001 "id": 40000000002');
  assert.deepEqual(plain(rep.atualizadas), [{ routeId: '1', idsNovos: 1 }]);
  assert.equal(a.routes.get('1').ids.size, 2);
});

test('bloco vazio NÃO desfaz a exclusão de uma rota', () => {
  const a = app();
  a.importRoutesFromHtml('"routeId":5 "cluster":"A" "id": 40000000005');
  a.deleteRoute('5');
  a.importRoutesFromHtml('"routeId":5 "cluster":"A"');
  assert.ok(a.deletedRoutes.has('5'));
  assert.ok(!a.revivedRoutes.has('5'));
  assert.ok(!a.routes.has('5'));
});

test('QR de placa no formato ^Ç^', () => {
  const p = app().parseScanPayload('^license_plate^Ç^abc1d23^,^carrier_name^Ç^Transp X^');
  assert.equal(p.kind, 'plate');
  assert.equal(p.plateKey, 'ABC1D23');
  assert.equal(p.plate.carrier_name, 'Transp X');
});

test('QR de rota (assignment) no formato ^Ç^ e em JSON', () => {
  const a = app();
  const caret = a.parseScanPayload('^assignment^Ç^j11_am5^,^container_id^Ç^123^');
  assert.equal(caret.kind, 'routeqr');
  assert.equal(caret.routeKey, 'assignment:J11_AM5');

  const json = a.parseScanPayload('{"container_id":123,"assignment":"J11_AM5"}');
  assert.equal(json.kind, 'routeqr');
  assert.equal(json.routeKey, 'container:123');
});

test('placa digitada (Mercosul e antiga) e pacote', () => {
  const a = app();
  assert.equal(a.parseScanPayload('ABC-1D23').kind, 'plate');
  assert.equal(a.parseScanPayload('ABC1234').kind, 'plate');
  assert.equal(a.parseScanPayload('40000000001').kind, 'shipment');
});

test('lista de IDs da busca aceita separadores variados e IDs curtos', () => {
  const ids = app().parseIdsList('276796395; 276796396,40000000001\n40000000001  12');
  assert.deepEqual(plain(ids), ['276796395', '276796396', '40000000001']);
});
