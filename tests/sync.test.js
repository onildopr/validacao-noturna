// Sincronização entre aparelhos pelo Supabase (servidor falso em memória).
const test = require('node:test');
const assert = require('node:assert/strict');
const { createServer, createDevice, addRoute, bipar, snapshot, DAY } = require('./helpers');

async function doisAparelhos() {
  const server = createServer();
  const A = createDevice('A', server);
  const B = createDevice('B', server);
  await A.applyWorkDay(DAY);
  await B.applyWorkDay(DAY);
  addRoute(A, '1', 'J11', ['40000000001', '40000000002']);
  addRoute(A, '2', 'J12', ['40000000003']);
  await A.flushDefsSave();
  await B.pullDefs(DAY);
  return { server, A, B };
}

const sincronizar = async (...devs) => {
  for (const d of devs) await d.flushEventQueue();
  for (const d of devs) await d.periodicSyncTick();
};

test('rotas importadas num aparelho chegam no outro', async () => {
  const { B } = await doisAparelhos();
  assert.equal(B.routes.size, 2);
  assert.equal(B.routes.get('1').ids.size, 2);
});

test('bipagens nos dois aparelhos: os dois terminam iguais', async () => {
  const { server, A, B } = await doisAparelhos();
  bipar(A, '1', '40000000001');
  bipar(B, '1', '40000000002');
  bipar(B, '1', '40000000003');
  bipar(A, '2', '40000000003');
  await sincronizar(A, B);

  assert.equal(server.tables.scan_events.length, 4);
  assert.equal(snapshot(A), snapshot(B));
  assert.equal(A.routes.get('1').faltantes.size, 0);
  assert.ok(!A.routes.get('1').foraDeRota.has('40000000003'));
});

test('sem internet: bipagem fica na fila e sobe uma vez só quando volta', async () => {
  const { server, B } = await doisAparelhos();
  server.offline.add('B');
  bipar(B, '1', '40000000001');
  await B.flushEventQueue();
  assert.equal(B.eventQueue.size, 1);
  assert.equal(B.cloudOffline, true);

  server.offline.delete('B');
  await B.flushEventQueue();
  await B.flushEventQueue();
  assert.equal(B.eventQueue.size, 0);
  assert.equal(B.cloudOffline, false);
  assert.equal(server.tables.scan_events.length, 1);
});

test('recarregar a página sem internet mantém fila e estado', async () => {
  const { server, B } = await doisAparelhos();
  server.offline.add('B');
  bipar(B, '1', '40000000001');
  B.persistEventsNow();

  const B2 = createDevice('B', server, { store: B.__store });
  await B2.applyWorkDay(DAY);
  assert.equal(B2.eventQueue.size, 1);
  assert.ok(B2.routes.get('1').conferidos.has('40000000001'));

  server.offline.delete('B');
  await B2.flushEventQueue();
  assert.equal(server.tables.scan_events.length, 1);
});

test('aparelho novo monta o mesmo estado a partir do banco', async () => {
  const { server, A, B } = await doisAparelhos();
  bipar(A, '1', '40000000001');
  bipar(B, '1', '40000000003');
  await sincronizar(A, B);

  const C = createDevice('C', server);
  await C.applyWorkDay(DAY);
  assert.equal(snapshot(C), snapshot(A));
});

test('checagem periódica sem mudanças quase não gasta banco', async () => {
  const { server, A, B } = await doisAparelhos();
  bipar(A, '1', '40000000001');
  await sincronizar(A, B);

  server.bytesOut = 0;
  const antes = server.calls.length;
  await B.periodicSyncTick();
  await B.periodicSyncTick();
  const novas = server.calls.slice(antes);

  assert.ok(server.bytesOut < 300, `baixou ${server.bytesOut} bytes`);
  assert.ok(!novas.some(c => c.table === 'routes_state' && /data/.test(c.cols)), 'não deveria baixar as rotas');
  assert.ok(!novas.some(c => c.table === 'scan_events' && !c.count), 'não deveria baixar bipagens');
});

test('importações ao mesmo tempo em dois aparelhos: nenhuma se perde', async () => {
  const { A, B } = await doisAparelhos();
  addRoute(A, '10', 'K1', ['40000000010']);
  addRoute(B, '11', 'K2', ['40000000011']);
  await A.flushDefsSave();
  await B.flushDefsSave();
  await A.pullDefs(DAY);
  for (const d of [A, B]) {
    assert.ok(d.routes.has('10') && d.routes.has('11'));
  }
});

test('exclusão de rota chega no outro aparelho', async () => {
  const { A, B } = await doisAparelhos();
  A.deleteRoute('2');
  await A.flushDefsSave();
  await B.pullDefs(DAY);
  assert.ok(!B.routes.has('2'));
});

test('placa vinculada chega no outro aparelho, e a exclusão da bipagem também', async () => {
  const { A, B } = await doisAparelhos();
  const pk = A.ensurePlate({ license_plate: 'ABC1D23', carrier_name: 'Transp' });
  A.carretas.currentPlateKey = pk;
  A.vincularRouteQrNaPlaca('assignment:J11', { raw: 'r', jsonText: '{"assignment":"J11"}' }, pk);
  await A.flushDefsSave();
  await B.pullDefs(DAY);
  assert.equal(B.routes.get('1').plateKey, 'ABC1D23');
  assert.ok(B.carretas.plates.get('ABC1D23').routes.has('assignment:J11'));

  const t = Date.now(); while (Date.now() === t) { /* garante horário posterior */ }
  A.clearBipagemForPlate('ABC1D23');
  await A.flushDefsSave();
  await B.pullDefs(DAY);
  assert.equal(B.routes.get('1').plateKey, '');
  assert.equal(B.carretas.plates.get('ABC1D23').routes.size, 0);
});

test('trocar de dia envia o pendente antes', async () => {
  const { server, A } = await doisAparelhos();
  bipar(A, '1', '40000000001');
  await A.applyWorkDay('2026-10-01');
  assert.equal(server.tables.scan_events.length, 1);
  assert.equal(server.tables.scan_events[0].day, DAY);
});
