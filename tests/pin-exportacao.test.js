// PIN (único, conferido no banco) e formato das exportações.
const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const { createServer, createDevice, addRoute, bipar, DAY } = require('./helpers');

// Parser de uma linha CSV com campos entre aspas ("" = aspas dentro do campo)
function parseCsvLine(line) {
  const out = [];
  let cur = '', q = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (q) {
      if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++; }
      else if (ch === '"') q = false;
      else cur += ch;
    } else if (ch === '"') q = true;
    else if (ch === ',') { out.push(cur); cur = ''; }
    else cur += ch;
  }
  out.push(cur);
  return out;
}

// PIN único, conferido no banco (funções pin_enabled / check_pin simuladas em helpers.js)
const PIN_HASH_225487 = crypto.createHash('sha256').update('conferencia:225487').digest('hex');

test('hash do PIN usado no SQL confere com o PIN 225487', () => {
  assert.equal(PIN_HASH_225487, 'a88ccd081e59aa21d14dc62edf90e7c6ac41d2d78c7b18a31f31183dbaab30a6');
});

test('sem PIN cadastrado: ações liberadas sem perguntar', async () => {
  const server = createServer();
  const a = createDevice('T', server);
  assert.equal(await a.pinEnabled(), false);
  assert.equal(await a.requirePin('teste'), true);
});

test('com PIN: o banco aceita só o PIN certo (o mesmo para qualquer operação)', async () => {
  const server = createServer();
  server.pinHash = PIN_HASH_225487;
  const erd1 = createDevice('A', server, { op: 'ERD1' });
  const erd2 = createDevice('B', server, { op: 'ERD2' });
  assert.equal(await erd1.pinEnabled(), true);
  assert.equal(await erd1.verifyPin('225487'), true);
  assert.equal(await erd1.verifyPin(' 225487 '), true);
  assert.equal(await erd1.verifyPin('000000'), false);
  assert.equal(await erd1.verifyPin(''), false);
  assert.equal(await erd2.verifyPin('225487'), true);
});

test('sem conexão com PIN cadastrado: ação destrutiva fica bloqueada', async () => {
  const server = createServer();
  server.pinHash = PIN_HASH_225487;
  const a = createDevice('T', server);
  server.offline.add('T');
  assert.equal(await a.requirePin('teste'), false);
});

test('PIN digitado certo vale por 10 minutos', async () => {
  const server = createServer();
  server.pinHash = PIN_HASH_225487;
  const a = createDevice('T', server);
  a.pinOkUntil = Date.now() + 60 * 1000;
  assert.equal(await a.requirePin('teste'), true);
});

test('cadastro de operação não envia PIN e valida o código', async () => {
  const a = createDevice('T', null);
  let enviado = null;
  a.getSb = () => ({ from: () => ({ upsert: async (op) => { enviado = op; return { error: null }; } }) });
  await assert.rejects(() => a.adminUpsertOperation('XX', '', true), /Código inválido/);
  await a.adminUpsertOperation('erd2', 'Expedição 2', true);
  assert.deepEqual(JSON.parse(JSON.stringify(enviado)), { code: 'ERD2', name: 'Expedição 2', active: true });
});

test('linha do CSV no formato do app de leitura', () => {
  const a = createDevice('T', null);
  const linha = a.buildScannerCsvRow(new Date('2026-09-30T12:34:56Z'), 'QR Code', '{"id":"1","t":"lm"}');
  const cols = parseCsvLine(linha);
  assert.equal(cols.length, 10);
  assert.equal(cols[3], 'QR Code');
  assert.equal(cols[4], '{"id":"1","t":"lm"}');
  assert.equal(cols[7], '2026-09-30'); // data UTC
  assert.equal(cols[8], '12:34:56');   // hora UTC
});

test('CSV com placa/rota: placa, QR da rota e depois os pacotes em ordem de horário', () => {
  const a = createDevice('T', null);
  a.workDay = DAY;
  a.markDefsDirty = () => {};
  a.scheduleEventFlush = () => {};
  addRoute(a, '1', 'J11', ['40000000001', '40000000002']);
  const pk = a.ensurePlate({ license_plate: 'ABC1D23', jsonText: '{"license_plate":"ABC1D23"}' });
  a.carretas.currentPlateKey = pk;
  a.vincularRouteQrNaPlaca('assignment:J11', { raw: 'r', jsonText: '{"assignment":"J11"}' }, pk);
  bipar(a, '1', '40000000002');
  bipar(a, '1', '40000000001');

  const linhas = a.buildScannerCsvLinesForRoute(a.routes.get('1'));
  assert.equal(linhas.length, 4);
  assert.match(linhas[0], /license_plate/);
  assert.match(linhas[1], /assignment/);
  assert.match(linhas[2], /40000000002/);
  assert.match(linhas[3], /40000000001/);
});
