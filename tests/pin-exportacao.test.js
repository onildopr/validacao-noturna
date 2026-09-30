// PIN por operação e formato das exportações.
const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const { createDevice, addRoute, bipar, DAY } = require('./helpers');

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

test('SHA-256 próprio bate com o do Node', () => {
  const a = createDevice('T', null);
  for (const msg of ['', 'abc', 'conferencia:ERD1:1234', 'x'.repeat(55), 'x'.repeat(64), 'ção 🚚']) {
    assert.equal(a.sha256Hex(msg), crypto.createHash('sha256').update(msg).digest('hex'));
  }
});

test('PIN: operação sem PIN libera; com PIN só o certo passa', () => {
  const a = createDevice('T', null);
  assert.equal(a.checkPin('ERD1', ''), true);

  a.opPins.set('ERD1', a.pinHash('ERD1', '4321'));
  assert.equal(a.checkPin('ERD1', '4321'), true);
  assert.equal(a.checkPin('ERD1', '1234'), false);
  assert.equal(a.checkPin('ERD1', ''), false);
  // o mesmo PIN em outra operação gera outro hash
  assert.notEqual(a.pinHash('ERD1', '4321'), a.pinHash('ERD2', '4321'));
});

test('requirePin libera direto quando a operação não tem PIN', async () => {
  const a = createDevice('T', null);
  assert.equal(await a.requirePin('teste', 'ERD1'), true);
});

test('adminUpsertOperation valida o PIN', async () => {
  const a = createDevice('T', null);
  a.getSb = () => ({ from: () => ({ upsert: async () => ({ error: null }) }) });
  await assert.rejects(() => a.adminUpsertOperation('ERD1', '', true, 'definir', '12'), /4 a 8 números/);
  await a.adminUpsertOperation('ERD1', '', true, 'definir', '2468');
  assert.equal(a.checkPin('ERD1', '2468'), true);
  await a.adminUpsertOperation('ERD1', '', true, 'remover');
  assert.equal(a.checkPin('ERD1', 'qualquer'), true);
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
