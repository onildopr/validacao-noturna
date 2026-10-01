// Senha do dia: dia − mês + ano (primeiro acesso do dia em cada aparelho).
const test = require('node:test');
const assert = require('node:assert/strict');
const { createDevice } = require('./helpers');

test('senha do dia é dia − mês + ano', () => {
  const a = createDevice('T', null);
  assert.equal(a.dailyPassword('2026-09-30'), '2047'); // 30 − 9 + 2026
  assert.equal(a.dailyPassword('2026-10-01'), '2017'); // 1 − 10 + 2026
  assert.equal(a.dailyPassword('2027-01-31'), '2057'); // 31 − 1 + 2027
});

test('aceita o resultado ou a conta escrita, com ou sem zero à esquerda', () => {
  const a = createDevice('T', null);
  for (const ok of ['2047', ' 2047 ', '30-9+2026', '30-09+2026', '30 - 09 + 2026']) {
    assert.equal(a.checkDailyPassword(ok, '2026-09-30'), true, ok);
  }
});

test('recusa senha de outro dia, conta de outro dia e vazio', () => {
  const a = createDevice('T', null);
  for (const bad of ['2048', '2017', '29-9+2026', '30-9+2025', '30092026', '', 'abc']) {
    assert.equal(a.checkDailyPassword(bad, '2026-09-30'), false, bad);
  }
});

test('liberado só no dia em que a senha foi digitada', () => {
  const store = {};
  const a = createDevice('T', null, { store });
  assert.equal(a.isUnlockedToday(), false);
  a.markUnlockedToday();
  assert.equal(a.isUnlockedToday(), true);

  // No dia seguinte o registro antigo não vale
  store['conf_unlock_day.v1'] = '2000-01-01';
  assert.equal(a.isUnlockedToday(), false);
});
