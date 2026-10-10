import test from 'node:test';
import assert from 'node:assert/strict';
import { createGas } from './gas-harness.mjs';

const files = [new URL('../Lifecycle.gs', import.meta.url).pathname, new URL('./fixtures/stub-parser.gs', import.meta.url).pathname];
const parse = (body, day = '2026-05-01') => JSON.parse(JSON.stringify(createGas({ files }).context.parseAirbnbEmailResult_(body, 'Bekræftet reservation', new Date(day), '', day)));

test('adjacent-column confirmation gets the fixed dates, nights and code', () => {
  const r = parse('Indtjekning   Udtjekning\ntirs. 7. jul.   søn. 12. jul.\n5 nætter');
  assert.equal(r.ok, true);
  assert.equal(r.value.checkIn, '2026-07-07');
  assert.equal(r.value.checkOut, '2026-07-12');
  assert.equal(r.value.nights, 5);
  assert.equal(r.value.code, 'HMAAAAAAA1');
});

test('Danish dates that cannot be parsed are a rejection with a reason, not a row', () => {
  assert.deepEqual(parse('Indtjekning\n\nUdtjekning\n'), { ok: false, reason: 'dates-missing' });
});

test('stated nights disagreeing with the dates is rejected', () => {
  assert.deepEqual(parse('Indtjekning   Udtjekning\ntirs. 7. jul.   søn. 12. jul.\n3 nætter'), { ok: false, reason: 'nights-mismatch' });
});

test('the legacy parser returning null or throwing becomes a reason', () => {
  assert.deepEqual(parse('BROKEN\nIndtjekning   Udtjekning\ntirs. 7. jul.   søn. 12. jul.\n'), { ok: false, reason: 'parse-failed' });
  assert.deepEqual(parse('THROW\nIndtjekning   Udtjekning\ntirs. 7. jul.   søn. 12. jul.\n'), { ok: false, reason: 'parse-error' });
});

test('non-Danish emails keep the legacy parser dates but are still validated', () => {
  const r = parse('Check-in Jan 1, 1999 ... 1 nights');
  assert.equal(r.ok, true);
  assert.equal(r.value.checkIn, '1999-01-01');
});
