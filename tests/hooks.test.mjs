// Integration test for code-gs-heartbeat-hooks.patch. Runs only when
// GUEST_MAP_CODE_GS points at a Code.gs with the patch applied (the served
// Code.gs is not committed here: it embeds a credential).
import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import { createGas } from './gas-harness.mjs';

const CODE = process.env.GUEST_MAP_CODE_GS;
const HB = new URL('../Heartbeat.gs', import.meta.url).pathname;
const skip = CODE ? false : 'set GUEST_MAP_CODE_GS to a patched Code.gs';

function load() {
  const g = createGas({ files: [CODE, HB] });
  const set = (name, fn) => { g.context[`__${name}`] = fn; vm.runInContext(`${name} = __${name};`, g.context); };
  set('getSheet', () => ({ getDataRange: () => ({ getValues: () => [['Listing', 'ID', 'Name']] }), getLastRow: () => 1 }));
  set('getExistingEmailIds', () => new Set());
  set('ensureNewColumns_', () => ({}));
  set('getExistingPayoutIds', () => new Set());
  set('buildPayoutSummary', () => null);
  return { ...g, set, beat: (c) => JSON.parse(g.store.get(`HEARTBEAT_${c}`)) };
}

test('quiet successful booking sync records success with zero new rows', { skip }, () => {
  const g = load();
  g.set('searchAll_', () => []);
  assert.equal(g.context.syncAirbnbEmails(), 0);
  assert.equal(g.beat('bookings').lastErrorCode, null);
  assert.ok(g.beat('bookings').lastSuccessAt);
});

test('Gmail search failure: no throw, error recorded, success not advanced', { skip }, () => {
  const g = load();
  g.set('searchAll_', () => { throw new Error('Gmail down: contains guest text'); });
  assert.doesNotThrow(() => g.context.syncAirbnbEmails());
  assert.equal(g.beat('bookings').lastErrorCode, 'gmail-search-failed');
  assert.equal(g.beat('bookings').lastSuccessAt, null);
  assert.doesNotMatch(g.store.get('HEARTBEAT_bookings'), /guest text/);
});

test('payout sync: missing sheet is a recorded failure; quiet success records zero rows', { skip }, () => {
  const g = load();
  g.set('getPayoutSheet', () => null);
  g.context.syncPayoutEmails();
  assert.equal(g.beat('payouts').lastErrorCode, 'payout-sheet-missing');
  g.set('getPayoutSheet', () => ({}));
  g.set('searchAll_', () => []);
  assert.equal(g.context.syncPayoutEmails(), 0);
  assert.equal(g.beat('payouts').lastErrorCode, null);
});

test('doGet exposes heartbeats only in the keyed private view', { skip }, () => {
  const g = load();
  g.set('searchAll_', () => []);
  g.context.syncAirbnbEmails();
  const key = vm.runInContext('PRIVATE_KEY', g.context);
  const priv = JSON.parse(g.context.doGet({ parameter: { key } }).text);
  assert.equal(priv.heartbeats.bookings.producer, 'gas:guest-map:syncAirbnbEmails');
  assert.equal(priv.heartbeats.payouts, null);
  const pub = JSON.parse(g.context.doGet({ parameter: {} }).text);
  assert.equal(Object.hasOwn(pub, 'heartbeats'), false);
  assert.deepEqual(Object.keys(pub), ['guests', 'payouts']);
});
