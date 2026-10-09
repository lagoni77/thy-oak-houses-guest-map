import test from 'node:test';
import assert from 'node:assert/strict';
import { createGas } from './gas-harness.mjs';

const HB = new URL('../Heartbeat.gs', import.meta.url).pathname;
let clock = new Date('2026-10-10T10:00:00.000Z');
const gas = () => createGas({ files: [HB], now: () => clock });
const beat = (g, c) => JSON.parse(g.store.get(`HEARTBEAT_${c}`));

test('a successful run with zero new rows advances lastSuccessAt', () => {
  const g = gas();
  clock = new Date('2026-10-10T10:00:00.000Z');
  const out = g.context.withHeartbeat_('bookings', () => 0);
  assert.equal(out, 0);
  assert.deepEqual(beat(g, 'bookings'), {
    producer: 'gas:guest-map:syncAirbnbEmails', lastAttemptAt: '2026-10-10T10:00:00.000Z',
    lastSuccessAt: '2026-10-10T10:00:00.000Z', lastErrorCode: null, lastNewRows: 0,
  });
});

test('a handled failure records the attempt and code but keeps the earlier success, and does not throw', () => {
  const g = gas();
  clock = new Date('2026-10-10T10:00:00.000Z');
  g.context.withHeartbeat_('payouts', () => 2);
  clock = new Date('2026-10-10T11:00:00.000Z');
  const out = g.context.withHeartbeat_('payouts', () => { throw g.context.heartbeatFailure_('gmail-search-failed'); });
  assert.equal(out, undefined);
  assert.deepEqual(beat(g, 'payouts'), {
    producer: 'gas:guest-map:syncPayoutEmails', lastAttemptAt: '2026-10-10T11:00:00.000Z',
    lastSuccessAt: '2026-10-10T10:00:00.000Z', lastErrorCode: 'gmail-search-failed', lastNewRows: 2,
  });
});

test('an unexpected exception is recorded as exception and rethrown (trigger still reports failure)', () => {
  const g = gas();
  assert.throws(() => g.context.withHeartbeat_('bookings', () => { throw new Error('sheet gone'); }), /sheet gone/);
  const b = beat(g, 'bookings');
  assert.equal(b.lastErrorCode, 'exception');
  assert.equal(b.lastSuccessAt, null);
});

test('success clears the previous error code', () => {
  const g = gas();
  g.context.withHeartbeat_('bookings', () => { throw g.context.heartbeatFailure_('gmail-search-failed'); });
  g.context.withHeartbeat_('bookings', () => 1);
  assert.equal(beat(g, 'bookings').lastErrorCode, null);
});

test('readHeartbeats_ returns both collectors, null when never run, and tolerates corrupt values', () => {
  const g = gas();
  g.context.withHeartbeat_('bookings', () => 3);
  g.store.set('HEARTBEAT_payouts', '{corrupt');
  const out = JSON.parse(JSON.stringify(g.context.readHeartbeats_()));
  assert.equal(out.bookings.lastNewRows, 3);
  assert.equal(out.payouts, null);
});

test('an unknown collector name is refused', () => {
  assert.throws(() => gas().context.withHeartbeat_('pricing', () => 0), /unknown collector/);
});
