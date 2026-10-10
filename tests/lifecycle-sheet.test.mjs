// Lifecycle.gs sheet/Gmail adapters against in-memory fakes. Synthetic data only.
import test from 'node:test';
import assert from 'node:assert/strict';
import { createGas, createSheet, createMessage } from './gas-harness.mjs';

const LC = new URL('../Lifecycle.gs', import.meta.url).pathname;
const HEAD = ['Listing', 'ID', 'Name', 'CheckIn', 'CheckOut', 'Location', 'Lat', 'Lng', 'EmailID', 'Nights', 'Payout', 'Status', 'Guests', 'BookingUrl'];
const row = (code, inD, outD, status = '') => ['Tiny Oak House', 'TOH-X', 'Synthetic Guest', inD, outD, 'Nowhere', 0, 0, `conf-${code}`, 3, '1.000,00 kr', status, '2', `https://www.airbnb.dk/hosting/reservations/details/${code}`];
const cancelMsg = (id, code, date = '2026-10-09T08:00:00Z') => createMessage({ id, subject: `Annulleret: Reservation ${code} for Tiny Oak House`, body: 'Reservationen er annulleret.', date });
const changeMsg = (id, code, dates, date = '2026-10-09T09:00:00Z') => createMessage({
  id, subject: `Reservation altered: ${code}`, date,
  body: typeof dates === 'function' ? dates : `Indtjekning   Udtjekning\n${dates}\n\nhttps://www.airbnb.dk/hosting/reservations/details/${code}`,
});

function setup({ rows = [row('HMAAAAAAA1', '2026-11-02', '2026-11-05')], messages = [], mode = null, now } = {}) {
  const guestData = createSheet('GuestData', [HEAD, ...rows]);
  const gas = createGas({ files: [LC], sheets: { GuestData: guestData }, messages, ...(now ? { now } : {}) });
  if (mode) gas.store.set('LIFECYCLE_APPLY', mode);
  return { ...gas, guestData, run: () => JSON.parse(JSON.stringify(gas.context.runLifecycleScan_(guestData))) };
}
const ledgerRows = (s) => (s.book.ReservationEvents ? s.book.ReservationEvents.data.length - 1 : 0);
const statusOf = (s, r = 1) => s.guestData.data[r][11];

test('parse issues: upsert counts attempts, resolve closes, only open ones are listed', () => {
  const s = setup();
  const c = s.context;
  c.recordParseIssue_('e1', 'confirmation', 'dates-missing', '2026-10-01');
  c.recordParseIssue_('e1', 'confirmation', 'dates-inverted', '2026-10-01');
  c.recordParseIssue_('e2', 'lifecycle', 'read-failed', '2026-10-02');
  let open = JSON.parse(JSON.stringify(c.openParseIssues_('confirmation')));
  assert.equal(open.length, 1);
  assert.equal(open[0].attempts, 2);
  assert.equal(open[0].reason, 'dates-inverted');
  c.resolveParseIssue_('e1');
  assert.equal(c.openParseIssues_('confirmation').length, 0);
  assert.equal(c.openParseIssues_('lifecycle').length, 1);
  assert.deepEqual(s.book.ParseIssues.data[0], ['EmailID', 'Collector', 'Reason', 'EmailDay', 'FirstSeenAt', 'LastAttemptAt', 'Attempts', 'ResolvedAt']);
});

test('mode off (default): cancellations are ledgered but GuestData is not touched', () => {
  const s = setup({ messages: [cancelMsg('c1', 'HMAAAAAAA1')] });
  const r = s.run();
  assert.equal(r.mode, 'off');
  assert.equal(ledgerRows(s), 1);
  assert.equal(statusOf(s), '');
  assert.equal(r.pending, 1);
});

test('mode cancellations: Status set once; a replayed run writes nothing new', () => {
  const s = setup({ messages: [cancelMsg('c1', 'HMAAAAAAA1')], mode: 'cancellations' });
  assert.equal(s.run().applied, 1);
  assert.equal(statusOf(s), 'Cancelled');
  const writes = s.guestData.writes;
  const again = s.run();
  assert.equal(ledgerRows(s), 1, 'same email never ledgered twice');
  assert.equal(s.guestData.writes, writes, 'no rewrite of an already-cancelled row');
  assert.equal(again.applied, 0);
});

test('cancellation never deletes the row; an already-cancelled row is left as is', () => {
  const s = setup({ rows: [row('HMAAAAAAA1', '2026-11-02', '2026-11-05', 'Aflyst')], messages: [cancelMsg('c1', 'HMAAAAAAA1')], mode: 'cancellations' });
  s.run();
  assert.equal(statusOf(s), 'Aflyst');
  assert.equal(s.guestData.data.length, 2);
});

test('mode cancellations: an (unverified) change is ledgered but not applied', () => {
  const s = setup({ messages: [changeMsg('x1', 'HMAAAAAAA1', 'tirs. 3. nov.   fre. 6. nov.')], mode: 'cancellations' });
  const r = s.run();
  assert.equal(ledgerRows(s), 1);
  assert.equal(s.guestData.data[1][3], '2026-11-02');
  assert.equal(r.pending, 1);
});

test('mode all: an accepted change updates CheckIn, CheckOut and Nights', () => {
  const s = setup({ messages: [changeMsg('x1', 'HMAAAAAAA1', 'tirs. 3. nov.   fre. 6. nov.')], mode: 'all' });
  assert.equal(s.run().applied, 1);
  assert.deepEqual(s.guestData.data[1].slice(3, 5), ['2026-11-03', '2026-11-06']);
  assert.equal(s.guestData.data[1][9], 3);
});

test('mode all: a cancellation after a change wins; processing order does not matter', () => {
  const msgs = [cancelMsg('c1', 'HMAAAAAAA1', '2026-10-09T12:00:00Z'), changeMsg('x1', 'HMAAAAAAA1', 'tirs. 3. nov.   fre. 6. nov.', '2026-10-09T09:00:00Z')];
  for (const order of [msgs, [...msgs].reverse()]) {
    const s = setup({ messages: order, mode: 'all' });
    s.run();
    assert.equal(statusOf(s), 'Cancelled');
  }
});

test('out of order: a cancellation seen before its booking row exists is applied once the row appears', () => {
  const s = setup({ rows: [], messages: [cancelMsg('c1', 'HMBBBBBBB2')], mode: 'cancellations' });
  const first = s.run();
  assert.equal(first.rowNotFound, 1);
  s.guestData.data.push(row('HMBBBBBBB2', '2026-12-01', '2026-12-04'));
  const res = JSON.parse(JSON.stringify(s.context.applyLifecycleFor_(s.guestData, ['HMBBBBBBB2'])));
  assert.equal(res.applied, 1);
  assert.equal(statusOf(s), 'Cancelled');
});

test('unparseable change is an open lifecycle issue that holds the lifecycle bookmark', () => {
  const s = setup({ messages: [changeMsg('x1', 'HMAAAAAAA1', 'ingen datoer', '2026-10-03T09:00:00Z')], mode: 'all' });
  const r = s.run();
  assert.equal(r.issues, 1);
  assert.equal(ledgerRows(s), 0, 'nothing ledgered from a failed parse');
  assert.equal(s.context.openParseIssues_('lifecycle')[0].reason, 'dates-missing');
  assert.equal(s.store.get('LAST_LIFECYCLE_SCAN_DATE'), '2026/10/03');
});

test('failure then recovery: an unreadable message is retried next run and its issue resolved', () => {
  let attempts = 0;
  const flaky = () => { attempts += 1; if (attempts === 1) throw new Error('Gmail hiccup'); return 'Indtjekning   Udtjekning\ntirs. 3. nov.   fre. 6. nov.\n'; };
  const s = setup({ messages: [changeMsg('x1', 'HMAAAAAAA1', flaky, '2026-10-05T09:00:00Z')], mode: 'all' });
  assert.equal(s.run().issues, 1);
  assert.equal(s.context.openParseIssues_('lifecycle')[0].reason, 'read-failed');
  assert.equal(s.store.get('LAST_LIFECYCLE_SCAN_DATE'), '2026/10/05');
  const second = s.run();
  assert.equal(second.issues, 0);
  assert.equal(s.context.openParseIssues_('lifecycle').length, 0);
  assert.equal(s.guestData.data[1][3], '2026-11-03');
  assert.equal(s.store.get('LAST_LIFECYCLE_SCAN_DATE'), '2026/10/10', 'bookmark advances only after recovery');
});

test('a change email without any reservation code is an issue, never applied to a guessed row', () => {
  const s = setup({ messages: [createMessage({ id: 'x2', subject: 'Reservation altered', body: 'Indtjekning   Udtjekning\ntirs. 3. nov.   fre. 6. nov.\n', date: '2026-10-09T09:00:00Z' })], mode: 'all' });
  s.run();
  assert.equal(s.context.openParseIssues_('lifecycle')[0].reason, 'code-missing');
  assert.equal(s.guestData.data[1][3], '2026-11-02');
});

test('issue and ledger sheets never contain guest names or email bodies', () => {
  const s = setup({ messages: [cancelMsg('c1', 'HMAAAAAAA1'), changeMsg('x1', 'HMAAAAAAA1', 'ingen datoer')] });
  s.run();
  const dump = JSON.stringify([s.book.ParseIssues?.data, s.book.ReservationEvents?.data]);
  assert.doesNotMatch(dump, /Synthetic Guest|annulleret\.|ingen datoer/);
});

test('mode values other than off/cancellations/all fall back to off', () => {
  const s = setup({ messages: [cancelMsg('c1', 'HMAAAAAAA1')], mode: 'yes please' });
  assert.equal(s.run().mode, 'off');
  assert.equal(statusOf(s), '');
});
