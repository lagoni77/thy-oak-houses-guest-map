// Integration: patched Code.gs (code-gs-coming-guests.patch applied to live v20) +
// Heartbeat.gs + Lifecycle.gs. Runs only when GUEST_MAP_CODE_GS points at a
// patched private copy (the deployed Code.gs embeds a credential and is never
// committed). Synthetic messages only; MailApp throws if anything tries to send.
import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import { createGas, createSheet, createMessage } from './gas-harness.mjs';

const CODE = process.env.GUEST_MAP_CODE_GS;
const skip = CODE ? false : 'set GUEST_MAP_CODE_GS to a patched Code.gs';
const HB = new URL('../Heartbeat.gs', import.meta.url).pathname;
const LC = new URL('../Lifecycle.gs', import.meta.url).pathname;
const HEAD = ['Listing', 'ID', 'Name', 'CheckIn', 'CheckOut', 'Location', 'Lat', 'Lng', 'EmailID', 'Nights', 'Payout', 'Status', 'Guests', 'BookingUrl', 'GuestPhoto', 'GuestMessage', 'OriginalMessage', 'DetectedLangHint'];

const confirmationBody = (code, dates = 'tirs. 3. nov.   fre. 6. nov.', nights = 3) =>
  `Hej Brian\n\nTiny Oak House | Synthetic\n\nIndtjekning   Udtjekning\n${dates}\n15:00   11:00\n\n${nights} nætter\n\nDu tjener 1.000,00 kr\n\nhttps://www.airbnb.dk/hosting/reservations/details/${code}`;
const confirm = (id, code, opts = {}) => createMessage({ id, subject: 'Bekræftet reservation – Synthetic Guest ankommer', body: opts.body || confirmationBody(code), date: opts.date || '2026-10-09T08:00:00Z' });
const cancel = (id, code, date = '2026-10-09T09:00:00Z') => createMessage({ id, subject: `Annulleret: Reservation ${code} for Tiny Oak House`, body: '', date });

function load(messages, mode) {
  const guestData = createSheet('GuestData', [HEAD]);
  const g = createGas({ files: [CODE, HB, LC], sheets: { GuestData: guestData }, messages });
  const set = (name, fn) => { g.context[`__${name}`] = fn; vm.runInContext(`${name} = __${name};`, g.context); };
  set('ensureNewColumns_', () => ({ Guests: 13, BookingUrl: 14, GuestPhoto: 15, GuestMessage: 16, OriginalMessage: 17, DetectedLangHint: 18 }));
  set('getOrGeocodeLocation', () => ({ lat: 0, lng: 0 }));
  set('buildPayoutSummary', () => null);
  if (mode) g.store.set('LIFECYCLE_APPLY', mode);
  const beat = () => JSON.parse(g.store.get('HEARTBEAT_bookings'));
  const feed = (priv) => JSON.parse(g.context.doGet({ parameter: priv ? { key: vm.runInContext('PRIVATE_KEY', g.context) } : {} }).text);
  return { ...g, guestData, beat, feed, sync: () => g.context.syncAirbnbEmails() };
}

test('adjacent-column confirmation lands with correct dates, a ledger event and a clean heartbeat', { skip }, () => {
  const s = load([confirm('m1', 'HMAAAAAAA1')]);
  assert.equal(s.sync(), 1);
  const row = s.guestData.data[1];
  assert.deepEqual([row[3], row[4], row[9]], ['2026-11-03', '2026-11-06', 3]);
  assert.equal(s.book.ReservationEvents.data.length, 2);
  assert.equal(s.store.get('LAST_SCAN_DATE'), '2026/10/10');
  assert.equal(s.beat().openIssues, 0);
  assert.ok(s.beat().lastSuccessAt);
});

test('parse failure: no row, an open issue, the bookmark held, and the heartbeat says so', { skip }, () => {
  const s = load([confirm('m1', 'HMAAAAAAA1', { body: confirmationBody('HMAAAAAAA1', 'ikke en dato'), date: '2026-10-04T08:00:00Z' })]);
  assert.equal(s.sync(), 0);
  assert.equal(s.guestData.data.length, 1);
  assert.equal(s.store.get('LAST_SCAN_DATE'), '2026/10/04');
  assert.equal(s.beat().openIssues, 1, 'collector ran, but it is not complete');
  assert.ok(s.beat().lastSuccessAt, 'freshness policy unchanged: a completed check is a successful check');
  assert.equal(s.book.ParseIssues.data[1][2], 'dates-missing');
});

test('failure then recovery: a transient read failure is retried and resolved', { skip }, () => {
  let calls = 0;
  const body = () => { calls += 1; if (calls === 1) throw new Error('Gmail hiccup'); return confirmationBody('HMAAAAAAA1'); };
  const s = load([confirm('m1', 'HMAAAAAAA1', { body, date: '2026-10-05T08:00:00Z' })]);
  s.sync();
  assert.equal(s.beat().openIssues, 1);
  assert.equal(s.store.get('LAST_SCAN_DATE'), '2026/10/05');
  s.sync();
  assert.equal(s.guestData.data.length, 2);
  assert.equal(s.beat().openIssues, 0);
  assert.equal(s.store.get('LAST_SCAN_DATE'), '2026/10/10');
});

test('replay: syncing twice adds no duplicate rows or ledger events', { skip }, () => {
  const s = load([confirm('m1', 'HMAAAAAAA1'), cancel('c1', 'HMAAAAAAA1')], 'cancellations');
  s.sync();
  const snapshot = JSON.stringify([s.guestData.data, s.book.ReservationEvents.data.map((r) => r.slice(0, 6))]);
  s.sync();
  assert.equal(JSON.stringify([s.guestData.data, s.book.ReservationEvents.data.map((r) => r.slice(0, 6))]), snapshot);
});

test('cancellation seen before its booking: applied when the booking arrives; both feeds drop it', { skip }, () => {
  const messages = [cancel('c1', 'HMBBBBBBB2', '2026-10-09T07:00:00Z')];
  const s = load(messages, 'cancellations');
  s.sync();
  messages.push(confirm('m2', 'HMBBBBBBB2'));
  s.sync();
  assert.equal(s.guestData.data[1][11], 'Cancelled');
  assert.equal(s.feed(false).guests.length, 0);
  assert.equal(s.feed(true).guests.length, 0);
});

test('default mode off: a cancellation is ledgered only; the booking stays in the feed', { skip }, () => {
  const s = load([confirm('m1', 'HMAAAAAAA1'), cancel('c1', 'HMAAAAAAA1')]);
  s.sync();
  assert.equal(s.guestData.data[1][11] ?? '', '');
  assert.equal(s.feed(true).guests.length, 1);
});

test('public/private separation: heartbeats (with openIssues) only in the keyed view; no issue or ledger data in either', { skip }, () => {
  const s = load([confirm('m1', 'HMAAAAAAA1'), confirm('m9', 'HMAAAAAAA9', { body: confirmationBody('HMAAAAAAA9', 'ikke en dato') })]);
  s.sync();
  const pub = s.feed(false);
  const priv = s.feed(true);
  assert.deepEqual(Object.keys(pub), ['guests', 'payouts']);
  assert.equal(priv.heartbeats.bookings.openIssues, 1);
  for (const f of [pub, priv]) assert.doesNotMatch(JSON.stringify(f), /dates-missing|ReservationEvents|ParseIssues|m9/);
});

test('non-confirmation messages inside a confirmation thread are skipped, not parsed or raised as issues', { skip }, () => {
  const s = load([
    confirm('m1', 'HMAAAAAAA1'),
    createMessage({ id: 'm1b', thread: 'm1', subject: 'Svar fra Synthetic Guest', body: 'Hej, vi glæder os!', date: '2026-10-09T10:00:00Z' }),
  ]);
  assert.equal(s.sync(), 1);
  assert.equal(s.guestData.data.length, 2);
  assert.equal(s.beat().openIssues, 0);
});

test('notice-only accepted change: the keyed feed carries DatesUnverifiedSince; the public feed does not', { skip }, () => {
  const s = load([
    confirm('m1', 'HMAAAAAAA1'),
    createMessage({ id: 'n1', subject: 'Reservationen blev opdateret', date: '2026-10-09T10:00:00Z',
      body: 'Reservationen blev opdateret\n\nÆndringen af reservationen er accepteret.\n\nhttps://www.airbnb.dk/hosting/reservations/details/HMAAAAAAA1' }),
  ], 'all');
  s.sync();
  const priv = s.feed(true).guests[0];
  assert.equal(priv.DatesUnverifiedSince, '2026-10-09T10:00:00.000Z');
  assert.equal(priv.CheckIn, '2026-11-03', 'last known dates kept, never inferred');
  assert.equal(Object.hasOwn(s.feed(false).guests[0], 'DatesUnverifiedSince'), false);
});
