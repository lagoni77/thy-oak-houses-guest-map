// Lifecycle.gs pure logic — synthetic fixtures only (no real guest data).
import test from 'node:test';
import assert from 'node:assert/strict';
import { createGas } from './gas-harness.mjs';

const LC = new URL('../Lifecycle.gs', import.meta.url).pathname;
const g = () => createGas({ files: [LC] }).context;

// ── stay dates: layouts ──
const stacked = 'Indtjekning\n\ntirs. 7. jul.\n\n15:00\n\nUdtjekning\n\nsøn. 12. jul.\n\n11:00\n\n5 nætter';
const adjacent = 'Indtjekning          Udtjekning\ntirs. 7. jul.          søn. 12. jul.\n15:00          11:00\n';
const adjacentTabs = 'Indtjekning\tUdtjekning\ntirs. 7. jul.\tsøn. 12. jul.\n';
const adjacentBlankLine = 'Indtjekning   Udtjekning\n\n   tirs. 7. jul.   søn. 12. jul.\n';
const adjacentReversed = 'Udtjekning   Indtjekning\nsøn. 12. jul.   tirs. 7. jul.\n';
const adjacentExplicitYears = 'Indtjekning   Udtjekning\nman. 28. dec. 2026   lør. 2. jan. 2027\n';

for (const [label, body] of [['stacked', stacked], ['adjacent columns (spaces)', adjacent], ['adjacent columns (tabs)', adjacentTabs],
  ['adjacent with blank line', adjacentBlankLine], ['adjacent, columns reversed', adjacentReversed]]) {
  test(`stay dates: ${label}`, () => {
    const r = g().parseStayDates_(body, '2026-05-01');
    assert.equal(r.ok, true, JSON.stringify(r));
    assert.equal(r.checkIn, '2026-07-07');
    assert.equal(r.checkOut, '2026-07-12');
  });
}

test('adjacent layout never copies the check-in date into check-out (the reported bug)', () => {
  const r = g().parseStayDates_(adjacent, '2026-05-01');
  assert.notEqual(r.checkOut, r.checkIn);
});

// ── stay dates: year rollover ──
test('booking made in December for January: next year', () => {
  const r = g().parseStayDates_('Indtjekning   Udtjekning\nfre. 8. jan.   søn. 10. jan.\n', '2026-12-20');
  assert.deepEqual([r.checkIn, r.checkOut], ['2027-01-08', '2027-01-10']);
});
test('stay crossing New Year: check-out rolls into the next year', () => {
  const r = g().parseStayDates_('Indtjekning   Udtjekning\nman. 28. dec.   lør. 2. jan.\n', '2026-11-03');
  assert.deepEqual([r.checkIn, r.checkOut], ['2026-12-28', '2027-01-02']);
});
test('booking made ten months ahead stays in the future, not the past', () => {
  const r = g().parseStayDates_('Indtjekning   Udtjekning\ntirs. 10. aug.   fre. 13. aug.\n', '2026-10-09');
  assert.deepEqual([r.checkIn, r.checkOut], ['2027-08-10', '2027-08-13']);
});
test('same-day / next-day bookings keep the email year', () => {
  assert.equal(g().parseStayDates_('Indtjekning   Udtjekning\nlør. 10. okt.   søn. 11. okt.\n', '2026-10-10').checkIn, '2026-10-10');
  assert.equal(g().parseStayDates_('Indtjekning   Udtjekning\nfre. 9. okt.   søn. 11. okt.\n', '2026-10-10').checkIn, '2026-10-09', 'one day of timezone grace');
});
test('explicit years are honoured', () => {
  const r = g().parseStayDates_(adjacentExplicitYears, '2026-01-01');
  assert.deepEqual([r.checkIn, r.checkOut], ['2026-12-28', '2027-01-02']);
});

// ── stay dates: rejection (observable reasons, never a silent bad row) ──
test('missing dates are rejected with a reason', () => {
  assert.deepEqual(JSON.parse(JSON.stringify(g().parseStayDates_('Hej Brian\nIngen datoer her', '2026-05-01'))), { ok: false, reason: 'dates-missing' });
});
test('only one date found is rejected', () => {
  assert.equal(g().parseStayDates_('Indtjekning\n\ntirs. 7. jul.\n', '2026-05-01').reason, 'dates-missing');
});
test('explicitly inverted dates are rejected', () => {
  assert.equal(g().parseStayDates_('Indtjekning   Udtjekning\nsøn. 12. jul. 2026   tirs. 7. jul. 2026\n', '2026-05-01').reason, 'dates-inverted');
});
test('implausibly long stays are rejected as a likely misparse', () => {
  assert.equal(g().parseStayDates_('Indtjekning   Udtjekning\ntirs. 7. jul. 2026   søn. 12. dec. 2026\n', '2026-05-01').reason, 'dates-implausible');
});
test('validateStay_ rejects stated nights that disagree with the dates', () => {
  const c = g();
  assert.equal(c.validateStay_({ checkIn: '2026-07-07', checkOut: '2026-07-12' }, 5).ok, true);
  assert.equal(c.validateStay_({ checkIn: '2026-07-07', checkOut: '2026-07-12' }, null).ok, true, 'no stated nights is fine');
  assert.equal(c.validateStay_({ checkIn: '2026-07-07', checkOut: '2026-07-07' }, 5).reason, 'nights-mismatch');
});

// ── subject classification ──
test('classifies confirmations, verified cancellation templates and (unverified) changes', () => {
  const c = g();
  assert.equal(c.classifyAirbnbSubject_('Bekræftet reservation – X ankommer 7. jul.').kind, 'confirmed');
  assert.deepEqual(JSON.parse(JSON.stringify(c.classifyAirbnbSubject_('Annulleret: Reservation HMABC12345 for Tiny Oak House'))), { kind: 'cancelled', code: 'HMABC12345', verified: true });
  assert.equal(c.classifyAirbnbSubject_('Canceled: Reservation HMABC12345 for X').code, 'HMABC12345');
  const change = c.classifyAirbnbSubject_('Reservation altered: HMABC12345');
  assert.equal(change.kind, 'changed');
  assert.equal(change.verified, false);
  assert.equal(c.classifyAirbnbSubject_('Ny besked fra gæst').kind, 'other');
});

// ── lifecycle fold: ordering, replay, idempotency ──
const ev = (emailId, kind, at, dates = {}) => ({ emailId, kind, eventAt: Date.parse(at), ...dates });
const confirm = ev('m1', 'confirmed', '2026-05-01T10:00:00Z', { checkIn: '2026-07-07', checkOut: '2026-07-12' });
const change = ev('m2', 'changed', '2026-05-10T10:00:00Z', { checkIn: '2026-07-08', checkOut: '2026-07-14' });
const change2 = ev('m3', 'changed', '2026-05-20T10:00:00Z', { checkIn: '2026-07-09', checkOut: '2026-07-14' });
const cancel = ev('m4', 'cancelled', '2026-06-01T10:00:00Z');

const permutations = (xs) => xs.length <= 1 ? [xs] : xs.flatMap((x, i) => permutations([...xs.slice(0, i), ...xs.slice(i + 1)]).map((p) => [x, ...p]));
const plain = (o) => JSON.parse(JSON.stringify(o));

test('fold: confirmed then accepted changes → latest change wins, in any processing order', () => {
  const expected = { status: 'confirmed', checkIn: '2026-07-09', checkOut: '2026-07-14', datesFrom: 'm3', eventCount: 3 };
  for (const order of permutations([confirm, change, change2])) assert.deepEqual(plain(g().foldReservationEvents_(order)), expected);
});

test('fold: cancellation is terminal regardless of order; later changes do not revive it', () => {
  for (const order of permutations([confirm, change, cancel])) assert.equal(g().foldReservationEvents_(order).status, 'cancelled');
  const lateChange = ev('m9', 'changed', '2026-07-01T10:00:00Z', { checkIn: '2026-08-01', checkOut: '2026-08-03' });
  assert.equal(g().foldReservationEvents_([confirm, cancel, lateChange]).status, 'cancelled');
});

test('fold: replaying the same email twice changes nothing (idempotent by email id)', () => {
  const once = plain(g().foldReservationEvents_([confirm, change]));
  const twice = plain(g().foldReservationEvents_([confirm, change, change, confirm]));
  assert.deepEqual(twice, once);
});

test('fold: an older change processed after a newer one does not overwrite it', () => {
  assert.equal(g().foldReservationEvents_([confirm, change2, change]).checkIn, '2026-07-09');
});

test('fold: a change seen before its confirmation still wins over the older confirmation dates', () => {
  assert.equal(g().foldReservationEvents_([change, confirm]).checkIn, '2026-07-08');
});

test('fold: cancellation with no known dates keeps dates unknown', () => {
  assert.deepEqual(plain(g().foldReservationEvents_([cancel])), { status: 'cancelled', checkIn: null, checkOut: null, datesFrom: null, eventCount: 1 });
});

// ── bookmark held by open parse issues ──
test('scan bookmark never advances past the oldest open parse issue', () => {
  const c = g();
  assert.equal(c.heldScanBookmark_('2026/10/10', []), '2026/10/10');
  assert.equal(c.heldScanBookmark_('2026/10/10', [{ emailDay: '2026-10-03' }, { emailDay: '2026-10-07' }]), '2026/10/03');
  assert.equal(c.heldScanBookmark_('2026/10/10', [{ emailDay: 'garbage' }]), '2026/10/10');
});

test('verified accepted-change subject "Reservationen blev opdateret" is a verified change', () => {
  const c = g().classifyAirbnbSubject_('Reservationen blev opdateret');
  assert.equal(c.kind, 'changed');
  assert.equal(c.verified, true);
});

test('change REQUESTS never classify as accepted changes', () => {
  for (const subject of ['Anmodning om ændring af reservation', 'Ændringsanmodning for reservation HMABC12345', 'Reservation change request: HMABC12345', 'Request to change reservation HMABC12345']) {
    assert.equal(g().classifyAirbnbSubject_(subject).kind, 'other', subject);
  }
});
