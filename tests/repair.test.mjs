// Rollout / repair tools: dry runs write nothing; confirmed repair is logged for rollback.
import test from 'node:test';
import assert from 'node:assert/strict';
import { createGas, createSheet, createMessage } from './gas-harness.mjs';

const LC = new URL('../Lifecycle.gs', import.meta.url).pathname;
const HEAD = ['Listing', 'ID', 'Name', 'CheckIn', 'CheckOut', 'Location', 'Lat', 'Lng', 'EmailID', 'Nights', 'Payout', 'Status', 'Guests', 'BookingUrl'];
const row = (emailId, code, inD, outD, nights) => ['Tiny Oak House', 'TOH-X', 'Synthetic Guest', inD, outD, 'Nowhere', 0, 0, emailId, nights, '1,00 kr', '', '2', `https://www.airbnb.dk/hosting/reservations/details/${code}`];
const confirmation = (id, body, date = '2026-05-01T08:00:00Z') => createMessage({ id, subject: 'Bekræftet reservation – Synthetic Guest ankommer', body, date });

function setup(rows, messages, mode) {
  const guestData = createSheet('GuestData', [HEAD, ...rows]);
  const gas = createGas({ files: [LC], sheets: { GuestData: guestData }, messages });
  if (mode) gas.store.set('LIFECYCLE_APPLY', mode);
  return { ...gas, guestData };
}

const ADJ = 'Indtjekning   Udtjekning\ntirs. 7. jul.   søn. 12. jul.\n5 nætter';

test('auditStayDates finds the adjacent-column damage (check-out == check-in) without writing', () => {
  const s = setup([row('m1', 'HMAAAAAAA1', '2026-07-07', '2026-07-07', 5), row('m2', 'HMAAAAAAA2', '2026-07-07', '2026-07-12', 5)],
    [confirmation('m1', ADJ), confirmation('m2', ADJ)]);
  const before = JSON.stringify(s.guestData.data);
  const r = JSON.parse(JSON.stringify(s.context.auditStayDates()));
  assert.equal(r.wouldFix, 1);
  assert.equal(r.unchanged, 1);
  assert.deepEqual(r.candidates.map((c) => c.code), ['HMAAAAAAA1']);
  assert.equal(r.candidates[0].kind, 'checkout-equals-checkin');
  assert.equal(JSON.stringify(s.guestData.data), before, 'dry run writes nothing');
  assert.equal(s.book.RepairLog, undefined);
  assert.doesNotMatch(s.logs.join('\n'), /Synthetic Guest/, 'logs carry row numbers and codes, never names');
});

test('auditStayDates flags a wrong year from the old heuristic', () => {
  const s = setup([row('m1', 'HMAAAAAAA1', '2026-08-10', '2026-08-13', 3)],
    [confirmation('m1', 'Indtjekning   Udtjekning\ntirs. 10. aug.   fre. 13. aug.\n3 nætter', '2026-10-09T08:00:00Z')]);
  const r = s.context.auditStayDates();
  assert.equal(r.candidates[0].kind, 'year');
});

test('repairStayDatesConfirmed fixes the row and logs old and new values for rollback', () => {
  const s = setup([row('m1', 'HMAAAAAAA1', '2026-07-07', '2026-07-07', 0)], [confirmation('m1', ADJ)]);
  const r = s.context.repairStayDatesConfirmed();
  assert.equal(r.fixed, 1);
  assert.deepEqual(s.guestData.data[1].slice(3, 5), ['2026-07-07', '2026-07-12']);
  assert.equal(s.guestData.data[1][9], 5);
  const log = s.book.RepairLog.data;
  assert.deepEqual(log[0], ['At', 'Row', 'Code', 'Field', 'Old', 'New', 'State']);
  assert.deepEqual(log.slice(1).map((r) => r.slice(1)), [[2, 'HMAAAAAAA1', 'CheckOut', '2026-07-07', '2026-07-12', 'applied'], [2, 'HMAAAAAAA1', 'Nights', 0, 5, 'applied']]);
  assert.equal(s.context.repairStayDatesConfirmed().fixed, 0, 'second run is a no-op');
});

test('repair never touches rows whose email is missing or unparseable; it reports them', () => {
  const s = setup([row('gone', 'HMAAAAAAA1', '2026-07-07', '2026-07-07', 5), row('m2', 'HMAAAAAAA2', '2026-07-07', '2026-07-07', 5)],
    [confirmation('m2', 'Indtjekning\n\nikke en dato\n\nUdtjekning\n')]);
  const r = JSON.parse(JSON.stringify(s.context.repairStayDatesConfirmed()));
  assert.equal(r.fixed, 0);
  assert.equal(r.emailMissing, 1);
  assert.deepEqual(r.unparseable, { 'dates-missing': 1 });
  assert.deepEqual(s.guestData.data[1].slice(3, 5), ['2026-07-07', '2026-07-07']);
});

test('previewLifecycle reports what a mode would change without writing GuestData', () => {
  const cancel = createMessage({ id: 'c1', subject: 'Annulleret: Reservation HMAAAAAAA1 for Tiny Oak House', body: '', date: '2026-10-09T08:00:00Z' });
  const s = setup([row('m1', 'HMAAAAAAA1', '2026-11-02', '2026-11-05', 3)], [cancel]);
  s.context.runLifecycleScan_(s.guestData); // mode off: ledger only
  const before = JSON.stringify(s.guestData.data);
  const p = JSON.parse(JSON.stringify(s.context.previewLifecycle('cancellations')));
  assert.equal(p.wouldApply, 1);
  assert.equal(JSON.stringify(s.guestData.data), before);
});

test('classifyRecentAirbnbSubjects counts templates without names or codes', () => {
  const s = setup([], [
    createMessage({ id: 'a', subject: 'Bekræftet reservation – Synthetic Guest ankommer 7. jul.', date: '2026-10-01T08:00:00Z' }),
    createMessage({ id: 'b', subject: 'Annulleret: Reservation HMAAAAAAA1 for Tiny Oak House', date: '2026-10-01T08:00:00Z' }),
    createMessage({ id: 'c', subject: 'Reservation altered: HMAAAAAAA2', date: '2026-10-01T08:00:00Z' }),
    createMessage({ id: 'd', subject: 'Synthetic Guest har sendt en besked', date: '2026-10-01T08:00:00Z' }),
  ]);
  const r = JSON.parse(JSON.stringify(s.context.classifyRecentAirbnbSubjects()));
  assert.equal(r.confirmed, 1);
  assert.equal(r.cancelled, 1);
  assert.equal(r.changed, 1);
  assert.equal(r.other, 1);
  assert.doesNotMatch(JSON.stringify(r) + s.logs.join('\n'), /Synthetic Guest|HMAAAAAAA/);
});
