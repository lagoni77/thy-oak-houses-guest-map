// Codex review round 2 repros and regressions (synthetic data only).
import test from 'node:test';
import assert from 'node:assert/strict';
import { createGas, createSheet, createMessage } from './gas-harness.mjs';

const LC = new URL('../Lifecycle.gs', import.meta.url).pathname;
const HEAD = ['Listing', 'ID', 'Name', 'CheckIn', 'CheckOut', 'Location', 'Lat', 'Lng', 'EmailID', 'Nights', 'Payout', 'Status', 'Guests', 'BookingUrl'];
const CODE = 'HMAAAAAAA1';
const row = (inD = '2026-11-02', outD = '2026-11-05', nights = 3, emailId = 'conf-1') => ['Tiny Oak House', 'TOH-X', 'Synthetic Guest', inD, outD, 'Nowhere', 0, 0, emailId, nights, '1,00 kr', '', '2', `https://www.airbnb.dk/hosting/reservations/details/${CODE}`];
const cancelMsg = (id = 'c1', date = '2026-10-09T08:00:00Z') => createMessage({ id, subject: `Annulleret: Reservation ${CODE} for Tiny Oak House`, body: '', date });
const datedChange = (id, dates, date = '2026-10-09T09:00:00Z', nights = null) => createMessage({ id, subject: 'Reservationen blev opdateret', date,
  body: `Reservationen blev opdateret\n\nIndtjekning   Udtjekning\n${dates}\n${nights != null ? `${nights} nætter\n` : ''}\nhttps://www.airbnb.dk/hosting/reservations/details/${CODE}` });
// Representative notice-only accepted change: subject, short notice, reservation-detail link, no dates.
const noticeOnly = (id = 'n1', date = '2026-10-09T10:00:00Z') => createMessage({ id, subject: 'Reservationen blev opdateret', date,
  body: `Reservationen blev opdateret\n\nHej Brian\n\nÆndringen af reservationen er accepteret.\n\nSe reservation\nhttps://www.airbnb.dk/hosting/reservations/details/${CODE}\n\nAirbnb` });
const confirmation = (id, body, date = '2026-05-01T08:00:00Z') => createMessage({ id, subject: 'Bekræftet reservation – Synthetic Guest ankommer', body, date });

function setup({ rows = [row()], messages = [], mode = null } = {}) {
  const guestData = createSheet('GuestData', [HEAD, ...rows]);
  const gas = createGas({ files: [LC], sheets: { GuestData: guestData }, messages });
  if (mode) gas.store.set('LIFECYCLE_APPLY', mode);
  const plain = (x) => JSON.parse(JSON.stringify(x));
  return { ...gas, guestData, plain, run: () => plain(gas.context.runLifecycleScan_(guestData)), mode: (m) => gas.store.set('LIFECYCLE_APPLY', m) };
}
const cell = (s, field) => s.guestData.data[1][s.guestData.data[0].indexOf(field)];
const log = (s, name = 'LifecycleLog') => (s.book[name] ? s.book[name].data.slice(1) : []);

// ── P1: already-ledgered events must be applied when the mode is switched on ──
test('off → preview → cancellations: the already-ledgered cancellation is applied on the next run', () => {
  const s = setup({ messages: [cancelMsg()] });
  s.run();
  assert.equal(s.plain(s.context.previewLifecycle('cancellations')).wouldApply, 1);
  s.mode('cancellations');
  const r = s.run();
  assert.equal(r.applied, 1);
  assert.equal(cell(s, 'Status'), 'Cancelled');
});

test('cancellations → all: an already-ledgered dated change is applied on the next run', () => {
  const s = setup({ messages: [datedChange('x1', 'tirs. 3. nov.   fre. 6. nov.')], mode: 'cancellations' });
  s.run();
  assert.equal(cell(s, 'CheckIn'), '2026-11-02');
  s.mode('all');
  s.run();
  assert.equal(cell(s, 'CheckIn'), '2026-11-03');
});

test('applyLifecycleNow() applies the whole ledger explicitly and is idempotent', () => {
  const s = setup({ messages: [cancelMsg()] });
  s.run();
  s.mode('cancellations');
  assert.equal(s.plain(s.context.applyLifecycleNow()).applied, 1);
  assert.equal(s.plain(s.context.applyLifecycleNow()).applied, 0);
  assert.equal(log(s).length, 1);
});

// ── P1: notice-only accepted changes ──
test('notice-only accepted change: ledgered without dates, no parse issue, dates kept but flagged unverified (mode all)', () => {
  const s = setup({ messages: [noticeOnly()], mode: 'all' });
  const r = s.run();
  assert.equal(r.issues, 0);
  assert.equal(s.book.ReservationEvents.data[1][4], '', 'no invented check-in');
  assert.equal(cell(s, 'CheckIn'), '2026-11-02');
  assert.equal(cell(s, 'DatesUnverifiedSince'), '2026-10-09T10:00:00.000Z');
  assert.deepEqual(log(s).map((l) => l[3]), ['DatesUnverifiedSince']);
});

test('a later dated change clears the unverified flag and applies its dates', () => {
  const messages = [noticeOnly()];
  const s = setup({ messages, mode: 'all' });
  s.run();
  messages.push(datedChange('x2', 'tirs. 3. nov.   fre. 6. nov.', '2026-10-09T11:00:00Z'));
  s.run();
  assert.equal(cell(s, 'CheckIn'), '2026-11-03');
  assert.equal(cell(s, 'DatesUnverifiedSince'), '');
});

test('notice-only change without any reservation code is an issue', () => {
  const s = setup({ messages: [createMessage({ id: 'n9', subject: 'Reservationen blev opdateret', body: 'Ændringen er accepteret.', date: '2026-10-09T10:00:00Z' })], mode: 'all' });
  s.run();
  assert.equal(s.context.openParseIssues_('lifecycle')[0].reason, 'code-missing');
});

test('notice-only change in mode cancellations is pending, not applied', () => {
  const s = setup({ messages: [noticeOnly()], mode: 'cancellations' });
  assert.equal(s.run().pending, 1);
  assert.equal(cell(s, 'DatesUnverifiedSince'), undefined);
});

// ── validation gaps on the change path ──
test('a dated change whose stated nights disagree is rejected as an issue, never applied', () => {
  const s = setup({ messages: [datedChange('x1', 'tirs. 3. nov.   fre. 6. nov.', '2026-10-09T09:00:00Z', 5)], mode: 'all' });
  const r = s.run();
  assert.equal(r.issues, 1);
  assert.equal(s.context.openParseIssues_('lifecycle')[0].reason, 'nights-mismatch');
  assert.equal(cell(s, 'CheckIn'), '2026-11-02');
});

test('a change with an impossible date is rejected', () => {
  const s = setup({ messages: [datedChange('x1', 'man. 31. feb.   tirs. 3. mar.')], mode: 'all' });
  s.run();
  assert.equal(s.context.openParseIssues_('lifecycle')[0].reason, 'dates-invalid');
});

// ── P1: before-images persisted before mutation; failure injection; resumable ──
test('a failing GuestData write leaves a logged before-image marked failed and no change; the next run retries', () => {
  const s = setup({ messages: [cancelMsg()], mode: 'cancellations' });
  s.guestData.failSetValue = () => true;
  assert.throws(() => s.run(), /injected setValue/);
  assert.equal(cell(s, 'Status'), '');
  assert.deepEqual(log(s).map((l) => [l[3], l[4], l[5], l[6]]), [['Status', '', 'Cancelled', 'failed']]);
  s.guestData.failSetValue = null;
  s.run();
  assert.equal(cell(s, 'Status'), 'Cancelled');
  assert.deepEqual(log(s).map((l) => l[6]), ['failed', 'applied']);
});

test('if the before-image cannot be logged, nothing is mutated', () => {
  const s2 = setup({ messages: [cancelMsg()], mode: 'cancellations' });
  s2.context.lcSheet_('LifecycleLog', ['At', 'Row', 'Code', 'Field', 'Old', 'New', 'State']);
  s2.book.LifecycleLog.failAppend = () => true;
  assert.throws(() => s2.run(), /injected appendRow/);
  assert.equal(cell(s2, 'Status'), '');
});

test('crash between mutation and marking applied: the pending entry is reconciled, never written twice', () => {
  const s = setup({ messages: [cancelMsg()], mode: 'cancellations' });
  s.context.lcSheet_('LifecycleLog', ['At', 'Row', 'Code', 'Field', 'Old', 'New', 'State']);
  s.book.LifecycleLog.failSetValue = (r, c, v) => v === 'applied';
  assert.throws(() => s.run(), /injected setValue/);
  assert.equal(cell(s, 'Status'), 'Cancelled', 'the write itself happened');
  assert.equal(log(s)[0][6], 'pending');
  s.book.LifecycleLog.failSetValue = null;
  const writes = s.guestData.writes;
  s.run();
  assert.equal(log(s).length, 1, 'no second write logged');
  assert.equal(log(s)[0][6], 'applied');
  assert.equal(s.guestData.writes, writes);
});

test('rollback restores before-images idempotently, refuses while lifecycle is on, and flags human edits', () => {
  const s = setup({ messages: [cancelMsg()], mode: 'cancellations' });
  s.run();
  assert.throws(() => s.context.rollbackLifecycleChanges('2000-01-01T00:00:00Z'), /LIFECYCLE_APPLY=off/);
  s.mode('off');
  assert.equal(s.plain(s.context.rollbackLifecycleChanges('2000-01-01T00:00:00Z')).rolledBack, 1);
  assert.equal(cell(s, 'Status'), '');
  assert.equal(s.plain(s.context.rollbackLifecycleChanges('2000-01-01T00:00:00Z')).rolledBack, 0, 'idempotent');
  assert.equal(log(s)[0][6], 'rolled-back');

  const t = setup({ messages: [cancelMsg()], mode: 'cancellations' });
  t.run();
  t.mode('off');
  t.guestData.data[1][11] = 'Aflyst (manuelt)'; // human edit after the automated write
  const r = t.plain(t.context.rollbackLifecycleChanges('2000-01-01T00:00:00Z'));
  assert.equal(r.conflicts, 1);
  assert.equal(cell(t, 'Status'), 'Aflyst (manuelt)', 'a human edit is never overwritten');
});

test('stay-date repair is logged before mutation per field and can be rolled back', () => {
  const s = setup({ rows: [row('2026-07-07', '2026-07-07', 0, 'm1')], messages: [confirmation('m1', 'Indtjekning   Udtjekning\ntirs. 7. jul.   søn. 12. jul.\n5 nætter')] });
  assert.equal(s.plain(s.context.repairStayDatesConfirmed()).fixed, 1);
  assert.deepEqual(log(s, 'RepairLog').map((l) => [l[3], l[4], l[5], l[6]]), [['CheckOut', '2026-07-07', '2026-07-12', 'applied'], ['Nights', 0, 5, 'applied']]);
  assert.equal(s.plain(s.context.rollbackStayDateRepairs('2000-01-01T00:00:00Z')).rolledBack, 2);
  assert.deepEqual([cell(s, 'CheckOut'), cell(s, 'Nights')], ['2026-07-07', 0]);
});

// ── P2: repair must not revert accepted modifications ──
test('stay-date repair never reverts an accepted change; repeated repair and replay keep the change', () => {
  const s = setup({ rows: [row('2026-11-02', '2026-11-05', 3, 'm1')], messages: [
    confirmation('m1', `Indtjekning   Udtjekning\nman. 2. nov.   tors. 5. nov.\n3 nætter\nhttps://www.airbnb.dk/hosting/reservations/details/${CODE}`, '2026-09-01T08:00:00Z'),
    datedChange('x1', 'tirs. 3. nov.   fre. 6. nov.'),
  ], mode: 'all' });
  s.run();
  assert.equal(cell(s, 'CheckIn'), '2026-11-03');
  for (let i = 0; i < 2; i++) {
    const r = s.plain(s.context.repairStayDatesConfirmed());
    assert.equal(r.fixed, 0);
    assert.equal(r.lifecycleOwned, 1);
    s.run();
    assert.deepEqual([cell(s, 'CheckIn'), cell(s, 'CheckOut')], ['2026-11-03', '2026-11-06']);
  }
});

test('repair also leaves rows with a notice-only (unverified) change alone', () => {
  const s = setup({ rows: [row('2026-07-07', '2026-07-07', 0, 'm1')], messages: [confirmation('m1', 'Indtjekning   Udtjekning\ntirs. 7. jul.   søn. 12. jul.\n5 nætter'), noticeOnly()] });
  s.run(); // mode off: ledger only
  assert.equal(s.plain(s.context.auditStayDates()).lifecycleOwned, 1);
});
