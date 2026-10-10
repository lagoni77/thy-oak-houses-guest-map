// ═══════════════════════════════════════════════════════════════════════════════
// RESERVATION LIFECYCLE + STAY-DATE PARSING — coming-guests repair (2026-10-10)
//
// Pure helpers first (tested in node: tests/lifecycle.test.mjs), then the sheet
// and Gmail adapters used by Code.gs (tests/lifecycle-sheet.test.mjs).
//
// - parseStayDates_: Danish "Indtjekning / Udtjekning" in both the stacked layout
//   and the adjacent-column table layout (both headers on one line, both dates
//   on the next). Year inferred from the email's Copenhagen day; a stay may
//   cross New Year. Bad parses are rejected with a reason, never written.
// - foldReservationEvents_: confirmed / changed / cancelled events keyed by the
//   Airbnb reservation code fold to one state. Idempotent per email id and
//   independent of processing order: cancellation is terminal; dates come from
//   the latest dated event.
// - Parse issues are recorded (email id + reason code only, no content) and
//   hold the scan bookmark so failed emails stay inside the re-scan window.
//
// Change-accepted subject patterns are NOT verified against real Airbnb mail;
// see CHANGE_SUBJECTS and COMING_GUESTS_REPAIR.md before enabling them.
// ═══════════════════════════════════════════════════════════════════════════════

var DK_MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, maj: 5, jun: 6, jul: 7, aug: 8, sep: 9, okt: 10, nov: 11, dec: 12 };
var DK_DATE_RE = /(\d{1,2})\.\s*(jan|feb|mar|apr|maj|jun|jul|aug|sep|okt|nov|dec)[a-zæøå]*\.?(?:[ \t]+(\d{4}))?/gi;
var MAX_STAY_NIGHTS = 60;
var DAY_MS = 86400000;

function ymd_(y, m, d) { return y + '-' + String(m).padStart(2, '0') + '-' + String(d).padStart(2, '0'); }
function dayMs_(s) { return Date.parse(s + 'T00:00:00Z'); }

function dateTokens_(line) {
  var out = [];
  var re = new RegExp(DK_DATE_RE.source, 'gi');
  var m;
  while ((m = re.exec(line)) !== null) {
    out.push({ day: Number(m[1]), month: DK_MONTHS[m[2].toLowerCase()], year: m[3] ? Number(m[3]) : null });
  }
  return out;
}

// Up to `limit` date tokens from the non-empty lines after line index `from`,
// stopping at `stopAt` (exclusive) when given.
function datesAfter_(lines, from, limit, stopAt) {
  var found = [];
  var seen = 0;
  for (var i = from + 1; i < lines.length && seen < 4; i++) {
    if (stopAt != null && i >= stopAt) break;
    if (!lines[i].trim()) continue;
    seen++;
    found = found.concat(dateTokens_(lines[i]));
    if (found.length >= limit) break;
  }
  return found.slice(0, limit);
}

/**
 * parseStayDates_(body, emailDay 'YYYY-MM-DD' in Europe/Copenhagen)
 *   -> { ok:true, checkIn, checkOut } | { ok:false, reason }
 */
function parseStayDates_(body, emailDay) {
  var lines = String(body || '').replace(/\r\n?/g, '\n').split('\n');
  var inIdx = -1;
  var outIdx = -1;
  for (var i = 0; i < lines.length; i++) {
    if (inIdx < 0 && /Indtjekning/i.test(lines[i])) inIdx = i;
    if (outIdx < 0 && /Udtjekning/i.test(lines[i])) outIdx = i;
  }
  if (inIdx < 0 || outIdx < 0) return { ok: false, reason: 'dates-missing' };

  var tIn;
  var tOut;
  if (inIdx === outIdx) {
    // Adjacent columns: one header line, dates in the same column order below it.
    var pair = datesAfter_(lines, inIdx, 2);
    if (pair.length < 2) return { ok: false, reason: 'dates-missing' };
    var outFirst = lines[inIdx].search(/Udtjekning/i) < lines[inIdx].search(/Indtjekning/i);
    tIn = outFirst ? pair[1] : pair[0];
    tOut = outFirst ? pair[0] : pair[1];
  } else {
    tIn = datesAfter_(lines, inIdx, 1, outIdx > inIdx ? outIdx : null)[0];
    tOut = datesAfter_(lines, outIdx, 1, inIdx > outIdx ? inIdx : null)[0];
  }
  if (!tIn || !tOut) return { ok: false, reason: 'dates-missing' };

  var emailYear = Number(String(emailDay).slice(0, 4));
  var emailMs = dayMs_(emailDay);
  if (!Number.isFinite(emailMs)) return { ok: false, reason: 'email-date-unknown' };

  // Check-in: the email year unless that puts it more than a day before the email.
  var inYear = tIn.year || emailYear;
  if (!tIn.year && dayMs_(ymd_(inYear, tIn.month, tIn.day)) < emailMs - DAY_MS) inYear += 1;
  var checkIn = ymd_(inYear, tIn.month, tIn.day);

  // Check-out: the check-in year unless that is not after check-in (stay across New Year).
  var outYear = tOut.year || inYear;
  if (!tOut.year && dayMs_(ymd_(outYear, tOut.month, tOut.day)) <= dayMs_(checkIn)) outYear += 1;
  var checkOut = ymd_(outYear, tOut.month, tOut.day);

  var nights = (dayMs_(checkOut) - dayMs_(checkIn)) / DAY_MS;
  if (!Number.isFinite(nights)) return { ok: false, reason: 'dates-missing' };
  if (nights <= 0) return { ok: false, reason: 'dates-inverted' };
  if (nights > MAX_STAY_NIGHTS) return { ok: false, reason: 'dates-implausible' };
  return { ok: true, checkIn: checkIn, checkOut: checkOut };
}

/** Stated nights (from "N nætter"), when present, must agree with the dates. */
function validateStay_(stay, statedNights) {
  var nights = (dayMs_(stay.checkOut) - dayMs_(stay.checkIn)) / DAY_MS;
  if (statedNights != null && Number(statedNights) !== nights) return { ok: false, reason: 'nights-mismatch' };
  return { ok: true, nights: nights };
}

// ── subject classification ───────────────────────────────────────────────────
var CONFIRMED_SUBJECTS = [/Bekræftet reservation/i, /Reservation confirmed/i, /Ny bekræftet booking/i];
// Verified: auditMissingBookings has matched these in production.
var CANCELLED_SUBJECTS = [/^(?:Annulleret|Cancelled|Canceled):\s*Reservation\s+([A-Z0-9]{6,12})\b/i];
// UNVERIFIED candidates for an accepted reservation change. Check with
// classifyRecentAirbnbSubjects() before enabling change application.
var CHANGE_SUBJECTS = [
  /Reservation\s+(?:altered|changed|change accepted)[:\s-]*(HM[A-Z0-9]{6,10})?/i,
  /(?:Reservationsændring|Ændring af reservation|Reservationen er ændret)[^A-Z0-9]*(HM[A-Z0-9]{6,10})?/i,
];

function classifyAirbnbSubject_(subject) {
  var s = String(subject || '');
  var m;
  for (var i = 0; i < CANCELLED_SUBJECTS.length; i++) {
    m = s.match(CANCELLED_SUBJECTS[i]);
    if (m) return { kind: 'cancelled', code: m[1].toUpperCase(), verified: true };
  }
  for (i = 0; i < CHANGE_SUBJECTS.length; i++) {
    m = s.match(CHANGE_SUBJECTS[i]);
    if (m) return { kind: 'changed', code: m[1] ? m[1].toUpperCase() : null, verified: false };
  }
  for (i = 0; i < CONFIRMED_SUBJECTS.length; i++) {
    if (CONFIRMED_SUBJECTS[i].test(s)) return { kind: 'confirmed', code: null, verified: true };
  }
  return { kind: 'other', code: null, verified: false };
}

function reservationCodeFrom_(text) {
  var m = String(text || '').match(/reservations\/details\/(HM[A-Z0-9]+)/i) || String(text || '').match(/\b(HM[A-Z0-9]{8})\b/);
  return m ? m[1].toUpperCase() : null;
}

// ── lifecycle fold ───────────────────────────────────────────────────────────
var KIND_ORDER = { confirmed: 0, changed: 1, cancelled: 2 };

/** events: [{ emailId, kind, eventAt(ms), checkIn?, checkOut? }] for ONE reservation code. */
function foldReservationEvents_(events) {
  var byId = {};
  (events || []).forEach(function (e) { if (e && e.emailId && !byId[e.emailId]) byId[e.emailId] = e; });
  var list = Object.keys(byId).map(function (k) { return byId[k]; }).sort(function (a, b) {
    return (a.eventAt - b.eventAt) || (KIND_ORDER[a.kind] - KIND_ORDER[b.kind]) || (a.emailId < b.emailId ? -1 : a.emailId > b.emailId ? 1 : 0);
  });
  var state = { status: 'confirmed', checkIn: null, checkOut: null, datesFrom: null, eventCount: list.length };
  list.forEach(function (e) {
    if (e.kind === 'cancelled') state.status = 'cancelled';
    if ((e.kind === 'confirmed' || e.kind === 'changed') && e.checkIn && e.checkOut) {
      state.checkIn = e.checkIn;
      state.checkOut = e.checkOut;
      state.datesFrom = e.emailId;
    }
  });
  if (!list.some(function (e) { return e.kind === 'confirmed' || e.kind === 'changed'; }) && state.status !== 'cancelled') {
    state.status = 'unknown';
  }
  return state;
}

// ── scan bookmark ────────────────────────────────────────────────────────────
/** today 'yyyy/MM/dd'; openIssues [{ emailDay:'YYYY-MM-DD' }] -> 'yyyy/MM/dd' no later than the oldest issue. */
function heldScanBookmark_(today, openIssues) {
  var oldest = null;
  (openIssues || []).forEach(function (issue) {
    var d = String(issue && issue.emailDay || '');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) return;
    if (!oldest || d < oldest) oldest = d;
  });
  if (!oldest) return today;
  var held = oldest.replace(/-/g, '/');
  return held < today ? held : today;
}

// ═══════════════════════════════════════════════════════════════════════════════
// Sheet + Gmail adapters (private spreadsheet only; nothing here is exposed by doGet)
// ═══════════════════════════════════════════════════════════════════════════════

var ISSUES_SHEET = 'ParseIssues';
var ISSUE_HEADERS = ['EmailID', 'Collector', 'Reason', 'EmailDay', 'FirstSeenAt', 'LastAttemptAt', 'Attempts', 'ResolvedAt'];
var LEDGER_SHEET = 'ReservationEvents';
var LEDGER_HEADERS = ['EmailID', 'Code', 'Kind', 'EventAt', 'CheckIn', 'CheckOut', 'RecordedAt'];
var LIFECYCLE_QUERY = 'from:automated@airbnb.com (subject:"Annulleret: Reservation" OR subject:"Cancelled: Reservation" OR ' +
  'subject:"Canceled: Reservation" OR subject:"Reservation altered" OR subject:"Reservationsændring" OR subject:"Ændring af reservation")';
var LIFECYCLE_MODES = { off: true, cancellations: true, all: true };

function lcSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function cellDay_(v) {
  if (v === '' || v == null) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, 'Europe/Copenhagen', 'yyyy-MM-dd');
  return String(v).slice(0, 10);
}

function messageDay_(message) {
  return Utilities.formatDate(message.getDate(), 'Europe/Copenhagen', 'yyyy-MM-dd');
}

function lcShiftDate_(slashYmd, deltaDays) {
  var p = String(slashYmd).split('/').map(Number);
  var d = new Date(Date.UTC(p[0], p[1] - 1, p[2]));
  d.setUTCDate(d.getUTCDate() + deltaDays);
  return Utilities.formatDate(d, 'UTC', 'yyyy/MM/dd');
}

// ── parse issues (email id + reason code only) ───────────────────────────────
function issueRows_() {
  var values = lcSheet_(ISSUES_SHEET, ISSUE_HEADERS).getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    out.push({ row: i + 1, emailId: String(r[0]), collector: r[1], reason: r[2], emailDay: cellDay_(r[3]), attempts: Number(r[6]) || 0, resolvedAt: r[7] || '' });
  }
  return out;
}

function recordParseIssue_(emailId, collector, reason, emailDay) {
  var sheet = lcSheet_(ISSUES_SHEET, ISSUE_HEADERS);
  var nowIso = new Date().toISOString();
  var rows = issueRows_().filter(function (x) { return x.emailId === String(emailId); });
  if (rows.some(function (x) { return /^dismissed /.test(String(x.resolvedAt)); })) return; // a human ruled it out
  var open = rows.filter(function (x) { return !x.resolvedAt; })[0];
  if (open) {
    sheet.getRange(open.row, 3).setValue(reason);
    sheet.getRange(open.row, 6).setValue(nowIso);
    sheet.getRange(open.row, 7).setValue(open.attempts + 1);
    return;
  }
  sheet.appendRow([String(emailId), collector, reason, emailDay || '', nowIso, nowIso, 1, '']);
}

function resolveParseIssue_(emailId) {
  var sheet = lcSheet_(ISSUES_SHEET, ISSUE_HEADERS);
  issueRows_().forEach(function (x) {
    if (x.emailId === String(emailId) && !x.resolvedAt) sheet.getRange(x.row, 8).setValue(new Date().toISOString());
  });
}

/**
 * Manual (editor): mark an email that will never parse — e.g. a non-booking
 * template caught by the subject query — as dismissed, with a reason, so it
 * stops holding the scan bookmark and is not re-opened. Returns false if no
 * open issue exists for that email id.
 */
function dismissParseIssue(emailId, why) {
  var sheet = lcSheet_(ISSUES_SHEET, ISSUE_HEADERS);
  var open = issueRows_().filter(function (x) { return x.emailId === String(emailId) && !x.resolvedAt; });
  open.forEach(function (x) { sheet.getRange(x.row, 8).setValue('dismissed ' + new Date().toISOString() + ' ' + String(why || '').slice(0, 80)); });
  return open.length > 0;
}

function openParseIssues_(collector) {
  return issueRows_().filter(function (x) { return !x.resolvedAt && (!collector || x.collector === collector); });
}

// ── reservation event ledger ─────────────────────────────────────────────────
function ledgerRows_() {
  var values = lcSheet_(LEDGER_SHEET, LEDGER_HEADERS).getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    out.push({ emailId: String(r[0]), code: String(r[1]), kind: r[2], eventAt: Number(r[3]), checkIn: cellDay_(r[4]) || null, checkOut: cellDay_(r[5]) || null });
  }
  return out;
}

function appendLedgerEvent_(event) {
  var sheet = lcSheet_(LEDGER_SHEET, LEDGER_HEADERS);
  if (ledgerRows_().some(function (e) { return e.emailId === String(event.emailId); })) return false;
  sheet.appendRow([String(event.emailId), event.code, event.kind, event.eventAt, event.checkIn || '', event.checkOut || '', new Date().toISOString()]);
  return true;
}

function lifecycleMode_() {
  var mode = String(PropertiesService.getScriptProperties().getProperty('LIFECYCLE_APPLY') || 'off').trim().toLowerCase();
  return LIFECYCLE_MODES[mode] ? mode : 'off';
}

// ── apply folded state to GuestData (idempotent; never deletes) ──────────────
function guestColumns_(headers) {
  var idx = {};
  headers.forEach(function (h, i) { idx[String(h || '').toLowerCase().replace(/[^a-z]/g, '')] = i + 1; });
  return { checkIn: idx.checkin, checkOut: idx.checkout, nights: idx.nights, status: idx.status, bookingUrl: idx.bookingurl };
}

function isCancelledStatus_(v) { return /(cancelled|canceled|aflyst|annuller|annuler|afbestilt)/i.test(String(v || '')); }

var LIFECYCLE_LOG_HEADERS = ['At', 'Row', 'Code', 'Field', 'Old', 'New'];

function applyReservationState_(sheet, values, cols, code, state, events, mode, dryRun) {
  var fieldName = function (c) { return c === cols.status ? 'Status' : c === cols.checkIn ? 'CheckIn' : c === cols.checkOut ? 'CheckOut' : 'Nights'; };
  // Every applied write is logged with its old value (rollback: write Old back).
  var write = function (r, c, v) {
    if (dryRun) return;
    var old = values[r - 1][c - 1];
    sheet.getRange(r, c).setValue(v);
    lcSheet_('LifecycleLog', LIFECYCLE_LOG_HEADERS).appendRow([new Date().toISOString(), r, code, fieldName(c),
      c === cols.checkIn || c === cols.checkOut ? cellDay_(old) : (old == null ? '' : old), v]);
  };
  if (mode === 'off') return 'disabled';
  if (!cols.bookingUrl) return 'row-not-found';
  var rowNo = -1;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][cols.bookingUrl - 1]).toUpperCase().indexOf('/DETAILS/' + code) >= 0) { rowNo = i + 1; break; }
  }
  if (rowNo < 0) return 'row-not-found';
  var row = values[rowNo - 1];
  var wrote = false;

  if (state.status === 'cancelled') {
    if (!cols.status) return 'no-status-column';
    if (!isCancelledStatus_(row[cols.status - 1])) { write(rowNo, cols.status, 'Cancelled'); wrote = true; }
    return wrote ? 'applied' : 'unchanged';
  }

  var datesEvent = events.filter(function (e) { return e.emailId === state.datesFrom; })[0];
  if (mode === 'all' && datesEvent && datesEvent.kind === 'changed' && state.checkIn && state.checkOut) {
    if (cellDay_(row[cols.checkIn - 1]) !== state.checkIn) { write(rowNo, cols.checkIn, state.checkIn); wrote = true; }
    if (cellDay_(row[cols.checkOut - 1]) !== state.checkOut) { write(rowNo, cols.checkOut, state.checkOut); wrote = true; }
    var nights = (dayMs_(state.checkOut) - dayMs_(state.checkIn)) / DAY_MS;
    if (cols.nights && Number(row[cols.nights - 1]) !== nights) { write(rowNo, cols.nights, nights); wrote = true; }
  }
  return wrote ? 'applied' : 'unchanged';
}

/** Re-folds the ledger for each code and applies it. Safe to call repeatedly. */
function applyLifecycleFor_(sheet, codes, mode, dryRun) {
  mode = mode || lifecycleMode_();
  var ledger = ledgerRows_();
  var values = sheet.getDataRange().getValues();
  var cols = guestColumns_(values[0] || []);
  var summary = { applied: 0, unchanged: 0, rowNotFound: 0, pending: 0 };
  var unique = {};
  (codes || []).forEach(function (c) { if (c) unique[c] = true; });
  Object.keys(unique).forEach(function (code) {
    var events = ledger.filter(function (e) { return e.code === code; });
    if (!events.length) return;
    var state = foldReservationEvents_(events);
    var outcome = applyReservationState_(sheet, values, cols, code, state, events, mode, dryRun);
    if (outcome === 'applied') { summary.applied++; if (!dryRun) values = sheet.getDataRange().getValues(); }
    else if (outcome === 'unchanged') summary.unchanged++;
    else if (outcome === 'row-not-found') summary.rowNotFound++;
    var needsChange = events.some(function (e) { return e.kind === 'changed'; }) && state.status !== 'cancelled';
    if (outcome === 'disabled' || (needsChange && mode === 'cancellations')) summary.pending++;
  });
  return summary;
}

/**
 * Scans cancellation / change emails since the lifecycle bookmark, ledgers each
 * readable event once, records unreadable ones as issues (holding the bookmark),
 * then applies the folded state per the LIFECYCLE_APPLY mode (default off).
 */
function runLifecycleScan_(sheet) {
  var props = PropertiesService.getScriptProperties();
  var mode = lifecycleMode_();
  var last = props.getProperty('LAST_LIFECYCLE_SCAN_DATE');
  var start = last ? lcShiftDate_(last, -2) : '2024/09/01';
  var threads = GmailApp.search(LIFECYCLE_QUERY + ' after:' + start);
  var known = {};
  ledgerRows_().forEach(function (e) { known[e.emailId] = true; });
  var touched = [];
  var issues = 0;

  threads.forEach(function (thread) {
    thread.getMessages().forEach(function (message) {
      var id = message.getId();
      if (known[id]) return;
      var cls = classifyAirbnbSubject_(message.getSubject());
      if (cls.kind !== 'cancelled' && cls.kind !== 'changed') return;
      var day = messageDay_(message);
      var body;
      try { body = message.getPlainBody(); } catch (e) { recordParseIssue_(id, 'lifecycle', 'read-failed', day); issues++; return; }
      var code = cls.code || reservationCodeFrom_(body);
      if (!code) { recordParseIssue_(id, 'lifecycle', 'code-missing', day); issues++; return; }
      var event = { emailId: id, code: code, kind: cls.kind, eventAt: message.getDate().getTime() };
      if (cls.kind === 'changed') {
        var dates = parseStayDates_(body, day);
        if (!dates.ok) { recordParseIssue_(id, 'lifecycle', dates.reason, day); issues++; return; }
        event.checkIn = dates.checkIn;
        event.checkOut = dates.checkOut;
      }
      appendLedgerEvent_(event);
      known[id] = true;
      resolveParseIssue_(id);
      touched.push(code);
    });
  });

  var today = Utilities.formatDate(new Date(), 'Europe/Copenhagen', 'yyyy/MM/dd');
  props.setProperty('LAST_LIFECYCLE_SCAN_DATE', heldScanBookmark_(today, openParseIssues_('lifecycle')));
  var summary = applyLifecycleFor_(sheet, touched, mode);
  summary.mode = mode;
  summary.issues = issues;
  summary.events = touched.length;
  return summary;
}


// ═══════════════════════════════════════════════════════════════════════════════
// Rollout tools — run manually from the editor. Logs carry counts, row numbers and
// reservation codes only (never names, messages or email bodies).
// ═══════════════════════════════════════════════════════════════════════════════

/** Read-only: what would LIFECYCLE_APPLY=<mode> change, given the ledger so far? */
function previewLifecycle(mode) {
  mode = LIFECYCLE_MODES[mode] ? mode : 'cancellations';
  var codes = ledgerRows_().map(function (e) { return e.code; });
  var r = applyLifecycleFor_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GuestData'), codes, mode, true);
  var out = { mode: mode, wouldApply: r.applied, unchanged: r.unchanged, rowNotFound: r.rowNotFound, pending: r.pending };
  Logger.log('previewLifecycle ' + JSON.stringify(out));
  return out;
}

var TEMPLATE_WORDS = /^(Bekræftet|Annulleret|Reservation|Reservationen|Reservationsændring|Ændring|Cancelled|Canceled|Confirmed|Ny|Din|Dit|Your|Booking|Udbetaling|Payout|Påmindelse|Reminder|Anmodning|Request|Altered|Changed|Change|Accepted|Opdatering|Update)$/;

function subjectTemplate_(subject) {
  return String(subject || '')
    .replace(/\bHM[A-Z0-9]{6,10}\b/g, '<CODE>')
    .replace(/\d+/g, '#')
    .split(/\s+/)
    .map(function (w) { return /^[A-ZÆØÅ]/.test(w) && !TEMPLATE_WORDS.test(w.replace(/[^A-Za-zÆØÅæøå]/g, '')) ? '<X>' : w; })
    .join(' ')
    .slice(0, 80);
}

/** Read-only: counts Airbnb subject templates of the last 400 days by classification. */
function classifyRecentAirbnbSubjects() {
  var counts = { confirmed: 0, cancelled: 0, changed: 0, other: 0, templates: {} };
  searchAllLc_('from:automated@airbnb.com newer_than:400d').forEach(function (thread) {
    thread.getMessages().forEach(function (message) {
      var cls = classifyAirbnbSubject_(message.getSubject());
      counts[cls.kind] = (counts[cls.kind] || 0) + 1;
      var key = cls.kind + ': ' + subjectTemplate_(message.getSubject());
      counts.templates[key] = (counts.templates[key] || 0) + 1;
    });
  });
  Logger.log('classifyRecentAirbnbSubjects ' + JSON.stringify(counts));
  return counts;
}

function searchAllLc_(query) {
  var out = [];
  for (var start = 0; ; start += 500) {
    var batch = GmailApp.search(query, start, 500);
    out = out.concat(batch);
    if (batch.length < 500) break;
  }
  return out;
}

var REPAIR_HEADERS = ['At', 'Row', 'Code', 'OldCheckIn', 'OldCheckOut', 'OldNights', 'NewCheckIn', 'NewCheckOut', 'NewNights'];

/** Dry run: re-parse each GuestData row's own confirmation email with the fixed parser. */
function auditStayDates() { return stayDateRepair_(false); }
/** Applies auditStayDates' candidates; every change is logged to RepairLog for rollback. */
function repairStayDatesConfirmed() { return stayDateRepair_(true); }

function stayDateRepair_(apply) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GuestData');
  var values = sheet.getDataRange().getValues();
  var headers = values[0] || [];
  var cols = guestColumns_(headers);
  var idx = {};
  headers.forEach(function (h, i) { idx[String(h || '').toLowerCase().replace(/[^a-z]/g, '')] = i + 1; });
  var emailCol = idx.emailid;
  var out = { unchanged: 0, wouldFix: 0, fixed: 0, emailMissing: 0, notDanish: 0, unparseable: {}, candidates: [] };
  var log = apply ? lcSheet_('RepairLog', REPAIR_HEADERS) : null;

  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var emailId = emailCol ? String(row[emailCol - 1] || '') : '';
    if (!emailId) continue;
    var message = GmailApp.getMessageById(emailId);
    if (!message) { out.emailMissing++; continue; }
    var body = message.getPlainBody();
    if (!/Indtjekning|Udtjekning/i.test(body)) { out.notDanish++; continue; }
    var dates = parseStayDates_(body, messageDay_(message));
    if (dates.ok) {
      var stated = (String(body).match(/(\d+)\s+(?:nætter|nights)/i) || [])[1];
      var v = validateStay_(dates, stated != null ? Number(stated) : null);
      if (!v.ok) dates = { ok: false, reason: v.reason };
    }
    if (!dates.ok) { out.unparseable[dates.reason] = (out.unparseable[dates.reason] || 0) + 1; continue; }

    var oldIn = cellDay_(row[cols.checkIn - 1]);
    var oldOut = cellDay_(row[cols.checkOut - 1]);
    var oldNights = cols.nights ? row[cols.nights - 1] : '';
    var newNights = (dayMs_(dates.checkOut) - dayMs_(dates.checkIn)) / DAY_MS;
    if (oldIn === dates.checkIn && oldOut === dates.checkOut && Number(oldNights) === newNights) { out.unchanged++; continue; }

    var code = cols.bookingUrl ? reservationCodeFrom_(row[cols.bookingUrl - 1]) : null;
    var kind = oldIn && oldIn === oldOut ? 'checkout-equals-checkin'
      : (oldIn.slice(5) === dates.checkIn.slice(5) && oldIn.slice(0, 4) !== dates.checkIn.slice(0, 4)) ? 'year' : 'other';
    out.candidates.push({ row: r + 1, code: code, kind: kind });
    out.wouldFix++;
    if (apply) {
      sheet.getRange(r + 1, cols.checkIn).setValue(dates.checkIn);
      sheet.getRange(r + 1, cols.checkOut).setValue(dates.checkOut);
      if (cols.nights) sheet.getRange(r + 1, cols.nights).setValue(newNights);
      log.appendRow([new Date().toISOString(), r + 1, code || '', oldIn, oldOut, oldNights, dates.checkIn, dates.checkOut, newNights]);
      out.fixed++;
    }
  }
  Logger.log((apply ? 'repairStayDatesConfirmed ' : 'auditStayDates ') + JSON.stringify(out));
  return out;
}

// ── confirmation parsing with an observable outcome ──────────────────────────
/**
 * Wraps Code.gs parseAirbnbEmail: Danish stay dates come from parseStayDates_
 * (adjacent-column safe), stated nights must agree, and every failure has a
 * reason the caller records as a parse issue instead of silently skipping.
 */
function parseAirbnbEmailResult_(body, subject, emailDate, htmlBody, emailDay) {
  var danish = /Indtjekning|Udtjekning/i.test(String(body || ''));
  var dates = danish ? parseStayDates_(body, emailDay) : null;
  if (dates && !dates.ok) return { ok: false, reason: dates.reason };
  var parsed;
  try { parsed = parseAirbnbEmail(body, subject, emailDate, htmlBody); } catch (e) { return { ok: false, reason: 'parse-error' }; }
  if (!parsed) return { ok: false, reason: 'parse-failed' };
  if (dates) { parsed.checkIn = dates.checkIn; parsed.checkOut = dates.checkOut; }
  var stated = (String(body || '').match(/(\d+)\s+(?:nætter|nights)/i) || [])[1];
  var v = validateStay_(parsed, stated != null ? Number(stated) : null);
  if (!v.ok) return { ok: false, reason: v.reason };
  parsed.nights = v.nights;
  parsed.code = reservationCodeFrom_(parsed.bookingUrl) || reservationCodeFrom_(htmlBody) || reservationCodeFrom_(body);
  return { ok: true, value: parsed };
}
