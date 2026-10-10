// ═══════════════════════════════════════════════════════════════════════════════
// PRODUCER HEARTBEATS — Signal House (LeonidOS), staged 2026-10-10
//
// Records, per collector, when it last ATTEMPTED and last SUCCESSFULLY checked
// Gmail, plus how many found emails it could not read (openIssues). A heartbeat
// proves the collector ran — never that the booking data is complete or correct. A run that finds no new emails is still a success; success advances
// only after the collector completed. Stored in Script Properties
// (HEARTBEAT_bookings / HEARTBEAT_payouts) and exposed read-only by doGet() in
// the private (keyed) view as `heartbeats`.
//
// Error codes are short fixed strings — never exception text, which can carry
// email content. Approved thresholds (Mainframe signal-house/CONTRACT.md):
// warning after 3 h, Telegram after 6 h without a successful check.
// ═══════════════════════════════════════════════════════════════════════════════

var HEARTBEAT_COLLECTORS = {
  bookings: 'gas:guest-map:syncAirbnbEmails',
  payouts: 'gas:guest-map:syncPayoutEmails'
};

function heartbeatKey_(collector) {
  if (!Object.prototype.hasOwnProperty.call(HEARTBEAT_COLLECTORS, collector)) {
    throw new Error('unknown collector: ' + collector);
  }
  return 'HEARTBEAT_' + collector;
}

function readHeartbeat_(collector) {
  var raw = PropertiesService.getScriptProperties().getProperty(heartbeatKey_(collector));
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function writeHeartbeat_(collector, patch) {
  var prev = readHeartbeat_(collector) || { lastSuccessAt: null, lastErrorCode: null, lastNewRows: null, openIssues: null };
  var next = {
    producer: HEARTBEAT_COLLECTORS[collector],
    lastAttemptAt: patch.lastAttemptAt !== undefined ? patch.lastAttemptAt : prev.lastAttemptAt,
    lastSuccessAt: patch.lastSuccessAt !== undefined ? patch.lastSuccessAt : prev.lastSuccessAt,
    lastErrorCode: patch.lastErrorCode !== undefined ? patch.lastErrorCode : prev.lastErrorCode,
    lastNewRows: patch.lastNewRows !== undefined ? patch.lastNewRows : prev.lastNewRows,
    openIssues: patch.openIssues !== undefined ? patch.openIssues : (prev.openIssues !== undefined ? prev.openIssues : null)
  };
  PropertiesService.getScriptProperties().setProperty(heartbeatKey_(collector), JSON.stringify(next));
}

/** A failure the collector already handles quietly (it used to just `return`). */
function heartbeatFailure_(code) {
  var e = new Error(code);
  e.heartbeatCode = code;
  return e;
}

/**
 * Runs a collector body and records its heartbeat. `run` returns the number of
 * new rows (0 is a healthy quiet run). A heartbeatFailure_ is recorded and
 * swallowed, preserving the old return-quietly behaviour; anything else is
 * recorded as 'exception' and rethrown so the trigger still reports it.
 */
function withHeartbeat_(collector, run) {
  heartbeatKey_(collector);
  writeHeartbeat_(collector, { lastAttemptAt: new Date().toISOString() });
  var newRows;
  try {
    newRows = run();
  } catch (e) {
    writeHeartbeat_(collector, { lastErrorCode: e && e.heartbeatCode ? e.heartbeatCode : 'exception' });
    if (e && e.heartbeatCode) return undefined;
    throw e;
  }
  writeHeartbeat_(collector, {
    lastSuccessAt: new Date().toISOString(),
    lastErrorCode: null,
    lastNewRows: typeof newRows === 'number' ? newRows : null
  });
  return newRows;
}

/**
 * Emails the collector found but could not read or parse (see Lifecycle.gs
 * ParseIssues). A successful check with open issues is NOT a complete one.
 */
function setHeartbeatOpenIssues_(collector, count) {
  heartbeatKey_(collector);
  writeHeartbeat_(collector, { openIssues: Number.isInteger(count) && count >= 0 ? count : null });
}

function readHeartbeats_() {
  return { bookings: readHeartbeat_('bookings'), payouts: readHeartbeat_('payouts') };
}
