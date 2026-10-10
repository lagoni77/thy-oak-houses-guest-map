// Minimal Apps Script sandbox for node tests: loads .gs sources into one vm
// context with in-memory PropertiesService, SpreadsheetApp and GmailApp fakes
// and a controllable clock. Synthetic data only.
import vm from 'node:vm';
import { readFileSync } from 'node:fs';

function formatDate(d, tz, fmt) {
  const parts = Object.fromEntries(new Intl.DateTimeFormat('en-CA', {
    timeZone: tz || 'UTC', year: 'numeric', month: '2-digit', day: '2-digit',
  }).formatToParts(new Date(d)).map((p) => [p.type, p.value]));
  const sep = String(fmt || 'yyyy/MM/dd').includes('-') ? '-' : '/';
  return [parts.year, parts.month, parts.day].join(sep);
}

export function createSheet(name, rows = []) {
  const data = rows.map((r) => [...r]);
  const width = () => Math.max(0, ...data.map((r) => r.length));
  const sheet = {
    name, data, writes: 0,
    getName: () => name,
    setName() {},
    getLastRow: () => data.length,
    getLastColumn: () => width(),
    getDataRange: () => ({ getValues: () => data.map((r) => { const out = [...r]; while (out.length < width()) out.push(''); return out; }) }),
    appendRow: (row) => { data.push([...row]); sheet.writes += 1; },
    getRange: (r, c, nr = 1, nc = 1) => ({
      getValue: () => (data[r - 1] || [])[c - 1] ?? '',
      setValue: (v) => { while (data.length < r) data.push([]); data[r - 1][c - 1] = v; sheet.writes += 1; },
      setValues: (vals) => { for (let i = 0; i < nr; i++) { while (data.length < r + i) data.push([]); for (let j = 0; j < nc; j++) data[r - 1 + i][c - 1 + j] = vals[i][j]; } sheet.writes += 1; },
    }),
    setFrozenRows() {},
    deleteRow: () => { throw new Error('tests forbid deleting rows'); },
  };
  return sheet;
}

export function createMessage({ id, subject, body = '', html = '', date }) {
  return { getId: () => id, getSubject: () => subject, getPlainBody: typeof body === 'function' ? body : () => body, getBody: () => html, getDate: () => new Date(date) };
}

export function createGas({ files, now = () => new Date('2026-10-10T10:00:00.000Z'), sheets = {}, messages = [] }) {
  const store = new Map();
  const logs = [];
  const RealDate = Date;
  class FakeDate extends RealDate { constructor(...a) { super(...(a.length ? a : [now().getTime()])); } static now() { return now().getTime(); } }
  const book = { ...sheets };
  const spreadsheet = {
    getSheetByName: (n) => book[n] || null,
    insertSheet: (n) => (book[n] = createSheet(n)),
    getSheets: () => Object.values(book),
  };
  const gmail = {
    queries: [],
    // Synthetic search: every message whose subject contains one of the quoted subject phrases.
    search: (query) => {
      gmail.queries.push(query);
      const phrases = [...String(query).matchAll(/subject:"([^"]+)"/g)].map((m) => m[1].toLowerCase());
      const hits = messages.filter((m) => !phrases.length || phrases.some((p) => m.getSubject().toLowerCase().includes(p)));
      return hits.map((m) => ({ getMessages: () => [m] }));
    },
    getMessageById: (id) => messages.find((m) => m.getId() === id) || null,
  };
  const context = vm.createContext({
    Date: FakeDate, JSON, Math, String, Number, Object, Array, Set, Map, RegExp, Error, isNaN, parseFloat, parseInt,
    Logger: { log: (m) => logs.push(String(m)) },
    PropertiesService: { getScriptProperties: () => ({
      getProperty: (k) => (store.has(k) ? store.get(k) : null),
      setProperty: (k, v) => { store.set(k, String(v)); },
      deleteProperty: (k) => { store.delete(k); },
    }) },
    Utilities: { formatDate },
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet },
    GmailApp: gmail,
    MailApp: { sendEmail: () => { throw new Error('tests must never send mail'); } },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: (text) => ({ text, setMimeType() { return this; } }) },
  });
  for (const file of files) vm.runInContext(readFileSync(file, 'utf8'), context, { filename: file });
  return { context, store, logs, book, gmail };
}
