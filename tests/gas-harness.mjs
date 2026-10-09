// Minimal Apps Script sandbox for node tests: loads .gs sources into one vm
// context with in-memory PropertiesService and controllable clock.
import vm from 'node:vm';
import { readFileSync } from 'node:fs';

export function createGas({ files, now = () => new Date('2026-10-10T10:00:00.000Z') }) {
  const store = new Map();
  const logs = [];
  const RealDate = Date;
  class FakeDate extends RealDate { constructor(...a) { super(...(a.length ? a : [now().getTime()])); } static now() { return now().getTime(); } }
  const context = vm.createContext({
    Date: FakeDate, JSON, Math, String, Number, Object, Array, Set, Map, RegExp, Error, isNaN, parseFloat, parseInt,
    Logger: { log: (m) => logs.push(String(m)) },
    PropertiesService: { getScriptProperties: () => ({
      getProperty: (k) => (store.has(k) ? store.get(k) : null),
      setProperty: (k, v) => { store.set(k, String(v)); },
      deleteProperty: (k) => { store.delete(k); },
    }) },
    Utilities: { formatDate: (d) => new RealDate(d).toISOString().slice(0, 10).replace(/-/g, '/') },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: (text) => ({ text, setMimeType() { return this; } }) },
  });
  for (const file of files) vm.runInContext(readFileSync(file, 'utf8'), context, { filename: file });
  return { context, store, logs };
}
