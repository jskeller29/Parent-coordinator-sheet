// A small fake spreadsheet for offline flow tests: tabs with a grid of values,
// hide / show / activate / insert / delete, and ranges that read and write values.
// Formatting calls (colors, fonts, notes, borders ...) are accepted and ignored.
// Not a full SpreadsheetApp: add what a new test needs.

function colNum(letters) { return letters.toUpperCase().split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0); }
function parseA1(a1) {
  const m = String(a1).match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i);
  if (!m) throw new Error("fake sheet: can't read range " + a1);
  const r1 = Number(m[2]), c1 = colNum(m[1]);
  const r2 = m[4] ? Number(m[4]) : r1, c2 = m[3] ? colNum(m[3]) : c1;
  return [r1, c1, r2 - r1 + 1, c2 - c1 + 1];
}

// Ranges and tabs are Proxies: anything not defined here is a formatting call,
// accepted and chained (returns the same object).
function makeRange(sheet, row, col, nr, nc) {
  const r = {
    getRow: () => row, getColumn: () => col, getNumRows: () => nr, getNumColumns: () => nc,
    getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => sheet._get(row + i, col + j))),
    getDisplayValues: () => r.getValues().map(rw => rw.map(v => v instanceof Date ? (v.getMonth() + 1) + '/' + v.getDate() + '/' + v.getFullYear() : String(v))),
    getValue: () => sheet._get(row, col),
    getDisplayValue: () => r.getDisplayValues()[0][0],
    setValues: vals => {
      if (vals.length !== nr || vals.some(v => v.length !== nc)) throw new Error("fake sheet: setValues size " + vals.length + "x" + (vals[0] || []).length + " into " + nr + "x" + nc);
      vals.forEach((rw, i) => rw.forEach((v, j) => sheet._set(row + i, col + j, v)));
      return p;
    },
    setValue: v => { sheet._set(row, col, v); return p; },
    clearContent: () => { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) sheet._set(row + i, col + j, ""); return p; },
    clear: () => r.clearContent(),
    setRichTextValue: v => { sheet._set(row, col, v && v.text !== undefined ? v.text : ""); return p; }
  };
  const p = new Proxy(r, { get(t, k) { return (k in t) ? t[k] : (typeof k === 'string' ? () => p : undefined); } });
  return p;
}

let nextId = 1000;
function makeSheet(ss, name) {
  const cells = new Map(); // "r,c" -> value
  const s = {
    _name: name, _id: nextId++, _hidden: false, _maxRows: 1000, _maxCols: 26,
    _get: (r, c) => (cells.has(r + ',' + c) ? cells.get(r + ',' + c) : ""),
    _set: (r, c, v) => { if (v === "" || v === null || v === undefined) cells.delete(r + ',' + c); else cells.set(r + ',' + c, v); },
    _cells: cells,
    getName: () => s._name, setName: n => { s._name = n; return p; }, getSheetId: () => s._id,
    isSheetHidden: () => s._hidden,
    showSheet: () => { s._hidden = false; return p; },
    hideSheet: () => {
      if (!ss._sheets.some(x => x !== p && !x.isSheetHidden())) throw new Error("fake sheet: can't hide every tab");
      s._hidden = true;
      if (ss._active === p) ss._active = ss._sheets.find(x => !x.isSheetHidden());
      return p;
    },
    activate: () => { s._hidden = false; ss._active = p; return p; },
    getMaxRows: () => s._maxRows, getMaxColumns: () => s._maxCols,
    insertRowsAfter: (a, n) => { s._maxRows += n; return p; }, insertColumnsAfter: (a, n) => { s._maxCols += n; return p; },
    deleteRows: (a, n) => { s._maxRows -= n; return p; }, deleteColumns: (a, n) => { s._maxCols -= n; return p; },
    getLastRow: () => { let m = 0; cells.forEach((v, k) => { m = Math.max(m, Number(k.split(',')[0])); }); return m; },
    getLastColumn: () => { let m = 0; cells.forEach((v, k) => { m = Math.max(m, Number(k.split(',')[1])); }); return m; },
    getRange: (a, b, c, d) => (typeof a === 'string' ? makeRange(p, ...parseA1(a)) : makeRange(p, a, b, c || 1, d || 1)),
    getDataRange: () => makeRange(p, 1, 1, Math.max(1, s.getLastRow()), Math.max(1, s.getLastColumn())),
    clear: () => { cells.clear(); return p; },
    clearContents: () => { cells.clear(); return p; },
    // test helper: put a 2D array at row, col
    _put: (rows, row, col) => { rows.forEach((rw, i) => rw.forEach((v, j) => s._set((row || 1) + i, (col || 1) + j, v))); return p; }
  };
  const p = new Proxy(s, { get(t, k) { return (k in t) ? t[k] : (typeof k === 'string' ? () => p : undefined); } });
  return p;
}

function makeSpreadsheet(tabNames) {
  const ss = {
    _sheets: [], _active: null,
    getId: () => "FAKE_SS",
    getSheets: () => ss._sheets.slice(),
    getSheetByName: n => ss._sheets.find(s => s.getName() === n) || null,
    getActiveSheet: () => ss._active,
    setActiveSheet: s => s.activate(),
    insertSheet: (a, b) => {
      const name = typeof a === 'string' ? a : "Sheet" + (ss._sheets.length + 1);
      const index = typeof a === 'number' ? a : (typeof b === 'number' ? b : ss._sheets.length);
      const s = makeSheet(ss, name);
      ss._sheets.splice(index, 0, s);
      ss._active = s; // like Sheets: a new tab is shown and becomes the active one
      return s;
    },
    deleteSheet: s => {
      if (ss._active === s) throw new Error("fake sheet: deleting the tab the user is on");
      ss._sheets = ss._sheets.filter(x => x !== s);
    },
    _toasts: [],
    toast: (msg, title) => { ss._toasts.push({ msg: String(msg), title: String(title || "") }); }
  };
  (tabNames || []).forEach(n => ss._sheets.push(makeSheet(ss, n)));
  ss._active = ss._sheets[0] || null;
  return ss;
}

/** SpreadsheetApp for loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) } }). ui._alerts / ui._menus record calls. */
function fakeSpreadsheetApp(ss) {
  const builder = () => new Proxy({}, { get(t, k) { return k === 'build' ? () => ({}) : () => builder(); } });
  const ui = { _alerts: [], _menus: [], ButtonSet: {}, Button: {},
    alert: (...a) => { ui._alerts.push(a.map(String).join(" | ")); }, showSidebar() {}, showModalDialog() {},
    createMenu: name => {
      const m = { name: name, items: [] };
      ui._menus.push(m);
      const p = new Proxy(m, { get(t, k) {
        if (k === 'addItem') return (label, fn) => { m.items.push(fn); return p; };
        return (k in t) ? t[k] : () => p;
      } });
      return p;
    } };
  return {
    _ui: ui,
    getActiveSpreadsheet: () => ss, flush() {}, newRichTextValue: builder, newTextStyle: builder,
    BorderStyle: { SOLID: 'SOLID' }, WrapStrategy: { CLIP: 'CLIP', WRAP: 'WRAP', OVERFLOW: 'OVERFLOW' },
    BandingTheme: new Proxy({}, { get: (t, k) => String(k) }),
    getUi: () => ui
  };
}

module.exports = { makeSpreadsheet, fakeSpreadsheetApp, parseA1 };
