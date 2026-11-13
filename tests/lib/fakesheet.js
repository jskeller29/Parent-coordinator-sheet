// A small fake spreadsheet for offline flow tests: tabs with a grid of values,
// hide / show / activate / insert / delete, and ranges that read and write values.
// Formatting calls (colors, fonts, notes, borders ...) are accepted and otherwise
// ignored, but written to ss._log ({ tab, row, col, nr, nc, fn, args }) so a test
// can check what was painted. Kept: merges (merge / getMergedRanges), fills
// (setBackground / getBackground), column widths, conditional format rules,
// data validations (sheet._dv: "r,c" -> rule) and bandings (sheet._bandings).
// Not a full SpreadsheetApp: add what a new test needs.

function colNum(letters) { return letters.toUpperCase().split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0); }
function parseA1(a1, maxRows) {
  const m = String(a1).match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+)?)?$/i);
  if (!m) throw new Error("fake sheet: can't read range " + a1);
  const r1 = Number(m[2]), c1 = colNum(m[1]);
  // "Z2:Z" = open-ended: down to the last row of the tab
  const r2 = m[4] ? Number(m[4]) : (m[3] ? (maxRows || r1) : r1), c2 = m[3] ? colNum(m[3]) : c1;
  return [r1, c1, r2 - r1 + 1, c2 - c1 + 1];
}

// Ranges and tabs are Proxies: anything not defined here is a formatting call,
// accepted and chained (returns the same object).
function colLetters(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

function makeRange(sheet, row, col, nr, nc) {
  const log = (fn, args) => { if (sheet._log) sheet._log.push({ tab: sheet.getName(), row, col, nr, nc, fn, args: Array.from(args) }); };
  const each = f => { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) f(row + i, col + j); };
  const r = {
    getRow: () => row, getColumn: () => col, getNumRows: () => nr, getNumColumns: () => nc,
    getLastRow: () => row + nr - 1, getLastColumn: () => col + nc - 1, getSheet: () => sheet,
    getA1Notation: () => colLetters(col) + row + (nr > 1 || nc > 1 ? ':' + colLetters(col + nc - 1) + (row + nr - 1) : ''),
    merge: function () { log('merge', arguments); sheet._merges.push({ r: row, c: col, nr, nc }); return p; },
    getMergedRanges: () => sheet._merges
      .filter(m => m.r < row + nr && m.r + m.nr > row && m.c < col + nc && m.c + m.nc > col)
      .map(m => makeRange(sheet, m.r, m.c, m.nr, m.nc)),
    setBackground: function (v) { log('setBackground', arguments); each((a, b) => { if (v) sheet._bg.set(a + ',' + b, v); else sheet._bg.delete(a + ',' + b); }); return p; },
    setBackgrounds: function (grid) { log('setBackgrounds', arguments); grid.forEach((rw, i) => rw.forEach((v, j) => { const k = (row + i) + ',' + (col + j); if (v) sheet._bg.set(k, v); else sheet._bg.delete(k); })); return p; },
    getBackground: () => sheet._bg.get(row + ',' + col) || '#ffffff',
    getDataValidations: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => sheet._dv.get((row + i) + ',' + (col + j)) || null)),
    getDataValidation: () => sheet._dv.get(row + ',' + col) || null,
    getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => sheet._get(row + i, col + j))),
    getDisplayValues: () => r.getValues().map(rw => rw.map(v => v instanceof Date ? (v.getMonth() + 1) + '/' + v.getDate() + '/' + v.getFullYear() : String(v))),
    getValue: () => sheet._get(row, col),
    // formulas (sheet._formulas: "r,c" -> text); writing a value or clearing a cell removes its formula, like Sheets
    getFormula: () => sheet._formulas.get(row + ',' + col) || "",
    getFormulas: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => sheet._formulas.get((row + i) + ',' + (col + j)) || "")),
    setFormula: function (f) { log('setFormula', arguments); sheet._formulas.set(row + ',' + col, String(f)); return p; },
    getDisplayValue: () => r.getDisplayValues()[0][0],
    setValues: vals => {
      if (vals.length !== nr || vals.some(v => v.length !== nc)) throw new Error("fake sheet: setValues size " + vals.length + "x" + (vals[0] || []).length + " into " + nr + "x" + nc);
      vals.forEach((rw, i) => rw.forEach((v, j) => sheet._set(row + i, col + j, v)));
      return p;
    },
    setValue: v => { sheet._set(row, col, v); return p; },
    clearContent: () => { log('clearContent', []); for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) sheet._set(row + i, col + j, ""); return p; },
    clear: () => r.clearContent(),
    setRichTextValue: v => { sheet._set(row, col, v && v.text !== undefined ? v.text : ""); return p; }
  };
  const p = new Proxy(r, { get(t, k) { return (k in t) ? t[k] : (typeof k === 'string' ? function () { log(k, arguments); return p; } : undefined); } });
  return p;
}

let nextId = 1000;
function makeSheet(ss, name) {
  const cells = new Map(); // "r,c" -> value
  const s = {
    _name: name, _id: nextId++, _hidden: false, _maxRows: 1000, _maxCols: 26,
    _log: ss._log, _merges: [], _bg: new Map(), _dv: new Map(), _widths: {}, _cf: [], _cfWrites: 0, _bandings: [],
    getColumnWidth: c => s._widths[c] || 100,
    setColumnWidth: (c, w) => { s._widths[c] = w; s._logSheet('setColumnWidth', [c, w]); return p; },
    getConditionalFormatRules: () => s._cf.slice(),
    setConditionalFormatRules: rules => { s._cf = rules.slice(); s._cfWrites++; return p; },
    getBandings: () => s._bandings.slice(),
    _logSheet: (fn, args) => { if (s._log) s._log.push({ tab: s._name, sheetCall: true, fn, args: Array.from(args) }); },
    _get: (r, c) => (cells.has(r + ',' + c) ? cells.get(r + ',' + c) : ""),
    _set: (r, c, v) => { s._formulas.delete(r + ',' + c); if (v === "" || v === null || v === undefined) cells.delete(r + ',' + c); else cells.set(r + ',' + c, v); },
    _cells: cells, _formulas: new Map(),
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
    deleteRows: (a, n) => { s._maxRows -= n; return p; },
    deleteColumns: (a, n) => { // the cells to the right move left, like Sheets
      const moved = new Map();
      cells.forEach((v, k) => { const [rr, cc] = k.split(',').map(Number); if (cc < a) moved.set(k, v); else if (cc >= a + n) moved.set(rr + ',' + (cc - n), v); });
      cells.clear(); moved.forEach((v, k) => cells.set(k, v));
      s._maxCols -= n; s._logSheet('deleteColumns', [a, n]); return p;
    },
    getLastRow: () => { let m = 0; cells.forEach((v, k) => { m = Math.max(m, Number(k.split(',')[0])); }); return m; },
    getLastColumn: () => { let m = 0; cells.forEach((v, k) => { m = Math.max(m, Number(k.split(',')[1])); }); return m; },
    getRange: (a, b, c, d) => (typeof a === 'string' ? makeRange(p, ...parseA1(a, s._maxRows)) : makeRange(p, a, b, c || 1, d || 1)),
    getDataRange: () => makeRange(p, 1, 1, Math.max(1, s.getLastRow()), Math.max(1, s.getLastColumn())),
    clear: () => { cells.clear(); s._formulas.clear(); return p; },
    clearContents: () => { cells.clear(); s._formulas.clear(); return p; },
    // test helper: put a 2D array at row, col
    _put: (rows, row, col) => { rows.forEach((rw, i) => rw.forEach((v, j) => s._set((row || 1) + i, (col || 1) + j, v))); return p; }
  };
  const p = new Proxy(s, { get(t, k) { return (k in t) ? t[k] : (typeof k === 'string' ? function () { s._logSheet(k, arguments); return p; } : undefined); } });
  return p;
}

function makeSpreadsheet(tabNames) {
  const ss = {
    _sheets: [], _active: null, _log: [],
    getId: () => "FAKE_SS",
    getUrl: () => "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/edit",
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
    _meta: [],
    addDeveloperMetadata: (key, value) => { ss._meta.push({ key, value }); return ss; },
    createDeveloperMetadataFinder: () => {
      let key = null;
      const f = { withKey: k => { key = k; return f; },
        find: () => ss._meta.filter(m => key === null || m.key === key).map(m => ({ getKey: () => m.key, getValue: () => m.value,
          remove: () => { ss._meta = ss._meta.filter(x => x !== m); } })) };
      return f;
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
  // Conditional format rules that can be read back (whenTextEqualTo / whenFormulaSatisfied, colors, ranges)
  const newConditionalFormatRule = () => {
    const st = {};
    const b = {
      whenTextEqualTo: t => { st.text = t; return b; }, whenFormulaSatisfied: f => { st.formula = f; return b; },
      setFontColor: c => { st.font = c; return b; }, setBackground: c => { st.bg = c; return b; },
      setRanges: rs => { st.ranges = rs; return b; },
      build: () => ({ _st: st, getRanges: () => st.ranges || [],
        getBooleanCondition: () => ({
          getCriteriaType: () => (st.text !== undefined ? 'TEXT_EQUAL_TO' : 'CUSTOM_FORMULA'),
          getCriteriaValues: () => [st.text !== undefined ? st.text : st.formula],
          getFontColor: () => st.font || null, getBackground: () => st.bg || null }) })
    };
    return b;
  };
  return {
    _ui: ui,
    getActiveSpreadsheet: () => ss, flush() {}, newRichTextValue: builder, newTextStyle: builder, newDataValidation: builder,
    newConditionalFormatRule,
    BorderStyle: { SOLID: 'SOLID', SOLID_MEDIUM: 'SOLID_MEDIUM', SOLID_THICK: 'SOLID_THICK' },
    WrapStrategy: { CLIP: 'CLIP', WRAP: 'WRAP', OVERFLOW: 'OVERFLOW' },
    BandingTheme: new Proxy({}, { get: (t, k) => String(k) }),
    BooleanCriteria: { TEXT_EQUAL_TO: 'TEXT_EQUAL_TO', CUSTOM_FORMULA: 'CUSTOM_FORMULA' },
    DataValidationCriteria: { VALUE_IN_RANGE: 'VALUE_IN_RANGE', VALUE_IN_LIST: 'VALUE_IN_LIST' },
    getUi: () => ui
  };
}

module.exports = { makeSpreadsheet, fakeSpreadsheetApp, parseA1 };
