// ======================================================================
// FILE: Sheet Lock.gs
// PURPOSE: "Baby-proofing" so nobody breaks the sheet by accident.
//
//  1. LOCKED AREAS (every tab except Raw Data and Backend):
//     - Row 1 (headers) of Contact Log, Notes and Send Out.
//     - Contact Log: column A (the "Last Sunday" formula), the END bar row
//       (except its O "FILTER?" checkbox) and everything below it.
//     - Events: the header row (row 2) of every event table. Row 1 (event
//       name / type / place / date / delete box) stays editable.
//     - PCAR: everything except H2:I2 ("PUT YOUR OWN INPUT" dates).
//     - Tabs the build writes (Master Table, Combined Contact Tracking,
//       Parents Divided, Directory, Phone Contacts, Students Override Report,
//       New Students): the whole tab.
//     Everywhere the fake data loads (Contact Log rows above END, Notes /
//     Send Out rows 2+, Events tables) stays editable.
//
//     Two layers:
//     a) Google "warning" protections (Data > Protect sheets and ranges):
//        Google asks "are you sure?" BEFORE an edit, a row delete, a sort...
//        Scripts are never blocked by them. All of ours carry a description
//        starting with SHEET_LOCK_TAG, so they can be found and redone.
//     b) If someone clicks through that warning anyway, installedOnEdit puts
//        the locked cells back (like the "Initial Setup first" lock in
//        onEdit). Row 1 / END bar / PCAR come back from a snapshot kept in
//        document properties; built tabs undo a one-cell edit (the next
//        build rewrites them anyway).
//
//  2. MULTI-ROW CLEAR CHECK (Contact Log, Notes, Send Out, Events): clearing
//     2+ rows at once (Delete key) asks "did you mean to do that?". No puts
//     the rows back from the copy onSelectionChange took when the rows were
//     selected; without that copy it tells the user to press Ctrl+Z.
//
// Everything is switched on by default. App Menu > "🔓 Unlock Sheet" turns
// it off (doc property SHEET_LOCK_OFF) for template edits; "🔒 Lock Sheet"
// turns it back on and retakes the snapshots.
// Shared code: works the same with or without z_MINE.gs. Owner-only tabs
// (Old Master Table, Old Phone Contacts) are not locked.
// ======================================================================

var SHEET_LOCK_TAG = "🔒 Sheet Lock";
var SHEET_LOCK_OFF_KEY = "SHEET_LOCK_OFF";
var SHEET_LOCK_SNAP_KEY = "SHEET_LOCK_SNAP_";     // + zone name; chunks "_0", "_1"... and count "_n"
var SHEET_LOCK_CL_END_KEY = "SHEET_LOCK_CL_END";  // END row the Contact Log snapshot was taken at
var SHEET_LOCK_SEL_KEY = "SHEET_LOCK_SEL_";       // + sheet id (document cache)
var SHEET_LOCK_SNAP_CHUNK = 2500;                 // chars per property (stays under the 9 KB value limit)

// Tabs the build writes -> how many header rows they have
var SHEET_LOCK_BUILT_TABS = {
  "Master Table": 4,
  "Combined Contact Tracking": 1,
  "Parents Divided": 1,
  "Directory": 1,
  "Phone Contacts": 1,
  "Students Override Report": 1,
  "New Students": 1
};
// Typed tabs whose row 1 is locked (Contact Log also gets its END bar)
var SHEET_LOCK_HEADER_TABS = ["Contact Log", "Notes", "Send Out"];
var SHEET_LOCK_PCAR = "PCAR";
var SHEET_LOCK_PCAR_INPUT = { row: 2, col: 8, numCols: 2 }; // H2:I2

// Row 2 of every event table (tables are 7 columns: 6 data + 1 spacer)
var EVENT_TABLE_HEADERS = ["OSIS", "Guardian(s)", "Parent Attendee", "Student Name", "# of Indv", "Notes"];
var EVENT_TABLE_WIDTH = 7;

// Tabs with the "clear 2+ rows?" check -> first data row
var CLEAR_CHECK_FIRST_ROW = { "Contact Log": 2, "Notes": 2, "Send Out": 2, "Events": 3 };

// Functions that make a cell a spill (array) anchor; their outputs have no formula of their own
var SPILL_FORMULA_RE_ = /\b(HSTACK|VSTACK|ARRAYFORMULA|SPLIT|FILTER|SORTN?|UNIQUE|SEQUENCE|QUERY|TRANSPOSE|MAP|BYROW|BYCOL|MAKEARRAY|IMPORTRANGE|FLATTEN|TOCOL|TOROW|CHOOSECOLS|CHOOSEROWS|WRAPROWS|WRAPCOLS|EXPAND)\s*\(/i;

function isSheetLockOn_() {
  return PropertiesService.getDocumentProperties().getProperty(SHEET_LOCK_OFF_KEY) !== "true";
}

function isLockedTab_(name) {
  return SHEET_LOCK_BUILT_TABS.hasOwnProperty(name) || SHEET_LOCK_HEADER_TABS.indexOf(name) !== -1 ||
    name === SHEET_LOCK_PCAR || name === "Events";
}

// ======================================================================
// MENU: lock / unlock
// ======================================================================
function toggleSheetLock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();

  if (isSheetLockOn_()) {
    const resp = ui.alert(
      "🔓 Unlock the sheet?",
      "This turns off the sheet lock so headers, the Contact Log END bar, PCAR and the built tabs can be changed, " +
      "and clearing several rows no longer asks first.\n\nUse it only to change the layout, then come back here and click '🔒 Lock Sheet'.",
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) return;
    props.setProperty(SHEET_LOCK_OFF_KEY, "true");
    removeSheetLocks_(ss);
    ss.toast("The sheet is unlocked. Lock it again from the App Menu when you're done.", "🔓 Unlocked", 6);
  } else {
    props.deleteProperty(SHEET_LOCK_OFF_KEY);
    applySheetLocks_(ss);
    ss.toast("Headers, the END bar, PCAR and the built tabs are protected again.", "🔒 Locked", 6);
  }
  onOpen(); // refresh the menu label
}

// ======================================================================
// PROTECTIONS + SNAPSHOTS
// ======================================================================

/**
 * Rectangles (1-based, inclusive) that are locked on one tab. maxR / maxC
 * default to the sheet size; the edit guard passes the edit's own extent
 * (same intersections, two fewer calls per edit).
 */
function sheetLockRects_(sheet, name, endRow, maxR, maxC) {
  maxR = maxR || sheet.getMaxRows();
  maxC = maxC || sheet.getMaxColumns();
  const rects = [];
  const add = (r1, c1, r2, c2, zone) => { if (r2 >= r1 && c2 >= c1) rects.push({ r1, c1, r2, c2, zone }); };

  if (SHEET_LOCK_BUILT_TABS.hasOwnProperty(name)) {
    const hdr = SHEET_LOCK_BUILT_TABS[name];
    add(1, 1, hdr, maxC, "HDR:" + name);
    add(hdr + 1, 1, maxR, maxC, "BODY");
  } else if (SHEET_LOCK_HEADER_TABS.indexOf(name) !== -1) {
    add(1, 1, 1, maxC, "HDR:" + name);
    if (name === "Contact Log" && endRow) {
      add(2, 1, endRow - 1, 1, "CLWEEK"); // column A "Last Sunday" formula of every log row
      add(endRow, 1, endRow, CL_COL_FILTER_TOGGLE - 1, "CLEND");
      add(endRow, CL_COL_FILTER_TOGGLE + 1, endRow, maxC, "CLEND");
      add(endRow + 1, 1, maxR, maxC, "CLEND");
    }
  } else if (name === SHEET_LOCK_PCAR) {
    const inp = SHEET_LOCK_PCAR_INPUT;
    add(1, 1, inp.row - 1, maxC, "PCAR");
    add(inp.row, 1, inp.row, inp.col - 1, "PCAR");
    add(inp.row, inp.col + inp.numCols, inp.row, maxC, "PCAR");
    add(inp.row + 1, 1, maxR, maxC, "PCAR");
  }
  return rects;
}

/** Start columns of the event tables that exist (row 1 has an event name). */
function eventTableStarts_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return [];
  const row1 = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const starts = [];
  for (let c = 1; c <= lastCol; c += EVENT_TABLE_WIDTH) {
    if (toText_(row1[c - 1])) starts.push(c);
  }
  return starts;
}

/** The warning protections one tab should have. */
function desiredSheetLocks_(sheet, name) {
  const out = [];
  const desc = SHEET_LOCK_TAG + " – ";
  if (SHEET_LOCK_BUILT_TABS.hasOwnProperty(name)) {
    out.push({ kind: "SHEET", unprotected: [], desc: desc + "built automatically (🏗️ Build Sheets Only rewrites it)" });
  } else if (name === SHEET_LOCK_PCAR) {
    const inp = SHEET_LOCK_PCAR_INPUT;
    out.push({ kind: "SHEET", unprotected: [sheet.getRange(inp.row, inp.col, 1, inp.numCols)], desc: desc + "PCAR formulas (only H2:I2 are yours to change)" });
  } else if (name === "Events") {
    eventTableStarts_(sheet).forEach(c => {
      out.push({ kind: "RANGE", range: sheet.getRange(2, c, 1, EVENT_TABLE_HEADERS.length), desc: desc + "event table headers" });
    });
  } else if (SHEET_LOCK_HEADER_TABS.indexOf(name) !== -1) {
    sheetLockRects_(sheet, name, name === "Contact Log" ? findEndRowInColumnE_(sheet) : null).forEach(r => {
      // Column A is protected down to the bottom of the sheet (it overlaps the END bar
      // lock), so the rows added above END land inside it and it grows with them
      const r2 = r.zone === "CLWEEK" ? sheet.getMaxRows() : r.r2;
      out.push({
        kind: "RANGE",
        range: sheet.getRange(r.r1, r.c1, r2 - r.r1 + 1, r.c2 - r.c1 + 1),
        desc: desc + (r.zone === "CLEND" ? "Contact Log END bar and below" : r.zone === "CLWEEK" ? "Contact Log column A (Last Sunday formula)" : "header row")
      });
    });
  }
  return out;
}

function sheetLockKey_(sheetName, kind, a1, unprotectedA1s) {
  return [sheetName, kind, a1, (unprotectedA1s || []).join(",")].join("|");
}

/**
 * Adds / refreshes the warning protections and the snapshots. Safe to run any
 * time (only our own tagged protections are touched; unchanged ones are kept).
 * names: optional list of tab names (default: every locked tab).
 */
function applySheetLocks_(ss, names) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (!isSheetLockOn_()) return;
  const sheets = ss.getSheets().filter(sh => {
    const n = sh.getName().trim();
    return isLockedTab_(n) && (!names || names.indexOf(n) !== -1);
  });
  if (!sheets.length) return;

  // What should exist
  const want = new Map();
  sheets.forEach(sh => {
    const n = sh.getName().trim();
    desiredSheetLocks_(sh, n).forEach(d => {
      const a1 = d.kind === "SHEET" ? "SHEET" : d.range.getA1Notation();
      want.set(sheetLockKey_(sh.getName(), d.kind, a1, d.unprotected && d.unprotected.map(r => r.getA1Notation())), Object.assign({ sheet: sh }, d));
    });
  });

  // What exists: keep matches, remove stale ones (only ours)
  const sheetIds = new Set(sheets.map(sh => sh.getSheetId()));
  const otherSheetProtection = new Set(); // tabs that already have someone else's sheet protection
  [SpreadsheetApp.ProtectionType.SHEET, SpreadsheetApp.ProtectionType.RANGE].forEach(type => {
    ss.getProtections(type).forEach(p => {
      let key = null, sheetId = null;
      try {
        const r = p.getRange();
        sheetId = r.getSheet().getSheetId();
        if (!sheetIds.has(sheetId)) return;
        if (type === SpreadsheetApp.ProtectionType.SHEET) {
          key = sheetLockKey_(r.getSheet().getName(), "SHEET", "SHEET", p.getUnprotectedRanges().map(u => u.getA1Notation()));
        } else {
          key = sheetLockKey_(r.getSheet().getName(), "RANGE", r.getA1Notation());
        }
      } catch (err) { key = null; }
      const ours = String(p.getDescription() || "").indexOf(SHEET_LOCK_TAG) === 0;
      if (!ours) {
        if (type === SpreadsheetApp.ProtectionType.SHEET && sheetId !== null) otherSheetProtection.add(sheetId);
        return;
      }
      if (key && want.has(key) && !want.get(key).kept) { want.get(key).kept = true; return; }
      try { p.remove(); } catch (err) { console.warn("Sheet lock: could not remove an old protection: " + err.message); }
    });
  });

  want.forEach(d => {
    if (d.kept) return;
    try {
      if (d.kind === "SHEET") {
        if (otherSheetProtection.has(d.sheet.getSheetId())) return; // never take over the owner's own protection
        const p = d.sheet.protect().setDescription(d.desc);
        p.setWarningOnly(true);
        if (d.unprotected.length) p.setUnprotectedRanges(d.unprotected);
      } else {
        d.range.protect().setDescription(d.desc).setWarningOnly(true);
      }
    } catch (err) { console.warn("Sheet lock: could not protect " + d.sheet.getName() + ": " + err.message); }
  });

  // Snapshots of the locked cells (used to put them back after an edit)
  sheets.forEach(sh => {
    try { takeSheetLockSnapshots_(sh, sh.getName().trim()); }
    catch (err) { console.warn("Sheet lock: snapshot of " + sh.getName() + " skipped: " + err.message); }
  });
}

/** Removes every protection this file added (unlock). Snapshots are kept. */
function removeSheetLocks_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  [SpreadsheetApp.ProtectionType.SHEET, SpreadsheetApp.ProtectionType.RANGE].forEach(type => {
    ss.getProtections(type).forEach(p => {
      if (String(p.getDescription() || "").indexOf(SHEET_LOCK_TAG) !== 0) return;
      try { p.remove(); } catch (err) { console.warn("Sheet lock: could not remove a protection: " + err.message); }
    });
  });
}

/** Re-takes the snapshots of one tab's locked areas (headers, END bar + below, PCAR). */
function takeSheetLockSnapshots_(sheet, name) {
  const lastR = Math.max(sheet.getLastRow(), 1);
  const lastC = Math.max(sheet.getLastColumn(), 1);
  if (SHEET_LOCK_BUILT_TABS.hasOwnProperty(name)) {
    saveLockSnapshot_("HDR:" + name, readLockSnapshot_(sheet, 1, 1, SHEET_LOCK_BUILT_TABS[name], lastC));
  } else if (SHEET_LOCK_HEADER_TABS.indexOf(name) !== -1) {
    saveLockSnapshot_("HDR:" + name, readLockSnapshot_(sheet, 1, 1, 1, lastC));
    if (name === "Contact Log") snapshotContactLogEnd_(sheet, findEndRowInColumnE_(sheet));
  } else if (name === SHEET_LOCK_PCAR) {
    saveLockSnapshot_("PCAR", readLockSnapshot_(sheet, 1, 1, lastR, lastC));
  }
}

/** Snapshot of the Contact Log END bar and everything below it (tied to that END row). */
function snapshotContactLogEnd_(sheet, endRow) {
  if (!endRow) return;
  const lastR = Math.max(sheet.getLastRow(), endRow);
  const snap = readLockSnapshot_(sheet, endRow, 1, lastR, Math.max(sheet.getLastColumn(), CL_WIDTH));
  snap.endRow = endRow;
  saveLockSnapshot_("CLEND", snap);
  PropertiesService.getDocumentProperties().setProperty(SHEET_LOCK_CL_END_KEY, String(endRow));
}

/** Called after the Contact Log END bar moved (a row was added above it). */
function refreshContactLogEndSnapshot_(sheet, endRow) {
  try {
    if (isSheetLockOn_()) snapshotContactLogEnd_(sheet, endRow);
  } catch (err) { console.warn("Sheet lock: END bar snapshot skipped: " + err.message); }
}

/**
 * Reads a block into { cells: {"row,col": [kind, value]}, hasFormulas }.
 * kind: "f" formula, "v" value, "d" date (ms), "s" probably a spill output of a
 * formula to its left / above (has no formula of its own; restoring it means
 * clearing it so the formula can spill again). Blank cells are left out.
 */
function readLockSnapshot_(sheet, r1, c1, r2, c2) {
  const nr = r2 - r1 + 1, nc = c2 - c1 + 1;
  const range = sheet.getRange(r1, c1, nr, nc);
  const values = range.getValues();
  const formulas = range.getFormulas();
  const cells = {};
  let hasFormulas = false;
  for (let i = 0; i < nr; i++) {
    for (let j = 0; j < nc; j++) {
      const key = (r1 + i) + "," + (c1 + j);
      if (formulas[i][j]) { cells[key] = ["f", formulas[i][j]]; hasFormulas = true; continue; }
      const v = values[i][j];
      if (v === "" || v === null) continue;
      if (nearestFormulaSpills_(formulas, i, j)) { cells[key] = ["s"]; continue; }
      cells[key] = v instanceof Date ? ["d", v.getTime()] : ["v", v];
    }
  }
  return { r1, c1, r2, c2, hasFormulas, cells };
}

function nearestFormulaSpills_(formulas, i, j) {
  for (let x = j - 1; x >= 0; x--) {
    if (formulas[i][x]) { if (SPILL_FORMULA_RE_.test(formulas[i][x])) return true; break; }
  }
  for (let y = i - 1; y >= 0; y--) {
    if (formulas[y][j]) return SPILL_FORMULA_RE_.test(formulas[y][j]);
  }
  return false;
}

function saveLockSnapshot_(zone, snap) {
  const props = PropertiesService.getDocumentProperties();
  const key = SHEET_LOCK_SNAP_KEY + zone;
  const json = JSON.stringify(snap);
  const oldCount = Number(props.getProperty(key + "_n")) || 0;
  const out = {};
  let n = 0;
  for (let i = 0; i < json.length;) {
    let end = Math.min(i + SHEET_LOCK_SNAP_CHUNK, json.length);
    const code = json.charCodeAt(end - 1);
    if (end < json.length && code >= 0xD800 && code <= 0xDBFF) end--; // never split an emoji / math-bold pair
    out[key + "_" + (n++)] = json.slice(i, end);
    i = end;
  }
  out[key + "_n"] = String(n);
  props.setProperties(out);
  for (let i = n; i < oldCount; i++) props.deleteProperty(key + "_" + i);
}

function loadLockSnapshot_(zone) {
  try {
    const props = PropertiesService.getDocumentProperties();
    const key = SHEET_LOCK_SNAP_KEY + zone;
    const n = Number(props.getProperty(key + "_n")) || 0;
    if (!n) return null;
    let json = "";
    for (let i = 0; i < n; i++) {
      const part = props.getProperty(key + "_" + i);
      if (part === null) return null;
      json += part;
    }
    return JSON.parse(json);
  } catch (err) {
    console.warn("Sheet lock: snapshot " + zone + " unreadable: " + err.message);
    return null;
  }
}

/** Writes one locked rectangle back from its snapshot. */
function restoreRectFromSnapshot_(sheet, h, snap) {
  const nr = h.r2 - h.r1 + 1, nc = h.c2 - h.c1 + 1;
  const range = sheet.getRange(h.r1, h.c1, nr, nc);
  const grid = [];
  const consts = [];
  for (let i = 0; i < nr; i++) {
    const row = [];
    for (let j = 0; j < nc; j++) {
      const cell = snap.cells[(h.r1 + i) + "," + (h.c1 + j)];
      if (!cell || cell[0] === "s") { row.push(""); continue; }
      if (cell[0] === "f") { row.push(cell[1]); continue; }
      const v = cell[0] === "d" ? new Date(cell[1]) : cell[1];
      consts.push([i, j, v]);
      row.push(v);
    }
    grid.push(row);
  }

  if (!snap.hasFormulas) { range.setValues(grid); return; }

  // Formula areas (PCAR, below the END bar): formulas first, flush, then only
  // fill constants into cells that are still empty, so a value never blocks a
  // formula that spills into that cell.
  consts.forEach(c => { grid[c[0]][c[1]] = ""; });
  range.setValues(grid);
  if (!consts.length) return;
  SpreadsheetApp.flush();
  const now = range.getValues();
  consts.forEach(c => {
    if (now[c[0]][c[1]] === "") sheet.getRange(h.r1 + c[0], h.c1 + c[1]).setValue(c[2]);
  });
}

// ======================================================================
// EDIT GUARD (called first thing by handleEdit_ for every tab but Raw Data)
// ======================================================================

/**
 * Puts locked cells back and runs the multi-row clear check.
 * Returns { stop, endRow }: stop = nothing else should run for this edit;
 * endRow = the Contact Log END row (so handleEdit_ need not look it up again).
 */
function sheetLockGuard_(e, sheet, name) {
  const out = { stop: false };
  if (!isLockedTab_(name)) return out;
  try {
    if (!isSheetLockOn_()) return out;
    const range = e.range;
    const er = { r1: range.getRow(), c1: range.getColumn(), r2: range.getLastRow(), c2: range.getLastColumn() };

    if (name === "Contact Log") out.endRow = contactLogEndRowForGuard_(sheet, er);

    let lockedCells = 0;
    const problems = { restored: {}, failed: {} };

    if (name === "Events") {
      lockedCells = restoreEventHeaders_(sheet, er, problems);
    } else {
      const hits = [];
      sheetLockRects_(sheet, name, out.endRow, er.r2, er.c2).forEach(r => {
        const h = { r1: Math.max(r.r1, er.r1), c1: Math.max(r.c1, er.c1), r2: Math.min(r.r2, er.r2), c2: Math.min(r.c2, er.c2), zone: r.zone };
        if (h.r1 <= h.r2 && h.c1 <= h.c2) hits.push(h);
      });
      hits.forEach(h => {
        lockedCells += (h.r2 - h.r1 + 1) * (h.c2 - h.c1 + 1);
        const kind = h.zone === "BODY" ? "built" : h.zone === "CLEND" ? "end" : h.zone === "CLWEEK" ? "week" : h.zone === "PCAR" ? "pcar" : (SHEET_LOCK_BUILT_TABS.hasOwnProperty(name) ? "built" : "header");
        // Each area on its own: one failing never skips the others or the clear check
        try {
          const ok = restoreLockedHit_(e, sheet, h, out.endRow);
          // Column A is only mentioned when the edit was in column A alone; when whole
          // rows were cleared, a "change undone" toast would read as if the rows came back
          if (kind === "week" && (er.c1 !== 1 || er.c2 !== 1)) return;
          if (ok) problems.restored[kind] = true;
          else problems.failed[kind] = true;
        } catch (err) {
          console.error("Sheet lock: could not put back " + h.zone + ": " + err.message);
          problems.failed[kind] = true;
        }
      });
      try {
        if (name === "Contact Log" && !hits.some(h => h.zone === "CLEND")) keepContactLogEndSnapshotCurrent_(sheet, out.endRow);
      } catch (err) { console.warn("Sheet lock: END bar snapshot check failed: " + err.message); }
    }

    toastSheetLock_(e.source, problems);
    if (lockedCells) {
      const total = (er.r2 - er.r1 + 1) * (er.c2 - er.c1 + 1);
      if (lockedCells >= total) { out.stop = true; return out; }
    }

    if (CLEAR_CHECK_FIRST_ROW[name] && !confirmMultiRowClear_(e, sheet, name, er, out.endRow)) out.stop = true;
  } catch (err) {
    console.error("Sheet lock guard failed: " + err.message);
  }
  return out;
}

/** Contact Log END row, also when the edit just typed over the "END" marker itself. */
function contactLogEndRowForGuard_(sheet, er) {
  const found = findEndRowInColumnE_(sheet);
  if (found) return found;
  const saved = Number(PropertiesService.getDocumentProperties().getProperty(SHEET_LOCK_CL_END_KEY)) || null;
  if (saved && er.r1 <= saved && er.r2 >= saved && er.c1 <= CL_COL_OSIS && er.c2 >= CL_COL_OSIS) return saved;
  return null;
}

/** The END bar moved without passing handleEdit_ (rows inserted / deleted by hand): retake its snapshot. */
function keepContactLogEndSnapshotCurrent_(sheet, endRow) {
  if (!endRow) return;
  const saved = PropertiesService.getDocumentProperties().getProperty(SHEET_LOCK_CL_END_KEY);
  if (saved && Number(saved) !== endRow) refreshContactLogEndSnapshot_(sheet, endRow);
}

/** Puts one locked rectangle back. Returns false when that was not possible. */
function restoreLockedHit_(e, sheet, h, endRow) {
  if (h.zone === "CLWEEK") {
    // Always the same formula, so no snapshot is needed
    const formulas = [];
    for (let r = h.r1; r <= h.r2; r++) formulas.push([weekFormula_(r)]);
    sheet.getRange(h.r1, 1, formulas.length, 1).setFormulas(formulas);
    return true;
  }
  const snap = h.zone === "BODY" ? null : loadLockSnapshot_(h.zone);
  if (snap && (h.zone !== "CLEND" || snap.endRow === endRow)) {
    restoreRectFromSnapshot_(sheet, h, snap);
    return true;
  }
  // No usable snapshot: a one-cell edit of plain values (built tabs, headers) can use oldValue
  const single = e.range.getNumRows() === 1 && e.range.getNumColumns() === 1;
  if (single && (h.zone === "BODY" || h.zone.indexOf("HDR:") === 0)) {
    if (e.oldValue !== undefined) { e.range.setValue(e.oldValue); return true; }
    if (e.value !== undefined) { e.range.clearContent(); return true; } // typed into an empty cell
  }
  return false;
}

/** Events: row 2 of each event table always reads EVENT_TABLE_HEADERS. Returns the locked cell count. */
function restoreEventHeaders_(sheet, er, problems) {
  if (er.r1 > 2 || er.r2 < 2) return 0;
  const firstStart = Math.floor((er.c1 - 1) / EVENT_TABLE_WIDTH) * EVENT_TABLE_WIDTH + 1;
  const row1 = sheet.getRange(1, firstStart, 1, er.c2 - firstStart + 1).getValues()[0];
  const row2Range = sheet.getRange(2, er.c1, 1, er.c2 - er.c1 + 1);
  const row2 = row2Range.getValues()[0];
  let locked = 0, changed = false;
  for (let c = er.c1; c <= er.c2; c++) {
    const start = Math.floor((c - 1) / EVENT_TABLE_WIDTH) * EVENT_TABLE_WIDTH + 1;
    const offset = c - start;
    if (offset >= EVENT_TABLE_HEADERS.length || !toText_(row1[start - firstStart])) continue;
    locked++;
    if (row2[c - er.c1] !== EVENT_TABLE_HEADERS[offset]) { row2[c - er.c1] = EVENT_TABLE_HEADERS[offset]; changed = true; }
  }
  if (changed) row2Range.setValues([row2]);
  if (locked) problems.restored.header = true;
  // Only row 2 counts as locked here; rows 1 and 3+ of the same edit carry on as normal
  return er.r1 === 2 && er.r2 === 2 ? locked : 0;
}

function toastSheetLock_(ss, problems) {
  const what = {
    header: "The header row can't be changed",
    end: "The Contact Log END bar and the area below it can't be changed",
    week: "Column A (Last Sunday) fills itself in from the Date and can't be changed",
    pcar: "The PCAR formulas can't be changed (only H2:I2 are yours)",
    built: "This tab is built automatically and can't be edited here"
  };
  const restored = Object.keys(problems.restored);
  const failed = Object.keys(problems.failed);
  if (failed.length) {
    const msg = what[failed[0]] + ". Press Ctrl+Z (⌘+Z on a Mac) to undo" +
      (failed.indexOf("built") !== -1 ? ", or click 🏗️ Build Sheets Only to rebuild it." : ".");
    ss.toast(msg, "🔒 Locked", 8);
  } else if (restored.length) {
    ss.toast(what[restored[0]] + ", so your change was undone.", "🔒 Locked", 5);
  }
}

// ======================================================================
// MULTI-ROW CLEAR CHECK
// ======================================================================

/** Simple trigger: remembers what was in a multi-row selection so a clear can be put back. */
function onSelectionChange(e) {
  try { rememberSelection_(e); } catch (err) { console.warn("Selection snapshot skipped: " + err.message); }
}

function rememberSelection_(e) {
  const range = e && e.range;
  if (!range || range.getNumRows() < 2) return;
  const sheet = range.getSheet();
  const firstRow = CLEAR_CHECK_FIRST_ROW[sheet.getName().trim()];
  if (!firstRow) return;

  const sel = { r1: range.getRow(), c1: range.getColumn(), r2: range.getLastRow(), c2: range.getLastColumn() };
  const r1 = Math.max(sel.r1, firstRow);
  const r2 = Math.min(sel.r2, sheet.getLastRow());
  const c2 = Math.min(sel.c2, sheet.getLastColumn());
  const cache = CacheService.getDocumentCache();
  const key = SHEET_LOCK_SEL_KEY + sheet.getSheetId();
  if (r2 - r1 + 1 < 2 || c2 < sel.c1 || (r2 - r1 + 1) * (c2 - sel.c1 + 1) > 20000) { cache.remove(key); return; }

  const block = sheet.getRange(r1, sel.c1, r2 - r1 + 1, c2 - sel.c1 + 1);
  const values = block.getValues();
  const formulas = block.getFormulas();
  const grid = values.map((row, i) => row.map((v, j) =>
    formulas[i][j] ? formulas[i][j] : (v instanceof Date ? { d: v.getTime() } : v)));
  const json = JSON.stringify({ sel, r1, c1: sel.c1, grid });
  if (json.length > 95000) { cache.remove(key); return; }
  cache.put(key, json, 1800);
}

/** Old contents of the rectangle from the selection snapshot, or null when there is none covering it. */
function selectionSnapshotFor_(cache, sheet, r1, c1, r2, c2) {
  const json = cache.get(SHEET_LOCK_SEL_KEY + sheet.getSheetId());
  if (!json) return null;
  const s = JSON.parse(json);
  if (r1 < s.sel.r1 || r2 > s.sel.r2 || c1 < s.sel.c1 || c2 > s.sel.c2) return null;
  const out = [];
  for (let r = r1; r <= r2; r++) {
    const row = [];
    for (let c = c1; c <= c2; c++) {
      const src = s.grid[r - s.r1];
      let v = src && (c - s.c1) < src.length ? src[c - s.c1] : "";
      if (v && typeof v === "object" && v.d !== undefined) v = new Date(v.d);
      row.push(v === null || v === undefined ? "" : v);
    }
    out.push(row);
  }
  return out;
}

/**
 * Clearing 2+ rows at once asks first. Returns false when the user said No
 * (the rows were put back, or they were told to press Ctrl+Z).
 */
function confirmMultiRowClear_(e, sheet, name, er, endRow) {
  const cache = CacheService.getDocumentCache();
  const selKey = SHEET_LOCK_SEL_KEY + sheet.getSheetId();
  const r1 = Math.max(er.r1, CLEAR_CHECK_FIRST_ROW[name]);
  let r2 = er.r2;
  if (name === "Contact Log") r2 = endRow ? Math.min(r2, endRow - 1) : 0;
  // Any other edit makes the remembered selection out of date
  if (r2 - r1 + 1 < 2 || e.value !== undefined) { cache.remove(selKey); return true; }
  const c2 = er.c2; // not the sheet's last column: the clear may just have emptied the right-hand columns

  // A clear leaves only blanks (and unticked checkboxes). Contact Log column A is
  // skipped: the sheet lock has just put its formula back.
  const skipCol = name === "Contact Log" ? 1 : 0;
  const isEmpty = (v, j) => v === "" || v === false || er.c1 + j === skipCol;
  const now = sheet.getRange(r1, er.c1, r2 - r1 + 1, c2 - er.c1 + 1).getValues();
  let sawBlank = false;
  for (let i = 0; i < now.length; i++) {
    for (let j = 0; j < now[i].length; j++) {
      const v = now[i][j];
      if (er.c1 + j === skipCol) continue;
      if (v === "") sawBlank = true;
      else if (v !== false) { cache.remove(selKey); return true; }
    }
  }
  if (!sawBlank) return true;

  let before = selectionSnapshotFor_(cache, sheet, r1, er.c1, r2, c2);
  cache.remove(selKey);
  // An all-empty copy can't be trusted: when Delete is pressed right after selecting,
  // onSelectionChange may only read the rows AFTER they were cleared. Ask anyway
  // (and fall back to Ctrl+Z) instead of guessing they were empty already.
  if (before && before.every(row => row.every(isEmpty))) before = null;

  const ui = SpreadsheetApp.getUi();
  const rowCount = r2 - r1 + 1;
  const resp = ui.alert(
    "🗑️ Clear " + rowCount + " rows?",
    'Rows ' + r1 + '–' + r2 + ' on "' + name + '" were just cleared.\n\nDid you mean to do that?\n\n' +
    "Yes = keep them cleared\nNo = " + (before ? "put them back" : "stop here (then press Ctrl+Z / ⌘+Z to bring them back)"),
    ui.ButtonSet.YES_NO
  );
  if (resp === ui.Button.YES) return true;

  if (before) {
    sheet.getRange(r1, er.c1, before.length, before[0].length).setValues(before);
    e.source.toast("Rows " + r1 + "–" + r2 + " were put back.", "↩️ Restored", 5);
  } else {
    ui.alert("↩️ Undo", "Press Ctrl+Z (⌘+Z on a Mac) now to bring the rows back.", ui.ButtonSet.OK);
  }
  return false;
}
