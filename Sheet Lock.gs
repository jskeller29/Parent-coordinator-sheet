// ======================================================================
// FILE: Sheet Lock.gs
// PURPOSE: "Baby-proofing": Google's own "Heads up, are you sure?" warning on
// the parts of the sheet nobody should change by accident. Nothing is undone
// by code afterwards (that made Ctrl+Z need several presses): if someone clicks
// through the warning, one Ctrl+Z (or File > Version history) puts it back.
//
// Warned (every tab except Raw Data, Backend and the owner-only tabs):
//   - Row 1 (headers) of Contact Log, Notes and Send Out.
//   - Contact Log: the END bar row (except its O "FILTER?" checkbox) and
//     everything below it.
//   - Events: the header row (row 2) of every event table.
//   - PCAR: everything except H2:I2 ("PUT YOUR OWN INPUT" dates).
//   - Tabs the build writes (Master Table, Combined Contact Tracking, Parents
//     Divided, Directory, Phone Contacts, Students Override Report, New
//     Students): the whole tab.
// Not warned: every place data is typed (Contact Log rows above END, incl.
// column A; Notes / Send Out rows 2+; Events row 1 and rows 3+).
//
// All are RANGE protections (a whole-tab SHEET protection puts a 🔒 on the tab);
// SHEET ones from the older version are removed as stale by applySheetLocks_.
// Warning-only protections never block scripts. Ours carry a description
// starting with SHEET_LOCK_TAG, so applySheetLocks_ can find and redo them
// (end of every build, Initial Setup, new / deleted event, Override Report).
// App Menu "🔓 Unlock Sheet" removes them (doc property SHEET_LOCK_OFF) for
// layout changes; "🔒 Lock Sheet" puts them back.
// Shared code: works the same with or without z_MINE.gs.
// ======================================================================

var SHEET_LOCK_TAG = "🔒 Sheet Lock";
var SHEET_LOCK_OFF_KEY = "SHEET_LOCK_OFF";

// Tabs the build writes (whole tab warned)
var SHEET_LOCK_BUILT_TABS = ["Master Table", "Combined Contact Tracking", "Parents Divided", "Directory",
  "Phone Contacts", "Students Override Report", "New Students"];
// Typed tabs whose row 1 is warned (Contact Log also gets its END bar)
var SHEET_LOCK_HEADER_TABS = ["Contact Log", "Notes", "Send Out"];
var SHEET_LOCK_PCAR = "PCAR";
var SHEET_LOCK_PCAR_INPUT = { row: 2, col: 8, numCols: 2 }; // H2:I2

// Row 2 of every event table (tables are 7 columns: 6 data + 1 spacer); buildEventTable writes these
var EVENT_TABLE_HEADERS = ["OSIS", "Guardian(s)", "Parent Attendee", "Student Name", "# of Indv", "Notes"];
var EVENT_TABLE_WIDTH = 7;

function isSheetLockOn_() {
  return PropertiesService.getDocumentProperties().getProperty(SHEET_LOCK_OFF_KEY) !== "true";
}

function isLockedTab_(name) {
  return SHEET_LOCK_BUILT_TABS.indexOf(name) !== -1 || SHEET_LOCK_HEADER_TABS.indexOf(name) !== -1 ||
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
      "This removes the 'are you sure?' warnings on the headers, the Contact Log END bar, PCAR and the built tabs." +
      "\n\nUse it only to change the layout, then come back here and click '🔒 Lock Sheet'.",
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
// PROTECTIONS
// ======================================================================

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
  const maxR = sheet.getMaxRows(), maxC = sheet.getMaxColumns();
  const range = (r1, c1, r2, c2, what) => {
    if (r2 >= r1 && c2 >= c1) out.push({ kind: "RANGE", range: sheet.getRange(r1, c1, r2 - r1 + 1, c2 - c1 + 1), desc: desc + what });
  };

  // Whole tabs are warned with RANGE protections over every cell, not a sheet
  // protection: same warning, but Google shows no 🔒 on the tab (owner's pick
  // 2026-11-05). A range covers the grid as it is now; the build re-runs this,
  // so rows a build adds are covered again right after.
  if (SHEET_LOCK_BUILT_TABS.indexOf(name) !== -1) {
    range(1, 1, maxR, maxC, "built automatically (🏗️ Build Sheets Only rewrites it)");
  } else if (name === SHEET_LOCK_PCAR) {
    // Everything except H2:I2, as 4 blocks around it
    const inp = SHEET_LOCK_PCAR_INPUT, what = "PCAR formulas (only H2:I2 are yours to change)";
    const lastInpCol = inp.col + inp.numCols - 1;
    range(1, 1, inp.row - 1, maxC, what);                            // rows above
    range(inp.row, 1, inp.row, inp.col - 1, what);                   // left of H2
    range(inp.row, lastInpCol + 1, inp.row, maxC, what);             // right of I2
    range(inp.row + 1, 1, maxR, maxC, what);                         // rows below
  } else if (name === "Events") {
    eventTableStarts_(sheet).forEach(c => range(2, c, 2, c + EVENT_TABLE_HEADERS.length - 1, "event table headers"));
  } else if (SHEET_LOCK_HEADER_TABS.indexOf(name) !== -1) {
    range(1, 1, 1, maxC, "header row");
    const endRow = name === "Contact Log" ? findEndRowInColumnE_(sheet) : null;
    if (endRow) {
      // The O cell (FILTER? checkbox) on the END bar stays free
      range(endRow, 1, endRow, CL_COL_FILTER_TOGGLE - 1, "Contact Log END bar and below");
      range(endRow, CL_COL_FILTER_TOGGLE + 1, endRow, maxC, "Contact Log END bar and below");
      range(endRow + 1, 1, maxR, maxC, "Contact Log END bar and below");
    }
  }
  return out;
}

function sheetLockKey_(sheetName, kind, a1, unprotectedA1s) {
  return [sheetName, kind, a1, (unprotectedA1s || []).join(",")].join("|");
}

/**
 * Adds / refreshes the warning protections. Safe to run any time: only our own
 * tagged protections are touched, unchanged ones are kept, stale ones removed.
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
  clearOldSheetLockData_();

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
}

/** Removes every protection this file added (unlock). */
function removeSheetLocks_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  [SpreadsheetApp.ProtectionType.SHEET, SpreadsheetApp.ProtectionType.RANGE].forEach(type => {
    ss.getProtections(type).forEach(p => {
      if (String(p.getDescription() || "").indexOf(SHEET_LOCK_TAG) !== 0) return;
      try { p.remove(); } catch (err) { console.warn("Sheet lock: could not remove a protection: " + err.message); }
    });
  });
}

/**
 * The first version of this file kept copies of the locked cells in document
 * properties (SHEET_LOCK_SNAP_*, SHEET_LOCK_CL_END). They are no longer used:
 * delete them so they don't take up the properties space.
 */
function clearOldSheetLockData_() {
  try {
    const props = PropertiesService.getDocumentProperties();
    Object.keys(props.getProperties()).forEach(k => {
      if (k.indexOf("SHEET_LOCK_SNAP_") === 0 || k === "SHEET_LOCK_CL_END") props.deleteProperty(k);
    });
  } catch (err) { console.warn("Sheet lock: old data not cleared: " + err.message); }
}
