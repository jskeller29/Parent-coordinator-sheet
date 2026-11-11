// ======================================================================
// FILE: Core.gs
// PURPOSE: Shared foundation used by every other file: the Contact Log column
// numbers, the Student List tab lookup, the hidden Backend tab helpers, the
// owner hook point, and the small text / phone / OSIS / date helpers.
// ======================================================================

// ======================================================================
// CONTACT LOG SCHEMA
// Single source of truth for the Contact Log columns (1-based), plus the
// shared data-row formatter, so a future column shift is changed here only.
// ======================================================================
const CL_WIDTH = 16;                // Full row width, A:P
const CL_COL_GUARDIAN = 3;          // C — Guardian name
const CL_COL_SPOKE_WITH = 4;        // D — Person spoken with
const CL_COL_OSIS = 5;              // E — OSIS (also holds the "END" marker on the end bar row)
const CL_COL_STUDENT = 6;           // F — Student name
const CL_COL_TYPE = 12;             // L — Type dropdown
const CL_COL_DIVIDER_NOTES = 13;    // M — divider line before Notes
const CL_COL_FILTER_TOGGLE = 15;    // O — master hide-filter checkbox on the end bar row
const CL_COL_DIVIDER_FOLLOWUP = 15; // O — divider line before Additional Notes
const CL_COL_CHECKBOX = 16;         // P — row checkbox

/**
 * Applies the standard Contact Log data-row formatting:
 * - no borders on the Guardian/Spoke With pair (C:D)
 * - no horizontal borders anywhere (only the dark brown end bar carries one)
 * - divider lines before Event Description (M) and Additional Notes (O)
 * - checkboxes planted in P
 */
function formatContactLogRows_(sheet, startRow, numRows) {
  if (!numRows || numRows < 1) return;
  sheet.getRange(startRow, CL_COL_GUARDIAN, numRows, 2).setBorder(false, false, false, false, false, false);
  sheet.getRange(startRow, 1, numRows, CL_WIDTH).setBorder(false, null, false, null, null, false);
  sheet.getRange(startRow, CL_COL_DIVIDER_NOTES, numRows, 1).setBorder(null, null, null, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(startRow, CL_COL_DIVIDER_FOLLOWUP, numRows, 1).setBorder(null, null, null, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(startRow, CL_COL_CHECKBOX, numRows, 1).insertCheckboxes().setFontColor("#000000");
}

/** Row number of the Contact Log's "END" bar (Column E), or null. */
function findEndRowInColumnE_(sheet) {
  // Only read down to the last data row — this runs twice per edit, and
  // "E:E" pulls every row in the sheet.
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return null;
  const values = sheet.getRange(1, CL_COL_OSIS, lastRow, 1).getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0].trim().toUpperCase() === "END") return i + 1;
  }
  return null;
}

// ======================================================================
// STUDENT LIST TAB (called "Master Table" before Nov 2026)
// The generated roster: OSIS, Student, Guardian, Site, Class, Sibling, Status,
// Last Interaction (header on row 4, data from row 5), ONE ROW PER STUDENT:
// no sibling-group rows. Every lookup goes through getStudentListSheet_ so an
// older copy's "Master Table" tab is renamed instead of a second, empty tab
// being made next to it.
// The autofill (ghost typists, Events dropdowns, quick updates) also needs the
// sibling-group rows ("111, 222" / "Ann X & Ben X"); it reads the full table
// from the Backend STUDENT_LIST section instead (studentLookupSource_ below).
// The owner's "Old Master Table" is a different tab and keeps its name.
// (Lists built when the files load - TAB_VISIBILITY, KEEP_TABS_,
// SHEET_LOCK_BUILT_TABS - spell the name out, so file load order can't matter.)
// ======================================================================
var STUDENT_LIST_SHEET = "Student List";
var LEGACY_STUDENT_LIST_SHEET = "Master Table";

/**
 * The Student List tab, or null when the sheet has neither name. An old
 * "Master Table" tab is renamed to "Student List" the first time, so its
 * Discharged students and statuses carry over to the next build.
 */
function getStudentListSheet_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(STUDENT_LIST_SHEET);
  if (sheet) return sheet;

  const legacy = ss.getSheetByName(LEGACY_STUDENT_LIST_SHEET);
  if (!legacy) return null;
  try {
    legacy.setName(STUDENT_LIST_SHEET);
  } catch (e) {
    // e.g. a view-only user opening the sheet; the old tab still works this run
    console.warn('Could not rename "Master Table" to "Student List": ' + e.message);
  }
  return legacy;
}

/**
 * Where the autofill looks students up: the Backend STUDENT_LIST section (BX:CD,
 * the full Student List WITH the sibling-group rows, written by every build).
 * Until a build has written it (a sheet updated to this version but not built
 * since), the Student List tab instead, whose last build still has those rows.
 * Returns { sheet, firstRow, col } (first data row; sheet column of OSIS, then
 * Student, Guardian, Site, Class, Sibling, Status) or null when neither exists.
 */
function studentLookupSource_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const b = BACKEND_BLOCKS.STUDENT_LIST;
  const backend = ss.getSheetByName(BACKEND_SHEET_NAME);
  if (backend && backend.getMaxColumns() >= b.col + b.width - 1 &&
      String(backend.getRange(BACKEND_HEADER_ROW, b.col).getValue()).trim().toUpperCase() === "OSIS") {
    return { sheet: backend, firstRow: BACKEND_HEADER_ROW + 1, col: b.col };
  }
  const list = getStudentListSheet_(ss);
  return list ? { sheet: list, firstRow: 5, col: 1 } : null;
}

/**
 * The lookup source's data rows (no header), `width` columns from OSIS; trailing
 * blank rows dropped. src: an already-found studentLookupSource_ (optional).
 */
function studentLookupRows_(ss, width, src) {
  src = src || studentLookupSource_(ss);
  if (!src) return [];
  const last = src.sheet.getLastRow();
  if (last < src.firstRow) return [];
  const vals = src.sheet.getRange(src.firstRow, src.col, last - src.firstRow + 1, width).getValues();
  let end = vals.length;
  while (end > 0 && !vals[end - 1].some(v => String(v).trim() !== "")) end--;
  return vals.slice(0, end);
}

/**
 * Open-ended range of one lookup column from the first data row down
 * (offset 0 = OSIS, 1 = Student, 2 = Guardian), e.g. Backend!BX3:BX.
 */
function studentLookupColumn_(src, offset) {
  const letter = columnLetter_(src.col + offset);
  return src.sheet.getRange(letter + src.firstRow + ":" + letter);
}

/** Sibling-group row of the full Student List (its OSIS cell lists several: "111, 222"). */
function isSiblingGroupRow_(row) {
  return String(row[0]).includes(",");
}

/** 1 -> "A", 76 -> "BX". */
function columnLetter_(col) {
  let s = "";
  for (let n = col; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s;
  return s;
}

/**
 * ==============================================================================
 * BACKEND SHEET
 * One hidden "Backend" tab holds several side-by-side sections that used to be
 * separate tabs. Row 1 carries a title over each section; each section's old
 * layout sits one row lower (old row 1 -> row 2).
 *
 *   A:E    Version & Settings   (was "Version")
 *   G:AV   New / Edit Student   (was "New/Edit Student", 42 columns)
 *   AX:BF  Backend Event Log    (was "Backend_Event_Log", 9 columns)
 *   BH:BJ  Send Out List        (feeds the "Send Out" sheet; not used by code)
 *   BL     Type List            (was "Type List")
 *   BN:BQ  Additional Information (OSIS / Student / Guardian / lines; written by
 *          every build from Your Lists, Z#OwnData.gs; read by the phone contacts)
 *   BS:BV  Your Lists setup     (one row per list tab, Z#OwnData.gs)
 *   BX:CD  Student List         (OSIS..Status WITH the sibling-group rows;
 *          written by every build, read by the autofill, studentLookupSource_)
 * The last three are "grow" sections: an older Backend tab gets their columns the
 * first time they are written; reading them before that returns nothing.
 *
 * Sections share rows, so NEVER delete/insert whole sheet rows, use
 * getDataRange()/getLastRow(), or clear to the right on this tab. Always go
 * through the helpers below, which only ever touch one section's columns.
 * ==============================================================================
 */
var BACKEND_SHEET_NAME = "Backend";
var BACKEND_HEADER_ROW = 2; // table sections: header row; data starts on row 3

var BACKEND_BLOCKS = {
  VERSION:   { col: 1,  width: 5,  title: "VERSION & SETTINGS", legacy: "Version" },
  OVERRIDES: { col: 7,  width: 42, title: "NEW / EDIT STUDENT", legacy: "New/Edit Student" },
  EVENT_LOG: { col: 50, width: 9,  title: "BACKEND EVENT LOG",  legacy: "Backend_Event_Log" },
  SEND_OUT:  { col: 60, width: 3,  title: "SEND OUT LIST",      legacy: null },
  TYPES:     { col: 64, width: 1,  title: "TYPE LIST",          legacy: "Type List" },
  ADDL_INFO: { col: 66, width: 4,  title: "ADDITIONAL INFORMATION", legacy: null, grow: true },
  OWN_LISTS: { col: 71, width: 4,  title: "YOUR LISTS (SETUP)",  legacy: null, grow: true },
  STUDENT_LIST: { col: 76, width: 7, title: "STUDENT LIST (WITH SIBLING GROUPS)", legacy: null, grow: true }
};

// Fixed cells inside the Version section (old "Version" tab, moved down one row)
var BACKEND_VERSION = {
  BUILD_DATE: "A2",          // was A1
  MIGRATE_LABEL: "A3",       // was A2
  MIGRATE_TOGGLE: "B3",      // was B2
  LINKS_ROW: 2,              // D2:E4 template links (was D1:E3)
  GUIDE_ROW: 5,              // D5:E5 guide link (was D4:E4)
  SETTINGS_HEADER_ROW: 5,    // A5 (was A4)
  SETTINGS_FIRST_ROW: 6      // A6:B (was A5:B)
};

/**
 * Returns the Backend tab, creating it (hidden) when missing and moving any
 * old separate tabs (Version, New/Edit Student, Backend_Event_Log, Type List)
 * into their sections the first time it runs.
 */
function getBackendSheet_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(BACKEND_SHEET_NAME);
  const legacyTabs = Object.keys(BACKEND_BLOCKS)
    .filter(k => BACKEND_BLOCKS[k].legacy && ss.getSheetByName(BACKEND_BLOCKS[k].legacy));

  if (sheet && legacyTabs.length === 0) return sheet;

  if (!sheet) {
    sheet = ss.insertSheet(BACKEND_SHEET_NAME);
    sheet.hideSheet();
  }
  const neededCols = Math.max.apply(null, Object.keys(BACKEND_BLOCKS).map(k => BACKEND_BLOCKS[k].col + BACKEND_BLOCKS[k].width - 1));
  if (sheet.getMaxColumns() < neededCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), neededCols - sheet.getMaxColumns());
  }
  writeBackendTitles_(sheet);
  legacyTabs.forEach(key => migrateLegacyTabToBackend_(ss, sheet, key));
  return sheet;
}

/** Fills in each section's row-1 title where the cell is still blank. */
function writeBackendTitles_(sheet) {
  Object.keys(BACKEND_BLOCKS).forEach(key => {
    const b = BACKEND_BLOCKS[key];
    const cell = sheet.getRange(1, b.col);
    if (String(cell.getValue()).trim() === "") cell.setValue(b.title).setFontWeight("bold");
  });
}

/**
 * Copies an old separate tab into its Backend section (shifted down one row)
 * and deletes the old tab. Skips (and keeps the old tab) when the section
 * already has data, so nothing is ever overwritten.
 */
function migrateLegacyTabToBackend_(ss, sheet, key) {
  const b = BACKEND_BLOCKS[key];
  const legacy = ss.getSheetByName(b.legacy);
  if (!legacy) return;

  if (backendBlockHasData_(sheet, key)) {
    console.warn(`Backend migration: "${b.legacy}" kept because the ${b.title} section already has data.`);
    return;
  }

  const lastRow = legacy.getLastRow();
  if (lastRow > 0) {
    const values = readLegacyTab_(legacy, b.width);
    ensureBackendRows_(sheet, BACKEND_HEADER_ROW + values.length - 1);
    sheet.getRange(BACKEND_HEADER_ROW, b.col, values.length, b.width).setValues(values);
  }

  if (key === "TYPES") repointTypeValidation_(ss, legacy, sheet);
  if (key === "VERSION") {
    // Keep the build date looking like a date after the move
    sheet.getRange(BACKEND_VERSION.BUILD_DATE).setNumberFormat("MM/dd/yy");
  }

  ss.deleteSheet(legacy);
}

/**
 * The Contact Log "Type" dropdown (column L) points at the old Type List tab;
 * repoint those rules to the Backend Type List section before that tab is deleted.
 */
function repointTypeValidation_(ss, legacyTypeSheet, backendSheet) {
  const log = ss.getSheetByName("Contact Log");
  if (!log || log.getMaxRows() < 2) return;
  const range = log.getRange(2, 12, log.getMaxRows() - 1, 1);
  const rules = range.getDataValidations();
  const typeCol = BACKEND_BLOCKS.TYPES.col;
  const newRange = backendSheet.getRange(BACKEND_HEADER_ROW, typeCol, backendSheet.getMaxRows() - 1, 1);
  let changed = false;

  const updated = rules.map(r => {
    const rule = r[0];
    if (!rule || rule.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) return [rule];
    const args = rule.getCriteriaValues();
    if (!args[0] || args[0].getSheet().getSheetId() !== legacyTypeSheet.getSheetId()) return [rule];
    changed = true;
    return [rule.copy().requireValueInRange(newRange, args[1] !== false).build()];
  });

  if (changed) range.setDataValidations(updated);
}

function ensureBackendRows_(sheet, neededRows) {
  if (sheet.getMaxRows() < neededRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), neededRows - sheet.getMaxRows());
  }
}

function backendBlockHasData_(sheet, key) {
  const b = BACKEND_BLOCKS[key];
  const lastRow = sheet.getLastRow();
  if (lastRow < BACKEND_HEADER_ROW) return false;
  const vals = sheet.getRange(BACKEND_HEADER_ROW, b.col, lastRow - BACKEND_HEADER_ROW + 1, b.width).getValues();
  return vals.some(r => r.some(v => String(v).trim() !== ""));
}

/**
 * Reads one table section like getDataRange().getValues() used to on the old
 * tab: index 0 = header row (sheet row 2), index i = sheet row 2 + i, and
 * trailing blank rows are dropped. Pass display=true for getDisplayValues().
 */
function backendReadBlock_(ss, key, display) {
  const sheet = getBackendSheet_(ss);
  return readBlockFromSheet_(sheet, key, display);
}

function readBlockFromSheet_(sheet, key, display) {
  const b = BACKEND_BLOCKS[key];
  if (!backendBlockReady_(sheet, b, false)) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < BACKEND_HEADER_ROW) return [];
  const range = sheet.getRange(BACKEND_HEADER_ROW, b.col, lastRow - BACKEND_HEADER_ROW + 1, b.width);
  const vals = display ? range.getDisplayValues() : range.getValues();
  let end = vals.length;
  while (end > 0 && !vals[end - 1].some(v => String(v).trim() !== "")) end--;
  return vals.slice(0, end);
}

/**
 * Reads a table section from ANOTHER spreadsheet (e.g. an older copy being
 * migrated) without changing it: its Backend section if it has one, otherwise
 * the old separate tab.
 */
function readBackendBlockFromOtherSpreadsheet_(otherSs, key) {
  const b = BACKEND_BLOCKS[key];
  const backend = otherSs.getSheetByName(BACKEND_SHEET_NAME);
  if (backend) return readBlockFromSheet_(backend, key, false);
  const legacy = b.legacy ? otherSs.getSheetByName(b.legacy) : null;
  if (!legacy || legacy.getLastRow() < 1) return [];
  return readLegacyTab_(legacy, b.width);
}

/** Old tab's values from A1, exactly `width` columns wide (padded if the tab is narrower). */
function readLegacyTab_(legacy, width) {
  const cols = Math.min(width, legacy.getMaxColumns());
  return legacy.getRange(1, 1, legacy.getLastRow(), cols).getValues().map(r => padRow_(r, width));
}

/** Sheet row number for index i of a backendReadBlock_ result. */
function backendRowFor_(i) {
  return BACKEND_HEADER_ROW + i;
}

/** Range inside a section: i = row index as in backendReadBlock_, c = 0-based column. */
function backendBlockRange_(ss, key, i, c, numRows, numCols) {
  if (key === "EVENT_LOG" && typeof clearNotesSourcesCache_ === "function") clearNotesSourcesCache_(); // cached copy is stale now
  if (key === "ADDL_INFO") clearAdditionalInfoCache_(ss); // Build.gs keeps a cached copy too
  const sheet = getBackendSheet_(ss);
  return sheet.getRange(backendRowFor_(i), BACKEND_BLOCKS[key].col + c, numRows || 1, numCols || 1);
}

/** Replaces a whole table section (header + rows) starting at row 2. */
function backendWriteBlock_(ss, key, values) {
  if (key === "EVENT_LOG" && typeof clearNotesSourcesCache_ === "function") clearNotesSourcesCache_(); // cached copy is stale now
  const sheet = getBackendSheet_(ss);
  const b = BACKEND_BLOCKS[key];
  backendBlockReady_(sheet, b, true);
  backendClearBlock_(ss, key, BACKEND_HEADER_ROW);
  if (!values || values.length === 0) return;
  const rows = values.map(r => padRow_(r, b.width));
  ensureBackendRows_(sheet, BACKEND_HEADER_ROW + rows.length - 1);
  sheet.getRange(BACKEND_HEADER_ROW, b.col, rows.length, b.width).setValues(rows);
}

/** Clears a section's contents from fromRow down (only its own columns). */
function backendClearBlock_(ss, key, fromRow) {
  if (key === "EVENT_LOG" && typeof clearNotesSourcesCache_ === "function") clearNotesSourcesCache_(); // cached copy is stale now
  if (key === "ADDL_INFO") clearAdditionalInfoCache_(ss); // Build.gs keeps a cached copy too (a write clears first, so it lands here)
  const sheet = getBackendSheet_(ss);
  const b = BACKEND_BLOCKS[key];
  if (!backendBlockReady_(sheet, b, false)) return; // a grow section that was never written: nothing to clear
  const rows = sheet.getMaxRows() - fromRow + 1;
  if (rows > 0) sheet.getRange(fromRow, b.col, rows, b.width).clearContent();
}

/**
 * A "grow" section (Additional Information, Your Lists setup, Student List) sits right of the
 * old ones, where an older Backend tab has no columns yet. create = true adds the
 * columns and the section title; false just says whether they are there. Other
 * sections answer true without any sheet call.
 */
function backendBlockReady_(sheet, b, create) {
  if (!b.grow) return true;
  const need = b.col + b.width - 1;
  if (sheet.getMaxColumns() < need) {
    if (!create) return false;
    sheet.insertColumnsAfter(sheet.getMaxColumns(), need - sheet.getMaxColumns());
  }
  if (create) {
    const title = sheet.getRange(1, b.col);
    if (String(title.getValue()).trim() === "") title.setValue(b.title).setFontWeight("bold");
  }
  return true;
}

/** Appends rows directly under the section's own last row. */
function backendAppendRows_(ss, key, rows) {
  if (key === "EVENT_LOG" && typeof clearNotesSourcesCache_ === "function") clearNotesSourcesCache_(); // cached copy is stale now
  if (key === "ADDL_INFO") clearAdditionalInfoCache_(ss); // Build.gs keeps a cached copy too
  if (!rows || rows.length === 0) return;
  const sheet = getBackendSheet_(ss);
  const b = BACKEND_BLOCKS[key];
  backendBlockReady_(sheet, b, true);
  const current = readBlockFromSheet_(sheet, key, false);
  const startRow = BACKEND_HEADER_ROW + Math.max(current.length, 1);
  const padded = rows.map(r => padRow_(r, b.width));
  ensureBackendRows_(sheet, startRow + padded.length - 1);
  sheet.getRange(startRow, b.col, padded.length, b.width).setValues(padded);
}

/**
 * Deletes rows from one section only (cells shift up within its columns), so
 * the other sections are untouched. idxs are backendReadBlock_ indices (>= 1).
 */
function backendDeleteRows_(ss, key, idxs) {
  if (key === "EVENT_LOG" && typeof clearNotesSourcesCache_ === "function") clearNotesSourcesCache_(); // cached copy is stale now
  const sheet = getBackendSheet_(ss);
  const b = BACKEND_BLOCKS[key];
  Array.from(new Set(idxs)).filter(i => i >= 1).sort((x, y) => y - x).forEach(i => {
    sheet.getRange(backendRowFor_(i), b.col, 1, b.width).deleteCells(SpreadsheetApp.Dimension.ROWS);
  });
}

/** Sorts a section's data rows. specs: [{col: 0-based, ascending: bool}] */
function backendSortBlock_(ss, key, specs) {
  if (key === "EVENT_LOG" && typeof clearNotesSourcesCache_ === "function") clearNotesSourcesCache_(); // cached copy is stale now
  const sheet = getBackendSheet_(ss);
  const b = BACKEND_BLOCKS[key];
  const n = readBlockFromSheet_(sheet, key, false).length - 1;
  if (n < 2) return;
  sheet.getRange(BACKEND_HEADER_ROW + 1, b.col, n, b.width)
    .sort(specs.map(s => ({ column: b.col + s.col, ascending: s.ascending })));
}

function padRow_(row, width) {
  const out = row.slice(0, width);
  while (out.length < width) out.push("");
  return out;
}

// ======================================================================
// OWNER HOOKS
// The owner's own sheet also has z_MINE.gs, which defines ownerHooks_().
// Shared code asks getOwnerHooks_() and, only when it returns something,
// calls the owner's version of that one step (history from "Old Master
// Table" / "Old Phone Contacts", the tall Phone Contacts layout, and so on).
// Every other school's sheet has no z_MINE.gs, so the shared path runs.
// No function is defined twice, so file load order no longer matters.
// ======================================================================
var OWNER_HOOKS_CACHE_; // looked up once per run

function getOwnerHooks_() {
  if (OWNER_HOOKS_CACHE_ === undefined) {
    OWNER_HOOKS_CACHE_ = typeof ownerHooks_ === "function" ? ownerHooks_() : null;
  }
  return OWNER_HOOKS_CACHE_;
}

/**
 * Runs one user-facing action (an edit, a build, a sync, opening the sheet).
 * Normally it just calls fn. On the owner's sheet with "Toggle Debug" on, it
 * also toasts whether the action worked (see ownerDebugRun_ in z_MINE.gs).
 * label may be a function so the name is only worked out when it is shown.
 */
function runAction_(label, fn) {
  const owner = getOwnerHooks_();
  if (owner && owner.debugOn()) return owner.debugRun(label, fn);
  return fn();
}

// ======================================================================
// EDIT DEBUG LOG (read by the optional 🐞 Contact Log Debug sidebar,
// Z#ContactLogDebug.gs). Off unless that sidebar is open: it keeps the
// document cache key CL_DEBUG_ON alive while it polls, so recording stops by
// itself a few minutes after the sidebar is closed. When off, each edit costs
// one cache read. Each run's steps (ms since the script started) go to the
// document cache key CL_DEBUG_LOG (newest 25 runs).
// ======================================================================
var CL_DEBUG_RUN_ = null;
var CL_DEBUG_LOG_KEY = "CL_DEBUG_LOG";
var CL_DEBUG_ON_KEY = "CL_DEBUG_ON";

function clDbgStart_(label) {
  CL_DEBUG_RUN_ = null;
  try {
    if (CacheService.getDocumentCache().get(CL_DEBUG_ON_KEY) !== "1") return;
    if (typeof label === "function") label = label(); // only worked out when recording
    CL_DEBUG_RUN_ = { label: String(label), at: new Date().toLocaleTimeString(), t0: Date.now(), steps: [] };
  } catch (e) { CL_DEBUG_RUN_ = null; }
}

/** One step of the current run (no-op when recording is off). */
function clDbg_(msg) {
  if (!CL_DEBUG_RUN_) return;
  CL_DEBUG_RUN_.steps.push([Date.now() - CL_DEBUG_RUN_.t0, String(msg).slice(0, 400)]);
}

function clDbgEnd_(error) {
  const run = CL_DEBUG_RUN_;
  CL_DEBUG_RUN_ = null;
  if (!run) return;
  try {
    if (error) run.steps.push([Date.now() - run.t0, "❌ ERROR: " + (error && error.message ? error.message : error)]);
    run.total = Date.now() - run.t0;
    delete run.t0;
    const cache = CacheService.getDocumentCache();
    let log = [];
    try { log = JSON.parse(cache.get(CL_DEBUG_LOG_KEY) || "[]"); } catch (e) { log = []; }
    log.unshift(run);
    log = log.slice(0, 25);
    let json = JSON.stringify(log);
    while (json.length > 90000 && log.length > 1) { log.pop(); json = JSON.stringify(log); }
    cache.put(CL_DEBUG_LOG_KEY, json, 21600);
  } catch (e) { /* debug only */ }
}

/** Runs fn as one recorded run (errors are recorded, then thrown as before). */
function clDbgRun_(label, fn) {
  clDbgStart_(label);
  try {
    const out = fn();
    clDbgEnd_();
    return out;
  } catch (err) {
    clDbgEnd_(err);
    throw err;
  }
}

// ======================================================================
// SMALL HELPERS (text, names, phones, OSIS, dates)
// ======================================================================
function writeSheet_(ss, name, values, startRow = 1, startCol = 1) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  // Guard: never crash on an empty payload. A builder that returns [] (or only
  // filtered everything out) would otherwise throw on values[0].length below.
  if (!Array.isArray(values) || values.length === 0 || !Array.isArray(values[0]) || values[0].length === 0) {
    return;
  }

  const neededRows = startRow - 1 + values.length;
  const neededCols = startCol - 1 + values[0].length;

  if (sheet.getMaxRows() < neededRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), neededRows - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < neededCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), neededCols - sheet.getMaxColumns());
  }

  const clearRows = sheet.getMaxRows() - startRow + 1;
  const clearCols = sheet.getMaxColumns() - startCol + 1;
  if (clearRows > 0 && clearCols > 0) {
    sheet.getRange(startRow, startCol, clearRows, clearCols).clearContent();
  }

  sheet.getRange(startRow, startCol, values.length, values[0].length).setValues(values);
}

// Same result as a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" }),
// but the collator is built once instead of on every comparison (faster sorts).
const NATURAL_COLLATOR_ = new Intl.Collator(undefined, { numeric: true, sensitivity: "base" });

function compareNatural_(a, b) {
  return NATURAL_COLLATOR_.compare(a, b);
}

function toText_(value) {
  return value == null ? "" : String(value).trim();
}

function firstWord_(text) {
  return toText_(text).split(/\s+/)[0] || "";
}

function isRealStudentId_(value) {
  return toText_(value).replace(/\D/g, "").length >= 6;
}

function stripTrailingParenthetical_(text) {
  return toText_(text).replace(/\s*\([^)]+\)\s*$/, "").trim();
}

function properCase_(text) {
  const raw = toText_(text);
  if (!raw) return "";

  // If the input already has intentional internal capitals (McDonald, DiMarco,
  // O'Brien, DeShawn), leave it exactly as typed. Only normalize inputs that are
  // clearly un-cased: ALL CAPS ("SMITH") or all lowercase ("smith").
  const hasLower = /\p{Ll}/u.test(raw);
  const hasUpper = /\p{Lu}/u.test(raw);
  if (hasLower && hasUpper) return raw;

  // \p{L} + u flag so accented initials (Ángel, Óscar) capitalize correctly.
  return raw.toLowerCase().replace(/(^|[\s\-'’])(\p{L})/gu, (m, sep, ch) => sep + ch.toUpperCase());
}

function splitFullName_(name) {
  // The owner's sheet splits off the LAST word as the last name instead
  const owner = getOwnerHooks_();
  if (owner) return owner.splitFullName(name);

  const parts = toText_(name).split(/\s+/).filter(Boolean);
  return {
    first: parts[0] || "",
    last: parts.slice(1).join(" ")
  };
}

function normalizeEmail_(value) {
  return toText_(value).toLowerCase().replace(/\s+/g, "");
}

function normalizePhoneForMatch_(value) {
  // Drop an extension ("x12", "ext. 12") so its digits don't make the number unmatchable
  const digits = toText_(value).replace(/\s*(?:ext\.?|extension|x)\s*\d+\s*$/i, "").replace(/\D/g, "");
  if (!digits) return "";
  return digits.length === 11 && digits.startsWith("1") ? digits.slice(1) : digits;
}

function formatPhone_(value) {
  const digits = normalizePhoneForMatch_(value);
  if (digits.length === 10) {
    return `(${digits.slice(0, 3)}) ${digits.slice(3, 6)}-${digits.slice(6)}`;
  }
  return "";
}

/* =========================
   OSIS MATCHING (shared)
   Tokenize a cell into whole OSIS numbers and match EXACTLY, so a short id can
   never substring-match inside a longer one (e.g. "123456" inside "812345678").
========================= */
function extractOsisTokens_(value) {
  const str = value == null ? "" : String(value);
  const matches = str.match(/\b\d{6,}\b/g);
  if (matches && matches.length) return matches;
  // Fallback for any stray formats that aren't purely numeric.
  return str.split(",").map(s => s.trim()).filter(Boolean);
}

function osisCellIncludes_(cellValue, targetOsis) {
  const target = String(targetOsis == null ? "" : targetOsis).trim();
  if (!target) return false;
  return extractOsisTokens_(cellValue).indexOf(target) !== -1;
}

/**
 * Every OSIS in column A of the given Student List rows (a sibling row "111, 222"
 * counts each one), as normalizeOsis_ keys, so a lost leading zero still matches.
 */
function knownOsisSet_(rows) {
  const set = new Set();
  rows.forEach(r => extractOsisTokens_(r[0]).forEach(o => {
    const key = normalizeOsis_(o);
    if (key) set.add(key);
  }));
  return set;
}

function groupGuardiansByStudent_(guardianRows) {
  const map = new Map();
  for (const g of guardianRows) {
    if (!map.has(g.student_id)) map.set(g.student_id, []);
    map.get(g.student_id).push(g);
  }
  return map;
}

function formatDateFull_(value) {
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value)) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "MMMM d, yyyy");
  }

  const text = toText_(value);
  if (!text) return "";

  const parsed = new Date(text);
  if (!isNaN(parsed)) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), "MMMM d, yyyy");
  }

  return text;
}

/**
 * Bold Unicode text: A-Z, a-z and 0-9 become Mathematical Sans-Serif Bold
 * letters/digits (𝗕𝗼𝗹𝗱 𝟭𝟮𝟯, the "Bold Sans" sets on yaytext / lingojam);
 * everything else (spaces, punctuation, accented letters) is left as is.
 * italic = true uses Sans-Serif Bold Italic letters (𝘽𝙤𝙡𝙙); Unicode has no
 * italic digits, so digits stay bold.
 * Sans-serif (owner's pick 2026-11-05) to match the sheet's Montserrat text;
 * it was the serif Mathematical Bold / Bold Italic set before.
 * Used for the notes-line headers: a formula (the search bar's Communications
 * box) can't show real bold for part of a cell, but these letters stay bold.
 */
function boldText_(text, italic) {
  const upper = italic ? 0x1D63C : 0x1D5D4;
  const lower = italic ? 0x1D656 : 0x1D5EE;
  return String(text == null ? "" : text).replace(/[A-Za-z0-9]/g, ch => {
    const c = ch.charCodeAt(0);
    if (c >= 65 && c <= 90) return String.fromCodePoint(upper + c - 65);   // A-Z
    if (c >= 97 && c <= 122) return String.fromCodePoint(lower + c - 97);  // a-z
    return String.fromCodePoint(0x1D7EC + c - 48);                         // 0-9 (sans-serif bold)
  });
}

/** Header of one notes line: "𝘿𝙖𝙩𝙚 - {𝗠𝗲𝘁𝗵𝗼𝗱 - 𝗣𝗲𝗿𝘀𝗼𝗻}" (date bold italic sans, braces bold sans). */
function noteLineHeader_(dateText, braceText) {
  return `${boldText_(dateText, true)} - {${boldText_(braceText)}}`;
}

// ======================================================================
// DIALOG STYLES
// Every dialog / sidebar shares one look, kept in Styles.html.
// HTML files pull it in with a scriptlet (<?!= include_('Styles') ?>) and are
// opened with createTemplateFromFile(...).evaluate(); dialogs built as strings
// in the .gs files put include_('Styles') at the front of their HTML.
// ======================================================================
function include_(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/** Opens an HTML file as a template so its include_('Styles') line is filled in. */
function htmlFromFile_(name) {
  return HtmlService.createTemplateFromFile(name).evaluate();
}
