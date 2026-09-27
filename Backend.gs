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
  TYPES:     { col: 64, width: 1,  title: "TYPE LIST",          legacy: "Type List" }
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
  const neededCols = BACKEND_BLOCKS.TYPES.col;
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
  const sheet = getBackendSheet_(ss);
  return sheet.getRange(backendRowFor_(i), BACKEND_BLOCKS[key].col + c, numRows || 1, numCols || 1);
}

/** Replaces a whole table section (header + rows) starting at row 2. */
function backendWriteBlock_(ss, key, values) {
  const sheet = getBackendSheet_(ss);
  const b = BACKEND_BLOCKS[key];
  backendClearBlock_(ss, key, BACKEND_HEADER_ROW);
  if (!values || values.length === 0) return;
  const rows = values.map(r => padRow_(r, b.width));
  ensureBackendRows_(sheet, BACKEND_HEADER_ROW + rows.length - 1);
  sheet.getRange(BACKEND_HEADER_ROW, b.col, rows.length, b.width).setValues(rows);
}

/** Clears a section's contents from fromRow down (only its own columns). */
function backendClearBlock_(ss, key, fromRow) {
  const sheet = getBackendSheet_(ss);
  const b = BACKEND_BLOCKS[key];
  const rows = sheet.getMaxRows() - fromRow + 1;
  if (rows > 0) sheet.getRange(fromRow, b.col, rows, b.width).clearContent();
}

/** Appends rows directly under the section's own last row. */
function backendAppendRows_(ss, key, rows) {
  if (!rows || rows.length === 0) return;
  const sheet = getBackendSheet_(ss);
  const b = BACKEND_BLOCKS[key];
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
  const sheet = getBackendSheet_(ss);
  const b = BACKEND_BLOCKS[key];
  Array.from(new Set(idxs)).filter(i => i >= 1).sort((x, y) => y - x).forEach(i => {
    sheet.getRange(backendRowFor_(i), b.col, 1, b.width).deleteCells(SpreadsheetApp.Dimension.ROWS);
  });
}

/** Sorts a section's data rows. specs: [{col: 0-based, ascending: bool}] */
function backendSortBlock_(ss, key, specs) {
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
