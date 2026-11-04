// ======================================================================
// FILE: Core.gs
// PURPOSE: Shared foundation used by every other file: the Contact Log column
// numbers, the hidden Backend tab helpers, the owner hook point, and the small
// text / phone / OSIS / date helpers.
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
const CL_COL_DIVIDER_FOLLOWUP = 15; // O — divider line before Followup Notes
const CL_COL_CHECKBOX = 16;         // P — row checkbox

/**
 * Applies the standard Contact Log data-row formatting:
 * - no borders on the Guardian/Spoke With pair (C:D)
 * - no horizontal borders anywhere (only the dark brown end bar carries one)
 * - divider lines before Notes (M) and Followup Notes (O)
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
 * Bold Unicode text (same characters as lingojam.com/BoldTextGenerator):
 * A-Z, a-z and 0-9 become Mathematical Bold letters/digits; everything else
 * (spaces, punctuation, accented letters) is left as is.
 * italic = true uses Bold Italic (serif) letters; Unicode has no italic digits,
 * so digits stay bold.
 */
function boldText_(text, italic) {
  const upper = italic ? 0x1D468 : 0x1D400;
  const lower = italic ? 0x1D482 : 0x1D41A;
  return String(text == null ? "" : text).replace(/[A-Za-z0-9]/g, ch => {
    const c = ch.charCodeAt(0);
    if (c >= 65 && c <= 90) return String.fromCodePoint(upper + c - 65);   // A-Z
    if (c >= 97 && c <= 122) return String.fromCodePoint(lower + c - 97);  // a-z
    return String.fromCodePoint(0x1D7CE + c - 48);                         // 0-9
  });
}

/** Header of one notes line: "𝑫𝒂𝒕𝒆 - {𝐌𝐞𝐭𝐡𝐨𝐝 - 𝐏𝐞𝐫𝐬𝐨𝐧}" (date bold italic, braces bold). */
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
