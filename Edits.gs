// ======================================================================
// FILE: Edits.gs
// PURPOSE: Everything that runs while someone types: the installable onEdit
// trigger, the ghost typists (Contact Log, Notes, Events), Raw Data paste
// clean-up, and the fast single-student "quick update" of Combined Contact
// Tracking / Phone Contacts. Also the Events tables (menu + watcher).
// ======================================================================

// =========================================================
// MAIN TRIGGER FUNCTION (installable onEdit)
// =========================================================
function installedOnEdit(e) {
  if (!e || !e.range) return;
  const result = runAction_(() => "Edit " + e.range.getSheet().getName() + "!" + e.range.getA1Notation(), () => handleEdit_(e));
  // Data update reminder while the sheet is open (at most one check an hour; Menus.gs)
  try { checkDataReminderWhileOpen_(); } catch (err) { console.warn("Data reminder check failed: " + err.message); }
  return result;
}

function handleEdit_(e) {

  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName().trim();

  // --- RAW DATA PASTE INTERCEPTOR ---
  if (sheetName === "RAW Data" || sheetName === "Raw Data") {
    if (range.getNumRows() > 1 || range.getNumColumns() > 1) {
      cleanRawDataPasteZone_(sheet, range);
    }
    fixRawDataNumberFormatting_(sheet);
    updateRawDataTimestamp_(range);
    lockRawDataRowHeights_(sheet);
    return; 
  }

  // --- TRIGGER THE EVENTS GHOST TYPIST & BACKEND ---
  if (sheetName === "Events") {
    watchEventsSheet_(e);
    refreshSearchBarCommunications_(null); // an event's notes show in the search bar too
    return; 
  }

  // --- NEW: TRIGGER THE NOTES GHOST TYPIST ---
  if (sheetName === "Notes") {
    watchNotesSheet_(e);
    refreshSearchBarCommunications_(null); // Notes-tab notes show in the search bar too
    return; 
  }

  // If not Contact Log, stop here.
  if (sheetName !== "Contact Log") return;
  
  const row = range.getRow();
  const numRows = range.getNumRows(); 
  const col = range.getColumn();
  // 🛡️ FIX: Extract editedValue safely early on so the whole function can use it
  const editedValue = String(e.value !== undefined ? e.value : range.getDisplayValue()).trim();

  if (row < 2) return;

  // 1. Shifted End Row Finder (Now Col E / 5). ONE cache call brings back the
  // END row and the ghost typist's data; the END row is then checked with a
  // one-cell read (a full column E scan only when it moved).
  const prefetch = contactLogPrefetch_(sheet);
  let endRow = prefetch.endRow;

  // --- NEW FILTER LOGIC ---
  // The Master Toggle lives on the end bar row in Col O
  if (endRow && row === endRow && col === CL_COL_FILTER_TOGGLE) {
    toggleHideFilter_(sheet, endRow);
    return; 
  }
  // ------------------------

  if (endRow && row >= endRow) {
    // An edit in the search bar below the END bar (e.g. looking up a student):
    // draw its Communications box right away
    if (row > endRow) refreshSearchBarCommunications_(sheet, endRow, null, true);
    return;
  }

  // =======================================================
  // CONTACT LOG GHOST TYPIST — PASTED COLUMN (fills every pasted row)
  // =======================================================
  if (numRows > 1 && (col === CL_COL_GUARDIAN || col === CL_COL_OSIS || col === CL_COL_STUDENT)) {
    try {
      const pasteRows = endRow ? Math.min(numRows, endRow - row) : numRows;
      if (pasteRows > 0) fillPastedContactLogRows_(e.source, sheet, row, pasteRows, col, endRow);
    } catch (err) {
      console.error(err);
      e.source.toast("Auto-fill failed: " + err.message, "⚠️ Ghost Typist Error", 5);
    }
  }

  // =======================================================
  // CONTACT LOG GHOST TYPIST (OSIS, Guardian, Student)
  // Runs FIRST so the auto-filled values land on screen before the
  // cosmetic border/checkbox repair work below.
  // =======================================================
  if (range.getNumRows() === 1 && range.getNumColumns() === 1) {
    try {
      // Watch columns: C (Guardian), E (OSIS), F (Student)
      if ((col === CL_COL_GUARDIAN || col === CL_COL_OSIS || col === CL_COL_STUDENT) && editedValue && editedValue.toUpperCase() !== "FALSE" && editedValue.toUpperCase() !== "TRUE") {
        // Master Table A5:C + Directory C:E, from a short cache when the
        // previous edit already read them (ghostTypistData_)
        const ghost = ghostTypistData_(e.source, prefetch.ghostRaw);
        if (ghost.master.length) {
          const masterData = ghost.master;
          const matchCol = (col === CL_COL_OSIS) ? 0 : (col === CL_COL_STUDENT) ? 1 : 2;
          const needle = editedValue.toLowerCase();
          let match = null;
          for (let i = 0; i < masterData.length; i++) {
            if (String(masterData[i][matchCol]).trim().toLowerCase() === needle) { match = masterData[i]; break; }
          }

          if (match) {
            const correctOsis = match[0];
            const correctStudent = match[1];
            const correctGuardian = match[2];

            // --- PULL SITE & CLASS FROM DIRECTORY (HANDLES SIBLINGS) ---
            const sites = new Set();
            const classes = new Set();

            if (correctOsis) {
              const osisSet = new Set(String(correctOsis).split(',').map(o => o.trim()).filter(String));
              // OSIS/Site/Class (Directory C:E), already read with the Master Table
              {
                const dirData = ghost.dir;
                for (let i = 0; i < dirData.length; i++) {
                  if (osisSet.has(String(dirData[i][0]).trim())) {
                    const siteVal = String(dirData[i][1]).trim();
                    const classVal = String(dirData[i][2]).trim();
                    if (siteVal) sites.add(siteVal);
                    if (classVal) classes.add(classVal);
                  }
                }
              }
            }

            const correctSite = Array.from(sites).join(", ");
            const correctClass = Array.from(classes).join(", ");

            // Write only the cells it fills (no read of the row first: writes
            // wait for the flush below, a read would cost a round trip).
            // D (Person Spoke With) and G (#) are not touched.
            if (col !== CL_COL_GUARDIAN) sheet.getRange(row, CL_COL_GUARDIAN).setValue(correctGuardian); // C
            if (col === CL_COL_GUARDIAN) sheet.getRange(row, CL_COL_OSIS, 1, 2).setValues([[correctOsis, correctStudent]]); // E:F
            else if (col === CL_COL_OSIS) sheet.getRange(row, CL_COL_STUDENT).setValue(correctStudent);  // F
            else sheet.getRange(row, CL_COL_OSIS).setValue(correctOsis);                                 // E
            sheet.getRange(row, 8, 1, 2).setValues([[correctSite, correctClass]]);                       // H:I
            SpreadsheetApp.flush(); // Paint the filled values immediately

          } else {
            e.source.toast("Could not find '" + editedValue + "' in the Master Table.", "Ghost Typist Failed", 4);
          }
        }
      }

      // =======================================================
      // AUTO-FILL DATE / # INTERACTIONS / FOLLOW-UP STATUS
      // Fires on the same single-cell edits. Fill-if-blank only, and skipped
      // entirely when "Disable Contact Log Auto-Fill" is on in Settings. The
      // identity lookup above always runs regardless of that setting.
      // =======================================================
      autofillContactLogDefaults_(sheet, row, col, endRow, editedValue);
    } catch (err) {
      // Surface ghost-typist/auto-fill failures instead of dying silently
      // mid-edit; the repair and delta sync below still get a chance to run.
      console.error(err);
      e.source.toast("Auto-fill failed: " + err.message, "⚠️ Ghost Typist Error", 5);
    }
  }

  // --- NOTES FIRST ---
  // A single-cell edit whose row already has an OSIS rewrites that student's
  // Combined Contact Tracking notes NOW, right after the ghost typist, so the
  // search bar's Communications formula shows them as early as possible. The
  // row's borders / checkbox / END bar clean-up below can wait.
  let notesDone = false;
  if (numRows === 1 && range.getNumColumns() === 1 && (!endRow || row < endRow)) {
    const earlyOsis = String(sheet.getRange(row, CL_COL_OSIS).getDisplayValue()).trim();
    if (isValidOsis_(earlyOsis)) notesDone = syncNotesForRowOsis_(earlyOsis);
  }

  // --- SEARCH BAR COMMUNICATIONS (only when drawn by code, see the section below) ---
  // Right after the ghost typist: if the search bar shows a student on the
  // edited row(s), redraw its notes now (built from the Contact Log itself, so
  // it does not wait for the Combined Contact Tracking update further down).
  if (isSearchBarCodeMode_()) {
    const lastEdited = endRow ? Math.min(row + numRows - 1, endRow - 1) : row + numRows - 1;
    if (lastEdited >= row) {
      const editedOsis = sheet.getRange(row, CL_COL_OSIS, lastEdited - row + 1, 1).getDisplayValues().map(r => r[0]);
      // An OSIS typed over / cleared: the student it used to be may be the one shown
      if (col === CL_COL_OSIS && e.oldValue !== undefined) editedOsis.push(String(e.oldValue));
      refreshSearchBarCommunications_(sheet, endRow, editedOsis);
    }
  }

  // --- SAFE ROW CEILING ---
  let safeNumRows = numRows;
  if (endRow && (row + safeNumRows) > endRow) {
    safeNumRows = endRow - row;
  }

  if (safeNumRows > 0) {
    // --- INSTANT ROW REPAIR ---
    sheet.getRange(row, 1, safeNumRows, CL_WIDTH).setBackground(null);
    formatContactLogRows_(sheet, row, safeNumRows);
    sheet.getRange(row, CL_COL_SPOKE_WITH, safeNumRows, 1).setFontFamily("Roboto");
  }

  autoAddBlankRow_(sheet, row, endRow);
  if (endRow && row === endRow - 1) rememberEndRow_(sheet, endRow + 1); // a row was added above END
  patchMissingFormulasInRow_(sheet, row);

  // --- LAG FIX: Only recalculate colors if you actually changed the Type (Col L) ---
  if (col === CL_COL_TYPE) {
    applyTypeDropdownColors();
  }

  // ==================================================
  // --- FORCE END BAR FORMATTING & SWEEP ---
  // ==================================================
  let finalEndRow = findEndRowInColumnE_(sheet);
  rememberEndRow_(sheet, finalEndRow);
  if (finalEndRow) {
    sheet.getRange(finalEndRow, 1, 1, CL_WIDTH).setBorder(true, null, true, null, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID);

    let maxRows = sheet.getMaxRows();
    if (maxRows >= finalEndRow) {
      let rowsToClear = maxRows - finalEndRow + 1;
      sheet.getRange(finalEndRow, CL_COL_CHECKBOX, rowsToClear, 1).clearDataValidations().clearContent();
    }
  }

  SpreadsheetApp.flush(); 

  if (notesDone) return; // already synced right after the ghost typist (above)

  // --- DYNAMIC DELTA SYNC TRIGGER ---
  let targetOsis = String(sheet.getRange(row, CL_COL_OSIS).getDisplayValue()).trim();

  if (!isValidOsis_(targetOsis)) {
    Utilities.sleep(1000);
    SpreadsheetApp.flush();
    targetOsis = String(sheet.getRange(row, CL_COL_OSIS).getDisplayValue()).trim();
  }

  if (isValidOsis_(targetOsis)) {
    const props = PropertiesService.getDocumentProperties();
    const autoNotes = props.getProperty('autoNotes') !== 'false'; 
    
    if (autoNotes) {
      syncNotesForRowOsis_(targetOsis);
    } else {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "Note saved locally. Background sync is paused via Settings.", 
        "⚡ Fast Entry Mode", 
        3
      );
    }

  } else {
    if (col === CL_COL_OSIS && !editedValue) {
      updateContactAndPhoneOnly();
    }
  }
}

// =========================================================
// CONTACT LOG SEARCH BAR — COMMUNICATIONS (drawn by code, owner's request 2026-11-05)
// The search bar sits a few rows below the END bar. Its "Communications" box
// was one formula in column D (HSTACK("Communications", notes)) reading
// Combined Contact Tracking, so a new Contact Log row only showed there after
// the quick update had rewritten Combined. Now the script writes the label in D
// and the notes in E itself, built straight from the Contact Log / Events /
// Notes data (notesTextForOsis_, the same text Combined gets), right after the
// ghost typist fills a row: no waiting for Combined. Display only.
// Same rules as the old formula:
//   - OSIS = column E of the row ABOVE the label (the search bar's OSIS box);
//     a "111, 222" group shows each student's notes, joined by a blank line
//   - "▶ Name" (Master Table B) above each student's notes only when F two
//     rows above the label is not blank
//   - a leading "Notes:" removed, spaces trimmed; nothing found = "No Notes";
//     empty OSIS box = blank
// Redrawn on: an edit in the search bar (below END); a Contact Log edit for a
// student the box is showing (right after the ghost typist); Notes / Events
// edits; every Build. To go back to the formula, run
// restoreSearchBarCommunicationsFormula from the Apps Script editor
// (useCodeForSearchBarCommunications switches back).
// =========================================================
const SEARCH_BAR_COMMS_LABEL = "Communications";
const SEARCH_BAR_SCAN_ROWS = 15;                        // rows below END searched for the label
const SEARCH_BAR_COMMS_PROP = "SEARCH_BAR_COMMS_MODE";  // "code" = drawn by the script; anything else = the formula (default)

/** True when the script draws the box (useCodeForSearchBarCommunications); default is the formula. */
function isSearchBarCodeMode_() {
  return PropertiesService.getDocumentProperties().getProperty(SEARCH_BAR_COMMS_PROP) === "code";
}
const SHEETS_CELL_LIMIT = 50000;

/**
 * The "Communications" label cell below the END bar as { row, col }, or null.
 * Any column (the box's position comes from the sheet, not the code). Also
 * matches the old HSTACK formula when it shows an error instead of the label
 * (e.g. #REF! because text left in the cell beside it blocks its spill).
 */
function findSearchBarComms_(sheet, endRow) {
  const first = endRow + 1;
  const count = Math.min(SEARCH_BAR_SCAN_ROWS, sheet.getMaxRows() - endRow);
  if (count < 1) return null;
  const area = sheet.getRange(first, 1, count, sheet.getMaxColumns());
  const shown = area.getDisplayValues();
  const formulas = area.getFormulas();
  for (let i = 0; i < shown.length; i++) {
    for (let j = 0; j < shown[i].length; j++) {
      if (String(shown[i][j]).trim() === SEARCH_BAR_COMMS_LABEL ||
          formulas[i][j].indexOf('"' + SEARCH_BAR_COMMS_LABEL + '"') !== -1) {
        return { row: first + i, col: j + 1 };
      }
    }
  }
  return null;
}

/** Same as Sheets' TRIM: drops leading/trailing spaces and collapses repeated spaces. */
function sheetsTrim_(text) {
  return String(text).replace(/ {2,}/g, " ").replace(/^ +| +$/g, "");
}

/**
 * The search bar's box: { row, col, osisBox, showNames }, or null when there is
 * none. Same cells as the old formula, relative to the label: OSIS = 1 row up,
 * 1 column right; the name switch = 2 rows up, 2 columns right.
 */
function readSearchBarBox_(sheet, endRow) {
  if (!endRow) { console.warn("Search bar: no END row found in Contact Log column E."); return null; }
  const at = findSearchBarComms_(sheet, endRow);
  if (!at) { console.warn("Search bar: no \"Communications\" cell in the " + SEARCH_BAR_SCAN_ROWS + " rows below END."); return null; }
  if (at.row - 2 <= endRow) { console.warn("Search bar: \"Communications\" is too close to END (row " + at.row + ")."); return null; }
  SpreadsheetApp.flush(); // the OSIS box is a formula: let it catch up with the edit first
  const v = sheet.getRange(at.row - 2, at.col + 1, 2, 2).getDisplayValues();
  return { row: at.row, col: at.col, osisBox: String(v[1][0]), showNames: String(v[0][1]) !== "" };
}

/** The OSIS numbers in an OSIS box / cell ("111, 222" -> ["111", "222"]). */
function searchBarOsisList_(text) {
  return String(text).split(",").map(o => sheetsTrim_(o)).filter(o => o !== "");
}

/**
 * One student's notes, read from Combined Contact Tracking (A = OSIS, D = notes):
 * the fast path for a search bar lookup, when Combined is already up to date.
 * Returns a function osis -> notes ("" when the student has no row).
 */
function combinedNotesReader_(ss) {
  const sh = ss.getSheetByName("Combined Contact Tracking");
  const last = sh ? sh.getLastRow() : 0;
  const ids = last >= 2 ? sh.getRange(2, 1, last - 1, 1).getDisplayValues() : [];
  return osis => {
    const key = String(osis).trim().toLowerCase();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim().toLowerCase() === key) return String(sh.getRange(i + 2, 4).getDisplayValue());
    }
    return "";
  };
}

/** One student's notes built from the Contact Log / Event Log / Notes data (never stale). */
function logsNotesReader_(ss, contactSheet, endRow) {
  const disableHiddenNotes = PropertiesService.getDocumentProperties().getProperty('disableHiddenNotes') === 'true';
  // Read each source once (Contact Log only down to the END bar)
  const contactData = contactSheet.getRange(1, 1, endRow, CL_WIDTH).getDisplayValues();
  const eventData = backendReadBlock_(ss, "EVENT_LOG", true);
  const notesSheet = ss.getSheetByName("Notes");
  const nData = notesSheet ? notesSheet.getDataRange().getValues() : null;
  return osis => notesTextForOsis_(osis, contactData, eventData, nData, disableHiddenNotes);
}

/**
 * The Communications text for the box (same rules as the old formula).
 * fromCombined: read the notes from Combined Contact Tracking (quickest; used for
 * a search bar lookup) instead of building them from the logs (used right after
 * a Contact Log edit, before the quick update has written Combined).
 */
function buildSearchBarCommunications_(ss, contactSheet, endRow, osisBox, showNames, fromCombined) {
  if (String(osisBox) === "") return "";
  const osisList = searchBarOsisList_(osisBox);
  const notesFor = fromCombined ? combinedNotesReader_(ss) : logsNotesReader_(ss, contactSheet, endRow);

  let names = null;
  if (showNames) {
    names = new Map();
    const master = ss.getSheetByName("Master Table");
    const last = master ? master.getLastRow() : 0;
    if (last >= 5) {
      master.getRange(5, 1, last - 4, 2).getDisplayValues().forEach(r => {
        const k = String(r[0]).trim().toLowerCase();
        if (k && !names.has(k)) names.set(k, r[1]);
      });
    }
  }

  const blocks = [];
  osisList.forEach(osis => {
    const notes = notesFor(osis);
    if (notes === "") return;
    if (!names) { blocks.push(notes); return; }
    const name = names.get(osis.toLowerCase());
    blocks.push("▶ " + (name === undefined ? "Unknown" : name) + "\n" + notes);
  });

  const n = blocks.join("\n\n");
  const cleaned = sheetsTrim_(n.replace(/^notes:\s*[\r\n]*/i, ""));
  if (n === "" || n === "0" || cleaned === "" || cleaned === "0") return "No Notes";
  // A cell holds at most 50,000 characters (the formula showed an error past that)
  return cleaned.length > SHEETS_CELL_LIMIT ? cleaned.slice(0, SHEETS_CELL_LIMIT - 1) + "…" : cleaned;
}

/**
 * Redraws the search bar's Communications box. Never throws (it runs inside edits).
 * onlyForOsis: optional list of OSIS values; the box is only redrawn when it
 * shows one of them (a Contact Log row edit for another student skips it).
 * fromCombined: true = read the notes from Combined Contact Tracking (a lookup
 * or after a build, when Combined is current); otherwise built from the logs.
 */
function refreshSearchBarCommunications_(sheet, endRow, onlyForOsis, fromCombined) {
  try {
    if (!isSearchBarCodeMode_()) return; // the formula draws it
    sheet = sheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Contact Log");
    if (!sheet) return;
    const ss = sheet.getParent();
    endRow = endRow || findEndRowInColumnE_(sheet);
    const box = readSearchBarBox_(sheet, endRow);
    if (!box) return;

    if (onlyForOsis) {
      const shown = new Set(searchBarOsisList_(box.osisBox));
      const wanted = [].concat(...onlyForOsis.map(searchBarOsisList_));
      if (!wanted.some(o => shown.has(o))) return;
    }

    // Combined is only current when quick updates are on (Settings "autoNotes")
    const useCombined = fromCombined === true &&
      PropertiesService.getDocumentProperties().getProperty('autoNotes') !== 'false';
    const text = buildSearchBarCommunications_(ss, sheet, endRow, box.osisBox, box.showNames, useCombined);

    // Label first (this replaces the old formula, whose spill would block E), then the notes
    const label = sheet.getRange(box.row, box.col);
    if (label.getFormula() !== "" || label.getValue() !== SEARCH_BAR_COMMS_LABEL) label.setValue(SEARCH_BAR_COMMS_LABEL);
    const cell = sheet.getRange(box.row, box.col + 1);
    if (cell.getNumberFormat() !== "@") cell.setNumberFormat("@"); // plain text: a note is never read as a number / formula
    if (cell.getFormula() !== "" || cell.getDisplayValue() !== text) cell.setValue(text);
  } catch (err) {
    console.error("Search bar Communications not updated: " + err.message);
  }
}

/**
 * Run from the Apps Script editor to see why the box is (not) filled: redraws it
 * and toasts what it found, or what stopped it.
 */
function checkSearchBarCommunications() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const say = (msg) => { ss.toast(msg, "Search Bar check", 15); console.log(msg); };
  if (!isSearchBarCodeMode_()) {
    return say("The formula draws Communications (the default). Run useCodeForSearchBarCommunications to let the script draw it.");
  }
  const sheet = ss.getSheetByName("Contact Log");
  if (!sheet) return say("No Contact Log tab.");
  const endRow = findEndRowInColumnE_(sheet);
  if (!endRow) return say("No END row found in Contact Log column E.");
  const at = findSearchBarComms_(sheet, endRow);
  if (!at) return say("END is row " + endRow + ", but no \"Communications\" cell in the " + SEARCH_BAR_SCAN_ROWS + " rows below it.");
  const box = readSearchBarBox_(sheet, endRow);
  if (!box) return say("\"Communications\" at row " + at.row + " is too close to END (row " + endRow + ").");
  const osisCell = sheet.getRange(box.row - 1, box.col + 1).getA1Notation();
  if (box.osisBox === "") return say("Found the box at " + sheet.getRange(box.row, box.col).getA1Notation() + ", but the OSIS cell " + osisCell + " is empty.");
  refreshSearchBarCommunications_(sheet, endRow, null, true);
  const out = sheet.getRange(box.row, box.col + 1);
  say("OSIS " + box.osisBox + " (from " + osisCell + ") -> wrote " + String(out.getDisplayValue()).length + " characters into " + out.getA1Notation() + ".");
}

/** Puts the original Communications formula back and stops the script from drawing it. */
function restoreSearchBarCommunicationsFormula() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Contact Log");
  const endRow = sheet ? findEndRowInColumnE_(sheet) : null;
  const at = endRow ? findSearchBarComms_(sheet, endRow) : null;
  if (!at) { ss.toast("Could not find the Communications box below END.", "Search Bar", 5); return; }

  PropertiesService.getDocumentProperties().setProperty(SEARCH_BAR_COMMS_PROP, "formula");
  sheet.getRange(at.row, at.col + 1).clearContent();
  sheet.getRange(at.row, at.col).setFormula(SEARCH_BAR_COMMS_FORMULA);
  ss.toast("Communications is a formula again.", "Search Bar", 5);
}

/** Lets the script draw Communications again (undoes restoreSearchBarCommunicationsFormula). */
function useCodeForSearchBarCommunications() {
  PropertiesService.getDocumentProperties().setProperty(SEARCH_BAR_COMMS_PROP, "code");
  refreshSearchBarCommunications_();
  SpreadsheetApp.getActiveSpreadsheet().toast("Communications is drawn by the script.", "Search Bar", 5);
}

/** The original formula (only used to put it back). */
const SEARCH_BAR_COMMS_FORMULA = `=LET(
  targetE, INDIRECT(ADDRESS(ROW()-1, COLUMN()+1)),
  targetF, INDIRECT(ADDRESS(ROW()-2, COLUMN()+2)),
  HSTACK(
    "Communications",
    IF(targetE="","", LET(
      osisArray, SPLIT(targetE, ",", FALSE, TRUE),
      allNotes, MAP(osisArray, LAMBDA(osis, 
         LET(
           oText, TO_TEXT(TRIM(osis)),
           oNum, IFERROR(VALUE(oText), oText),
           
           name, IFNA(VLOOKUP(oNum, 'Master Table'!A:B, 2, FALSE), IFNA(VLOOKUP(oText, 'Master Table'!A:B, 2, FALSE), "Unknown")),
           notes, IFNA(VLOOKUP(oNum, 'Combined Contact Tracking'!A:D, 4, FALSE), IFNA(VLOOKUP(oText, 'Combined Contact Tracking'!A:D, 4, FALSE), "")),
           
           IF(notes="", "", IF(targetF="", notes, "▶ " & name & CHAR(10) & notes))
         )
      )),
      n, TEXTJOIN(CHAR(10) & CHAR(10), TRUE, allNotes),
      cleaned, TRIM(REGEXREPLACE(n,"(?i)^notes:\\s*[\\r\\n]*","")),
      IF(OR(n="", n="0", cleaned="", cleaned="0"), "No Notes", cleaned)
    ))
  )
)`;

/**
 * The quick notes update for one Contact Log row's OSIS (a sibling group
 * "111, 222" updates the group row and then each sibling). Returns true when it
 * ran, false when "update notes as I type" is off in Settings (the caller then
 * shows the Fast Entry Mode toast at the end of the edit).
 */
function syncNotesForRowOsis_(targetOsis) {
  if (PropertiesService.getDocumentProperties().getProperty('autoNotes') === 'false') return false;
  const osisList = [targetOsis];
  if (targetOsis.includes(",")) {
    targetOsis.split(",").map(o => o.trim()).forEach(individualOsis => {
      if (isValidOsis_(individualOsis)) osisList.push(individualOsis);
    });
  }
  updateOsisDeltas_(osisList);
  return true;
}

// =========================================================
// GHOST TYPIST DATA CACHE
// The Contact Log ghost typist needs Master Table A5:C and Directory C:E on
// every edit. Only a build writes those tabs, so the first edit reads them and
// keeps a copy in the document cache for 10 minutes; the next edits skip both
// reads. Every build clears it (clearGhostTypistCache_). Too big for the cache
// (> ~95 KB): read every time, as before.
// =========================================================
const GHOST_CACHE_KEY_ = "GHOST_TYPIST_DATA";
const GHOST_CACHE_SECONDS_ = 600;

/** ghostRaw: the cached JSON already fetched by contactLogPrefetch_ (skips a cache call), or undefined. */
function ghostTypistData_(ss, ghostRaw) {
  let cache = null;
  try {
    cache = CacheService.getDocumentCache();
    const hit = ghostRaw !== undefined ? ghostRaw : (cache && cache.get(GHOST_CACHE_KEY_));
    if (hit) return JSON.parse(hit);
  } catch (err) { console.warn("Ghost typist cache not read: " + err.message); }

  const master = ss.getSheetByName("Master Table");
  const mLast = master ? master.getLastRow() : 0;
  const dir = ss.getSheetByName("Directory");
  const dLast = dir ? dir.getLastRow() : 0;
  const data = {
    master: mLast >= 5 ? master.getRange(5, 1, mLast - 4, 3).getValues() : [],
    dir: dLast > 0 ? dir.getRange(1, 3, dLast, 3).getValues() : []
  };
  try {
    const json = JSON.stringify(data);
    if (cache && json.length < 95000) cache.put(GHOST_CACHE_KEY_, json, GHOST_CACHE_SECONDS_);
  } catch (err) { console.warn("Ghost typist cache not saved: " + err.message); }
  return data;
}

/** Cache key of the Contact Log's END row (per tab, since a copy of the tab has its own id). */
function endRowCacheKey_(sheet) {
  return "CL_END_ROW_" + sheet.getSheetId();
}

/** Remembers where END is for the next edit (6 hours). */
function rememberEndRow_(sheet, endRow) {
  if (!endRow) return;
  try { CacheService.getDocumentCache().put(endRowCacheKey_(sheet), String(endRow), 21600); } catch (err) { /* next edit just scans */ }
}

/**
 * Start of a Contact Log edit: { endRow, ghostRaw } with as few round trips as
 * possible. One cache call returns the remembered END row + the ghost typist's
 * cached data; the END row is trusted only after a one-cell check that column
 * E still says END there, else column E is scanned (and remembered).
 */
function contactLogPrefetch_(sheet) {
  const endKey = endRowCacheKey_(sheet);
  let got = {};
  try { got = CacheService.getDocumentCache().getAll([endKey, GHOST_CACHE_KEY_]) || {}; } catch (err) { got = {}; }

  let endRow = null;
  const remembered = parseInt(got[endKey], 10);
  if (remembered > 1) {
    try {
      if (String(sheet.getRange(remembered, CL_COL_OSIS).getDisplayValue()).trim().toUpperCase() === "END") endRow = remembered;
    } catch (err) { /* row no longer exists: scan below */ }
  }
  if (!endRow) {
    endRow = findEndRowInColumnE_(sheet);
    rememberEndRow_(sheet, endRow);
  }
  return { endRow: endRow, ghostRaw: got[GHOST_CACHE_KEY_] || null };
}

function clearGhostTypistCache_() {
  try { CacheService.getDocumentCache().remove(GHOST_CACHE_KEY_); } catch (err) { console.warn("Ghost typist cache not cleared: " + err.message); }
}

// =========================================================
// HELPER FUNCTIONS FOR TRIGGERS
// =========================================================
function isValidOsis_(value) {
  if (!value) return false;
  const strVal = String(value);
  if (strVal.includes("#")) return false; 
  if (strVal.toUpperCase() === "END") return false;
  if (strVal.toLowerCase().includes("loading")) return false;
  return true;
}

// =========================================================
// CONTACT LOG ROW HELPERS
// =========================================================

/** Column A "week of" formula for one Contact Log row (the Sunday of B's week). */
function weekFormula_(row) {
  return `=IF(B${row}="", "", B${row} - WEEKDAY(B${row}, 1) + 1)`;
}

function patchMissingFormulasInRow_(sheet, row) {
  const range = sheet.getRange(row, 1, 1, 1);
  const values = range.getDisplayValues()[0];
  const formulas = range.getFormulas()[0];

  if (String(values[0]).trim() === "" || String(values[0]).indexOf("#") === 0 || formulas[0] === "") {
    sheet.getRange(row, 1).setFormula(weekFormula_(row));
  }
}

function FIX_ALL_MISSING_FORMULAS_BATCH() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Contact Log");
  if (!sheet) return;

  const endRow = findEndRowInColumnE_(sheet);
  if (!endRow || endRow <= 2) {
    ss.toast("Could not find END in Column E or table is empty.", "Error", 5);
    return;
  }

  const numRows = endRow - 2;
  const startRow = 2;
  const colAFormulas = [];

  for (let i = 0; i < numRows; i++) {
    colAFormulas.push([weekFormula_(startRow + i)]);
  }

  sheet.getRange(startRow, 1, numRows, 1).setFormulas(colAFormulas);
  formatContactLogRows_(sheet, startRow, numRows);

  ss.toast("Week formulas restored and borders fixed!", "Success", 3);
}

function autoAddBlankRow_(sheet, editedRow, endRow) {
  if (editedRow === endRow - 1) {
    sheet.insertRowAfter(editedRow);
    formatContactLogRows_(sheet, editedRow + 1, 1);
    sheet.getRange(editedRow + 1, CL_COL_CHECKBOX).setValue(false);

    // The end bar's black border sticks to the edited row's bottom edge when
    // the new row is inserted above it. Strip its horizontal borders so only
    // the dark brown end bar (re-drawn by the onEdit sweep) keeps a border.
    sheet.getRange(editedRow, 1, 1, CL_WIDTH).setBorder(null, null, false, null, null, false);
  }
}

function applyTypeDropdownColors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const contactLog = ss.getSheetByName('Contact Log');
  // Type List section of the Backend tab (BL2:BL)
  const typeValues = backendReadBlock_(ss, "TYPES", true).flat().filter(String);
  const typeColumnRange = contactLog.getRange('L2:L'); 
  let existingRules = contactLog.getConditionalFormatRules();
  
  let typeRules = [];   
  let otherRules = [];  
  
  let existingTypesMapped = []; 
  let usedColors = []; 
  
  existingRules.forEach(rule => {
    const ranges = rule.getRanges();
    if (ranges.length === 0) {
      otherRules.push(rule);
      return;
    }
    
    const firstCol = ranges[0].getColumn();
    const lastCol = ranges[0].getLastColumn();
    
    if (firstCol === 12 && lastCol === 12) {
      try {
        const condition = rule.getBooleanCondition();
        if (condition && condition.getCriteriaType() === SpreadsheetApp.BooleanCriteria.TEXT_EQUAL_TO) {
          const textBeingColored = condition.getCriteriaValues()[0];
          if (typeValues.includes(textBeingColored)) {
             existingTypesMapped.push(textBeingColored);
             typeRules.push(rule); 
             const ruleColor = condition.getFontColor(); 
             if (ruleColor && !usedColors.includes(ruleColor.toLowerCase())) {
               usedColors.push(ruleColor.toLowerCase());
             }
          }
        }
      } catch(e) { console.error(e); }
    } else {
      otherRules.push(rule);
    }
  });
  
  function generateVibrantHex() {
    let rgb = [
      Math.floor(Math.random() * 40),       
      Math.floor(Math.random() * 60) + 140, 
      Math.floor(Math.random() * 200)       
    ];
    for (let i = rgb.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [rgb[i], rgb[j]] = [rgb[j], rgb[i]];
    }
    return "#" + rgb.map(x => x.toString(16).padStart(2, '0')).join('');
  }
  
  typeValues.forEach(type => {
    if (!existingTypesMapped.includes(type)) {
      let colorToUse = generateVibrantHex();
      while (usedColors.includes(colorToUse.toLowerCase())) {
        colorToUse = generateVibrantHex();
      }
      const newRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(type)
        .setFontColor(colorToUse) 
        .setRanges([typeColumnRange])
        .build();
      typeRules.push(newRule); 
      existingTypesMapped.push(type);
      usedColors.push(colorToUse.toLowerCase());
    }
  });
  
  const finalRulesToApply = [...typeRules, ...otherRules];
  contactLog.setConditionalFormatRules(finalRulesToApply);
}

function toggleHideFilter_(sheet, endRow) {
  if (!endRow || endRow < 3) return;
  const isChecked = sheet.getRange(endRow, 15).getValue();
  
  let filter = sheet.getFilter();
  if (filter) {
    filter.remove();
  }
  
  if (isChecked === true || isChecked === 'TRUE') {
    createContactLogHideFilter_(sheet, endRow - 1);
  }
}

/**
 * The Contact Log's master hide filter: rows 1..lastRow, A:P, showing only rows
 * whose Column P checkbox is black (the grey ones are the hidden notes).
 */
function createContactLogHideFilter_(sheet, lastRow) {
  const filter = sheet.getRange(1, 1, lastRow, 16).createFilter();
  const blackColor = SpreadsheetApp.newColor().setRgbColor("#000000").build();
  const criteria = SpreadsheetApp.newFilterCriteria()
    .setVisibleForegroundColor(blackColor) 
    .build();
  filter.setColumnFilterCriteria(16, criteria);
  return filter;
}

// =========================================================
// NEW: CONTACT LOG AUTO-FILL (Date / # of interactions / Follow-up status)
// =========================================================
// When a coordinator enters real interaction content into a fresh row, stamp the
// three routine defaults so they don't have to type them every time:
//   B (2)  = today's date (a real midnight Date, so the week formula in Col A and
//            the "today = green" / week-band conditional formatting all work)
//   G (7)  = 1  (# of interactions; this also clears the red "missing" flag)
//   N (14) = "No Followup Needed"  (the follow-up status dropdown's default)
//
// Rules:
//   • FILL-IF-BLANK: never overwrite a cell the user already filled. Back-dating a
//     date, setting G to 3, or choosing "Followup Needed - ASAP" in N is all preserved.
//   • Skipped entirely when 'disableContactAutofill' is set in Settings. The
//     separate parent/student/OSIS lookup is NOT affected by that switch.
//   • Only fires on genuine single-cell content edits inside the live table
//     (the caller already restricts to single-cell edits above the END row).
function autofillContactLogDefaults_(sheet, row, col, endRow, editedValue) {
  const props = PropertiesService.getDocumentProperties();
  if (props.getProperty('disableContactAutofill') === 'true') return;

  if (!endRow || row < 2 || row >= endRow) return;

  // Columns that represent a real interaction being entered:
  // Guardian(3), Person Spoke With(4), OSIS(5), Student(6),
  // Method 1(10), Method 2(11), Type(12), Event Description(13), Additional Notes(15).
  const CONTENT_COLS = [3, 4, 5, 6, 10, 11, 12, 13, 15];
  if (CONTENT_COLS.indexOf(col) === -1) return;

  // Ignore blank edits and checkbox toggles so clearing a cell never stamps a row.
  const val = String(editedValue == null ? "" : editedValue).trim().toUpperCase();
  if (val === "" || val === "TRUE" || val === "FALSE") return;

  // ONE read covers B (date), G (count), and N (follow-up) instead of three.
  const rowVals = sheet.getRange(row, 2, 1, 13).getDisplayValues()[0]; // B..N

  // B (2): Date — only if empty.
  if (String(rowVals[0]).trim() === "") {
    const today = new Date();
    today.setHours(0, 0, 0, 0); // midnight => $B2 = TODAY() and Col A week math both work
    sheet.getRange(row, 2).setValue(today);
  }

  // G (7): # of interactions — only if empty.
  if (String(rowVals[5]).trim() === "") sheet.getRange(row, 7).setValue(1);

  // N (14): Follow-up status dropdown — only if empty (never flip a chosen value).
  if (String(rowVals[12]).trim() === "") sheet.getRange(row, 14).setValue("No Followup Needed");
}

// =========================================================
// PASTE GHOST TYPIST HELPERS
// =========================================================

/**
 * Reads Master Table A5:C ONCE and returns lookup(value, col) for pasted values,
 * col 0 = OSIS, 1 = Student, 2 = Guardian. Case/spacing don't matter, and a
 * single OSIS or parent name matches a sibling group / multi-parent cell.
 */
function buildGhostLookup_(ss) {
  const master = ss.getSheetByName("Master Table");
  const last = master ? master.getLastRow() : 0;
  const data = last >= 5 ? master.getRange(5, 1, last - 4, 3).getValues() : [];
  const key = v => String(v).trim().toLowerCase().replace(/\s+/g, " ");
  const byCol = [new Map(), new Map(), new Map()];
  const add = (c, k, rec) => { if (k && !byCol[c].has(k)) byCol[c].set(k, rec); };

  data.forEach(r => {
    const rec = { osis: r[0], student: r[1], guardian: r[2] };
    for (let c = 0; c < 3; c++) add(c, key(r[c]), rec);
    String(r[0]).split(",").forEach(o => add(0, key(o), rec));
    String(r[1]).split(",").forEach(n => add(1, key(n), rec));
    String(r[2]).split(/,|\s+and\s+/i).forEach(n => add(2, key(stripTrailingParenthetical_(n)), rec)); // "Homer Simpson and Marge Simpson"
  });

  return (value, col) => byCol[col].get(key(value)) || byCol[col].get(key(stripTrailingParenthetical_(value))) || null;
}

/** Reads Directory C:E ONCE; returns siteClass(osisGroup) -> { site, cls } like the single-cell typist. */
function buildDirectorySiteClassLookup_(ss) {
  const byOsis = new Map();
  const dir = ss.getSheetByName("Directory");
  const last = dir ? dir.getLastRow() : 0;
  if (last > 0) {
    dir.getRange(1, 3, last, 3).getValues().forEach(r => {
      const o = String(r[0]).trim();
      if (!o) return;
      if (!byOsis.has(o)) byOsis.set(o, { sites: new Set(), classes: new Set() });
      if (String(r[1]).trim()) byOsis.get(o).sites.add(String(r[1]).trim());
      if (String(r[2]).trim()) byOsis.get(o).classes.add(String(r[2]).trim());
    });
  }
  return osisGroup => {
    const sites = new Set(), classes = new Set();
    String(osisGroup).split(",").map(o => o.trim()).filter(Boolean).forEach(o => {
      const hit = byOsis.get(o);
      if (hit) { hit.sites.forEach(x => sites.add(x)); hit.classes.forEach(x => classes.add(x)); }
    });
    return { site: Array.from(sites).join(", "), cls: Array.from(classes).join(", ") };
  };
}

function toastGhostMisses_(ss, filled, missing) {
  if (!missing.length) {
    ss.toast(`Filled ${filled} row${filled === 1 ? "" : "s"}.`, "Ghost Typist", 3);
    return;
  }
  const shown = missing.slice(0, 5).join(", ") + (missing.length > 5 ? ` (+${missing.length - 5} more)` : "");
  ss.toast(`Filled ${filled}. Not found in the Master Table: ${shown}`, "Ghost Typist", 8);
}

/** Contact Log: fills Guardian / OSIS / Student / Site / Class for every pasted row. */
function fillPastedContactLogRows_(ss, sheet, row, numRows, col, endRow) {
  const lookup = buildGhostLookup_(ss);
  const siteClass = buildDirectorySiteClassLookup_(ss);
  const matchCol = (col === CL_COL_OSIS) ? 0 : (col === CL_COL_STUDENT) ? 1 : 2;

  const updateRange = sheet.getRange(row, 3, numRows, 7); // C:I
  const rows = updateRange.getValues();
  const missing = [];
  const filledRows = [];

  rows.forEach((r, i) => {
    const val = String(r[col - 3]).trim();
    if (!val || val.toUpperCase() === "TRUE" || val.toUpperCase() === "FALSE") return;
    const m = lookup(val, matchCol);
    if (!m) { missing.push(val); return; }
    const sc = siteClass(m.osis);
    if (col !== CL_COL_GUARDIAN) r[0] = m.guardian; // C
    if (col !== CL_COL_OSIS) r[2] = m.osis;         // E
    if (col !== CL_COL_STUDENT) r[3] = m.student;   // F
    r[5] = sc.site;                                 // H
    r[6] = sc.cls;                                  // I
    filledRows.push({ row: row + i, val });
  });

  updateRange.setValues(rows);
  SpreadsheetApp.flush();

  // Date / # / Follow-up defaults and the week formula for each filled row
  filledRows.forEach(f => {
    autofillContactLogDefaults_(sheet, f.row, col, endRow, f.val);
    patchMissingFormulasInRow_(sheet, f.row);
  });
  // Keep a blank row above the END bar when the paste reached it
  autoAddBlankRow_(sheet, row + numRows - 1, endRow);

  toastGhostMisses_(ss, filledRows.length, missing);
}

/** Notes: fills OSIS / Guardian / Student for every pasted row; families already on the sheet are removed. */
function fillPastedNotesRows_(ss, sheet, row, numRows, col) {
  const lookup = buildGhostLookup_(ss);
  const matchCol = (col === 1) ? 0 : (col === 2) ? 2 : 1; // A=OSIS, B=Guardian, C=Student

  const block = sheet.getRange(row, 1, numRows, 3).getValues();

  // OSIS already logged outside the pasted rows
  const seen = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 1).getValues().forEach((r, i) => {
      const rr = i + 2;
      const o = String(r[0]).trim();
      if (o && (rr < row || rr >= row + numRows)) seen.add(o);
    });
  }

  const missing = [];
  const dupes = [];
  const filledRows = [];

  block.forEach((r, i) => {
    const val = String(r[col - 1]).trim();
    if (!val || val.toUpperCase() === "TRUE" || val.toUpperCase() === "FALSE") return;
    const m = lookup(val, matchCol);
    if (!m) { missing.push(val); return; }
    const osis = String(m.osis).trim();
    if (seen.has(osis)) { dupes.push(m.student); block[i] = ["", "", ""]; return; }
    seen.add(osis);
    block[i] = [m.osis, m.guardian, m.student];
    filledRows.push(row + i);
  });

  sheet.getRange(row, 1, numRows, 3).setValues(block);
  filledRows.forEach(r => sheet.getRange(r, 5).insertCheckboxes().setFontColor("#000000"));

  if (dupes.length) {
    try {
      SpreadsheetApp.getUi().alert("⚠️ Duplicates Removed",
        `Already logged in the Notes sheet, so these were removed: ${dupes.join(", ")}`,
        SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (err) {}
  }
  toastGhostMisses_(ss, filledRows.length, missing);
}

// =========================================================
// NEW: NOTES SHEET GHOST TYPIST
// =========================================================
function watchNotesSheet_(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const row = range.getRow();
  const col = range.getColumn();
  
  // Ignore headers. Pastes (and multi-cell deletes) in A:D fill identities and sync notes for every row.
  if (row < 2) return;
  if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) {
    if (col > 4) return;
    const numRows = range.getNumRows();
    if (col <= 3) fillPastedNotesRows_(e.source, sheet, row, numRows, col);
    // A paste that touches Notes (D) syncs every row, so cleared notes are removed too
    const touchesNotes = col + range.getNumColumns() - 1 >= 4;
    syncNotesRows_(sheet, row, numRows, !touchesNotes);
    return;
  }

  // 🛡️ BULLETPROOF CHECKBOX DETECTION:
  const rawValue = range.getValue(); 
  const isChecked = (rawValue === true || String(rawValue).toUpperCase() === "TRUE");
  const editedValue = String(e.value !== undefined ? e.value : range.getDisplayValue()).trim();

  // =========================================================
  // --- ROW DELETION WITH CONFIRMATION (COLUMN E / 5) ---
  // =========================================================
  if (col === 5 && isChecked) {
    let isConfirmed = false;
    try {
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
        "⚠️ Delete Note", 
        "Are you sure you want to permanently delete this row?", 
        ui.ButtonSet.YES_NO
      );
      isConfirmed = (response === ui.Button.YES);
    } catch (err) {
      // Fallback if UI alert is blocked: default to executing the delete safely
      isConfirmed = true; 
    }

    if (isConfirmed) {
      // Capture the OSIS before deleting so we can trigger a sync just in case
      const targetOsis = String(sheet.getRange(row, 1).getDisplayValue()).trim();
      sheet.deleteRow(row);
      
      try { e.source.toast("Note permanently removed.", "🗑️ Note Deleted", 3); } catch(err){}
      
      SpreadsheetApp.flush();
      
      // ✨ NEW: Smart Sibling Check during deletion!
      if (targetOsis) updateOsisDeltas_(targetOsis.split(",").map(o => o.trim()).filter(Boolean));
    } else {
      // If they click NO or close the window, uncheck the box safely!
      range.uncheck();
    }
    return; 
  }

  // =========================================================
  // --- NEW: SYNC WHEN NOTE IS TYPED (COLUMN D / 4) ---
  // =========================================================
  // If user actually edits the manual note, push the update!
  if (col === 4) {
    const targetOsis = String(sheet.getRange(row, 1).getDisplayValue()).trim();
    if (targetOsis) updateOsisDeltas_(targetOsis.split(",").map(o => o.trim()).filter(Boolean));
    return; 
  }

  // =========================================================
  // --- GHOST TYPIST & DUPLICATE CHECKER ---
  // =========================================================
  if (col === 5) return; // Prevent Typist from running on checkbox clicks
  if (!editedValue || editedValue.toUpperCase() === "FALSE" || editedValue.toUpperCase() === "TRUE") return;

  // Assuming Columns are: A (1) = OSIS, B (2) = Guardian(s), C (3) = Student Name
  if (col === 1 || col === 2 || col === 3) {
    const masterSheet = e.source.getSheetByName("Master Table");
    if (masterSheet) {
      // Typed OSIS (A) -> Master Table A, Guardian (B) -> C, Student (C) -> B
      const found = findInMasterTable_(masterSheet, col === 1 ? 1 : col === 2 ? 3 : 2, editedValue);

      if (found) {
        // Pull exact data from Master Table
        const correctOsis = found.osis;
        const correctStudent = found.student;
        const correctGuardian = found.guardian;

        // ⚡ NEW: CHECK FOR DUPLICATES ON THE NOTES SHEET ⚡
        const lastRow = sheet.getLastRow();
        let isDuplicate = false;
        
        if (lastRow >= 2) {
          const existingData = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
          for (let i = 0; i < existingData.length; i++) {
            // Check if OSIS matches AND make sure we aren't looking at the row we are currently typing in
            if ((i + 2) !== row && String(existingData[i]).trim() === String(correctOsis).trim()) {
              isDuplicate = true;
              break;
            }
          }
        }

        if (isDuplicate) {
          const ui = SpreadsheetApp.getUi();
          ui.alert(
            "⚠️ Duplicate Removed", 
            `This family (${correctStudent}) is already logged in the Notes sheet. The duplicate has been removed.`, 
            ui.ButtonSet.OK
          );
          range.clearContent(); // Wipes out the duplicate text they just typed
          return; // Stops the script here
        }

        // Batch write all 3 identity columns into the Notes sheet instantly
        const updateRange = sheet.getRange(row, 1, 1, 3); // Columns A to C
        const rowData = updateRange.getValues();

        if (col !== 1) rowData[0][0] = correctOsis;      // Col 1 (OSIS)
        if (col !== 2) rowData[0][1] = correctGuardian;  // Col 2 (Guardian)
        if (col !== 3) rowData[0][2] = correctStudent;   // Col 3 (Student Name)

        updateRange.setValues(rowData);
        
        // Ensure a checkbox is securely placed in Column E if one isn't there already
        sheet.getRange(row, 5).insertCheckboxes().setFontColor("#000000");

        // A note typed before the student was filled in syncs now that the row has an OSIS
        syncNotesRows_(sheet, row, 1, true);

      } else {
        e.source.toast("Could not find '" + editedValue + "' in the Master Table.", "Ghost Typist Failed", 4);
      }
    }
  }
}

/**
 * Single-cell ghost typist lookup (Notes and Events): finds a typed value in one
 * Master Table column (1 = OSIS, 2 = Student, 3 = Guardian; row 5 down, whole
 * cell, any case) and returns that row's { osis, student, guardian }, or null.
 */
function findInMasterTable_(masterSheet, masterCol, value) {
  const letter = ["A", "B", "C"][masterCol - 1];
  const found = masterSheet.getRange(`${letter}5:${letter}`).createTextFinder(value).matchEntireCell(true).matchCase(false).findNext();
  if (!found) return null;
  const vals = masterSheet.getRange(found.getRow(), 1, 1, 3).getValues()[0];
  return { osis: vals[0], student: vals[1], guardian: vals[2] };
}

/**
 * Pushes Notes rows into Combined Contact Tracking / Phone Contacts.
 * onlyWithNotes: skip rows whose Notes (D) is blank. Many families at once
 * use one full rebuild instead of a delta per student.
 */
function syncNotesRows_(sheet, row, numRows, onlyWithNotes) {
  const ids = new Set();
  sheet.getRange(row, 1, numRows, 4).getValues().forEach(r => {
    if (onlyWithNotes && !String(r[3]).trim()) return;
    String(r[0]).split(",").map(o => o.trim()).filter(Boolean).forEach(o => ids.add(o));
  });
  if (!ids.size) return;
  if (ids.size > 5) {
    updateContactAndPhoneOnly();
  } else {
    updateOsisDeltas_(Array.from(ids));
  }
}

function updateRawDataTimestamp_(range) {
  const rowStart = range.getRow();
  if (rowStart < 2) return; 

  const colStart = range.getColumn();
  const colEnd = range.getLastColumn();
  const sheet = range.getSheet();
  const ss = sheet.getParent();
  
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yy");
  let msgs = [];

  if (colStart <= 31 && colEnd >= 1) {
    sheet.getRange("A1").setValue("ATS - DATE UPDATED: " + today);
    msgs.push("ATS");
  }

  if (colStart <= 40 && colEnd >= 33) {
    sheet.getRange("AG1").setValue("Parent Square - DATE UPDATED: " + today);
    msgs.push("Parent Square");
  }

  if (colStart <= 58 && colEnd >= 42) {
    sheet.getRange("AP1").setValue("NYCSA - DATE UPDATED: " + today);
    msgs.push("NYCSA");
  }

  if (colStart <= 61 && colEnd >= 60) {
    msgs.push("Site/Class Data");
  }

  if (msgs.length > 0) {
    ss.toast("Dates updated for: " + msgs.join(", ") + ". Repainting formatting...", "Timestamp Triggered", 6);
    
    forceVibrantColors_(sheet);
    applyBandingToRawData_(sheet);
  }
}

function cleanRawDataPasteZone_(sheet, range) {
  const startCol = range.getColumn();
  const pastedData = range.getValues();
  if (!pastedData || pastedData.length === 0) return;
  
  const atsHeaders = ["School Year","DBN","Student ID","Last Name","First Name","Race","Gender","Age","IEP Flag","ELL Flag","504 Flag","NYSAA Eligible","Housing Status","Temporary Housing Flag","Economic Disadvantage Flag","Cohort Year","Grade","Official Class","Home Language","Guardian Name 1","Guardian Relationship 1","Guardian Phone Number 1","Guardian Email 1","Guardian Spoken Language 1","Guardian Written Language 1","Guardian Name 2","Guardian Relationship 2","Guardian Phone Number 2","Guardian Email 2","Guardian Spoken Language 2","Guardian Written Language 2"];
  const psHeaders = ["Name","Students","Email","Phone","Secondary Phone","Record Created","Registered?","SIS Synced"];
  const nycsaHeaders = ["Location Code","School District","School Name","Class Code","Grade Level","Student ID","Student First Name","Student Last Name","Guardian First Name","Guardian Last Name","Another Custodial NYCSA Account Exist","Guardian Home Number","Guardian Work Number","Guardian Cell Number","Relationship","Role","Account Creation Code"];
  
  let targetZone = "";
  let expectedHeaders = [];
  let pasteStartCol = 1;
  
  if (startCol <= 31) {
    targetZone = "ATS";
    pasteStartCol = 1;
    expectedHeaders = atsHeaders;
  } else if (startCol >= 33 && startCol <= 40) {
    targetZone = "Parent Square";
    pasteStartCol = 33;
    expectedHeaders = psHeaders;
  } else if (startCol >= 42 && startCol <= 58) {
    targetZone = "NYCSA";
    pasteStartCol = 42;
    expectedHeaders = nycsaHeaders;
  } else {
    return; 
  }
  
  let headerRowIndex = -1;
  const expectedCheck = expectedHeaders.map(h => String(h).toLowerCase().replace(/[^a-z0-9]/g, ''));
  
  for (let i = 0; i < Math.min(20, pastedData.length); i++) {
    const rowStr = pastedData[i].join("").toLowerCase().replace(/[^a-z0-9]/g, '');
    let matchCount = 0;
    for (let j = 0; j < expectedCheck.length; j++) {
       if (rowStr.includes(expectedCheck[j])) matchCount++;
    }
    
    if ((targetZone === "Parent Square" && matchCount >= 2) || matchCount >= 3) {
      headerRowIndex = i;
      break;
    }
  }
  
  const finalData = [];
  const isFullExport = (headerRowIndex !== -1);
  
  if (isFullExport) {
    const headerRow = pastedData[headerRowIndex].map(h => String(h).toLowerCase().replace(/[^a-z0-9]/g, ''));
    const indexMap = expectedCheck.map(targetH => headerRow.indexOf(targetH));
    
    for (let i = headerRowIndex + 1; i < pastedData.length; i++) {
      if (pastedData[i].every(c => String(c).trim() === "")) continue;
      const newRow = indexMap.map(idx => {
        if (idx !== -1 && idx < pastedData[i].length) return pastedData[i][idx]; 
        return "";
      });
      finalData.push(newRow);
    }
  } else {
    for (let i = 0; i < pastedData.length; i++) {
       if (pastedData[i].every(c => String(c).trim() === "")) continue;
       const newRow = pastedData[i].slice(0, expectedHeaders.length);
       while (newRow.length < expectedHeaders.length) newRow.push("");
       finalData.push(newRow);
    }
  }
  
  if (finalData.length === 0 && !isFullExport) return;
  
  range.clear();
  
  const maxCols = sheet.getMaxColumns();
  if (maxCols > 61) {
     sheet.deleteColumns(62, maxCols - 61);
  } else if (maxCols < 61) {
     sheet.insertColumnsAfter(maxCols, 61 - maxCols);
  }
  SpreadsheetApp.flush(); 
  
  const lastRow = sheet.getMaxRows();
  if (isFullExport && lastRow >= 7) {
     sheet.getRange(7, pasteStartCol, lastRow - 6, expectedHeaders.length).clearContent();
  }
  
  sheet.getRange(6, 1, 1, atsHeaders.length).setValues([atsHeaders]);
  sheet.getRange(6, 33, 1, psHeaders.length).setValues([psHeaders]);
  sheet.getRange(6, 42, 1, nycsaHeaders.length).setValues([nycsaHeaders]);
  
  sheet.getRange("BH6").setValue("Site");
  sheet.getRange("BI6").setValue("Class");
  
  sheet.getRange("A5").setValue("PASTE").setHorizontalAlignment("center");
  sheet.getRange("AG6").setValue("PASTE").setHorizontalAlignment("center");
  sheet.getRange("AP2").setValue("PASTE").setHorizontalAlignment("center");
  
  try {
    sheet.getRange("BH3:BI3").merge().setValue("Class / Site").setHorizontalAlignment("center");
  } catch(e) {
    sheet.getRange("BH3").setValue("Class / Site").setHorizontalAlignment("center");
  }
  
  sheet.getRange("AF:AF").clearContent().setBackground("#FFFFFF");
  sheet.getRange("AO:AO").clearContent().setBackground("#FFFFFF");
  sheet.getRange("BG:BG").clearContent().setBackground("#FFFFFF");
  
  const targetRow = isFullExport ? 7 : Math.max(7, range.getRow());
  if (finalData.length > 0) {
    sheet.getRange(targetRow, pasteStartCol, finalData.length, expectedHeaders.length).setValues(finalData);
  }
  
  forceVibrantColors_(sheet);
  applyBandingToRawData_(sheet);
  
  SpreadsheetApp.getActiveSpreadsheet().toast(`Global headers locked. Data mapped to Row ${targetRow}.`, `${targetZone} Cleaned ✨`, 6);
}

/**
 * Keeps Raw Data rows 7+ at the normal height. A Site / Class cell (BI) holds one
 * class per line, and Sheets would otherwise make the whole row (ATS, Parent
 * Square, NYCSA too) as tall as that cell. The extra lines stay in the cell.
 */
function lockRawDataRowHeights_(sheet) {
  const rows = sheet.getMaxRows() - 6;
  if (rows > 0) sheet.setRowHeightsForced(7, rows, 21);
}

function fixRawDataNumberFormatting_(sheet) {
  const targetCols = [3, 22, 28, 36, 37, 47, 53, 54, 55, 58];
  const maxRows = sheet.getMaxRows();
  if (maxRows <= 6) return; 
  const numRowsToFormat = maxRows - 6;
  targetCols.forEach(col => {
    sheet.getRange(7, col, numRowsToFormat, 1).setNumberFormat("General"); 
  });
}

// ======================================================================
// --- FAST OSIS DELTA SYNC ---
// Rewrites one student's notes (Combined Contact Tracking, column D) from the
// Contact Log, Event Log and Notes without a full rebuild. A student who isn't
// on that tab yet triggers updateContactAndPhoneOnly() instead.
// ======================================================================
function updateSingleOsisDelta_(targetOsis) {
  updateOsisDeltas_([targetOsis]);
}

/**
 * Quick update for several OSIS in a row (e.g. a sibling group, then each
 * sibling). Same result as calling updateSingleOsisDelta_ once per OSIS, but
 * the Contact Log, Event Log and Notes are read only once.
 */
function updateOsisDeltas_(osisList) {
  // The owner's sheet also rewrites that student's Phone Contacts rows
  const owner = getOwnerHooks_();
  if (owner) return owner.updateOsisDeltas(osisList);

  const targets = osisList.filter(o => o);
  if (!targets.length) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const props = PropertiesService.getDocumentProperties();
  const disableHiddenNotes = props.getProperty('disableHiddenNotes') === 'true';

  // Read each source once; none of them change while the notes are rewritten
  const contactSheet = ss.getSheetByName("Contact Log");
  const contactData = contactSheet ? contactSheet.getDataRange().getDisplayValues() : null;
  const eventData = backendReadBlock_(ss, "EVENT_LOG", true);
  const notesSheet = ss.getSheetByName("Notes");
  const nData = notesSheet ? notesSheet.getDataRange().getValues() : null;

  let combinedOsisColumn = null; // Combined Contact Tracking column A (re-read after a rebuild)

  targets.forEach(targetOsis => {
    const finalNotesList = notesTextForOsis_(targetOsis, contactData, eventData, nData, disableHiddenNotes);

    const combinedSheet = ss.getSheetByName("Combined Contact Tracking");
    if (!combinedSheet) return;

    if (!combinedOsisColumn) combinedOsisColumn = combinedSheet.getRange("A:A").getDisplayValues();
    const cleanTarget = String(targetOsis).replace(/[^a-zA-Z0-9]/g, '');
    let foundInBackend = false;

    for (let i = 1; i < combinedOsisColumn.length; i++) {
      // We strictly match the Combined Contact Tracking single-row OSIS
      const sheetOsis = String(combinedOsisColumn[i][0]).replace(/[^a-zA-Z0-9]/g, '');

      if (sheetOsis === cleanTarget && cleanTarget !== "") {
        combinedSheet.getRange(i + 1, 4).setValue(finalNotesList); // Column D
        foundInBackend = true;
        break;
      }
    }

    if (!foundInBackend && finalNotesList) {
      ss.toast("New student profile syncing to backend...", "Auto-Build", 3);
      updateContactAndPhoneOnly();
      combinedOsisColumn = null; // that rebuild rewrote the tab
    }
  });
}

/**
 * One student's notes text (what Combined Contact Tracking column D holds after a
 * quick update) from already-read Contact Log (display values), Backend Event
 * Log and Notes data. Used by the quick update and the search bar's
 * Communications box, so both always show the same text.
 */
function notesTextForOsis_(targetOsis, contactData, eventData, nData, disableHiddenNotes) {
  const combinedInteractions = [];

  if (contactData) {
    for (let i = 1; i < contactData.length; i++) {
      const row = contactData[i];
      // ⚡ SHIFTED: OSIS boundary is now array index [4] (Col E)
      if (String(row[4]).trim().toUpperCase() === "END") break;
      if (!osisCellIncludes_(row[4], targetOsis)) continue;

      const entry = contactLogNoteLine_(row, disableHiddenNotes);
      if (entry) combinedInteractions.push(entry);
    }
  }

  if (eventData) {
    for (let i = 1; i < eventData.length; i++) {
      const row = eventData[i];
      if (!osisCellIncludes_(row[3], targetOsis)) continue;
      combinedInteractions.push(eventNoteLine_(row));
    }
  }

  // ✨ NEW: Fetch Manual Notes using .includes() so siblings get it too!
  let mergedManualNotes = "";
  if (nData) {
    for (let i = 1; i < nData.length; i++) {
      if (!nData[i]) continue;
      if (!osisCellIncludes_(nData[i][0], targetOsis)) continue;

      const nText = String(nData[i][3] || "").trim();
      if (nText) {
        mergedManualNotes = mergedManualNotes ? mergedManualNotes + "\n\n" + nText : nText;
      }
    }
  }

  // Merge the manual note block over the interactions
  return joinNoteLines_(mergedManualNotes, combinedInteractions);
}

function updateContactAndPhoneOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName("RAW Data") || ss.getSheetByName("Raw Data");
  if (!rawSheet) return;

  const rawData = rawSheet.getDataRange().getValues();
  if (rawData.length < 4) return;

  const overrides = getOverrideData_(ss);
  const studentMap = buildStudentMap_(rawData, overrides.students);
  const guardianRows = buildGuardianObjects_(rawData, overrides.guardians);
  const masterTableData = getExistingMasterTable_(ss);

  // buildPhoneContactsData_ requires these two indexes (args 6 & 7). Without them
  // getParentsSquareStatus_ dereferences `undefined.emailStatus` and throws on the
  // first guardian that has an email or phone, silently killing this whole path.
  const parentSquareIndex = buildParentSquareIndex_(rawData);
  const nycsaIndex = buildNYCSAIndex_(rawData);

  // Use getDisplayValues() so date handling matches the main build pipeline
  // (buildAllDerivedSheets reads the Contact Log the same way).
  const contactLogSheet = ss.getSheetByName("Contact Log");
  const contactLogData = contactLogSheet ? contactLogSheet.getDataRange().getDisplayValues() : [];
  
  const eventLogData = backendReadBlock_(ss, "EVENT_LOG");

  const combinedContactTracking = buildCombinedContactTrackingData_(ss, masterTableData, contactLogData, eventLogData);
  writeSheet_(ss, "Combined Contact Tracking", combinedContactTracking);

  const phoneContacts = buildPhoneContactsData_(ss, studentMap, guardianRows, masterTableData, combinedContactTracking, parentSquareIndex, nycsaIndex);
  writeSheet_(ss, "Phone Contacts", phoneContacts);
}

function getExistingMasterTable_(ss) {
  const sheet = ss.getSheetByName("Master Table");
  if (!sheet) return [["OSIS", "Student", "Guardian", "Class"]];

  const lastRow = sheet.getLastRow();
  if (lastRow < 5) return [["OSIS", "Student", "Guardian", "Class"]];

  const data = sheet.getRange(5, 1, lastRow - 4, 4).getValues(); 
  return data.filter(row => row[0] && !String(row[0]).includes("#"));
}

// ----------------------------------------------------------------------
// ROW WRITING FUNCTIONS
// ----------------------------------------------------------------------
function writeSingleRowDelta_(ss, sheetName, newArrayData, targetOsis, osisColIndex) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return null;
  if (!Array.isArray(newArrayData) || newArrayData.length === 0) return null;

  const searchOsis = String(targetOsis).trim();
  let dataRowIndex = -1;

  for (let i = 0; i < newArrayData.length; i++) {
    const cellOsisList = String(newArrayData[i][osisColIndex]).split(',').map(s => s.trim());
    if (cellOsisList.includes(searchOsis)) {
      dataRowIndex = i;
      break;
    }
  }

  const lastRow = sheet.getLastRow();
  let sheetRowIndex = -1;
  
  if (lastRow > 0) {
    const sheetOsisValues = sheet.getRange(1, osisColIndex + 1, lastRow, 1).getValues();
    for (let i = 0; i < sheetOsisValues.length; i++) {
      const cellOsisList = String(sheetOsisValues[i][0]).split(',').map(s => s.trim());
      if (cellOsisList.includes(searchOsis)) {
        sheetRowIndex = i + 1;
        break;
      }
    }
  }

  if (dataRowIndex === -1) {
    if (sheetRowIndex !== -1) {
      sheet.deleteRow(sheetRowIndex);
    }
    return null;
  }

  const targetDataRow = newArrayData[dataRowIndex];

  if (sheetRowIndex !== -1) {
    sheet.getRange(sheetRowIndex, 1, 1, targetDataRow.length).setValues([targetDataRow]);
  } else {
    sheet.appendRow(targetDataRow);
  }
  
  return targetDataRow; 
}

function writeMultipleRowDeltas_(ss, sheetName, newArrayData, targetOsis, osisColIndex) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  if (!newArrayData || newArrayData.length === 0) return [];

  const searchOsis = String(targetOsis).trim();

  const newMatchingRows = newArrayData.filter(row => {
    if (row[osisColIndex] == null) return false;
    return osisCellIncludes_(row[osisColIndex], searchOsis);
  });

  const existingRowIndices = [];
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow > 0 && lastCol >= (osisColIndex + 1)) {
    const sheetOsisValues = sheet.getRange(1, osisColIndex + 1, lastRow, 1).getValues();
    for (let i = 0; i < sheetOsisValues.length; i++) {
      if (sheetOsisValues[i][0] != null) {
        if (osisCellIncludes_(sheetOsisValues[i][0], searchOsis)) {
          existingRowIndices.push(i + 1); 
        }
      }
    }
  }

  let newRowIndex = 0;
  let existingIndexPtr = 0;

  while (newRowIndex < newMatchingRows.length && existingIndexPtr < existingRowIndices.length) {
    const targetRow = existingRowIndices[existingIndexPtr];
    const newData = newMatchingRows[newRowIndex];
    sheet.getRange(targetRow, 1, 1, newData.length).setValues([newData]);
    newRowIndex++;
    existingIndexPtr++;
  }

  while (newRowIndex < newMatchingRows.length) {
    sheet.appendRow(newMatchingRows[newRowIndex]);
    newRowIndex++;
  }

  const rowsToDelete = [];
  while (existingIndexPtr < existingRowIndices.length) {
    rowsToDelete.push(existingRowIndices[existingIndexPtr]);
    existingIndexPtr++;
  }
  
  rowsToDelete.sort((a, b) => b - a);
  for (let i = 0; i < rowsToDelete.length; i++) {
    sheet.deleteRow(rowsToDelete[i]);
  }

  return newMatchingRows;
}

// ==============================================================
// 1. ONE-BOX EVENT FORM (Custom UI with Calendar Picker)
// ==============================================================
function createNewEventTable() {
  const html = include_('Styles') + `
    <style>
      .ev { padding: 8px 12px 12px; display: flex; flex-direction: column; gap: 14px; }
      .chips { display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 8px; }
      .chip { border: 1px solid var(--line); background: #ffffff; border-radius: var(--radius); padding: 5px 10px; font: 600 12.5px var(--font); color: var(--ink); cursor: pointer; }
      .chip:hover { background: var(--panel); }
      .chip[aria-pressed="true"] { background: var(--accent-soft); border-color: var(--accent); color: var(--accent); }
      .foot { display: flex; justify-content: flex-end; gap: 8px; margin-top: 4px; }
    </style>
    <div class="ev">
      <div>
        <label class="lab" for="name">Event name <span class="req">*</span></label>
        <input type="text" id="name" placeholder="e.g. Spring Festival" autofocus>
      </div>

      <div>
        <label class="lab" for="type">Type <span class="req">*</span></label>
        <div class="chips" id="typeChips">
          <button type="button" class="chip">Parent Teacher Conference</button>
          <button type="button" class="chip">Workshop</button>
          <button type="button" class="chip">Family event</button>
          <button type="button" class="chip">Meeting</button>
        </div>
        <input type="text" id="type" placeholder="Pick one above or type your own">
      </div>

      <div>
        <label class="lab" for="loc">Location</label>
        <input type="text" id="loc" placeholder="e.g. Gym, Zoom">
      </div>

      <div>
        <label class="lab" for="date">Date <span class="req">*</span></label>
        <input type="date" id="date">
      </div>

      <div class="foot">
        <button type="button" class="btn sec" onclick="google.script.host.close()">Cancel</button>
        <button type="button" class="btn pri" id="createBtn" onclick="submit()">Create table</button>
      </div>

      <script>
        // The type buttons only fill in the Type box; any typed type works too.
        const typeBox = document.getElementById('type');
        const chips = document.querySelectorAll('#typeChips .chip');
        function markChip() {
          chips.forEach(function(c) { c.setAttribute('aria-pressed', String(c.textContent === typeBox.value.trim())); });
        }
        chips.forEach(function(c) {
          c.addEventListener('click', function() { typeBox.value = c.textContent; markChip(); });
        });
        typeBox.addEventListener('input', markChip);
        markChip();

        function submit() {
          const name = document.getElementById('name').value.trim();
          const type = document.getElementById('type').value.trim();
          const loc = document.getElementById('loc').value.trim();
          const rawDate = document.getElementById('date').value; // Comes out as YYYY-MM-DD

          if(!name) {
            alert("⚠️ Event Name is required!");
            return;
          }
          if(!type) {
            alert("⚠️ Event Type is required!");
            return;
          }
          if(!rawDate) {
            alert("⚠️ Date is required! Please select a date from the calendar.");
            return;
          }

          // Prevent double-clicking by instantly disabling the button
          const btn = document.getElementById('createBtn');
          btn.disabled = true;
          btn.innerText = 'Creating…';

          // Reformat the calendar output from YYYY-MM-DD to MM/DD/YYYY
          const parts = rawDate.split('-');
          const formattedDate = parts[1] + '/' + parts[2] + '/' + parts[0];

          google.script.run.withSuccessHandler(google.script.host.close).buildEventTable(name, loc, formattedDate, type);
        }
      </script>
    </div>
  `;
  const ui = HtmlService.createHtmlOutput(html).setWidth(440).setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(ui, 'Create New Event');
}

// ==============================================================
// 2. THE TABLE BUILDER & DROPDOWN GENERATOR
// ==============================================================
function buildEventTable(eventName, eventLoc, eventDate, eventType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Events");
  
  if (!sheet) {
    sheet = ss.insertSheet("Events");
  }
  
  sheet.activate(); 

  // Insert 7 columns on the far left (6 data + 1 spacer)
  sheet.insertColumns(1, 7);
  
  // Clear any residual formatting/validation from neighboring columns
  sheet.getRange(1, 1, sheet.getMaxRows(), 7).clearDataValidations().setBackground(null);

  // Set Row 1 Layout: Merged Title (Cols 1, 2), Category (Col 3), Location (Col 4), Date (Col 5), Delete Box (Col 6)
  sheet.getRange(1, 1, 1, 2).merge().setValue(eventName).setFontWeight("bold").setBackground("#d9ead3").setHorizontalAlignment("center");
  sheet.getRange(1, 3).setValue(eventType).setFontWeight("bold").setBackground("#DAA06D").setHorizontalAlignment("center"); // NEW: Event Type
  sheet.getRange(1, 4).setValue(eventLoc).setFontWeight("bold").setBackground("#fff2cc").setHorizontalAlignment("center");
  sheet.getRange(1, 5).setValue(eventDate).setFontWeight("bold").setBackground("#c9daf8").setHorizontalAlignment("center");
  sheet.getRange(1, 6).insertCheckboxes().setBackground("#ea9999");
  
  // Set Row 2 Headers
  const headers = EVENT_TABLE_HEADERS; // Sheet Lock.gs warns on row 2 of every table
  sheet.getRange("A2:F2").setValues([headers]).setFontWeight("bold").setBackground("#efefef");

  // Format Column Widths
  sheet.setColumnWidth(1, 100); // OSIS
  sheet.setColumnWidth(2, 150); // Guardian(s)
  sheet.setColumnWidth(3, 150); // Parent Attendee
  sheet.setColumnWidth(4, 150); // Student
  sheet.setColumnWidth(5, 80);  // # of Indv
  sheet.setColumnWidth(6, 200); // Notes
  sheet.setColumnWidth(7, 30);  // Spacer

  // Apply Borders
  sheet.getRange("A1:F").setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  
  // =======================================================
  // STRICT DROPDOWN MAPPING
  // =======================================================
  const masterSheet = ss.getSheetByName("Master Table"); 
  if (masterSheet) {
    const osisRule = SpreadsheetApp.newDataValidation().requireValueInRange(masterSheet.getRange("A5:A"), true).build();
    sheet.getRange("A3:A").setDataValidation(osisRule);

    const guardianRule = SpreadsheetApp.newDataValidation().requireValueInRange(masterSheet.getRange("C5:C"), true).build();
    sheet.getRange("B3:B").setDataValidation(guardianRule);
    
    // Parent Attendee is manual entry, no strict validation needed.

    const studentRule = SpreadsheetApp.newDataValidation().requireValueInRange(masterSheet.getRange("B5:B"), true).build();
    sheet.getRange("D3:D").setDataValidation(studentRule);
  }

  // Ensure the Backend tab's Event Log section has its header row
  if (backendReadBlock_(ss, "EVENT_LOG").length === 0) {
    backendWriteBlock_(ss, "EVENT_LOG", [["Date", "Event Name", "Category", "OSIS", "Student Name", "Guardian(s)", "Parent Attendee", "# of Indv", "Notes"]]);
  }

  // The tables moved right: re-lock every table's header row (Sheet Lock.gs)
  try { applySheetLocks_(ss, ["Events"]); } catch (e) { console.warn("Sheet lock skipped: " + e.message); }
}

// ==============================================================
// 3. EVENTS SHEET WATCHER: FULL RECONCILIATION SWEEP & SORT
// ==============================================================
function watchEventsSheet_(e) {
  if (!e || !e.range) return;

  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== "Events") return; 

  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();

  const maxRows = sheet.getLastRow();
  const maxCols = sheet.getLastColumn();
  if (maxRows < 3 || maxCols < 6) return;

  // Tables are 7 columns wide (6 data + 1 spacer)
  const tableStartCol = Math.floor((col - 1) / 7) * 7 + 1;
  const relativeCol = col - tableStartCol; // 0=OSIS, 1=Guardian, 2=Attendee, 3=Student, 4=#Indv, 5=Notes/Box
  
  // 1. GHOST TYPIST LOGIC (Runs instantly before anything else)
  if (row >= 3) {
    const masterSheet = e.source.getSheetByName("Master Table");
    if (masterSheet) {
      
      // ⚡ CASE A: VERTICAL PASTE (Copying & Pasting a column of OSIS numbers)
      if (range.getNumRows() > 1 && range.getNumColumns() === 1 && relativeCol === 0) {
        const pastedValues = range.getDisplayValues();
        const masterData = masterSheet.getDataRange().getValues();
        const updateData = [];
        let hasUpdates = false;

        // Fetch existing data for these rows to preserve manual entries (like attendee/notes)
        const existingData = sheet.getRange(row, tableStartCol, pastedValues.length, 6).getValues();

        for (let i = 0; i < pastedValues.length; i++) {
          let targetOsis = String(pastedValues[i][0]).trim();
          
          if (!targetOsis) {
             updateData.push([existingData[i][1], existingData[i][2], existingData[i][3], existingData[i][4]]); 
             continue;
          }
          
          // Fast memory sweep
          const master = findMasterRowByOsis_(masterData, targetOsis);
          const foundGuardian = master ? master[2] : "";
          const foundStudent = master ? master[1] : "";
          
          updateData.push([
            foundGuardian || existingData[i][1], 
            existingData[i][2], 
            foundStudent || existingData[i][3], 
            (String(existingData[i][4]).trim() === "") ? 1 : existingData[i][4]
          ]);
          hasUpdates = true;
        }
        
        if (hasUpdates) {
           sheet.getRange(row, tableStartCol + 1, updateData.length, 4).setValues(updateData);
        }
      }
      
      // ⚡ CASE A2: VERTICAL PASTE of Guardian or Student names
      else if (range.getNumRows() > 1 && (relativeCol === 1 || relativeCol === 3)) {
        const lookup = buildGhostLookup_(e.source);
        const block = sheet.getRange(row, tableStartCol, range.getNumRows(), 5).getValues(); // OSIS..#Indv
        const missing = [];
        let filled = 0;

        block.forEach(r => {
          const val = String(r[relativeCol]).trim();
          if (!val) return;
          const m = lookup(val, relativeCol === 1 ? 2 : 1);
          if (!m) { missing.push(val); return; }
          r[0] = m.osis;
          if (relativeCol !== 1) r[1] = m.guardian;
          if (relativeCol !== 3) r[3] = m.student;
          if (String(r[4]).trim() === "") r[4] = 1;
          filled++;
        });

        sheet.getRange(row, tableStartCol, block.length, 5).setValues(block);
        toastGhostMisses_(e.source, filled, missing);
      }

      // ⚡ CASE B: SINGLE CELL EDIT OR COMMA-SEPARATED PASTE
      else if (range.getNumRows() === 1 && range.getNumColumns() === 1) {
        const editedValue = String(e.value || range.getDisplayValue() || "").trim();
        
        if ((relativeCol === 0 || relativeCol === 1 || relativeCol === 3) && editedValue && editedValue !== "undefined") {
          
          // ⚡ NEW: MULTI-OSIS BATCH PASTE LOGIC (Comma or Line-break separated in ONE cell)
          if (relativeCol === 0 && (editedValue.includes(",") || editedValue.includes("\n"))) {
             // Split by comma or newline, remove whitespace, and filter out empties
             const osisArray = editedValue.split(/[,\n]+/).map(o => o.trim()).filter(Boolean);
             
             if (osisArray.length > 0) {
               const masterData = masterSheet.getDataRange().getValues();
               const numNewRows = osisArray.length - 1;
               
               if (numNewRows > 0) {
                 // Safely shift ONLY these 6 columns down to make room (Protects side-by-side tables!)
                 sheet.getRange(row + 1, tableStartCol, numNewRows, 6).insertCells(SpreadsheetApp.Dimension.ROWS);
                 
                 // Copy Data Validation and Formatting into the newly created gap
                 const templateRange = sheet.getRange(row, tableStartCol, 1, 6);
                 templateRange.copyTo(sheet.getRange(row + 1, tableStartCol, numNewRows, 6), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
                 templateRange.copyTo(sheet.getRange(row + 1, tableStartCol, numNewRows, 6), SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
               }
               
               const updateData = [];
               for (let i = 0; i < osisArray.length; i++) {
                 let targetOsis = osisArray[i];

                 // Fast memory sweep of the Master Table
                 const master = findMasterRowByOsis_(masterData, targetOsis);
                 const foundGuardian = master ? master[2] : "";
                 const foundStudent = master ? master[1] : "";
                 
                 // Preserve any manual notes typed into the original row, but keep subsequent rows clean
                 if (i === 0) {
                   const existingRowData = sheet.getRange(row, tableStartCol, 1, 6).getValues()[0];
                   updateData.push([
                     targetOsis, 
                     foundGuardian || existingRowData[1], 
                     existingRowData[2], 
                     foundStudent || existingRowData[3], 
                     (String(existingRowData[4]).trim() === "") ? 1 : existingRowData[4], 
                     existingRowData[5]
                   ]);
                 } else {
                   updateData.push([targetOsis, foundGuardian, "", foundStudent, 1, ""]);
                 }
               }
               
               // Blast the array back onto the spreadsheet instantly
               sheet.getRange(row, tableStartCol, osisArray.length, 6).setValues(updateData);
             }
          } 
          // ⚡ ORIGINAL LOGIC: Single typed entry
          else {
            // Typed OSIS -> Master Table A, Guardian -> C, Student -> B
            const found = findInMasterTable_(masterSheet, relativeCol === 0 ? 1 : relativeCol === 1 ? 3 : 2, editedValue);
            if (found) {
              const correctOsis = found.osis;
              const correctStudent = found.student;
              const correctGuardian = found.guardian;

              if (relativeCol !== 0) sheet.getRange(row, tableStartCol).setValue(correctOsis);
              if (relativeCol !== 1) sheet.getRange(row, tableStartCol + 1).setValue(correctGuardian);
              if (relativeCol !== 3) sheet.getRange(row, tableStartCol + 3).setValue(correctStudent);
              
              // Checks strictly for a completely empty string before forcing a 1
              if (String(sheet.getRange(row, tableStartCol + 4).getValue()).trim() === "") {
                sheet.getRange(row, tableStartCol + 4).setValue(1); 
              }
            }
          }
        }
      }
    }
  }

  SpreadsheetApp.flush(); 

  const eventName = String(sheet.getRange(1, tableStartCol).getValue()).trim();
  if (!eventName) return; 
  const category = String(sheet.getRange(1, tableStartCol + 2).getValue()).trim() || "Other";
  const dateString = sheet.getRange(1, tableStartCol + 4).getValue();
  let eventDate = new Date(dateString);
  if (isNaN(eventDate)) eventDate = new Date(); 
  const fDateString = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

  // =======================================================
  // 2. DELETE EVENT LOGIC (Red Checkbox)
  // =======================================================
  if (row === 1 && relativeCol === 5) {
    if (String(e.value).toUpperCase() === "TRUE") {
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert("⚠️ Delete Event", `Are you sure you want to completely delete the event "${eventName}"?\n\nThis will permanently delete all logs associated with it from the backend.`, ui.ButtonSet.YES_NO);
      
      if (response === ui.Button.YES) {
        sheet.deleteColumns(tableStartCol, 7);
        SpreadsheetApp.flush();
        try { applySheetLocks_(e.source, ["Events"]); } catch (err) { console.warn("Sheet lock skipped: " + err.message); }
        
        const freshData = buildBackendEventLogData_(e.source);
        backendWriteBlock_(e.source, "EVENT_LOG", freshData);
        ui.alert("✅ Event Removed", "The event has been deleted. To push this deletion to the Contact Tracker, click 'Build Sheets Only' in your App Menu.", ui.ButtonSet.OK);
        return; 
      } else {
        range.uncheck(); 
        return; 
      }
    } else {
      return; 
    }
  }

  // =======================================================
  // 3. RENAME / DATE CHANGE INTERCEPTOR 
  // =======================================================
  if (row === 1) {
      const oldValue = String(e.oldValue || "").trim();
      if (oldValue && oldValue !== "undefined") {
          e.source.toast("Rebuilding event data...", "⚙️ Syncing", 3);
          const freshData = buildBackendEventLogData_(e.source);
          backendWriteBlock_(e.source, "EVENT_LOG", freshData);
          let affected = new Set();
          for(let r = 1; r < freshData.length; r++) {
             if (String(freshData[r][1]).trim() === eventName) affected.add(String(freshData[r][3]).trim());
          }
          SpreadsheetApp.flush();
          updateOsisDeltas_(Array.from(affected).filter(Boolean));
          return; 
      }
  }

  // =======================================================
  // 4. THE ABSOLUTE RECONCILIATION SWEEP
  // =======================================================
  const allTableData = sheet.getRange(1, 1, maxRows, maxCols).getValues();
  const frontendRecords = new Map(); 
  const rowsToClear = []; 

  for (let c = 0; c < maxCols; c += 7) {
      const tName = String(allTableData[0][c] || "").trim();
      if (!tName) continue;

      const tDateStr = allTableData[0][c + 4];
      let tDate = new Date(tDateStr);
      if (isNaN(tDate)) tDate = new Date();
      const tDateString = Utilities.formatDate(tDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

      if (tName === eventName && tDateString === fDateString) {
          for (let r = 2; r < maxRows; r++) { 
              let rOsisRaw = String(allTableData[r][c] || "").trim();
              if (!rOsisRaw) continue; 

              let rGuardian = String(allTableData[r][c + 1] || "").trim();
              let rAttendee = String(allTableData[r][c + 2] || "").trim();
              let rStudentRaw = String(allTableData[r][c + 3] || "").trim(); 
              
              // ⚡ THE ZERO FIX: Evaluates the literal raw cell value!
              let rawNumCell = allTableData[r][c + 4];
              let rNum = 1; 
              if (String(rawNumCell).trim() !== "") {
                  rNum = Number(rawNumCell);
              }

              let rNotes = String(allTableData[r][c + 5] || "").trim();

              const osisArray = rOsisRaw.split(',').map(o => o.trim()).filter(o => o !== "");
              const nameArray = rStudentRaw.split(',').map(n => n.trim()).filter(n => n !== "");

              for (let j = 0; j < osisArray.length; j++) {
                  let singleOsis = osisArray[j];
                  let singleName = nameArray[j] !== undefined ? nameArray[j] : rStudentRaw;
                  let assignedNum = (j === 0) ? rNum : ""; 

                  if (frontendRecords.has(singleOsis)) {
                      rowsToClear.push({ row: r + 1, col: c + 1 }); 
                  } else {
                      frontendRecords.set(singleOsis, {
                         guardian: rGuardian,
                         attendee: rAttendee,
                         student: singleName,
                         numIndv: assignedNum,
                         notes: rNotes,
                         synced: false 
                      });
                  }
              }
          }
      }
  }

  if (rowsToClear.length > 0) {
      SpreadsheetApp.getUi().alert("⚠️ Duplicate Removed", "This family is already logged for this event. The duplicate has been removed.", SpreadsheetApp.getUi().ButtonSet.OK);
      rowsToClear.forEach(target => sheet.getRange(target.row, target.col, 1, 6).clearContent());
  }

  // Backend tab's Event Log section (index 0 = header row)
  const backendData = backendReadBlock_(e.source, "EVENT_LOG");
  const affectedOsisNumbers = new Set(); 

  for (let i = backendData.length - 1; i > 0; i--) {
      let bEventName = String(backendData[i][1]).trim();
      let bDate = backendData[i][0];
      
      let bDateString = "";
      if (bDate instanceof Date) {
        bDateString = Utilities.formatDate(bDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else if (bDate) {
        let parsedBDate = new Date(bDate);
        if (!isNaN(parsedBDate)) bDateString = Utilities.formatDate(parsedBDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }

      if (bEventName === eventName && bDateString === fDateString) {
         let bOsis = String(backendData[i][3]).trim();

         if (frontendRecords.has(bOsis)) {
           let fData = frontendRecords.get(bOsis);
           
           if (backendData[i][4] != fData.student || backendData[i][5] != fData.guardian || backendData[i][6] != fData.attendee || backendData[i][7] != fData.numIndv || backendData[i][8] != fData.notes) {
              backendBlockRange_(e.source, "EVENT_LOG", i, 4, 1, 5).setValues([[fData.student, fData.guardian, fData.attendee, fData.numIndv, fData.notes]]);
              affectedOsisNumbers.add(bOsis);
           }
           fData.synced = true; 
         } else {
           backendDeleteRows_(e.source, "EVENT_LOG", [i]);
           affectedOsisNumbers.add(bOsis);
         }
      }
  }

  const newRows = [];
  for (let [fOsis, fData] of frontendRecords.entries()) {
     if (!fData.synced) {
        newRows.push([eventDate, eventName, category, fOsis, fData.student, fData.guardian, fData.attendee, fData.numIndv, fData.notes]);
        affectedOsisNumbers.add(fOsis);
     }
  }
  backendAppendRows_(e.source, "EVENT_LOG", newRows);
  
  backendSortBlock_(e.source, "EVENT_LOG", [{col: 0, ascending: false}, {col: 1, ascending: true}]);

  SpreadsheetApp.flush(); 
  updateOsisDeltas_(Array.from(affectedOsisNumbers).filter(Boolean));
}

/** First Master Table data row (row 5 down) whose OSIS (A) is exactly `osis`, from getDataRange() values. */
function findMasterRowByOsis_(masterData, osis) {
  for (let r = 4; r < masterData.length; r++) {
    if (String(masterData[r][0]).trim() === osis) return masterData[r];
  }
  return null;
}
