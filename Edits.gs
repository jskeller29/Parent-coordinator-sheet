// ======================================================================
// FILE: Edits.gs
// PURPOSE: Everything that runs while someone types: the installable onEdit
// trigger, the ghost typists (Contact Log, Notes, Events), Raw Data paste
// clean-up, and the fast single-student "quick update" of Combined Contact
// Tracking / Phone Contacts. Also the Events tables (menu + watcher). (The
// temporary 🐞 Contact Log Debug sidebar is the owner's test file z_ContactLogDebug.gs.)
// ======================================================================

// =========================================================
// MAIN TRIGGER FUNCTION (installable onEdit)
// =========================================================
function installedOnEdit(e) {
  if (!e || !e.range) return;
  const result = runAction_(() => "Edit " + e.range.getSheet().getName() + "!" + e.range.getA1Notation(),
    () => clDbgRun_(() => "Edit " + e.range.getSheet().getName() + "!" + e.range.getA1Notation() +
      " = " + JSON.stringify(String(e.value === undefined ? "(paste / multi-cell or cleared)" : e.value).slice(0, 60)), () => handleEdit_(e)));
  // Data update reminder while the sheet is open (at most one check an hour; Menus.gs)
  try { checkDataReminderWhileOpen_(); } catch (err) { console.warn("Data reminder check failed: " + err.message); }
  return result;
}

/**
 * Installable change trigger (added by ensureChangeTrigger_, Menus.gs). Row
 * deletes do not fire onEdit, so a deleted Contact Log row's note would stay in
 * Combined Contact Tracking (and the search bar) until the next build.
 */
function installedOnChange(e) {
  if (!e || (e.changeType !== "REMOVE_ROW" && e.changeType !== "INSERT_ROW")) return;
  const removed = e.changeType === "REMOVE_ROW";
  return runAction_(removed ? "Row deleted" : "Row inserted", () => clDbgRun_(removed ? "Row(s) deleted" : "Row(s) inserted", () => {
    const ss = e.source || SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    clDbg_("active tab: " + (sheet ? sheet.getName() : "?"));
    if (!sheet || sheet.getName().trim() !== "Contact Log") return;
    // Row 2 holds the Person Spoke With formula: deleting row 2 takes it, a row inserted above row 2
    // pushes it down. Put it back now (own try/catch, so it never stops the notes below)
    try { clDbg_("Person Spoke With formula (R2): " + restoreSpokeWithChoicesAfterRowChange_(ss, sheet)); }
    catch (err) { console.warn("Person Spoke With formula not checked: " + err.message); }
    if (removed) resyncNotesAfterRowDelete_(ss, sheet); // an inserted (empty) row changes no notes
  }));
}

/**
 * After Contact Log rows are deleted (the change event doesn't say which):
 * redraws the END bar's border, then works out every student's notes from the
 * logs and updates only those whose Combined Contact Tracking text changed
 * (normally just the deleted row's student). When many differ, Combined is
 * rebuilt instead (updateContactAndPhoneOnly).
 */
function resyncNotesAfterRowDelete_(ss, sheet, knownEndRow) {
  // knownEndRow: the edit handler already knows END (it just redrew its border)
  let endRow = knownEndRow;
  if (!endRow) {
    // A row delete: END moved up; refresh its remembered row and its top/bottom border
    endRow = findEndRowInColumnE_(sheet);
    rememberEndRow_(sheet, endRow);
    // Only the line ABOVE the END bar: the line under it is the lookup box's thick top line, and a thin
    // one drawn there replaced it (release check 4.4: "the border above search is no longer thick")
    if (endRow) sheet.getRange(endRow, 1, 1, CL_WIDTH).setBorder(true, null, null, null, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID);
    clDbg_("END is now row " + endRow + " (border redrawn)");
  }

  const props = PropertiesService.getDocumentProperties();
  if (props.getProperty('autoNotes') === 'false') { clDbg_("skipped: 'update notes as I type' is off in Settings"); return false; }
  const combined = ss.getSheetByName("Combined Contact Tracking");
  const last = combined ? combined.getLastRow() : 0;
  if (last < 2) return false;

  const disableHiddenNotes = props.getProperty('disableHiddenNotes') === 'true';
  const contactData = endRow ? sheet.getRange(1, 1, endRow, CL_WIDTH).getDisplayValues() : sheet.getDataRange().getDisplayValues();
  const sources = notesSourcesData_(ss); // Event Log + Notes tab (a row delete changes neither)
  const eventData = sources.eventData;
  const nData = sources.nData;

  // Group each source's rows by OSIS once (same rows osisCellIncludes_ would
  // pick, same order), so each student only looks at their own rows
  const byOsis = (data, col, stopAtEnd) => {
    const map = new Map();
    if (!data) return map;
    for (let i = 1; i < data.length; i++) {
      if (stopAtEnd && String(data[i][col]).trim().toUpperCase() === "END") break;
      new Set(extractOsisTokens_(data[i][col])).forEach(o => {
        if (!map.has(o)) map.set(o, []);
        map.get(o).push(data[i]);
      });
    }
    return map;
  };
  const contactBy = byOsis(contactData, 4, true), eventBy = byOsis(eventData, 3, false), notesBy = byOsis(nData, 0, false);
  const withHeader = (map, osis) => [[]].concat(map.get(osis) || []); // notesTextForOsis_ skips row 0

  const changed = [];
  combined.getRange(2, 1, last - 1, 4).getDisplayValues().forEach(r => {
    const osis = String(r[0]).trim();
    if (!isValidOsis_(osis)) return;
    const text = notesTextForOsis_(osis, withHeader(contactBy, osis), withHeader(eventBy, osis), withHeader(notesBy, osis), disableHiddenNotes);
    if (text !== String(r[3])) changed.push(osis);
  });
  clDbg_("notes compared for " + (last - 1) + " Combined rows; different: " + (changed.length ? changed.slice(0, 10).join(", ") + (changed.length > 10 ? " …" : "") : "none"));
  if (!changed.length) return true;

  if (changed.length > 25) { clDbg_("more than 25 differ -> full rebuild of Combined"); updateContactAndPhoneOnly(); }
  else {
    // The quick one-cell writes first (the search bar formula follows), then on
    // the owner's sheet the update that also rewrites Phone Contacts
    const owner = getOwnerHooks_();
    const quick = updateCombinedNotesQuick_(changed, !owner);
    if (owner) updateOsisDeltas_(changed, { combinedAlreadyWritten: quick.missing.length === 0 });
  }
  clDbg_("notes updated");
  return true;
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
    watchEventsSheet_(e); // fills the row in only; the Backend Event Log / notes wait for the build
    return; 
  }

  // --- NEW: TRIGGER THE NOTES GHOST TYPIST ---
  if (sheetName === "Notes") {
    clearNotesSourcesCache_();
    watchNotesSheet_(e);
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

  // ONE CONTACT LOG EDIT AT A TIME for the part that reads END and adds the new row (2026-11-18,
  // release check 7.2: "the new row didn't have a checkbox"). Typing two cells of the empty row
  // quickly starts two runs side by side: both saw the old END row, both added a blank row, and the
  // first run's clean-up under END then took the checkbox off the other's new row. The second run
  // now waits here until the first has added its row, so it finds END where it really is. Let go
  // before the notes (they take the same lock for their own write); Google lets it go anyway when
  // the run ends. No lock within 10 s: on as before.
  const rowLock = LockService.getDocumentLock();
  let rowLocked = false;
  try { rowLocked = rowLock.tryLock(10000); } catch (err) { rowLocked = false; }
  const releaseRowLock = () => {
    if (!rowLocked) return;
    rowLocked = false;
    try { rowLock.releaseLock(); } catch (err) { /* let go when the run ends */ }
  };
  if (!rowLocked) clDbg_("another Contact Log edit held the row lock for 10 s: going on without it");

  // 1. Shifted End Row Finder (Now Col E / 5). ONE cache call brings back the
  // END row and the ghost typist's data; the END row is then checked with a
  // one-cell read (a full column E scan only when it moved).
  const prefetch = contactLogPrefetch_(sheet, row);
  let endRow = prefetch.endRow;
  clDbg_("END row " + endRow + (prefetch.endFromCache ? " (remembered)" : " (scanned column E)") + "; typist data " + (prefetch.ghostRaw ? "cached" : "not cached"));

  // --- NEW FILTER LOGIC ---
  // The Master Toggle lives on the end bar row in Col O
  if (endRow && row === endRow && col === CL_COL_FILTER_TOGGLE) {
    toggleHideFilter_(sheet, endRow);
    releaseRowLock();
    return; 
  }
  // ------------------------

  if (endRow && row >= endRow) {
    // An edit in the search bar below the END bar (e.g. looking up a student):
    // its formulas update by themselves
    releaseRowLock();
    clDbg_("edit is in the search bar area (below END): its formulas update by themselves");
    return;
  }

  // =======================================================
  // CONTACT LOG GHOST TYPIST — PASTED COLUMN (fills every pasted row)
  // =======================================================
  let pasteAddedRows = 0; // the paste reached END and empty rows were added there
  if (numRows > 1 && (col === CL_COL_GUARDIAN || col === CL_COL_OSIS || col === CL_COL_STUDENT)) {
    try {
      const pasteRows = endRow ? Math.min(numRows, endRow - row) : numRows;
      if (pasteRows > 0) pasteAddedRows = Number(fillPastedContactLogRows_(e.source, sheet, row, pasteRows, col, endRow, range.getNumColumns())) || 0;
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
  let typistOsis = null; // the OSIS the ghost typist wrote (saves re-reading the cell)
  if (range.getNumRows() === 1 && range.getNumColumns() === 1) {
    try {
      // Watch columns: C (Guardian), E (OSIS), F (Student)
      if ((col === CL_COL_GUARDIAN || col === CL_COL_OSIS || col === CL_COL_STUDENT) && editedValue && editedValue.toUpperCase() !== "FALSE" && editedValue.toUpperCase() !== "TRUE") {
        // Student List A5:C + Directory C:E, from a short cache when the
        // previous edit already read them (ghostTypistData_)
        const tTypist = Date.now();
        const ghost = ghostTypistData_(e.source, prefetch.ghostRaw);
        const tLookupData = Date.now() - tTypist;
        if (ghost.master.length) {
          const masterData = ghost.master;
          const matchCol = (col === CL_COL_OSIS) ? 0 : (col === CL_COL_STUDENT) ? 1 : 2;
          // Stray spaces don't matter (typed or in the Student List): "Ann  Smith " = "Ann Smith"
          const needle = typistKey_(editedValue);
          let match = null;
          for (let i = 0; i < masterData.length; i++) {
            if (typistKey_(masterData[i][matchCol]) === needle) { match = masterData[i]; break; }
          }

          if (match) {
            const correctOsis = match[0];
            typistOsis = String(correctOsis).trim();
            const correctStudent = match[1];
            const correctGuardian = match[2];

            // --- PULL SITE & CLASS FROM DIRECTORY (HANDLES SIBLINGS) ---
            const sites = new Set();
            const classes = new Set();

            if (correctOsis) {
              const osisSet = new Set(String(correctOsis).split(',').map(o => o.trim()).filter(String));
              // OSIS/Site/Class (Directory C:E), already read with the Student List
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
            // (painted together with the auto-fill below: one flush for both)
            clDbg_("⏱ ghost typist " + (Date.now() - tTypist) + " ms (Student List / Directory data " + tLookupData +
              " ms" + (prefetch.ghostRaw ? ", cached" : ", read from the tabs") + "): OSIS " + correctOsis + " / " + correctStudent + " / " + correctGuardian);

          } else {
            clDbg_("ghost typist: '" + editedValue + "' not found in the Student List");
            e.source.toast("Could not find '" + editedValue + "' in the Student List.", "Ghost Typist Failed", 4);
          }
        }
      }

      // =======================================================
      // AUTO-FILL DATE / # INTERACTIONS / FOLLOW-UP STATUS
      // Fires on the same single-cell edits. Fill-if-blank only, and skipped
      // entirely when "Disable Contact Log Auto-Fill" is on in Settings. The
      // identity lookup above always runs regardless of that setting.
      // =======================================================
      const tAuto = Date.now();
      autofillContactLogDefaults_(sheet, row, col, endRow, editedValue, prefetch.rowVals);
      clDbg_("⏱ auto-fill (Date / # / Follow-up) " + (Date.now() - tAuto) + " ms" + (prefetch.rowVals ? " (row already read)" : ""));
      // Paint the ghost typist's cells and the auto-fill together, right now
      const tPaint = Date.now();
      SpreadsheetApp.flush();
      clDbg_("⏱ painted typist + auto-fill in " + (Date.now() - tPaint) + " ms");
    } catch (err) {
      // Surface ghost-typist/auto-fill failures instead of dying silently
      // mid-edit; the repair and delta sync below still get a chance to run.
      console.error(err);
      e.source.toast("Auto-fill failed: " + err.message, "⚠️ Ghost Typist Error", 5);
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
    formatContactLogRows_(sheet, row, safeNumRows); // also the font (was: column D set to Roboto here)
  }

  // A row was added above END (by this edit, or by the paste above when it reached END: before
  // 2026-11-18 the paste's row was missed, so END's clean-up cleared the new row's checkbox)
  // (the paste's row first: a paste that started on the last row would otherwise get a second one)
  const addedRows = pasteAddedRows || autoAddBlankRow_(sheet, row, endRow); // how many empty rows went in above END
  // (the week formula check, patchMissingFormulasInRow_, runs at the end of the
  // edit: its reads would hold up the clean-up and the notes)

  // --- LAG FIX: Only recalculate colors if you actually changed the Type (Col L) ---
  if (col === CL_COL_TYPE) {
    applyTypeDropdownColors();
  }

  // ==================================================
  // --- FORCE END BAR FORMATTING & SWEEP ---
  // ==================================================
  // Where END is now: it was checked at the start of this edit, and moves down
  // exactly one row when a blank row was added above it (no read needed);
  // column E is only scanned when END was not known
  const finalEndRow = endRow ? endRow + addedRows : findEndRowInColumnE_(sheet);
  rememberEndRow_(sheet, finalEndRow);
  if (finalEndRow) {
    // the line above the END bar only (under it is the lookup box's thick line; release check 4.4)
    sheet.getRange(finalEndRow, 1, 1, CL_WIDTH).setBorder(true, null, null, null, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID);

    // Checkboxes under END come only from a paste reaching down there. A one-cell edit never
    // changes a row under END (the new row goes above it), so it skips this (one write less,
    // and it can't clear a new row's checkbox when END moved meanwhile; 2026-11-18)
    if (numRows > 1 || range.getNumColumns() > 1) {
      let maxRows = sheet.getMaxRows();
      if (maxRows >= finalEndRow) {
        let rowsToClear = maxRows - finalEndRow + 1;
        sheet.getRange(finalEndRow, CL_COL_CHECKBOX, rowsToClear, 1).clearDataValidations().clearContent();
      }
    }
  }

  SpreadsheetApp.flush(); 
  releaseRowLock(); // the new row and END are on the sheet: the next edit may go on

  clDbg_("row clean-up + END bar done");

  // --- NOTES (right after the row clean-up; owner's order 2026-11-05) ---
  // A single-cell edit whose row already has an OSIS rewrites that student's
  // Combined Contact Tracking notes now (the quick one-cell write), and pushes
  // it to the sheet so the search bar's Communications formula shows them.
  let notesDone = false;
  const multiCell = numRows > 1 || range.getNumColumns() > 1;
  if (!multiCell && (!endRow || row < endRow)) {
    // The typist just wrote it (no read); otherwise read the row's OSIS cell
    const earlyOsis = typistOsis !== null ? typistOsis : String(sheet.getRange(row, CL_COL_OSIS).getDisplayValue()).trim();
    // OSIS cleared or typed over: the student it USED to be loses this row's note
    const oldOsis = (col === CL_COL_OSIS && e.oldValue !== undefined) ? String(e.oldValue).trim() : "";
    const extra = (isValidOsis_(oldOsis) && oldOsis !== earlyOsis) ? [oldOsis] : [];
    const mainOsis = isValidOsis_(earlyOsis) ? earlyOsis : (extra.length ? extra.shift() : "");
    // A notes problem is logged, and the normal sync below tries once more
    clDbg_("row OSIS for the notes: " + JSON.stringify(earlyOsis) + (oldOsis ? " (was " + JSON.stringify(oldOsis) + ")" : ""));
    if (mainOsis) {
      try {
        notesDone = syncNotesForRowOsis_(mainOsis, true, extra);
        clDbg_(notesDone ? "notes synced (quick write, pushed to the sheet)" : "notes NOT synced: 'update notes as I type' is off");
      }
      catch (err) { clDbg_("❌ quick notes sync failed: " + err.message); console.error("Quick notes sync failed (retried below): " + err.message); }
    }
  } else if (multiCell && (!endRow || row < endRow)) {
    // A paste or cleared row(s): which students changed is not known (a cleared
    // row's old OSIS is gone), so compare every student's notes with the logs
    // and update the ones that changed (same check as a deleted row)
    try {
      clDbg_("multi-cell edit (paste / cleared rows): checking which students' notes changed");
      notesDone = resyncNotesAfterRowDelete_(e.source, sheet, finalEndRow);
    } catch (err) { clDbg_("❌ notes check after the paste / clear failed: " + err.message); console.error("Notes check after a paste / clear failed: " + err.message); }
  }

  // Week formula in column A (reads, so after the notes)
  try { patchMissingFormulasInRow_(sheet, row); } catch (err) { console.warn("Week formula not checked: " + err.message); }

  if (notesDone) {
    // Already synced by the quick write (above). Owner's sheet: now the update
    // that also rewrites Phone Contacts.
    const pending = PENDING_OWNER_NOTES_;
    PENDING_OWNER_NOTES_ = null;
    if (pending) {
      // A student with no Combined row yet: the owner's update adds it (full version)
      try { updateOsisDeltas_(pending.list, { combinedAlreadyWritten: pending.combinedAlreadyWritten }); }
      catch (err) { clDbg_("❌ owner notes update failed: " + err.message); console.error("Owner notes update failed: " + err.message); }
    }
    return;
  }

  // --- DYNAMIC DELTA SYNC TRIGGER ---
  let targetOsis = String(sheet.getRange(row, CL_COL_OSIS).getDisplayValue()).trim();

  // A formula OSIS still "Loading...": wait a second and look again. A blank OSIS
  // has nothing to wait for (before 2026-11-12 every such edit waited 1 s).
  if (!isValidOsis_(targetOsis) && /loading/i.test(targetOsis)) {
    Utilities.sleep(1000);
    SpreadsheetApp.flush();
    targetOsis = String(sheet.getRange(row, CL_COL_OSIS).getDisplayValue()).trim();
  }

  clDbg_("row OSIS at the end of the edit: " + JSON.stringify(targetOsis));
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
    // OSIS cleared and its old value unknown: rebuild Combined + Phone Contacts
    if (col === CL_COL_OSIS && !editedValue && !multiCell) {
      updateContactAndPhoneOnly();
    }
  }
}

// =========================================================
// CONTACT LOG SEARCH BAR — COMMUNICATIONS
// The search bar sits a few rows below the END bar. Its "Communications" box is
// the owner's formula (SEARCH_BAR_COMMS_FORMULA, typed in the sheet): it reads
// Combined Contact Tracking, so it updates the instant an OSIS is typed, and
// after a Contact Log edit / row delete as soon as the quick notes update has
// rewritten that student's Combined row. (A script-drawn version was tried on
// 2026-11-05 and removed at the owner's request: a script can never be instant
// on a lookup.) The code below only finds the box (debug sidebar) and puts the
// formula back.
// =========================================================
const SEARCH_BAR_COMMS_LABEL = "Communications";
const SEARCH_BAR_SCAN_ROWS = 15; // rows below END searched for the label

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

/**
 * How many rows above the Communications label the OSIS box is: 1 in the old lookup box
 * layout (OSIS row right above Communications), 6 in the owner's 2026-11-11 layout (the
 * guardian table sits between them). Found from the "OSIS" label (the box is the cell right
 * of it); 1 when there is no such label above Communications.
 */
function searchBarOsisRowsAbove_(sheet, endRow, at) {
  const rows = at.row - endRow - 1;
  if (rows < 1) return 1;
  const label = sheet.getRange(endRow + 1, 1, rows, sheet.getMaxColumns())
    .createTextFinder("OSIS").matchEntireCell(true).findNext();
  return label ? at.row - label.getRow() : 1;
}

/**
 * The Communications formula for a label cell whose OSIS box is rowsAbove rows up (and the
 * name switch one row above that): SEARCH_BAR_COMMS_FORMULA with its row counts changed.
 */
function searchBarCommsFormula_(rowsAbove) {
  return SEARCH_BAR_COMMS_FORMULA
    .replace("ADDRESS(ROW()-1, COLUMN()+1)", "ADDRESS(ROW()-" + rowsAbove + ", COLUMN()+1)")
    .replace("ADDRESS(ROW()-2, COLUMN()+2)", "ADDRESS(ROW()-" + (rowsAbove + 1) + ", COLUMN()+2)");
}

/** Same as Sheets' TRIM: drops leading/trailing spaces and collapses repeated spaces. */
function sheetsTrim_(text) {
  return String(text).replace(/ {2,}/g, " ").replace(/^ +| +$/g, "");
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
  const sources = notesSourcesData_(ss);
  return osis => notesTextForOsis_(osis, contactData, sources.eventData, sources.nData, disableHiddenNotes);
}

/** Puts the original Communications formula back (Apps Script editor, or "Put formula back" in the debug sidebar). */
function restoreSearchBarCommunicationsFormula() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Contact Log");
  const endRow = sheet ? findEndRowInColumnE_(sheet) : null;
  const at = endRow ? findSearchBarComms_(sheet, endRow) : null;
  if (!at) { ss.toast("Could not find the Communications box below END.", "Search Bar", 5); return; }

  getStudentListSheet_(ss); // an older copy's "Master Table" is renamed first (the formula reads 'Student List')
  sheet.getRange(at.row, at.col + 1).clearContent();
  sheet.getRange(at.row, at.col).setFormula(searchBarCommsFormula_(searchBarOsisRowsAbove_(sheet, endRow, at)));
  ss.toast("Communications is a formula again.", "Search Bar", 5);
}

/** The original formula (only used to put it back; searchBarCommsFormula_ sets its row counts). */
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
           
           name, IFNA(VLOOKUP(oNum, 'Student List'!A:B, 2, FALSE), IFNA(VLOOKUP(oText, 'Student List'!A:B, 2, FALSE), "Unknown")),
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

// =========================================================
// CONTACT LOG LOOKUP BOX (the search bar's whole box below END)
// Owner's layout (2026-11-11, made by hand; small rows = 10 px): small row,
// OSIS | Student | Class, small row, guardian table, small row,
// Communications | Student Info. The code finds it, outlines it (reset) and
// keeps the guardian table as tall as the student with the most guardians
// needs (every build and reset). Its formulas are typed in the sheet.
// =========================================================
const LOOKUP_BOX_GUARDIAN_ROWS_ = 2;      // guardian rows the box is built with (never fewer)
const LOOKUP_BOX_MAX_GUARDIAN_ROWS_ = 10; // never more
const LOOKUP_BOX_SPACER_MAX_PX_ = 15;     // the small rows (10 px) between the parts
const LOOKUP_BOX_ROW_COLORS_ = ["#ffffff", "#fff8ee"]; // guardian 1, 3, 5... white; 2, 4... cream

/**
 * Thick black outline around the lookup box: top = the row above OSIS, left = the
 * "Communications" label's column, right = the box's last column, bottom = its last row
 * (new layout: the bottom of the Communications rows). None of these depend on the student
 * shown (the old outline looked for the first empty guardian row, which a reset has just
 * emptied, so it stopped at the guardian header). Returns false when the box is not found.
 */
function drawContactLogLookupBoxOutline_(sheet, endRow) {
  const box = findContactLogLookupBox_(sheet, endRow);
  if (!box) return false;
  const cols = box.right - box.left + 1;
  // Clear the old outline around the box (incl. a stray bottom line one row lower), then draw it
  const clearRows = Math.min(box.bottom - box.top + 2, sheet.getMaxRows() - box.top + 1);
  sheet.getRange(box.top, box.left, clearRows, cols).setBorder(false, false, false, false, null, null);
  sheet.getRange(box.top, box.left, box.bottom - box.top + 1, cols)
    .setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_THICK);
  return true;
}

/**
 * Where the lookup box is, or null when its "Communications" label is not found below END:
 * { top, left, right, bottom, lowest, old, header (0 = no "Relationship" header), commRow }.
 * - top: the row above the "OSIS" label (the Communications formula's name switch); END + 2 without it
 * - right: the furthest of the "Student Information" / "Student Info" cell's merge, the merges in
 *   the Communications row and the guardian header's last filled cell
 * - bottom: new layout = last Communications row (a merge there can make them taller);
 *   old layout (guardian table under Communications) = END + 8, the last row every reset kept
 * - lowest: the lowest row any part of the box was found on
 */
function findContactLogLookupBox_(sheet, endRow) {
  const maxRows = sheet.getMaxRows();
  const rows = maxRows - endRow;
  if (rows < 1) return null;
  const maxCols = sheet.getMaxColumns();
  const area = sheet.getRange(endRow + 1, 1, rows, maxCols);
  const found = text => area.createTextFinder(text).matchEntireCell(true).findNext();

  const commCell = found("Communications");
  // An erroring Communications formula shows no label: findSearchBarComms_ also matches its text
  const comm = commCell ? { row: commCell.getRow(), col: commCell.getColumn() } : findSearchBarComms_(sheet, endRow);
  if (!comm) return null;
  const osis = found("OSIS");
  const info = found("Student Information") || found("Student Info");
  const rel = found("Relationship");
  const header = rel ? rel.getRow() : 0;

  let right = comm.col;
  if (info) right = Math.max(right, (info.isPartOfMerge() ? info.getMergedRanges()[0] : info).getLastColumn());
  let commEnd = comm.row;
  sheet.getRange(comm.row, comm.col, 1, maxCols - comm.col + 1).getMergedRanges().forEach(m => {
    commEnd = Math.max(commEnd, m.getLastRow());
    right = Math.max(right, m.getLastColumn());
  });
  if (header) {
    const shown = sheet.getRange(header, 1, 1, maxCols).getDisplayValues()[0];
    for (let c = shown.length; c > comm.col; c--) if (shown[c - 1] !== "") { right = Math.max(right, c); break; }
  }

  const top = osis && osis.getRow() > endRow + 1 && osis.getRow() < comm.row ? osis.getRow() - 1 : endRow + 2;
  const old = header > commEnd && !!info && info.getRow() < comm.row;
  const bottom = old ? Math.max(header, Math.min(endRow + 8, maxRows)) : commEnd;
  const lowest = Math.max(commEnd, header, info ? info.getRow() : 0, osis ? osis.getRow() : 0);
  if (top > comm.row) return null;
  return { top: top, left: comm.col, right: right, bottom: bottom, lowest: lowest, old: old, header: header, commRow: comm.row };
}

/**
 * The most guardians any one student has in Parents List data (header row first; A = OSIS,
 * B = guardian name), counted the way the guardian table's FILTER picks them (B not blank).
 */
function lookupBoxGuardianCount_(parentsDivided) {
  const counts = new Map();
  let max = 0;
  for (let i = 1; i < (parentsDivided || []).length; i++) {
    const osis = String(parentsDivided[i][0]).trim(), name = String(parentsDivided[i][1]).trim();
    if (!osis || !name) continue;
    const n = (counts.get(osis) || 0) + 1;
    counts.set(osis, n);
    if (n > max) max = n;
  }
  return max;
}

/** Same, read from the Parents List tab (0 when it is missing or empty). */
function lookupBoxGuardianCountFromSheet_(ss) {
  const sh = ss ? getParentsListSheet_(ss) : null;
  const last = sh ? sh.getLastRow() : 0;
  return last >= 2 ? lookupBoxGuardianCount_(sh.getRange(1, 1, last, 2).getDisplayValues()) : 0;
}

/**
 * End of a build: the lookup box's thick outline drawn again (release check 4.4: after a row was
 * inserted and deleted, a thin END bar line had replaced the thick line above the box; edits now draw
 * only the line above the END bar). Same outline as every reset (drawContactLogLookupBoxOutline_).
 */
function redrawLookupBoxOutlineForBuild_(ss) {
  const sheet = ss.getSheetByName("Contact Log");
  const endRow = sheet ? findEndRowInColumnE_(sheet) : null;
  return endRow ? drawContactLogLookupBoxOutline_(sheet, endRow) : false;
}

/** End of a build: guardian rows for the Parents List rows just built. */
function syncLookupBoxGuardianRowsForBuild_(ss, parentsDivided) {
  const sheet = ss.getSheetByName("Contact Log");
  const endRow = sheet ? findEndRowInColumnE_(sheet) : null;
  return endRow ? syncLookupBoxGuardianRows_(sheet, endRow, lookupBoxGuardianCount_(parentsDivided)) : 0;
}

/**
 * Person Spoke With choices (owner's request 2026-11-13; columns in Core.gs CL_COL_SPOKE_CHOICES).
 * Every build makes sure the formula is in Contact Log R2, because Migration clears rows 2+ in every
 * column and deleting row 2 takes the formula with it. The formula then lists rows 2 .. the row above
 * END by itself (rows added / deleted / pasted need nothing from the script).
 * - R2 has a formula (this one, or one changed by hand) -> kept. If it shows #REF! (its list is
 *   blocked, e.g. a whole row was copied and pasted, which pastes the names as plain text), the rest
 *   of R2:AA(END-1) is cleared so it can list again: those cells only ever hold the formula's list.
 * - no formula, but a row was inserted ABOVE row 2 (2026-11-14): Sheets pushed the formula down
 *   (R3, R4 ...), where it lists each row's names one row too low (its list starts at C2), so every
 *   dropdown showed the row above's guardians, and the check below took its own names for typed
 *   data. It is moved back to R2 (its own text, so one changed by hand stays as it was), when the
 *   inserted rows' R:AA are empty.
 * - no formula -> written, only when R2:AA(END-1) is empty (never over a school's own data there);
 *   then R:AA are hidden.
 * light (a row change, restoreSpokeWithChoicesAfterRowChange_): only the formula is written, and
 * R:AA hidden only when they aren't (fewer script writes after a user's delete / insert).
 * Returns "kept" / "unblocked" / "added" / "moved back" / "skipped: <why>". Called at the end of every build.
 */
/**
 * WHERE THE CHOICES LIVE (2026-11-14, release check 4.3 / 4.4 / 2.19). The formula's results used to
 * fill Contact Log R:AA on every log row, and Google then refused "Delete table row" ("You cannot insert
 * or delete cells over an array formula"); deleting row 2 took the formula, a row inserted above row 2
 * pushed it down, and two copies (one #REF!) could end up there. Now the SAME formula also sits on the
 * Backend (BACKEND_BLOCKS.SPOKE_WITH: CF2, results on Backend row n for Contact Log row n), where no
 * Contact Log row change can touch it, and nothing runs on an edit. Column D's dropdown is set BY HAND in
 * the Table: 'Backend'!$CF2:$CO2 (it was 'Contact Log'!$R2:$AA2). Every build:
 *   - writes / keeps the Backend formula (ensureSpokeChoicesOnBackend_);
 *   - reads D2's dropdown rule once (spokeDropdownSource_): it reads the Backend -> the old formula in
 *     R:AA is taken out (that frees "Delete table row"); it reads R:AA (or can't be read) -> the old
 *     formula is kept in R2 as before (ensureSpokeWithChoicesInLog_), with any extra copy removed.
 * The answer is kept in doc property SPOKE_CHOICES_MODE, so a row delete / insert reads nothing more
 * once the dropdown reads the Backend.
 * COLUMNS AFTER P (owner 2026-11-15, "checkboxes in Q+ on build when P should be the last column"): the
 * old place needs R:AA, and the build used to add them to a tab that ends at P whenever column D's rule
 * didn't read the Backend (no rule seen too); Google gives new columns what the column on their left
 * has (P: the row checkboxes), so Q:AA filled with checkboxes. Now columns are only added while column
 * D is SEEN to read 'Contact Log'!R:AA ("contactlog"; and then without P's look or checkboxes), and
 * every build tidies the columns after P (tidyContactLogAfterP_): a checkbox there is taken out, and
 * when nothing reads R:AA for sure ("backend" / "none") the empty columns at the right end are deleted,
 * so the tab ends at P again. Never deleted on "unknown" (a rule the script can't see could read R:AA:
 * deleting them would leave that rule pointing at nothing).
 */
var SPOKE_MODE_KEY_ = 'SPOKE_CHOICES_MODE'; // 'backend' | 'contactlog' | 'none' | 'unknown' (spokeDropdownSource_)
var SPOKE_RULE_ROWS_ = 50;                    // log rows read to find column D's dropdown rule
var SPOKE_NOTE_START_ = "Person Spoke With choices"; // the start of R2's cell note (the old place)

function spokeBackendFormula_() {
  return "=ARRAYFORMULA(IFERROR(SPLIT(INDEX('Contact Log'!C:C, 2):INDEX('Contact Log'!C:C, " +
    "MATCH(TRUE, TRIM('Contact Log'!E:E)=\"END\", 0) - 1), \" and \", FALSE), \"\"))";
}

/**
 * Which copy column D's dropdown reads, from the first rule on D2:D(up to 50 log rows) (one call; a row
 * pasted from elsewhere can leave D2 without one):
 *   "backend"    'Backend'!$CF2:$CO2 (DEPLOY.md step 9);
 *   "contactlog" the old 'Contact Log'!$R2:$AA2;
 *   "none"       a rule that reads no list: a dropdown of typed items, a range on another tab;
 *   "unknown"    no rule found, a dropdown from a range that can't be read, or the read failed. A Sheets
 *                Table keeps a Dropdown column's rule with the Table, and whether a script always sees it
 *                isn't known: so "no rule" never counts as "nothing reads R:AA". The old place is then
 *                kept as it is (put back as before), but no column is added for it or deleted.
 */
function spokeDropdownSource_(sheet, endRow) {
  try {
    const n = Math.max(1, Math.min(SPOKE_RULE_ROWS_, (endRow || 3) - 2));
    const rule = sheet.getRange(2, CL_COL_SPOKE_WITH, n, 1).getDataValidations().map(r => r[0]).find(Boolean);
    if (!rule) return "unknown";
    const first = (rule.getCriteriaValues() || [])[0];
    if (!first || typeof first.getSheet !== "function") {
      const type = typeof rule.getCriteriaType === "function" ? rule.getCriteriaType() : "";
      const fromRange = !type || type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE || String(type) === "VALUE_IN_RANGE";
      return fromRange ? "unknown" : "none";
    }
    const name = String(first.getSheet().getName()).trim();
    if (name === BACKEND_SHEET_NAME) return "backend";
    if (name === "Contact Log") return "contactlog";
    return "none";
  } catch (e) {
    console.warn("Person Spoke With dropdown not read: " + e.message);
    return "unknown";
  }
}

/**
 * The Backend copy: the formula in CF2 (written when missing or changed), the section title, and enough
 * Backend rows for every Contact Log row (+300 to grow until the next build: an array that runs out of
 * rows shows #REF! and every dropdown goes empty). Returns "kept" / "added" / "replaced" / "skipped: ...".
 */
function ensureSpokeChoicesOnBackend_(ss, endRow) {
  const b = BACKEND_BLOCKS.SPOKE_WITH;
  const sh = getBackendSheet_(ss);
  backendBlockReady_(sh, b, true);
  const need = Math.max(endRow || 0, 2) + 300;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  const cell = sh.getRange(2, b.col);
  const want = spokeBackendFormula_();
  const have = String(cell.getFormula() || "");
  if (have === want && String(cell.getDisplayValue()).trim() !== "#REF!") return "kept";
  if (have && have !== want && !/SPLIT\(/i.test(have)) return "skipped: Backend CF2 holds something else";
  // The results area is this section's alone: anything left in it would block the list (#REF!)
  sh.getRange(2, b.col, sh.getMaxRows() - 1, b.width).clearContent();
  cell.setFormula(want);
  return have ? "replaced" : "added";
}

/**
 * Takes the old formula out of Contact Log R (rows from..END): every cell whose formula splits column C.
 * R2's own cell note (written with the formula, SPOKE_NOTE_START_) goes with it.
 */
function removeSpokeFormulasFromLog_(sheet, endRow, fromRow) {
  fromRow = fromRow || 2;
  if (!endRow || endRow < fromRow || sheet.getMaxColumns() < CL_COL_SPOKE_CHOICES) return 0;
  const f = sheet.getRange(fromRow, CL_COL_SPOKE_CHOICES, endRow - fromRow + 1, 1).getFormulas();
  let n = 0;
  f.forEach((r, i) => {
    const text = String(r[0] || "");
    if (!/SPLIT\(/i.test(text) || !/C:C/i.test(text)) return;
    const cell = sheet.getRange(fromRow + i, CL_COL_SPOKE_CHOICES).clearContent();
    if (fromRow + i === 2 && String(cell.getNote() || "").indexOf(SPOKE_NOTE_START_) === 0) cell.clearNote();
    n++;
  });
  return n;
}

function ensureSpokeWithChoices_(ss, light) {
  const sheet = ss.getSheetByName("Contact Log");
  if (!sheet) return "skipped: no Contact Log";
  const props = PropertiesService.getDocumentProperties();
  if (light) {
    // A row delete / insert: the Backend copy needs nothing (no lag), nor a sheet whose column D reads no
    // copy; the old copy only while D reads it (or the last build couldn't tell: then no column is added)
    const stored = props.getProperty(SPOKE_MODE_KEY_);
    if (stored === "backend" || stored === "none") return stored + ": nothing to do";
    return ensureSpokeWithChoicesInLog_(ss, sheet, true, stored === "contactlog");
  }
  const endRow = findEndRowInColumnE_(sheet);
  let backend;
  try { backend = ensureSpokeChoicesOnBackend_(ss, endRow); }
  catch (e) { backend = "not written (" + e.message + ")"; console.warn("Person Spoke With choices (Backend): " + e.message); }
  const mode = spokeDropdownSource_(sheet, endRow);
  if (props.getProperty(SPOKE_MODE_KEY_) !== mode) props.setProperty(SPOKE_MODE_KEY_, mode);
  let result;
  if (mode === "backend" || mode === "none") {
    // Nothing reads the old place (Contact Log R:AA) for sure: its formula goes (that frees "Delete table row")
    const removed = removeSpokeFormulasFromLog_(sheet, endRow);
    if (removed) SpreadsheetApp.flush(); // its list leaves R:AA before the columns after P are read
    result = (mode === "backend" ? "Backend " + backend : "column D's dropdown reads no list (Backend copy " + backend + ")") +
      (removed ? "; the old formula in Contact Log R taken out (" + removed + ")" : "");
  } else {
    console.log("Person Spoke With choices: Backend copy " + backend);
    result = ensureSpokeWithChoicesInLog_(ss, sheet, false, mode === "contactlog");
  }
  // The columns after P: no checkboxes; the tab ends at P when nothing reads R:AA (owner 2026-11-15)
  try {
    const tidy = tidyContactLogAfterP_(sheet, mode === "backend" || mode === "none");
    if (tidy) result += "; " + tidy;
  } catch (e) { console.warn("Contact Log after P not tidied: " + e.message); }
  // In the build's log (Apps Script → Executions): which dropdown the build saw in column D, what it did
  console.log("Person Spoke With choices [column D reads: " + mode + "]: " + result);
  return result;
}

/**
 * The columns after P (owner 2026-11-15: "checkboxes in Q+ on build when P should be the last column"):
 *   - a checkbox there is taken out (the rule, and its TRUE / FALSE where no formula is): the code never
 *     puts one after P, Google copied P's into the columns the old place was given;
 *   - trim (nothing reads R:AA): the empty columns at the right end are deleted, so the tab ends at P
 *     (or right after the last column that holds something). Empty = no value, formula, note (R2's
 *     own excepted), merged cell or other dropdown; a column with anything in it stays, and so does
 *     every column on its left.
 * One call (the column count) once the tab ends at P; when it doesn't, one read (the rules), a few more
 * to trim. Returns what changed ("" = nothing).
 */
function tidyContactLogAfterP_(sheet, trim) {
  const first = CL_WIDTH + 1; // Q
  const maxCols = sheet.getMaxColumns();
  if (maxCols < first) return "";
  const n = maxCols - first + 1, maxRows = sheet.getMaxRows();
  const area = sheet.getRange(1, first, maxRows, n);
  const rules = area.getDataValidations();
  const isBox = rule => {
    if (!rule) return false;
    try {
      const t = rule.getCriteriaType();
      return t === SpreadsheetApp.DataValidationCriteria.CHECKBOX || String(t) === "CHECKBOX";
    } catch (e) { return false; }
  };
  let boxes = 0;
  const box = rules.map(r => r.map(rule => { const b = isBox(rule); if (b) boxes++; return b; }));
  if (!boxes && !trim) return "";

  const vals = area.getValues(), forms = area.getFormulas();
  const done = [];
  // Clears runs of marked cells, one column at a time (one call per run)
  const clearRuns = (marks, how) => {
    for (let j = 0; j < n; j++) {
      let start = -1;
      for (let i = 0; i <= maxRows; i++) {
        const on = i < maxRows && marks[i][j];
        if (on && start < 0) start = i;
        if (!on && start >= 0) {
          const r = sheet.getRange(1 + start, first + j, i - start, 1);
          if (how === "rules") r.clearDataValidations(); else r.clearContent();
          start = -1;
        }
      }
    }
  };
  if (boxes) {
    // The rules: one call when every rule after P is a checkbox (the usual case), else run by run
    const others = rules.some((r, i) => r.some((rule, j) => rule && !box[i][j]));
    if (others) clearRuns(box, "rules"); else area.clearDataValidations();
    // Their TRUE / FALSE (never a formula, never a list's names)
    const values = box.map((r, i) => r.map((b, j) => b && typeof vals[i][j] === "boolean" && !forms[i][j]));
    if (values.some(r => r.some(Boolean))) {
      clearRuns(values, "values");
      values.forEach((r, i) => r.forEach((v, j) => { if (v) vals[i][j] = ""; }));
    }
    done.push(boxes + (boxes === 1 ? " checkbox" : " checkboxes") + " after P taken out");
  }
  if (trim) {
    const notes = area.getNotes();
    let keep = first - 1; // the last column that stays (P at least)
    for (let j = n - 1; j >= 0 && keep < first; j--) {
      for (let i = 0; i < maxRows; i++) {
        const v = vals[i][j], note = String(notes[i][j] || "");
        if ((v !== "" && v !== null) || forms[i][j] || (rules[i][j] && !box[i][j]) ||
            (note && note.indexOf(SPOKE_NOTE_START_) !== 0)) { keep = first + j; break; }
      }
    }
    area.getMergedRanges().forEach(m => { keep = Math.max(keep, m.getLastColumn()); });
    if (keep < maxCols) {
      sheet.deleteColumns(keep + 1, maxCols - keep);
      done.push("empty column" + (maxCols - keep === 1 ? " " : "s ") + columnLetter_(keep + 1) +
        (maxCols - keep === 1 ? "" : ":" + columnLetter_(maxCols)) + " deleted (the tab ends at " + columnLetter_(keep) + ")");
    }
  }
  return done.join("; ");
}

/**
 * The old place (column D's dropdown still reads 'Contact Log'!$R2:$AA2): the formula in R2, as before.
 * canGrow (D is known to read R:AA): a tab narrower than AA gets its columns, plain (no P look, no
 * checkboxes). Without it no column is ever added (owner 2026-11-15: P is the Contact Log's last column).
 */
function ensureSpokeWithChoicesInLog_(ss, sheet, light, canGrow) {
  const col = CL_COL_SPOKE_CHOICES, width = CL_SPOKE_CHOICES_WIDTH, lastCol = col + width - 1;
  const maxCols = sheet.getMaxColumns();
  const anchor = maxCols >= col ? sheet.getRange(2, col) : null;
  const formula = anchor ? String(anchor.getFormula() || "") : "";
  if (formula && String(anchor.getDisplayValue()).trim() !== "#REF!") {
    if (light) return "kept";
    // A second copy further down (release check 2.19: "two arrayformulas, one with #REF!"): only R2's stays
    const extra = removeSpokeFormulasFromLog_(sheet, findEndRowInColumnE_(sheet), 3);
    return extra ? "kept (" + extra + " extra copy removed)" : "kept";
  }

  if (maxCols < col && !canGrow) return "skipped: the Contact Log ends before R (no column is added)";
  const endRow = findEndRowInColumnE_(sheet);
  if (!endRow || endRow < 3) return "skipped: no log rows above END";
  const rows = endRow - 2; // rows 2 .. END-1
  const growCols = () => {
    if (!canGrow || maxCols >= lastCol) return;
    sheet.insertColumnsAfter(maxCols, lastCol - maxCols);
    // Google gives new columns the look and the rules of the column on their left (P: the row
    // checkboxes; owner 2026-11-15, "checkboxes in Q+"): these are plain helper columns
    sheet.getRange(1, maxCols + 1, sheet.getMaxRows(), lastCol - maxCols).clearDataValidations().clearFormat();
  };
  const avail = () => Math.min(width, sheet.getMaxColumns() - col + 1); // R:AA, or the part the tab has

  if (formula) { // #REF!: something sits where the list goes (or the tab is too narrow for it)
    growCols();
    const w = avail();
    if (w > 1) sheet.getRange(2, col + 1, 1, w - 1).clearContent();
    if (rows > 1) sheet.getRange(3, col, rows - 1, w).clearContent();
    return "unblocked";
  }

  const have = Math.min(maxCols, lastCol) - col + 1; // the part of R:AA the tab already has
  const hasData = (r, n) => n > 0 && have > 0 &&
    sheet.getRange(r, col, n, have).getValues().some(row => row.some(v => v !== "" && v !== null));
  const write = (text, how) => {
    growCols();
    const cell = sheet.getRange(2, col).setFormula(text);
    if (!light) {
      cell.setNote(SPOKE_NOTE_START_ + ": one guardian per cell for column D's dropdown " +
                   "('Contact Log'!$R2:$AA2). Hidden on purpose; every build puts this formula back if it goes missing.");
      sheet.hideColumns(col, avail());
    } else if (!sheet.isColumnHiddenByUser(col)) {
      sheet.hideColumns(col, avail());
    }
    return how;
  };

  // Pushed down by rows inserted above row 2? (R3 .. R(END): a formula that splits column C)
  if (maxCols >= col) {
    const below = sheet.getRange(3, col, endRow - 2, 1).getFormulas();
    for (let i = 0; i < below.length; i++) {
      const text = String(below[i][0] || "");
      if (!/SPLIT\(/i.test(text) || !/C:C/i.test(text)) continue;
      const row = 3 + i;
      if (hasData(2, row - 2)) { // something typed into the inserted rows' R:AA: leave it all
        console.warn("Person Spoke With formula found in R" + row + ", not moved: R2:AA" + (row - 1) + " already holds data");
        return "skipped: R:AA hold data";
      }
      sheet.getRange(row, col).clearContent(); // its one-row-too-low list goes with it
      const how = write(text, "moved back");
      // any other copy further down goes too (only one may list the names)
      if (!light) { const extra = removeSpokeFormulasFromLog_(sheet, endRow, row + 1); if (extra) return how + " (" + extra + " extra copy removed)"; }
      return how;
    }
  }

  if (hasData(2, rows)) {
    console.warn("Person Spoke With choices not added: Contact Log R2:AA" + (endRow - 1) + " already holds data");
    return "skipped: R:AA hold data";
  }
  return write(SPOKE_WITH_CHOICES_FORMULA, "added");
}

/**
 * After rows are deleted from / inserted into the Contact Log (owner's request 2026-11-14, "only row
 * 2"): the Person Spoke With formula lives only in R2. Deleting row 2 takes it with it (every row's
 * dropdown goes empty); a row inserted above row 2 pushes it down (every dropdown one row off). The
 * change event doesn't say which row, so this reads R2 (one call, the only cost on any other delete /
 * insert); no formula there -> ensureSpokeWithChoices_(ss, light) puts it back now (same checks: never
 * over anything typed into R:AA; only the formula is written). Returns what happened.
 */
function restoreSpokeWithChoicesAfterRowChange_(ss, sheet) {
  // Column D reads the Backend copy (set by hand; the last build saw it), or no copy at all: a row
  // change can't touch it
  const mode = PropertiesService.getDocumentProperties().getProperty(SPOKE_MODE_KEY_);
  if (mode === "backend" || mode === "none") return mode + ": nothing to do";
  let formula;
  try { formula = sheet.getRange(2, CL_COL_SPOKE_CHOICES).getFormula(); }
  catch (e) { return "skipped: no column R"; } // a tab that ends before R (the Contact Log ends at P)
  if (formula) return "kept"; // row 2 is still there
  return ensureSpokeWithChoices_(ss, true);
}

/**
 * Makes the lookup box's guardian table exactly as tall as maxGuardians needs (at least
 * LOOKUP_BOX_GUARDIAN_ROWS_), so a student with a 3rd / 4th / 5th guardian (New/Edit Student)
 * does not spill into the small row above Communications (#REF!), and the extra rows go again
 * when no student needs them. Rows are added / deleted under the last guardian row:
 * - added rows: same height as that row, no line between guardian rows, white / #fff8ee
 *   taking turns (guardian 3 white, 4 cream, 5 white ...), the box's thick sides, and the thick
 *   line under the new last row;
 * - a row is only deleted when it holds no formula; the thick line goes under the new last row;
 * - every ROW()-n / ROW()+n in the box's formulas (Communications, Student Info: they count rows
 *   up to the OSIS box) is re-counted so it still points at the same cell.
 * Only in the owner's layout (guardian table above Communications, a small row between them).
 * Returns the number of rows added (negative = deleted), 0 when nothing changed.
 */
function syncLookupBoxGuardianRows_(sheet, endRow, maxGuardians) {
  const box = findContactLogLookupBox_(sheet, endRow);
  if (!box || box.old || !box.header || box.header >= box.commRow) return 0;

  // Guardian rows: under the header down to the small row(s) right above Communications
  let last = box.commRow - 1;
  while (last > box.header && sheet.getRowHeight(last) <= LOOKUP_BOX_SPACER_MAX_PX_) last--;
  if (last === box.commRow - 1) return 0; // no small row above Communications: not the layout we know
  const have = last - box.header;
  if (have < 1) return 0;
  const want = Math.min(LOOKUP_BOX_MAX_GUARDIAN_ROWS_, Math.max(LOOKUP_BOX_GUARDIAN_ROWS_, maxGuardians || 0));
  if (want === have) return 0;

  const cols = box.right - box.left + 1;
  const black = "#000000", thick = SpreadsheetApp.BorderStyle.SOLID_THICK;
  const maxCols = sheet.getMaxColumns();

  if (want > have) {
    const n = want - have;
    const height = sheet.getRowHeight(last);
    sheet.insertRowsAfter(last, n);
    sheet.setRowHeights(last + 1, n, height);
    const colors = [];
    for (let i = 0; i < n; i++) colors.push(new Array(cols).fill(LOOKUP_BOX_ROW_COLORS_[(have + i) % 2]));
    sheet.getRange(last + 1, box.left, n, cols).setBackgrounds(colors);
    // No lines between the guardian rows; the box's thick sides; the thick line under the last row
    sheet.getRange(last, box.left, n + 1, cols).setBorder(null, null, null, null, null, false);
    sheet.getRange(last + 1, box.left, n, cols).setBorder(null, true, true, true, null, null, black, thick);
    recountLookupBoxRowOffsets_(sheet, box.top, box.bottom + n, maxCols,
      R => R <= last ? R : R > last + n ? R - n : null,   // where a row was before
      r => r <= last ? r : r + n);                         // where an old row is now
    return n;
  }

  // Fewer: delete the extra rows from the bottom (not if one of them holds a formula: the
  // guardian table is one formula in the header row, so its rows never do)
  const k = have - want;
  const first = last - k + 1;
  if (sheet.getRange(first, 1, k, maxCols).getFormulas().some(row => row.some(f => f !== ""))) {
    console.warn("Lookup box: extra guardian rows " + first + "-" + last + " hold a formula; not deleted");
    return 0;
  }
  sheet.deleteRows(first, k);
  sheet.getRange(first - 1, box.left, 1, cols).setBorder(null, null, true, null, null, null, black, thick);
  recountLookupBoxRowOffsets_(sheet, box.top, box.bottom - k, maxCols,
    R => R < first ? R : R + k,
    r => r < first ? r : r > last ? r - k : null);
  return -k;
}

/**
 * After rows were added / deleted inside the lookup box: re-counts every ROW()-n / ROW()+n
 * in the formulas of rows fromRow..toRow so each still points at the same cell
 * (oldRowOf(R) = where row R was, null = a new row; newRowOf(r) = where old row r is now,
 * null = deleted, which is left as it was).
 */
function recountLookupBoxRowOffsets_(sheet, fromRow, toRow, maxCols, oldRowOf, newRowOf) {
  toRow = Math.min(toRow, sheet.getMaxRows());
  if (toRow < fromRow) return;
  const formulas = sheet.getRange(fromRow, 1, toRow - fromRow + 1, maxCols).getFormulas();
  formulas.forEach((rowFormulas, i) => {
    const R = fromRow + i, r = oldRowOf(R);
    if (r === null) return;
    rowFormulas.forEach((f, j) => {
      if (!f) return;
      const fixed = f.replace(/(^|[^A-Za-z0-9_.])ROW\(\s*\)\s*([+-])\s*(\d+)/gi, (m, pre, sign, n) => {
        const target = newRowOf(sign === "-" ? r - Number(n) : r + Number(n));
        if (target === null) return m;
        const d = R - target;
        return pre + "ROW()" + (d > 0 ? "-" + d : d < 0 ? "+" + (-d) : "");
      });
      if (fixed !== f) sheet.getRange(R, j + 1).setFormula(fixed);
    });
  });
}

/**
 * The quick notes update for one Contact Log row's OSIS (a sibling group
 * "111, 222" updates the group row and then each sibling). Returns true when it
 * ran, false when "update notes as I type" is off in Settings (the caller then
 * shows the Fast Entry Mode toast at the end of the edit).
 */
function syncNotesForRowOsis_(targetOsis, early, extraOsis) {
  if (PropertiesService.getDocumentProperties().getProperty('autoNotes') === 'false') return false;
  const osisList = [];
  const add = (o) => {
    if (!o || osisList.indexOf(o) !== -1) return;
    osisList.push(o);
    if (o.includes(",")) {
      o.split(",").map(x => x.trim()).forEach(individualOsis => {
        if (isValidOsis_(individualOsis) && osisList.indexOf(individualOsis) === -1) osisList.push(individualOsis);
      });
    }
  };
  add(targetOsis);
  (extraOsis || []).forEach(add); // e.g. the OSIS this row had before it was cleared / typed over
  if (!early) { updateOsisDeltas_(osisList); return true; }

  // Early (right after the ghost typist): only the quick one-cell Combined write,
  // so the search bar formula updates as soon as possible. On the owner's sheet
  // the heavy update (Combined + Phone Contacts) waits for the end of the edit.
  const owner = getOwnerHooks_();
  const quick = updateCombinedNotesQuick_(osisList, !owner);
  PENDING_OWNER_NOTES_ = owner ? { list: osisList, combinedAlreadyWritten: quick.missing.length === 0 } : null;
  return true;
}


// OSIS list whose owner-sheet notes update (Phone Contacts) is still to run at the end of this edit
var PENDING_OWNER_NOTES_ = null;

// =========================================================
// GHOST TYPIST DATA CACHE
// The Contact Log ghost typist needs Student List A5:C and Directory C:E on
// every edit. Only a build writes those tabs, so the first edit reads them and
// keeps a copy in the document cache for 6 hours; the next edits skip both
// reads. Every build clears it (clearGhostTypistCache_). A copy bigger than one
// cache key (100 KB, a big school) is split over several keys (bigCachePut_,
// Core.gs; 2026-11-18): before, it was read from the tabs on EVERY edit.
// =========================================================
const GHOST_CACHE_KEY_ = "GHOST_TYPIST_DATA";
const GHOST_CACHE_SECONDS_ = 21600; // 6 hours (the most the cache keeps); every build clears it

/** ghostRaw: the cached JSON already fetched by contactLogPrefetch_ (skips a cache call), or undefined. */
function ghostTypistData_(ss, ghostRaw) {
  let cache = null;
  try {
    cache = CacheService.getDocumentCache();
    const hit = ghostRaw !== undefined ? ghostRaw : (cache && bigCacheGet_(cache.getAll(bigCacheKeys_(GHOST_CACHE_KEY_)), GHOST_CACHE_KEY_));
    if (hit) return JSON.parse(hit);
  } catch (err) { console.warn("Ghost typist cache not read: " + err.message); }

  const dir = ss.getSheetByName("Directory");
  const dLast = dir ? dir.getLastRow() : 0;
  const data = {
    master: studentLookupRows_(ss, 3), // OSIS / Student / Guardian incl. sibling groups (Backend copy)
    dir: dLast > 0 ? dir.getRange(1, 3, dLast, 3).getValues() : []
  };
  try {
    if (cache) bigCachePut_(cache, GHOST_CACHE_KEY_, JSON.stringify(data), GHOST_CACHE_SECONDS_);
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
function contactLogPrefetch_(sheet, row) {
  const endKey = endRowCacheKey_(sheet);
  let got = {};
  try { got = CacheService.getDocumentCache().getAll([endKey].concat(bigCacheKeys_(GHOST_CACHE_KEY_))) || {}; } catch (err) { got = {}; }

  let endRow = null;
  const remembered = parseInt(got[endKey], 10);
  let endFromCache = false;
  let rowVals = null; // the edited row's B..N (display), when the END check could read it too
  if (remembered > 1) {
    try {
      if (row && row >= 2 && row < remembered && remembered - row <= 60) {
        // ONE read from the edited row down to the remembered END: checks END
        // (column E of the last row) and gives the auto-fill its B..N values
        const span = sheet.getRange(row, 2, remembered - row + 1, 13).getDisplayValues();
        if (String(span[span.length - 1][CL_COL_OSIS - 2]).trim().toUpperCase() === "END") {
          endRow = remembered; endFromCache = true; rowVals = span[0];
        }
      } else if (String(sheet.getRange(remembered, CL_COL_OSIS).getDisplayValue()).trim().toUpperCase() === "END") {
        endRow = remembered; endFromCache = true;
      }
    } catch (err) { /* row no longer exists: scan below */ }
  }
  if (!endRow) {
    endRow = findEndRowInColumnE_(sheet);
    rememberEndRow_(sheet, endRow);
  }
  return { endRow: endRow, endFromCache: endFromCache, rowVals: rowVals, ghostRaw: bigCacheGet_(got, GHOST_CACHE_KEY_) };
}

// =========================================================
// NOTES SOURCES CACHE
// A student's notes also need the Backend Event Log and the Notes tab. A
// Contact Log edit never changes those, but reading them cost ~1.5 s each on a
// big sheet. So they are kept in the document cache for 6 hours; any
// Event Log write (the backend* helpers in Core.gs), any Events / Notes edit
// and every build clear it. A copy bigger than one cache key is split over
// several (bigCachePut_, Core.gs; 2026-11-18). A Notes cell holding a date:
// read every time, as before.
// =========================================================
const NOTES_SRC_EVENTS_KEY_ = "NOTES_SRC_EVENTS";
const NOTES_SRC_NOTES_KEY_ = "NOTES_SRC_NOTES";

/** { eventData, nData }: Event Log (display values) and Notes tab (values), cached when possible. */
function notesSourcesData_(ss) {
  let cache = null, got = {};
  try {
    cache = CacheService.getDocumentCache();
    got = cache.getAll(bigCacheKeys_(NOTES_SRC_EVENTS_KEY_).concat(bigCacheKeys_(NOTES_SRC_NOTES_KEY_))) || {};
  } catch (err) { got = {}; }
  const keep = (key, value) => {
    try {
      if (cache) bigCachePut_(cache, key, JSON.stringify(value), 21600); // 6 hours; cleared on every change
    } catch (err) { /* read again next time */ }
  };
  const cachedEvents = bigCacheGet_(got, NOTES_SRC_EVENTS_KEY_), cachedNotes = bigCacheGet_(got, NOTES_SRC_NOTES_KEY_);

  let eventData;
  if (cachedEvents) eventData = JSON.parse(cachedEvents);
  else { eventData = backendReadBlock_(ss, "EVENT_LOG", true); keep(NOTES_SRC_EVENTS_KEY_, eventData); }

  let nData;
  if (cachedNotes) nData = JSON.parse(cachedNotes);
  else {
    const notesSheet = ss.getSheetByName("Notes");
    nData = notesSheet ? notesSheet.getDataRange().getValues() : null;
    // JSON would turn a date into different text, so only plain values are cached
    if (nData && !nData.some(r => r.some(v => v instanceof Date))) keep(NOTES_SRC_NOTES_KEY_, nData);
  }
  clDbg_("notes sources: Event Log " + (cachedEvents ? "cached" : "read") + ", Notes tab " + (cachedNotes ? "cached" : "read"));
  return { eventData: eventData, nData: nData };
}

function clearNotesSourcesCache_() {
  try { CacheService.getDocumentCache().removeAll(bigCacheKeys_(NOTES_SRC_EVENTS_KEY_).concat(bigCacheKeys_(NOTES_SRC_NOTES_KEY_))); } catch (err) { console.warn("Notes sources cache not cleared: " + err.message); }
}

function clearGhostTypistCache_() {
  try { bigCacheRemove_(CacheService.getDocumentCache(), GHOST_CACHE_KEY_); } catch (err) { console.warn("Ghost typist cache not cleared: " + err.message); }
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

/**
 * Typing in the LAST empty row above END adds CL_NEW_ROWS_ (3) empty rows under it (owner 2026-11-18; was 1):
 * the next two entries need no insert, which is one of the slowest steps of an edit. Returns how many rows
 * were added (0 = none; END moved down that many).
 */
var CL_NEW_ROWS_ = 3;
function autoAddBlankRow_(sheet, editedRow, endRow) {
  if (!endRow || editedRow !== endRow - 1) return 0;
  const n = CL_NEW_ROWS_;
  sheet.insertRowsAfter(editedRow, n);
  formatContactLogRows_(sheet, editedRow + 1, n);
  sheet.getRange(editedRow + 1, CL_COL_CHECKBOX, n, 1).setValue(false);

  // The end bar's black border sticks to the edited row's bottom edge when
  // the new row is inserted above it. Strip its horizontal borders so only
  // the dark brown end bar (re-drawn by the onEdit sweep) keeps a border.
  sheet.getRange(editedRow, 1, 1, CL_WIDTH).setBorder(null, null, false, null, null, false);
  return n;
}

/**
 * Contact Log Type column (L): each type in the Type List (Backend BL) gets a
 * text color, the same one every time: its place in the list picks it from
 * SHEET_LOOK_.typeColors (Core.gs). Before 2026-11-12 the colors were random,
 * so every reset and 10 PM run (resetContactLogFormatting clears the rules)
 * gave the types new colors. Every rule that covers only column L is a Type
 * rule and is drawn again (as before, a type taken off the list loses its
 * rule); the other rules are kept. Writes nothing when the rules already match.
 */
function applyTypeDropdownColors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const contactLog = ss.getSheetByName('Contact Log');
  if (!contactLog) return;
  // Type List section of the Backend tab (BL2:BL)
  const typeValues = backendReadBlock_(ss, "TYPES", true).flat().map(v => String(v)).filter(String);
  const uniqueTypes = typeValues.filter((t, i) => typeValues.indexOf(t) === i);
  const palette = SHEET_LOOK_.typeColors;
  const wanted = uniqueTypes.map((type, i) => ({ type: type, color: palette[i % palette.length] }));

  const existingRules = contactLog.getConditionalFormatRules();
  const otherRules = [];
  const typeNow = []; // what the column L rules say now: { type, color }, or null for any other L rule
  existingRules.forEach(rule => {
    const ranges = rule.getRanges();
    const onlyL = ranges.length > 0 && ranges[0].getColumn() === CL_COL_TYPE && ranges[0].getLastColumn() === CL_COL_TYPE;
    if (!onlyL) { otherRules.push(rule); return; }
    let entry = null;
    try {
      const condition = rule.getBooleanCondition();
      if (condition && condition.getCriteriaType() === SpreadsheetApp.BooleanCriteria.TEXT_EQUAL_TO) {
        entry = { type: String(condition.getCriteriaValues()[0]), color: String(condition.getFontColor() || "").toLowerCase() };
      }
    } catch (e) { console.error(e); }
    typeNow.push(entry);
  });

  // Already right (same types, same order, same colors, nothing else on L): no write
  const same = typeNow.length === wanted.length && wanted.every((w, i) =>
    typeNow[i] && typeNow[i].type === w.type && typeNow[i].color === w.color.toLowerCase());
  if (same) return;

  const typeColumnRange = contactLog.getRange('L2:L');
  const typeRules = wanted.map(w => SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(w.type)
    .setFontColor(w.color)
    .setRanges([typeColumnRange])
    .build());
  contactLog.setConditionalFormatRules([...typeRules, ...otherRules]);
}

/**
 * Contact Log "Method missing" color (owner 2026-11-23): a row with a Date but no Method (J) is
 * light red until a Method is picked. resetContactLogFormatting (every 10 PM run) draws it with
 * the other rules; every build also adds it when the Contact Log doesn't have it, so a sheet whose
 * nightly run is off gets it too. One read; it writes only when the rule is missing, and a rule
 * someone recolored by hand (same formula) is left alone.
 */
var METHOD_MISSING_FORMULA_ = '=AND($B2<>"", $J2="")';

function methodMissingRule_(sheet) {
  return SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(METHOD_MISSING_FORMULA_)
    .setBackground(SHEET_LOOK_.missingMethod)
    .setRanges([sheet.getRange("J2:J" + sheet.getMaxRows())]).build();
}

function hasMethodMissingRule_(rules) {
  const want = METHOD_MISSING_FORMULA_.replace(/\s+/g, "");
  return rules.some(rule => {
    try {
      const c = rule.getBooleanCondition();
      return !!c && String((c.getCriteriaValues() || [])[0] || "").replace(/\s+/g, "") === want;
    } catch (e) { return false; }
  });
}

/** Returns true when the rule was added. */
function ensureMethodMissingRule_(ss) {
  const sheet = ss.getSheetByName("Contact Log");
  if (!sheet) return false;
  const rules = sheet.getConditionalFormatRules();
  if (hasMethodMissingRule_(rules)) return false;
  rules.push(methodMissingRule_(sheet));
  sheet.setConditionalFormatRules(rules);
  return true;
}

/**
 * Contact Log "newer contact" color (owner 2026-11-24): a Follow-up (N) that is still open turns
 * orange when the same student (OSIS, E) has a NEWER interaction in the log: a later date anywhere
 * (at least a second later), or the same date further down. Only the older row turns orange, never
 * the newest one; marking it "Skipped / Newer Interaction" or "Followed up / Resolved" clears it.
 * Rows under END (the lookup box) have no date, so they never count. Open = the values the hide rule
 * keeps showing (getContactLogHideFormula_, Regenerate.gs). The second of slack: Sheets turns a date
 * with a time into text with ~15 digits, so "later than this row's own date" could count the row itself.
 * Drawn like the Method color: resetContactLogFormatting (the 10 PM run), and every build when missing.
 * Checked with a formula engine on sample rows (newer below / above, backdated, same-day pairs, text OSIS).
 */
var FOLLOWUP_OPEN_VALUES_ = ["Followup Needed", "Followup Needed - ASAP", "Waiting for Response", "Need to respond", "Waiting - Check Up"];
var FOLLOWUP_NEWER_FORMULA_ = '=IF(OR(' + FOLLOWUP_OPEN_VALUES_.map(v => 'TRIM($N2)="' + v + '"').join(", ") + '), ' +
  'IF(AND($E2<>"", ISNUMBER($B2)), COUNTIFS($E$2:$E, $E2, $B$2:$B, ">"&($B2+1/86400)) + ' +
  'COUNTIFS($E3:$E, $E2, $B3:$B, ">="&($B2-1/86400)) > 0))';

function followupNewerRule_(sheet) {
  return SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(FOLLOWUP_NEWER_FORMULA_)
    .setBackground(SHEET_LOOK_.followupNewer)
    .setRanges([sheet.getRange("N2:N" + sheet.getMaxRows())]).build();
}

function hasFollowupNewerRule_(rules) {
  const want = FOLLOWUP_NEWER_FORMULA_.replace(/\s+/g, "");
  return rules.some(rule => {
    try {
      const c = rule.getBooleanCondition();
      return !!c && String((c.getCriteriaValues() || [])[0] || "").replace(/\s+/g, "") === want;
    } catch (e) { return false; }
  });
}

/** Returns true when the rule was added (first in the list, so it wins over the Workshop row color). */
function ensureFollowupNewerRule_(ss) {
  const sheet = ss.getSheetByName("Contact Log");
  if (!sheet) return false;
  const rules = sheet.getConditionalFormatRules();
  if (hasFollowupNewerRule_(rules)) return false;
  rules.unshift(followupNewerRule_(sheet));
  sheet.setConditionalFormatRules(rules);
  return true;
}

// ======================================================================
// METHOD 2 CHOICES (owner 2026-11-23): Contact Log column K's dropdown = the fixed choices below +
// every Method 2 already typed above END, so the list grows by itself. ONE formula in Backend CQ3
// (BACKEND_BLOCKS.METHOD2), written by every build when it is missing. The dropdown on K is set BY
// HAND in the Contact Log table: ='Backend'!$CQ$3:$CQ with "Show a warning" (DEPLOY.md). To change
// the fixed choices, edit the formula in CQ3 on the sheet: a formula there that reads Contact Log K
// is kept as it is.
// ======================================================================
// Always first in the list, whatever is typed (owner 2026-11-24; was Call, Text, Email, Voicemail, Meeting, In School)
var METHOD2_FIXED_CHOICES_ = ["Call", "Email", "Voicemail", "Meeting", "In School", "Parents Square"];

function method2ChoicesFormula_() {
  const fixed = METHOD2_FIXED_CHOICES_.map(c => '"' + c + '"').join(", ");
  return "=ARRAYFORMULA(LET(endRow, IFERROR(MATCH(TRUE, TRIM('Contact Log'!E:E)=\"END\", 0), 3), " +
    "typed, TRIM(INDEX('Contact Log'!K:K, 2):INDEX('Contact Log'!K:K, MAX(2, endRow - 1))), " +
    "all, UNIQUE(VSTACK(" + fixed + ", typed)), FILTER(all, all <> \"\")))";
}

/** "added" | "kept" | "unblocked" | "skipped: ..." (never throws on a sheet that has something else there). */
function ensureMethod2ChoicesOnBackend_(ss) {
  const b = BACKEND_BLOCKS.METHOD2;
  const sh = getBackendSheet_(ss);
  backendBlockReady_(sh, b, true);
  const header = sh.getRange(BACKEND_HEADER_ROW, b.col);
  if (String(header.getValue()).trim() === "") header.setValue("Method 2").setFontWeight("bold");
  const first = BACKEND_HEADER_ROW + 1;
  const cell = sh.getRange(first, b.col);
  const have = String(cell.getFormula() || "");
  if (have && /Contact Log'?!\$?K/i.test(have)) {
    // Kept (one changed by hand too); a list blocked by something typed under it (#REF!) is freed
    if (String(cell.getDisplayValue()).trim() !== "#REF!") return "kept";
    const below = sh.getMaxRows() - first;
    if (below > 0) sh.getRange(first + 1, b.col, below, 1).clearContent();
    return "unblocked";
  }
  const where = "Backend " + columnLetter_(b.col) + first;
  if (have) return "skipped: " + where + " holds another formula";
  // Empty: written only when nothing was typed in that column either (never over someone's values)
  const n = Math.max(1, sh.getLastRow() - BACKEND_HEADER_ROW);
  if (sh.getRange(first, b.col, n, 1).getValues().some(r => String(r[0]).trim() !== "")) {
    return "skipped: " + where + " and below hold values";
  }
  cell.setFormula(method2ChoicesFormula_());
  return "added";
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
/** preRowVals: the row's B..N display values already read this edit (skips the read), or null. */
function autofillContactLogDefaults_(sheet, row, col, endRow, editedValue, preRowVals) {
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
  const rowVals = preRowVals || sheet.getRange(row, 2, 1, 13).getDisplayValues()[0]; // B..N

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
 * Reads the Student List's OSIS / Student / Guardian ONCE (the Backend copy, with
 * the sibling-group rows) and returns lookup(value, col) for pasted values,
 * col 0 = OSIS, 1 = Student, 2 = Guardian. Case/spacing don't matter, and a
 * single OSIS or parent name matches a sibling group / multi-parent cell.
 */
function buildGhostLookup_(ss) {
  const data = studentLookupRows_(ss, 3);
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
  ss.toast(`Filled ${filled}. Not found in the Student List: ${shown}`, "Ghost Typist", 8);
}

/**
 * Contact Log: fills Guardian / OSIS / Student / Site / Class for every pasted row.
 * numCols: a paste wider than one column that also covers the OSIS column (E) looks a
 * row up by its own OSIS when it has one (2026-11-12): looked up by the parent's name,
 * rows of a parent with two children all got the first child's OSIS.
 */
function fillPastedContactLogRows_(ss, sheet, row, numRows, col, endRow, numCols) {
  const lookup = buildGhostLookup_(ss);
  const siteClass = buildDirectorySiteClassLookup_(ss);
  const startCol = (col === CL_COL_OSIS) ? 0 : (col === CL_COL_STUDENT) ? 1 : 2;
  const lastCol = col + (numCols || 1) - 1;
  const osisPasted = col !== CL_COL_OSIS && col <= CL_COL_OSIS && lastCol >= CL_COL_OSIS;

  const updateRange = sheet.getRange(row, 3, numRows, 7); // C:I
  const rows = updateRange.getValues();
  const missing = [];
  const filledRows = [];

  rows.forEach((r, i) => {
    let val = String(r[col - 3]).trim();
    let matchCol = startCol;
    const ownOsis = String(r[CL_COL_OSIS - 3]).trim();
    if (osisPasted && ownOsis) { val = ownOsis; matchCol = 0; }
    if (!val || val.toUpperCase() === "TRUE" || val.toUpperCase() === "FALSE") return;
    const m = lookup(val, matchCol);
    if (!m) { missing.push(val); return; }
    const sc = siteClass(m.osis);
    const byCol = matchCol === 0 ? CL_COL_OSIS : matchCol === 1 ? CL_COL_STUDENT : CL_COL_GUARDIAN;
    if (byCol !== CL_COL_GUARDIAN) r[0] = m.guardian; // C
    if (byCol !== CL_COL_OSIS) r[2] = m.osis;         // E
    if (byCol !== CL_COL_STUDENT) r[3] = m.student;   // F
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
  const added = autoAddBlankRow_(sheet, row + numRows - 1, endRow);

  toastGhostMisses_(ss, filledRows.length, missing);
  return added; // rows added (END moved down that many): the edit's END clean-up must know (it cleared the new row's checkbox)
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

  // OSIS (A) cleared or typed over: the student(s) it was lose this row's note (2026-11-12:
  // the note stayed in their notes until the next build). The typist below adds it to the new one.
  if (col === 1 && e.oldValue !== undefined && String(e.oldValue).trim() && String(e.oldValue).trim() !== editedValue) {
    try { updateOsisDeltas_(String(e.oldValue).split(",").map(o => o.trim()).filter(Boolean)); }
    catch (err) { console.warn("Notes: old OSIS not updated: " + err.message); }
  }

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
      // The question couldn't be shown: never delete a note without a yes (2026-11-12)
      isConfirmed = false;
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
    const tNotesTypist = Date.now();
    const lookupSrc = studentLookupSource_(e.source); // Student List incl. sibling groups (Backend copy)
    if (lookupSrc) {
      // Typed OSIS (A) -> Student List A, Guardian (B) -> C, Student (C) -> B
      const found = findInMasterTable_(lookupSrc, col === 1 ? 1 : col === 2 ? 3 : 2, editedValue);
      clDbg_("Notes typist lookup " + (Date.now() - tNotesTypist) + " ms: " + (found ? "found OSIS " + found.osis : "'" + editedValue + "' not found"));

      if (found) {
        // Pull exact data from Student List
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

        // Write only the 2 identity cells it fills (no read of the row first), then
        // paint them before the checkbox and the notes sync below
        if (col !== 1) sheet.getRange(row, 1).setValue(correctOsis);      // Col 1 (OSIS)
        if (col !== 2) sheet.getRange(row, 2).setValue(correctGuardian);  // Col 2 (Guardian)
        if (col !== 3) sheet.getRange(row, 3).setValue(correctStudent);   // Col 3 (Student Name)
        SpreadsheetApp.flush();
        clDbg_("⏱ Notes typist " + (Date.now() - tNotesTypist) + " ms (painted)");
        
        // Ensure a checkbox is securely placed in Column E if one isn't there already
        sheet.getRange(row, 5).insertCheckboxes().setFontColor("#000000");

        // A note typed before the student was filled in syncs now that the row has an OSIS
        syncNotesRows_(sheet, row, 1, true);
        clDbg_("Notes: notes sync done");

      } else {
        e.source.toast("Could not find '" + editedValue + "' in the Student List.", "Ghost Typist Failed", 4);
      }
    }
  }
}

/**
 * Single-cell ghost typist lookup (Notes and Events): finds a typed value in one
 * Student List column (1 = OSIS, 2 = Student, 3 = Guardian; whole cell, any case)
 * of src (studentLookupSource_: the Backend copy, sibling groups included) and
 * returns that row's { osis, student, guardian }, or null.
 */
function findInMasterTable_(src, masterCol, value) {
  const found = studentLookupColumn_(src, masterCol - 1).createTextFinder(value).matchEntireCell(true).matchCase(false).findNext();
  if (found) {
    const vals = src.sheet.getRange(found.getRow(), src.col, 1, 3).getValues()[0];
    return { osis: vals[0], student: vals[1], guardian: vals[2] };
  }
  // Not found as typed: try again ignoring stray spaces ("Ann  Smith " = "Ann Smith", typed
  // or in the Student List; owner 2026-11-12). Only read when the quick lookup misses.
  const want = typistKey_(value);
  if (!want) return null;
  const hit = studentLookupRows_(null, 3, src).find(r => typistKey_(r[masterCol - 1]) === want);
  return hit ? { osis: hit[0], student: hit[1], guardian: hit[2] } : null;
}

/** How the ghost typists compare a typed value with a Student List cell: case and stray spaces don't count. */
function typistKey_(v) {
  return String(v === null || v === undefined ? "" : v).trim().replace(/\s+/g, " ").toLowerCase();
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
  const nycsaHeaders = ["Location Code","School District","School Name","Class Code","Grade Level","Student ID","Student First Name","Student Last Name","Guardian First Name","Guardian Last Name","Another Primary NYCSA Account Exist","Guardian Home Number","Guardian Work Number","Guardian Cell Number","Relationship","Role","Account Creation Code"];
  
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
/**
 * opts.combinedAlreadyWritten: the quick write already rewrote these students'
 * Combined rows this edit, so the owner's update only does Phone Contacts.
 */
function updateOsisDeltas_(osisList, opts) {
  // The owner's sheet also rewrites that student's Phone Contacts rows
  const owner = getOwnerHooks_();
  if (owner) {
    clDbg_("notes update (owner version, z_MINE.gs: " + (opts && opts.combinedAlreadyWritten ? "Phone Contacts only" : "Combined + Phone Contacts") + ") for " + osisList.join(" | "));
    const out = owner.updateOsisDeltas(osisList, opts || {});
    clDbg_("notes update (owner version) finished");
    return out;
  }
  return updateCombinedNotesQuick_(osisList, true);
}

/**
 * The quick notes update: rewrites only each student's Combined Contact
 * Tracking cell (D) from the logs. allowRebuild: a student with no Combined row
 * yet rebuilds Combined + Phone Contacts (shared layout, so never on the owner's
 * sheet: there the owner's update at the end of the edit adds the row).
 */
function updateCombinedNotesQuick_(osisList, allowRebuild) {
  // One edit at a time: two quick edits running side by side could otherwise
  // write older notes over newer ones. The writes are pushed to the sheet before
  // the lock is let go, so the next edit reads them.
  let lock = null, locked = false;
  try { lock = LockService.getDocumentLock(); locked = lock.tryLock(20000); } catch (err) { locked = false; }
  if (!locked) clDbg_("notes: another edit held the lock for 20 s, writing anyway");
  try {
    const out = updateCombinedNotesQuickLocked_(osisList, allowRebuild);
    SpreadsheetApp.flush();
    return out;
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (err) { /* released when the run ends */ } }
  }
}

function updateCombinedNotesQuickLocked_(osisList, allowRebuild) {
  const targets = osisList.filter(o => o);
  clDbg_("notes update for " + targets.join(" | "));
  const texts = {}, missing = [];
  if (!targets.length) return { texts: texts, missing: missing };
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const props = PropertiesService.getDocumentProperties();
  const disableHiddenNotes = props.getProperty('disableHiddenNotes') === 'true';

  // Read each source once; none of them change while the notes are rewritten.
  // Event Log + Notes tab come from a short cache (notesSourcesData_).
  const contactSheet = ss.getSheetByName("Contact Log");
  const contactData = contactSheet ? contactSheet.getDataRange().getDisplayValues() : null;
  clDbg_("notes: Contact Log read (" + (contactData ? contactData.length : 0) + " rows)");
  const sources = notesSourcesData_(ss);
  const eventData = sources.eventData;
  const nData = sources.nData;

  let combinedOsisColumn = null; // Combined Contact Tracking column A (re-read after a rebuild)
  // Only needed for a student with no Combined row yet (read once, usually from the ghost typist cache)
  let knownOsis = null;
  const isKnown = osis => {
    if (!knownOsis) knownOsis = knownStudentOsis_(ss);
    return knownOsis.has(normalizeOsis_(osis));
  };

  targets.forEach(targetOsis => {
    const finalNotesList = notesTextForOsis_(targetOsis, contactData, eventData, nData, disableHiddenNotes);
    texts[targetOsis] = finalNotesList;

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
        clDbg_("Combined row " + (i + 1) + " (OSIS " + targetOsis + ") rewritten: " + finalNotesList.length + " characters");
        foundInBackend = true;
        break;
      }
    }

    // Same rule as the build: a student who isn't on the Student List gets no row (and no rebuild
    // on every edit, which could never add it)
    if (!foundInBackend && finalNotesList && !isKnown(targetOsis)) {
      clDbg_("OSIS " + targetOsis + " is not on the Student List: no Combined row");
      return;
    }
    if (!foundInBackend && finalNotesList) missing.push(targetOsis);
    if (!foundInBackend) clDbg_("OSIS " + targetOsis + " has no Combined row" + (!finalNotesList ? " and no notes (nothing to write)" : allowRebuild ? " -> rebuilding Combined + Phone Contacts" : " (added by the owner's update at the end)"));
    if (!foundInBackend && finalNotesList && allowRebuild) {
      ss.toast("New student profile syncing to backend...", "Auto-Build", 3);
      updateContactAndPhoneOnly();
      combinedOsisColumn = null; // that rebuild rewrote the tab
    }
  });
  return { texts: texts, missing: missing };
}

/**
 * OSIS keys (normalizeOsis_) of every student who may have a Combined Contact
 * Tracking row: the Student List (the ghost typist's cached copy, so usually no
 * sheet read) plus, on the owner's sheet, the Old Master Table.
 */
function knownStudentOsis_(ss) {
  const known = knownOsisSet_(ghostTypistData_(ss).master);
  const owner = getOwnerHooks_();
  if (owner) owner.oldMasterTableOsis(ss).forEach(o => known.add(o));
  return known;
}

/**
 * One student's notes text (what Combined Contact Tracking column D holds after a
 * quick update) from already-read Contact Log (display values), Backend Event
 * Log and Notes data. Used by the quick update and the search bar's
 * Communications box, so both always show the same text.
 */
function notesTextForOsis_(targetOsis, contactData, eventData, nData, disableHiddenNotes) {
  const combinedInteractions = [];
  const range = notesDateRange_(); // ⚙️ Settings date range, same as the full build

  if (contactData) {
    for (let i = 1; i < contactData.length; i++) {
      const row = contactData[i];
      // ⚡ SHIFTED: OSIS boundary is now array index [4] (Col E)
      if (String(row[4]).trim().toUpperCase() === "END") break;
      if (!osisCellIncludes_(row[4], targetOsis)) continue;

      const entry = contactLogNoteLine_(row, disableHiddenNotes);
      if (entry && inNotesDateRange_(entry, range)) combinedInteractions.push(entry);
    }
  }

  if (eventData) {
    for (let i = 1; i < eventData.length; i++) {
      const row = eventData[i];
      if (!osisCellIncludes_(row[3], targetOsis)) continue;
      const entry = eventNoteLine_(row);
      if (inNotesDateRange_(entry, range)) combinedInteractions.push(entry);
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
  const guardianRows = buildGuardianObjects_(rawData, overrides.guardians, overrides.phones);
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

/**
 * Student List rows (OSIS, Student, Guardian, Site) incl. sibling groups, from the
 * Backend copy, WITH a header row first: the builders it feeds skip row 0 like the
 * build's own Student List (before 2026-11-12 there was no header, so the first
 * student lost their Student List details in Phone Contacts after a quick rebuild).
 */
function getExistingMasterTable_(ss) {
  const header = ["OSIS", "Student", "Guardian", "Site"];
  const data = studentLookupRows_(ss, 4);
  return [header].concat(data.filter(row => row[0] && !String(row[0]).includes("#")));
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

          google.script.run
            .withSuccessHandler(openPanel)
            .withFailureHandler(function (err) {
              alert('The event table could not be made: ' + (err && err.message ? err.message : err));
              btn.disabled = false; btn.innerText = 'Create table';
            })
            .buildEventTable(name, loc, formattedDate, type);
        }
        // The Events panel (EventsSidebar): fill-in tips + "Done: build and close Events"
        function openPanel() {
          google.script.run.withSuccessHandler(function () { google.script.host.close(); })
            .withFailureHandler(function () { google.script.host.close(); }).openEventsSidebar();
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
  
  if (sheet.isSheetHidden()) sheet.showSheet(); // hidden while no event is being filled in (EventsSidebar)
  sheet.activate(); 

  // Insert 7 columns on the far left (6 data + 1 spacer)
  sheet.insertColumns(1, 7);
  
  // Clear any residual formatting/validation from neighboring columns
  const maxRows = sheet.getMaxRows();
  sheet.getRange(1, 1, maxRows, 7).clearDataValidations().setBackground(null);

  // The table's look: title row, delete box, header row (EVENT_TABLE_HEADERS; Sheet Lock.gs
  // warns on row 2 of every table), lines, widths, font. One place: styleEventTableAt_
  styleEventTableAt_(sheet, 1, maxRows);

  // Row 1: Title (merged over columns 1-2), Category (3), Location (4), Date (5); 6 = Delete Box
  sheet.getRange(1, 1, 1, 2).setValue(eventName);
  sheet.getRange(1, 3).setValue(eventType);
  sheet.getRange(1, 4).setValue(eventLoc);
  sheet.getRange(1, 5).setValue(eventDate);

  // =======================================================
  // STRICT DROPDOWN MAPPING
  // =======================================================
  // Lists come from the Backend copy of the Student List, so sibling groups can be picked
  const rules = eventDropdownRules_(ss);
  sheet.getRange("A3:A").setDataValidation(rules.osis);
  sheet.getRange("B3:B").setDataValidation(rules.guardian);
  // Parent Attendee is manual entry, no strict validation needed.
  sheet.getRange("D3:D").setDataValidation(rules.student);

  // Ensure the Backend tab's Event Log section has its header row
  if (backendReadBlock_(ss, "EVENT_LOG").length === 0) {
    backendWriteBlock_(ss, "EVENT_LOG", [["Date", "Event Name", "Category", "OSIS", "Student Name", "Guardian(s)", "Parent Attendee", "# of Indv", "Notes"]]);
  }

  // One event at a time (owner 2026-11-24): the new table shows, every other one is hidden
  // (📆 Events → 👁️ View Events brings them back). Never stops the new table from being made.
  try {
    sheet.showColumns(1, EVENT_TABLE_WIDTH);
    showOnlyEventTables_(sheet, [1]);
    wrapAllEventNotes_(sheet); // older tables' Notes wrap too
    sheet.setActiveRange(sheet.getRange(3, 1));
  } catch (e) { console.warn("Other events not hidden: " + e.message); }

  // The tables moved right: re-lock every table's header row (Sheet Lock.gs)
  try { applySheetLocks_(ss, ["Events"]); } catch (e) { console.warn("Sheet lock skipped: " + e.message); }
}

// ==============================================================
// 2b. ONE EVENT AT A TIME + 👁️ VIEW EVENTS (owner 2026-11-24)
// ==============================================================
// 📋 New Event hides every other event table, so only the one being filled in shows.
// 📆 Events → 👁️ View Events lists every table on the Events tab (one row per table:
// name, type, date, place, how many rows filled in) to tick one, several or all; the
// ticked ones show, the others are hidden. Only HIDDEN: nothing is moved or deleted, and
// the build / notes / PCAR read hidden columns like any other. A table = 7 columns
// (EVENT_TABLE_WIDTH) from column 1 with its name in row 1 (eventTableStarts_, Sheet Lock.gs).

/** Every table on the Events tab, left to right (newest first: New Event adds at column A). */
function eventTablesList_(sheet) {
  const lastCol = sheet.getLastColumn(), lastRow = sheet.getLastRow();
  if (lastCol < 1 || lastRow < 1) return [];
  const vals = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const tz = Session.getScriptTimeZone();
  const text = v => (v instanceof Date ? Utilities.formatDate(v, tz, "MM/dd/yyyy") : toText_(v));
  const out = [];
  for (let c = 1; c <= lastCol; c += EVENT_TABLE_WIDTH) {
    const name = text(vals[0][c - 1]);
    if (!name) continue;
    let rows = 0;
    for (let r = 2; r < vals.length; r++) { // OSIS, Guardian(s), Parent Attendee or Student filled = a row
      if ([0, 1, 2, 3].some(i => c - 1 + i < lastCol && text(vals[r][c - 1 + i]))) rows++;
    }
    out.push({
      col: c, name: name,
      type: c + 1 < lastCol ? text(vals[0][c + 1]) : "",
      place: c + 2 < lastCol ? text(vals[0][c + 2]) : "",
      date: c + 3 < lastCol ? text(vals[0][c + 3]) : "",
      rows: rows,
      shown: !sheet.isColumnHiddenByUser(c)
    });
  }
  return out;
}

/**
 * Shows the tables starting at the columns in `cols`, hides every other table. One
 * showColumns / hideColumns call per run of tables in the same state. Columns that are
 * not part of a table (no name in row 1) are left as they are. Returns how many show.
 */
function showOnlyEventTables_(sheet, cols) {
  const want = {};
  (cols || []).forEach(c => { want[Number(c)] = true; });
  const maxCols = sheet.getMaxColumns();
  const starts = eventTableStarts_(sheet);
  let run = null, shown = 0;
  const flush = () => {
    if (!run) return;
    if (run.show) sheet.showColumns(run.from, run.to - run.from + 1);
    else sheet.hideColumns(run.from, run.to - run.from + 1);
    run = null;
  };
  starts.forEach(c => {
    const show = !!want[c];
    if (show) shown++;
    const to = Math.min(c + EVENT_TABLE_WIDTH - 1, maxCols);
    if (run && run.show === show && run.to === c - 1) run.to = to;
    else { flush(); run = { show: show, from: c, to: to }; }
  });
  flush();
  return shown;
}

/** 📆 Events → 👁️ View Events: tick the event tables to show. */
function openViewEventsDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Events");
  const list = sheet ? eventTablesList_(sheet) : [];
  const data = JSON.stringify(list).replace(/</g, "\\u003c");
  const html = include_('Styles') + `
    <style>
      .ve { padding: 8px 12px 12px; display: flex; flex-direction: column; gap: 10px; height: 100vh; box-sizing: border-box; }
      .ve p { margin: 0; color: var(--muted); font-size: 13px; }
      .bar { display: flex; gap: 8px; align-items: center; }
      .bar input[type=text] { flex: 1; }
      .list { flex: 1; overflow-y: auto; border: 1px solid var(--line); border-radius: var(--radius); }
      .row { display: flex; gap: 10px; align-items: flex-start; padding: 9px 10px; border-bottom: 1px solid var(--line); cursor: pointer; }
      .row:last-child { border-bottom: 0; }
      .row:hover { background: var(--panel); }
      .row input { width: 16px; height: 16px; margin: 2px 0 0; accent-color: var(--accent); flex: 0 0 auto; cursor: pointer; }
      .nm { font-weight: 600; color: var(--ink); }
      .sub { font-size: 12.5px; color: var(--muted); margin-top: 2px; }
      .empty { padding: 18px; color: var(--muted); text-align: center; }
      .foot { display: flex; justify-content: space-between; align-items: center; gap: 8px; }
      .foot .r { display: flex; gap: 8px; }
    </style>
    <div class="ve">
      <p>Tick the events to show on the Events tab. The others are hidden, not deleted.</p>
      <div class="bar">
        <input type="text" id="find" placeholder="Search by name, type or date">
        <button type="button" class="btn sec sm" id="allBtn">Select all</button>
      </div>
      <div class="list" id="list"></div>
      <div class="foot">
        <span class="sub" id="count"></span>
        <div class="r">
          <button type="button" class="btn sec" onclick="google.script.host.close()">Cancel</button>
          <button type="button" class="btn pri" id="showBtn">Show selected</button>
        </div>
      </div>
    </div>
    <script>
      const EVENTS = ${data};
      const picked = {};
      EVENTS.forEach(function (ev) { if (ev.shown) picked[ev.col] = true; });
      const listEl = document.getElementById('list'), findEl = document.getElementById('find');
      const allBtn = document.getElementById('allBtn'), showBtn = document.getElementById('showBtn');
      function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
      function visible() {
        const q = findEl.value.trim().toLowerCase();
        return EVENTS.filter(function (ev) { return !q || [ev.name, ev.type, ev.date, ev.place].join(' ').toLowerCase().indexOf(q) !== -1; });
      }
      function draw() {
        const vis = visible();
        if (!EVENTS.length) listEl.innerHTML = '<div class="empty">No events yet. Make one with 📆 Events → 📋 New Event.</div>';
        else if (!vis.length) listEl.innerHTML = '<div class="empty">No event matches your search.</div>';
        else listEl.innerHTML = vis.map(function (ev) {
          const bits = [ev.type, ev.date, ev.place].filter(Boolean);
          bits.push(ev.rows + (ev.rows === 1 ? ' row' : ' rows') + ' filled in');
          return '<label class="row"><input type="checkbox" data-col="' + ev.col + '"' + (picked[ev.col] ? ' checked' : '') + '>'
            + '<div><div class="nm">' + esc(ev.name) + (ev.shown ? ' <span class="pill">showing now</span>' : '') + '</div>'
            + '<div class="sub">' + esc(bits.join(' · ')) + '</div></div></label>';
        }).join('');
        const n = EVENTS.filter(function (ev) { return picked[ev.col]; }).length;
        document.getElementById('count').textContent = EVENTS.length ? n + ' of ' + EVENTS.length + ' selected' : '';
        showBtn.disabled = n === 0;
        showBtn.textContent = n === EVENTS.length && n > 1 ? 'Show all ' + n : 'Show ' + (n === 1 ? '1 event' : n + ' events');
        allBtn.textContent = vis.length && vis.every(function (ev) { return picked[ev.col]; }) ? 'Unselect all' : 'Select all';
        allBtn.disabled = !vis.length;
      }
      listEl.addEventListener('change', function (e) {
        const c = e.target.getAttribute('data-col');
        if (c) { picked[c] = e.target.checked; draw(); }
      });
      findEl.addEventListener('input', draw);
      allBtn.addEventListener('click', function () {
        const vis = visible(), on = !vis.every(function (ev) { return picked[ev.col]; });
        vis.forEach(function (ev) { picked[ev.col] = on; });
        draw();
      });
      showBtn.addEventListener('click', function () {
        const keys = EVENTS.filter(function (ev) { return picked[ev.col]; }).map(function (ev) { return { col: ev.col, name: ev.name }; });
        if (!keys.length) return;
        showBtn.disabled = true; showBtn.textContent = 'Showing…';
        google.script.run
          .withSuccessHandler(function () {
            google.script.run.withSuccessHandler(function () { google.script.host.close(); })
              .withFailureHandler(function () { google.script.host.close(); }).openEventsSidebar();
          })
          .withFailureHandler(function (err) {
            alert('The events could not be shown: ' + (err && err.message ? err.message : err));
            draw();
          })
          .showEventTables(keys);
      });
      draw();
    </script>
  `;
  const ui = HtmlService.createHtmlOutput(html).setWidth(480).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(ui, 'View Events');
}

/**
 * View Events' Show button: keys = [{ col, name }] of the ticked tables. A table is found
 * at its column with the same name; when tables moved since the window opened (an event
 * added or deleted), by its name. Shows them, hides the rest, goes to the first one.
 */
function showEventTables(keys) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Events");
  if (!sheet) throw new Error("There is no Events tab.");
  const list = eventTablesList_(sheet);
  const cols = [];
  (keys || []).forEach(k => {
    const name = toText_(k && k.name);
    const hit = list.find(t => t.col === Number(k && k.col) && t.name === name) || list.find(t => t.name === name && cols.indexOf(t.col) === -1);
    if (hit && cols.indexOf(hit.col) === -1) cols.push(hit.col);
  });
  if (!cols.length) throw new Error("Those events are no longer on the Events tab. Close this window and open 👁️ View Events again.");
  showOnlyEventTables_(sheet, cols);
  try { wrapAllEventNotes_(sheet); } catch (e) { console.warn("Event notes not wrapped: " + e.message); }
  if (sheet.isSheetHidden()) sheet.showSheet();
  sheet.activate();
  sheet.setActiveRange(sheet.getRange(1, Math.min.apply(null, cols)));
  ss.toast(cols.length === list.length ? "Showing every event." : "Showing " + cols.length + " of " + list.length + " events.", "👁️ View Events", 4);
  return { shown: cols.length, total: list.length };
}

// ==============================================================
// 2c. THE EVENTS PANEL (EventsSidebar.html; owner 2026-11-24)
// ==============================================================
// Opens after 📋 New Event and 👁️ View Events. Typing in an event only fills the row in
// (watchEventsSheet_); the panel says so and has "Done: build and close Events" (hides the
// Events tab, then the Build sidebar runs 🏗️ Build Sheets Only, which puts the attendance
// into the notes / Combined Contact Tracking / PCAR) and "Close Events, build later" (the
// 10 PM rebuild does it). The Events tab is hidden whenever no event is being filled in
// (not in ⚙️ Settings → Tabs & columns any more).

function openEventsSidebar() {
  SpreadsheetApp.getUi().showSidebar(htmlFromFile_('EventsSidebar').setTitle('Events').setWidth(320));
}

/** What the panel shows: the tables on screen, how many in all, whether tonight's rebuild runs. */
function getEventsPanel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Events");
  const list = sheet ? eventTablesList_(sheet) : [];
  const props = PropertiesService.getDocumentProperties();
  return {
    showing: list.filter(t => t.shown).map(t => ({ name: t.name, type: t.type, date: t.date, rows: t.rows })),
    total: list.length,
    tabShown: !!sheet && !sheet.isSheetHidden(),
    nightlyOff: props.getProperty('disableNightlySync') === 'true'
  };
}

/** Hides the Events tab (the user goes to the Contact Log first). Never throws. */
function hideEventsTab_(ss) {
  const sheet = ss.getSheetByName("Events");
  if (!sheet || sheet.isSheetHidden()) return false;
  try {
    if (ss.getActiveSheet().getSheetId() === sheet.getSheetId()) {
      const log = ss.getSheetByName("Contact Log");
      const to = (log && !log.isSheetHidden()) ? log
        : ss.getSheets().filter(sh => sh.getSheetId() !== sheet.getSheetId() && !sh.isSheetHidden())[0];
      if (!to) return false; // Sheets can't hide the last tab showing
      to.activate();
    }
    sheet.hideSheet();
    return true;
  } catch (e) { console.warn("Events tab not hidden: " + e.message); return false; }
}

/**
 * The panel's buttons. build = true: hides the Events tab and opens the Build sidebar
 * (🏗️ Build Sheets Only, in place of this panel). false: hides the tab, the panel closes.
 */
function eventsFinish(build) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  hideEventsTab_(ss);
  if (build) { runBuildOnly(); return { building: true }; }
  const nightlyOff = PropertiesService.getDocumentProperties().getProperty('disableNightlySync') === 'true';
  ss.toast(nightlyOff ? "Your events go into the notes at the next 🏗️ Build Sheets Only." : "Your events go into the notes at tonight's rebuild, or at the next 🏗️ Build Sheets Only.", "📆 Events closed", 6);
  return { building: false };
}

/**
 * One Events table's look (SHEET_LOOK_, Core.gs; owner 2026-11-12): the 6 columns from
 * column c and the spacer after them. Title row (event name over the first 2 columns,
 * type, place, date) in the typing tabs' brown with white bold text; the delete box on a
 * soft red; the header row (EVENT_TABLE_HEADERS) in light tan; thin lines with a brown
 * outline; Montserrat; the widths. Was a green / tan / yellow / blue / red title row, a grey
 * header row and black lines, typed out twice. Used by buildEventTable (a new event),
 * restyleEventTables_ (Migration) and 🎨 Apply Sheet Style. Writes the header row; other
 * values are not touched (the delete box becomes a fresh, unticked checkbox).
 */
function styleEventTableAt_(sheet, c, maxRows) {
  const L = SHEET_LOOK_, T = L.typing, S = SpreadsheetApp.BorderStyle;
  sheet.getRange(1, c, maxRows, 6).setFontFamily(L.font).setFontColor(L.ink)
    .setBorder(true, true, true, true, true, true, T.line, S.SOLID)
    .setBorder(true, true, true, true, null, null, T.header, S.SOLID);
  sheet.getRange(1, c, 1, 2).merge();
  sheet.getRange(1, c, 1, 5).setBackground(T.header).setFontColor(T.headerText)
    .setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange(1, c + 5).insertCheckboxes().setBackground(L.danger.fill).setFontColor(L.danger.text);
  sheet.getRange(2, c, 1, 6).setValues([EVENT_TABLE_HEADERS]).setFontWeight("bold")
    .setBackground(T.subHeader).setFontColor(T.subHeaderText);
  [100, 150, 150, 150, 80, 200, 30].forEach((w, i) => sheet.setColumnWidth(c + i, w)); // OSIS, Guardian(s), Parent Attendee, Student, # of Indv, Notes, spacer
  wrapEventNotesAt_(sheet, c, maxRows);
}

/** The Notes column of the table at column c wraps its text (owner 2026-11-24), header row down. */
function wrapEventNotesAt_(sheet, c, maxRows) {
  if (maxRows >= 2) sheet.getRange(2, c + 5, maxRows - 1, 1).setWrap(true);
}

/** Every table's Notes column wraps (tables made before 2026-11-24 too). One call per table. */
function wrapAllEventNotes_(sheet) {
  const maxRows = sheet.getMaxRows();
  eventTableStarts_(sheet).forEach(c => wrapEventNotesAt_(sheet, c, maxRows));
}

/**
 * Migration copies the Events tab as plain values (Menus.gs ma_migrateSimpleSheets): this
 * gives every table (7 columns wide from column A, title in row 1) its look back, the same
 * as buildEventTable makes it (styleEventTableAt_), and the OSIS / Guardian / Student
 * dropdowns. Values stay as they are (the delete box starts unticked). lookOnly (🎨 Apply
 * Sheet Style): the look only, the dropdowns are left as they are. Returns how many
 * tables it found.
 */
function restyleEventTables_(ss, sheet, lookOnly) {
  const lastCol = sheet.getLastColumn(), maxRows = sheet.getMaxRows();
  if (lastCol < 1 || sheet.getLastRow() < 1) return 0;
  const row1 = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const rules = lookOnly ? null : eventDropdownRules_(ss);
  let tables = 0;
  for (let c = 1; c <= lastCol; c += 7) {
    if (String(row1[c - 1]).trim() === "") continue;
    try {
      if (sheet.getMaxColumns() < c + 6) sheet.insertColumnsAfter(sheet.getMaxColumns(), c + 6 - sheet.getMaxColumns());
      styleEventTableAt_(sheet, c, maxRows);
      if (rules && maxRows >= 3) {
        sheet.getRange(3, c, maxRows - 2, 1).setDataValidation(rules.osis);
        sheet.getRange(3, c + 1, maxRows - 2, 1).setDataValidation(rules.guardian);
        sheet.getRange(3, c + 3, maxRows - 2, 1).setDataValidation(rules.student);
      }
      tables++;
    } catch (e) { console.warn("Event table at column " + c + " not restyled: " + e.message); }
  }
  return tables;
}

/**
 * The Events tables' OSIS / Guardian / Student dropdown rules: they list the
 * Backend copy of the Student List (Backend!BX3:BX, BY3:BY, BZ3:BZ), sibling
 * groups included. Before the first build with this version the list is empty.
 */
function eventDropdownRules_(ss) {
  const sheet = getBackendSheet_(ss);
  const b = BACKEND_BLOCKS.STUDENT_LIST;
  backendBlockReady_(sheet, b, true); // an older Backend tab gets the section's columns first
  const src = { sheet: sheet, firstRow: BACKEND_HEADER_ROW + 1, col: b.col };
  const rule = offset => SpreadsheetApp.newDataValidation().requireValueInRange(studentLookupColumn_(src, offset), true).build();
  return { osis: rule(0), student: rule(1), guardian: rule(2) };
}

/**
 * Event tables made before the sibling-group rows moved to the Backend list the
 * Student List tab (now one row per student). Points those columns at the Backend
 * copy. Reads row 3 only (buildEventTable gives each column one rule from row 3
 * down) and rewrites a column only while its rule still lists the Student List
 * tab, so after the first time this is one read. Run by every build.
 */
function pointEventDropdownsAtBackend_(ss) {
  const events = ss.getSheetByName("Events");
  const list = getStudentListSheet_(ss);
  if (!events || !list || events.getMaxRows() < 3) return;
  const listId = list.getSheetId();
  const keyFor = { 1: "osis", 2: "student", 3: "guardian" }; // Student List column -> rule
  let rules = null;
  events.getRange(3, 1, 1, events.getMaxColumns()).getDataValidations()[0].forEach((rule, i) => {
    if (!rule || rule.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) return;
    const listed = rule.getCriteriaValues()[0];
    if (!listed || listed.getSheet().getSheetId() !== listId || !keyFor[listed.getColumn()]) return;
    rules = rules || eventDropdownRules_(ss);
    const letter = columnLetter_(i + 1);
    events.getRange(letter + "3:" + letter).setDataValidation(rules[keyFor[listed.getColumn()]]);
  });
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

  // Tables are 7 columns wide (6 data + 1 spacer)
  const tableStartCol = Math.floor((col - 1) / 7) * 7 + 1;
  const relativeCol = col - tableStartCol; // 0=OSIS, 1=Guardian, 2=Attendee, 3=Student, 4=#Indv, 5=Notes/Box

  // Row 1: the red delete box. A new name / type / date needs nothing now: the next build reads it.
  if (row === 1) {
    if (relativeCol === 5 && range.getNumRows() === 1 && range.getNumColumns() === 1) deleteEventTable_(e, sheet, tableStartCol);
    return;
  }

  const maxRows = sheet.getLastRow();
  const maxCols = sheet.getLastColumn();
  if (maxRows < 3 || maxCols < 6) return;
  
  // 1. GHOST TYPIST LOGIC (Runs instantly before anything else)
  const tEventsTypist = Date.now();
  if (row >= 3) {
    const lookupSrc = studentLookupSource_(e.source); // Student List incl. sibling groups (Backend copy)
    if (lookupSrc) {
      
      // ⚡ CASE A: VERTICAL PASTE (Copying & Pasting a column of OSIS numbers)
      if (range.getNumRows() > 1 && range.getNumColumns() === 1 && relativeCol === 0) {
        const pastedValues = range.getDisplayValues();
        const masterData = studentLookupRows_(e.source, 3, lookupSrc);
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
               const masterData = studentLookupRows_(e.source, 3, lookupSrc);
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

                 // Fast memory sweep of the Student List
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
            // Typed OSIS -> Student List A, Guardian -> C, Student -> B
            const found = findInMasterTable_(lookupSrc, relativeCol === 0 ? 1 : relativeCol === 1 ? 3 : 2, editedValue);
            clDbg_("Events typist lookup " + (Date.now() - tEventsTypist) + " ms: " + (found ? "found OSIS " + found.osis : "'" + editedValue + "' not found"));
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
  clDbg_("⏱ Events typist " + (Date.now() - tEventsTypist) + " ms (painted)");

  // FAST EVENTS (owner 2026-11-24): an edit only fills the row in (the typist above). The
  // Backend Event Log, the notes, Combined Contact Tracking and the PCAR get the attendance
  // at the next build: the Events panel's "Done: build" (EventsSidebar), 🏗️ Build Sheets Only or
  // the 10 PM rebuild (buildBackendEventLogData_ reads every table). Before, every edit re-read
  // every table, rewrote the Backend Event Log and the students' notes.
  // Left: a family already in this table is taken out again (one read of its OSIS column).
  if (row >= 3 && relativeCol <= 3) removeDuplicateEventRows_(sheet, tableStartCol, maxRows);
}

/**
 * The same OSIS twice in one event table: the later row is cleared and the user is told
 * (as before). Several OSIS in one cell ("111, 222") count one by one.
 */
function removeDuplicateEventRows_(sheet, tableStartCol, maxRows) {
  if (maxRows < 4) return;
  const osisCol = sheet.getRange(3, tableStartCol, maxRows - 2, 1).getDisplayValues();
  const seen = {}, dupRows = [];
  osisCol.forEach((r, i) => {
    const list = String(r[0]).split(',').map(o => o.trim()).filter(Boolean);
    if (!list.length) return;
    if (list.some(o => seen[o])) { dupRows.push(i + 3); return; }
    list.forEach(o => { seen[o] = true; });
  });
  if (!dupRows.length) return;
  dupRows.forEach(r => sheet.getRange(r, tableStartCol, 1, 6).clearContent());
  const ui = SpreadsheetApp.getUi();
  ui.alert("⚠️ Duplicate Removed", "This family is already logged for this event. The duplicate has been removed.", ui.ButtonSet.OK);
}

/** The red box in row 1 was ticked: asks, then deletes that event's 7 columns. */
function deleteEventTable_(e, sheet, tableStartCol) {
  const range = e.range;
  if (String(e.value).toUpperCase() !== "TRUE") return;
  const eventName = String(sheet.getRange(1, tableStartCol).getValue()).trim();
  if (!eventName) return;
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert("⚠️ Delete Event", `Are you sure you want to completely delete the event "${eventName}"?\n\nIts attendance leaves the notes at the next build.`, ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) { range.uncheck(); return; }

  sheet.deleteColumns(tableStartCol, EVENT_TABLE_WIDTH);
  SpreadsheetApp.flush();
  // Every other event hidden (one at a time, 👁️ View Events): show the newest one, so
  // the tab isn't left with no table on it
  let showedNewest = "";
  try {
    const left = eventTablesList_(sheet);
    if (left.length && !left.some(t => t.shown)) { showOnlyEventTables_(sheet, [left[0].col]); showedNewest = left[0].name; }
  } catch (err) { console.warn("Events not shown after the delete: " + err.message); }
  try { applySheetLocks_(e.source, ["Events"]); } catch (err) { console.warn("Sheet lock skipped: " + err.message); }

  ui.alert("✅ Event Removed", "The event has been deleted. It leaves the notes and the PCAR at the next build (Done: build in the Events panel, 🏗️ Build Sheets Only, or tonight's rebuild)."
    + (showedNewest ? "\n\nShowing \"" + showedNewest + "\" now. To see other events: 📆 Events → 👁️ View Events." : ""), ui.ButtonSet.OK);
}

/** First Student List row whose OSIS is exactly `osis`, from studentLookupRows_ (data rows only). */
function findMasterRowByOsis_(masterData, osis) {
  for (let r = 0; r < masterData.length; r++) {
    if (String(masterData[r][0]).trim() === osis) return masterData[r];
  }
  return null;
}
