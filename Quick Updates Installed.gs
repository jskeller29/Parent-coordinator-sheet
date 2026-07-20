// ======================================================================
// FILE: installedOnEdit.gs
// ======================================================================

// =========================================================
// MAIN TRIGGER FUNCTION
// =========================================================

function installedOnEdit(e) {  
  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName().trim();

  // --- RAW DATA PASTE INTERCEPTOR ---
  if (sheetName === "RAW Data" || sheetName === "Raw Data") {
    if (range.getNumRows() > 1 || range.getNumColumns() > 1) {
      if (typeof cleanRawDataPasteZone_ === "function") cleanRawDataPasteZone_(sheet, range);
    }
    if (typeof fixRawDataNumberFormatting_ === "function") fixRawDataNumberFormatting_(sheet);
    if (typeof updateRawDataTimestamp_ === "function") updateRawDataTimestamp_(range);
    return; 
  }

  // --- TRIGGER THE EVENTS GHOST TYPIST & BACKEND ---
  if (sheetName === "Events") {
    if (typeof watchEventsSheet_ === "function") watchEventsSheet_(e);
    return; 
  }

  // --- NEW: TRIGGER THE NOTES GHOST TYPIST ---
  if (sheetName === "Notes") {
    if (typeof watchNotesSheet_ === "function") watchNotesSheet_(e);
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

  // 1. Shifted End Row Finder (Now Col E / 5)
  let endRow = typeof findEndRowInColumnE_ === "function" ? findEndRowInColumnE_(sheet) : null;

  // --- NEW FILTER LOGIC ---
  // The Master Toggle shifted from Col 18 to Col 15 (O)
  if (endRow && row === endRow && col === 15) {
    if (typeof toggleHideFilter_ === "function") toggleHideFilter_(sheet, endRow);
    return; 
  }
  // ------------------------

  if (endRow && row >= endRow) return;

  // --- SAFE ROW CEILING ---
  let safeNumRows = numRows;
  if (endRow && (row + safeNumRows) > endRow) {
    safeNumRows = endRow - row; 
  }

  if (safeNumRows > 0) {
    // --- INSTANT ROW REPAIR (Shifted overall width to 16 / P) ---
    sheet.getRange(row, 1, safeNumRows, 16).setBackground(null);

    // Strip any stray horizontal borders from the edited rows — only the dark
    // brown end bar should carry one (re-drawn by the sweep at the bottom).
    sheet.getRange(row, 1, safeNumRows, 16).setBorder(false, null, false, null, null, false);

    // Parent and Person Spoke With (Shifted to Col 3 and 4)
    sheet.getRange(row, 3, safeNumRows, 2).setBorder(false, false, false, false, false, false);
    sheet.getRange(row, 4, safeNumRows, 1).setFontFamily("Roboto");
    
    // Draw the divider line before Notes (Shifted to Column 13 / M)
    sheet.getRange(row, 13, safeNumRows, 1).setBorder(null, null, null, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID);
    
    // Draw the divider line before Followup Notes (Shifted to Column 15 / O)
    sheet.getRange(row, 15, safeNumRows, 1).setBorder(null, null, null, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID);
    
    // Checkboxes firmly planted in P (16)
    sheet.getRange(row, 16, safeNumRows, 1).insertCheckboxes().setFontColor("#000000");
  }

  if (typeof autoAddBlankRow_ === "function") autoAddBlankRow_(sheet, row, endRow);
  if (typeof patchMissingFormulasInRow_ === "function") patchMissingFormulasInRow_(sheet, row);

  // =======================================================
  // CONTACT LOG GHOST TYPIST (OSIS, Guardian, Student)
  // =======================================================
  if (range.getNumRows() === 1 && range.getNumColumns() === 1) {
    // Watch new columns: C (3, Guardian), E (5, OSIS), F (6, Student)
    if ((col === 3 || col === 5 || col === 6) && editedValue && editedValue.toUpperCase() !== "FALSE" && editedValue.toUpperCase() !== "TRUE") {
      const masterSheet = e.source.getSheetByName("Master Table");
      if (masterSheet) {
        let searchRange;
        if (col === 5) searchRange = masterSheet.getRange("A5:A"); // Typed OSIS
        else if (col === 6) searchRange = masterSheet.getRange("B5:B"); // Typed Student
        else if (col === 3) searchRange = masterSheet.getRange("C5:C"); // Typed Guardian

        const found = searchRange.createTextFinder(editedValue).matchEntireCell(true).matchCase(false).findNext();
        if (found) {
          const foundRow = found.getRow();
          
          // Pull core identifying data from the Master Table
          const correctOsis = masterSheet.getRange(foundRow, 1).getValue();
          const correctStudent = masterSheet.getRange(foundRow, 2).getValue();
          const correctGuardian = masterSheet.getRange(foundRow, 3).getValue();
          
          // --- PULL SITE & CLASS FROM DIRECTORY (HANDLES SIBLINGS) ---
          const sites = new Set();
          const classes = new Set();
          
          const dirSheet = e.source.getSheetByName("Directory");
          if (dirSheet && correctOsis) {
            const osisArray = String(correctOsis).split(',').map(o => o.trim()).filter(String);
            const dirSearchRange = dirSheet.getRange("C:C");
            
            osisArray.forEach(individualOsis => {
              const dirFound = dirSearchRange.createTextFinder(individualOsis).matchEntireCell(true).findNext();
              if (dirFound) {
                const dirRow = dirFound.getRow();
                const dirCol = dirFound.getColumn(); 
                
                const siteClassVals = dirSheet.getRange(dirRow, dirCol + 1, 1, 2).getValues()[0];
                const siteVal = String(siteClassVals[0]).trim();
                const classVal = String(siteClassVals[1]).trim();
                
                if (siteVal) sites.add(siteVal);
                if (classVal) classes.add(classVal);
              }
            });
          }
          
          const correctSite = Array.from(sites).join(", ");
          const correctClass = Array.from(classes).join(", ");
          
          const updateRange = sheet.getRange(row, 3, 1, 7); 
          const rowData = updateRange.getValues();
          
          if (col !== 3) rowData[0][0] = correctGuardian; // Col 3 
          if (col !== 5) rowData[0][2] = correctOsis;     // Col 5 
          if (col !== 6) rowData[0][3] = correctStudent;  // Col 6 
          rowData[0][5] = correctSite;                    // Col 8 
          rowData[0][6] = correctClass;                   // Col 9 
          
          updateRange.setValues(rowData); 
          
        } else {
          e.source.toast("Could not find '" + editedValue + "' in the Master Table.", "Ghost Typist Failed", 4);
        }
      }
    }

    // =======================================================
    // NEW: AUTO-FILL DATE / # INTERACTIONS / FOLLOW-UP STATUS
    // Fires on the same single-cell edits. Fill-if-blank only, and skipped
    // entirely when "Disable Contact Log Auto-Fill" is on in Settings. The
    // identity lookup above always runs regardless of that setting.
    // =======================================================
    if (typeof autofillContactLogDefaults_ === "function") {
      autofillContactLogDefaults_(sheet, row, col, endRow, editedValue);
    }
  }

  // --- LAG FIX: Only recalculate colors if you actually changed the Type (Shifted to Col 12 / L) ---
  if (col === 12) {
    if (typeof applyTypeDropdownColors === "function") applyTypeDropdownColors(); 
  }

  // ==================================================
  // --- FORCE END BAR FORMATTING & SWEEP ---
  // ==================================================
  let finalEndRow = typeof findEndRowInColumnE_ === "function" ? findEndRowInColumnE_(sheet) : null;
  if (finalEndRow) {
    sheet.getRange(finalEndRow, 1, 1, 16).setBorder(true, null, true, null, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID);
    
    let maxRows = sheet.getMaxRows();
    if (maxRows >= finalEndRow) {
      let rowsToClear = maxRows - finalEndRow + 1;
      sheet.getRange(finalEndRow, 16, rowsToClear, 1).clearDataValidations().clearContent();
    }
  }

  SpreadsheetApp.flush(); 

  // --- DYNAMIC DELTA SYNC TRIGGER ---
  let targetOsis = String(sheet.getRange(row, 5).getDisplayValue()).trim();

  if (typeof isValidOsis_ === "function" && !isValidOsis_(targetOsis)) {
    Utilities.sleep(1000);
    SpreadsheetApp.flush();
    targetOsis = String(sheet.getRange(row, 5).getDisplayValue()).trim();
  }

  if (typeof isValidOsis_ === "function" && isValidOsis_(targetOsis)) {
    const props = PropertiesService.getDocumentProperties();
    const autoNotes = props.getProperty('autoNotes') !== 'false'; 
    
    if (autoNotes) {
      if (targetOsis.includes(",")) {
        if (typeof updateSingleOsisDelta_ === "function") updateSingleOsisDelta_(targetOsis);
        const siblingOsisArray = targetOsis.split(",").map(o => o.trim());
        siblingOsisArray.forEach(individualOsis => {
          if (typeof isValidOsis_ === "function" && isValidOsis_(individualOsis)) {
            if (typeof updateSingleOsisDelta_ === "function") updateSingleOsisDelta_(individualOsis);
          }
        });
      } else {
        if (typeof updateSingleOsisDelta_ === "function") updateSingleOsisDelta_(targetOsis);
      }
    } else {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "Note saved locally. Background sync is paused via Settings.", 
        "⚡ Fast Entry Mode", 
        3
      );
    }

  } else {
    if (col === 5 && !editedValue) {
      if (typeof updateContactAndPhoneOnly === "function") updateContactAndPhoneOnly();
    }
  }
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
// NEW: CONTACT LOG AUTO-FILL (Date / # of interactions / Follow-up status)
// =========================================================
// When a coordinator enters real interaction content into a fresh row, stamp the
// three routine defaults so they don't have to type them every time:
//   B (2)  = today's date (a real midnight Date, so the week formula in Col A and
//            the "today = green" / week-band conditional formatting all work)
//   G (7)  = 1  (# of interactions; this also clears the red "missing" flag)
//   N (14) = "No"  (the follow-up status dropdown's default)
//
// Rules:
//   • FILL-IF-BLANK: never overwrite a cell the user already filled. Back-dating a
//     date, setting G to 3, or choosing "Yes - Important" in N is all preserved.
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
  // Method 1(10), Method 2(11), Type(12), Notes(13), Follow-up Notes(15).
  const CONTENT_COLS = [3, 4, 5, 6, 10, 11, 12, 13, 15];
  if (CONTENT_COLS.indexOf(col) === -1) return;

  // Ignore blank edits and checkbox toggles so clearing a cell never stamps a row.
  const val = String(editedValue == null ? "" : editedValue).trim().toUpperCase();
  if (val === "" || val === "TRUE" || val === "FALSE") return;

  // B (2): Date — only if empty.
  const dateCell = sheet.getRange(row, 2);
  if (String(dateCell.getDisplayValue()).trim() === "") {
    const today = new Date();
    today.setHours(0, 0, 0, 0); // midnight => $B2 = TODAY() and Col A week math both work
    dateCell.setValue(today);
  }

  // G (7): # of interactions — only if empty.
  const countCell = sheet.getRange(row, 7);
  if (String(countCell.getDisplayValue()).trim() === "") countCell.setValue(1);

  // N (14): Follow-up status dropdown — only if empty (never flip a chosen value).
  const followupCell = sheet.getRange(row, 14);
  if (String(followupCell.getDisplayValue()).trim() === "") followupCell.setValue("No");
}

// =========================================================
// NEW: NOTES SHEET GHOST TYPIST
// =========================================================
function watchNotesSheet_(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const row = range.getRow();
  const col = range.getColumn();
  
  // Ignore headers and multi-cell edits
  if (row < 2) return;
  if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) return;

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
      if (targetOsis && typeof updateSingleOsisDelta_ === "function") {
        const osisParts = targetOsis.split(",");
        osisParts.forEach(o => {
          const cleanO = o.trim();
          if (cleanO) updateSingleOsisDelta_(cleanO);
        });
      }
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
    if (targetOsis && typeof updateSingleOsisDelta_ === "function") {
      const osisParts = targetOsis.split(",");
      osisParts.forEach(o => {
        const cleanO = o.trim();
        if (cleanO) updateSingleOsisDelta_(cleanO);
      });
    }
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
      let searchRange;
      if (col === 1) searchRange = masterSheet.getRange("A5:A");      // Typed OSIS
      else if (col === 2) searchRange = masterSheet.getRange("C5:C"); // Typed Guardian
      else if (col === 3) searchRange = masterSheet.getRange("B5:B"); // Typed Student

      if (!searchRange) return;

      const found = searchRange.createTextFinder(editedValue).matchEntireCell(true).matchCase(false).findNext();

      if (found) {
        const foundRow = found.getRow();
        
        // Pull exact data from Master Table
        const correctOsis = masterSheet.getRange(foundRow, 1).getValue();
        const correctStudent = masterSheet.getRange(foundRow, 2).getValue();
        const correctGuardian = masterSheet.getRange(foundRow, 3).getValue();

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

      } else {
        e.source.toast("Could not find '" + editedValue + "' in the Master Table.", "Ghost Typist Failed", 4);
      }
    }
  }
}
