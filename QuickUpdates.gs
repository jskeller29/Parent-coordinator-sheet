// ======================================================================
// FILE: QuickUpdates.gs
// ======================================================================

// =========================================================
// HELPER FUNCTIONS (Must be at the top so they load first!)
// =========================================================

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

// Kept safely so old files don't crash before we update them
function findEndRowInColumnG_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return null;
  const values = sheet.getRange(1, 7, lastRow, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim().toUpperCase() === "END") return i + 1;
  }
  return null;
}

function patchMissingFormulasInRow_(sheet, row) {
  const range = sheet.getRange(row, 1, 1, 1);
  const values = range.getDisplayValues()[0];
  const formulas = range.getFormulas()[0];

  const formulaWeek = `=IF(B${row}="", "", B${row} - WEEKDAY(B${row}, 1) + 1)`;
  if (String(values[0]).trim() === "" || String(values[0]).indexOf("#") === 0 || formulas[0] === "") {
    sheet.getRange(row, 1).setFormula(formulaWeek);
  }
}

function FIX_ALL_MISSING_FORMULAS_BATCH() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Contact Log");
  if (!sheet) return;

  const endRow = typeof findEndRowInColumnE_ === "function" ? findEndRowInColumnE_(sheet) : null;
  if (!endRow || endRow <= 2) {
    ss.toast("Could not find END in Column E or table is empty.", "Error", 5);
    return;
  }

  const numRows = endRow - 2;
  const startRow = 2;
  const colAFormulas = [];

  for (let i = 0; i < numRows; i++) {
    const rowNum = startRow + i;
    colAFormulas.push([`=IF(B${rowNum}="", "", B${rowNum} - WEEKDAY(B${rowNum}, 1) + 1)`]);
  }

  sheet.getRange(startRow, 1, numRows, 1).setFormulas(colAFormulas);
  formatContactLogRows_(sheet, startRow, numRows);

  ss.toast("Week formulas restored and borders fixed!", "Success", 3);
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

function updateContactAndPhoneOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName("RAW Data") || ss.getSheetByName("Raw Data");
  if (!rawSheet) return;

  const rawData = rawSheet.getDataRange().getValues();
  if (rawData.length < 4) return;

  if (typeof buildStudentMap_ !== "function") return;
  if (typeof buildGuardianObjects_ !== "function") return;
  if (typeof buildCombinedContactTrackingData_ !== "function") return;
  if (typeof buildPhoneContactsData_ !== "function") return;
  if (typeof buildParentSquareIndex_ !== "function") return;
  if (typeof buildNYCSAIndex_ !== "function") return;
  if (typeof writeSheet_ !== "function") return;

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
  
  const eventLogSheet = ss.getSheetByName("Backend_Event_Log");
  const eventLogData = eventLogSheet ? eventLogSheet.getDataRange().getValues() : [];

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
    
    if (typeof forceVibrantColors_ === "function") forceVibrantColors_(sheet);
    if (typeof applyBandingToRawData_ === "function") applyBandingToRawData_(sheet);
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
    sheet.getRange("BH3:BI3").merge().setValue("MAKE OWN DATA").setHorizontalAlignment("center");
  } catch(e) {
    sheet.getRange("BH3").setValue("MAKE OWN DATA").setHorizontalAlignment("center");
  }
  
  sheet.getRange("AF:AF").clearContent().setBackground("#FFFFFF");
  sheet.getRange("AO:AO").clearContent().setBackground("#FFFFFF");
  sheet.getRange("BG:BG").clearContent().setBackground("#FFFFFF");
  
  const targetRow = isFullExport ? 7 : Math.max(7, range.getRow());
  if (finalData.length > 0) {
    sheet.getRange(targetRow, pasteStartCol, finalData.length, expectedHeaders.length).setValues(finalData);
  }
  
  if (typeof forceVibrantColors_ === "function") forceVibrantColors_(sheet);
  if (typeof applyBandingToRawData_ === "function") applyBandingToRawData_(sheet);
  
  SpreadsheetApp.getActiveSpreadsheet().toast(`Global headers locked. Data mapped to Row ${targetRow}.`, `${targetZone} Cleaned ✨`, 6);
}

function applyTypeDropdownColors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const contactLog = ss.getSheetByName('Contact Log');
  const typeList = ss.getSheetByName('Type List');
  
  const typeValues = typeList.getRange('A1:A').getDisplayValues().flat().filter(String);
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
    filter = sheet.getRange(1, 1, endRow - 1, 16).createFilter();
    const blackColor = SpreadsheetApp.newColor().setRgbColor("#000000").build();
    const criteria = SpreadsheetApp.newFilterCriteria()
      .setVisibleForegroundColor(blackColor) 
      .build();
    filter.setColumnFilterCriteria(16, criteria);
  }
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
// ======================================================================
function updateSingleOsisDelta_(targetOsis) {
  if (!targetOsis) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const props = PropertiesService.getDocumentProperties();
  const disableHiddenNotes = props.getProperty('disableHiddenNotes') === 'true';

  const contactSheet = ss.getSheetByName("Contact Log");
  const combinedInteractions = [];
  
  if (contactSheet) {
    const contactData = contactSheet.getDataRange().getDisplayValues();
    for (let i = 1; i < contactData.length; i++) {
      const row = contactData[i];
      
      // ⚡ SHIFTED: OSIS boundary is now array index [4] (Col E)
      if (String(row[4]).trim().toUpperCase() === "END") break; 
      
      if (!osisCellIncludes_(row[4], targetOsis)) continue;

      // ⚡ SHIFTED: Hide Checkbox is now array index [15] (Col P)
      const isHidden = row[15]; 
      const hiddenBool = (isHidden === true || String(isHidden).toUpperCase() === "TRUE");
      
      if (!disableHiddenNotes && hiddenBool) continue;

      let dateVal = row[1];
      if (!dateVal) {
        dateVal = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");
      }

      // ⚡ SHIFTED ALL INDEXES (-3)
      const mVal = String(row[9]).trim();  // Method 1
      const nVal = String(row[10]).trim(); // Method 2
      const eVal = String(row[3]).trim();  // Person Spoke With
      const pVal = String(row[12]).trim(); // Notes
      const rVal = String(row[14]).trim(); // Followup Notes

      if (pVal === "" && rVal === "") continue;

      const methodLabel = nVal !== "" ? nVal : mVal;
      const dateText = typeof formatDateFull_ === "function" ? formatDateFull_(dateVal) : dateVal;
      
      let line = `• ${dateText} - {${methodLabel} - ${eVal}} ${pVal}`;
      if (rVal !== "") line += ` [${rVal}]`;

      let dObj = new Date(dateVal); 
      if (isNaN(dObj)) dObj = new Date(0);
      
      combinedInteractions.push({ dateObj: dObj, text: line.trim() });
    }
  }

  const eventSheet = ss.getSheetByName("Backend_Event_Log");
  if (eventSheet) {
    const eventData = eventSheet.getDataRange().getDisplayValues();
    for (let i = 1; i < eventData.length; i++) {
      const row = eventData[i];
      if (!osisCellIncludes_(row[3], targetOsis)) continue;

      const dateVal = row[0];
      const eventName = String(row[1]).trim();
      const studentName = String(row[4]).trim();
      const guardian = String(row[5]).trim(); 
      const attendee = String(row[6]).trim(); 
      const notes = String(row[8]).trim();    

      let person = attendee ? attendee : guardian;
      if (!person) person = "Parent";

      const dateText = typeof formatDateFull_ === "function" ? formatDateFull_(dateVal) : dateVal;
      let line = `• ${dateText} - {Event - [${person}]} |${eventName}|`;
      if (notes) line += ` ${notes}`;

      let dateObj = new Date(dateVal); 
      if (isNaN(dateObj)) dateObj = new Date(0);

      combinedInteractions.push({ dateObj: dateObj, text: line.trim() });
    }
  }

  // ✨ NEW: Fetch Manual Notes using .includes() so siblings get it too!
  let mergedManualNotes = "";
  const notesSheet = ss.getSheetByName("Notes");
  if (notesSheet) {
    const nData = notesSheet.getDataRange().getValues();
    for (let i = 1; i < nData.length; i++) {
      if (!nData[i]) continue;
      if (!osisCellIncludes_(nData[i][0], targetOsis)) continue;
      
      const nText = String(nData[i][3] || "").trim();
      if (nText) {
        mergedManualNotes = mergedManualNotes ? mergedManualNotes + "\n\n" + nText : nText;
      }
    }
  }

  // Combine and sort the interaction bullets
  let baseNotes = "";
  if (combinedInteractions.length > 0) {
    combinedInteractions.sort((a, b) => a.dateObj.getTime() - b.dateObj.getTime());
    baseNotes = combinedInteractions.map(interaction => interaction.text).join("\n");
  }

  // Merge the manual note block over the interactions
  const manualNotes = mergedManualNotes !== "" ? `𝑵𝒐𝒕𝒆𝒔: ${mergedManualNotes}` : "";
  let finalNotesList = "";
  if (manualNotes !== "" && baseNotes !== "") {
    finalNotesList = manualNotes + "\n\n" + baseNotes;
  } else if (manualNotes !== "") {
    finalNotesList = manualNotes;
  } else if (baseNotes !== "") {
    finalNotesList = baseNotes;
  }

  const combinedSheet = ss.getSheetByName("Combined Contact Tracking");
  if (!combinedSheet) return;
  
  const combinedData = combinedSheet.getRange("A:A").getDisplayValues();
  let foundInBackend = false;
  
  for (let i = 1; i < combinedData.length; i++) {
    // We strictly match the Combined Contact Tracking single-row OSIS
    const sheetOsis = String(combinedData[i][0]).replace(/[^a-zA-Z0-9]/g, '');
    const cleanTarget = String(targetOsis).replace(/[^a-zA-Z0-9]/g, '');
    
    if (sheetOsis === cleanTarget && cleanTarget !== "") {
      combinedSheet.getRange(i + 1, 4).setValue(finalNotesList); // Column D
      foundInBackend = true;
      break;
    }
  }
  
  if (!foundInBackend && finalNotesList) {
     ss.toast("New student profile syncing to backend...", "Auto-Build", 3);
     if (typeof updateContactAndPhoneOnly === "function") updateContactAndPhoneOnly();
  }
}
