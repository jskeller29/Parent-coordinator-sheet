// ======================================================================
// FILE: MigrationAssistant.gs
// PURPOSE: Safely migrates data from an older version of the CRM to the
// newest version, using Smart Mapping to handle column schema changes.
// ======================================================================

function openMigrationWizard() {
  const html = HtmlService.createHtmlOutputFromFile('MigrationUI')
    .setTitle('🔄 System Upgrade Wizard')
    .setWidth(550)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'System Upgrade Wizard');
}

/**
 * STEP 1: Verification
 * Ensures the URL is valid and the user has access.
 */
function ma_verify(url) {
  try {
    const oldSS = SpreadsheetApp.openByUrl(url);
    return { success: true, name: oldSS.getName() };
  } catch (e) {
    return { success: false, error: "Could not open sheet. Ensure the link is correct and you have access." };
  }
}

/**
 * STEP 2: Recover Custom Settings (Smart Inference)
 * Because background properties are locked, we infer settings based on old sheet visibility!
 */
function ma_migrateSettings(url) {
  const oldSS = SpreadsheetApp.openByUrl(url);
  const props = PropertiesService.getDocumentProperties();
  
  const pPhone = oldSS.getSheetByName("Phone Contacts");
  const pDiv = oldSS.getSheetByName("Parents Divided");
  
  if (pPhone) {
      const isPhoneEnabled = !pPhone.isSheetHidden();
      props.setProperty('phoneContacts', isPhoneEnabled.toString());
      props.setProperty('syncToPhones', isPhoneEnabled.toString()); // Tie sync setting to sheet visibility
  }
  if (pDiv) props.setProperty('parentsDivided', (!pDiv.isSheetHidden()).toString());
  
  // Trigger a visual update on the new sheet so tabs hide/show correctly
  if (typeof updateVisualSettings_ === "function" && typeof getSettings === "function") {
     updateVisualSettings_(getSettings());
  }
}

/**
 * STEP 3: Migrate Raw Data
 */
function ma_migrateRawData(url) {
  const oldSS = SpreadsheetApp.openByUrl(url);
  const newSS = SpreadsheetApp.getActiveSpreadsheet();
  
  const oldRaw = oldSS.getSheetByName("Raw Data");
  const newRaw = newSS.getSheetByName("Raw Data");
  if (!oldRaw || !newRaw) return;
  
  const data = oldRaw.getDataRange().getValues();
  newRaw.clearContents();
  if (data.length > 0) {
    newRaw.getRange(1, 1, data.length, data[0].length).setValues(data);
  }
}

/**
 * STEP 4: Migrate Standard "Simple" Sheets
 * (New/Edit Student, Send Out, Events)
 */
function ma_migrateSimpleSheets(url) {
  const oldSS = SpreadsheetApp.openByUrl(url);
  const newSS = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheetsToCopy = ["New/Edit Student", "Send Out", "Events"];
  
  sheetsToCopy.forEach(sheetName => {
    const oldSheet = oldSS.getSheetByName(sheetName);
    const newSheet = newSS.getSheetByName(sheetName);
    
    if (oldSheet && newSheet) {
      const data = oldSheet.getDataRange().getValues();
      if (data.length > 1) { 
        const maxRows = newSheet.getMaxRows();
        if (maxRows > 1) {
          newSheet.getRange(2, 1, maxRows - 1, newSheet.getMaxColumns()).clearContent();
        }
        
        if (newSheet.getMaxRows() < data.length) newSheet.insertRowsAfter(newSheet.getMaxRows(), data.length - newSheet.getMaxRows());
        if (newSheet.getMaxColumns() < data[0].length) newSheet.insertColumnsAfter(newSheet.getMaxColumns(), data[0].length - newSheet.getMaxColumns());
        
        newSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
      }
    }
  });
}

/**
 * STEP 5: The Smart Mapper for Contact Log
 * Handles dynamic column shifts and strips blank rows
 */
function ma_migrateContactLog(url) {
  const oldSS = SpreadsheetApp.openByUrl(url);
  const newSS = SpreadsheetApp.getActiveSpreadsheet();
  
  const oldLog = oldSS.getSheetByName("Contact Log");
  const newLog = newSS.getSheetByName("Contact Log");
  if (!oldLog || !newLog) throw new Error("Contact Log missing");
  
  // 1. Get Old Data & Find END
  const oldData = oldLog.getDataRange().getValues();
  let oldEndIdx = oldData.length;
  
  for (let i = 0; i < oldData.length; i++) {
    const colA = String(oldData[i][0]).trim().toUpperCase();
    const colE = String(oldData[i][4]).trim().toUpperCase(); 
    if (colA === "END" || colE === "END" || String(oldData[i][6]).trim().toUpperCase() === "END") {
      oldEndIdx = i;
      break;
    }
  }
  
  const oldHeaders = oldData[0].map(h => String(h).trim().toLowerCase());
  
  // SCRUBBER: Strip headers, END row, AND perfectly empty/ghost rows
  const isGhostValue = (val) => {
    if (val instanceof Date && val.getFullYear() < 1970) return true; // Destroy 1899 phantom dates
    const s = String(val).trim().toUpperCase();
    return s === "" || s.startsWith("#") || s === "FALSE" || s === "-";
  };

  const oldRecords = oldData.slice(1, oldEndIdx).filter(row => {
    // Check the first 10 core columns; if ANY of them contain real data, keep the row
    return row.slice(0, 10).some(val => !isGhostValue(val));
  });
  
  if (oldRecords.length === 0) return; // Nothing to migrate
  
  // 2. Get New Headers & Find END
  const newData = newLog.getDataRange().getValues();
  const newHeaders = newData[0].map(h => String(h).trim().toLowerCase());
  let newEndRow = 2; // Default
  
  for (let i = 0; i < newData.length; i++) {
    if (String(newData[i][0]).trim().toUpperCase() === "END" || 
        String(newData[i][4]).trim().toUpperCase() === "END" ||
        String(newData[i][6]).trim().toUpperCase() === "END") {
      newEndRow = i + 1; 
      break;
    }
  }
  
  // 3. Build Smart Mapping Dictionary
  const headerMap = {}; 
  oldHeaders.forEach((oldH, oldIdx) => {
    if (!oldH) return;
    const newIdx = newHeaders.indexOf(oldH);
    if (newIdx !== -1) {
      headerMap[oldIdx] = newIdx;
    }
  });
  
  // 4. Construct Perfectly Mapped Array
  const mappedRecords = [];
  oldRecords.forEach(oldRow => {
    const newRow = new Array(newHeaders.length).fill(""); 
    Object.keys(headerMap).forEach(oldIdx => {
      const newIdx = headerMap[oldIdx];
      let val = oldRow[oldIdx];
      
      // Fix stringified dates on the fly if this maps to a Date column
      if (newHeaders[newIdx] === "date" && typeof val === 'string' && val.includes(',')) {
        val = val.split(',')[1].trim();
        let parsedDate = new Date(val);
        if (!isNaN(parsedDate)) val = parsedDate;
      }
      
      // Final sweep to prevent stray 1899 ghost dates from entering the array
      if (val instanceof Date && val.getFullYear() < 1970) val = "";
      
      newRow[newIdx] = val;
    });
    mappedRecords.push(newRow);
  });
  
  // 5. Inject into New Sheet (UPGRADED: Starts exactly at Row 2, Leaves 1 Blank Row)
  const currentSpace = newEndRow - 2;
  const neededSpace = mappedRecords.length + 1; // 💡 +1 explicitly ensures a blank row is ALWAYS left above the END row

  if (currentSpace > 0) {
    // Erase the blank spacer rows
    newLog.getRange(2, 1, currentSpace, newLog.getMaxColumns()).clearContent();
  }

  // Adjust rows dynamically so the END row perfectly hugs the bottom (leaving 1 spacer)
  if (neededSpace > currentSpace) {
    const rowsToAdd = neededSpace - currentSpace;
    newLog.insertRowsBefore(newEndRow, rowsToAdd);
    newEndRow += rowsToAdd;
  } else if (neededSpace < currentSpace) {
    const rowsToRemove = currentSpace - neededSpace;
    newLog.deleteRows(2 + neededSpace, rowsToRemove);
    newEndRow -= rowsToRemove;
  }

  // Paste the migrated data starting exactly at row 2
  if (mappedRecords.length > 0) {
    newLog.getRange(2, 1, mappedRecords.length, newHeaders.length).setValues(mappedRecords);
  }
  
  // 💡 CRITICAL FORMATTING OVERRIDE:
  // Instead of letting Google Sheets guess the formatting (which might pull the brown 'END' bar upwards), 
  // we take the clean formatting from Row 2 and physically paint it down over everything, 
  // ensuring the new blank row perfectly matches the data rows above it.
  if (neededSpace > 1) {
    const formatSource = newLog.getRange(2, 1, 1, newLog.getMaxColumns());
    const formatTarget = newLog.getRange(3, 1, neededSpace - 1, newLog.getMaxColumns());
    formatSource.copyTo(formatTarget, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    formatSource.copyTo(formatTarget, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  }
  
  // ==========================================
  // --- SPACER ROW SANITIZATION ---
  // ==========================================
  const spacerRow = newEndRow - 1;
  
  // 1. Wipe Columns A through O completely clean of any cloned phantom text
  newLog.getRange(spacerRow, 1, 1, 15).clearContent().clearDataValidations();
  
  // 2. 💡 THE FIX: Target ONLY Column P (16). Wipe any "FALSE" text, then cleanly plant the checkbox.
  // Using the exact syntax provided to ensure a pure, working checkbox.
  newLog.getRange(spacerRow, 16).clearContent().clearDataValidations();
  newLog.getRange(spacerRow, 16).insertCheckboxes().setFontColor("#000000");

  // Repaint all background formulas instantly (Like the Week formula in A)
  // 🚨 FIX: Changed 'ss' to 'newSS' so it uses the correct defined variable!
  if (typeof fillContactLogHstackFormulas_ === "function") {
    fillContactLogHstackFormulas_(newSS); 
  }
  
  // ==============================================================
  // --- NEW: AUTO-APPLY MASTER FILTER VISUALLY ---
  // ==============================================================
  try {
    // Ensure the END row itself is clean in Column P (16)
    newLog.getRange(newEndRow, 16).clearContent();
    
    // Clear any old tangled filters
    let filter = newLog.getFilter();
    if (filter) filter.remove();
    
    // Re-cast the fresh filter box encompassing the data up to Column P (16)
    // 💡 newEndRow - 1 safely encapsulates the data AND the guaranteed blank spacer row!
    filter = newLog.getRange(1, 1, newEndRow - 1, 16).createFilter();
    
    // Force filter visibility criteria (Only show items where font color is black)
    const blackColor = SpreadsheetApp.newColor().setRgbColor("#000000").build();
    const criteria = SpreadsheetApp.newFilterCriteria()
      .setVisibleForegroundColor(blackColor)
      .build();
    
    filter.setColumnFilterCriteria(16, criteria); // Tie the rule to Column P (16)
    SpreadsheetApp.flush(); // Lock the filter in first
    
  } catch(e) {
    console.error("Filter reset failed:", e);
  }
  
  SpreadsheetApp.getUi().alert(`Success! Migrated ${oldRecords.length} records into your Typed Table and applied your active Master Filter.`);
}

/**
 * STEP 6: Final Rebuild
 */
function ma_finish() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("Migration complete! Rebuilding master directories...", "Almost Done", 5);
  
  if (typeof buildAllDerivedSheets === "function") {
    buildAllDerivedSheets();
  }
  
  // 💡 NEW: Instantly Unlock and Repaint the App Menu without requiring a refresh!
  PropertiesService.getDocumentProperties().setProperty('FULL_MENU_UNLOCKED', 'true');
  if (typeof onOpen === "function") {
    try { 
      onOpen(); 
    } catch(e) {}
  }
  
  return { success: true };
}
