/**
 * ==============================================================================
 * STUDENT OVERRIDES BACKEND
 * Handles saving overrides, memory settings, and surgical spreadsheet updates.
 * ==============================================================================
 */

function openNewStudentDialog() {
  const html = HtmlService.createHtmlOutputFromFile('NewStudentDialog')
      .setTitle('New / Edit Student Override')
      .setWidth(600) 
      .setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, '✨ New / Edit Student');
}

function openManageOverridesDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ManageOverridesDialog')
      .setTitle('Manage Student Overrides')
      .setWidth(500)
      .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, '🗂️ Manage Overrides');
}

function safeStr_(val) {
  if (val === null || val === undefined) return "";
  return String(val).trim();
}

/**
 * Triggered when the user clicks 'Save' in the HTML pop-up.
 * Bulk saves an array of student payloads, then hands off the heavy processing to a background trigger.
 */
function saveStudentOverrides(payloads) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("New/Edit Student");

  if (!sheet) throw new Error("Could not find the 'New/Edit Student' tab! Please ensure it is created.");

  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");
  const existingData = sheet.getDataRange().getValues();
  const osisSet = new Set(payloads.map(p => safeStr_(p.osis)));

  // 1. DUPLICATE REMOVAL
  for (let i = existingData.length - 1; i > 0; i--) {
    const rowOsis = safeStr_(existingData[i][2]); 
    if (rowOsis !== "" && osisSet.has(rowOsis)) sheet.deleteRow(i + 1);
  }

  // 2. COMPILE ALL NEW ROWS IN MEMORY
  const rowsToAppend = payloads.map(data => {
      // Expanded to 42 columns to support Label and Hide From Report
      const row = new Array(42).fill("");
      row[0] = dateStr;           
      row[1] = data.entryType;    
      row[2] = data.osis;         
      row[3] = data.lastName;     
      row[4] = data.firstName;    
      row[5] = data.gender;       
      row[6] = data.age;          
      row[7] = data.grade;        
      row[8] = data.classCode;    
      row[9] = data.homeLanguage; 

      // Guardians Section (Up to 5)
      for (let i = 0; i < data.guardians.length; i++) {
        if (i >= 5) break;
        const g = data.guardians[i];
        const baseCol = 10 + (i * 6);
        row[baseCol] = g.name;
        row[baseCol + 1] = g.relationship;
        row[baseCol + 2] = g.phone;
        row[baseCol + 3] = g.email;
        // Map the single language UI input perfectly to BOTH legacy columns
        row[baseCol + 4] = g.language;
        row[baseCol + 5] = g.language;
      }
      
      // Save Label & Stealth Mode
      row[40] = data.label;
      row[41] = data.hideFromReport;

      return row;
  });

  // 3. BULK PASTE INSTANTLY
  if (rowsToAppend.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, 42).setValues(rowsToAppend);
  }
  SpreadsheetApp.flush();

 // 4. CHECK SETTINGS & TRIGGER ASYNC REBUILD
  const settings = typeof getSettings === "function" ? getSettings() : {}; 
  const isPaused = (settings.disableRebuilds === true || settings.disableRebuilds === 'true');

  if (!isPaused) {
      createAsyncOverrideTrigger(Array.from(osisSet));
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "Student data saved! The spreadsheet will begin rebuilding in the background in ~1 minute.", 
        "⏳ Update Queued", 8
      );
      return "Rebuilding in background...";
  } else {
      SpreadsheetApp.getActiveSpreadsheet().toast("Overrides saved locally. Background sync is paused via Settings.", "⚡ Fast Entry Mode", 4);
      return "Saved locally.";
  }
}

function createAsyncOverrideTrigger(osisArray) {
   const cache = CacheService.getScriptCache();
   const batchId = "OVERRIDE_BATCH_" + Date.now();
   cache.put(batchId, JSON.stringify(osisArray), 600); 
   PropertiesService.getDocumentProperties().setProperty('PENDING_OVERRIDE_BATCH', batchId);

   const triggers = ScriptApp.getProjectTriggers();
   triggers.forEach(t => {
     if (t.getHandlerFunction() === "processAsyncOverrides") ScriptApp.deleteTrigger(t);
   });

   ScriptApp.newTrigger("processAsyncOverrides").timeBased().after(500).create();
}

function processAsyncOverrides() {
   const triggers = ScriptApp.getProjectTriggers();
   triggers.forEach(t => {
     if (t.getHandlerFunction() === "processAsyncOverrides") ScriptApp.deleteTrigger(t);
   });

   const batchId = PropertiesService.getDocumentProperties().getProperty('PENDING_OVERRIDE_BATCH');
   if (!batchId) return;

   const cache = CacheService.getScriptCache();
   const dataStr = cache.get(batchId);
   PropertiesService.getDocumentProperties().deleteProperty('PENDING_OVERRIDE_BATCH');
   
   if (!dataStr) return;

   try {
       const ss = SpreadsheetApp.getActiveSpreadsheet();
       if (ss) ss.toast("Background rebuild is starting now...", "⚙️ Building Sheets", 5);

       if (typeof buildAllDerivedSheets === "function") buildAllDerivedSheets();
       
       const osisArray = JSON.parse(dataStr);
       if (osisArray && osisArray.length > 0 && typeof updateSingleOsisDelta_ === "function") {
           osisArray.forEach(osis => updateSingleOsisDelta_(osis));
       }

       if (ss) ss.toast("Background build complete! Your new student overrides are fully integrated.", "✅ Success", 8);
   } catch (e) {
       console.error("Background Rebuild Error: " + e.message);
       const ss = SpreadsheetApp.getActiveSpreadsheet();
       if (ss) ss.toast("Background rebuild failed. Please run a manual build.", "❌ Error", 10);
   }
}

// =========================================
//  DATA FETCHERS & MANAGE UI LOGIC
// =========================================

function getAllStudentsForEdit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName("Raw Data");
  const overrideSheet = ss.getSheetByName("New/Edit Student");
  const masterSheet = ss.getSheetByName("Master Table");

  // Dynamically map current Statuses from the Master Table
  const statusMap = new Map();
  if (masterSheet) {
    const mData = masterSheet.getDataRange().getValues();
    let osisIdx = 0, statIdx = 6; 
    for(let i=0; i < Math.min(10, mData.length); i++) {
        for(let c=0; c<mData[i].length; c++) {
            if(String(mData[i][c]).toUpperCase() === "OSIS") osisIdx = c;
            if(String(mData[i][c]).toUpperCase() === "STATUS") statIdx = c;
        }
    }
    for(let i = 4; i < mData.length; i++) {
        statusMap.set(safeStr_(mData[i][osisIdx]), safeStr_(mData[i][statIdx]));
    }
  }

  const studentMap = new Map();

  // 1. Pull ATS Base Data
  if (rawSheet) {
    const rawData = rawSheet.getDataRange().getValues();
    for (let i = 1; i < rawData.length; i++) {
      const r = rawData[i];
      const osis = safeStr_(r[2]);
      if (osis.replace(/\D/g, "").length < 6) continue;

      const g1Name = safeStr_(r[19]);
      const g2Name = safeStr_(r[25]);
      const guardians = [];

      if (g1Name) guardians.push({ name: g1Name, relationship: safeStr_(r[20]), phone: safeStr_(r[21]), email: safeStr_(r[22]), language: safeStr_(r[23]) || safeStr_(r[24]) });
      if (g2Name) guardians.push({ name: g2Name, relationship: safeStr_(r[26]), phone: safeStr_(r[27]), email: safeStr_(r[28]), language: safeStr_(r[29]) || safeStr_(r[30]) });

      studentMap.set(osis, {
        osis: osis,
        lastName: safeStr_(r[3]),
        firstName: safeStr_(r[4]),
        gender: safeStr_(r[6]),
        age: safeStr_(r[7]),
        grade: safeStr_(r[16]),
        classCode: safeStr_(r[17]),
        homeLanguage: safeStr_(r[18]),
        guardians: guardians,
        label: statusMap.get(osis) || "Current",
        hideFromReport: false
      });
    }
  }

  // 2. Pull Overrides (And overwrite any base data)
  if (overrideSheet) {
    const overData = overrideSheet.getDataRange().getValues();
    for (let i = 1; i < overData.length; i++) {
      const r = overData[i];
      const osis = safeStr_(r[2]);
      if (osis.replace(/\D/g, "").length < 6) continue;

      const guardians = [];
      for (let g = 0; g < 5; g++) {
        const baseCol = 10 + (g * 6);
        const name = safeStr_(r[baseCol]);
        if (name) {
          guardians.push({
            name: name,
            relationship: safeStr_(r[baseCol+1]),
            phone: safeStr_(r[baseCol+2]),
            email: safeStr_(r[baseCol+3]),
            language: safeStr_(r[baseCol+4]) || safeStr_(r[baseCol+5])
          });
        }
      }

      studentMap.set(osis, {
        osis: osis,
        lastName: safeStr_(r[3]),
        firstName: safeStr_(r[4]),
        gender: safeStr_(r[5]),
        age: safeStr_(r[6]),
        grade: safeStr_(r[7]),
        classCode: safeStr_(r[8]),
        homeLanguage: safeStr_(r[9]),
        guardians: guardians,
        label: safeStr_(r[40]) || statusMap.get(osis) || "Current",
        hideFromReport: String(r[41]).toLowerCase() === 'true'
      });
    }
  }

  return Array.from(studentMap.values()).sort((a, b) => a.lastName.localeCompare(b.lastName));
}

function getOverridesList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("New/Edit Student");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const list = [];

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const osis = safeStr_(r[2]);
    if (osis.replace(/\D/g, "").length < 6) continue;

    let dateStr = r[0];
    if (Object.prototype.toString.call(dateStr) === "[object Date]" && !isNaN(dateStr)) {
       dateStr = Utilities.formatDate(dateStr, Session.getScriptTimeZone(), "MM/dd/yyyy");
    }

    list.push({
      date: dateStr,
      type: safeStr_(r[1]),
      osis: osis,
      lastName: safeStr_(r[3]),
      firstName: safeStr_(r[4]),
      classCode: safeStr_(r[8])
    });
  }
  return list.reverse();
}

function deleteOverride(osis) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("New/Edit Student");
  if (!sheet) throw new Error("Sheet not found");

  const data = sheet.getDataRange().getValues();
  let deleted = false;

  for (let i = data.length - 1; i > 0; i--) {
    if (safeStr_(data[i][2]) === safeStr_(osis)) {
      sheet.deleteRow(i + 1);
      deleted = true;
    }
  }

  if (deleted) {
    const settings = typeof getSettings === "function" ? getSettings() : {};
    const isPaused = (settings.disableRebuilds === true || settings.disableRebuilds === 'true');

    if (!isPaused) {
      createAsyncOverrideTrigger([osis]);
      SpreadsheetApp.getActiveSpreadsheet().toast("Override removed! The spreadsheet will begin rebuilding in the background in ~1 minute.", "⏳ Update Queued", 8);
    } else {
      SpreadsheetApp.getActiveSpreadsheet().toast("Override deleted locally. Sheets will update on next manual build.", "⚡ Fast Mode", 4);
    }
  }
  return true;
}

function setPrefillAndOpenEdit(osis) {
  PropertiesService.getUserProperties().setProperty('PREFILL_OSIS', osis);
  openNewStudentDialog();
}

function getPrefillOsis() {
  const props = PropertiesService.getUserProperties();
  const osis = props.getProperty('PREFILL_OSIS');
  props.deleteProperty('PREFILL_OSIS');
  return osis || "";
}
