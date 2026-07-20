// ==========================================
// ⚙️ SETTINGS BACKEND LOGIC
// ==========================================
function openSettingsDialog() {
  const html = HtmlService.createHtmlOutputFromFile('SettingsDialog')
      .setTitle('Dashboard Settings')
      .setWidth(360)
      .setHeight(830); 
  SpreadsheetApp.getUi().showModalDialog(html, 'Settings');
}

// =========================================
//  MAIN SETTINGS MEMORY 
// =========================================

function getSettings() {
  const props = PropertiesService.getDocumentProperties();
  return {
    syncToPhones: props.getProperty('syncToPhones') === 'true',
    
    // --- CHANGED: Tabs now strictly default to FALSE (Hidden) ---
    phoneContacts: props.getProperty('phoneContacts') === 'true', 
    parentsDivided: props.getProperty('parentsDivided') === 'true', 
    notesTab: props.getProperty('notesTab') === 'true', 
    sendOut: props.getProperty('sendOut') === 'true', 
    
    // --- CHANGED: Columns now strictly default to FALSE (Hidden) ---
    siteClass: props.getProperty('siteClass') === 'true',
    typeCol: props.getProperty('typeCol') === 'true',
    
    // --- LIVE UPDATE SETTINGS ---
    disableRebuilds: props.getProperty('disableRebuilds') === 'true',
    disableNightlySync: props.getProperty('disableNightlySync') === 'true',
    autoNotes: props.getProperty('autoNotes') !== 'false',

    // Contact Log auto-fill of Date / # / Follow-up (default OFF = auto-fill ON).
    // Does NOT affect the parent/student/OSIS lookup.
    disableContactAutofill: props.getProperty('disableContactAutofill') === 'true',
    
    // --- BYPASS SWITCH & DATES ---
    disableHiddenNotes: props.getProperty('disableHiddenNotes') === 'true', 
    notesStartDate: props.getProperty('notesStartDate') || '',
    notesEndDate: props.getProperty('notesEndDate') || '',

    // --- DATA MANAGEMENT ---
    hideDischarged: props.getProperty('hideDischarged') === 'true'
  };
}

function updateVisualSettings_(settings) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Tab Visibility ---
  const tabMap = {
    phoneContacts:  "Phone Contacts",
    parentsDivided: "Parents Divided",
    sendOut: "Send Out",
    notesTab: "Notes" 
  };

  for (const [key, sheetName] of Object.entries(tabMap)) {
    try {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) continue;
      if (settings[key]) sheet.showSheet();
      else sheet.hideSheet();
    } catch (e) {
      console.warn("Could not hide/show sheet: " + sheetName);
    }
  }

  // --- Contact Log Column Visibility ---
  try {
    const contactLog = ss.getSheetByName("Contact Log");
    if (contactLog) {
      if (settings.siteClass) {
        contactLog.showColumns(8, 2);
      } else {
        contactLog.hideColumns(8, 2);
      }

      if (settings.typeCol) {
        contactLog.showColumns(12);
      } else {
        contactLog.hideColumns(12);
      }
    }
  } catch (e) {
    console.warn("Could not update Contact Log columns.");
  }

  // --- PCAR Visibility Logic ---
  try {
    const pcarSheet = ss.getSheetByName("PCAR");
    if (pcarSheet && pcarSheet.getMaxRows() >= 12) {
      if (settings.sendOut) {
        pcarSheet.showRows(11, 2); // Shows row 11 and 12
      } else {
        pcarSheet.hideRows(11, 2); // Hides row 11 and 12
      }
    }
  } catch (e) {
    console.warn("Could not update PCAR rows.");
  }

  // --- Master Table Filter (Hide Discharged) ---
  try {
    const masterSheet = ss.getSheetByName("Master Table");
    if (masterSheet) {
      const range = masterSheet.getDataRange();
      let filter = range.getFilter();
      
      if (settings.hideDischarged) {
        if (!filter) filter = range.createFilter();
        
        const headers = range.getValues()[0];
        if (headers) {
          const statusCol = headers.indexOf("Status") + 1; 
          
          if (statusCol > 0) {
            const criteria = SpreadsheetApp.newFilterCriteria().setHiddenValues(["Discharged"]).build();
            filter.setColumnFilterCriteria(statusCol, criteria);
          }
        }
      } else {
        if (filter) filter.remove();
      }
    }
  } catch (e) {
    console.warn("Could not update Master Table filter.");
  }
}

// --- TRIGGER MANAGEMENT ---
function manageNightlyTrigger_(disableSync) {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const syncFunctionName = 'nightlyMasterSync'; 

    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === syncFunctionName) {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }

    if (!disableSync) {
      ScriptApp.newTrigger(syncFunctionName)
        .timeBased()
        .atHour(2) 
        .everyDays(1)
        .create();
    }
  } catch (e) {
    console.warn("Trigger management failed: " + e.message);
  }
}

// --- PERMANENT PURGE FUNCTION ---
function purgeDischargedStudents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("Master Table");
  if (!masterSheet) return false;
  
  const data = masterSheet.getDataRange().getValues();
  if (data.length < 2) return 0; 
  
  const headers = data[0];
  const statusColIdx = headers.indexOf("Status");
  if (statusColIdx === -1) return 0;
  
  const currentData = [headers];
  let purgedCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    const status = String(data[i][statusColIdx]).trim();
    if (status.toLowerCase() !== "discharged") {
      currentData.push(data[i]);
    } else {
      purgedCount++;
    }
  }
  
  if (purgedCount > 0) {
    const filter = masterSheet.getFilter();
    if (filter) filter.remove(); 
    
    masterSheet.clearContents();
    masterSheet.getRange(1, 1, currentData.length, currentData[0].length).setValues(currentData);
    
    const props = PropertiesService.getDocumentProperties();
    if (props.getProperty('hideDischarged') === 'true') {
      const newFilter = masterSheet.getDataRange().createFilter();
      const criteria = SpreadsheetApp.newFilterCriteria().setHiddenValues(["Discharged"]).build();
      newFilter.setColumnFilterCriteria(statusColIdx + 1, criteria);
    }
  }
  
  return purgedCount;
}

function applySettings(settings) {
  const props = PropertiesService.getDocumentProperties();
  
  const oldDisableHidden = props.getProperty('disableHiddenNotes') === 'true';
  const oldStartDate = props.getProperty('notesStartDate') || '';
  const oldEndDate = props.getProperty('notesEndDate') || '';

  const newDisableHidden = settings.disableHiddenNotes === true;

  // 2. Save settings - Using String() safely handles undefined payload data instead of crashing!
  props.setProperty('syncToPhones', String(settings.syncToPhones));
  props.setProperty('phoneContacts', String(settings.phoneContacts));
  props.setProperty('parentsDivided', String(settings.parentsDivided));
  props.setProperty('notesTab', String(settings.notesTab)); 
  props.setProperty('sendOut', String(settings.sendOut));
  props.setProperty('siteClass', String(settings.siteClass));
  props.setProperty('typeCol', String(settings.typeCol));
  props.setProperty('disableRebuilds', String(settings.disableRebuilds));
  props.setProperty('disableNightlySync', String(settings.disableNightlySync));
  props.setProperty('autoNotes', String(settings.autoNotes));
  props.setProperty('disableContactAutofill', String(settings.disableContactAutofill));
  props.setProperty('disableHiddenNotes', String(newDisableHidden));
  props.setProperty('notesStartDate', settings.notesStartDate || '');
  props.setProperty('notesEndDate', settings.notesEndDate || '');
  props.setProperty('hideDischarged', String(settings.hideDischarged));
  
  // 3. Update Visuals & Master Table Filter safely
  if (typeof updateVisualSettings_ === "function") {
    updateVisualSettings_(settings);
  }
  
  // 4. Update Triggers Programmatically safely
  manageNightlyTrigger_(settings.disableNightlySync);
  
  // 5. Rebuild Menu Instantly (Calling the true onOpen directly)
  try {
    if (typeof onOpen === "function") {
      onOpen(); 
    }
  } catch (e) {
    console.warn("Could not instantly rebuild UI menus: " + e.message);
  }

  // 6. Check if we need to rebuild and show the appropriate toast
  let needRebuild = (oldDisableHidden !== newDisableHidden) || 
                    (oldStartDate !== settings.notesStartDate) || 
                    (oldEndDate !== settings.notesEndDate);
  
  if (settings.disableRebuilds) {
    needRebuild = false;
  }
  
  if (needRebuild) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Settings saved. Updating notes across sheets in the background...", 
      "Processing", 
      10
    );
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Settings applied successfully.", 
      "Settings Saved", 
      5
    );
  }

  return needRebuild; 
}
