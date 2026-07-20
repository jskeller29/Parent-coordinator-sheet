/**
 * The master orchestrator function.
 * Reads the 'Raw Data' tab once, builds all internal data maps,
 * and then generates/updates all the derived output sheets.
 */
function buildAllDerivedSheets() { 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName("Raw Data");
  if (!rawSheet) throw new Error('Sheet "Raw Data" not found.');
  const rawData = rawSheet.getDataRange().getValues();
  if (rawData.length < 2) throw new Error('"Raw Data" does not have enough rows.');

  // FIX 3: Pre-read logs ONCE. Uses getDisplayValues() for perfectly formatted string dates.
  const contactLogSheet = ss.getSheetByName("Contact Log");
  const contactLogData = contactLogSheet ? contactLogSheet.getDataRange().getDisplayValues() : [];
  const eventLogSheet = ss.getSheetByName("Backend_Event_Log");
  let eventLogData = eventLogSheet ? eventLogSheet.getDataRange().getValues() : [];

  // =========================================================
  // --- PRE-FETCH OLD DATA SNAPSHOT FOR REPORTS ---
  // =========================================================
  const oldPdMap = new Map();
  const oldPdSheet = ss.getSheetByName("Parents Divided");
  if (oldPdSheet) {
    const oldPdData = oldPdSheet.getDataRange().getValues();
    for (let i = 1; i < oldPdData.length; i++) {
      oldPdMap.set(String(oldPdData[i][1]).trim(), { ps: oldPdData[i][8], nycsa: oldPdData[i][9] });
    }
  }

  const oldMasterStatusMap = new Map();
  const oldMasterSheetForReport = ss.getSheetByName("Master Table");
  if (oldMasterSheetForReport) {
    const oldMasterData = oldMasterSheetForReport.getDataRange().getValues();
    for (let i = 4; i < oldMasterData.length; i++) {
      const osis = String(oldMasterData[i][0]).trim();
      if (osis && !osis.includes(",")) oldMasterStatusMap.set(osis, oldMasterData[i][5]);
    }
  }

  // --- Step 0: Pre-read the Override Sheet --- 
  const overrides = getOverrideData_(ss);

  // --- Step 1: Build Core Data Maps in Memory --- 
  const studentMap = buildStudentMap_(rawData, overrides.students);
  const guardianRows = buildGuardianObjects_(rawData, overrides.guardians);
  const siblingMap = buildSiblingMap_(studentMap, guardianRows);
  const parentSquareIndex = buildParentSquareIndex_(rawData);
  const nycsaIndex = buildNYCSAIndex_(rawData);

  // --- Step 2: Build and Write Output Sheets ---
  const masterTable = buildMasterTableData_(ss, studentMap, guardianRows, siblingMap);
  writeSheet_(ss, "Master Table", masterTable, 4, 1);

// ==========================================
  // REBUILD BACKEND EVENT LOG
  // ==========================================
  const backendEventSheet = ss.getSheetByName("Backend_Event_Log");
  if (backendEventSheet) {
    // 1. Call our updated helper function to scrape and split the Events sheet
    const freshEventData = buildBackendEventLogData_(ss);
    
    // 2. Clear out any old, dead, or ghost data from the backend
    backendEventSheet.clearContents();
    
    // 3. Paste the perfectly parsed, chronologically sorted data back in
    if (freshEventData && freshEventData.length > 0) {
      backendEventSheet.getRange(1, 1, freshEventData.length, freshEventData[0].length).setValues(freshEventData);
      
      // ⚡ THE FIX: Overwrite the old memory array with the perfectly formatted new data!
      eventLogData = freshEventData; 
    }
  }

  // Build Contact Tracking FIRST so we can use its dates
  const combinedContactTracking = buildCombinedContactTrackingData_(ss, masterTable, contactLogData, eventLogData);
  writeSheet_(ss, "Combined Contact Tracking", combinedContactTracking);

  const parentsDivided = buildParentsDividedData_(ss, guardianRows, siblingMap, parentSquareIndex, nycsaIndex, studentMap, contactLogData, eventLogData);
  writeSheet_(ss, "Parents Divided", parentsDivided);

  const directory = buildDirectoryData_(studentMap, guardianRows);
  writeSheet_(ss, "Directory", directory);

  if (typeof formatDirectoryAuto_ === "function") formatDirectoryAuto_(ss.getSheetByName("Directory"));

  // UPDATED: Pass the already-built parentSquareIndex and nycsaIndex down
  const phoneContacts = buildPhoneContactsData_(ss, studentMap, guardianRows, masterTable, combinedContactTracking, parentSquareIndex, nycsaIndex);
  writeSheet_(ss, "Phone Contacts", phoneContacts);

  // --- Sync Column Visibility (Static Schema) ---
  if (typeof fillContactLogHstackFormulas_ === "function") {
    fillContactLogHstackFormulas_(ss);
  }

  const dirSheet = ss.getSheetByName("Directory");
  let hasSite = false;
  for (const student of studentMap.values()) {
    if (student.site && String(student.site).trim() !== "") {
      hasSite = true;
      break;
    }
  }

  if (dirSheet) {
    if (hasSite) dirSheet.showColumns(4);
    else dirSheet.hideColumns(4);
  }
  if (contactLogSheet) {
    if (hasSite) contactLogSheet.showColumns(8);
    else contactLogSheet.hideColumns(8);
  }

  // =========================================================
  // --- GENERATE BUILD REPORT ---
  // =========================================================
  const newStudentsLog = [];
  const dischargedLog = [];
  const nycsaLog = [];
  const psLog = [];

  for (let i = 1; i < masterTable.length; i++) {
    const osis = String(masterTable[i][0]).trim();
    const name = masterTable[i][1];
    const status = String(masterTable[i][5]).trim();
    if (osis && !osis.includes(",")) {
      const oldStatus = oldMasterStatusMap.get(osis);
      if (status === "Discharged" && oldStatus !== "Discharged" && oldMasterStatusMap.size > 0) dischargedLog.push(`${name} (${osis})`);
      if (status !== "Discharged" && oldStatus === "Discharged" && oldMasterStatusMap.size > 0) newStudentsLog.push(`${name} (${osis})`);
      if (status !== "Discharged" && !oldMasterStatusMap.has(osis) && oldMasterStatusMap.size > 0) newStudentsLog.push(`${name} (${osis})`);
    }
  }

  if (oldPdMap.size > 0) {
    for (let i = 1; i < parentsDivided.length; i++) {
      const gName = String(parentsDivided[i][1]).trim();
      const student = parentsDivided[i][2];
      const newPs = String(parentsDivided[i][8]).trim();
      const newNycsa = String(parentsDivided[i][9]).trim();

      if (oldPdMap.has(gName)) {
        const old = oldPdMap.get(gName);
        const oldPs = String(old.ps).trim();
        const oldNycsa = String(old.nycsa).trim();

        if (oldNycsa.startsWith("No") && (newNycsa.startsWith("Yes") || newNycsa === "Yes/Unknown")) {
          nycsaLog.push(`${gName} (Student: ${student})`);
        }
        if (oldPs === "No" && newPs === "Yes") {
          psLog.push(`${gName} (Student: ${student})`);
        }
      }
    }
  }

  CacheService.getScriptCache().put('BUILD_REPORT_LOGS', JSON.stringify({
    newStudents: newStudentsLog,
    discharged: dischargedLog,
    nycsaUpgrades: nycsaLog,
    psUpgrades: psLog
  }), 21600);

  // FIX 4: Set Montserrat font ONCE at the very end instead of inside writeSheet_
  ["Master Table", "Parents Divided", "Directory", "Combined Contact Tracking", "Phone Contacts"]
    .forEach(name => {
      const sh = ss.getSheetByName(name);
      if (sh) sh.getDataRange().setFontFamily("Montserrat");
    });
}
