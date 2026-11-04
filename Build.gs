// ======================================================================
// FILE: Build.gs
// PURPOSE: The full rebuild ("Build Sheets"): reads Raw Data once and writes the
// Master Table, Combined Contact Tracking, Parents Divided, Directory, Phone
// Contacts and New Students tabs, plus the build report for the sidebar.
// ======================================================================

/**
 * The master orchestrator function.
 * Reads the 'Raw Data' tab once, builds all internal data maps,
 * and then generates/updates all the derived output sheets.
 * On the owner's sheet (z_MINE.gs) a few steps use the owner's versions:
 * history from "Old Master Table" / "Old Phone Contacts" and the tall
 * Phone Contacts layout (see ownerHooks_).
 */
function buildAllDerivedSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName("Raw Data");
  if (!rawSheet) throw new Error('Sheet "Raw Data" not found.');
  const rawData = rawSheet.getDataRange().getValues();
  if (rawData.length < 2) throw new Error('"Raw Data" does not have enough rows.');

  const owner = getOwnerHooks_();

  // Pre-read logs ONCE. Uses getDisplayValues() for perfectly formatted string dates.
  const contactLogSheet = ss.getSheetByName("Contact Log");
  const contactLogData = contactLogSheet ? contactLogSheet.getDataRange().getDisplayValues() : [];
  let eventLogData = backendReadBlock_(ss, "EVENT_LOG");

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

  // OSIS -> Status from the Master Table before it is rewritten (build report + New Students log)
  const oldMasterStatusMap = owner ? owner.readMasterStatusSnapshot(ss) : readMasterStatusSnapshot_(ss);

  // --- Step 0: Pre-read the Override Sheet ---
  // Remove New Student overrides that ATS has caught up with (and exact duplicates of Raw Data)
  // Cleanup must never block the build, so failures are only logged
  try { cleanRedundantOverrides(); } catch (e) { console.warn("Override cleanup skipped: " + e.message); }
  const overrides = getOverrideData_(ss);

  // --- Step 1: Build Core Data Maps in Memory ---
  const studentMap = buildStudentMap_(rawData, overrides.students);
  const guardianRows = buildGuardianObjects_(rawData, overrides.guardians);
  const siblingMap = buildSiblingMap_(studentMap, guardianRows);
  const parentSquareIndex = buildParentSquareIndex_(rawData);
  const nycsaIndex = buildNYCSAIndex_(rawData);

  // --- Step 2: Build and Write Output Sheets ---
  let masterTable = buildMasterTableData_(ss, studentMap, guardianRows, siblingMap);
  if (owner) masterTable = owner.adjustMasterTable(ss, masterTable); // Old Master Table history
  writeSheet_(ss, "Master Table", masterTable, 4, 1);

  // Log students who just appeared in Raw Data (ATS) to the "New Students" tab.
  // New / discharged students come from Raw Data only (same result feeds the tab and the build report)
  const logAtsChanges = () => {
    const changes = detectAtsChanges_(ss, masterTable, oldMasterStatusMap, rawData);
    try {
      logNewStudents_(ss, masterTable, changes, rawData, guardianRows);
    } catch (e) { console.warn("New Students log skipped: " + e.message); }
    return changes;
  };
  // Shared sheets log right after the Master Table; the owner's sheet after the Directory (as before)
  let atsChanges = owner ? null : logAtsChanges();

  // ==========================================
  // REBUILD BACKEND EVENT LOG
  // ==========================================
  {
    // 1. Call our updated helper function to scrape and split the Events sheet
    const freshEventData = buildBackendEventLogData_(ss);

    // 2+3. Replace the Backend tab's Event Log section (clears old/ghost rows;
    // only its own columns, so the other Backend sections are untouched)
    if (freshEventData && freshEventData.length > 0) {
      backendWriteBlock_(ss, "EVENT_LOG", freshEventData);

      // ⚡ THE FIX: Overwrite the old memory array with the perfectly formatted new data!
      eventLogData = freshEventData;
    }
  }

  // Build Contact Tracking FIRST so we can use its dates
  const combinedContactTracking = buildCombinedContactTrackingData_(ss, masterTable, contactLogData, eventLogData);
  writeSheet_(ss, "Combined Contact Tracking", combinedContactTracking);

  const parentsDivided = buildParentsDividedData_(ss, guardianRows, siblingMap, parentSquareIndex, nycsaIndex, studentMap, contactLogData, eventLogData);
  writeSheet_(ss, "Parents Divided", parentsDivided);

  // Directory draws only from Raw Data: no New/Edit Student overrides, no discharged students
  // (the owner's sheet then merges in Old Phone Contacts and "MASTER" Old Master Table students)
  let directory = buildDirectoryData_(buildStudentMap_(rawData, new Map()), buildGuardianObjects_(rawData, new Map()));
  if (owner) directory = owner.adjustDirectory(ss, directory);
  writeSheet_(ss, "Directory", directory);

  if (owner) atsChanges = logAtsChanges();

  formatDirectoryAuto_(ss.getSheetByName("Directory"));

  // UPDATED: Pass the already-built parentSquareIndex and nycsaIndex down
  const phoneContacts = buildPhoneContactsData_(ss, studentMap, guardianRows, masterTable, combinedContactTracking, parentSquareIndex, nycsaIndex);
  writeSheet_(ss, "Phone Contacts", phoneContacts);

  // --- Sync Column Visibility (Static Schema) ---
  fillContactLogHstackFormulas_(ss);

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
    const siteCol = owner ? owner.contactLogSiteColumn : 8;
    if (hasSite) contactLogSheet.showColumns(siteCol);
    else contactLogSheet.hideColumns(siteCol);
  }

  // =========================================================
  // --- GENERATE BUILD REPORT ---
  // =========================================================
  const { newStudents: newStudentsLog, discharged: dischargedLog } = atsChangeReportLines_(masterTable, atsChanges);
  const nycsaLog = [];
  const psLog = [];

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

  // Remember this build's Raw Data roster for the next comparison
  try { saveAtsSnapshot_(ss, atsChanges.atsOsis); } catch (e) { console.warn("ATS snapshot not saved: " + e.message); }

  // Set Montserrat font ONCE at the very end instead of inside writeSheet_
  const fontSheets = owner ? owner.montserratSheets
    : ["Master Table", "Parents Divided", "Directory", "Combined Contact Tracking", "Phone Contacts", "New Students"];
  fontSheets.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) sh.getDataRange().setFontFamily("Montserrat");
  });
}

/** Nightly trigger: full rebuild, Contact Log formatting, then (if opted in) the phone sync. */
function nightlyMasterSync() {
  return runAction_("Nightly sync", () => nightlyMasterSyncRun_());
}

function nightlyMasterSyncRun_() {
  try {
    buildAllDerivedSheets();
    resetContactLogFormatting();

    const props = PropertiesService.getDocumentProperties();
    const syncToPhones = props.getProperty('syncToPhones') === 'true';

    if (syncToPhones) {
       console.log("Nightly Rebuild finished. Phone sync TRIGGERED (Opted In).");
       syncToGoogleContacts();
    } else {
       console.log("Nightly Rebuild finished. Phone sync skipped (Opted Out).");
    }

    // Refresh the resolved links (template copies + User Guide) from their Drive
    // folders so the stored/inherited links stay current with no manual step
    // (see refreshResolvedLinks_ in Menus.gs). Safe/no-UI; never blocks the sync.
    try {
      refreshResolvedLinks_();
    } catch (e) { console.error("Resolved link refresh failed: " + e.message); }

  } catch(e) {
    console.error("Nightly Master Sync failed: " + e.message);
  }
}

// ==========================================
// 🚀 SIDEBAR BACKEND ENDPOINTS
// ==========================================
function getRunMode() { 
  return CacheService.getScriptCache().get('RUN_MODE') || 'FULL_SYNC'; 
}

function getBuildLogData() { 
  const cache = CacheService.getScriptCache().get('BUILD_REPORT_LOGS'); 
  return cache ? JSON.parse(cache) : { newStudents: [], discharged: [], nycsaUpgrades: [], psUpgrades: [] }; 
}

function startBuildOnlyProcess() {
  return runAction_("Build Sheets Only", () => startBuildOnlyProcessRun_());
}

function startBuildOnlyProcessRun_() {
  const cache = CacheService.getScriptCache();
  cache.put('SYNC_UI_STATE', JSON.stringify({ phase: 'building' }), 21600); 
  try { 
    buildAllDerivedSheets();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    formatDirectoryAuto_(ss.getSheetByName("Directory"));

    cache.put('SYNC_UI_STATE', JSON.stringify({ phase: 'done_build_only' }), 21600); 
  } catch(e) { 
    console.error("Build failed: " + (e.stack || e.message));
    cache.put('SYNC_UI_STATE', JSON.stringify({ phase: 'killed', error: e.message || String(e) }), 21600); 
  } 
}

function cleanRedundantOverrides() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName("RAW Data") || ss.getSheetByName("Raw Data");
  if (!rawSheet) return;

  const rawData = rawSheet.getDataRange().getValues();
  const overrideData = backendReadBlock_(ss, "OVERRIDES");

  const rawMap = new Map();
  for (let i = 6; i < rawData.length; i++) {
    const osis = String(rawData[i][2]).trim();
    if (osis) rawMap.set(osis, rawData[i]);
  }

  function extractGuardians(row, startCol, maxGuardians) {
    const guardians = [];
    for (let i = 0; i < maxGuardians; i++) {
      const base = startCol + (i * 6);
      const name = String(row[base] || "").trim();
      if (!name) continue; 
      
      guardians.push({
        name: name,
        phone: String(row[base + 2] || "").trim(),
        email: String(row[base + 3] || "").trim(),
        lang: String(row[base + 4] || "").trim() 
      });
    }
    return guardians;
  }

  let deletedOsisList = []; 

  for (let i = overrideData.length - 1; i >= 1; i--) {
    const oRow = overrideData[i];
    const osis = String(oRow[2]).trim();
    if (!osis) continue; 

    if (!rawMap.has(osis)) continue; 

    // "New Student" rows exist only until ATS has the student; once they're in Raw Data, drop them
    if (String(oRow[1]).trim() === "New Student") {
      backendDeleteRows_(ss, "OVERRIDES", [i]);
      deletedOsisList.push(osis);
      continue;
    }

    const rRow = rawMap.get(osis);

    const newStudentName = `${String(oRow[4] || "").trim()} ${String(oRow[3] || "").trim()}`.trim();
    const oldStudentName = `${String(rRow[4] || "").trim()} ${String(rRow[3] || "").trim()}`.trim();
    if (newStudentName.toLowerCase() !== oldStudentName.toLowerCase()) continue; 

    const newGuardians = extractGuardians(oRow, 10, 5);
    const oldGuardians = extractGuardians(rRow, 19, 2);

    const newNames = newGuardians.map(g => g.name.toLowerCase());
    const oldNames = oldGuardians.map(g => g.name.toLowerCase());

    let structureChanged = (newNames.length !== oldNames.length);
    if (!structureChanged) {
      for (let name of newNames) {
        if (!oldNames.includes(name)) structureChanged = true;
      }
    }
    if (structureChanged) continue; 

    let contactChanged = false;
    newGuardians.forEach(ng => {
      const og = oldGuardians.find(g => g.name.toLowerCase() === ng.name.toLowerCase());
      if (og && (og.phone !== ng.phone || og.email !== ng.email || og.lang !== ng.lang)) {
        contactChanged = true;
      }
    });
    
    if (contactChanged) continue; 

    backendDeleteRows_(ss, "OVERRIDES", [i]); 
    deletedOsisList.push(osis); 
  }

  if (deletedOsisList.length > 0) {
    let displayList = deletedOsisList.length > 5 
        ? deletedOsisList.slice(0, 5).join(", ") + ` (+${deletedOsisList.length - 5} more)`
        : deletedOsisList.join(", ");
        
    ss.toast(`Removed ${deletedOsisList.length} obsolete override(s) because ATS caught up.\nOSIS: ${displayList}`, "🧹 Auto-Cleanup", 8);
    console.log(`Self-Cleaning Complete: Removed ${deletedOsisList.length} redundant override(s). OSIS: ${deletedOsisList.join(", ")}`);
  }
}

/* =========================
   OVERRIDE MAP GENERATOR (New / Edit Student section of the Backend tab)
========================= */
function getOverrideData_(ss) {
  const overrideStudents = new Map();
  const overrideGuardians = new Map();

  // New/Edit Student section of the Backend tab (index 0 = header row)
  const data = backendReadBlock_(ss, "OVERRIDES");
  
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const id = toText_(r[2]);
    if (!isRealStudentId_(id)) continue;
    
    const lastName = toText_(r[3]);
    const firstName = toText_(r[4]);
    const classCode = toText_(r[8]).toUpperCase();
    
    overrideStudents.set(id, {
      student: properCase_(`${firstName} ${lastName}`.trim()),
      classCode: classCode,
      site: "",
      teachers: "",
      label: toText_(r[40]) // <-- Retrieves the new Custom Label
    });

    const guardianList = [];
    for (let g = 0; g < 5; g++) {
      const baseCol = 10 + (g * 6); 
      if (toText_(r[baseCol])) {
        const gInfo = buildGuardianObject_(
          id, 
          r[baseCol],
          r[baseCol + 1],
          r[baseCol + 4],
          r[baseCol + 5], // Passed twice in UI mapping
          r[baseCol + 3],
          r[baseCol + 2]
        );
        if (gInfo) guardianList.push(gInfo);
      }
    }
    overrideGuardians.set(id, guardianList);
  }

  return { students: overrideStudents, guardians: overrideGuardians };
}

/* =========================
   MASTER TABLE PIPELINE
========================= */
function buildStudentMap_(rawData, overrideStudents) {
  const map = new Map();
  const classRefMap = buildClassReferenceMap_(rawData);
  // Raw Data's Official Class is the source of truth for where a student is now
  const rawClassById = new Map();

  for (let i = 0; i < rawData.length; i++) {
    const r = rawData[i];
    const id = toText_(r[2]);
    if (!isRealStudentId_(id)) continue;

    const classCode = toText_(r[17]).toUpperCase();
    if (classCode && !rawClassById.has(id)) rawClassById.set(id, classCode);

    if (overrideStudents.has(id)) continue; 

    const studentName = properCase_(`${toText_(r[4])} ${toText_(r[3])}`.trim());

    let assignedSite = "";
    let assignedTeacher = "";
    if (classRefMap.has(classCode)) {
      const ref = classRefMap.get(classCode);
      assignedSite = ref.site;
      assignedTeacher = ref.teacher;
    }

    if (!map.has(id)) {
      map.set(id, {
        student: studentName,
        classCode: classCode,
        site: assignedSite,
        teachers: assignedTeacher,
        label: "Current" // Default ATS Status
      });
    }
  }

  for (const [id, studentObj] of overrideStudents.entries()) {
    // An override's class goes stale when ATS moves the student (e.g. Site 276 -> 188),
    // so prefer the current Raw Data class; override-only students keep their own.
    if (rawClassById.has(id)) studentObj.classCode = rawClassById.get(id);
    if (classRefMap.has(studentObj.classCode)) {
      const ref = classRefMap.get(studentObj.classCode);
      studentObj.site = ref.site;
      studentObj.teachers = ref.teacher;
    }
    map.set(id, studentObj);
  }

  return map;
}

/* =========================
   BASE BUILDERS
========================= */
function buildGuardianObjects_(rawData, overrideGuardians) {
  const out = [];

  for (let i = 0; i < rawData.length; i++) {
    const r = rawData[i];
    const id = toText_(r[2]);
    if (!isRealStudentId_(id)) continue;

    if (overrideGuardians.has(id)) continue;

    const g1 = buildGuardianObject_(id, r[19], r[20], r[23], r[24], r[22], r[21]);
    if (g1) out.push(g1);

    const g2 = buildGuardianObject_(id, r[25], r[26], r[29], r[30], r[28], r[27]);
    if (g2) out.push(g2);
  }

  for (const guardianList of overrideGuardians.values()) {
    out.push(...guardianList);
  }

  return out;
}

function buildGuardianObject_(studentId, fullName, relationship, spokenLanguage, writtenLanguage, email, phone) {
  const name = toText_(fullName);
  if (!name) return null;

  let language = "";
  if (toText_(spokenLanguage)) {
    language = properCase_(firstWord_(spokenLanguage));
  } else if (toText_(writtenLanguage)) {
    language = properCase_(firstWord_(writtenLanguage));
  }

  let relTag = properCase_(relationship);
  if (language && language.toLowerCase() !== "english") {
    relTag = relTag ? `${relTag} - ${language}` : language;
  }

  return {
    student_id: studentId,
    guardianOutput: properCase_(name) + (relTag ? ` (${relTag})` : ""),
    relationship: properCase_(relationship),
    language: language,
    email: toText_(email),
    mobile: formatPhone_(phone)
  };
}

function buildClassReferenceMap_(rawData) {
  const map = new Map();
  const regex = /([A-Za-z0-9]+):([^:]*?)(?=\s+[A-Za-z0-9]+:|$)/g;

  for (let i = 0; i < rawData.length; i++) {
    const siteVal = toText_(rawData[i][59]);
    const classesVal = toText_(rawData[i][60]);

    if (!classesVal) continue; 

    const matches = [...classesVal.matchAll(regex)];
    
    for (const m of matches) {
      const classCode = m[1].trim().toUpperCase();
      const teacher = m[2].trim();
      
      map.set(classCode, { 
        site: siteVal, 
        teacher: teacher 
      });
    }
  }
  return map;
}

/* =========================
   SIBLINGS
========================= */
function buildSiblingMap_(studentMap, guardianRows) {
  const nameEmailMap = new Map();
  const namePhoneMap = new Map();

  for (const g of guardianRows) {
    const id = g.student_id;
    if (!studentMap.has(id)) continue;

    const guardianName = stripTrailingParenthetical_(toText_(g.guardianOutput)).toLowerCase();
    if (!guardianName) continue; 

    const email = normalizeEmail_(g.email);
    const phone = normalizePhoneForMatch_(g.mobile);

    if (email) {
      const emailKey = `${guardianName}|${email}`;
      if (!nameEmailMap.has(emailKey)) nameEmailMap.set(emailKey, new Set());
      nameEmailMap.get(emailKey).add(id);
    }

    if (phone) {
      const phoneKey = `${guardianName}|${phone}`;
      if (!namePhoneMap.has(phoneKey)) namePhoneMap.set(phoneKey, new Set());
      namePhoneMap.get(phoneKey).add(id);
    }
  }

  const rowsByStudent = new Map();
  for (const g of guardianRows) {
    if (!rowsByStudent.has(g.student_id)) rowsByStudent.set(g.student_id, []);
    rowsByStudent.get(g.student_id).push(g);
  }

  const siblingMap = new Map();

  for (const [id] of studentMap) {
    const siblings = new Map();
    const rows = rowsByStudent.get(id) || [];

    for (const g of rows) {
      const guardianName = stripTrailingParenthetical_(toText_(g.guardianOutput)).toLowerCase();
      if (!guardianName) continue;

      const email = normalizeEmail_(g.email);
      const phone = normalizePhoneForMatch_(g.mobile);

      if (email) {
        const emailKey = `${guardianName}|${email}`;
        if (nameEmailMap.has(emailKey)) {
          nameEmailMap.get(emailKey).forEach(otherId => {
            if (otherId !== id && studentMap.has(otherId)) {
              siblings.set(otherId, `${otherId} - ${studentMap.get(otherId).student}`);
            }
          });
        }
      }

      if (phone) {
        const phoneKey = `${guardianName}|${phone}`;
        if (namePhoneMap.has(phoneKey)) {
          namePhoneMap.get(phoneKey).forEach(otherId => {
            if (otherId !== id && studentMap.has(otherId)) {
              siblings.set(otherId, `${otherId} - ${studentMap.get(otherId).student}`);
            }
          });
        }
      }
    }

    siblingMap.set(id, Array.from(siblings.values()).sort().join(", "));
  }

  return siblingMap;
}

/* ========================= MASTER TABLE ========================= */
/**
 * Compiles the Master Table.
 * Preserves "Discharged" students by comparing 
 * current raw data with previously generated Master Table data.
 */
function buildMasterTableData_(ss, studentMap, guardianRows, siblingMap) {
  const out = [[
    "OSIS",
    "Student",
    "Guardian",
    "Site",
    "Class",
    "Sibling",
    "Status"
  ]];
  
  const guardiansByStudent = new Map();
  for (const g of guardianRows) {
    if (!guardiansByStudent.has(g.student_id)) guardiansByStudent.set(g.student_id, []);
    guardiansByStudent.get(g.student_id).push(stripTrailingParenthetical_(g.guardianOutput));
  }
  
  const oldMasterMap = new Map();
  const oldMasterSheet = ss.getSheetByName("Master Table");
  
  if (oldMasterSheet) {
    const oldData = oldMasterSheet.getDataRange().getValues();
    
    let headerRowIdx = -1;
    let oldHeaderMap = {};
    for (let i = 0; i < Math.min(10, oldData.length); i++) {
       if (String(oldData[i][0]).trim().toUpperCase() === "OSIS") {
           headerRowIdx = i;
           for (let c = 0; c < oldData[i].length; c++) {
               oldHeaderMap[String(oldData[i][c]).trim()] = c;
           }
           break;
       }
    }
    
    const startRow = headerRowIdx !== -1 ? headerRowIdx + 1 : 4;
    const hStudent = oldHeaderMap["Student"] !== undefined ? oldHeaderMap["Student"] : 1;
    const hGuardian = oldHeaderMap["Guardian"] !== undefined ? oldHeaderMap["Guardian"] : 2;
    const hSite = oldHeaderMap["Site"] !== undefined ? oldHeaderMap["Site"] : -1;
    const hClass = oldHeaderMap["Class"] !== undefined ? oldHeaderMap["Class"] : 3;
    const hSibling = oldHeaderMap["Sibling"] !== undefined ? oldHeaderMap["Sibling"] : 4;
    const hStatus = oldHeaderMap["Status"] !== undefined ? oldHeaderMap["Status"] : 5;

    for (let i = startRow; i < oldData.length; i++) {
      const row = oldData[i];
      const osis = toText_(row[0]);
      
      if (osis.includes(",")) continue;
      if (!isRealStudentId_(osis)) continue;
      
      if (!oldMasterMap.has(osis)) {
        oldMasterMap.set(osis, {
          student: toText_(row[hStudent]),
          guardian: toText_(row[hGuardian]),
          site: hSite !== -1 ? toText_(row[hSite]) : "",
          classCode: toText_(row[hClass]),
          sibling: toText_(row[hSibling]),
          status: toText_(row[hStatus])
        });
      }
    }
  }
  
  const currentIds = new Set(Array.from(studentMap.keys()));
  const allIds = new Set([ ...Array.from(oldMasterMap.keys()), ...Array.from(studentMap.keys()) ]);
  const ids = Array.from(allIds).sort();
  
  for (const id of ids) {
    const currentStudent = studentMap.get(id);
    const oldStudent = oldMasterMap.get(id);
    const isCurrent = currentIds.has(id);
    
    let student = "";
    let guardian = "";
    let site = "";
    let classCode = "";
    let sibling = "";
    let status = isCurrent ? "Current" : "Discharged";

    if (isCurrent) {
      student = currentStudent.student;
      site = currentStudent.site || "";
      
      classCode = currentStudent.teachers && currentStudent.teachers.trim() !== "" 
          ? `${currentStudent.classCode}: ${currentStudent.teachers}` 
          : currentStudent.classCode;
          
      guardian = Array.from(new Set(guardiansByStudent.get(id) || []))
        .filter(Boolean)
        .join(" and ");
      sibling = siblingMap.get(id) || "";
      
      // 🚀 INJECTS THE CUSTOM LABEL INSTEAD OF FORCING "CURRENT"
      if (currentStudent.label && currentStudent.label.trim() !== "") {
         status = currentStudent.label;
      }
      
    } else if (oldStudent) {
      student = oldStudent.student;
      guardian = oldStudent.guardian;
      site = oldStudent.site || "";
      classCode = oldStudent.classCode;
      sibling = oldStudent.sibling;
      
      const rawStatus = String(oldStudent.status).trim();
      const statLower = rawStatus.toLowerCase();
      if (statLower !== "current" && statLower !== "master" && rawStatus !== "") {
          status = rawStatus;
      }
    }
    
    out.push([ id, student, guardian, site, classCode, sibling, status ]);
  }
  
  // ==========================================
  // --- UPGRADED: FAMILY ALIAS GENERATOR ---
  // ==========================================
  const familyMap = {};
  
  for (let i = 1; i < out.length; i++) {
    const row = out[i];
    const osis = String(row[0]).trim();
    const student = row[1];
    const guardian = row[2];
    const site = row[3];
    const classCode = row[4];
    const siblingString = row[5];
    const status = row[6];
    
    let siblingOsisArray = [];
    if (siblingString && String(siblingString).trim() !== "") {
       const matches = String(siblingString).match(/\b\d{6,}\b/g);
       if (matches) siblingOsisArray = matches;
    }
    
    if (siblingOsisArray.length > 0) {
      const fullFamily = [osis, ...siblingOsisArray].sort();
      const familyKey = fullFamily.join("_");
      
      if (!familyMap[familyKey]) familyMap[familyKey] = [];
      familyMap[familyKey].push({
         osis: osis,
         student: student,
         site: site,
         pClass: classCode,
         status: status,
         originalRow: row
      });
    }
  }
  
  for (const key in familyMap) {
    const kids = familyMap[key];
    if (kids.length > 1) {
      const combinedOsis = kids.map(k => k.osis).join(", ");
      const combinedStudent = kids.map(k => k.student).join(" & ");
      const combinedSite = Array.from(new Set(kids.map(k => k.site).filter(Boolean))).join(" / ");
      const combinedClass = kids.map(k => k.pClass).join(" / ");
      
      const isFamilyCurrent = kids.some(k => k.status === "Current");
      const familyStatus = isFamilyCurrent ? "Current" : "Discharged";
      
      let comboRow = [...kids[0].originalRow];
      comboRow[0] = combinedOsis;
      comboRow[1] = combinedStudent;
      comboRow[3] = combinedSite;
      comboRow[4] = combinedClass;
      comboRow[6] = familyStatus;
      
      out.push(comboRow);
    }
  }
  
  // ==========================================
  // --- ALPHABETIZE BY LAST NAME ---
  // ==========================================
  const headers = out[0];
  const body = out.slice(1);
  body.sort(byStudentLastName_(1));

  return [headers, ...body];
}

function buildBackendEventLogData_(ss) {
  const out = [["Date", "Event Name", "Category", "OSIS", "Student Name", "Guardian(s)", "Parent Attendee", "# of Indv", "Notes"]];
  const eventsSheet = ss.getSheetByName("Events");
  
  if (!eventsSheet) return out;

  const maxRows = eventsSheet.getLastRow();
  const maxCols = eventsSheet.getLastColumn();

  if (maxRows < 3 || maxCols < 6) return out;

  const data = eventsSheet.getRange(1, 1, maxRows, maxCols).getValues();
  const records = [];

  for (let col = 0; col < maxCols; col += 7) {
    const eventName = String(data[0][col] || "").trim();
    if (!eventName) continue;

    const category = String(data[0][col + 2] || "Other").trim();
    
    let dateStr = data[0][col + 4];
    let eventDate = new Date(dateStr);
    if (isNaN(eventDate)) eventDate = new Date();

    for (let row = 2; row < maxRows; row++) {
      const rawOsis = String(data[row][col] || "").trim();
      if (!rawOsis) continue;

      const guardian = String(data[row][col + 1] || "").trim();
      const attendee = String(data[row][col + 2] || "").trim();
      const rawStudentName = String(data[row][col + 3] || "").trim();
      
      // ⚡ THE ZERO FIX: Evaluates the literal raw cell value!
      let rawNumCell = data[row][col + 4];
      let numIndv = 1;
      if (String(rawNumCell).trim() !== "") {
          numIndv = Number(rawNumCell);
      }

      const notes = String(data[row][col + 5] || "").trim();

      const osisArray = rawOsis.split(',').map(o => o.trim()).filter(o => o !== "");
      const nameArray = rawStudentName.split(',').map(n => n.trim()).filter(n => n !== "");

      for (let i = 0; i < osisArray.length; i++) {
        const singleOsis = osisArray[i];
        const singleName = nameArray[i] !== undefined ? nameArray[i] : rawStudentName;
        const assignedNumIndv = (i === 0) ? numIndv : "";

        records.push([eventDate, eventName, category, singleOsis, singleName, guardian, attendee, assignedNumIndv, notes]);
      }
    }
  }

  records.sort((a, b) => {
    const dateDiff = b[0].getTime() - a[0].getTime();
    if (dateDiff !== 0) return dateDiff;
    return a[1].localeCompare(b[1]);
  });

  return out.concat(records);
}

/* =========================
   COMBINED CONTACT TRACKING
========================= */
function buildCombinedContactTrackingData_(ss, masterTable, contactLogData, eventLogData) {

  const props = PropertiesService.getDocumentProperties();
  const disableHiddenNotes = props.getProperty('disableHiddenNotes') === 'true';

  // --- Retrieve Custom Date Range Filter ---
  const notesStartDateStr = props.getProperty('notesStartDate');
  const notesEndDateStr = props.getProperty('notesEndDate');

  let filterStartDate = null;
  let filterEndDate = null;

  if (notesStartDateStr) {
    const parts = notesStartDateStr.split("-");
    if (parts.length === 3) filterStartDate = new Date(parts[0], parts[1] - 1, parts[2], 0, 0, 0);
  }
  if (notesEndDateStr) {
    const parts = notesEndDateStr.split("-");
    if (parts.length === 3) filterEndDate = new Date(parts[0], parts[1] - 1, parts[2], 23, 59, 59);
  }

  const out = [[ "OSIS", "Guardians", "Student", "Notes" ]];

  const startIndex = (masterTable.length > 0 && String(masterTable[0][0]).toUpperCase() === "OSIS") ? 1 : 0;

  // OSIS -> guardians/student. A sibling-group row ("111, 222" / "Ann X, Ben X") gives each
  // OSIS its own student name when the counts line up, so those rows aren't left blank.
  const masterMap = new Map();
  masterTable.slice(startIndex).forEach(r => {
    const ids = String(r[0]).split(",").map(o => o.trim()).filter(o => isRealStudentId_(o));
    const students = String(r[1]).split(",").map(n => n.trim());
    ids.forEach((id, i) => {
      if (masterMap.has(id) && ids.length > 1) return; // a student's own row wins over a group row
      masterMap.set(id, {
        guardians: toText_(r[2]),
        student: ids.length > 1 && students.length === ids.length ? students[i] : toText_(r[1])
      });
    });
  });

  const oldCombinedSheet = ss.getSheetByName("Combined Contact Tracking");
  const oldCombinedMap = new Map();
  if (oldCombinedSheet) {
    const oldData = oldCombinedSheet.getDataRange().getValues();
    for (let i = 1; i < oldData.length; i++) {
      const row = oldData[i];
      const osis = toText_(row[0]);
      if (!isRealStudentId_(osis)) continue;

      if (!oldCombinedMap.has(osis)) {
        oldCombinedMap.set(osis, {
          guardians: toText_(row[1]),
          student: toText_(row[2])
        });
      }
    }
  }

  // --- NEW: Retrieve Manual Notes (Matches Contact Log Array Logic) ---
  const notesSheet = ss.getSheetByName("Notes");
  const usableNotesRows = [];
  if (notesSheet) {
    const nData = notesSheet.getDataRange().getValues();
    for (let i = 1; i < nData.length; i++) {
      if (!nData[i]) continue;
      const rawOsis = toText_(nData[i][0]);
      if (rawOsis.trim() !== "") {
        usableNotesRows.push(nData[i]);
      }
    }
  }

  const usableContactRows = [];
  for (let i = 1; i < contactLogData.length; i++) {
    const row = contactLogData[i];
    const rowMarker = toText_(row[4]);
    if (rowMarker.toUpperCase() === "END") break;
    usableContactRows.push(row);
  }

  const usableEventRows = eventLogData && eventLogData.length > 1 ? eventLogData.slice(1) : [];

  const osisSet = new Set();

  // Tokenize into whole OSIS numbers so siblings formatted as "222222222, 333333333"
  // (or with & / separators) each register independently.
  const addOsisToSet = (osisString) => {
    extractOsisTokens_(osisString).forEach(cleanOsis => {
      if (isRealStudentId_(cleanOsis)) osisSet.add(cleanOsis);
    });
  };

  usableContactRows.forEach(r => addOsisToSet(r[4]));
  usableEventRows.forEach(r => addOsisToSet(r[3]));
  usableNotesRows.forEach(r => addOsisToSet(r[0]));

  // Each source's rows grouped by OSIS (sheet order kept), so every student only
  // looks at their own rows instead of re-scanning every row of every source
  const contactRowsByOsis = rowsByOsis_(usableContactRows, 4);
  const eventRowsByOsis = rowsByOsis_(usableEventRows, 3);
  const notesRowsByOsis = rowsByOsis_(usableNotesRows, 0);

  for (const osis of osisSet) {
    const combinedInteractions = [];

    // --- Process Contact Log ---
    for (const row of contactRowsByOsis.get(osis) || []) {
      const entry = contactLogNoteLine_(row, disableHiddenNotes);
      if (!entry) continue;
      if (filterStartDate && entry.dateObj < filterStartDate) continue;
      if (filterEndDate && entry.dateObj > filterEndDate) continue;
      combinedInteractions.push(entry);
    }

    // --- Process Event Log ---
    for (const row of eventRowsByOsis.get(osis) || []) {
      const entry = eventNoteLine_(row);
      if (filterStartDate && entry.dateObj < filterStartDate) continue;
      if (filterEndDate && entry.dateObj > filterEndDate) continue;
      combinedInteractions.push(entry);
    }

    // --- Process Manual Notes ---
    let mergedManualNotes = "";
    for (const nRow of notesRowsByOsis.get(osis) || []) {
      const nText = toText_(nRow[3]);
      if (nText.trim() !== "") {
        mergedManualNotes = mergedManualNotes ? mergedManualNotes + "\n\n" + nText.trim() : nText.trim();
      }
    }

    const mergedNotes = joinNoteLines_(mergedManualNotes, combinedInteractions);

    if (mergedNotes !== "") {
      const master = masterMap.get(osis);
      const old = oldCombinedMap.get(osis);

      const guardians = master?.guardians || old?.guardians || "";
      const student = master?.student || old?.student || "";

      out.push([osis, guardians, student, mergedNotes.trim()]);
    }
  }

  const headers = out[0];
  const body = out.slice(1);
  body.sort(byStudentLastName_(2));

  return [headers, ...body];
}

/** Rows grouped by every OSIS listed in their `col` cell (a sibling row lands under each OSIS), sheet order kept. */
function rowsByOsis_(rows, col) {
  const map = new Map();
  rows.forEach(row => {
    new Set(extractOsisTokens_(row[col])).forEach(osis => {
      if (!map.has(osis)) map.set(osis, []);
      map.get(osis).push(row);
    });
  });
  return map;
}

/* =========================
   NOTES LINES
   One place that turns Contact Log / Event Log rows into the notes lines
   "𝑫𝒂𝒕𝒆 - {𝐌𝐞𝐭𝐡𝐨𝐝 - 𝐏𝐞𝐫𝐬𝐨𝐧} note [follow-up]", used by both the full build
   (buildCombinedContactTrackingData_) and the quick update (updateOsisDeltas_),
   so the two always write the same text.
========================= */

/**
 * A Contact Log row (B date, D person, J/K method, M notes, O follow-up, P hide)
 * as { dateObj, text }, or null when the row is hidden (unless hidden notes are
 * included) or has neither notes nor follow-up notes.
 */
function contactLogNoteLine_(row, disableHiddenNotes) {
  const isHidden = row[15];
  const hiddenBool = (isHidden === true || String(isHidden).toUpperCase() === "TRUE");
  if (!disableHiddenNotes && hiddenBool) return null;

  let dateVal = row[1];
  if (!dateVal) dateVal = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");

  let dateObj = new Date(dateVal);
  if (isNaN(dateObj)) dateObj = new Date(0);

  const eVal = toText_(row[3]);  // Person Spoke With
  const mVal = toText_(row[9]);  // Method 1
  const nVal = toText_(row[10]); // Method 2
  const pVal = toText_(row[12]); // Notes
  const rVal = toText_(row[14]); // Followup Notes

  if (pVal === "" && rVal === "") return null;

  const methodLabel = nVal !== "" ? nVal : mVal;
  const dateText = formatDateFull_(dateVal);

  // Only append the "- Person Spoke With" when Col D was actually filled;
  // a blank D should read "{Method}", not "{Method - }".
  const spokeWith = String(eVal).trim() ? ` - ${eVal}` : "";
  let line = `${noteLineHeader_(dateText, methodLabel + spokeWith)} ${pVal}`;
  if (rVal !== "") line += ` [${rVal}]`;

  return { dateObj: dateObj, text: line.trim() };
}

/** A Backend Event Log row (A date, B event, F guardian, G attendee, I notes) as { dateObj, text }. */
function eventNoteLine_(row) {
  const dateVal = row[0];

  let dateObj = new Date(dateVal);
  if (isNaN(dateObj)) dateObj = new Date(0);

  const eventName = toText_(row[1]);
  const guardian = toText_(row[5]);
  const attendee = toText_(row[6]);
  const notes = toText_(row[8]);

  let person = attendee ? attendee : guardian;
  if (!person) person = "Parent";

  const dateText = formatDateFull_(dateVal);
  let line = `${noteLineHeader_(dateText, `Event - [${person}]`)} |${eventName}|`;
  if (notes) line += ` ${notes}`;

  return { dateObj: dateObj, text: line.trim() };
}

/** Manual notes ("𝑵𝒐𝒕𝒆𝒔: ...") first, then the dated lines oldest to newest. */
function joinNoteLines_(manualNotesText, interactions) {
  let baseNotes = "";
  if (interactions.length > 0) {
    interactions.sort((a, b) => a.dateObj.getTime() - b.dateObj.getTime());
    baseNotes = interactions.map(interaction => interaction.text).join("\n");
  }

  const manualNotes = manualNotesText !== "" ? `𝑵𝒐𝒕𝒆𝒔: ${manualNotesText}` : "";

  if (manualNotes !== "" && baseNotes !== "") return manualNotes + "\n\n" + baseNotes;
  if (manualNotes !== "") return manualNotes;
  return baseNotes;
}

/**
 * Sort comparator for rows with a student name in `col`: the first student's
 * last name (a sibling row "Ann X & Ben Y" uses Ann's), then the full name.
 * Used by the Master Table and Combined Contact Tracking.
 */
function byStudentLastName_(col) {
  const extractLastName = (name) => {
    const firstStudent = String(name || "").split("&")[0].trim();
    const parts = firstStudent.split(/\s+/);
    return parts.length > 1 ? parts.pop().toLowerCase() : firstStudent.toLowerCase();
  };
  return (a, b) => {
    const lastNameCompare = compareNatural_(extractLastName(a[col]), extractLastName(b[col]));
    if (lastNameCompare === 0) {
      return compareNatural_(String(a[col] || ""), String(b[col] || ""));
    }
    return lastNameCompare;
  };
}

/* =========================
   NAME MATCHING (shared by Parent Square + NYCSA)
   Names are compared as sets of lowercase parts, so accents, punctuation,
   commas and word order never block a match ("SIMPSON, Homer" = "Homer Simpson").
========================= */
function nameTokens_(text) {
  const s = stripTrailingParenthetical_(text)
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .toLowerCase()
    .replace(/['’`.]/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
  return s ? s.split(" ") : [];
}

function sameNameExactly_(a, b) {
  const A = new Set(a), B = new Set(b);
  return A.size > 0 && A.size === B.size && [...A].every(t => B.has(t));
}

// Same name, or the shorter name (2+ parts) is fully inside the longer one:
// ATS "Maria Del Carmen Lopez" matches NYCSA "Maria" + "Lopez".
function namesMatch_(a, b) {
  const A = new Set(a), B = new Set(b);
  if (!A.size || !B.size) return false;
  const small = A.size <= B.size ? A : B;
  const big = small === A ? B : A;
  for (const t of small) if (!big.has(t)) return false;
  return small.size >= 2 || A.size === B.size;
}

function shareNameToken_(a, b) {
  const B = new Set(b);
  return a.some(t => B.has(t));
}

// Digits only, padded to 9 so an OSIS stored as a number (lost leading zero) still matches
function normalizeOsis_(value) {
  const digits = toText_(value).replace(/\D/g, "");
  return digits && digits.length < 9 ? digits.padStart(9, "0") : digits;
}

/* =========================
   PARENT SQUARE INDEX
========================= */
function buildParentSquareIndex_(rawData) {
  const emailStatus = new Map();
  const phoneStatus = new Map();
  const rowsByStudentToken = new Map(); // student-name part -> PS rows listing that student
  const studentTokensByOsis = new Map(); // OSIS -> [first name, last name] from ATS

  for (let i = 0; i < rawData.length; i++) {
    const r = rawData[i];

    if (isRealStudentId_(r[2])) {
      const studentTokens = [nameTokens_(r[4])[0], nameTokens_(r[3])[0]].filter(Boolean);
      studentTokensByOsis.set(normalizeOsis_(r[2]), studentTokens);
    }

    const email = normalizeEmail_(r[34]);
    const phones = [normalizePhoneForMatch_(r[35]), normalizePhoneForMatch_(r[36])].filter(Boolean);
    const registered = normalizeYesNo_(r[38]);

    if (email) mergeStatusIntoMap_(emailStatus, email, registered);
    phones.forEach(p => mergeStatusIntoMap_(phoneStatus, p, registered));

    const students = new Set(nameTokens_(r[33]));
    if (!students.size) continue;
    const row = { name: nameTokens_(r[32]), students, email, phones, registered };
    students.forEach(t => {
      if (!rowsByStudentToken.has(t)) rowsByStudentToken.set(t, []);
      rowsByStudentToken.get(t).push(row);
    });
  }

  return { emailStatus, phoneStatus, rowsByStudentToken, studentTokensByOsis };
}

function mergeStatusIntoMap_(map, key, newStatus) {
  if (!key) return;
  const current = map.get(key) || "";
  map.set(key, mergeYesNoStatus_(current, newStatus));
}

function mergeYesNoStatus_(a, b) {
  const aa = normalizeYesNo_(a);
  const bb = normalizeYesNo_(b);

  if (aa === "Yes" || bb === "Yes") return "Yes";
  if (aa === "No" || bb === "No") return "No";
  return "";
}

function normalizeYesNo_(value) {
  const v = toText_(value).toLowerCase();
  if (v === "yes") return "Yes";
  if (v === "no") return "No";
  return "";
}

// PS rows whose "Students" column names this guardian's student (first + last name)
function psRowsForStudent_(studentId, parentSquareIndex) {
  const tokens = parentSquareIndex.studentTokensByOsis.get(normalizeOsis_(studentId)) || [];
  if (tokens.length < 2) return [];
  const candidates = parentSquareIndex.rowsByStudentToken.get(tokens[1]) || [];
  return candidates.filter(r => tokens.every(t => r.students.has(t)));
}

function getParentsSquareStatus_(guardian, parentSquareIndex) {
  const email = normalizeEmail_(guardian.email);
  const phone = normalizePhoneForMatch_(guardian.mobile);
  const reduceStatus = rows => rows.reduce((s, r) => mergeYesNoStatus_(s, r.registered), "");

  // 1. Rows for this guardian's own student: name + contact, then name, then contact.
  //    Keeps a shared household phone/email from giving one parent the other's status.
  const studentRows = psRowsForStudent_(guardian.student_id, parentSquareIndex);
  if (studentRows.length) {
    const name = nameTokens_(guardian.guardianOutput);
    const byName = studentRows.filter(r => namesMatch_(name, r.name));
    const byContact = studentRows.filter(r => (email && r.email === email) || (phone && r.phones.indexOf(phone) !== -1));
    const both = byContact.filter(r => byName.indexOf(r) !== -1);
    const pick = both.length ? both : byName.length ? byName : byContact;
    if (pick.length) return reduceStatus(pick);
  }

  // 2. Email/phone anywhere in Parent Square
  let status = "";

  if (email && parentSquareIndex.emailStatus.has(email)) {
    status = mergeYesNoStatus_(status, parentSquareIndex.emailStatus.get(email));
  }

  if (phone && parentSquareIndex.phoneStatus.has(phone)) {
    status = mergeYesNoStatus_(status, parentSquareIndex.phoneStatus.get(phone));
  }

  return status;
}

/* =========================
   NYCSA INDEX (UPGRADED)
========================= */
function buildNYCSAIndex_(rawData) {
  const byStudent = new Map(); // OSIS -> NYCSA guardian rows for that student

  for (let i = 0; i < rawData.length; i++) {
    const r = rawData[i];

    const osis = normalizeOsis_(r[46]); // NYCSA Student ID is index 46
    const name = nameTokens_(`${toText_(r[49])} ${toText_(r[50])}`);
    if (!osis || !name.length) continue;

    if (!byStudent.has(osis)) byStudent.set(osis, []);
    byStudent.get(osis).push({
      name,
      phones: [r[52], r[53], r[54]].map(normalizePhoneForMatch_).filter(Boolean), // Home, Work, Cell
      anotherCustodial: normalizeYesNo_(r[51]),
      accountCode: toText_(r[57]) // NYCSA Account Creation Code (Column BF)
    });
  }

  return { byStudent };
}

// Within one student's NYCSA rows: exact name, then phone (with a shared name part),
// then partial name. Each later step must point to a single row, and never to a row
// whose name is exactly the other guardian's.
function findNYCSARow_(g, rows, otherGuardian) {
  if (!g || !g.guardianOutput || !rows.length) return null;
  const name = nameTokens_(g.guardianOutput);

  const exact = rows.filter(r => sameNameExactly_(name, r.name));
  if (exact.length) return exact[0];

  const otherName = otherGuardian && otherGuardian.guardianOutput ? nameTokens_(otherGuardian.guardianOutput) : [];
  const open = rows.filter(r => !sameNameExactly_(otherName, r.name));

  const phone = normalizePhoneForMatch_(g.mobile);
  const byPhone = phone ? open.filter(r => r.phones.indexOf(phone) !== -1 && shareNameToken_(name, r.name)) : [];
  if (byPhone.length === 1) return byPhone[0];

  const byName = open.filter(r => namesMatch_(name, r.name));
  if (byName.length === 1) return byName[0];

  return null;
}

function getNYCSAStatus_(currentGuardian, otherGuardian, index) {
  const rows = index.byStudent.get(normalizeOsis_(currentGuardian.student_id)) || [];

  const currentMatch = findNYCSARow_(currentGuardian, rows, otherGuardian);
  const otherMatch = findNYCSARow_(otherGuardian, rows, currentGuardian);

  if (currentMatch) {
    const code = currentMatch.accountCode ? ` ${currentMatch.accountCode}` : "";
    const type = currentMatch.anotherCustodial === "Yes" ? "OG" : "NG";
    return `No (${type})${code}`;
  }

  if (otherMatch) {
    if (otherMatch.anotherCustodial === "No") return "-";
    if (otherMatch.anotherCustodial === "Yes") return "Yes/Unknown";
  }

  return "Yes/Unknown";
}

/* =========================
   PARENTS DIVIDED
========================= */
function buildParentsDividedData_(ss, guardianRows, siblingMap, parentSquareIndex, nycsaIndex, studentMap, contactLogData, eventLogData) {
  const out = [[
    "student_id",
    "Guardian output",
    "Student",
    "Class / Site",
    "relationship",
    "language",
    "email",
    "mobile",
    "Sibling",
    "Parents Square",
    "NYCSA",
    "Last Interaction"
  ]];

  const guardiansByStudentId = new Map();
  for (const g of guardianRows) {
    if (!guardiansByStudentId.has(g.student_id)) guardiansByStudentId.set(g.student_id, []);
    guardiansByStudentId.get(g.student_id).push(g);
  }

  const allInteractions = [];

  for (let i = 1; i < contactLogData.length; i++) {
    const row = contactLogData[i];
    
    // ✨ FIX: Contact Log shifted columns! Read the correct indexes:
    // OSIS is now index 4 (Col E), Hidden Checkbox is 15 (Col P), Person Spoke With is 3 (Col D)
    const rowMarker = toText_(row[4]);
    if (rowMarker.toUpperCase() === "END") break;
    
    const isHidden = row[15]; 
    if (isHidden === true || String(isHidden).toLowerCase() === "true") continue;
    
    const dateVal = row[1];
    const personStr = String(row[3] || row[2] || "").toLowerCase().replace("▶", "").trim();
    const osisStr = String(row[4] || "").trim();
    
    if (dateVal && osisStr) {
       let d = new Date(dateVal);
       if (!isNaN(d)) allInteractions.push({ osisStr, personStr, dateObj: d, dateRaw: dateVal });
    }
  }

  for (let i = 1; i < eventLogData.length; i++) {
    const row = eventLogData[i];
    const dateVal = row[0];
    const osisStr = String(row[3] || "").trim();
    const personStr = String(row[5] || row[4] || "").toLowerCase().trim();
    
    if (dateVal && osisStr) {
       let d = new Date(dateVal);
       if (!isNaN(d)) allInteractions.push({ osisStr, personStr, dateObj: d, dateRaw: dateVal });
    }
  }

  allInteractions.sort((a, b) => a.dateObj.getTime() - b.dateObj.getTime());

  const interactionsByOsis = new Map();
  for (const interaction of allInteractions) {
    // ✨ UPGRADED: Smart number extraction handles siblings even if you separated them with "&" or "/"
    const matches = interaction.osisStr.match(/\b\d{6,}\b/g) || interaction.osisStr.split(",").map(p => p.trim()).filter(Boolean);
    for (const osis of matches) {
      if (!interactionsByOsis.has(osis)) interactionsByOsis.set(osis, []);
      interactionsByOsis.get(osis).push(interaction);
    }
  }

  for (const g of guardianRows) {
    const student = studentMap.get(g.student_id) || { student: "", classCode: "", site: "" };

    let classSiteStr = String(student.classCode || "");
    if (student.site && String(student.site).trim() !== "") {
      classSiteStr += classSiteStr ? ` (Site: ${student.site})` : `Site: ${student.site}`;
    }

    let lastInteraction = "None logged";
    const cleanGuardianName = stripTrailingParenthetical_(g.guardianOutput).toLowerCase().trim();
    const gFirstName = cleanGuardianName.split(" ")[0];
    const targetOsis = String(g.student_id).trim();

    const guardianInteractions = interactionsByOsis.get(targetOsis) || [];
    for (let i = guardianInteractions.length - 1; i >= 0; i--) {
      const interaction = guardianInteractions[i];
      
      if (interaction.personStr) {
        const logName = interaction.personStr;
        const logFirstName = logName.split(" ")[0];

        if (cleanGuardianName.includes(logName) || logName.includes(cleanGuardianName) || logFirstName === gFirstName) {
          lastInteraction = formatDateFull_(interaction.dateRaw);
          break;
        }
      }
    }

    const otherGuardian = (guardiansByStudentId.get(g.student_id) || []).find(row => row !== g) || null;

    out.push([
      g.student_id,
      stripTrailingParenthetical_(g.guardianOutput), 
      student.student,
      classSiteStr,
      g.relationship,
      g.language,
      g.email,
      g.mobile,
      siblingMap.get(g.student_id) || "",
      getParentsSquareStatus_(g, parentSquareIndex),
      getNYCSAStatus_(g, otherGuardian, nycsaIndex),
      lastInteraction
    ]);
  }

  const headers = out[0];
  const body = out.slice(1);

  body.sort((a, b) => {
    const extractLastName = (name) => {
      const parts = String(name || "").trim().split(/\s+/);
      return parts.length > 1 ? parts.pop().toLowerCase() : String(name || "").toLowerCase();
    };
    
    const lastNameA = extractLastName(a[2]);
    const lastNameB = extractLastName(b[2]);
    
    const lastNameCompare = compareNatural_(lastNameA, lastNameB);
    
    if (lastNameCompare === 0) {
      const studentCompare = compareNatural_(String(a[2] || ""), String(b[2] || ""));
      if (studentCompare === 0) {
        return compareNatural_(String(a[1] || ""), String(b[1] || ""));
      }
      return studentCompare;
    }
    return lastNameCompare;
  });

  return [headers, ...body];
}

/* =========================
   DIRECTORY
========================= */
function buildDirectoryData_(studentMap, guardianRows) {
  const headers = ["Last Name", "First Name", "OSIS", "Site", "Class/ Teacher", "Parent Contact"];
  const out = [headers];
  const grouped = groupGuardiansByStudent_(guardianRows);

  for (const [id, student] of studentMap) {
    const guardians = grouped.get(id) || [];
    const parentContact = guardians.map(formatGuardianDirectoryLine_).filter(Boolean).join("\n");

    const parsedStudent = splitFullName_(student.student);
    const classTeacherInfo = student.teachers 
      ? `${student.classCode}: ${student.teachers}` 
      : student.classCode;

    const row = [parsedStudent.last, parsedStudent.first, id, student.site || "", classTeacherInfo, parentContact];
    out.push(row);
  }

  const body = out.slice(1).sort(compareDirectoryRows_);

  return [out[0], ...body];
}

/** Directory order: Site, then Class / Teacher, then last name. */
function compareDirectoryRows_(a, b) {
  const siteA = String(a[3] || "").trim();
  const siteB = String(b[3] || "").trim();
  const primarySort = compareNatural_(siteA, siteB);
  if (primarySort !== 0) return primarySort;

  const classA = String(a[4] || "").trim();
  const classB = String(b[4] || "").trim();
  const secondarySort = compareNatural_(classA, classB);
  if (secondarySort !== 0) return secondarySort;

  const lastNameA = String(a[0] || "").trim();
  const lastNameB = String(b[0] || "").trim();
  return compareNatural_(lastNameA, lastNameB);
}

function formatDirectoryAuto_(sheet) {
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;

  const siteValues = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
  const hasSiteData = siteValues.some(r => String(r[0]).trim() !== "");

  const targetCol = hasSiteData ? 4 : 5; 
  const targetValues = sheet.getRange(2, targetCol, lastRow - 1, 1).getValues();

  const colors = [
    "#FFB3BA", "#BAE1FF", "#BAFFC9", "#FFFFBA", "#CBAACB", "#FFDFBA", 
    "#A9DEF9", "#D0F0C0", "#FFFACD", "#E2B2E8", "#FFDAB3", "#F6D2E6",
    "#B6CCFE", "#FCF6BD", "#FF99C8", "#D7E3FC", "#D0F4DE", "#FFE5D9"
  ];

  const map = new Map();
  let idx = 0;
  const columnAColors = [];

  for (let i = 0; i < targetValues.length; i++) {
    const val = String(targetValues[i][0] || "").trim();
    if (!val) {
      columnAColors.push([null]); 
      continue;
    }
    if (!map.has(val)) {
      map.set(val, colors[idx % colors.length]);
      idx++;
    }
    columnAColors.push([map.get(val)]);
  }

  sheet.getRange(2, 1, lastRow - 1, sheet.getMaxColumns()).setBackground(null);
  sheet.getRange(2, 1, columnAColors.length, 1).setBackgrounds(columnAColors);
  sheet.autoResizeColumns(1, sheet.getLastColumn());
}

function formatGuardianDirectoryLine_(g) {
  const left = g.guardianOutput;
  const parts = [toText_(g.email), toText_(g.mobile)].filter(Boolean);
  if (!left) return "";
  if (parts.length === 0) return left;
  return `${left} - ${parts.join(", ")}`;
}

/* =========================
   PHONE CONTACTS
========================= */
function buildPhoneContactsData_(ss, studentMap, guardianRows, masterTable, combinedContactTrackingArray, parentSquareIndex, nycsaIndex) {
  // The owner's sheet uses its own "tall" layout (one row per phone, plus Old Phone Contacts)
  const owner = getOwnerHooks_();
  if (owner) return owner.buildPhoneContactsData(ss, studentMap, guardianRows, masterTable, combinedContactTrackingArray, parentSquareIndex, nycsaIndex);

  const out = [[
    "Name Prefix", "First Name", "Middle Name", "Last Name", "Name Suffix", 
    "Phonetic First Name", "Phonetic Middle Name", "Phonetic Last Name", "Nickname", "File As", 
    "E-mail 1 - Label", "E-mail 1 - Value", "E-mail 2 - Label", "E-mail 2 - Value", "E-mail 3 - Label", 
    "E-mail 3 - Value", "Phone 1 - Type", "Phone 1 - Value", "Phone 2 - Type", "Phone 2 - Value", 
    "Phone 3 - Type", "Phone 3 - Value", "Organization Name", "Organization Department", "Birthday", 
    "Event 1 - Label", "Event 1 - Value", "Relation 1 - Label", "Relation 1 - Value", "Website 1 - Label", 
    "Website 1 - Value", "Custom Field 1 - Label", "Custom Field 1 - Value", "Notes", "Labels", "OSIS"
  ]];

  const masterMap = new Map(
    masterTable.slice(1).filter(r => isRealStudentId_(r[0])).map(r => [toText_(r[0]), r])
  );
  
  const tracking = buildTrackingNotesMap_(ss, combinedContactTrackingArray); 

  const guardiansByStudentId = new Map();
  for (const g of guardianRows) {
    if (!guardiansByStudentId.has(g.student_id)) guardiansByStudentId.set(g.student_id, []);
    guardiansByStudentId.get(g.student_id).push(g);
  }

  const groupedGuardians = new Map();

  for (const g of guardianRows) {
    const id = g.student_id;
    const master = masterMap.get(id) || [];
    const student = studentMap.get(id) || { student: "", classCode: "", site: "", teachers: "" };
    const guardianName = stripTrailingParenthetical_(g.guardianOutput);

    const key = `${guardianName.toLowerCase()}|${toText_(g.email).toLowerCase()}|${toText_(g.mobile)}`;

    if (!groupedGuardians.has(key)) {
      groupedGuardians.set(key, {
        guardianName: guardianName,
        relationship: toText_(g.relationship),
        email: toText_(g.email),
        phone: toText_(g.mobile),
        students: []
      });
    }

    const group = groupedGuardians.get(key);

    let classDisplay = toText_(student.classCode);
    if (student.teachers && toText_(student.teachers).trim() !== "") {
      classDisplay += ` - ${toText_(student.teachers)}`;
    } else if (master[4] && toText_(master[4]).trim() !== "") { 
      classDisplay = toText_(master[4]); 
    }

    const otherGuardian = (guardiansByStudentId.get(id) || []).find(row => row !== g) || null;

    const psStatus = getParentsSquareStatus_(g, parentSquareIndex);
    const nycsaStatus = getNYCSAStatus_(g, otherGuardian, nycsaIndex); 

    let studentBlock = `STUDENT NAME: ${toText_(student.student)}\n` +
                       `RELATIONSHIP: ${toText_(g.relationship)}\n` +
                       `LANGUAGE: ${toText_(g.language)}\n` +
                       `OSIS: ${id}\n` +
                       `Parent Name: ${toText_(master[2] || guardianName)}\n` +
                       `Class: ${classDisplay}`;
    
    if (student.site && toText_(student.site) !== "") studentBlock += `\nSite: ${toText_(student.site)}`;
    
    studentBlock += `\nParents Square: ${psStatus || "No"}`;
    studentBlock += `\nNYCSA: ${nycsaStatus || "No"}`;

    const studentNotes = tracking.get(id);
    const cleanedNotes = (studentNotes && studentNotes.trim() !== "") ? studentNotes.trim() : "No Notes";
    
    studentBlock += `\n\nCommunications:\n${cleanedNotes}`;

    group.students.push({
      studentName: toText_(student.student),
      osis: id,
      block: studentBlock,
      site: toText_(student.site),
      classCode: classDisplay
    });
  }

  for (const group of groupedGuardians.values()) {
    const studentNames = group.students.map(s => s.studentName).join(", ");
    const osisList = group.students.map(s => s.osis).join(", ");
    const finalNotes = group.students.map(s => s.block).join("\n\n---\n");

    const hasSiblings = group.students.length > 1;

    const phoneLabels = group.students.map(s => {
      let label = "";
      if (hasSiblings) {
        const firstName = String(s.studentName).split(" ")[0];
        label += `{${firstName}} `;
      }
      if (s.site && s.site !== "") label += `Site: ${s.site}, `;
      label += `Class: ${s.classCode}`;
      return label.trim();
    }).join(" & ");

    const orgNames = group.students.map(s => {
      if (hasSiblings) {
        const firstName = String(s.studentName).split(" ")[0];
        return `{${firstName}} OSIS: ${s.osis}`;
      } else {
        return `OSIS: ${s.osis}`;
      }
    }).join(", ");

    const row = new Array(36).fill("");
    
    row[1]  = `(${group.relationship}) ${group.guardianName}`.trim();
    row[3]  = `| Student: ${studentNames} |`;
    row[11] = group.email;
    row[16] = phoneLabels;
    row[17] = group.phone;
    row[22] = orgNames;
    row[33] = finalNotes;
    row[34] = "School Directory";
    row[35] = osisList;

    out.push(row);
  }
  
  return out;
}

function buildTrackingNotesMap_(ss, combinedContactTrackingArray) { 
  const map = new Map();
  
  if (combinedContactTrackingArray && combinedContactTrackingArray.length > 0) {
    for (let i = 1; i < combinedContactTrackingArray.length; i++) {
      const osis = toText_(combinedContactTrackingArray[i][0]);
      if (!isRealStudentId_(osis)) continue;
      if (!map.has(osis)) {
        map.set(osis, toText_(combinedContactTrackingArray[i][3])); 
      }
    }
    return map;
  }

  const sheet = ss.getSheetByName("Combined Contact Tracking");
  if (!sheet) return map;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const osis = toText_(data[i][0]);
    if (!isRealStudentId_(osis)) continue;
    if (!map.has(osis)) {
      map.set(osis, toText_(data[i][3]));
    }
  }
  return map;
}

// ======================================================================
// SECTION: NEW STUDENTS TAB (was NewStudents.gs)
// Keeps a running "New Students" log. Every build (including the
// nightly sync) compares the fresh Master Table with the one it is about to
// replace and records students who just appeared in Raw Data (ATS) or came
// back from Discharged. Only Raw Data counts; Source is always "ATS".
// The tab is created once and only ever added to (a row is re-dated when its
// student re-enters).
// ======================================================================
const NEW_STUDENTS_SHEET = "New Students";

/**
 * "75M094" -> "P94M" (D75 school naming). Falls back to "" when the DBN
 * doesn't look like a D75 code.
 */
function schoolShortName_(rawData) {
  for (let i = 0; i < rawData.length; i++) {
    const m = toText_(rawData[i][1]).toUpperCase().match(/^75([MXKQR])(\d{3})$/);
    if (m) return `P${Number(m[2])}${m[1]}`;
  }
  return "";
}

/**
 * Reads the Status of every student in the current Master Table (before it is
 * rewritten). Status is found by header, so column shifts don't break it.
 */
function readMasterStatusSnapshot_(ss) {
  const map = new Map();
  const sheet = ss.getSheetByName("Master Table");
  if (!sheet) return map;
  const data = sheet.getDataRange().getValues();
  const headerIdx = data.findIndex(r => String(r[0]).trim().toUpperCase() === "OSIS");
  if (headerIdx === -1) return map;
  let statusIdx = data[headerIdx].indexOf("Status");
  if (statusIdx === -1) statusIdx = 6;
  for (let i = headerIdx + 1; i < data.length; i++) {
    const osis = String(data[i][0]).trim();
    if (osis && !osis.includes(",")) map.set(osis, String(data[i][statusIdx]).trim());
  }
  return map;
}

const ATS_SNAPSHOT_PROP = "NEW_STUDENTS_ATS_OSIS";

const ATS_SNAPSHOT_CHUNK = 8000; // Document properties cap each value at ~9KB

/**
 * OSIS that were in Raw Data at the previous build, or null when there is no
 * snapshot for this spreadsheet yet (first run, or a copied sheet's leftovers).
 */
function loadAtsSnapshot_(ss) {
  const props = PropertiesService.getDocumentProperties();
  if (props.getProperty(ATS_SNAPSHOT_PROP + "_SHEET") !== ss.getId()) return null;
  const count = Number(props.getProperty(ATS_SNAPSHOT_PROP + "_N")) || 0;
  let text = "";
  for (let i = 0; i < count; i++) text += props.getProperty(ATS_SNAPSHOT_PROP + "_" + i) || "";
  return new Set(text.split(",").filter(Boolean));
}

function saveAtsSnapshot_(ss, osisSet) {
  const props = PropertiesService.getDocumentProperties();
  const text = Array.from(osisSet).join(",");
  const oldCount = Number(props.getProperty(ATS_SNAPSHOT_PROP + "_N")) || 0;
  const values = {};
  let n = 0;
  for (let i = 0; i < text.length; i += ATS_SNAPSHOT_CHUNK) values[ATS_SNAPSHOT_PROP + "_" + n++] = text.slice(i, i + ATS_SNAPSHOT_CHUNK);
  values[ATS_SNAPSHOT_PROP + "_N"] = String(n);
  values[ATS_SNAPSHOT_PROP + "_SHEET"] = ss.getId();
  props.setProperties(values);
  for (let i = n; i < oldCount; i++) props.deleteProperty(ATS_SNAPSHOT_PROP + "_" + i);
}

/**
 * Compares Raw Data (ATS, A:AE) with the previous build. Shared by the New
 * Students tab and the build report (sidebar / email) so they always agree.
 *   newOsis:        in Raw Data now but not at the last build, or back from Discharged
 *   dischargedOsis: "Current" in the old Master Table, in Raw Data at the last
 *                   build, and no longer in Raw Data
 * Call saveAtsSnapshot_(ss, changes.atsOsis) once the build has used them.
 *
 * @param {Array[]} masterTable        Fresh Master Table (header row first)
 * @param {Map}     oldMasterStatusMap OSIS -> Status from the Master Table before this build
 * @param {Array[]} rawData            Raw Data values
 */
function detectAtsChanges_(ss, masterTable, oldMasterStatusMap, rawData) {
  const atsOsis = new Set();
  for (let i = 0; i < rawData.length; i++) {
    const id = toText_(rawData[i][2]);
    if (isRealStudentId_(id)) atsOsis.add(id);
  }
  const newOsis = new Set();
  const dischargedOsis = new Set();

  // Previous Raw Data roster. Without a snapshot, fall back to the old Master Table.
  const prevAts = loadAtsSnapshot_(ss);
  // First build has nothing to compare against, so it would list every student
  if (!prevAts && oldMasterStatusMap.size === 0) return { atsOsis, newOsis, dischargedOsis };

  for (let i = 1; i < masterTable.length; i++) {
    const m = masterTable[i];
    const osis = String(m[0]).trim();
    if (!osis || osis.includes(",")) continue;
    const oldStatus = String(oldMasterStatusMap.has(osis) ? oldMasterStatusMap.get(osis) : "").trim();

    if (atsOsis.has(osis)) {
      if (/discharge/i.test(String(m[6]))) continue;
      const hadBefore = prevAts ? prevAts.has(osis) : oldMasterStatusMap.has(osis);
      if (!hadBefore || /discharge/i.test(oldStatus)) newOsis.add(osis);
    } else if (/^current$/i.test(oldStatus)) {
      // Without a snapshot, only count students the Master Table itself now marks Discharged
      const wasInAts = prevAts ? prevAts.has(osis) : /discharge/i.test(String(m[6]));
      if (wasInAts) dischargedOsis.add(osis);
    }
  }
  return { atsOsis, newOsis, dischargedOsis };
}

/**
 * Logs the students detectAtsChanges_ found new to Raw Data. Only Raw Data
 * counts: New/Edit Student and other Directory sources are ignored, and Source
 * is always "ATS".
 *
 * @param {Array[]} masterTable  Fresh Master Table (header row first)
 * @param {Object}  changes      Result of detectAtsChanges_
 * @param {Array[]} rawData      Raw Data values (A:AE)
 * @param {Array}   guardianRows Guardian objects (for Language)
 */
function logNewStudents_(ss, masterTable, changes, rawData, guardianRows) {
  let sheet = ss.getSheetByName(NEW_STUDENTS_SHEET);
  if (!sheet) sheet = ss.insertSheet(NEW_STUDENTS_SHEET);
  if (!toText_(sheet.getRange(1, 1).getValue())) {
    const school = schoolShortName_(rawData);
    const headers = [school ? `Date Entered ${school}` : "Date Entered", "Source", "OSIS", "Student", "Guardian", "Site", "Class", "Language"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  if (!changes.newOsis.size) return;

  const homeLanguage = new Map();
  for (let i = 0; i < rawData.length; i++) {
    const id = toText_(rawData[i][2]);
    if (changes.newOsis.has(id) && !homeLanguage.has(id) && toText_(rawData[i][18])) homeLanguage.set(id, properCase_(firstWord_(rawData[i][18])));
  }

  const languagesById = new Map();
  for (const g of guardianRows) {
    if (!g.language) continue;
    if (!languagesById.has(g.student_id)) languagesById.set(g.student_id, []);
    const list = languagesById.get(g.student_id);
    if (list.indexOf(g.language) === -1) list.push(g.language);
  }

  // Existing log rows by OSIS
  const lastRow = sheet.getLastRow();
  const logged = new Map();
  if (lastRow > 1) {
    sheet.getRange(2, 3, lastRow - 1, 1).getValues().forEach((r, i) => {
      const osis = String(r[0]).trim();
      if (osis) logged.set(osis, i + 2);
    });
  }

  const today = new Date();
  const toAppend = [];
  for (let i = 1; i < masterTable.length; i++) {
    const m = masterTable[i];
    const osis = String(m[0]).trim();
    if (!changes.newOsis.has(osis)) continue;

    const language = (languagesById.get(osis) || []).join(", ") || homeLanguage.get(osis) || "";
    const row = [today, "ATS", m[0], m[1], m[2], m[3], m[4], language];

    const priorRow = logged.get(osis);
    if (priorRow) sheet.getRange(priorRow, 1, 1, row.length).setValues([row]);
    else toAppend.push(row);
  }

  if (toAppend.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, toAppend.length, toAppend[0].length).setValues(toAppend);
  }
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).setNumberFormat("M/d/yyyy");
}

/** Build report lines ("Name (OSIS)") for the sidebar / email. */
function atsChangeReportLines_(masterTable, changes) {
  const newStudents = [];
  const discharged = [];
  for (let i = 1; i < masterTable.length; i++) {
    const osis = String(masterTable[i][0]).trim();
    const line = `${masterTable[i][1]} (${osis})`;
    if (changes.newOsis.has(osis)) newStudents.push(line);
    if (changes.dischargedOsis.has(osis)) discharged.push(line);
  }
  return { newStudents, discharged };
}
