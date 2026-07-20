/* =========================
   HELPERS
========================= */

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

function groupGuardiansByStudent_(guardianRows) {
  const map = new Map();
  for (const g of guardianRows) {
    if (!map.has(g.student_id)) map.set(g.student_id, []);
    map.get(g.student_id).push(g);
  }
  return map;
}

function normalizeEmail_(value) {
  return toText_(value).toLowerCase();
}

function normalizePhoneForMatch_(value) {
  const digits = toText_(value).replace(/\D/g, "");
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

function firstWord_(text) {
  return toText_(text).split(/\s+/)[0] || "";
}

function toText_(value) {
  return value == null ? "" : String(value).trim();
}

function isRealStudentId_(value) {
  return toText_(value).replace(/\D/g, "").length >= 6;
}

function compareText_(a, b) {
  return toText_(a).localeCompare(toText_(b), undefined, { sensitivity: "base" });
}

function stripTrailingParenthetical_(text) {
  return toText_(text).replace(/\s*\([^)]+\)\s*$/, "").trim();
}

function splitFullName_(name) {
  const parts = toText_(name).split(/\s+/).filter(Boolean);
  return {
    first: parts[0] || "",
    last: parts.slice(1).join(" ")
  };
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

function nightlyMasterSync() {
  try {
    if (typeof buildAllDerivedSheets === "function") {
      buildAllDerivedSheets();
    }
    
    if (typeof resetContactLogFormatting === "function") {
      resetContactLogFormatting();
    }
    
    const props = PropertiesService.getDocumentProperties();
    const syncToPhones = props.getProperty('syncToPhones') === 'true';
    
    if (syncToPhones && typeof syncToGoogleContacts === "function") {
       console.log("Nightly Rebuild finished. Phone sync TRIGGERED (Opted In).");
       syncToGoogleContacts();
    } else {
       console.log("Nightly Rebuild finished. Phone sync skipped (Opted Out).");
    }
    
  } catch(e) {
    console.error("Nightly Master Sync failed: " + e.message);
  }
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

function cleanRedundantOverrides() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName("RAW Data") || ss.getSheetByName("Raw Data");
  const overrideSheet = ss.getSheetByName("New/Edit Student");
  
  if (!rawSheet || !overrideSheet) return;

  const rawData = rawSheet.getDataRange().getValues();
  const overrideData = overrideSheet.getDataRange().getValues();

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

    overrideSheet.deleteRow(i + 1); 
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

// ==========================================
// 🚀 SIDEBAR BACKEND ENDPOINTS
// ==========================================

function getRunMode() { 
  return CacheService.getScriptCache().get('RUN_MODE') || 'FULL_SYNC'; 
}

// NOTE: getSyncStatus() is defined once in the On-Open file (returns {phase:'idle'}
// when nothing is cached). EMERGENCY_STOP_SYNC() is defined once in ZProduction.gs
// (the full version that also deletes the auto-resume triggers). Both duplicates
// were removed from here so load order can no longer decide which version wins.

function getBuildLogData() { 
  const cache = CacheService.getScriptCache().get('BUILD_REPORT_LOGS'); 
  return cache ? JSON.parse(cache) : { newStudents: [], discharged: [], nycsaUpgrades: [], psUpgrades: [] }; 
}

function startBuildOnlyProcess() { 
  const cache = CacheService.getScriptCache();
  cache.put('SYNC_UI_STATE', JSON.stringify({ phase: 'building' }), 21600); 
  try { 
    if (typeof buildAllDerivedSheets === "function") buildAllDerivedSheets(); 
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (typeof formatDirectoryAuto_ === "function") {
      formatDirectoryAuto_(ss.getSheetByName("Directory"));
    }
    
    cache.put('SYNC_UI_STATE', JSON.stringify({ phase: 'done_build_only' }), 21600); 
  } catch(e) { 
    cache.put('SYNC_UI_STATE', JSON.stringify({ phase: 'killed', error: e.message }), 21600); 
  } 
}
