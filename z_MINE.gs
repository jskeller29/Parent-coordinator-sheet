// ======================================================================
// FILE: z_MINE.gs   (OWNER'S SHEET ONLY — never copy this into another school's sheet)
// PURPOSE: Everything that only the owner's own sheet does:
//   • Build: history from "Old Master Table" (Master Table + Directory, which
//     takes non-Raw-Data students only when column F Status is "MASTER") and
//     "Old Phone Contacts" (Directory + Phone Contacts), the last-word name split,
//     the "New Students" tab
//   • Phone Contacts: the "tall" layout (one row per phone number)
//   • Quick update: also rewrites the student's Phone Contacts rows
//   • Contacts sync: reads up to 4 phones / 3 emails per row, emails a summary
//   • 🐞 Debug menu: Reset + Load Fake Data, Reset to Display Mode, Copy + Reset the Copy, Copy + Fake Data, Debug toasts
//   • Dev tools: generateFakeLegacyData, formatting, email toggles
//
// HOW IT PLUGS IN: shared code calls getOwnerHooks_() (Core.gs), which returns
// ownerHooks_() below when this file is present and null everywhere else. No
// shared function is redefined here, so the file load order doesn't matter.
// ======================================================================

/** The owner-only steps the shared code asks for (see getOwnerHooks_ in Core.gs). */
function ownerHooks_() {
  return {
    // --- Build (buildAllDerivedSheets) ---
    readMasterStatusSnapshot: ownerReadMasterStatusSnapshot_,
    adjustMasterTable: injectHistoricalMasterTable_,
    adjustDirectory: (ss, directory) => injectHistoricalDirectory_(ss, mergeOldPhoneContactsIntoDirectory_(ss, directory)),
    splitFullName: ownerSplitFullName_,
    buildPhoneContactsData: ownerBuildPhoneContactsData_,
    contactLogSiteColumn: 11, // the owner's Contact Log hides/shows K (shared sheets: H)
    logNewStudents: ownerLogNewStudents_, // the "New Students" tab
    montserratSheets: ["Master Table", "Parents Divided", "Directory", "Combined Contact Tracking", "Phone Contacts"],

    // --- Quick update (updateOsisDeltas_) ---
    updateOsisDeltas: ownerUpdateOsisDeltas_,

    // --- Google Contacts sync (syncToGoogleContacts) ---
    contactsSync: {
      listContactGroups: () => People.ContactGroups.list(),
      rowReader: ownerPhoneContactsRowReader_,
      afterSync: ownerEmailSyncSummary_
    },

    // --- 🐞 Debug menu + toasts (onOpen / runAction_) ---
    addMenus: ownerAddMenus_,
    debugOn: ownerDebugOn_,
    debugRun: ownerDebugRun_
  };
}

// ======================================================================
// 🐞 DEBUG MENU
// "Reset + Load Fake Data" / "Reset to Display Mode" run the Regenerate.gs
// tools (after you type RESET, then their own "Are you sure?"). "Copy Sheet, Then
// Reset the Copy" wipes only a new copy, never this sheet; "Copy Sheet, Then Load
// Fake Data on the Copy" fills only a new copy, never this sheet. "Toggle Debug"
// turns on a toast after every edit, build, sync and sheet open saying whether
// it worked, including errors the code normally only writes to the log.
// ======================================================================
const OWNER_DEBUG_PROP_ = "OWNER_DEBUG_TOASTS";
let ownerDebugDepth_ = 0; // only the outermost action reports

function ownerAddMenus_(ui) {
  ui.createMenu("🐞 Debug")
    .addItem("🧪 Reset + Load Fake Data", "ownerMenuResetAndLoadFakeData")
    .addItem("🧹 Reset to Display Mode", "ownerMenuResetToDisplayMode")
    .addItem("📄 Copy Sheet, Then Reset the Copy", "ownerMenuCopyAndResetToDisplayMode")
    .addItem("🧪 Copy Sheet, Then Load Fake Data on the Copy", "ownerMenuCopyForFakeData")
    .addSeparator()
    .addItem(`🐞 Toggle Debug (now ${ownerDebugOn_() ? "ON" : "OFF"})`, "ownerMenuToggleDebug")
    .addToUi();
}

function ownerMenuResetAndLoadFakeData() {
  if (ownerConfirmWipe_("Reset + Load Fake Data")) runAction_("Reset + Load Fake Data", () => resetAndLoadFakeData());
}

function ownerMenuResetToDisplayMode() {
  if (ownerConfirmWipe_("Reset to Display Mode")) runAction_("Reset to Display Mode", () => resetToDisplayMode());
}

/**
 * Makes a copy of this sheet (same Drive folder) and runs the Reset to Display Mode steps
 * (resetToDisplayModeSteps_, Regenerate.gs) on the COPY only. This sheet is never changed.
 * The copy also gets this script, z_MINE.gs included: delete z_MINE.gs from the copy's
 * Extensions > Apps Script before handing it to another school.
 */
function ownerMenuCopyAndResetToDisplayMode() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resp = ui.alert(
    "📄 Copy Sheet, Then Reset the Copy",
    "This makes a copy of this sheet in the same Drive folder and wipes the COPY to Display Mode " +
    "(empty Raw Data, Contact Log, Events and built tabs; every tab not on the keep list is deleted).\n\n" +
    "THIS sheet is not changed. Continue?",
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  runAction_("Copy + Reset to Display Mode", () => {
    ss.toast("Copying the sheet…", "📄 Copy + Reset", 30);
    const copy = ownerCopyThisSheet_(ss, "Display Mode");

    ss.toast("Resetting the copy to Display Mode…", "📄 Copy + Reset", 60);
    try {
      resetToDisplayModeSteps_(copy); // Regenerate.gs: same steps as Reset to Display Mode
      SpreadsheetApp.flush();
    } catch (e) {
      ownerShowCopyLink_(copy.getUrl(), "⚠️ The copy was made but its reset stopped partway",
        "Error: " + (e && e.message ? e.message : e) + "<br>This sheet was not changed. You can delete the copy.");
      throw e;
    }

    ownerShowCopyLink_(copy.getUrl(), "✅ Copy reset to Display Mode",
      "This sheet was not changed. Before sharing the copy with another school, delete z_MINE.gs from " +
      "the copy's Extensions &gt; Apps Script.");
  });
}

/** Copies this sheet into the same Drive folder as "<name> - <label> (MM/dd/yy)" and opens the copy by ID. */
function ownerCopyThisSheet_(ss, label) {
  SpreadsheetApp.flush(); // the copy should include every pending write
  const file = DriveApp.getFileById(ss.getId());
  const parents = file.getParents();
  const folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yy");
  const copyFile = file.makeCopy(`${ss.getName()} - ${label} (${stamp})`, folder);
  return SpreadsheetApp.openById(copyFile.getId());
}

// ----------------------------------------------------------------------
// COPY SHEET, THEN LOAD FAKE DATA ON THE COPY
// Loading fake data runs code that always works on the ACTIVE spreadsheet (the
// Events tables, the build, the Override Report) and the build keeps its memory
// (new-student snapshot) in this script's document properties. So ownerRunOnCopy_
// makes the copy the active spreadsheet for the run, then switches back and puts
// this sheet's document properties back exactly as they were.
// ----------------------------------------------------------------------
function ownerMenuCopyForFakeData() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resp = ui.alert(
    "🧪 Copy Sheet, Then Load Fake Data on the Copy",
    "This makes a copy of this sheet in the same Drive folder, wipes the COPY and loads the fake data " +
    "into it (takes a minute or two).\n\nTHIS sheet is not changed. Continue?",
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  runAction_("Copy + Load Fake Data", () => {
    ss.toast("Copying the sheet…", "🧪 Copy + Fake Data", 30);
    const copy = ownerCopyThisSheet_(ss, "Fake Data");

    ss.toast("Loading fake data into the copy…", "🧪 Copy + Fake Data", 120);
    try {
      ownerRunOnCopy_(copy, () => loadFakeDataSteps_(copy)); // Regenerate.gs: same steps as Reset + Load Fake Data
      SpreadsheetApp.flush();
    } catch (e) {
      ownerShowCopyLink_(copy.getUrl(), "⚠️ The copy was made but loading fake data stopped partway",
        "Error: " + (e && e.message ? e.message : e) + "<br>This sheet was not changed. You can delete the copy.");
      throw e;
    }

    ownerShowCopyLink_(copy.getUrl(), "✅ Copy loaded with fake data",
      "This sheet was not changed. The copy was built like every other school's sheet: delete z_MINE.gs " +
      "from the copy's Extensions &gt; Apps Script.<br><br>The copy has no automatic triggers (edits, nightly): " +
      "run 🚀 App Menu → Initial Setup there if you want to test typing in it.");
  });
}

/**
 * Runs fn with copy as the active spreadsheet, then makes this sheet active again and
 * restores this script's document properties (settings, build snapshot, ...) exactly.
 * Refuses to run fn unless the switch really happened, so this sheet is never wiped.
 * The copy is built like every other school's sheet (no owner hooks: regular Phone
 * Contacts layout and name split, no New Students tab), because z_MINE.gs gets deleted
 * from the copy.
 */
function ownerRunOnCopy_(copy, fn) {
  const original = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getDocumentProperties();
  const savedProps = props.getProperties();
  const savedHooks = OWNER_HOOKS_CACHE_;
  SpreadsheetApp.setActiveSpreadsheet(copy);
  try {
    if (SpreadsheetApp.getActiveSpreadsheet().getId() !== copy.getId() || copy.getId() === original.getId()) {
      throw new Error("Could not switch to the copy, so nothing was loaded (this sheet was not touched).");
    }
    OWNER_HOOKS_CACHE_ = null; // getOwnerHooks_() (Core.gs) now answers "no owner file" for this run
    return fn();
  } finally {
    OWNER_HOOKS_CACHE_ = savedHooks;
    SpreadsheetApp.setActiveSpreadsheet(original);
    props.setProperties(savedProps, true); // true = also delete keys the run added
  }
}

function ownerShowCopyLink_(url, title, message) {
  const html = HtmlService.createHtmlOutput(
    `<div style="font-family:Arial,sans-serif;font-size:14px">` +
    `<p>${message}</p><p><a href="${url}" target="_blank">Open the copy</a></p></div>`)
    .setWidth(440).setHeight(260);
  SpreadsheetApp.getUi().showModalDialog(html, title);
}

function ownerMenuToggleDebug() {
  const props = PropertiesService.getDocumentProperties();
  const on = props.getProperty(OWNER_DEBUG_PROP_) !== "true";
  props.setProperty(OWNER_DEBUG_PROP_, String(on));
  SpreadsheetApp.getActiveSpreadsheet().toast(
    on ? "Debug toasts are ON: every edit, build, sync and open now reports ✅ / ⚠️ / ❌." : "Debug toasts are OFF.",
    "🐞 Debug", 8);
  onOpen(); // refresh the menu label
}

/** Extra guard for the menu: these wipe the sheet, including the Old Master Table / Old Phone Contacts tabs. */
function ownerConfirmWipe_(what) {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    `⚠️ ${what}`,
    "This wipes the data on THIS sheet (Raw Data, Contact Log, Events and every built tab) and deletes every " +
    "tab not on the keep list (KEEP_TABS_ in Regenerate.gs), incl. Old Master Table and Old Phone Contacts. " +
    "Only use it on a test copy.\n\nType RESET to continue.",
    ui.ButtonSet.OK_CANCEL);
  return res.getSelectedButton() === ui.Button.OK && res.getResponseText().trim().toUpperCase() === "RESET";
}

function ownerDebugOn_() {
  return PropertiesService.getDocumentProperties().getProperty(OWNER_DEBUG_PROP_) === "true";
}

/** Runs one action and toasts ✅ / ⚠️ (logged errors or warnings) / ❌ (it threw). */
function ownerDebugRun_(label, fn) {
  if (ownerDebugDepth_ > 0) return fn();
  const name = typeof label === "function" ? label() : label;
  const problems = [];
  const original = { error: console.error, warn: console.warn };
  const describe = args => Array.prototype.map.call(args, a => (a && a.message) ? a.message : String(a)).join(" ");
  try {
    console.error = function () { problems.push("❌ " + describe(arguments)); return original.error.apply(console, arguments); };
    console.warn = function () { problems.push("⚠️ " + describe(arguments)); return original.warn.apply(console, arguments); };
  } catch (e) { /* console can't be wrapped: only thrown errors are reported */ }

  ownerDebugDepth_++;
  const started = Date.now();
  try {
    const result = fn();
    const secs = ((Date.now() - started) / 1000).toFixed(1);
    if (problems.length) ownerDebugToast_(`${name}: finished in ${secs}s with ${problems.length} problem(s): ${problems.slice(0, 3).join(" | ")}`, "⚠️ Debug", 15);
    else ownerDebugToast_(`${name}: OK (${secs}s)`, "✅ Debug", 4);
    return result;
  } catch (e) {
    ownerDebugToast_(`${name}: FAILED — ${e && e.message ? e.message : e}`, "❌ Debug", 20);
    throw e;
  } finally {
    ownerDebugDepth_--;
    try { console.error = original.error; console.warn = original.warn; } catch (e) {}
  }
}

function ownerDebugToast_(message, title, seconds) {
  try { SpreadsheetApp.getActiveSpreadsheet().toast(message, title, seconds); } catch (e) {}
}

// ======================================================================
// BUILD: OLD MASTER TABLE / OLD PHONE CONTACTS HISTORY
// ======================================================================

/** OSIS -> Status from the Master Table before the build rewrites it (Status column found by header). */
function ownerReadMasterStatusSnapshot_(ss) {
  const oldMasterStatusMap = new Map();
  const oldMasterSheetForReport = ss.getSheetByName("Master Table");
  if (oldMasterSheetForReport) {
    const oldMasterData = oldMasterSheetForReport.getDataRange().getValues();
    
    // Dynamically find the Status column to prevent breaking when transitioning to the new 7-column format!
    let statusIndex = 5; 
    const headerRow = oldMasterData.find(r => String(r[0]).trim().toUpperCase() === "OSIS") || oldMasterData[3];
    if (headerRow) {
      const foundIdx = headerRow.indexOf("Status");
      if (foundIdx !== -1) statusIndex = foundIdx;
    }
    
    for (let i = 4; i < oldMasterData.length; i++) {
      const osis = String(oldMasterData[i][0]).trim();
      if (osis && !osis.includes(",")) oldMasterStatusMap.set(osis, oldMasterData[i][statusIndex]);
    }
  }
  return oldMasterStatusMap;
}

/** Splits a full name for the Directory: the LAST word is the last name (used by splitFullName_ on this sheet). */
function ownerSplitFullName_(fullName) {
  const parts = String(fullName).trim().split(" ");
  if (parts.length === 1) return { first: parts[0], last: "" };
  const last = parts.pop();
  const first = parts.join(" ");
  return { first: first, last: last };
}

/**
 * Helper 1: Injects Historical Master Table into Master Table.
 * CRITICAL FIX: Proactively scans the Master Table array. If any names are "Unknown",
 * or if they are non-ATS ("Discharged"), it aggressively patches the rows with 
 * legacy data so the user never has to delete the Master Table!
 */
function injectHistoricalMasterTable_(ss, masterTable) {
  const oldSheet = ss.getSheetByName("Old Master Table");
  if (!oldSheet) return masterTable;

  const oldData = oldSheet.getDataRange().getValues();

  // Dynamically find header indices so it grabs the real columns
  let headerIdx = 0;
  let headers = [];
  for (let i = 0; i < Math.min(10, oldData.length); i++) {
      if (String(oldData[i][0]).toUpperCase().includes("OSIS")) {
          headerIdx = i;
          headers = oldData[i];
          break;
      }
  }
  
  const h = (colName) => {
      const idx = headers.findIndex(col => String(col).trim().toLowerCase().includes(colName.toLowerCase()));
      return idx !== -1 ? idx : -1;
  };

  const osisCol = h("OSIS") !== -1 ? h("OSIS") : 0;
  const studentCol = h("Student") !== -1 ? h("Student") : 1;
  const guardianCol = h("Guardian") !== -1 ? h("Guardian") : 2;
  const siteCol = h("Site") !== -1 ? h("Site") : 3;
  const classCol = h("Class") !== -1 ? h("Class") : (h("Teacher") !== -1 ? h("Teacher") : 4);
  const statusCol = h("Status") !== -1 ? h("Status") : 5;

  // Create a map of existing Master Table rows so we know exactly which row to overwrite
  const existingOsisMap = new Map();
  for (let r = 1; r < masterTable.length; r++) {
     existingOsisMap.set(String(masterTable[r][0]).trim(), r);
  }

  // Start reading AFTER the discovered header row
  for (let i = headerIdx + 1; i < oldData.length; i++) {
    // Retain numeric types to prevent #N/A lookups elsewhere
    let osisVal = oldData[i][osisCol];
    const osisStr = String(osisVal).trim();
    if (!osisStr || osisStr === "#N/A") continue;
    
    if (/^\d+$/.test(osisStr)) osisVal = Number(osisStr);

    const rawStatus = String(oldData[i][statusCol]).trim();
    let finalStatus = rawStatus;
    if (finalStatus === "" || finalStatus === "#N/A") finalStatus = "Unknown";

    const studentName = String(oldData[i][studentCol]).trim() === "#N/A" || !String(oldData[i][studentCol]).trim() ? "Unknown" : String(oldData[i][studentCol]).trim();
    
    // Safely pull Guardian dynamically, strip tags
    let rawGuardian = String(oldData[i][guardianCol]).trim();
    if (rawGuardian === "#N/A" || !rawGuardian) rawGuardian = "Unknown";
    let cleanGuardian = rawGuardian.replace(/\([^)]+\)/g, '').replace(/\s+/g, ' ').trim();
    if (!cleanGuardian) cleanGuardian = "Unknown";
    
    const site = String(oldData[i][siteCol]).trim() === "#N/A" ? "" : String(oldData[i][siteCol]).trim();
    const classCode = String(oldData[i][classCol]).trim() === "#N/A" ? "" : String(oldData[i][classCol]).trim();

    if (existingOsisMap.has(osisStr)) {
        const rowIndex = existingOsisMap.get(osisStr);
        const currentStudent = String(masterTable[rowIndex][1]).trim();
        const currentGuardian = String(masterTable[rowIndex][2]).trim();
        const currentSite = String(masterTable[rowIndex][3]).trim();
        const currentClass = String(masterTable[rowIndex][4]).trim();
        const currentStatus = String(masterTable[rowIndex][6]).trim().toLowerCase();

        // 🎯 THE OVERRIDE: Aggressively rescue ANY "Unknown" or blank names using Old Master Table
        if (currentStudent.toLowerCase() === "unknown" || currentStudent === "") {
            masterTable[rowIndex][1] = studentName;
        }
        if (currentGuardian.toLowerCase() === "unknown" || currentGuardian === "") {
            masterTable[rowIndex][2] = cleanGuardian;
        }

        // 🎯 Rescue non-ATS students: If current status includes Discharge/Unknown, forcefully apply legacy Status/Class/Site
        if (currentStatus.includes("discharge") || currentStatus.includes("unknown") || currentStatus === "") {
            masterTable[rowIndex][6] = finalStatus;
            
            if (currentSite === "" || currentSite.toLowerCase() === "unknown") {
                masterTable[rowIndex][3] = site;
            }
            if (currentClass === "" || currentClass.toLowerCase() === "unknown") {
                masterTable[rowIndex][4] = classCode;
            }
        }
    } else {
        // Student wasn't in the Master Table at all, append them to the bottom!
        masterTable.push([
          osisVal,
          studentName,
          cleanGuardian,
          site,
          classCode,
          "",           // Placeholder for Sibling
          finalStatus
        ]);
        existingOsisMap.set(osisStr, masterTable.length - 1); // Track in case of duplicates
    }
  }

  return masterTable;
}

/**
 * Merges Old Phone Contacts into the (Raw Data) Directory built by buildDirectoryData_.
 * FIX: Safely splits multi-parent homes, matches them using clean names, 
 * appends to the current parent, and reformats historical tags perfectly!
 */
function mergeOldPhoneContactsIntoDirectory_(ss, directory) {
  const oldPhoneSheet = ss.getSheetByName("Old Phone Contacts");
  if (!oldPhoneSheet) return directory;
  const oldPhoneMap = buildOldPhoneMap_(oldPhoneSheet);

  // 2. Loop through the active directory and append extra phones to ATS parents
  for (let i = 1; i < directory.length; i++) {
    const row = directory[i];
    const osis = String(row[2]).trim();
    const parentContactRaw = String(row[5]); 
    
    if (!oldPhoneMap.has(osis)) continue;
    
    const contacts = oldPhoneMap.get(osis);
    const matchedOldContacts = new Set();
    
    // Split the cell by newlines in case there are multiple Active ATS parents
    const activeLines = parentContactRaw.split("\n");
    const newParentLines = [];
    
    activeLines.forEach(line => {
        const parentNameWithRel = line.split(" - ")[0].trim(); // e.g. "Jane Doe (Mother)"
        
        // Strip the relationship temporarily just so we can match "Jane Doe" perfectly
        let searchName = parentNameWithRel;
        if (searchName.includes("(")) {
            searchName = searchName.substring(0, searchName.lastIndexOf("(")).trim();
        }
        
        const match = contacts.find(c => isFuzzyMatch_(c.name, searchName));
        
        if (match) {
            matchedOldContacts.add(match);
            
            // Extract existing parts to avoid duplicating phones/emails ATS already gave us
            const existingPartsStr = line.includes(" - ") ? line.substring(line.indexOf(" - ") + 3) : "";
            const existingParts = existingPartsStr ? existingPartsStr.split(", ").map(s => s.trim()) : [];
            
            if (match.emails.length > 0) match.emails.forEach(e => { if (!existingParts.includes(e)) existingParts.push(e); });
            if (match.phones.length > 0) match.phones.forEach(p => { if (!existingParts.includes(p)) existingParts.push(p); });
            
            if (existingParts.length > 0) {
                // APPENDS to the current parent keeping the relationship tag
                newParentLines.push(`${parentNameWithRel} - ${existingParts.join(", ")}`);
            } else {
                newParentLines.push(parentNameWithRel);
            }
        } else {
            newParentLines.push(line); // No match found, leave it alone
        }
    });
    
    // 3. Append un-matched "other contacts" from historical data
    contacts.forEach(old => {
        if (!matchedOldContacts.has(old)) {
            const parts = [];
            if (old.emails.length > 0) parts.push(old.emails.join(", "));
            if (old.phones.length > 0) parts.push(old.phones.join(", "));
            
            // If historical, the Old Contacts sheet has "(Mother) Jane Doe". 
            // We flip it back to "Jane Doe (Mother)" to match the ATS format standard!
            let displayName = old.originalFirstName || old.name;
            if (displayName.startsWith("(")) {
                const closeParen = displayName.indexOf(")");
                if (closeParen !== -1) {
                    const rel = displayName.substring(0, closeParen + 1); // "(Mother)"
                    const namePart = displayName.substring(closeParen + 1).trim(); // "Jane Doe"
                    displayName = `${namePart} ${rel}`; // "Jane Doe (Mother)"
                }
            }
            
            if (parts.length > 0) {
                newParentLines.push(`${displayName} - ${parts.join(", ")}`);
            } else {
                newParentLines.push(displayName);
            }
        }
    });
    
    // Save the beautifully merged rich-text back into the directory cell
    directory[i][5] = newParentLines.join("\n");
  }
  return directory;
}

/**
 * Helper 2: Injects Historical Master Table into Directory.
 * UPGRADED: Allows all historical non-ATS students to safely port into the Directory.
 */
function injectHistoricalDirectory_(ss, directory) {
  const oldMaster = ss.getSheetByName("Old Master Table");
  const oldPhone = ss.getSheetByName("Old Phone Contacts");
  if (!oldMaster) return directory;

  const currentOsis = new Set(directory.map(row => String(row[2]).trim()));
  const oldPhoneMap = buildOldPhoneMap_(oldPhone);
  const oldMasterData = oldMaster.getDataRange().getValues();

  // Dynamically find headers here as well
  let headerIdx = 0;
  let headers = [];
  for (let i = 0; i < Math.min(10, oldMasterData.length); i++) {
      if (String(oldMasterData[i][0]).toUpperCase().includes("OSIS")) {
          headerIdx = i;
          headers = oldMasterData[i];
          break;
      }
  }
  
  const h = (colName) => {
      const idx = headers.findIndex(col => String(col).trim().toLowerCase().includes(colName.toLowerCase()));
      return idx !== -1 ? idx : -1;
  };

  const osisCol = h("OSIS") !== -1 ? h("OSIS") : 0;
  const studentCol = h("Student") !== -1 ? h("Student") : 1;
  const guardianCol = h("Guardian") !== -1 ? h("Guardian") : 2;
  const siteCol = h("Site") !== -1 ? h("Site") : 3;
  const classCol = h("Class") !== -1 ? h("Class") : (h("Teacher") !== -1 ? h("Teacher") : 4);

  for (let i = headerIdx + 1; i < oldMasterData.length; i++) {
    let osisVal = oldMasterData[i][osisCol];
    const osisStr = String(osisVal).trim();
    if (!osisStr || osisStr === "#N/A" || currentOsis.has(osisStr)) continue;
    
    if (/^\d+$/.test(osisStr)) osisVal = Number(osisStr); // Prevent VLOOKUP mismatches

    // 🛑 GATEKEEPER: Non-Raw-Data students only join the Directory when Old Master Table
    // column F (Status) is "MASTER"; Discharged, blank, etc. stay out
    const rawStatus = String(oldMasterData[i][5]).trim();
    if (rawStatus.toUpperCase() !== "MASTER") continue; 

    const studentName = String(oldMasterData[i][studentCol]).trim() === "#N/A" || !String(oldMasterData[i][studentCol]).trim() ? "Unknown" : String(oldMasterData[i][studentCol]).trim();
    const parsedStudent = splitFullName_(studentName);
    
    let rawGuardian = String(oldMasterData[i][guardianCol]).trim();
    if (rawGuardian === "#N/A" || !rawGuardian) rawGuardian = "Unknown";
    const cleanGuardian = rawGuardian.replace(/\([^)]+\)/g, '').replace(/\s+/g, ' ').trim();
    
    const site = String(oldMasterData[i][siteCol]).trim() === "#N/A" ? "" : String(oldMasterData[i][siteCol]).trim();
    const classCode = String(oldMasterData[i][classCol]).trim() === "#N/A" ? "" : String(oldMasterData[i][classCol]).trim();

    let contactStr = cleanGuardian;
    
    if (oldPhoneMap.has(osisStr)) {
      const contacts = oldPhoneMap.get(osisStr);
      const lines = [];
      
      contacts.forEach(c => {
          const parts = [];
          if (c.emails.length > 0) parts.push(c.emails.join(", "));
          if (c.phones.length > 0) parts.push(c.phones.join(", "));
          
          let displayName = c.originalFirstName || c.name;
          if (displayName.startsWith("(")) {
              const closeParen = displayName.indexOf(")");
              if (closeParen !== -1) {
                  const rel = displayName.substring(0, closeParen + 1); 
                  const namePart = displayName.substring(closeParen + 1).trim(); 
                  displayName = `${namePart} ${rel}`; 
              }
          }

          lines.push(parts.length > 0 ? `${displayName} - ${parts.join(", ")}` : displayName);
      });
      
      if (lines.length > 0) contactStr = lines.join("\n");
    }

    directory.push([parsedStudent.last, parsedStudent.first, osisVal, site, classCode, contactStr]);
  }

  // Final Directory Sort (same order as the shared Directory)
  const dirHeaders = directory[0];
  const body = directory.slice(1).sort(compareDirectoryRows_);

  return [dirHeaders, ...body];
}

/**
 * Shared Helper: Builds the fast memory map of Old Phone Contacts
 * FIX: Strips out the "(Mother)" and "| Student: Bobby |" bloat so the names match!
 */
function buildOldPhoneMap_(oldPhoneSheet) {
  const oldPhoneMap = new Map();
  if (!oldPhoneSheet) return oldPhoneMap;

  const phoneData = oldPhoneSheet.getDataRange().getValues();
  if (phoneData.length < 2) return oldPhoneMap;
  const headers = phoneData[0];
  const h = (name) => headers.indexOf(name);

  const osisIdx = h("OSIS");
  const fNameIdx = h("First Name");
  const lNameIdx = h("Last Name");
  
  const p1Idx = h("Phone 1 - Value");
  const p2Idx = h("Phone 2 - Value");
  const p3Idx = h("Phone 3 - Value");
  const p4Idx = h("Phone 4 - Value");
  
  const e1Idx = h("E-mail 1 - Value");
  const e2Idx = h("E-mail 2 - Value");
  const e3Idx = h("E-mail 3 - Value");

  if (osisIdx === -1) return oldPhoneMap;

  for (let i = 1; i < phoneData.length; i++) {
    const osis = String(phoneData[i][osisIdx]).trim();
    if (!osis) continue;

    const rawFName = fNameIdx > -1 ? String(phoneData[i][fNameIdx] || "") : "";
    const rawLName = lNameIdx > -1 ? String(phoneData[i][lNameIdx] || "") : "";
    
    let fNameClean = rawFName;
    let lNameClean = rawLName;
    
    // Remove "(Relationship)" from the start of the name so we can match it later
    if (fNameClean.startsWith("(")) {
        const closeParen = fNameClean.indexOf(")");
        if (closeParen !== -1) {
            fNameClean = fNameClean.substring(closeParen + 1).trim();
        }
    }
    
    // Ignore the "| Student: Bobby |" block placed in the Last Name column
    if (lNameClean.startsWith("|")) lNameClean = "";
    
    const cleanFullName = `${fNameClean} ${lNameClean}`.trim();

    const emails = [];
    [e1Idx, e2Idx, e3Idx].forEach(idx => {
        if (idx > -1 && phoneData[i][idx]) emails.push(String(phoneData[i][idx]).trim());
    });

    const phones = [];
    [p1Idx, p2Idx, p3Idx, p4Idx].forEach(idx => {
        if (idx > -1 && phoneData[i][idx]) phones.push(String(phoneData[i][idx]).trim());
    });

    if (!oldPhoneMap.has(osis)) oldPhoneMap.set(osis, []);
    
    oldPhoneMap.get(osis).push({
        name: cleanFullName,          // Just "Jane Doe" (Allows a perfect match!)
        originalFirstName: rawFName,  // Preserves "(Mother) Jane Doe" so we can flip it later
        emails: [...new Set(emails)],
        phones: [...new Set(phones)]
    });
  }
  return oldPhoneMap;
}

// ======================================================================
// BUILD: NEW STUDENTS TAB (was in Build.gs; owner's sheet only)
// Keeps a running "New Students" log. Every build (including the nightly
// sync) logs the students detectAtsChanges_ (Build.gs) found new to Raw Data
// (ATS) or back from Discharged. Only Raw Data counts; Source is always "ATS".
// The tab is created once and only ever added to (a row is re-dated when its
// student re-enters). The shared build report (sidebar) still lists new
// students on every sheet; only this tab is owner-only.
// ======================================================================
const OWNER_NEW_STUDENTS_SHEET_ = "New Students";

/**
 * "75M094" -> "P94M" (D75 school naming). Falls back to "" when the DBN
 * doesn't look like a D75 code.
 */
function ownerSchoolShortName_(rawData) {
  for (let i = 0; i < rawData.length; i++) {
    const m = toText_(rawData[i][1]).toUpperCase().match(/^75([MXKQR])(\d{3})$/);
    if (m) return `P${Number(m[2])}${m[1]}`;
  }
  return "";
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
function ownerLogNewStudents_(ss, masterTable, changes, rawData, guardianRows) {
  let sheet = ss.getSheetByName(OWNER_NEW_STUDENTS_SHEET_);
  if (!sheet) sheet = ss.insertSheet(OWNER_NEW_STUDENTS_SHEET_);
  if (!toText_(sheet.getRange(1, 1).getValue())) {
    const school = ownerSchoolShortName_(rawData);
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

// ======================================================================
// PHONE CONTACTS ("tall" layout) + Old Phone Contacts
// ======================================================================

/**
 * The owner's Phone Contacts (used by buildPhoneContactsData_ on this sheet).
 * GENERATES A "TALL" DATASET: One row per phone number to prevent Apple Contacts merging.
 * NEW: Now includes historical parents and unlisted ATS parents.
 */
function ownerBuildPhoneContactsData_(ss, studentMap, guardianRows, masterTable, combinedContactTrackingArray, parentSquareIndex, nycsaIndex) {
  const headers = [
    "Name Prefix", "First Name", "Middle Name", "Last Name", "Name Suffix", 
    "Phonetic First Name", "Phonetic Middle Name", "Phonetic Last Name", "Nickname", "File As", 
    "E-mail 1 - Label", "E-mail 1 - Value", "E-mail 2 - Label", "E-mail 2 - Value", "E-mail 3 - Label", "E-mail 3 - Value",
    "Phone 1 - Type", "Phone 1 - Value", "Phone 2 - Type", "Phone 2 - Value", "Phone 3 - Type", "Phone 3 - Value", "Phone 4 - Type", "Phone 4 - Value",
    "Address 1 - Label", "Address 1 - Country", "Address 1 - Street", "Address 1 - Extended Address", "Address 1 - City", "Address 1 - Region", 
    "Address 1 - Postal Code", "Address 1 - PO Box", "Organization Name", "Organization Department", "Birthday", 
    "Event 1 - Label", "Event 1 - Value", "Relation 1 - Label", "Relation 1 - Value", "Website 1 - Label", 
    "Website 1 - Value", "Custom Field 1 - Label", "Custom Field 1 - Value", "Notes", "Labels", "OSIS"
  ];
  
  const out = [headers];
  const h = (colName) => headers.indexOf(colName); 

  const masterMap = new Map(masterTable.slice(1).filter(r => isRealStudentId_(r[0])).map(r => [toText_(r[0]), r]));
  const tracking = buildTrackingNotesMap_(ss, combinedContactTrackingArray); 
  
  const oldContactsMap = getOldContactsMap_(ss);

  // 🔥 FIX: Build Old Master Guardians map to rescue any "Unknown" ATS parents
  const oldMasterGuardians = new Map();
  const oldMasterSheet = ss.getSheetByName("Old Master Table");
  if (oldMasterSheet) {
    const oldData = oldMasterSheet.getDataRange().getValues();
    let headerIdx = 0, headersArr = [];
    for (let i = 0; i < Math.min(10, oldData.length); i++) {
      if (String(oldData[i][0]).toUpperCase().includes("OSIS")) {
        headerIdx = i; headersArr = oldData[i]; break;
      }
    }
    const oCol = headersArr.findIndex(c => String(c).trim().toLowerCase().includes("osis"));
    const gCol = headersArr.findIndex(c => String(c).trim().toLowerCase().includes("guardian") || String(c).trim().toLowerCase().includes("parent"));
    
    if (oCol !== -1 && gCol !== -1) {
      for (let i = headerIdx + 1; i < oldData.length; i++) {
        const oStr = String(oldData[i][oCol]).trim();
        const gVal = String(oldData[i][gCol]).trim();
        if (oStr && gVal && gVal.toLowerCase() !== "unknown" && gVal !== "#N/A") {
          oldMasterGuardians.set(oStr, gVal);
        }
      }
    }
  }

  const groupedGuardians = new Map();
  const matchedOldKeys = new Set(); // TRACKS WHICH OLD RECORDS WERE MERGED

  // Each student's guardians in order (looked up once instead of filtering every row per guardian)
  const guardiansByStudentId = groupGuardiansByStudent_(guardianRows);

  // Group by identical guardian to combine notes and student details
  for (const g of guardianRows) {
    const id = g.student_id;
    const master = masterMap.get(id) || [];
    const student = studentMap.get(id) || { student: "", classCode: "", site: "", teachers: "" };
    
    // 🔥 OVERRIDE: Rescue Unknown Guardians using Old Master Table
    let rawGuardianName = toText_(g.guardianOutput);
    let currentRel = toText_(g.relationship);
    
    if (!rawGuardianName || rawGuardianName.toLowerCase() === "unknown" || rawGuardianName === "#N/A") {
      if (oldMasterGuardians.has(id)) {
         const oldG = oldMasterGuardians.get(id); // e.g. "Homer Simpson (Father)"
         
         // Remove parentheticals for the base name
         rawGuardianName = oldG.replace(/\([^)]+\)/g, '').replace(/\s+/g, ' ').trim();
         
         // Extract the relationship if present
         const match = oldG.match(/\(([^)]+)\)/);
         if (match) currentRel = match[1];
         else currentRel = "Guardian";
         
         // Overwrite the object so it propagates properly downstream
         g.guardianOutput = rawGuardianName;
         g.relationship = currentRel;
      }
    }

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
    if (student.teachers && toText_(student.teachers).trim() !== "") classDisplay += ` - ${toText_(student.teachers)}`;
    else if (master[4] && toText_(master[4]).trim() !== "") classDisplay = toText_(master[4]);

    const otherGuardian = (guardiansByStudentId.get(id) || []).find(row => row !== g) || null;
    const psStatus = getParentsSquareStatus_(g, parentSquareIndex);
    const nycsaStatus = getNYCSAStatus_(g, otherGuardian, nycsaIndex); 

    // FALLBACK OVERRIDE: If Master Table gave us "Unknown", use the Raw ATS Data (or the rescued data!)
    let pName = toText_(master[2] || guardianName);
    if (!pName || pName.toLowerCase() === "unknown" || pName === "#N/A") pName = guardianName;

    let studentBlock = `STUDENT NAME: ${toText_(student.student)}\nRELATIONSHIP: ${toText_(g.relationship)}\nLANGUAGE: ${toText_(g.language)}\nOSIS: ${id}\nParent Name: ${pName}\nClass: ${classDisplay}`;
    if (student.site && toText_(student.site) !== "") studentBlock += `\nSite: ${toText_(student.site)}`;
    studentBlock += `\nParents Square: ${psStatus || "No"}\nNYCSA: ${nycsaStatus || "No"}`;

    const studentNotes = tracking.get(id);
    if (studentNotes && studentNotes.trim() !== "") studentBlock += `\n\nNote:\n${studentNotes}`;

    group.students.push({ studentName: toText_(student.student), osis: id, block: studentBlock, site: toText_(student.site), classCode: classDisplay });
  }

  // Build output rows for ACTIVE ATS Parents
  for (const group of groupedGuardians.values()) {
    const studentNames = group.students.map(s => s.studentName).join(", ");
    const osisList = group.students.map(s => s.osis).join(", ");
    const finalNotes = group.students.map(s => s.block).join("\n\n---\n");
    const hasSiblings = group.students.length > 1;

    const basePhoneLabel = group.students.map(s => {
      let label = hasSiblings ? `{${String(s.studentName).split(" ")[0]}} ` : "";
      if (s.site && s.site !== "") label += `Site: ${s.site}, `;
      return (label + `Class: ${s.classCode}`).trim();
    }).join(" & ");

    const orgNames = group.students.map(s => hasSiblings ? `{${String(s.studentName).split(" ")[0]}} OSIS: ${s.osis}` : `OSIS: ${s.osis}`).join(", ");

    let allPhones = [];
    let extraEmails = [];
    
    const primaryPhoneClean = normalizePhoneForMerge_(group.phone);
    const primaryEmailClean = String(group.email).toLowerCase().trim();
    
    // 1. Add Old Matched Phones/Emails for ALL SIBLINGS (Old Phone Contacts numbers come first, so they are Contact 1+)
    group.students.forEach(studentObj => {
      const sOsis = String(studentObj.osis).trim();
      const oldRecords = oldContactsMap.get(sOsis) || [];
      
      // Find all matching old records for this specific sibling's OSIS
      const matchedOldRecords = oldRecords.filter(old => isFuzzyMatch_(group.guardianName, old.firstName));
      
      matchedOldRecords.forEach(matchedOldRecord => {
        // MARK AS MERGED SO WE DON'T DUPLICATE THEM LATER (Trimmed to prevent mismatch)
        matchedOldKeys.add(`${sOsis}_${String(matchedOldRecord.firstName).trim()}`); 

        // Merge phones
        matchedOldRecord.phones.forEach(oldP => {
          const cleanOldP = normalizePhoneForMerge_(oldP.val);
          if (cleanOldP && !allPhones.some(p => p.clean === cleanOldP)) {
            allPhones.push({ val: oldP.val, clean: cleanOldP });
          }
        });
        
        // Merge emails
        matchedOldRecord.emails.forEach(oldE => {
          const cleanOldE = String(oldE.val).toLowerCase().trim();
          if (cleanOldE && cleanOldE !== primaryEmailClean && !extraEmails.some(ee => ee.clean === cleanOldE)) {
            extraEmails.push({ label: `Old Email ${extraEmails.length + 1}`, val: oldE.val, clean: cleanOldE });
          }
        });
      });
    });

    // 2. Add the ATS Phone after the Old Phone Contacts numbers (skipped if it duplicates one)
    if (primaryPhoneClean && !allPhones.some(p => p.clean === primaryPhoneClean)) {
      allPhones.push({ val: group.phone, clean: primaryPhoneClean });
    }

    // Safety fallback: if no phones exist, output at least one row for emails/notes
    if (allPhones.length === 0) allPhones.push({ val: "", clean: "" });

    // 3. THE PIVOT: Create a distinct row for EVERY phone number
    allPhones.forEach((phoneObj, index) => {
      const row = new Array(headers.length).fill("");
      
      row[h("First Name")] = `(${group.relationship}) ${group.guardianName}`.trim(); 
      row[h("Last Name")] = `| Student: ${studentNames} |`;                         
      row[h("Organization Name")] = orgNames;                                       
      row[h("Notes")] = finalNotes;                                                 
      row[h("Labels")] = "School Directory";                                     
      row[h("OSIS")] = osisList;                                                
      
      row[h("E-mail 1 - Value")] = group.email; 
      if (extraEmails.length > 0) { row[h("E-mail 2 - Label")] = extraEmails[0].label; row[h("E-mail 2 - Value")] = extraEmails[0].val; }
      if (extraEmails.length > 1) { row[h("E-mail 3 - Label")] = extraEmails[1].label; row[h("E-mail 3 - Value")] = extraEmails[1].val; }
      
      let typeLabel = basePhoneLabel;
      if (allPhones.length > 1) typeLabel = `Contact ${index + 1}`; 
      
      row[h("Phone 1 - Type")] = typeLabel;                                         
      row[h("Phone 1 - Value")] = phoneObj.val;  

      out.push(row);
    });
  }

  // 4. ADD HISTORICAL OR UNLISTED ATS PARENTS
  for (const [oldOsis, oldRecords] of oldContactsMap.entries()) {
    for (const old of oldRecords) {
      if (!matchedOldKeys.has(`${String(oldOsis).trim()}_${String(old.firstName).trim()}`)) {
        
        let allPhones = old.phones;
        if (allPhones.length === 0) allPhones.push({ val: "", label: "Historical" });

        // Grab student data (From Active raw data or Injected Master Table)
        const sOsisStr = String(oldOsis).trim();
        const activeStudent = studentMap.get(sOsisStr);
        const isActive = !!activeStudent && !!activeStudent.student; 
        const masterRow = masterMap.get(sOsisStr);

        // ✨ FIX: Fully generate the normal Notes profile block for these parents with strict "Unknown" bypass
        let sNameVal = masterRow ? toText_(masterRow[1]) : `Historical OSIS: ${sOsisStr}`;
        if (sNameVal.toLowerCase() === "unknown" || sNameVal === "#N/A") sNameVal = `Historical OSIS: ${sOsisStr}`;
        let sName = isActive ? toText_(activeStudent.student) : sNameVal;

        let sRel = "Guardian";
        if (old.originalFirstName && old.originalFirstName.includes("(")) {
            const match = old.originalFirstName.match(/\(([^)]+)\)/);
            if (match) sRel = match[1];
        }
        
        let sClassVal = masterRow ? toText_(masterRow[4]) : "";
        if (sClassVal.toLowerCase() === "unknown" || sClassVal === "#N/A") sClassVal = "";
        let sClass = isActive ? toText_(activeStudent.classCode) : sClassVal;
        if (isActive && activeStudent.teachers && toText_(activeStudent.teachers).trim() !== "") {
            sClass += ` - ${toText_(activeStudent.teachers)}`;
        }
        
        let sSiteVal = masterRow ? toText_(masterRow[3]) : "";
        if (sSiteVal.toLowerCase() === "unknown" || sSiteVal === "#N/A") sSiteVal = "";
        let sSite = isActive ? toText_(activeStudent.site) : sSiteVal;

        let sParentNameVal = masterRow ? toText_(masterRow[2]) : old.firstName;
        if (!sParentNameVal || sParentNameVal.toLowerCase() === "unknown" || sParentNameVal === "#N/A") sParentNameVal = old.firstName;
        let sParentName = sParentNameVal;
        
        let finalNotes = `STUDENT NAME: ${sName}\nRELATIONSHIP: ${sRel}\nLANGUAGE: Unknown\nOSIS: ${sOsisStr}\nParent Name: ${sParentName}\nClass: ${sClass}`;
        if (sSite) finalNotes += `\nSite: ${sSite}`;
        finalNotes += `\nParents Square: Unknown\nNYCSA: Unknown`;
        
        const studentNotes = tracking.get(sOsisStr);
        if (studentNotes && studentNotes.trim() !== "") finalNotes += `\n\nNote:\n${studentNotes}`;
        
        // Append the original historical notes underneath if they had any
        if (old.notes && old.notes.trim() !== "") {
            finalNotes += `\n\n[Historical Record Notes]:\n${old.notes}`;
        }

        allPhones.forEach((phoneObj, index) => {
          const row = new Array(headers.length).fill("");
          
          row[h("First Name")] = old.firstName;
          row[h("Labels")] = "School Directory";
          row[h("OSIS")] = oldOsis;
          row[h("Last Name")] = `| Student: ${sName} |`;
          row[h("Organization Name")] = old.orgName || `OSIS: ${oldOsis}`;
          row[h("Notes")] = finalNotes;
          
          if (old.emails.length > 0) { row[h("E-mail 1 - Label")] = old.emails[0].label || "home"; row[h("E-mail 1 - Value")] = old.emails[0].val; }
          if (old.emails.length > 1) { row[h("E-mail 2 - Label")] = old.emails[1].label; row[h("E-mail 2 - Value")] = old.emails[1].val; }
          if (old.emails.length > 2) { row[h("E-mail 3 - Label")] = old.emails[2].label; row[h("E-mail 3 - Value")] = old.emails[2].val; }
          
          let typeLabel = phoneObj.label;
          if (!typeLabel) typeLabel = allPhones.length > 1 ? `Contact ${index + 1}` : "Contact";
          
          row[h("Phone 1 - Type")] = typeLabel;
          row[h("Phone 1 - Value")] = phoneObj.val;
          
          out.push(row);
        });
      }
    }
  }

  return out;
}

function getOldContactsMap_(ss) {
  const map = new Map();
  const sheet = ss.getSheetByName("Old Phone Contacts");
  if (!sheet) return map;
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return map;
  
  const headers = data[0];
  const h = (name) => headers.indexOf(name);
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const osis = toText_(row[h("OSIS")]);
    if (!osis) continue;
    
    if (!map.has(osis)) map.set(osis, []);
    
    map.get(osis).push({
      firstName: row[h("First Name")],
      lastName: row[h("Last Name")],
      orgName: row[h("Organization Name")],
      notes: row[h("Notes")],
      emails: [
        { label: row[h("E-mail 1 - Label")], val: row[h("E-mail 1 - Value")] },
        { label: row[h("E-mail 2 - Label")], val: row[h("E-mail 2 - Value")] },
        { label: row[h("E-mail 3 - Label")], val: row[h("E-mail 3 - Value")] }
      ].filter(e => e.val),
      phones: [
        { label: row[h("Phone 1 - Label")], val: row[h("Phone 1 - Value")] },
        { label: row[h("Phone 2 - Label")], val: row[h("Phone 2 - Value")] },
        { label: row[h("Phone 3 - Label")], val: row[h("Phone 3 - Value")] },
        { label: row[h("Phone 4 - Label")], val: row[h("Phone 4 - Value")] }
      ].filter(p => p.val)
    });
  }
  return map;
}

/** * CRITICAL FIX: Strips all parentheticals prior to distance checking.
 * This guarantees "Homer Simpson" and "Homer Simpson (Father)" will perfectly match.
 */
function isFuzzyMatch_(newName, oldName) {
  if (!newName || !oldName) return false;
  
  // Clean names by removing anything wrapped in parentheses before comparing
  const cleanNew = String(newName).replace(/\([^)]+\)/g, '').trim();
  const cleanOld = String(oldName).replace(/\([^)]+\)/g, '').trim();
  
  const n1 = cleanNew.toLowerCase().replace(/[^a-z]/g, '');
  const n2 = cleanOld.toLowerCase().replace(/[^a-z]/g, '');
  
  if (n1 === '' || n2 === '') return false;
  
  if (n1.includes(n2) || n2.includes(n1)) return true; 
  
  const track = Array(n2.length + 1).fill(null).map(() => Array(n1.length + 1).fill(null));
  for (let i = 0; i <= n1.length; i += 1) track[0][i] = i;
  for (let j = 0; j <= n2.length; j += 1) track[j][0] = j;
  for (let j = 1; j <= n2.length; j += 1) {
    for (let i = 1; i <= n1.length; i += 1) {
      const ind = n1[i - 1] === n2[j - 1] ? 0 : 1;
      track[j][i] = Math.min(track[j][i - 1] + 1, track[j - 1][i] + 1, track[j - 1][i - 1] + ind);
    }
  }
  return track[n2.length][n1.length] <= 2; 
}

function normalizePhoneForMerge_(value) {
  const digits = toText_(value).replace(/\D/g, "");
  if (!digits) return "";
  return digits.length === 11 && digits.startsWith("1") ? digits.slice(1) : digits;
}

// ======================================================================
// QUICK UPDATE (updateOsisDeltas_ on this sheet)
// Rebuilds Combined Contact Tracking + Phone Contacts in memory ONCE, then
// rewrites only the rows of each OSIS in the list.
// ======================================================================
function ownerUpdateOsisDeltas_(osisList) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let combinedContactTracking, phoneContacts, osisIndex;
  try {
    const rawSheet = ss.getSheetByName("Raw Data") || ss.getSheetByName("RAW Data");
    if (!rawSheet) return;

    const rawData = rawSheet.getDataRange().getValues();
    if (rawData.length < 4) return;

    const overrides = getOverrideData_(ss);
    const studentMap = buildStudentMap_(rawData, overrides.students);
    const guardianRows = buildGuardianObjects_(rawData, overrides.guardians);
    
    const parentSquareIndex = buildParentSquareIndex_(rawData);
    const nycsaIndex = buildNYCSAIndex_(rawData);
    
    const masterTableData = getExistingMasterTable_(ss);

    // ✨ THE FIX: Fetch the Contact Log and Event Log data here so the builder doesn't crash!
    const contactLogSheet = ss.getSheetByName("Contact Log");
    const contactLogData = contactLogSheet ? contactLogSheet.getDataRange().getValues() : [];
    
    const eventLogData = backendReadBlock_(ss, "EVENT_LOG");

    // Pass ALL 4 arguments into the builder
    combinedContactTracking = buildCombinedContactTrackingData_(ss, masterTableData, contactLogData, eventLogData);
    phoneContacts = ownerBuildPhoneContactsData_(ss, studentMap, guardianRows, masterTableData, combinedContactTracking, parentSquareIndex, nycsaIndex);

    const phoneHeaders = phoneContacts[0];
    osisIndex = phoneHeaders.indexOf("OSIS") !== -1 ? phoneHeaders.indexOf("OSIS") : 5;
  } catch (err) {
    // If it fails, log it so it doesn't fail completely silently
    console.error("Delta update failed: " + err.message);
    return;
  }

  osisList.forEach(targetOsis => {
    try {
      writeSingleRowDelta_(ss, "Combined Contact Tracking", combinedContactTracking, targetOsis, 0);
      writeMultipleRowDeltas_(ss, "Phone Contacts", phoneContacts, targetOsis, osisIndex);
    } catch (err) {
      console.error("Delta update failed: " + err.message);
    }
  });
}

// ======================================================================
// GOOGLE CONTACTS SYNC (the shared syncToGoogleContacts asks for these)
// ======================================================================

/** Reads one tall Phone Contacts row: up to 3 emails and 4 phones, each with its label. */
function ownerPhoneContactsRowReader_(headers) {
  const col = name => headers.indexOf(name);
  return row => {
    const parentName = row[col("First Name")];
    const studentNames = row[col("Last Name")];
    const notes = row[col("Notes")];
    const orgName = row[col("Organization Name")];
    const orgTitle = row[col("Organization Department")];

    let emailsArr = [];
    let phonesArr = [];
    
    [1, 2, 3].forEach(num => {
      const vIdx = headers.indexOf(`E-mail ${num} - Value`);
      const lIdx = headers.indexOf(`E-mail ${num} - Label`);
      if (vIdx !== -1 && row[vIdx]) emailsArr.push({ val: row[vIdx], label: lIdx !== -1 ? row[lIdx] : "home" });
    });
    
    [1, 2, 3, 4].forEach(num => {
      const vIdx = headers.indexOf(`Phone ${num} - Value`);
      const lIdx = headers.indexOf(`Phone ${num} - Type`);
      if (vIdx !== -1 && row[vIdx]) phonesArr.push({ val: row[vIdx], label: lIdx !== -1 ? row[lIdx] : "mobile" });
    });

    return {
      parentName: parentName,
      upsert: directoryResourceName => ownerUpsertContact_(parentName, studentNames, emailsArr, phonesArr, notes, orgName, orgTitle, directoryResourceName)
    };
  };
}

/** After a finished sync: emails the results and the build's new / discharged students (when turned on). */
function ownerEmailSyncSummary_(deletedCount, createdCount, editedCount) {
  // --- NEW: EMAIL NOTIFICATION LOGIC ---
  const emailEnabled = PropertiesService.getDocumentProperties().getProperty('EMAIL_NOTIFICATIONS_ENABLED');
  
  if (emailEnabled === 'true') {
    const userEmail = Session.getActiveUser().getEmail(); 
    const subject = "🔄 Nightly Sync Complete";
    // New / discharged students from the last build (Raw Data only; see detectAtsChanges_)
    let studentLines = "";
    try {
      const report = getBuildLogData();
      if (report) {
        const list = arr => (arr && arr.length) ? arr.map(x => `  - ${x}`).join("\n") : "  (none)";
        studentLines = `\n\nNew Students (ATS): ${(report.newStudents || []).length}\n${list(report.newStudents)}` +
                       `\n\nDischarged (no longer in ATS): ${(report.discharged || []).length}\n${list(report.discharged)}`;
      }
    } catch (e) {}
    const body = `Your Google Contacts sync has finished!\n\nResults:\nDeleted: ${deletedCount}\nCreated: ${createdCount}\nEdited: ${editedCount}${studentLines}\n\nCheck your spreadsheet for full logs.`;
    
    try {
      MailApp.sendEmail(userEmail, subject, body);
    } catch(e) {
      // Fails silently if email quota is reached or permissions are missing
    }
  }
}

/**
 * The owner's single-contact push (all of a row's emails / phones).
 * STRICT MATCHING: Searches Google Contacts by Name, Phone, and Email without making assumptions.
 */
function ownerUpsertContact_(parentName, studentNames, emailsArr, phonesArr, notes, orgName, orgTitle, directoryResourceName = null) {
  try {
    const fullName = String(parentName || "").trim();
    const fNameStr = String(studentNames || "").trim(); 
    if (!fullName) return null;

    const cache = CacheService.getScriptCache();
    
    // 1. STRICT CACHE KEY: Stop assuming blanks are matches. Extract exact identifiers.
    const primaryEmail = emailsArr.length > 0 ? String(emailsArr[0].val).toLowerCase().trim() : "";
    const primaryPhoneRaw = phonesArr.length > 0 ? String(phonesArr[0].val).trim() : "";
    const primaryPhoneClean = normalizePhoneForMerge_(primaryPhoneRaw);
    
    // Fallback to raw text if no digits (e.g. they typed "discharge") so it stays perfectly unique
    const searchPhone = primaryPhoneClean || primaryPhoneRaw.toLowerCase(); 
    
    const rowIdentity = fullName.toLowerCase() + "|" + fNameStr.toLowerCase() + "|" + searchPhone + "|" + primaryEmail;
    
    let hash = 0;
    for (let i = 0; i < rowIdentity.length; i++) hash = ((hash << 5) - hash) + rowIdentity.charCodeAt(i) | 0; 
    const cacheKey = "CONTACT_V7_" + Math.abs(hash); // Bumping to V7 to clear old buggy cached identities

    let existingContact = null;
    let cachedResourceName = cache.get(cacheKey);

    if (cachedResourceName) {
      try { existingContact = People.People.get(cachedResourceName, { personFields: 'names,phoneNumbers,emailAddresses,memberships,organizations,biographies' }); } 
      catch (e) { cache.remove(cacheKey); }
    }

    if (!existingContact) {
      let pageToken;
      do {
        const response = People.People.Connections.list('people/me', { pageSize: 1000, personFields: 'names,phoneNumbers,emailAddresses,memberships,organizations,biographies', pageToken: pageToken });
        if (response.connections) {
          for (const person of response.connections) {
            if (person.names && person.names.length > 0) {
              const apiGiven = String(person.names[0].givenName || "").toLowerCase().trim();
              const apiFamily = String(person.names[0].familyName || "").toLowerCase().trim();
              
              if (apiGiven === fullName.toLowerCase() && apiFamily === fNameStr.toLowerCase()) {
                
                // NO ASSUMPTIONS: Strictly evaluate exact strings to prevent false overwriting
                const apiPhonesClean = person.phoneNumbers ? person.phoneNumbers.map(p => normalizePhoneForMerge_(p.value)) : [];
                const apiPhonesRaw = person.phoneNumbers ? person.phoneNumbers.map(p => String(p.value).toLowerCase().trim()) : [];
                const apiEmails = person.emailAddresses ? person.emailAddresses.map(e => String(e.value).toLowerCase().trim()) : [];
                
                let hasMatchingPhone = false;
                
                if (searchPhone === "") {
                   // If the row on the sheet literally has no phone, check the email
                   if (primaryEmail !== "") {
                       hasMatchingPhone = apiEmails.includes(primaryEmail);
                   } else {
                       // If the row has NO phone and NO email, only match if the API contact also has neither
                       hasMatchingPhone = (apiPhonesRaw.length === 0 && apiEmails.length === 0);
                   }
                } else {
                   // Match by exact digit sequence OR exact raw text (e.g. "discharge")
                   hasMatchingPhone = (primaryPhoneClean !== "" && apiPhonesClean.includes(primaryPhoneClean)) || apiPhonesRaw.includes(searchPhone);
                }
                
                if (hasMatchingPhone) {
                  existingContact = person; 
                  break; 
                }
              }
            }
          }
        }
        if (existingContact) break;
        pageToken = response.nextPageToken;
      } while (pageToken);
    }

    // Build the payload
    const contactData = {
      names: [{ givenName: fullName, familyName: fNameStr }],
      organizations: [], phoneNumbers: [], emailAddresses: [], biographies: notes ? [{ value: String(notes) }] : []
    };

    if (String(orgName || "").trim() !== "" || String(orgTitle || "").trim() !== "") {
      contactData.organizations.push({ name: String(orgName || "").trim(), title: String(orgTitle || "").trim() });
    }
    
    emailsArr.forEach(e => { if (String(e.val).trim() !== "") contactData.emailAddresses.push({ value: String(e.val), type: String(e.label || "home") }); });
    phonesArr.forEach(p => { if (String(p.val).trim() !== "") contactData.phoneNumbers.push({ value: String(p.val), type: String(p.label || "mobile") }); });
    
    let attempts = 0;
    while (attempts < 3) {
      attempts++;
      try {
        if (directoryResourceName) contactData.memberships = [{ contactGroupMembership: { contactGroupResourceName: directoryResourceName } }];

        if (existingContact) {
          contactData.etag = existingContact.etag;
          if (existingContact.memberships && directoryResourceName) {
            const others = existingContact.memberships.filter(m => m.contactGroupMembership?.contactGroupResourceName !== directoryResourceName);
            contactData.memberships = contactData.memberships.concat(others);
          }
          const updated = People.People.updateContact(contactData, existingContact.resourceName, { updatePersonFields: 'names,phoneNumbers,emailAddresses,biographies,memberships,organizations' });
          cache.put(cacheKey, updated.resourceName, 21600);
          return { resourceName: updated.resourceName, isNew: false, isEdited: true, name: fullName, logStr: "Strict Update" }; 
        } else {
          const created = People.People.createContact(contactData);
          cache.put(cacheKey, created.resourceName, 21600);
          return { resourceName: created.resourceName, isNew: true, isEdited: false, name: fullName }; 
        }
      } catch (e) {
        if (String(e).includes("404")) throw e;
        if (attempts >= 3) throw e; 
        Utilities.sleep(2000 * attempts); 
      }
    }
  } catch (err) { return null; }
}

// ======================================================================
// OWNER DEV TOOLS (run from the Apps Script editor)
// ======================================================================
function TOGGLE_EMAIL_NOTIFICATIONS() {
  const props = PropertiesService.getDocumentProperties();
  const currentState = props.getProperty('EMAIL_NOTIFICATIONS_ENABLED');
  const newState = (currentState === 'false') ? 'true' : 'false';
  props.setProperty('EMAIL_NOTIFICATIONS_ENABLED', newState);
  
  const status = (newState === 'true') ? "ENABLED ✅" : "DISABLED ❌";
  SpreadsheetApp.getUi().alert("📧 Email Preference", `Notifications are now: ${status}`, SpreadsheetApp.getUi().ButtonSet.OK);
}

function DEBUG_TEST_EMAIL() {
  const userEmail = "jkeller7@schools.nyc.gov"; // <--- UPDATE THIS TO YOUR REAL EMAIL
  try {
    MailApp.sendEmail(userEmail, "🧪 TEST EMAIL", "This confirms your @p94m.org account allows script emails.");
    SpreadsheetApp.getUi().alert("🚀 Success", "Check your inbox: " + userEmail, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert("❌ Error", "Mail service failed: " + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Generates fake data for testing BOTH 'Old Master Table' and 'Old Phone Contacts'.
 * Perfectly mimics current live RAW data, keeps Kirk Van Houten as a legacy test, 
 * and strictly enforces the 《§COL§》 and 《§ROW§》 string templates.
 */
function generateFakeLegacyData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // =====================================================================
  // 1. BUILD "OLD PHONE CONTACTS"
  // =====================================================================
  const contactsSheetName = "Old Phone Contacts";
  let contactsSheet = ss.getSheetByName(contactsSheetName);
  if (!contactsSheet) {
    contactsSheet = ss.insertSheet(contactsSheetName);
  } else {
    contactsSheet.clear().clearFormats().clearConditionalFormatRules();
    if (contactsSheet.getFilter()) contactsSheet.getFilter().remove();
    contactsSheet.getBandings().forEach(b => b.remove());
  }

  const contactHeaders = [
    "OSIS", "Name Prefix", "First Name", "Middle Name", "Last Name", "Name Suffix", 
    "Phonetic First Name", "Phonetic Middle Name", "Phonetic Last Name", "Nickname", 
    "File As", "E-mail 1 - Label", "E-mail 1 - Value", "E-mail 2 - Label", "E-mail 2 - Value", 
    "E-mail 3 - Label", "E-mail 3 - Value", "Phone 1 - Label", "Phone 1 - Value", 
    "Phone 2 - Label", "Phone 2 - Value", "Phone 3 - Label", "Phone 3 - Value", 
    "Phone 4 - Label", "Phone 4 - Value", "Address 1 - Label", "Address 1 - Country", 
    "Address 1 - Street", "Address 1 - Extended Address", "Address 1 - City", 
    "Address 1 - Region", "Address 1 - Postal Code", "Address 1 - PO Box", 
    "Organization Name", "Organization Title", "Organization Department", "Birthday", 
    "Event 1 - Label", "Event 1 - Value", "Relation 1 - Label", "Relation 1 - Value", 
    "Website 1 - Label", "Website 1 - Value", "Custom Field 1 - Label", 
    "Custom Field 1 - Value", "Notes", "Labels"
  ];
  
  contactsSheet.getRange(1, 1, 1, contactHeaders.length).setValues([contactHeaders]);

  // HELPER 1: Easily map data into the 47 columns
  function buildContactRow(osis, firstName, lastName, emails, phones) {
    let row = new Array(47).fill("");
    row[0] = osis;                                    
    row[2] = firstName;                               
    row[4] = "Student: " + lastName;                  
    row[46] = "Master Copy";                          
    
    emails.forEach((email, i) => {
      if (i < 3) {
        row[11 + (i * 2)] = firstName.split("(")[0].trim(); 
        row[12 + (i * 2)] = email;                          
      }
    });

    phones.forEach((phone, i) => {
      if (i < 4) {
        row[17 + (i * 2)] = firstName.split("(")[0].trim(); 
        row[18 + (i * 2)] = phone;                          
      }
    });
    return row;
  }

  // Generate Phone Contact Data 
  const fakeContacts = [
    buildContactRow(123456789, "Homer Simpson (Father)", "Bart Simpson", ["HOMERSIMPSON@gmail.com", "HOMERSIMPSON2@gmail.com"], ["(111) 111-1111", "(111) 111-1112"]),
    buildContactRow(123456789, "Marge Simpson (Mother)", "Bart Simpson", ["MARGE.SIMPSON@gmail.com"], ["(222) 222-2222"]),
    buildContactRow(234567891, "Morticia Addams (Mother)", "Wednesday Addams", ["MORTICIA.ADDAMS@gmail.com"], ["(333) 333-3333"]),
    buildContactRow(234567891, "Gomez Addams (Father)", "Wednesday Addams", ["GOMEZ.ADDAMS@gmail.com"], ["(333) 333-4444"]),
    buildContactRow(345678912, "Mr. Incredible (Father)", "Dash Incredible", ["MR.INCREDIBLE@gmail.com"], ["(444) 444-4444"]),
    buildContactRow(345678912, "Mrs. Incredible (Mother)", "Dash Incredible", [], ["(555) 555-5555"]),
    buildContactRow(789123456, "Rio Morales (Mother)", "Miles Morales", ["RIO.MORALES@gmail.com"], ["(777) 777-7777"]),
    buildContactRow(789123456, "Jefferson Morales (Father)", "Miles Morales", ["JEFFERSON.MORALES@gmail.com"], ["(888) 888-8888"]),
    buildContactRow(222222222, "Kirk Van Houten (Father)", "Milhouse Van Houten", ["KIRK@yahoo.com"], ["(222) 222-2222"]),
    buildContactRow(222222222, "Luann Van Houten (Mother)", "Milhouse Van Houten", [], ["(222) 222-2223", "(222) 222-2224"])
  ];

  contactsSheet.getRange(2, 1, fakeContacts.length, contactHeaders.length).setValues(fakeContacts);
  const fullContactsTableRange = contactsSheet.getRange(1, 1, fakeContacts.length + 1, contactHeaders.length);
  fullContactsTableRange.createFilter();
  fullContactsTableRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  contactsSheet.setFrozenRows(1);

  // =====================================================================
  // 2. BUILD "OLD MASTER TABLE"
  // =====================================================================
  const masterSheetName = "Old Master Table";
  let masterSheet = ss.getSheetByName(masterSheetName);
  if (!masterSheet) {
    masterSheet = ss.insertSheet(masterSheetName);
  } else {
    masterSheet.clear().clearFormats().clearConditionalFormatRules();
    if (masterSheet.getFilter()) masterSheet.getFilter().remove();
    masterSheet.getBandings().forEach(b => b.remove());
  }

  const masterHeaders = ["OSIS", "Student", "Guardian", "Site", "Teacher", "Status", "Code"];
  masterSheet.getRange("A4:G4").setValues([masterHeaders]);

  // HELPER 2: Strictly enforces the 《§COL§》 and 《§ROW§》 template
  // If email or phone is passed as "", it perfectly leaves a blank gap between delimiters
  function buildMasterCode(guardianList) {
    let codeStr = "NameName《§COL§》Relationship《§COL§》Language《§COL§》Email《§COL§》Phone《§COL§》Primary";
    guardianList.forEach(g => {
      codeStr += `《§ROW§》${g.name}《§COL§》${g.rel}《§COL§》${g.lang}《§COL§》${g.email}《§COL§》${g.phone}《§COL§》${g.primary}`;
    });
    return codeStr;
  }

  // The actual injection array using the Template Builder
  const fakeMasterData = [
    [
      123456789, "Bart Simpson", "Homer Simpson (Father) and Marge Simpson (Mother)", 15, "Y41: Tessa", "Master",
      buildMasterCode([
        { name: "Homer Simpson", rel: "Father", lang: "English", email: "HOMERSIMPSON@gmail.com\nHOMERSIMPSON2@gmail.com", phone: "(111) 111-1111\n(111) 111-1112", primary: "Yes" },
        { name: "Marge Simpson", rel: "Mother", lang: "English", email: "MARGE.SIMPSON@gmail.com", phone: "(222) 222-2222", primary: "No" }
      ])
    ],
    [
      234567891, "Wednesday Addams", "Morticia Addams (Mother) and Gomez Addams (Father)", 51, "Y66: Carl", "Master",
      buildMasterCode([
        { name: "Morticia Addams", rel: "Mother", lang: "Spanish", email: "MORTICIA.ADDAMS@gmail.com", phone: "(333) 333-3333", primary: "Yes" },
        { name: "Gomez Addams", rel: "Father", lang: "English", email: "GOMEZ.ADDAMS@gmail.com", phone: "(333) 333-4444", primary: "No" }
      ])
    ],
    [
      345678912, "Dash Incredible", "Mr. Incredible (Father) and Mrs. Incredible (Mother)", 188, "X30: Ham", "Master",
      buildMasterCode([
        { name: "Mr. Incredible", rel: "Father", lang: "English", email: "MR.INCREDIBLE@gmail.com", phone: "(444) 444-4444", primary: "Yes" },
        { name: "Mrs. Incredible", rel: "Mother", lang: "English", email: "", phone: "(555) 555-5555", primary: "No" } // Intentionally blank email
      ])
    ],
    [
      789123456, "Miles Morales", "Rio Morales (Mother) and Jefferson Morales (Father)", 75, "Y24", "Master",
      buildMasterCode([
        { name: "Rio Morales", rel: "Mother", lang: "Spanish", email: "RIO.MORALES@gmail.com", phone: "(777) 777-7777", primary: "Yes" },
        { name: "Jefferson Morales", rel: "Father", lang: "English", email: "JEFFERSON.MORALES@gmail.com", phone: "(888) 888-8888", primary: "No" }
      ])
    ],
    [
      222222222, "Milhouse Van Houten", "Kirk Van Houten (Father) and Luann Van Houten (Mother)", 15, "Y42: Hoover", "Master",
      buildMasterCode([
        { name: "Kirk Van Houten", rel: "Father", lang: "English", email: "KIRK@yahoo.com", phone: "(222) 222-2222", primary: "Yes" },
        { name: "Luann Van Houten", rel: "Mother", lang: "English", email: "", phone: "(222) 222-2223\n(222) 222-2224", primary: "No" } // Intentionally blank email
      ])
    ]
  ];

  masterSheet.getRange(5, 1, fakeMasterData.length, masterHeaders.length).setValues(fakeMasterData);
  const fullMasterTableRange = masterSheet.getRange(4, 1, fakeMasterData.length + 1, masterHeaders.length);
  fullMasterTableRange.createFilter();
  fullMasterTableRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREEN, true, false);
  
  masterSheet.setFrozenRows(4);
  masterSheet.autoResizeColumns(1, 6);
  masterSheet.setColumnWidth(7, 400); 

  // Map Master_2 Named Range
  const rangeName = "Master_2";
  const masterDataRange = masterSheet.getRange("A5:G");
  const existingNamedRanges = ss.getNamedRanges();
  existingNamedRanges.forEach(nr => { if (nr.getName() === rangeName) nr.remove(); });
  ss.setNamedRange(rangeName, masterDataRange);

  ss.toast("✅ Old Master Table & Old Phone Contacts fully generated!", "Success", 8);
}

/**
 * Instantly builds the Legacy Phone Override Data UI into the Contact Log.
 * Injects a dynamic BYROW/LAMBDA formula to unpack the legacy data payload.
 */
function formatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const contactLog = ss.getSheetByName("Contact Log");
  
  if (!contactLog) {
    ss.toast("Contact Log sheet not found!", "Error");
    return;
  }

  // 1. Find the "END" row in Column E
  const searchRange = contactLog.getRange("E1:E");
  const textFinder = searchRange.createTextFinder("^END$").matchEntireCell(true).useRegularExpression(true);
  const endCell = textFinder.findNext();
  
  if (!endCell) {
    ss.toast("Could not find the 'END' row in Contact Log.", "Error");
    return;
  }

  const endRow = endCell.getRow();
  const startRow = endRow + 9; // The UI block starts exactly 9 rows below "END"
  const osisCell = `$G$${endRow + 2}`; // The specific [INPUT] cell you requested

  // 2. Clear out any old text so the new dynamic array formula has room to #SPILL
  contactLog.getRange(startRow, 5, 6, 12).clearContent().clearFormat();

  // 3. Format the block (Allows room for headers + up to 3 parent rows)
  const uiRange = contactLog.getRange(startRow, 5, 4, 12); 
  uiRange.setBackground("#b45f06") // Your signature dark orange / brown
         .setFontColor("#ffffff")
         .setFontWeight("bold")
         .setVerticalAlignment("middle")
         .setWrap(true);

  // 4. Construct your master BYROW/LAMBDA array formula
  // Wrapped in an IFERROR so it remains completely blank/clean if no OSIS is searched
  const customFormula = 
    `=IFERROR(BYROW(` +
      `WRAPROWS(` +
        `SPLIT(` +
          `SUBSTITUTE(` +
            `VLOOKUP(${osisCell},'Old Master Table'!$A$5:$G,7,FALSE),` +
            `"《§ROW§》","《§COL§》"` +
          `),` +
          `"《§COL§》", FALSE, FALSE` +
        `),` +
        `6` +
      `),` +
      `LAMBDA(r, HSTACK(INDEX(r,,1), "", INDEX(r,,2),"", INDEX(r,,4),"","",INDEX(r,,5), "",INDEX(r,,3),"",INDEX(r,,6)))` +
    `), "")`;

  // 5. Inject the formula into Column E (Col 5), exactly 9 rows below END
  contactLog.getRange(startRow, 5).setFormula(customFormula);

  ss.toast("✅ Dynamic Legacy Data Formula successfully injected!", "Success", 5);
}
