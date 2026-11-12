// ======================================================================
// FILE: Input Data.gs (+ InputDataSidebar.html, YourListsSidebar.html)
// PURPOSE: Getting the user's own data into the sheet, in two parts:
//   1. INPUT DATA MODE: "📥 Input Data Mode" (🚀 App Menu) walks the user
//      through the ATS / Parent Square / NYCSA exports, then Site / Class,
//      then Your Lists.
//   2. YOUR LISTS: step 5 of Input Data Mode; the user's own lists (bus
//      routes, PTA ...) matched to students or guardians, used by every build.
// Was the two trial files Z#InputDataMode.gs and Z#OwnData.gs (coalesced
// 2026-11-11): shared code now calls these functions directly (no typeof
// guards), and the Input Data Mode step 5 always has Your Lists.
// Function names (idm_* / od_*), sidebar endpoints and the document
// property / Backend keys are unchanged, so existing sheets carry on as before.
// ======================================================================

// ======================================================================
// PART 1: INPUT DATA MODE
// "📥 Input Data Mode" walks the user through loading their three exports
// (Student Data / ATS, Parent Square, NYCSA) one at a time:
//   1. Every visible tab is hidden (and remembered).
//   2. A temporary tab ("Student Data", then "Parent Square Data", then
//      "NYCSA Data") is created; the user pastes anywhere on it, or uploads
//      the file in the sidebar instead.
//   3. The sidebar checks the tab every 2 seconds. Once data is there it finds
//      the header row, maps the columns, checks for numbers Excel cut off
//      (7.19E+09 / ####), and replaces that section of Raw Data.
//   4. The temp tab is cleared and renamed for the next step.
//   5. At the end (or on Exit) the temp tab is deleted and the hidden tabs come back.
// Step 4 "Site / Class" is filled in by hand instead: one row per class code
// (from the ATS data), with Teacher Name and Site (optional, dropdown that grows
// as new sites are typed). Save writes Raw Data BH:BI (Class / Site) as one row
// per site: BH = site, BI = that site's classes, one per line ("Y25: Carl" or
// just "Y24" with no teacher); read by parseClassCell_ (Build.gs).
// Step 5 "Your Lists": each list on its own "📥" tab, set up in the
// YourListsSidebar (PART 2 below); this part only shows them.
//
// Called from shared code: addInputDataModeMenuItem_ (Menus.gs buildMenus_,
// both 🚀 App Menus) and openInputDataMode (data update reminder, Getting
// Started). Uses shared helpers (fixRawDataNumberFormatting_,
// forceVibrantColors_, applyBandingToRawData_, lockRawDataRowHeights_,
// parseClassCell_, runBuildOnly).
// ======================================================================

var IDM_STATE_KEY = 'INPUT_DATA_MODE_STATE';

// Debug log: only with the owner's test file z_InputDataDebug.gs (it defines
// idm_debugEnabled_ and the tab snapshot idm_debugSheets_). With it, every sidebar
// call returns a step-by-step log for the sidebar's Debug panel; without it the
// idm_log_ lines do nothing and the panel never shows (delete that file on the
// copies that go to schools). Script errors are always shown, never retried silently.
var IDM_DBG = null;
function idm_debugOn_() {
  try { return typeof idm_debugEnabled_ === "function" && !!idm_debugEnabled_(); } catch (e) { return false; }
}
function idm_log_(msg) {
  if (!IDM_DBG) return;
  IDM_DBG.push("+" + (Date.now() - IDM_DBG.t0) + "ms " + msg);
}
var IDM_PLACEHOLDER = 'PASTE HERE';
// Apps Script breaks ("Sheet 123 not found") for the rest of the session if a
// call deletes the tab the user is on. So the temp tab is reused (cleared and
// renamed) between steps, and at the end it is deleted in a SEPARATE call,
// after the user has been moved off it. Its id waits here until then.
var IDM_PENDING_DELETE_KEY = 'INPUT_DATA_MODE_DELETE_TAB';

// Same headers / columns / date stamps as cleanRawDataPasteZone_ and
// updateRawDataTimestamp_ (Edits.gs). Keep them in step.
var IDM_STEPS = [
  {
    key: 'ats', tab: 'Student Data', short: 'Student Data', zone: 'Student Data (ATS / Insight)', required: true,
    col: 1, stampCell: 'A1', stampPrefix: 'ATS - DATE UPDATED: ', minMatch: 3,
    headers: ["School Year","DBN","Student ID","Last Name","First Name","Race","Gender","Age","IEP Flag","ELL Flag","504 Flag","NYSAA Eligible","Housing Status","Temporary Housing Flag","Economic Disadvantage Flag","Cohort Year","Grade","Official Class","Home Language","Guardian Name 1","Guardian Relationship 1","Guardian Phone Number 1","Guardian Email 1","Guardian Spoken Language 1","Guardian Written Language 1","Guardian Name 2","Guardian Relationship 2","Guardian Phone Number 2","Guardian Email 2","Guardian Spoken Language 2","Guardian Written Language 2"],
    numberCols: ["Student ID","Guardian Phone Number 1","Guardian Phone Number 2"],
    howTo: [
      "Go to <b>TeachHub</b> and sign in.",
      "In the search bar, look up and open <b>Insight</b>.",
      "Go to <b>Tools → Downloader</b>.",
      "Under \"Data Available for Download,\" select only <b>Guardian Contact Information</b>, then download."
    ]
  },
  {
    key: 'ps', tab: 'Parent Square Data', short: 'Parent Square', zone: 'Parent Square', required: false,
    col: 33, stampCell: 'AG1', stampPrefix: 'Parent Square - DATE UPDATED: ', minMatch: 2,
    headers: ["Name","Students","Email","Phone","Secondary Phone","Record Created","Registered?","SIS Synced"],
    numberCols: ["Phone","Secondary Phone"],
    howTo: [
      "Go to <b>ParentSquare</b>, sign in, and open <b>Admin</b>.",
      "Under <b>Data Assistant</b>, select <b>Parents</b>.",
      "Click <b>Export CSV</b>."
    ]
  },
  {
    key: 'nycsa', tab: 'NYCSA Data', short: 'NYCSA', zone: 'NYCSA', required: false,
    col: 42, stampCell: 'AP1', stampPrefix: 'NYCSA - DATE UPDATED: ', minMatch: 3,
    headers: ["Location Code","School District","School Name","Class Code","Grade Level","Student ID","Student First Name","Student Last Name","Guardian First Name","Guardian Last Name","Another Custodial NYCSA Account Exist","Guardian Home Number","Guardian Work Number","Guardian Cell Number","Relationship","Role","Account Creation Code"],
    numberCols: ["Student ID","Guardian Home Number","Guardian Work Number","Guardian Cell Number"],
    howTo: [
      "Go to <b>TeachHub</b> and sign in.",
      "Open <b>Family Access Management</b>.",
      "Export the <b>\"Students with Non-NYCSA Account\"</b> report."
    ]
  },
  {
    key: 'siteclass', tab: 'Site / Class', short: 'Site / Class', zone: 'Site / Class', required: false, manual: true,
    howTo: [
      "The <b>Site / Class</b> tab lists every class code from your Student Data.",
      "Type each class's <b>Teacher Name</b>.",
      "<b>Site</b> is optional: pick one from the dropdown or type a new one (it's added to the dropdown).",
      "Add a class that isn't listed by typing its code at the bottom of column A.",
      "Click <b>✅ Save Site / Class</b> below when you're done."
    ]
  },
  {
    key: 'lists', tab: 'Your Lists', short: 'Your Lists', zone: 'Your Lists', required: false, lists: true,
    howTo: [
      "Any other list you get: bus routes, trip forms, after school, PTA…",
      "Each list gets its own tab. Upload the file (CSV or Excel) or paste the whole sheet there; the tool works out who each row is.",
      "What you keep shows on the phone contacts (Additional Information) and as a column in the Student List or Parents Divided after the next build."
    ]
  }
];
var IDM_SITECLASS_UPDATED_KEY = 'INPUT_DATA_MODE_SITECLASS_UPDATED';

// ---------------------------------------------------------------------
// MENU + LAUNCHER
// ---------------------------------------------------------------------

/** Called from onOpen (buildMenus_). Adds the menu item; reminds the user if the mode was left on. */
function addInputDataModeMenuItem_(menu) {
  menu.addItem("📥 Input Data Mode", "openInputDataMode");
  try { idm_deleteLeftoverTab_(); } catch (e) { console.warn("Input Data Mode leftover tab: " + e.message); }
  try {
    if (idm_loadState_()) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "Input Data Mode is still on (some tabs may be hidden). Open 🚀 App Menu → 📥 Input Data Mode to finish or exit.",
        "📥 Input Data Mode", 15);
    }
  } catch (e) { console.warn("Input Data Mode reminder: " + e.message); }
}

/** Menu handler: starts the mode (or resumes it) and opens the sidebar. */
function openInputDataMode() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    try { idm_deleteLeftoverTab_(); } catch (e) { console.warn("Input Data Mode leftover tab: " + e.message); }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let state = idm_loadState_();
    if (!state) {
      if (!idm_rawSheet_(ss)) {
        SpreadsheetApp.getUi().alert('📥 Input Data Mode', 'The "Raw Data" tab is missing, so there is nowhere to put the data.', SpreadsheetApp.getUi().ButtonSet.OK);
        return;
      }
      const active = ss.getActiveSheet();
      state = { step: 0, tempId: null, hiddenIds: [], prevActiveId: active ? active.getSheetId() : null, results: {} };
      idm_saveState_(state);
      idm_ensureTempSheet_(ss, state);
      idm_hideOtherSheets_(ss, state);
    } else {
      idm_ensureTempSheet_(ss, state);
    }
    idm_saveState_(state);
  } finally {
    lock.releaseLock();
  }
  const html = htmlFromFile_('InputDataSidebar').setTitle('Input Data Mode'); // template: pulls in Styles.html
  SpreadsheetApp.getUi().showSidebar(html.setWidth(340));
}

// ---------------------------------------------------------------------
// SIDEBAR ENDPOINTS
// ---------------------------------------------------------------------

/** Current step + "last updated" dates for the sidebar. */
function idm_getState() {
  return idm_withLock_(() => idm_view_(SpreadsheetApp.getActiveSpreadsheet(), idm_loadState_()));
}

/**
 * Polled every ~2s. Only LOOKS at the temp tab (fast); when there is data it
 * answers pending=true and the sidebar shows "moving…" and calls idm_processPaste.
 */
function idm_check() {
  return idm_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const state = idm_loadState_();
    if (!state) return idm_view_(ss, null);
    const temp = idm_ensureTempSheet_(ss, state);
    idm_saveState_(state);
    if (IDM_DBG && typeof idm_debugSheets_ === "function") idm_debugSheets_(ss, state, temp); // z_InputDataDebug.gs
    if (IDM_STEPS[state.step].manual || IDM_STEPS[state.step].lists) return idm_view_(ss, state);
    const hasData = idm_tempHasData_(temp);
    idm_log_("temp tab has data: " + hasData);
    const view = idm_view_(ss, state);
    view.pending = hasData;
    return view;
  });
}

/** Moves the pasted data on the temp tab to Raw Data (can take a while). */
function idm_processPaste() {
  return idm_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const state = idm_loadState_();
    if (!state) return idm_view_(ss, null);
    const temp = idm_ensureTempSheet_(ss, state);
    if (!idm_tempHasData_(temp)) return idm_view_(ss, state);
    return idm_processTemp_(ss, state, temp, false);
  });
}

/** "Use it anyway" after a cut-off-numbers warning on a paste. */
function idm_useAnyway() {
  return idm_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const state = idm_loadState_();
    if (!state) return idm_view_(ss, null);
    const temp = idm_ensureTempSheet_(ss, state);
    return idm_processTemp_(ss, state, temp, true);
  });
}

/** "Clear & try again": empties the temp tab. */
function idm_retry() {
  return idm_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const state = idm_loadState_();
    if (!state) return idm_view_(ss, null);
    const temp = idm_ensureTempSheet_(ss, state);
    idm_resetTempSheet_(temp);
    temp.activate();
    idm_saveState_(state);
    return idm_view_(ss, state);
  });
}

/** Uploaded CSV / TXT file (raw text, parsed here so every number stays exactly as written). */
function idm_submitCsv(text, ignoreCutoff) {
  return idm_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const state = idm_loadState_();
    if (!state) return idm_view_(ss, null);
    let rows;
    try {
      rows = idm_parseCsv_(text);
    } catch (e) {
      return idm_view_(ss, state, { type: 'error', code: 'badFile', message: "That file couldn't be read as a CSV file (" + e.message + ")." });
    }
    return idm_processRows_(ss, state, rows, rows.map(r => r.map(c => String(c))), !!ignoreCutoff, true);
  });
}

/** Uploaded Excel file, already read in the sidebar (2D array of strings / numbers). */
function idm_submitRows(rows, ignoreCutoff) {
  return idm_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const state = idm_loadState_();
    if (!state) return idm_view_(ss, null);
    rows = (rows || []).map(r => (r || []).map(c => (c === null || c === undefined) ? "" : c));
    return idm_processRows_(ss, state, rows, rows.map(r => r.map(c => String(c))), !!ignoreCutoff, true);
  });
}

/** "Skip": keep this section's old Raw Data and go to the next step. */
function idm_skip() {
  return idm_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const state = idm_loadState_();
    if (!state) return idm_view_(ss, null);
    state.results[IDM_STEPS[state.step].key] = { skipped: true };
    return idm_advance_(ss, state, null);
  });
}

/** "Exit Input Data Mode": delete the temp tab, bring hidden tabs back. */
/**
 * "💾 Save & exit": saves what is on the current tab first (Site / Class entries,
 * or a paste that hasn't been moved yet), then exits. A problem (wrong export,
 * cut-off numbers, bad class code) is shown instead and the mode stays on.
 */
function idm_saveAndExit() {
  return idm_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let state = idm_loadState_();
    if (!state) return idm_view_(ss, null);
    const step = IDM_STEPS[state.step];
    const temp = idm_sheetById_(ss, state.tempId);
    let result = null;
    if (step.manual) {
      result = idm_saveSiteClass_(ss, state);
    } else if (step.lists) {
      idm_log_("save & exit: lists are saved one by one in their own sidebar");
    } else if (temp && idm_tempHasData_(temp)) {
      result = idm_processTemp_(ss, state, temp, false);
    } else {
      idm_log_("save & exit: nothing waiting on the tab");
    }
    if (result && result.notice && result.notice.type === 'error') return result; // stay, show the problem
    if (result && result.finished) { result.exited = true; return result; }       // that was the last step
    state = idm_loadState_();
    if (!state) return { active: false, finished: true, exited: true, summary: [] };
    const results = state.results;
    idm_cleanup_(ss, state, idm_anyImported_(results));
    return { active: false, finished: true, exited: true, notice: result ? result.notice : null,
             anyImported: idm_finalBuildNeeded_(state), builtAlready: idm_builtAlready_(state), guideNext: idm_guideNext_(ss), summary: idm_summary_(results) };
  });
}

function idm_exit() {
  return idm_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const state = idm_loadState_();
    if (!state) return idm_view_(ss, null);
    const results = state.results;
    idm_cleanup_(ss, state, false);
    return { active: false, finished: true, exited: true, anyImported: idm_finalBuildNeeded_(state), builtAlready: idm_builtAlready_(state), guideNext: idm_guideNext_(ss), summary: idm_summary_(results) };
  });
}

/** Progress bar click: jump to any step (back or forward). The temp tab is renamed and reset for it. */
function idm_goToStep(index) {
  return idm_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const state = idm_loadState_();
    if (!state) return idm_view_(ss, null);
    index = Number(index);
    if (!(index >= 0 && index < IDM_STEPS.length) || index === state.step) return idm_view_(ss, state);
    idm_log_("jump from step " + state.step + " to " + index);
    const temp = idm_sheetById_(ss, state.tempId);
    state.step = index;
    if (temp) {
      idm_renameTab_(ss, temp, IDM_STEPS[index].tab);
      idm_prepareTab_(ss, temp, IDM_STEPS[index]);
      temp.activate();
    } else {
      state.tempId = null;
      idm_ensureTempSheet_(ss, state);
    }
    idm_saveState_(state);
    return idm_view_(ss, state);
  });
}

/** Step 5 "Finish ✓": ends the mode (the lists themselves were saved one by one). */
function idm_finishLists() {
  return idm_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const state = idm_loadState_();
    if (!state) return idm_view_(ss, null);
    const lists = od_listSummaries_(ss).filter(l => l.ready);
    state.results.lists = lists.length ? { saved: true, rows: lists.length, lists: true } : { skipped: true };
    return idm_advance_(ss, state, null);
  });
}

// ---------------------------------------------------------------------
// BUILD BEFORE STEP 5 (owner's request 2026-11-12)
// Your Lists matches each row to the Student List (+ Raw Data guardians), so on
// a new sheet (Student List never built) or after data was moved to Raw Data in
// this run, step 5 builds the sheets first. The sidebar calls idm_buildForLists
// when the view says needsBuild. It builds like 🏗️ Build Sheets Only, but inside
// this sidebar: the user stays on the temp tab, and any tab the build made or
// showed is hidden again (and shown at the end with the others). Tried once per
// change: a failed build is reported on the lists card, not retried in a loop.
// ---------------------------------------------------------------------

function idm_listsNeedBuild_(ss, state) {
  if (state.dataChanged) return true;
  return !state.listsBuild && !studentListIsBuilt_(ss);
}

/** Raw Data has students in the ATS section (Student ID, column C, from row 7). */
function idm_hasAtsRows_(raw) {
  return raw.getLastRow() >= 7 && String(raw.getRange(7, 3).getValue()).trim() !== "";
}

/** Step 5's build ran and nothing changed since: the end screen doesn't ask for another build. */
function idm_builtAlready_(state) {
  return !!(state.listsBuild && state.listsBuild.ok && !state.dataChanged && !state.listsChanged);
}

/** The end screen offers "Build sheets now": something came in, and the sheets don't have it yet. */
function idm_finalBuildNeeded_(state) {
  return idm_anyImported_(state.results) && !idm_builtAlready_(state);
}

/** The end screen's "Next steps" (the ✅ Getting Started checklist) while it still has steps to do. */
function idm_guideNext_(ss) {
  try {
    if (PropertiesService.getDocumentProperties().getProperty('ZG_AUTO_OPEN') === 'off') return false;
    return !zg_checklist_(ss).allDone;
  } catch (e) { console.warn("Getting Started check skipped: " + e.message); return false; }
}

/** A list was saved / deleted (Your Lists sidebar): its columns need a build at the end of the mode. */
function idm_noteListsChanged_() {
  const state = idm_loadState_();
  if (!state) return;
  state.listsChanged = true;
  idm_saveState_(state);
}

function idm_buildForLists() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // 1. Which tabs show now (short lock; the build itself runs without the document lock)
  const before = idm_withLock_(() => {
    const state = idm_loadState_();
    if (!state || !IDM_STEPS[state.step].lists) return { view: idm_view_(ss, state) };
    reportBuildStep_(0); // the sidebar shows the build's steps (getSyncStatus)
    return { visible: ss.getSheets().filter(s => !s.isSheetHidden()).map(s => s.getSheetId()) };
  });
  if (!before.visible) return before.view || before; // not on step 5 any more, or an error view

  // 2. The build (holds the build lock; another build / sync running = a message, not a wait)
  let outcome;
  const raw = idm_rawSheet_(ss);
  if (!raw || !idm_hasAtsRows_(raw)) {
    outcome = { ok: false, message: "Load your Student Data first (step 1): your lists are matched to those students." };
  } else {
    try {
      buildAllDerivedSheets();
      try { formatDirectoryAuto_(ss.getSheetByName("Directory")); } catch (e) { console.warn("Directory format skipped: " + e.message); }
      outcome = { ok: true, message: "Sheets built, so your lists can find your students." };
    } catch (e) {
      console.error("Input Data Mode build before Your Lists: " + (e && e.stack ? e.stack : e));
      outcome = { ok: false, message: "The sheets couldn't be built (" + (e && e.message ? e.message : e) + "). " +
        "Your lists can only find students who are already on the Student List." };
    }
  }

  // 3. Back to how Input Data Mode had the tabs (short lock)
  return idm_withLock_(() => {
    const state = idm_loadState_();
    if (!state) return idm_view_(ss, null);
    const temp = idm_sheetById_(ss, state.tempId);
    state.hiddenIds = state.hiddenIds || [];
    if (temp) { if (temp.isSheetHidden()) temp.showSheet(); temp.activate(); }
    ss.getSheets().forEach(s => {
      const id = s.getSheetId();
      if (s.isSheetHidden() || before.visible.indexOf(id) !== -1 || (temp && id === temp.getSheetId())) return;
      try {
        s.hideSheet();
        if (state.hiddenIds.indexOf(id) === -1) state.hiddenIds.push(id); // shown again at the end
      } catch (e) { console.warn("Input Data Mode: could not hide " + s.getName() + ": " + e.message); }
    });
    state.dataChanged = false;
    if (outcome.ok) state.listsChanged = false;
    state.listsBuild = outcome;
    idm_saveState_(state);
    idm_log_("build before Your Lists: " + (outcome.ok ? "ok" : outcome.message));
    return idm_view_(ss, state);
  });
}

/** Brings the user back to the paste tab. */
function idm_goToTempSheet() {
  return idm_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const state = idm_loadState_();
    if (!state) return idm_view_(ss, null);
    idm_ensureTempSheet_(ss, state).activate();
    idm_saveState_(state);
    return idm_view_(ss, state);
  });
}

/** Last screen: run the normal Build Sheets Only (opens the Build sidebar). */
function idm_build() {
  idm_deleteLeftoverTab_();
  runBuildOnly();
  return true;
}

/** Sidebar calls this right after the last screen shows: deletes the finished temp tab. */
function idm_deleteLeftoverTab() {
  return idm_withLock_(() => {
    const deleted = idm_deleteLeftoverTab_();
    const pending = !!PropertiesService.getDocumentProperties().getProperty(IDM_PENDING_DELETE_KEY);
    // The mode is over and the user is off the temp tab: the first build's full menus can show now
    // (not earlier: buildMenus_ also deletes the leftover tab, never in the call that moved the user)
    if (!pending) { try { refreshMenusAfterFirstBuild_(); } catch (e) { console.warn("Menus not refreshed: " + e.message); } }
    return { deleted: deleted, pending: pending }; // pending -> the sidebar tries again shortly
  });
}

/** Deletes the finished temp tab unless the user is on it (then it waits for next time). */
function idm_deleteLeftoverTab_() {
  const props = PropertiesService.getDocumentProperties();
  const id = props.getProperty(IDM_PENDING_DELETE_KEY);
  if (!id) return false;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = idm_sheetById_(ss, Number(id));
  if (!sheet) { props.deleteProperty(IDM_PENDING_DELETE_KEY); return false; }
  const active = ss.getActiveSheet();
  if (active && active.getSheetId() === sheet.getSheetId()) {
    // Still on it (the switch to Raw Data didn't take). Move off now; the next try deletes it.
    const raw = idm_rawSheet_(ss);
    const landing = (raw && !raw.isSheetHidden()) ? raw
      : ss.getSheets().filter(s => !s.isSheetHidden() && s.getSheetId() !== sheet.getSheetId())[0];
    if (landing) landing.activate();
    idm_log_("leftover tab is the active tab; moved to " + (landing ? landing.getName() : "nothing") + ", will retry");
    return false;
  }
  ss.deleteSheet(sheet);
  props.deleteProperty(IDM_PENDING_DELETE_KEY);
  idm_log_("leftover temp tab deleted");
  return true;
}

// ---------------------------------------------------------------------
// PROCESSING
// ---------------------------------------------------------------------

function idm_processTemp_(ss, state, temp, ignoreCutoff) {
  const range = temp.getDataRange();
  const values = range.getValues();
  const displays = range.getDisplayValues();
  // The "PASTE HERE" label is not data
  if (values.length && String(values[0][0]).trim() === IDM_PLACEHOLDER) { values[0][0] = ""; displays[0][0] = ""; }
  idm_log_("read temp tab: " + values.length + " rows x " + (values[0] ? values[0].length : 0) + " cols");
  return idm_processRows_(ss, state, values, displays, ignoreCutoff, false);
}

/**
 * Finds the header row, checks it is this step's export, maps the columns,
 * checks for cut-off numbers and writes the section of Raw Data.
 * fromUpload only changes which buttons the sidebar shows on an error.
 */
function idm_processRows_(ss, state, values, displays, ignoreCutoff, fromUpload) {
  const step = IDM_STEPS[state.step];
  if (step.manual || step.lists) return idm_view_(ss, state, { type: 'error', code: 'manual', message: "This step is filled in on the tab — no file needed.", fromUpload: true });
  // Wrong data pasted on the tab is cleared off it right away (owner's request 2026-11-12), so the
  // right export can be pasted straight in; notice.cleared tells the sidebar. A missing Raw Data tab
  // is not the paste's fault, so that paste stays.
  const fail = (code, message, extra) => {
    const notice = Object.assign({ type: 'error', code: code, message: message, fromUpload: fromUpload }, extra || {});
    const temp = (!fromUpload && code !== 'noRaw') ? idm_sheetById_(ss, state.tempId) : null;
    if (temp) {
      idm_resetTempSheet_(temp);
      notice.cleared = true;
      idm_log_("wrong data (" + code + "): paste tab cleared");
    }
    return idm_view_(ss, state, notice);
  };

  idm_log_("step " + state.step + " (" + step.key + "), " + values.length + " rows in, fromUpload=" + fromUpload + ", ignoreCutoff=" + ignoreCutoff);
  const found = idm_findHeader_(values);
  idm_log_("header: " + (found ? found.step.key + " at row " + (found.row + 1) + " (" + found.count + " of " + found.step.headers.length + " names match)" : "NOT FOUND"));
  if (!found) {
    for (let i = 0; i < Math.min(5, values.length); i++) idm_log_("row " + (i + 1) + ": " + JSON.stringify(values[i].slice(0, 8)));
    return fail('noHeader', "I couldn't find the top row with the column names (like \"" + step.headers[2] + "\"). " +
      "Make sure you select everything with Ctrl+A (Mac: ⌘+A) before copying, so the header row comes too.");
  }
  if (found.step.key !== step.key) {
    return fail('wrongType', "This looks like <b>" + found.step.zone + "</b> data, but this step is <b>" + step.zone + "</b>. " +
      "Use the " + step.zone + " export instead (or Skip to the " + found.step.zone + " step).");
  }

  const mapped = idm_mapRows_(step, values, displays, found.row);
  idm_log_("mapped " + mapped.rows.length + " data rows; missing columns: " +
    (step.headers.filter(h => values[found.row].map(idm_norm_).indexOf(idm_norm_(h)) === -1).join(", ") || "none"));
  if (!mapped.rows.length) {
    return fail('noRows', "I found the header row, but no rows of data under it.");
  }

  if (!ignoreCutoff) {
    const problems = idm_findCutoffNumbers_(step, mapped.rows, mapped.displays);
    idm_log_("cut-off check: " + (problems.length ? JSON.stringify(problems) : "ok"));
    if (problems.length) {
      const list = problems.map(p => "<b>" + p.column + "</b> (e.g. " + p.examples.join(", ") + ")").join("<br>");
      return fail('cutoff', "Some numbers look cut off by Excel, so the real digits are missing:<br>" + list +
        "<br><br>In Excel, press Ctrl+A (Mac: ⌘+A), then double-click the line between any two column letters to widen every column. " +
        "Then copy and paste again" + (fromUpload ? ", or upload the original downloaded file (not one re-saved from Excel)." : "."),
        // A paste is cleared off the tab, so "Use it anyway" sends these rows back (idm_submitRows)
        { canUseAnyway: true, heldRows: fromUpload ? null : idm_heldRows_(values, displays) });
    }
  }

  const raw = idm_rawSheet_(ss);
  if (!raw) return fail('noRaw', 'The "Raw Data" tab is missing.');
  idm_log_("writing to " + raw.getName() + " (maxRows=" + raw.getMaxRows() + ", maxCols=" + raw.getMaxColumns() + ")");
  idm_writeSection_(raw, step, mapped.rows);
  idm_log_("Raw Data written");

  state.results[step.key] = { rows: mapped.rows.length };
  state.dataChanged = true; // step 5 builds first (idm_buildForLists)
  return idm_advance_(ss, state, { type: 'imported', message: "✅ " + step.zone + ": " + mapped.rows.length + " row(s) moved to Raw Data." });
}

/** A paste's cells for the sidebar to keep (Dates as the text shown: google.script.run can't send a Date). */
function idm_heldRows_(values, displays) {
  return values.map((r, i) => r.map((c, j) => (c instanceof Date) ? String((displays[i] || [])[j] || "") : c));
}

/** Header row = the row (in the first 20) that matches the most column names of one export. */
function idm_findHeader_(values) {
  let best = null;
  const limit = Math.min(20, values.length);
  for (let i = 0; i < limit; i++) {
    const cells = {};
    values[i].forEach(c => { const n = idm_norm_(c); if (n) cells[n] = true; });
    IDM_STEPS.forEach(step => {
      if (!step.headers) return;
      const count = step.headers.filter(h => cells[idm_norm_(h)]).length;
      if (count >= step.minMatch && (!best || count > best.count)) best = { step: step, row: i, count: count };
    });
  }
  return best;
}

/** Same column mapping as a full-export paste on Raw Data (cleanRawDataPasteZone_). */
function idm_mapRows_(step, values, displays, headerRow) {
  const header = values[headerRow].map(idm_norm_);
  const indexMap = step.headers.map(h => header.indexOf(idm_norm_(h)));
  const rows = [], disp = [];
  for (let i = headerRow + 1; i < values.length; i++) {
    const row = indexMap.map(idx => (idx !== -1 && idx < values[i].length) ? values[i][idx] : "");
    if (row.every(c => String(c).trim() === "")) continue;
    rows.push(row);
    disp.push(indexMap.map(idx => (idx !== -1 && displays[i] && idx < displays[i].length) ? String(displays[i][idx]) : ""));
  }
  return { rows: rows, displays: disp };
}

/**
 * Numbers Excel shortened because the column was too narrow:
 *  - any cell shown as 7.19E+09 / 7.19e+09 or ##### (certain);
 *  - an ID / phone column where many values end in 000 (7.19E+09 pasted as 7190000000).
 *    One real number ending in 000 is normal; a narrow Excel column cuts them all.
 */
function idm_findCutoffNumbers_(step, rows, displays) {
  const problems = [];
  const sci = /^-?\d+(\.\d+)?e[+-]?\d+$/i;
  step.headers.forEach((h, c) => {
    const examples = [];
    displays.forEach(r => {
      const d = String(r[c]).trim();
      if ((sci.test(d) || /^#{3,}$/.test(d)) && examples.length < 3) examples.push(d);
    });
    if (examples.length) { problems.push({ column: h, examples: examples }); return; }

    if (step.numberCols.indexOf(h) === -1) return;
    let total = 0, zeros = 0, long = 0;
    const zeroExamples = [];
    rows.forEach(r => {
      const digits = String(r[c]).replace(/\D/g, "");
      if (digits.length < 7) return;
      total++;
      const m = digits.match(/0+$/);
      const n = m ? m[0].length : 0;
      if (n >= 3) {
        zeros++;
        if (n >= 5) long++;
        if (zeroExamples.length < 3) zeroExamples.push(digits);
      }
    });
    if ((zeros >= 2 && zeros / total >= 0.25) || (total < 8 && long >= 1)) {
      problems.push({ column: h, examples: zeroExamples });
    }
  });
  return problems;
}

/** Replaces one section of Raw Data (rows 7+) and stamps its date, like a paste on Raw Data does. */
function idm_writeSection_(raw, step, rows) {
  const width = step.headers.length;
  if (raw.getMaxColumns() < 61) raw.insertColumnsAfter(raw.getMaxColumns(), 61 - raw.getMaxColumns());
  const need = 6 + rows.length;
  if (raw.getMaxRows() < need) raw.insertRowsAfter(raw.getMaxRows(), need - raw.getMaxRows());

  const maxRows = raw.getMaxRows();
  idm_log_("clearing old rows 7-" + maxRows);
  raw.getRange(7, step.col, maxRows - 6, width).clearContent();
  raw.getRange(6, step.col, 1, width).setValues([step.headers]);
  fixRawDataNumberFormatting_(raw);
  idm_log_("setValues " + rows.length + " x " + width);
  raw.getRange(7, step.col, rows.length, width).setValues(rows);

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yy");
  raw.getRange(step.stampCell).setValue(step.stampPrefix + today);

  idm_log_("stamp set; repainting colors");
  forceVibrantColors_(raw);
  idm_log_("banding");
  applyBandingToRawData_(raw);
  lockRawDataRowHeights_(raw);
  SpreadsheetApp.flush();
}

// ---------------------------------------------------------------------
// SITE / CLASS STEP (filled in by hand, saved to Raw Data BH:BI)
// ---------------------------------------------------------------------

/** Current Raw Data BH:BI as class code -> { site, teacher } (same reading as the build). */
function idm_readSiteClass_(raw) {
  const map = new Map();
  if (raw.getMaxRows() < 7 || raw.getMaxColumns() < 61) return map;
  raw.getRange(7, 60, raw.getMaxRows() - 6, 2).getValues().forEach(r => {
    const site = String(r[0]).trim();
    const classes = String(r[1]).trim();
    if (!classes) return;
    parseClassCell_(classes).forEach(c => map.set(c.code, { site: site, teacher: c.teacher }));
  });
  return map;
}

/**
 * Class | Teacher Name | Site (Optional): one row per ATS class code (Raw Data R),
 * plus classes only in the current BH:BI, prefilled from BH:BI. The Site dropdown
 * reads a hidden list (column Z = every site typed in column C), so a new site is
 * offered in the dropdown as soon as it is typed; other values are allowed (warning only).
 */
function idm_fillSiteClassTab_(ss, sheet) {
  const raw = idm_rawSheet_(ss);
  const existing = raw ? idm_readSiteClass_(raw) : new Map();
  const codes = new Set();
  if (raw && raw.getMaxRows() >= 7) {
    raw.getRange(7, 18, raw.getMaxRows() - 6, 1).getValues().forEach(r => {
      const c = String(r[0]).trim().toUpperCase();
      if (c) codes.add(c);
    });
  }
  existing.forEach((v, code) => codes.add(code));
  const sorted = [...codes].sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
  idm_log_("site/class tab: " + sorted.length + " class codes (" + existing.size + " already in BH:BI)");

  idm_clearTab_(sheet);
  const rows = sorted.map(code => {
    const e = existing.get(code) || { site: "", teacher: "" };
    return [code, e.teacher, e.site];
  });
  const total = Math.max(rows.length + 50, 100); // room to add classes by hand
  if (sheet.getMaxRows() < total + 1) sheet.insertRowsAfter(sheet.getMaxRows(), total + 1 - sheet.getMaxRows());

  sheet.getRange(1, 1, 1, 3).setValues([["Class", "Teacher Name", "Site (Optional)"]])
    .setFontWeight("bold").setBackground("#d9ead3").setFontSize(11);
  if (rows.length) sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  sheet.getRange(2, 1, total, 1).setFontWeight("bold");
  sheet.getRange(1, 1, total + 1, 3).setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 150);

  // Hidden site list (column Z) that feeds the dropdown
  sheet.getRange("Z1").setValue("Site list (for the dropdown)");
  sheet.getRange("Z2").setFormula('=IFERROR(SORT(UNIQUE(FILTER(C2:C, C2:C<>""))), "")');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sheet.getRange("Z2:Z"), true)
    .setAllowInvalid(true)
    .setHelpText("Pick a site, or type a new one (it's added to this list).")
    .build();
  sheet.getRange(2, 3, total, 1).setDataValidation(rule).setHorizontalAlignment("right");
  try { sheet.hideColumns(26); } catch (e) { idm_log_("could not hide column Z: " + e.message); }
}

/** "✅ Save Site / Class": writes Raw Data BH:BI (rows 7+) and finishes the step. */
function idm_saveSiteClass() {
  return idm_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const state = idm_loadState_();
    if (!state) return idm_view_(ss, null);
    return idm_saveSiteClass_(ss, state);
  });
}

function idm_saveSiteClass_(ss, state) {
  {
    const step = IDM_STEPS[state.step];
    if (!step.manual) return idm_view_(ss, state);
    const temp = idm_ensureTempSheet_(ss, state);
    const raw = idm_rawSheet_(ss);
    if (!raw) return idm_view_(ss, state, { type: 'error', code: 'noRaw', message: 'The "Raw Data" tab is missing.', fromUpload: true });

    const last = temp.getLastRow();
    const values = last >= 2 ? temp.getRange(2, 1, last - 1, 3).getDisplayValues() : [];
    const bySite = new Map(); // site -> ["Y25: Carl", "Y24", ...] in the tab's order
    let classCount = 0;
    const bad = [];
    const seen = new Set();
    values.forEach(r => {
      const code = String(r[0]).trim().toUpperCase();
      const teacher = String(r[1]).replace(/:/g, " ").replace(/\s+/g, " ").trim(); // ":" / line breaks would break the cell
      const site = String(r[2]).trim();
      if (!code) return;
      if (!/^[A-Z0-9]+$/.test(code)) { bad.push(code); return; }
      if (seen.has(code)) return;
      seen.add(code);
      if (!teacher && !site) return; // nothing to save for this class
      if (!bySite.has(site)) bySite.set(site, []);
      bySite.get(site).push(teacher ? code + ": " + teacher : code);
      classCount++;
    });
    // One row per site (sites in number order, classes with no site last)
    const sites = [...bySite.keys()].sort((a, b) => (a === "") - (b === "") || a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
    const out = sites.map(site => [site, bySite.get(site).join("\n")]);
    idm_log_("site/class save: " + classCount + " classes in " + out.length + " site rows, " + bad.length + " bad codes");
    if (bad.length) {
      return idm_view_(ss, state, { type: 'error', code: 'badClass', fromUpload: true,
        message: "Class codes can only have letters and numbers (no spaces or dashes): <b>" + bad.slice(0, 5).join(", ") + "</b>. Fix them on the tab, then save again." });
    }

    if (raw.getMaxColumns() < 61) raw.insertColumnsAfter(raw.getMaxColumns(), 61 - raw.getMaxColumns());
    const need = 6 + out.length;
    if (raw.getMaxRows() < need) raw.insertRowsAfter(raw.getMaxRows(), need - raw.getMaxRows());
    raw.getRange(7, 60, raw.getMaxRows() - 6, 2).clearContent();
    raw.getRange(6, 60, 1, 2).setValues([["Site", "Class"]]);
    if (out.length) raw.getRange(7, 60, out.length, 2).setValues(out);
    forceVibrantColors_(raw);
    applyBandingToRawData_(raw);
    lockRawDataRowHeights_(raw); // one class per line in BI must not make rows tall
    SpreadsheetApp.flush();

    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yy");
    PropertiesService.getDocumentProperties().setProperty(IDM_SITECLASS_UPDATED_KEY, today);
    state.results[step.key] = { rows: classCount, saved: true };
    state.dataChanged = true; // step 5 builds first (idm_buildForLists)
    return idm_advance_(ss, state, { type: 'imported', message: "✅ Site / Class: " + classCount + " class(es) saved to Raw Data." });
  }
}

// ---------------------------------------------------------------------
// STEPS, TABS, HIDING
// ---------------------------------------------------------------------

/** Moves to the next step (new temp tab), or finishes. */
function idm_advance_(ss, state, notice) {
  const oldTemp = idm_sheetById_(ss, state.tempId);
  state.step++;
  idm_log_("advance to step " + state.step);
  if (state.step >= IDM_STEPS.length) {
    const results = state.results;
    idm_cleanup_(ss, state, true);
    return { active: false, finished: true, notice: notice, anyImported: idm_finalBuildNeeded_(state), builtAlready: idm_builtAlready_(state), guideNext: idm_guideNext_(ss), summary: idm_summary_(results) };
  }
  if (oldTemp) {
    // Same tab, cleared and renamed for the next export (never deleted while the user is on it)
    idm_renameTab_(ss, oldTemp, IDM_STEPS[state.step].tab);
    idm_prepareTab_(ss, oldTemp, IDM_STEPS[state.step]);
    oldTemp.activate();
    idm_log_("temp tab reused as " + oldTemp.getName());
  } else {
    state.tempId = null;
    idm_ensureTempSheet_(ss, state);
  }
  idm_saveState_(state);
  return idm_view_(ss, state, notice);
}

/** The current step's temp tab; (re)created if missing. */
function idm_ensureTempSheet_(ss, state) {
  let sheet = idm_sheetById_(ss, state.tempId);
  idm_log_("temp tab id " + state.tempId + (sheet ? " found" : " missing -> creating"));
  if (sheet) {
    if (sheet.isSheetHidden()) sheet.showSheet();
    return sheet;
  }
  const step = IDM_STEPS[state.step];
  sheet = ss.insertSheet(0);
  idm_renameTab_(ss, sheet, step.tab);
  if (sheet.getMaxColumns() < 40) sheet.insertColumnsAfter(sheet.getMaxColumns(), 40 - sheet.getMaxColumns());
  sheet.setTabColor("#f1c232");
  idm_prepareTab_(ss, sheet, step);
  sheet.activate();
  state.tempId = sheet.getSheetId();
  return sheet;
}

/** "Student Data", or "Student Data (2)" if the user already has a tab with that name. */
function idm_uniqueTabName_(ss, base, self) {
  let name = base, n = 2;
  for (;;) {
    const other = ss.getSheetByName(name);
    if (!other || (self && other.getSheetId() === self.getSheetId())) return name;
    name = base + " (" + (n++) + ")";
  }
}

/** Renames the temp tab (falls back to "Site & Class" if "/" isn't allowed in a tab name). */
function idm_renameTab_(ss, sheet, base) {
  try {
    sheet.setName(idm_uniqueTabName_(ss, base, sheet));
  } catch (e) {
    idm_log_("tab name '" + base + "' refused (" + e.message + ")");
    sheet.setName(idm_uniqueTabName_(ss, base.replace(/\s*\/\s*/g, " & "), sheet));
  }
}

/** Paste tab ("PASTE HERE") or the Site / Class form. */
function idm_prepareTab_(ss, sheet, step) {
  if (step.manual) idm_fillSiteClassTab_(ss, sheet);
  else if (step.lists) idm_fillListsTab_(ss, sheet);
  else idm_resetTempSheet_(sheet);
}

/** Step 5's tab: just a list of your lists (they are opened from the sidebar). */
function idm_fillListsTab_(ss, sheet) {
  idm_clearTab_(sheet);
  const lists = od_listSummaries_(ss);
  const rows = [["Your Lists"], ["Add a list or open one in the sidebar. Each list has its own tab."], [""],
                ["List", "Tab", "By", "Rows", "Matched", "Need you", "Updated"]];
  lists.forEach(l => rows.push([l.name, l.tabName, l.by === "guardian" ? "guardian" : "student",
    l.stats ? l.stats.rows : "", l.stats ? l.stats.matched : "", l.stats ? l.stats.review : "", l.updated || ""]));
  const width = 7;
  sheet.getRange(1, 1, rows.length, width).setValues(rows.map(r => padRow_(r, width)));
  sheet.getRange("A1").setFontWeight("bold").setFontSize(14);
  sheet.getRange(4, 1, 1, width).setFontWeight("bold").setBackground("#d0e0e3");
  sheet.setColumnWidth(1, 160);
}

/**
 * Empties the temp tab completely before the next step uses it: values, formats,
 * EVERY dropdown (the Site / Class step puts them on 100+ rows, more than its data,
 * so clearing only the data range left some behind), the hidden site-list column Z
 * and the frozen header row.
 */
function idm_clearTab_(sheet) {
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
  sheet.clear();
  try { sheet.showColumns(1, sheet.getMaxColumns()); } catch (e) { /* none hidden */ }
  try { sheet.setFrozenRows(0); } catch (e) { /* cosmetic */ }
}

function idm_resetTempSheet_(sheet) {
  idm_clearTab_(sheet);
  sheet.getRange("A1").setValue(IDM_PLACEHOLDER)
    .setBackground("#FFF2CC").setFontWeight("bold").setFontSize(14)
    .setNote("Paste anywhere on this tab — it doesn't have to be A1.");
  sheet.setColumnWidth(1, 160);
}

function idm_tempHasData_(sheet) {
  if (sheet.getLastRow() > 1 || sheet.getLastColumn() > 1) return true;
  return String(sheet.getRange("A1").getValue()).trim() !== IDM_PLACEHOLDER;
}

/** Hides every visible tab except the temp tab and remembers which ones they were. */
function idm_hideOtherSheets_(ss, state) {
  const hidden = [];
  ss.getSheets().forEach(s => {
    if (s.getSheetId() === state.tempId || s.isSheetHidden()) return;
    try { s.hideSheet(); hidden.push(s.getSheetId()); } catch (e) { console.warn("Input Data Mode could not hide " + s.getName() + ": " + e.message); }
  });
  state.hiddenIds = hidden;
}

/** Shows the tabs we hid, deletes the temp tab, clears the saved state. */
function idm_cleanup_(ss, state, goToRawData) {
  // Your Lists tabs stay hidden once the mode ends (owner's request 2026-11-12): they are opened
  // from step 5 (od_openList shows the tab again). Hidden after the others are back, so a tab still shows.
  let listIds = [];
  try { listIds = od_loadLists_(ss).filter(l => l.sheet).map(l => l.sheet.getSheetId()); }
  catch (e) { console.warn("Your Lists tabs not hidden: " + e.message); }
  (state.hiddenIds || []).forEach(id => {
    if (listIds.indexOf(id) !== -1) return;
    const s = idm_sheetById_(ss, id);
    if (s) { try { s.showSheet(); } catch (e) { console.warn(e.message); } }
  });
  listIds.forEach(id => {
    const s = idm_sheetById_(ss, id);
    if (s && !s.isSheetHidden()) { try { s.hideSheet(); } catch (e) { console.warn("Could not hide " + s.getName() + ": " + e.message); } }
  });
  const raw = idm_rawSheet_(ss);
  const target = (goToRawData && raw && !raw.isSheetHidden()) ? raw : (idm_sheetById_(ss, state.prevActiveId) || raw);
  const temp = idm_sheetById_(ss, state.tempId);
  let landing = (target && !target.isSheetHidden() && (!temp || target.getSheetId() !== temp.getSheetId())) ? target : null;
  if (!landing) landing = ss.getSheets().filter(s => !s.isSheetHidden() && (!temp || s.getSheetId() !== temp.getSheetId()))[0] || null;
  if (landing) landing.activate();
  const props = PropertiesService.getDocumentProperties();
  // Deleted by the sidebar's next call (idm_deleteLeftoverTab), once the user is off it
  if (temp && landing) props.setProperty(IDM_PENDING_DELETE_KEY, String(temp.getSheetId()));
  props.deleteProperty(IDM_STATE_KEY);
}

// ---------------------------------------------------------------------
// SMALL HELPERS
// ---------------------------------------------------------------------

function idm_view_(ss, state, notice) {
  if (!state) return { active: false, notice: notice || null };
  const raw = idm_rawSheet_(ss);
  const step = IDM_STEPS[state.step];
  const temp = idm_sheetById_(ss, state.tempId);
  return {
    active: true,
    stepIndex: state.step,
    stepCount: IDM_STEPS.length,
    step: { key: step.key, zone: step.zone, required: step.required, howTo: step.howTo, manual: !!step.manual, lists: !!step.lists },
    // Step 5: your lists (+ whether the sidebar must build first, and how that build went)
    lists: step.lists ? od_listSummaries_(ss) : undefined,
    needsBuild: step.lists ? idm_listsNeedBuild_(ss, state) : false,
    listsBuild: step.lists ? (state.listsBuild || null) : null,
    tabName: temp ? temp.getName() : step.tab,
    lastUpdated: raw ? idm_lastUpdated_(raw, step) : null,
    // Finished / skipped steps (other than this one), for the "✅ "Student Data" Successful" lines under Option 1
    done: IDM_STEPS.map((st, i) => ({ st: st, i: i, r: (state.results || {})[st.key] }))
      .filter(x => x.r && x.i !== state.step)
      .map(x => ({ name: x.st.tab, rows: x.r.rows || 0, skipped: !x.r.rows && !x.r.saved })),
    // Progress bar at the top of the sidebar
    progress: IDM_STEPS.map((st, i) => {
      const r = (state.results || {})[st.key];
      return { label: st.short, current: i === state.step,
               status: r ? ((r.rows || r.saved) ? 'done' : 'skipped') : 'todo' };
    }),
    notice: notice || null
  };
}

/** "MM/dd/yy" from the section's DATE UPDATED stamp, plus how many days ago (null = never). */
function idm_lastUpdated_(raw, step) {
  if (step.lists) {
    const dates = od_listSummaries_(SpreadsheetApp.getActiveSpreadsheet())
      .map(l => idm_dateInfo_(l.updated)).filter(Boolean).sort((a, b) => (a.days === null ? 1e9 : a.days) - (b.days === null ? 1e9 : b.days));
    return dates[0] || null;
  }
  if (step.manual) {
    const saved = PropertiesService.getDocumentProperties().getProperty(IDM_SITECLASS_UPDATED_KEY);
    return saved ? idm_dateInfo_(saved) : null;
  }
  const text = String(raw.getRange(step.stampCell).getDisplayValue());
  if (/00\/00\/00/.test(text)) return null;
  return idm_dateInfo_(text);
}

/** { text: "MM/dd/yy", days } from the first date in a string (null if none). */
function idm_dateInfo_(text) {
  const m = String(text).match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (!m) return null;
  let year = Number(m[3]);
  if (year < 100) year += 2000;
  const date = new Date(year, Number(m[1]) - 1, Number(m[2]));
  if (isNaN(date.getTime())) return { text: m[0], days: null };
  const days = Math.max(0, Math.floor((Date.now() - date.getTime()) / 86400000));
  return { text: m[0], days: days };
}

function idm_summary_(results) {
  return IDM_STEPS.map(step => {
    const r = (results || {})[step.key];
    if (r && r.lists) return "✅ " + step.zone + ": " + r.rows + " list(s) set up";
    if (r && r.saved) return "✅ " + step.zone + ": " + r.rows + " class(es) saved";
    if (r && r.rows) return "✅ " + step.zone + ": " + r.rows + " row(s) imported";
    if (r && r.skipped) return "⏭️ " + step.zone + ": skipped (old data kept)";
    return "— " + step.zone + ": not loaded (old data kept)";
  });
}

function idm_anyImported_(results) {
  return IDM_STEPS.some(step => (results || {})[step.key] && (results[step.key].rows > 0 || results[step.key].saved));
}

function idm_parseCsv_(text) {
  text = String(text || "").replace(/^﻿/, "");
  const firstLine = text.split(/\r?\n/)[0] || "";
  const delim = (firstLine.indexOf("\t") !== -1 && firstLine.indexOf(",") === -1) ? "\t" : ",";
  return Utilities.parseCsv(text, delim);
}

function idm_norm_(v) {
  return String(v).toLowerCase().replace(/[^a-z0-9]/g, '');
}

function idm_rawSheet_(ss) {
  return ss.getSheetByName("Raw Data") || ss.getSheetByName("RAW Data");
}

function idm_sheetById_(ss, id) {
  if (id === null || id === undefined) return null;
  return ss.getSheets().filter(s => s.getSheetId() === id)[0] || null;
}

function idm_loadState_() {
  const raw = PropertiesService.getDocumentProperties().getProperty(IDM_STATE_KEY);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function idm_saveState_(state) {
  PropertiesService.getDocumentProperties().setProperty(IDM_STATE_KEY, JSON.stringify(state));
}

function idm_withLock_(fn) {
  IDM_DBG = idm_debugOn_() ? [] : null;
  if (IDM_DBG) IDM_DBG.t0 = Date.now();
  let result;
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) {
    idm_log_("could not get the document lock in 20s");
    result = idm_errorView_(new Error("The sheet is busy — try again in a moment."));
  } else {
    try {
      idm_log_("lock ok");
      result = fn();
    } catch (e) {
      idm_log_("EXCEPTION: " + (e && e.stack ? e.stack : e));
      result = idm_errorView_(e);
    } finally {
      lock.releaseLock();
    }
  }
  if (IDM_DBG && result && typeof result === "object") {
    idm_log_("done");
    result.debug = IDM_DBG.slice();
  }
  IDM_DBG = null;
  return result;
}

/** A script error as a sidebar error (instead of a silent retry). */
function idm_errorView_(e) {
  const notice = { type: 'error', code: 'exception', message: "❌ Script error: " + (e && e.message ? e.message : e), fromUpload: true };
  try {
    return idm_view_(SpreadsheetApp.getActiveSpreadsheet(), idm_loadState_(), notice);
  } catch (e2) {
    idm_log_("idm_view_ also failed: " + e2.message);
    return { active: true, stepIndex: 0, stepCount: IDM_STEPS.length,
             step: { key: '', zone: '(unknown)', required: false, howTo: [] }, tabName: '?', lastUpdated: null, notice: notice };
  }
}

// ======================================================================
// PART 2: YOUR LISTS (step 5 of 📥 Input Data Mode; YourListsSidebar.html)
// Any outside list (bus routes, trip forms, after school, PTA ...) is pasted
// on its own tab ("📥 Bus Routes"). The tool finds the header row, guesses what
// each column is, matches every row to a student (student lists) or a guardian
// (guardian lists), and the columns you keep ("variables") go to:
//   - the Backend tab's ADDITIONAL INFORMATION section (OSIS / Student /
//     Guardian / lines), which the phone contacts show as "Additional
//     Information:" above Communications (readAdditionalInfo_, Build.gs);
//   - the Student List (student lists) or Parents Divided (guardian lists),
//     one column per variable.
// Everyone on a list gets it (owner's pick "A"); Combined Contact Tracking is
// not changed.
//
// List tab layout: row 1 title band, row 2 status band, row 3 what each column
// is, row 4 the list's header row (a paste is moved there, title rows above the
// headers dropped), data from row 5, then a spacer and the "✔ Matched" / "How"
// columns the tool fills. The setup (columns, variables, your picks) is one JSON
// cell per list in the Backend's YOUR LISTS (SETUP) section.
//
// Every build (od_buildOutputs_, called by buildAllDerivedSheets in a
// try/catch) matches every list again, so new students, sorting or a fixed name
// are picked up; rows only a person can decide (two possible students, a best
// guess you haven't accepted, not found) stay unmatched until you decide in the
// sidebar. Your decisions are kept per row (by what the row says) for the next paste.
// ======================================================================

var OD_TAB_PREFIX = "📥 ";
var OD_PLACEHOLDER = "PASTE HERE";
var OD_HEADER_ROW = 4;               // a list's header row after a paste; data from row 5
var OD_WIZARD_KEY = "OD_WIZARD_TAB"; // tab id the setup sidebar works on
var OD_BAND_TITLE = "📥 OWN DATA · ";
var OD_TOOL_HEADERS = ["✔ Matched", "How"];
var OD_COLORS = {
  band1: "#134f5c", band2: "#2f6b77", chips: "#fbfcfe", header: "#d0e0e3", paste: "#FFF2CC",
  ok: "#e6f4ea", guess: "#fef3d6", none: "#fce8e6", skip: "#f1f3f4"
};
var OD_HEADER_WORDS_ = /osis|student|name|first|last|parent|guardian|mother|father|phone|cell|mobile|e-?mail|class|grade|\bid\b|route|bus|stop|date|address|site|room|teacher/i;

// ======================================================================
// SETUP STORAGE (Backend "YOUR LISTS (SETUP)": Tab ID | List | Setup | Updated)
// ======================================================================

function od_emptySetup_() {
  return { v: 1, ready: false, by: "student", roles: {}, vars: [], picks: {}, headers: [], title: "", stats: null };
}

function od_fillSetup_(s) {
  const e = od_emptySetup_();
  s = s && typeof s === "object" ? s : {};
  Object.keys(e).forEach(k => { if (s[k] === undefined) s[k] = e[k]; });
  if (!s.picks || typeof s.picks !== "object") s.picks = {};
  if (!Array.isArray(s.vars)) s.vars = [];
  return s;
}

/** Every list: { tabId, name, setup, updated, sheet } (sheet = null when its tab is gone). */
function od_loadLists_(ss) {
  return backendReadBlock_(ss, "OWN_LISTS", true).slice(1)
    .filter(r => String(r[0]).trim() !== "")
    .map(r => {
      let setup = null;
      try { setup = JSON.parse(String(r[2] || "{}")); } catch (e) { setup = null; }
      const tabId = Number(r[0]);
      return { tabId: tabId, name: toText_(r[1]), setup: od_fillSetup_(setup), updated: toText_(r[3]), sheet: od_sheetById_(ss, tabId) };
    });
}

/** Writes the lists back (a list whose tab was deleted by hand is dropped). */
function od_saveLists_(ss, lists) {
  const rows = lists.filter(l => l.sheet).map(l => {
    let json = JSON.stringify(l.setup);
    // A cell holds 50,000 characters: forget the oldest decisions first if a list ever gets there
    while (json.length > 45000 && Object.keys(l.setup.picks).length) {
      delete l.setup.picks[Object.keys(l.setup.picks)[0]];
      json = JSON.stringify(l.setup);
    }
    return [l.tabId, l.name, json, l.updated || ""];
  });
  backendWriteBlock_(ss, "OWN_LISTS", [["Tab ID", "List", "Setup", "Updated"]].concat(rows));
}

function od_findList_(lists, tabId) {
  return lists.filter(l => l.tabId === Number(tabId))[0] || null;
}

// ======================================================================
// LIST TABS
// ======================================================================

function od_sheetById_(ss, id) {
  if (id === null || id === undefined || id === "" || isNaN(Number(id))) return null;
  return ss.getSheets().filter(s => s.getSheetId() === Number(id))[0] || null;
}

function od_uniqueTabName_(ss, base, self) {
  let name = base, n = 2;
  for (;;) {
    const other = ss.getSheetByName(name);
    if (!other || (self && other.getSheetId() === self.getSheetId())) return name;
    name = base + " (" + (n++) + ")";
  }
}

function od_cleanName_(name) {
  return String(name || "").replace(/[\[\]:*?\/\\'"]/g, " ").replace(/\s+/g, " ").trim().slice(0, 60);
}

/** A variable's name ("Bus Route"): one line, no ":" (the line reads "Name: value"); quotes are fine. */
function od_cleanLabel_(label) {
  return String(label || "").replace(/[:\r\n\t]/g, " ").replace(/\s+/g, " ").trim().slice(0, 60);
}

function od_colLetter_(c) {
  let s = "";
  while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
  return s;
}

function od_today_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yy");
}

/** Rows 1-3 (title, status, what each column is). */
function od_writeBands_(sheet, name, updated, status, chips) {
  const width = Math.max(chips.length, 8);
  if (sheet.getMaxColumns() < width) sheet.insertColumnsAfter(sheet.getMaxColumns(), width - sheet.getMaxColumns());
  const row1 = new Array(width).fill(""), row2 = new Array(width).fill(""), row3 = new Array(width).fill("");
  row1[0] = OD_BAND_TITLE + String(name).toUpperCase() + " — DATE UPDATED: " + (updated || "00/00/00");
  row2[0] = status;
  chips.forEach((c, i) => { row3[i] = c || ""; });
  sheet.getRange(1, 1, 3, width).setValues([row1, row2, row3]);
  sheet.getRange(1, 1, 1, width).setBackground(OD_COLORS.band1).setFontColor("#ffffff").setFontWeight("bold");
  sheet.getRange(2, 1, 1, width).setBackground(OD_COLORS.band2).setFontColor("#ffffff").setFontWeight("normal");
  sheet.getRange(3, 1, 1, width).setBackground(OD_COLORS.chips).setFontColor("#174ea6").setFontWeight("bold").setFontSize(9);
}

/** An empty list tab: bands + the yellow PASTE HERE box in A4. */
function od_resetListTab_(sheet, name) {
  try { sheet.setFrozenRows(0); } catch (e) { /* none */ }
  sheet.getDataRange().clearDataValidations();
  sheet.clear();
  od_writeBands_(sheet, name, "", "Paste your whole list (headers too) in the yellow box (A4) as values: Ctrl+Shift+V, or upload the file in the sidebar. The sidebar sets it up.", []);
  sheet.getRange(OD_HEADER_ROW, 1).setValue(OD_PLACEHOLDER).setBackground(OD_COLORS.paste).setFontWeight("bold")
    .setNote("Paste your whole list here (headers too). Title rows above the headers are fine.");
  sheet.setColumnWidth(1, 160);
  try { sheet.setTabColor(OD_COLORS.band1); } catch (e) { /* cosmetic */ }
}

/** Index of the first row that is part of the paste (rows 1-3 are ours while the title band is still in A1). */
function od_pasteStart_(values) {
  const bandIntact = values.length > 0 && String(values[0][0]).indexOf(OD_BAND_TITLE) === 0;
  return bandIntact ? OD_HEADER_ROW - 1 : 0;
}

function od_hasPaste_(sheet) {
  const values = sheet.getDataRange().getValues();
  const start = od_pasteStart_(values);
  for (let r = start; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const v = values[r][c];
      if (v === "" || v === null) continue;
      if (r === OD_HEADER_ROW - 1 && c === 0 && String(v).trim() === OD_PLACEHOLDER) continue;
      return true;
    }
  }
  return false;
}

/**
 * Header row of a pasted block (display values): among the first 15 rows, the
 * first clearly header-like one (2+ cells, mostly words, about as wide as the
 * table, column-name words like "Name" / "OSIS" / "Phone" help). Title rows
 * ("District 75 ... Fall Routes", "Printed 11/03/26") have one cell and are skipped.
 */
function od_findHeaderRow_(d) {
  const width = r => (d[r] || []).filter(x => String(x).trim() !== "").length;
  let maxWidth = 0;
  for (let r = 0; r < Math.min(30, d.length); r++) maxWidth = Math.max(maxWidth, width(r));
  let best = 0, bestScore = -1;
  for (let r = 0; r < Math.min(15, d.length); r++) {
    const cells = d[r].map(x => String(x).trim()).filter(Boolean);
    if (!cells.length || (cells.length < 2 && maxWidth >= 2)) continue;
    const words = cells.filter(x => !/^[\d\s.,:\/()+#$%-]+$/.test(x)).length / cells.length;
    const wide = cells.length / maxWidth;
    if (words < 0.6 || wide < 0.5) continue;
    const keywords = Math.min(4, cells.filter(x => OD_HEADER_WORDS_.test(x)).length);
    const score = wide + words + keywords * 0.5 + (width(r + 1) ? 0.5 : 0);
    if (score > bestScore + 0.25) { best = r; bestScore = score; }
  }
  return best;
}

/**
 * Moves what was pasted (anywhere on the tab) into the fixed layout: header row
 * on row 4, data rows under it (empty rows dropped), title rows above the
 * header remembered as the list's title. Cells are written as the text you saw
 * (plain text), so a leading zero or a phone format is kept.
 * Returns { headers, rows, title } or null when the tab is empty.
 */
function od_takePaste_(sheet, name) {
  const range = sheet.getDataRange();
  const values = range.getValues(), displays = range.getDisplayValues();
  const start = od_pasteStart_(values);
  const block = [];
  for (let r = start; r < displays.length; r++) {
    const row = displays[r].map(x => String(x));
    if (r === OD_HEADER_ROW - 1 && row[0].trim() === OD_PLACEHOLDER) row[0] = "";
    block.push(row);
  }
  let minC = Infinity, maxC = -1, minR = Infinity, maxR = -1;
  block.forEach((row, r) => row.forEach((v, c) => {
    if (v.trim() === "") return;
    minC = Math.min(minC, c); maxC = Math.max(maxC, c); minR = Math.min(minR, r); maxR = Math.max(maxR, r);
  }));
  if (maxC < 0) return null;
  const d = block.slice(minR, maxR + 1).map(r => r.slice(minC, maxC + 1).map(x => x.trim()));
  const h = od_findHeaderRow_(d);
  const title = d.slice(0, h).map(r => r.filter(Boolean).join(" ")).filter(Boolean).join(" · ");
  const headers = d[h].map((x, i) => x || ("Column " + od_colLetter_(i + 1)));
  const rows = d.slice(h + 1).filter(r => r.some(x => x !== ""));

  sheet.getDataRange().clearDataValidations();
  sheet.clear();
  const width = headers.length;
  const needRows = OD_HEADER_ROW + rows.length + 1, needCols = width + 3;
  if (sheet.getMaxRows() < needRows) sheet.insertRowsAfter(sheet.getMaxRows(), needRows - sheet.getMaxRows());
  if (sheet.getMaxColumns() < needCols) sheet.insertColumnsAfter(sheet.getMaxColumns(), needCols - sheet.getMaxColumns());
  const table = [headers].concat(rows);
  sheet.getRange(OD_HEADER_ROW, 1, table.length, width).setNumberFormat("@").setValues(table);
  sheet.getRange(OD_HEADER_ROW, 1, 1, width).setFontWeight("bold").setBackground(OD_COLORS.header);
  try { sheet.setFrozenRows(OD_HEADER_ROW); } catch (e) { /* cosmetic */ }
  od_writeBands_(sheet, name, od_today_(), "Pasted: " + rows.length + " row(s). Finish the setup in the sidebar.", headers.map(() => ""));
  return { headers: headers, rows: rows, title: title };
}

/** The list's table as it is on the tab now: { headers, rows } (display text; tool columns left out). */
function od_readTable_(sheet, setup) {
  const lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (lastRow < OD_HEADER_ROW || lastCol < 1) return { headers: [], rows: [] };
  const width = Math.min(lastCol, setup && setup.headers && setup.headers.length ? setup.headers.length : lastCol);
  const d = sheet.getRange(OD_HEADER_ROW, 1, lastRow - OD_HEADER_ROW + 1, width).getDisplayValues();
  let headers = d[0].map(x => String(x).trim());
  const tool = headers.indexOf(OD_TOOL_HEADERS[0]);
  if (tool !== -1) headers = headers.slice(0, Math.max(0, tool - 1)); // spacer before the tool columns
  const rows = d.slice(1).map(r => r.slice(0, headers.length).map(x => String(x).trim()));
  return { headers: headers, rows: rows };
}

// ======================================================================
// STUDENTS / GUARDIANS INDEX (what a row can be matched to)
// ======================================================================

/** "SIMPSON, BART" / "Bart Simpson" -> { raw, tokens, first, last } (lowercase parts), or null. */
function od_nameParts_(text) {
  const s = toText_(text);
  if (!s) return null;
  let firstT, lastT;
  if (s.indexOf(",") !== -1) {
    const i = s.indexOf(",");
    lastT = nameTokens_(s.slice(0, i));
    firstT = nameTokens_(s.slice(i + 1));
  } else {
    const t = nameTokens_(s);
    firstT = t.slice(0, 1);
    lastT = t.slice(1);
  }
  const tokens = firstT.concat(lastT);
  if (!tokens.length) return null;
  return { raw: s, tokens: tokens, first: firstT[0] || "", last: lastT.length ? lastT[lastT.length - 1] : "",
           firstTokens: firstT, lastTokens: lastT };
}

/** Name from two columns (first name, last name). */
function od_namePartsTwo_(firstText, lastText) {
  const firstT = nameTokens_(firstText), lastT = nameTokens_(lastText);
  const tokens = firstT.concat(lastT);
  if (!tokens.length) return null;
  return { raw: (toText_(firstText) + " " + toText_(lastText)).trim(), tokens: tokens, first: firstT[0] || "", last: lastT.length ? lastT[lastT.length - 1] : "",
           firstTokens: firstT, lastTokens: lastT };
}

function od_tokenKey_(tokens) {
  return tokens.slice().sort().join(" ");
}

function od_addTo_(map, key, item) {
  if (!key) return;
  if (!map.has(key)) map.set(key, []);
  const list = map.get(key);
  if (list.indexOf(item) === -1) list.push(item);
}

/**
 * studentRows: Student List rows (OSIS, Student, Guardian, Site, Class, ...);
 * guardianRows: buildGuardianObjects_ output (Raw Data + overrides) for phones
 * and emails. Students who aren't in Raw Data (Discharged, the owner's Old
 * Master Table) get their guardians from the Student List's Guardian text.
 */
function od_index_(studentRows, guardianRows) {
  const idx = {
    byOsis: new Map(), students: [], byNameKey: new Map(), byToken: new Map(), byFirst: new Map(), byLast: new Map(),
    persons: [], pByKey: new Map(), pByNameKey: new Map(), pByToken: new Map(), pByLast: new Map(), byPhone: new Map(), byEmail: new Map(), classes: new Set(),
    byInitial: new Map(), pByInitial: new Map(),        // first letter of any name part -> students / guardians
    byLastInitial: new Map(), pByLastInitial: new Map() // first letter of a LAST name part (who to score for a full name)
  };
  const addInitials = (map, tokens, item) => new Set(tokens.map(t => t.charAt(0))).forEach(ch => od_addTo_(map, ch, item));
  studentRows.forEach(r => {
    const osis = toText_(r[0]);
    if (!osis || osis.indexOf(",") !== -1 || !isRealStudentId_(osis)) return;
    const key = normalizeOsis_(osis);
    if (idx.byOsis.has(key)) return;
    const parts = od_nameParts_(r[1]) || { tokens: [], first: "", last: "", firstTokens: [], lastTokens: [] };
    const st = {
      osis: osis, key: key, name: toText_(r[1]), tokens: parts.tokens, first: parts.first, last: parts.last,
      firstTokens: parts.firstTokens, lastTokens: parts.lastTokens,
      cls: toText_(r[4]).split(/[:\/]/)[0].trim().toUpperCase(), guardianText: toText_(r[2]), persons: []
    };
    idx.byOsis.set(key, st);
    idx.students.push(st);
    od_addTo_(idx.byNameKey, od_tokenKey_(st.tokens), st);
    st.tokens.forEach(t => od_addTo_(idx.byToken, t, st));
    od_addTo_(idx.byFirst, st.first, st);
    od_addTo_(idx.byLast, st.last, st);
    addInitials(idx.byInitial, st.tokens, st);
    addInitials(idx.byLastInitial, st.lastTokens.length ? st.lastTokens : st.tokens, st);
    if (st.cls) idx.classes.add(st.cls);
  });

  const addPerson = (st, name, phone, email) => {
    const parts = od_nameParts_(name);
    if (!parts) return;
    const nameKey = od_tokenKey_(parts.tokens);
    // Same name = same person, unless both have phones / emails and none of them agree
    let person = (idx.pByNameKey.get(nameKey) || []).filter(p => {
      const phonesDiffer = phone && p.phones.size && !p.phones.has(phone);
      const emailsDiffer = email && p.emails.size && !p.emails.has(email);
      const phonesSame = phone && p.phones.has(phone), emailsSame = email && p.emails.has(email);
      return phonesSame || emailsSame || !(phonesDiffer || emailsDiffer);
    })[0];
    if (!person) {
      const n = (idx.pByNameKey.get(nameKey) || []).length;
      person = { key: nameKey + (n ? "#" + (n + 1) : ""), name: name, tokens: parts.tokens, first: parts.first, last: parts.last,
                 firstTokens: parts.firstTokens, lastTokens: parts.lastTokens, phones: new Set(), emails: new Set(), kids: [] };
      idx.persons.push(person);
      idx.pByKey.set(person.key, person);
      od_addTo_(idx.pByNameKey, nameKey, person);
      od_addTo_(idx.pByLast, person.last, person);
      parts.tokens.forEach(t => od_addTo_(idx.pByToken, t, person));
      addInitials(idx.pByInitial, parts.tokens, person);
      addInitials(idx.pByLastInitial, parts.lastTokens.length ? parts.lastTokens : parts.tokens, person);
    }
    if (phone) { person.phones.add(phone); od_addTo_(idx.byPhone, phone, person); }
    if (email) { person.emails.add(email); od_addTo_(idx.byEmail, email, person); }
    if (person.kids.indexOf(st) === -1) person.kids.push(st);
    if (st.persons.indexOf(person) === -1) st.persons.push(person);
  };

  guardianRows.forEach(g => {
    const st = idx.byOsis.get(normalizeOsis_(g.student_id));
    if (!st) return;
    addPerson(st, stripTrailingParenthetical_(g.guardianOutput), normalizePhoneForMatch_(g.mobile), normalizeEmail_(g.email));
  });
  idx.students.forEach(st => {
    if (st.persons.length || !st.guardianText) return;
    st.guardianText.split(/\s+and\s+|,/i).map(x => stripTrailingParenthetical_(x).trim()).filter(Boolean)
      .forEach(name => addPerson(st, name, "", ""));
  });
  return idx;
}

/** Student List rows (OSIS..Class, the Backend copy; od_index_ skips its sibling-group rows) + Raw Data guardians, for the sidebar (outside a build). */
function od_indexFromSheet_(ss) {
  const studentRows = studentLookupRows_(ss, 5);
  const raw = ss.getSheetByName("Raw Data") || ss.getSheetByName("RAW Data");
  let guardianRows = [];
  if (raw && raw.getLastRow() > 1) {
    const overrides = getOverrideData_(ss);
    guardianRows = buildGuardianObjects_(raw.getDataRange().getValues(), overrides.guardians);
  }
  return od_index_(studentRows, guardianRows);
}

// ======================================================================
// MATCHING
// ======================================================================

// Common short names (both ways). "Sam" / "Mira" style short names that START
// the full name are caught without this list (od_firstLike_).
var OD_NICKNAMES_ = {
  william: "will bill billy liam", robert: "rob bob bobby robbie", richard: "rich rick ricky dick", james: "jim jimmy jamie",
  john: "jack johnny", jonathan: "jon jonny", joseph: "joe joey", michael: "mike mikey mick", christopher: "chris topher",
  christina: "chris tina", christine: "chris tina", elizabeth: "liz lizzie beth betsy eliza lisa", katherine: "kate katie kathy kat",
  catherine: "cathy cat kate", margaret: "maggie meg peggy", jennifer: "jen jenny", jessica: "jess jessie", alexander: "alex xander",
  alexandra: "alex lexi sandra", anthony: "tony", benjamin: "ben benny", daniel: "dan danny", david: "dave davey", edward: "ed eddie ted",
  matthew: "matt", nicholas: "nick nicky", patrick: "pat", samuel: "sam sammy", samantha: "sam sammy", stephen: "steve", steven: "steve",
  thomas: "tom tommy", timothy: "tim timmy", victoria: "vicky tori", zachary: "zach zack", gabriel: "gabe", abigail: "abby",
  rebecca: "becky", deborah: "debbie deb", susan: "sue susie", jacqueline: "jackie", francisco: "paco pancho", jose: "pepe",
  guillermo: "memo", alejandro: "alex", eduardo: "lalo eddie", gerardo: "jerry"
};
var OD_NICK_LOOKUP_ = null;
function od_nicknames_() {
  if (OD_NICK_LOOKUP_) return OD_NICK_LOOKUP_;
  OD_NICK_LOOKUP_ = {};
  Object.keys(OD_NICKNAMES_).forEach(full => {
    const group = [full].concat(OD_NICKNAMES_[full].split(" "));
    group.forEach(a => { OD_NICK_LOOKUP_[a] = OD_NICK_LOOKUP_[a] || {}; group.forEach(b => { if (a !== b) OD_NICK_LOOKUP_[a][b] = full; }); });
  });
  return OD_NICK_LOOKUP_;
}

/** How two first names are alike: "same", "start" (Sam / Samuel), "nickname" (Bill / William), "initial" (D / Diana) or "". */
function od_firstLike_(a, b) {
  if (!a || !b) return "";
  if (a === b) return "same";
  if ((a.length === 1 && b.indexOf(a) === 0) || (b.length === 1 && a.indexOf(b) === 0)) return "initial";
  const short = a.length <= b.length ? a : b, long = short === a ? b : a;
  if (short.length >= 3 && long.indexOf(short) === 0) return "start";
  const nick = od_nicknames_();
  if (nick[a] && nick[a][b]) return "nickname";
  return "";
}

function od_lastLike_(a, b) {
  if (!a || !b) return "";
  if (a === b) return "same";
  if ((a.length === 1 && b.indexOf(a) === 0) || (b.length === 1 && a.indexOf(b) === 0)) return "initial";
  return "";
}

/** Why a short first name is a best guess, in words: "“Mira” is how “Mirabel” starts". */
function od_firstWhy_(how, rowFirst, name) {
  const full = String(name || "").split(/\s+/)[0];
  const cap = s => s ? s.charAt(0).toUpperCase() + s.slice(1) : s;
  if (how === "start") return "“" + cap(rowFirst) + "” is how “" + full + "” starts";
  if (how === "nickname") return "“" + cap(rowFirst) + "” is a short name for “" + full + "”";
  if (how === "initial") return "“" + cap(rowFirst) + ".” could be “" + full + "”";
  return "";
}

/** The identity cells of one list row, already split into name parts / phone digits. */
function od_rowCells_(row, roles) {
  const at = c => (c ? toText_(row[c - 1]) : "");
  const name = r => {
    if (!r) return null;
    if (r.first || r.last) return od_namePartsTwo_(at(r.first), at(r.last));
    return od_nameParts_(at(r.col));
  };
  return {
    osis: at(roles.osis),
    name: name(roles.student),
    guardian: name(roles.guardian),
    phone: normalizePhoneForMatch_(at(roles.phone)),
    email: normalizeEmail_(at(roles.email)),
    cls: at(roles["class"]).toUpperCase()
  };
}

/** What a row says about who it is, short (for the remembered decisions). */
function od_rowKey_(cells) {
  const key = JSON.stringify([normalizeOsis_(cells.osis), cells.name ? od_tokenKey_(cells.name.tokens) : "",
    cells.guardian ? od_tokenKey_(cells.guardian.tokens) : "", cells.phone, cells.email]);
  // Two different 32-bit hashes -> a short, practically unique key
  let h1 = 0x811c9dc5, h2 = 0x01000193;
  for (let i = 0; i < key.length; i++) {
    const c = key.charCodeAt(i);
    h1 = Math.imul(h1 ^ c, 0x01000193) >>> 0;
    h2 = Math.imul(h2 + c, 0x5bd1e995) >>> 0;
  }
  return h1.toString(36) + h2.toString(36);
}

// ----------------------------------------------------------------------
// HOW ALIKE TWO NAMES ARE (owner's request 2026-11-12): every possible student /
// guardian gets a confidence 0-100% from the first and last names (last name 55%,
// first name 45%), so typos ("Simpsn"), the other order ("SIMPSON BART"), short
// names, nicknames and initials still find someone. A confidence is only ever a
// recommendation: a person says yes in the sidebar (Your pick), nothing is used
// before that. High = OD_CONF_HIGH_+ with first AND last name each close; medium
// = OD_CONF_MEDIUM_+; low = OD_CONF_MIN_+; under that the closest few are still
// listed as suggestions. Two candidates within OD_CONF_TIE_ points = "pick one".
// ----------------------------------------------------------------------
var OD_CONF_HIGH_ = 90, OD_CONF_MEDIUM_ = 75, OD_CONF_MIN_ = 60, OD_CONF_TIE_ = 4;
var OD_SCORE_FLOOR_ = 0.5; // below this a name is not worth suggesting
var OD_SIM_MEMO_ = null;

function od_pct_(x) { return Math.round(Math.min(1, Math.max(0, Number(x) || 0)) * 100); }
function od_confLevel_(pct) {
  if (pct === null || pct === undefined || isNaN(pct)) return "";
  return pct >= OD_CONF_HIGH_ ? "high" : pct >= OD_CONF_MEDIUM_ ? "medium" : pct >= OD_CONF_MIN_ ? "low" : "weak";
}

/** Jaro-Winkler similarity of two lowercase name parts (0-1). */
function od_jaroWinkler_(a, b) {
  if (a === b) return 1;
  if (!a || !b) return 0;
  const range = Math.max(0, Math.floor(Math.max(a.length, b.length) / 2) - 1);
  const aHit = new Array(a.length).fill(false), bHit = new Array(b.length).fill(false);
  let matches = 0;
  for (let i = 0; i < a.length; i++) {
    const lo = Math.max(0, i - range), hi = Math.min(b.length - 1, i + range);
    for (let j = lo; j <= hi; j++) {
      if (!bHit[j] && a[i] === b[j]) { aHit[i] = bHit[j] = true; matches++; break; }
    }
  }
  if (!matches) return 0;
  let k = 0, half = 0;
  for (let i = 0; i < a.length; i++) {
    if (!aHit[i]) continue;
    while (!bHit[k]) k++;
    if (a[i] !== b[k]) half++;
    k++;
  }
  const jaro = (matches / a.length + matches / b.length + (matches - half / 2) / matches) / 3;
  let prefix = 0;
  while (prefix < 4 && prefix < a.length && prefix < b.length && a[prefix] === b[prefix]) prefix++;
  return jaro + prefix * 0.1 * (1 - jaro);
}

/** 1 - edits / longer length (0-1). */
function od_editRatio_(a, b) {
  if (a === b) return 1;
  if (!a || !b) return 0;
  let prev = [];
  for (let j = 0; j <= b.length; j++) prev[j] = j;
  for (let i = 1; i <= a.length; i++) {
    const cur = [i];
    for (let j = 1; j <= b.length; j++) cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    prev = cur;
  }
  return 1 - prev[b.length] / Math.max(a.length, b.length);
}

/** Spelling likeness of two name parts: Jaro-Winkler and edit distance averaged (JW alone over-rewards a shared start: Mary / Mark). */
function od_spellAlike_(a, b) {
  if (a === b) return 1;
  if (!a || !b) return 0;
  if (!OD_SIM_MEMO_) OD_SIM_MEMO_ = new Map();
  const key = a < b ? a + "|" + b : b + "|" + a;
  if (OD_SIM_MEMO_.has(key)) return OD_SIM_MEMO_.get(key);
  let v = (od_jaroWinkler_(a, b) + od_editRatio_(a, b)) / 2;
  if ((a.length === 1 && b.charAt(0) === a) || (b.length === 1 && a.charAt(0) === b)) v = Math.max(v, 0.7); // an initial
  OD_SIM_MEMO_.set(key, v);
  return v;
}

/** First names: the rules (same / nickname / "Sam" for Samuel / initial) or the spelling, whichever is more. */
function od_firstScore_(a, b) {
  if (!a || !b) return 0;
  const how = od_firstLike_(a, b);
  const rule = how === "same" ? 1 : how === "nickname" ? 0.95 : how === "start" ? 0.9 : how === "initial" ? 0.72 : 0;
  return Math.max(rule, od_spellAlike_(a, b));
}

/** Best likeness between any part of one list and any part of the other. */
function od_bestPair_(as, bs) {
  let best = 0;
  as.forEach(a => bs.forEach(b => { best = Math.max(best, od_spellAlike_(a, b)); }));
  return best;
}

/** Share of the shorter name's parts found (closely) in the longer one: middle names, two last names. */
function od_tokenShare_(a, b) {
  const small = a.length <= b.length ? a : b, big = small === a ? b : a;
  if (small.length < 2) return { score: 0, gate: false };
  let sum = 0, worst = 1;
  small.forEach(t => { let m = 0; big.forEach(u => { m = Math.max(m, od_spellAlike_(t, u)); }); sum += m; worst = Math.min(worst, m); });
  return { score: 0.95 * sum / small.length, gate: worst >= 0.9 };
}

/**
 * How alike a row's name is to a student's / guardian's (0-1). Variants: the
 * normal order, the last name written first without a comma, and the share of
 * name parts; the best one counts. "High" needs each name part close (else 0.89 max).
 */
function od_nameScore_(n, p) {
  if (!n || !p || !n.tokens.length || !p.tokens.length) return 0;
  const nf = n.firstTokens && n.firstTokens.length ? n.firstTokens : n.tokens.slice(0, 1);
  const nl = n.lastTokens && n.firstTokens && n.firstTokens.length ? n.lastTokens : n.tokens.slice(1);
  const pf = p.firstTokens && p.firstTokens.length ? p.firstTokens : p.tokens.slice(0, 1);
  const pl = p.lastTokens && p.firstTokens && p.firstTokens.length ? p.lastTokens : p.tokens.slice(1);
  const highGate = (score, ok) => (score >= OD_CONF_HIGH_ / 100 && !ok ? (OD_CONF_HIGH_ - 1) / 100 : score);
  if (!nl.length || !pl.length) {
    // One name part on one side ("Bart"): it can be the first or the last name, never more than low
    return 0.7 * od_bestPair_(n.tokens, p.tokens);
  }
  const first = od_firstScore_(nf[0], pf[0]), last = od_bestPair_(nl, pl);
  const normal = highGate(0.45 * first + 0.55 * last, first >= 0.9 && last >= 0.9);
  const sf = od_firstScore_(nl[nl.length - 1], pf[0]), sl = od_bestPair_(nf, pl);
  const swapped = highGate(0.45 * sf + 0.55 * sl - 0.02, sf >= 0.9 && sl >= 0.9);
  const share = od_tokenShare_(n.tokens, p.tokens);
  return Math.max(normal, swapped, highGate(share.score, share.gate));
}

/**
 * Who is worth scoring. A full name: people whose LAST name starts with the same
 * letter as one of the row's name parts (either order); a one-word name: anyone
 * with a name part starting with its letter. (A typo in the first letter of a last
 * name is not caught; it keeps big lists fast.)
 */
function od_candidates_(parts, byInitial, byLastInitial) {
  const seen = new Set();
  const map = parts.tokens.length > 1 && byLastInitial ? byLastInitial : byInitial;
  parts.tokens.forEach(t => (map.get(t.charAt(0)) || []).forEach(x => seen.add(x)));
  return [...seen];
}

/** "Close spelling: last name “Simpsn” ≈ “Simpson”" (what differs). */
function od_alikeWhy_(n, p) {
  const cap = x => x ? x.charAt(0).toUpperCase() + x.slice(1) : x;
  const bits = [];
  if (n.first && p.first && n.first !== p.first) bits.push("first name “" + cap(n.first) + "” ≈ “" + cap(p.first) + "”");
  if (n.last && p.last && n.last !== p.last) bits.push("last name “" + cap(n.last) + "” ≈ “" + cap(p.last) + "”");
  return bits.length ? "Close spelling: " + bits.join(", ") : "Same name parts in another order";
}

/** Students for a row, best first: [{ item: student, score 0-1, why }]. */
function od_rankStudents_(cells, idx, contact) {
  const best = new Map();
  const put = (st, score, why) => {
    const o = best.get(st);
    if (!o || score > o.score) best.set(st, { item: st, score: score, why: why || (o ? o.why : "") });
  };
  const n = cells.name;
  if (n) {
    od_candidates_(n, idx.byInitial, idx.byLastInitial).forEach(st => {
      const sc = od_nameScore_(n, st);
      if (sc >= OD_SCORE_FLOOR_) put(st, sc, od_alikeWhy_(n, st));
    });
  } else if (cells.guardian) {
    // No student name on the row: guardians with a close name -> their children
    od_candidates_(cells.guardian, idx.pByInitial, idx.pByLastInitial).forEach(p => {
      const sc = od_nameScore_(cells.guardian, p) * 0.95;
      if (sc >= OD_SCORE_FLOOR_) p.kids.forEach(st => put(st, sc, "Parent “" + cells.guardian.raw + "” looks like " + p.name));
    });
  }
  // The rules (short name, nickname, initial, part of the name) give their reason and a floor.
  // With a student name on the row, the parent's name only adds the small boost below
  // (else every sibling would get the parent rule's floor).
  od_studentGuesses_(n ? { name: n, guardian: null } : cells, idx).forEach(g => {
    const o = best.get(g.st);
    best.set(g.st, { item: g.st, score: Math.max(o ? o.score : 0, g.floor), why: g.why });
  });
  // What else on the row agrees with it: the parent's name, the phone / email, the class
  best.forEach(o => {
    const st = o.item, extra = [];
    if (n && cells.guardian && st.persons.some(p => od_personFits_(cells.guardian, p))) { o.score += 0.05; extra.push("parent matches"); }
    if (contact && contact.kids.indexOf(st) !== -1) { o.score += 0.06; extra.push("same " + (cells.phone ? "phone" : "email")); }
    if (cells.cls && st.cls === cells.cls) { o.score += 0.04; extra.push("same class"); }
    o.score = Math.min(o.score, 0.99);
    if (extra.length) o.why = (o.why ? o.why + "; " : "") + extra.join(", ");
  });
  return [...best.values()].sort((a, b) => b.score - a.score);
}

/** Guardians for a row of a guardian list, best first (same shape). */
function od_rankPersons_(cells, idx, contactPersons) {
  const best = new Map();
  const put = (p, score, why) => {
    const o = best.get(p);
    if (!o || score > o.score) best.set(p, { item: p, score: score, why: why || (o ? o.why : "") });
  };
  if (cells.guardian) {
    od_candidates_(cells.guardian, idx.pByInitial, idx.pByLastInitial).forEach(p => {
      const sc = od_nameScore_(cells.guardian, p);
      if (sc >= OD_SCORE_FLOOR_) put(p, sc, od_alikeWhy_(cells.guardian, p));
    });
    // The old rules (one name inside the other, same last name + short / initial first name) as a floor
    od_personCandidates_(cells.guardian, idx).forEach(p => {
      const fit = od_personFits_(cells.guardian, p);
      if (!fit) return;
      const o = best.get(p);
      best.set(p, { item: p, score: Math.max(o ? o.score : 0, fit === "same" || fit === "inside" ? 0.9 : 0.8), why: "“" + cells.guardian.raw + "” looks like " + p.name });
    });
  }
  if (cells.name) {
    // The list also names the child: that child's guardians
    od_rankStudents_({ name: cells.name, guardian: null, cls: cells.cls }, idx, null).forEach(r => {
      if (r.score < OD_CONF_MEDIUM_ / 100) return;
      r.item.persons.forEach(p => {
        const o = best.get(p);
        if (o) { o.score += 0.06; o.why += "; child " + r.item.name; }
        else if (!cells.guardian) put(p, r.score * 0.85, "Guardian of " + r.item.name + " (child's name " + od_pct_(r.score) + "% alike)");
      });
    });
  }
  best.forEach(o => {
    if (contactPersons && contactPersons.indexOf(o.item) !== -1) { o.score += 0.06; o.why += (o.why ? "; " : "") + "same " + (cells.phone ? "phone" : "email"); }
    o.score = Math.min(o.score, 0.99);
  });
  return [...best.values()].sort((a, b) => b.score - a.score);
}

/** Ranked candidates -> a best guess or "pick one" (null when nobody reaches OD_CONF_MIN_). */
function od_decideRanked_(ranked, keyOf, label, word, note) {
  const good = ranked.filter(r => od_pct_(r.score) >= OD_CONF_MIN_);
  if (!good.length) return null;
  const top = good[0];
  const scores = {};
  good.slice(0, 5).forEach(r => { scores[keyOf(r.item)] = od_pct_(r.score); });
  const close = good.filter(r => od_pct_(top.score) - od_pct_(r.score) < OD_CONF_TIE_);
  // "Pick one" only when they are at least medium: two LOW ties are usually siblings who only share
  // the last name with someone who isn't on the list ("Tina Belcher" vs Louise / Gene): not found,
  // with them as the closest suggestions (test data 2026-11-12)
  if (close.length > 1 && od_pct_(top.score) < OD_CONF_MEDIUM_) return null;
  if (close.length > 1) {
    return od_pick_(close.slice(0, 5).map(r => keyOf(r.item)), close.length + " " + word + " are about as close to “" + label + "”" + note, scores);
  }
  return od_guess_(keyOf(top.item), (top.why || "The names are alike") + note, od_pct_(top.score),
    good.slice(1, 3).map(r => keyOf(r.item)), scores);
}

/** The closest few under OD_CONF_MIN_, still offered on a "not found" row. */
function od_suggestions_(ranked, keyOf) {
  const list = ranked.filter(r => r.score >= OD_SCORE_FLOOR_).slice(0, 3);
  const scores = {};
  list.forEach(r => { scores[keyOf(r.item)] = od_pct_(r.score); });
  return { targets: list.map(r => keyOf(r.item)), scores: scores };
}

/** { target: 100 } for every target (exact matches offered side by side). */
function od_fullScores_(targets) {
  const out = {};
  targets.forEach(t => { out[t] = 100; });
  return out;
}

function od_ok_(target, how) { return { status: "ok", target: target, how: how }; }
// confidence: 0-100 (how alike the names are); scores: { target: 0-100 } for the options
function od_guess_(target, why, confidence, alternatives, scores) {
  return { status: "guess", target: target, how: "Best guess", reason: why, confidence: confidence === undefined ? null : confidence,
           options: alternatives || [], scores: scores || {} };
}
function od_pick_(targets, why, scores) { return { status: "pick", options: targets, how: "Check", reason: why, scores: scores || {} }; }
// suggestions: the closest few (under OD_CONF_MIN_), still offered to pick from
function od_none_(why, suggestions, scores) { return { status: "none", how: "—", reason: why, options: suggestions || [], scores: scores || {} }; }

/** Guardian name fits a person: same name, one name inside the other, or same last name with a short / initial first name. */
function od_personFits_(parts, p) {
  if (!parts) return "";
  if (sameNameExactly_(parts.tokens, p.tokens)) return "same";
  if (namesMatch_(parts.tokens, p.tokens)) return "inside";
  const f = od_firstLike_(parts.first, p.first), l = od_lastLike_(parts.last, p.last);
  return f && l && !(f === "initial" && l === "initial") ? f + "/" + l : "";
}

/** Guardians worth comparing with a name: sharing a name part, the first name or the last name (not everyone). */
function od_personCandidates_(parts, idx) {
  const seen = [];
  const add = p => { if (seen.indexOf(p) === -1) seen.push(p); };
  parts.tokens.forEach(t => (idx.pByToken.get(t) || []).forEach(add));
  if (parts.last) (idx.pByLast.get(parts.last) || []).forEach(add);
  if (parts.last && parts.last.length === 1) idx.persons.forEach(p => { if (p.first === parts.first) add(p); }); // "Martha K."
  return seen;
}

/** Students a row's phone / email belongs to (through their guardians). */
function od_studentsByContact_(cells, idx) {
  const persons = [];
  if (cells.phone) (idx.byPhone.get(cells.phone) || []).forEach(p => { if (persons.indexOf(p) === -1) persons.push(p); });
  if (cells.email) (idx.byEmail.get(cells.email) || []).forEach(p => { if (persons.indexOf(p) === -1) persons.push(p); });
  const kids = [];
  persons.forEach(p => p.kids.forEach(k => { if (kids.indexOf(k) === -1) kids.push(k); }));
  return { persons: persons, kids: kids };
}

/** Best-guess students for a name (never used without a person saying yes). */
function od_studentGuesses_(cells, idx) {
  const out = [];
  // floor: the least confidence this kind of match gets (the name score can be higher)
  const add = (st, why, floor) => { if (!out.some(o => o.st === st)) out.push({ st: st, why: why, floor: floor || 0.8 }); };
  const n = cells.name;
  if (n) {
    const seen = new Set();
    n.tokens.forEach(t => (idx.byToken.get(t) || []).forEach(st => seen.add(st)));
    if (n.first) (idx.byFirst.get(n.first) || []).forEach(st => seen.add(st));
    // a short first name ("Mira") with no shared token yet: students with that last name
    if (n.last) (idx.byLast.get(n.last) || []).forEach(st => seen.add(st));
    seen.forEach(st => {
      if (namesMatch_(n.tokens, st.tokens) && !sameNameExactly_(n.tokens, st.tokens)) { add(st, "“" + n.raw + "” is part of " + st.name, 0.9); return; }
      if (n.last) {
        const f = od_firstLike_(n.first, st.first), l = od_lastLike_(n.last, st.last);
        if (f && l && !(f === "initial" && l === "initial")) {
          add(st, l === "initial" ? "Same first name, “" + n.last.toUpperCase() + ".” could be " + st.last.charAt(0).toUpperCase() + st.last.slice(1)
                                  : (f === "same" ? "Same name" : od_firstWhy_(f, n.first, st.name) + ", same last name"),
              l === "initial" ? 0.78 : f === "initial" ? 0.76 : 0.88);
        }
      } else if (n.tokens.length === 1 && st.first === n.first) {
        add(st, "First name only", 0.6);
      }
    });
  }
  if (!out.length && cells.guardian) {
    // No student name to go on: the parent's name ("Martha K.") -> their children
    od_personCandidates_(cells.guardian, idx).forEach(p => {
      const fit = od_personFits_(cells.guardian, p);
      if (fit) p.kids.forEach(st => add(st, "Parent “" + cells.guardian.raw + "” looks like " + p.name, fit === "same" ? 0.85 : 0.75));
    });
  }
  return out;
}

/** One row of a STUDENT list -> ok / guess / pick / none (targets = normalized OSIS). */
function od_matchStudentRow_(cells, idx) {
  let osisUnknown = false;
  if (cells.osis && normalizeOsis_(cells.osis)) {
    const st = idx.byOsis.get(normalizeOsis_(cells.osis));
    if (st) return od_ok_(st.key, "OSIS");
    osisUnknown = true;
  }
  const exact = cells.name ? (idx.byNameKey.get(od_tokenKey_(cells.name.tokens)) || []) : [];
  const contact = od_studentsByContact_(cells, idx);
  const narrow = list => list.filter(st =>
    (cells.guardian && st.persons.some(p => od_personFits_(cells.guardian, p))) ||
    contact.kids.indexOf(st) !== -1 || (cells.cls && st.cls === cells.cls));

  if (exact.length === 1) {
    if (!osisUnknown) return od_ok_(exact[0].key, "Name");
    return od_guess_(exact[0].key, "Same name, but OSIS " + cells.osis + " isn't on the Student List", 100);
  }
  if (exact.length > 1) {
    const n = narrow(exact);
    if (n.length === 1 && !osisUnknown) return od_ok_(n[0].key, "Name + parent");
    const targets = (n.length ? n : exact).map(st => st.key);
    return od_pick_(targets, exact.length + " students named " + exact[0].name, od_fullScores_(targets));
  }
  if (contact.kids.length === 1 && !osisUnknown) {
    // One family with one child: the phone / email says who it is (a name, if any, must not say otherwise)
    const st = contact.kids[0];
    if (!cells.name || od_studentGuesses_(cells, idx).some(g => g.st === st) || !cells.name.last) return od_ok_(st.key, cells.phone ? "Phone" : "Email");
  }
  if (contact.kids.length > 1) {
    const byName = cells.name ? od_studentGuesses_(cells, idx).filter(g => contact.kids.indexOf(g.st) !== -1) : [];
    if (byName.length === 1 && !osisUnknown) return od_ok_(byName[0].st.key, (cells.phone ? "Phone" : "Email") + " + name");
  }
  // Best guesses with a confidence (how alike the names are); never used before a person says yes
  const ranked = od_rankStudents_(cells, idx, contact);
  const osisNote = osisUnknown ? " (OSIS " + cells.osis + " isn't on the Student List)" : "";
  const label = cells.name ? cells.name.raw : (cells.guardian ? cells.guardian.raw : "this row");
  const decided = od_decideRanked_(ranked, st => st.key, label, "students", osisNote);
  if (decided) return decided;
  if (contact.kids.length > 1) {
    const targets = contact.kids.map(st => st.key);
    return od_pick_(targets, "That phone / email belongs to " + contact.kids.length + " students");
  }
  const sugg = od_suggestions_(ranked, st => st.key);
  if (osisUnknown) return od_none_("OSIS " + cells.osis + " isn't on the Student List", sugg.targets, sugg.scores);
  if (!cells.name && !cells.guardian && !cells.phone && !cells.email) return od_none_("This row has no OSIS, name, phone or email");
  return od_none_(sugg.targets.length ? "No close match on the Student List" : "No student on the Student List matches this row", sugg.targets, sugg.scores);
}

/** One row of a GUARDIAN list -> ok / guess / pick / none (targets = person keys). */
function od_matchGuardianRow_(cells, idx) {
  if (cells.osis && normalizeOsis_(cells.osis)) {
    const st = idx.byOsis.get(normalizeOsis_(cells.osis));
    if (st) {
      if (cells.guardian) {
        const m = st.persons.filter(p => od_personFits_(cells.guardian, p));
        if (m.length === 1) return od_ok_(m[0].key, "OSIS + name");
      } else if (st.persons.length === 1) {
        return od_ok_(st.persons[0].key, "OSIS");
      }
      if (st.persons.length) {
        const scores = {};
        if (cells.guardian) st.persons.forEach(p => { scores[p.key] = od_pct_(od_nameScore_(cells.guardian, p)); });
        return od_pick_(st.persons.map(p => p.key), "Which guardian of " + st.name + "?", scores);
      }
    }
  }
  const exact = cells.guardian ? (idx.pByNameKey.get(od_tokenKey_(cells.guardian.tokens)) || []) : [];
  const contact = od_studentsByContact_(cells, idx).persons;
  if (exact.length === 1) return od_ok_(exact[0].key, "Name");
  if (exact.length > 1) {
    const n = exact.filter(p => contact.indexOf(p) !== -1);
    if (n.length === 1) return od_ok_(n[0].key, "Name + phone");
    const targets = exact.map(p => p.key);
    return od_pick_(targets, exact.length + " guardians named " + exact[0].name, od_fullScores_(targets));
  }
  if (contact.length === 1) return od_ok_(contact[0].key, cells.phone ? "Phone" : "Email");
  // Best guesses with a confidence (guardian's name, the child's name, phone / email)
  const ranked = od_rankPersons_(cells, idx, contact);
  const label = cells.guardian ? cells.guardian.raw : (cells.name ? cells.name.raw : "this row");
  const decided = od_decideRanked_(ranked, p => p.key, label, "guardians", "");
  if (decided) return decided;
  if (contact.length > 1) return od_pick_(contact.map(p => p.key), "That phone / email belongs to " + contact.length + " guardians");
  const sugg = od_suggestions_(ranked, p => p.key);
  return od_none_(sugg.targets.length ? "No close match among the guardians" : "No guardian matches this row", sugg.targets, sugg.scores);
}

/** Every data row of a list -> its result, your decisions (setup.picks) first. */
function od_matchAll_(setup, rows, idx) {
  const by = setup.by === "guardian" ? "guardian" : "student";
  const exists = t => (by === "student" ? idx.byOsis.has(t) : idx.pByKey.has(t));
  return rows.map(row => {
    const cells = od_rowCells_(row, setup.roles || {});
    const key = od_rowKey_(cells);
    const decided = setup.picks[key];
    let res;
    if (decided === "skip") res = { status: "skip", how: "Skipped" };
    else if (decided && exists(decided)) res = od_ok_(decided, "Your pick");
    else res = by === "student" ? od_matchStudentRow_(cells, idx) : od_matchGuardianRow_(cells, idx);
    res.key = key;
    res.cells = cells;
    return res;
  });
}

function od_targetText_(by, target, idx) {
  if (by === "student") {
    const st = idx.byOsis.get(target);
    return st ? { text: st.osis + " · " + st.name, sub: st.guardianText } : { text: target, sub: "" };
  }
  const p = idx.pByKey.get(target);
  return p ? { text: p.name, sub: p.kids.map(k => k.name).join(", ") } : { text: target, sub: "" };
}

function od_stats_(results) {
  const s = { rows: results.length, matched: 0, review: 0, skipped: 0, none: 0 };
  results.forEach(r => {
    if (r.status === "ok") s.matched++;
    else if (r.status === "skip") s.skipped++;
    else if (r.status === "none") { s.none++; s.review++; }
    else s.review++;
  });
  return s;
}

// ======================================================================
// OUTPUTS (every build)
// ======================================================================

/** A variable's value for one row: the cell, or the blank text, or null (nothing to show). */
function od_varValue_(v, row) {
  const cell = toText_(row[v.col - 1]);
  if (cell !== "") return cell;
  return v.blank && v.blank.mode === "text" && toText_(v.blank.text) ? toText_(v.blank.text) : null;
}

/**
 * Called by every build (buildAllDerivedSheets, in a try/catch), after the
 * Student List is built in memory. Matches every list, refreshes each list tab's
 * "✔ Matched" columns and bands, writes the Backend's Additional Information and
 * returns { studentCols, guardianCols } for the Student List / Parents Divided.
 */
function od_buildOutputs_(ss, masterTable, guardianRows) {
  const lists = od_loadLists_(ss);
  const ready = lists.filter(l => l.sheet && l.setup.ready);
  const out = { studentCols: [], guardianCols: [] };
  if (!lists.length) {
    backendClearBlock_(ss, "ADDL_INFO", BACKEND_HEADER_ROW);
    return out;
  }
  const idx = od_index_(masterTable.slice(1), guardianRows);
  const lines = new Map(); // OSIS -> { student: [], byGuardian: Map(lower name -> { name, lines }) }
  const entry = osis => {
    if (!lines.has(osis)) lines.set(osis, { student: [], byGuardian: new Map() });
    return lines.get(osis);
  };
  const addLine = (osis, guardianName, text) => {
    const e = entry(osis);
    if (!guardianName) { if (e.student.indexOf(text) === -1) e.student.push(text); return; }
    const k = guardianName.toLowerCase();
    if (!e.byGuardian.has(k)) e.byGuardian.set(k, { name: guardianName, lines: [] });
    const g = e.byGuardian.get(k);
    if (g.lines.indexOf(text) === -1) g.lines.push(text);
  };

  ready.forEach(list => {
    try {
      const setup = list.setup;
      const table = od_readTable_(list.sheet, setup);
      const results = od_matchAll_(setup, table.rows, idx);
      od_writeToolColumns_(list.sheet, list, table, results, idx);
      setup.stats = od_stats_(results);
      const by = setup.by === "guardian" ? "guardian" : "student";
      setup.vars.filter(v => v.col && (v.show || v.column)).forEach(v => {
        const label = toText_(v.label) || toText_(v.header) || ("Column " + od_colLetter_(v.col));
        const values = new Map(); // target -> [values]
        results.forEach((res, i) => {
          if (res.status !== "ok") return;
          const val = od_varValue_(v, table.rows[i]);
          if (val === null) return;
          if (!values.has(res.target)) values.set(res.target, []);
          if (values.get(res.target).indexOf(val) === -1) values.get(res.target).push(val);
        });
        const notText = v.notOnList && v.notOnList.mode === "text" ? toText_(v.notOnList.text) : "";
        if (by === "student") {
          const byOsis = new Map();
          idx.students.forEach(st => {
            const vals = values.get(st.key);
            const value = vals ? vals.join(" / ") : notText;
            if (!value) return;
            byOsis.set(st.osis, value);
            if (v.show) addLine(st.osis, "", label + ": " + value);
          });
          if (v.column) out.studentCols.push({ header: label, byOsis: byOsis });
        } else {
          const byGuardian = new Map();
          idx.persons.forEach(p => {
            const vals = values.get(p.key);
            const value = vals ? vals.join(" / ") : notText;
            if (!value) return;
            p.kids.forEach(st => {
              byGuardian.set(st.osis + "|" + p.name.toLowerCase(), value);
              if (v.show) addLine(st.osis, p.name, label + ": " + value);
            });
          });
          if (v.column) out.guardianCols.push({ header: label, byGuardian: byGuardian });
        }
      });
    } catch (e) {
      console.warn("Your Lists: \"" + list.name + "\" skipped: " + e.message);
    }
  });

  // Backend Additional Information: student lines (Guardian blank), then each guardian's
  const rows = [];
  idx.students.forEach(st => {
    const e = lines.get(st.osis);
    if (!e) return;
    if (e.student.length) rows.push([st.osis, st.name, "", e.student.join("\n")]);
    e.byGuardian.forEach(g => rows.push([st.osis, st.name, g.name, g.lines.join("\n")]));
  });
  rows.sort(byStudentLastName_(1));
  backendWriteBlock_(ss, "ADDL_INFO", [["OSIS", "Student", "Guardian", "Additional Information"]].concat(rows));
  od_saveLists_(ss, lists); // keeps each list's counts for the sidebar
  return out;
}

/** Rows 1-3 + the "✔ Matched" / "How" columns of a list tab. */
function od_writeToolColumns_(sheet, list, table, results, idx) {
  const setup = list.setup;
  const width = table.headers.length;
  if (!width) return;
  const toolCol = width + 2;
  if (sheet.getMaxColumns() < toolCol + 1) sheet.insertColumnsAfter(sheet.getMaxColumns(), toolCol + 1 - sheet.getMaxColumns());
  const by = setup.by === "guardian" ? "guardian" : "student";
  const cells = [OD_TOOL_HEADERS.slice()], colors = [[OD_COLORS.header, OD_COLORS.header]];
  results.forEach(r => {
    let text, color;
    if (r.status === "ok") { text = od_targetText_(by, r.target, idx).text; color = OD_COLORS.ok; }
    else if (r.status === "guess") { text = "? " + od_targetText_(by, r.target, idx).text + " (best guess" + (r.confidence !== null && r.confidence !== undefined ? ", " + r.confidence + "%" : "") + ")"; color = OD_COLORS.guess; }
    else if (r.status === "pick") { text = "? " + r.reason; color = OD_COLORS.guess; }
    else if (r.status === "skip") { text = "— skipped"; color = OD_COLORS.skip; }
    else {
      const closest = (r.options || [])[0];
      text = "✗ " + r.reason + (closest ? " (closest: " + od_targetText_(by, closest, idx).text + ", " + ((r.scores || {})[closest] || 0) + "%)" : "");
      color = OD_COLORS.none;
    }
    cells.push([text, r.how || ""]);
    colors.push([color, color]);
  });
  const last = sheet.getMaxRows();
  if (last >= OD_HEADER_ROW) sheet.getRange(OD_HEADER_ROW, toolCol - 1, last - OD_HEADER_ROW + 1, 3).clearContent().setBackground(null);
  sheet.getRange(OD_HEADER_ROW, toolCol, cells.length, 2).setValues(cells).setBackgrounds(colors);
  sheet.getRange(OD_HEADER_ROW, toolCol, 1, 2).setFontWeight("bold");

  const stats = od_stats_(results);
  const status = stats.review
    ? stats.matched + " of " + stats.rows + " rows matched · " + stats.review + " need you (🚀 App Menu → 📥 Input Data Mode → Your Lists)"
    : "All " + stats.matched + " row(s) matched" + (stats.skipped ? " (" + stats.skipped + " skipped)" : "") + " ✓";
  const roleName = { osis: "OSIS", "class": "CLASS", phone: "PHONE", email: "EMAIL" };
  const chips = table.headers.map((h, i) => {
    const c = i + 1, roles = setup.roles || {};
    for (const k of Object.keys(roleName)) if (roles[k] === c) return roleName[k];
    const nameRole = (r, label) => r && (r.col === c || r.first === c || r.last === c) ? label + (r.first === c ? " (FIRST)" : r.last === c ? " (LAST)" : "") : "";
    const nr = nameRole(roles.student, "STUDENT") || nameRole(roles.guardian, "GUARDIAN");
    if (nr) return nr;
    const v = setup.vars.filter(x => x.col === c)[0];
    if (v && (v.show || v.column)) return "SHOWN: " + (toText_(v.label) || h);
    return "TAB ONLY";
  });
  while (chips.length < toolCol - 1) chips.push("");
  chips.push("FILLED BY MATCHING");
  od_writeBands_(sheet, list.name, list.updated || od_today_(), status, chips);
}

// ======================================================================
// COLUMN GUESSES (what each column is)
// ======================================================================

function od_columnStats_(headers, rows, idx) {
  return headers.map((h, i) => {
    const vals = rows.map(r => toText_(r[i])).filter(Boolean);
    const n = vals.length || 1;
    let osisLike = 0, osisKnown = 0, phoneLike = 0, emailLike = 0, studentKnown = 0, guardianKnown = 0, classKnown = 0, comma = 0, multi = 0;
    vals.forEach(v => {
      const digits = v.replace(/\D/g, "");
      if (/^[\d\s-]+$/.test(v) && digits.length >= 6 && digits.length <= 9) { osisLike++; if (idx.byOsis.has(normalizeOsis_(v))) osisKnown++; }
      if (digits.length >= 10 && digits.length <= 11 && normalizePhoneForMatch_(v).length === 10) phoneLike++;
      if (v.indexOf("@") !== -1) emailLike++;
      if (v.indexOf(",") !== -1) comma++;
      const parts = od_nameParts_(v);
      if (parts && parts.tokens.length >= 2) multi++; // a full name, not just a first name
      if (parts && /[a-z]/i.test(v) && v.indexOf("@") === -1) {
        const key = od_tokenKey_(parts.tokens);
        if (idx.byNameKey.has(key) || (parts.last && (idx.byLast.get(parts.last) || []).some(st => od_firstLike_(parts.first, st.first)))) studentKnown++;
        if (idx.pByNameKey.has(key) || (parts.last && (idx.pByLast.get(parts.last) || []).some(p => od_firstLike_(parts.first, p.first)))) guardianKnown++;
      }
      if (idx.classes.has(v.toUpperCase())) classKnown++;
    });
    return { col: i + 1, header: h, n: vals.length, osisLike: osisLike / n, osisKnown: osisKnown, osisKnownShare: osisKnown / n,
             phone: phoneLike / n, email: emailLike / n, student: studentKnown / n, guardian: guardianKnown / n, cls: classKnown / n, comma: comma / n,
             multi: multi / n, samples: vals.slice(0, 3) };
  });
}

/** Best guess of what each column is + student or guardian list. */
function od_guessRoles_(headers, rows, idx) {
  const stats = od_columnStats_(headers, rows, idx);
  const used = new Set();
  const roles = { osis: 0, student: null, guardian: null, phone: 0, email: 0, "class": 0 };
  const H = s => String(s.header).toLowerCase();
  const pick = (test, score) => {
    let best = null, bestScore = 0;
    stats.forEach(s => { if (used.has(s.col) || !test(s)) return; const sc = score(s); if (sc > bestScore) { best = s; bestScore = sc; } });
    if (best) used.add(best.col);
    return best;
  };
  const guardianWord = /parent|guardian|mother|father|\bmom\b|\bdad\b|family|adult|contact/;
  const studentWord = /student|child|kid|scholar/;

  const osis = pick(s => s.osisKnownShare >= 0.3 || (/osis|student\s*(id|#|no|num)|\bid\b|id\s*#/.test(H(s)) && s.osisLike >= 0.5),
                    s => s.osisKnownShare * 2 + s.osisLike + (/osis/.test(H(s)) ? 1 : 0));
  if (osis) roles.osis = osis.col;
  const email = pick(s => s.email >= 0.6 || (/e-?mail/.test(H(s)) && s.email >= 0.3), s => s.email + (/e-?mail/.test(H(s)) ? 1 : 0));
  if (email) roles.email = email.col;
  const phone = pick(s => s.phone >= 0.6 || (/phone|cell|mobile|\btel/.test(H(s)) && s.phone >= 0.3), s => s.phone + (/phone|cell|mobile/.test(H(s)) ? 1 : 0));
  if (phone) roles.phone = phone.col;

  // Two-column names ("First Name" / "Last Name", "Parent First" / "Parent Last")
  const firstCols = stats.filter(s => !used.has(s.col) && /first/.test(H(s)));
  const lastCols = stats.filter(s => !used.has(s.col) && /last|surname/.test(H(s)));
  const pairFor = isGuardian => {
    const f = firstCols.filter(s => guardianWord.test(H(s)) === isGuardian)[0];
    const l = lastCols.filter(s => guardianWord.test(H(s)) === isGuardian)[0];
    if (f && l) { used.add(f.col); used.add(l.col); return { first: f.col, last: l.col }; }
    return null;
  };
  roles.guardian = pairFor(true);
  roles.student = pairFor(false);

  if (!roles.student) {
    const st = pick(s => !guardianWord.test(H(s)) && (studentWord.test(H(s)) || /name/.test(H(s)) || s.student >= 0.3) && s.osisLike < 0.5 && s.phone < 0.5,
                    s => s.student * 2 + (studentWord.test(H(s)) ? 1 : 0) + (/name/.test(H(s)) ? 0.5 : 0) - (s.guardian > s.student ? 1 : 0));
    if (st && (st.student >= 0.2 || studentWord.test(H(st)) || /^(full\s*)?name$/.test(H(st).trim()))) roles.student = { col: st.col };
    else if (st) used.delete(st.col);
  }
  if (!roles.guardian) {
    const g = pick(s => (guardianWord.test(H(s)) || s.guardian >= 0.3) && s.osisLike < 0.5 && s.phone < 0.5 && s.email < 0.5,
                   s => s.guardian * 2 + (guardianWord.test(H(s)) ? 1 : 0));
    if (g) roles.guardian = { col: g.col };
  }
  const cls = pick(s => s.cls >= 0.3 && /class|homeroom|official|section|room/.test(H(s)), s => s.cls);
  if (cls) roles["class"] = cls.col;

  let by = (roles.osis || roles.student) ? "student" : (roles.guardian || roles.phone || roles.email) ? "guardian" : "student";
  // A guardian list that also names the child by first name only ("Parent First | Parent Last | Child"):
  // the parent is who the row is (test data 2026-11-12: the PTA list was guessed as a student list)
  if (by === "student" && !roles.osis && roles.guardian && roles.student && roles.student.col) {
    const childCol = stats[roles.student.col - 1];
    const guardianFull = roles.guardian.first || (stats[roles.guardian.col - 1] || {}).multi >= 0.5;
    if (childCol && childCol.multi < 0.5 && guardianFull) by = "guardian";
  }
  return { roles: roles, by: by, stats: stats };
}

/** Columns that aren't used to find who the row is (the variables you can keep). */
function od_roleCols_(roles) {
  const cols = [];
  ["osis", "phone", "email", "class"].forEach(k => { if (roles[k]) cols.push(roles[k]); });
  ["student", "guardian"].forEach(k => {
    const r = roles[k];
    if (!r) return;
    ["col", "first", "last"].forEach(x => { if (r[x]) cols.push(r[x]); });
  });
  return cols;
}

/**
 * Variables for the headers: kept from the saved setup by header name (a new
 * version of the same list keeps its settings), new columns start "tab only"
 * (a list with just one other column starts shown).
 */
function od_varsFor_(headers, roles, savedVars) {
  const roleCols = od_roleCols_(roles);
  const byHeader = {};
  (savedVars || []).forEach(v => { byHeader[String(v.header).toLowerCase()] = v; });
  const vars = [];
  headers.forEach((h, i) => {
    const col = i + 1;
    if (roleCols.indexOf(col) !== -1) return;
    const old = byHeader[String(h).toLowerCase()];
    vars.push(old ? Object.assign({}, old, { col: col, header: h }) :
      { col: col, header: h, label: h, show: false, column: false, blank: { mode: "nothing", text: "" }, notOnList: { mode: "nothing", text: "" } });
  });
  if (vars.length === 1 && !(savedVars || []).length) { vars[0].show = true; vars[0].column = true; }
  return vars;
}

/** Saved roles moved to where those headers are now (null when one is missing). */
function od_rolesByHeader_(savedSetup, headers) {
  const old = savedSetup.headers || [];
  if (!old.length) return null;
  const pos = {};
  headers.forEach((h, i) => { pos[String(h).toLowerCase()] = i + 1; });
  const move = c => (c ? pos[String(old[c - 1] || "").toLowerCase()] || -1 : 0);
  const r = savedSetup.roles || {}, out = {};
  let missing = false;
  ["osis", "phone", "email", "class"].forEach(k => { const m = move(r[k]); if (m === -1) missing = true; out[k] = m > 0 ? m : 0; });
  ["student", "guardian"].forEach(k => {
    if (!r[k]) { out[k] = null; return; }
    const o = {};
    ["col", "first", "last"].forEach(x => { if (r[k][x]) { const m = move(r[k][x]); if (m === -1) missing = true; else o[x] = m; } });
    out[k] = Object.keys(o).length ? o : null;
  });
  return missing ? null : out;
}

// ======================================================================
// SIDEBAR ENDPOINTS (YourListsSidebar.html; the list buttons in Input Data Mode)
// ======================================================================

function od_withLock_(fn) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) return { error: "The sheet is busy. Try again in a moment." };
  try {
    return fn();
  } catch (e) {
    console.error("Your Lists: " + (e && e.stack ? e.stack : e));
    return { error: "❌ Script error: " + (e && e.message ? e.message : e) };
  } finally {
    lock.releaseLock();
  }
}

/** For Input Data Mode's step 5: one line per list (from the Backend only, no tab reads). */
function od_listSummaries_(ss) {
  return od_loadLists_(ss).filter(l => l.sheet).map(l => ({
    tabId: l.tabId, name: l.name, tabName: l.sheet.getName(), by: l.setup.by, ready: !!l.setup.ready,
    updated: l.updated, stats: l.setup.stats || null
  }));
}

/** "Add another list": makes the "📥 Name" tab and opens the setup sidebar on it. */
function od_createList(name) {
  const res = od_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    name = od_cleanName_(name);
    if (!name) return { error: "Give the list a name first." };
    const sheet = ss.insertSheet(od_uniqueTabName_(ss, OD_TAB_PREFIX + name));
    if (sheet.getMaxColumns() < 26) sheet.insertColumnsAfter(sheet.getMaxColumns(), 26 - sheet.getMaxColumns());
    od_resetListTab_(sheet, name);
    const lists = od_loadLists_(ss);
    lists.push({ tabId: sheet.getSheetId(), name: name, setup: od_emptySetup_(), updated: "", sheet: sheet });
    od_saveLists_(ss, lists);
    sheet.activate();
    return { tabId: sheet.getSheetId() };
  });
  if (res && res.tabId !== undefined) od_showWizard_(res.tabId);
  return res;
}

/** "Edit setup": the setup sidebar on an existing list. */
function od_openList(tabId) {
  const res = od_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const list = od_findList_(od_loadLists_(ss), tabId);
    if (!list || !list.sheet) return { error: "That list's tab is gone." };
    if (list.sheet.isSheetHidden()) list.sheet.showSheet();
    list.sheet.activate();
    return { tabId: list.tabId };
  });
  if (res && res.tabId !== undefined) od_showWizard_(res.tabId);
  return res;
}

/** "Paste new version": empties the list's tab (its setup and your decisions are kept) and waits for the paste. */
function od_repasteList(tabId) {
  const res = od_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const list = od_findList_(od_loadLists_(ss), tabId);
    if (!list || !list.sheet) return { error: "That list's tab is gone." };
    if (list.sheet.isSheetHidden()) list.sheet.showSheet();
    od_resetListTab_(list.sheet, list.name);
    list.sheet.activate();
    return { tabId: list.tabId };
  });
  if (res && res.tabId !== undefined) od_showWizard_(res.tabId);
  return res;
}

function od_renameList(tabId, name) {
  return od_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const lists = od_loadLists_(ss);
    const list = od_findList_(lists, tabId);
    name = od_cleanName_(name);
    if (!list || !list.sheet) return { error: "That list's tab is gone." };
    if (!name) return { error: "Give the list a name." };
    list.name = name;
    list.sheet.setName(od_uniqueTabName_(ss, OD_TAB_PREFIX + name, list.sheet));
    const title = String(list.sheet.getRange(1, 1).getValue());
    if (title.indexOf(OD_BAND_TITLE) === 0) list.sheet.getRange(1, 1).setValue(title.replace(/^📥 OWN DATA · .*? — /, OD_BAND_TITLE + name.toUpperCase() + " — "));
    od_saveLists_(ss, lists);
    return { ok: true };
  });
}

/** Deletes the list's tab and its setup (never the tab the user is standing on: they are moved off it first). */
function od_deleteList(tabId) {
  return od_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const lists = od_loadLists_(ss);
    const list = od_findList_(lists, tabId);
    if (!list) return { ok: true };
    if (list.sheet) {
      const active = ss.getActiveSheet();
      if (active && active.getSheetId() === list.sheet.getSheetId()) {
        const other = ss.getSheets().filter(s => !s.isSheetHidden() && s.getSheetId() !== list.sheet.getSheetId())[0];
        if (other) other.activate();
      }
      ss.deleteSheet(list.sheet);
    }
    od_saveLists_(ss, lists.filter(l => l !== list));
    idm_noteListsChanged_();
    return { ok: true };
  });
}

function od_showWizard_(tabId) {
  PropertiesService.getDocumentProperties().setProperty(OD_WIZARD_KEY, String(tabId));
  const html = htmlFromFile_("YourListsSidebar").setTitle("Your Lists"); // template: pulls in Styles.html
  SpreadsheetApp.getUi().showSidebar(html);
}

function od_wizardList_(ss) {
  const tabId = PropertiesService.getDocumentProperties().getProperty(OD_WIZARD_KEY);
  const lists = od_loadLists_(ss);
  return { lists: lists, list: od_findList_(lists, tabId) };
}

/** The setup sidebar's view: "paste" (waiting for the list) or "setup" (columns / variables). */
function od_getWizard() {
  return od_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const w = od_wizardList_(ss);
    if (!w.list || !w.list.sheet) return { error: "This list's tab is gone. Go back to Input Data Mode." , gone: true };
    const table = od_readTable_(w.list.sheet, null);
    if (!table.headers.length || !od_hasPaste_(w.list.sheet)) return od_pasteView_(w.list);
    return od_setupView_(ss, w.list, table, null);
  });
}

function od_pasteView_(list) {
  return { phase: "paste", list: { tabId: list.tabId, name: list.name, tabName: list.sheet.getName(), ready: !!list.setup.ready } };
}

/** Polled every 2 s while waiting: has something been pasted on the list's tab? */
function od_wizardCheck() {
  return od_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const w = od_wizardList_(ss);
    if (!w.list || !w.list.sheet) return { error: "This list's tab is gone.", gone: true };
    return { pasted: od_hasPaste_(w.list.sheet) };
  });
}

/** Moves the paste into place, then the column setup (a new version of a known list keeps its setup). */
function od_readPaste() {
  return od_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const w = od_wizardList_(ss);
    if (!w.list || !w.list.sheet) return { error: "This list's tab is gone.", gone: true };
    return od_readTab_(ss, w, false);
  });
}

/**
 * The list's tab -> its column setup (shared by a paste and an upload). A list
 * with no rows under the header is wrong data: the tab is emptied again (owner's
 * request 2026-11-12), so the right list can go straight in.
 */
function od_readTab_(ss, w, fromUpload) {
  const taken = od_takePaste_(w.list.sheet, w.list.name);
  if (!taken) return od_pasteView_(w.list);
  if (!taken.rows.length) {
    od_resetListTab_(w.list.sheet, w.list.name);
    return Object.assign(od_pasteView_(w.list), { cleared: true, error: "I found the header row, but no rows under it. " +
      (fromUpload ? "Upload the whole list, headers and rows." : "Paste the whole list, headers and rows.") + " The tab was cleared, so you can try again." });
  }
  w.list.updated = od_today_();
  w.list.setup.title = taken.title;
  od_saveLists_(ss, w.lists);
  return od_setupView_(ss, w.list, { headers: taken.headers, rows: taken.rows }, taken);
}

/** Uploaded CSV / TXT list (raw text, parsed here like Input Data Mode's uploads). */
function od_uploadCsv(text) {
  let rows;
  try {
    rows = idm_parseCsv_(text);
  } catch (e) {
    return { error: "That file couldn't be read as a CSV file (" + e.message + ")." };
  }
  return od_uploadRows(rows);
}

/**
 * Uploaded list (CSV parsed above, or Excel read in the sidebar): written on the
 * list's tab as plain text where a paste goes (row 4), then read exactly like a paste.
 */
function od_uploadRows(rows) {
  return od_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const w = od_wizardList_(ss);
    if (!w.list || !w.list.sheet) return { error: "This list's tab is gone.", gone: true };
    const cells = (rows || []).map(r => (r || []).map(c => (c === null || c === undefined) ? "" : String(c)));
    const width = cells.reduce((m, r) => Math.max(m, r.length), 0);
    if (!width || !cells.some(r => r.some(c => c.trim() !== ""))) {
      return Object.assign(od_pasteView_(w.list), { error: "That file is empty." });
    }
    const table = cells.map(r => r.concat(new Array(width - r.length).fill("")));
    const sheet = w.list.sheet;
    od_resetListTab_(sheet, w.list.name);
    const needRows = OD_HEADER_ROW + table.length, needCols = width + 3;
    if (sheet.getMaxRows() < needRows) sheet.insertRowsAfter(sheet.getMaxRows(), needRows - sheet.getMaxRows());
    if (sheet.getMaxColumns() < needCols) sheet.insertColumnsAfter(sheet.getMaxColumns(), needCols - sheet.getMaxColumns());
    sheet.getRange(OD_HEADER_ROW, 1, table.length, width).setNumberFormat("@").setValues(table);
    return od_readTab_(ss, w, true);
  });
}

/** Columns + variables for the sidebar (saved setup when it still fits, else the guesses). */
function od_setupView_(ss, list, table, taken) {
  const idx = od_indexFromSheet_(ss);
  const guess = od_guessRoles_(table.headers, table.rows, idx);
  const saved = list.setup;
  const savedRoles = saved.ready ? (taken ? od_rolesByHeader_(saved, table.headers) : saved.roles) : null;
  const roles = savedRoles || guess.roles;
  const by = savedRoles ? saved.by : guess.by;
  return {
    phase: "setup",
    list: { tabId: list.tabId, name: list.name, tabName: list.sheet.getName(), ready: !!saved.ready, title: saved.title || "" },
    kept: !!savedRoles,                 // a new version of the list: the old setup was kept
    headerRow: OD_HEADER_ROW,
    rowCount: table.rows.length,
    columns: guess.stats.map(s => ({ col: s.col, letter: od_colLetter_(s.col), header: s.header, samples: s.samples,
      hint: { osisKnown: s.osisKnown, n: s.n, comma: s.comma >= 0.5 } })),
    roles: roles,
    by: by,
    vars: od_varsFor_(table.headers, roles, saved.vars)
  };
}

/** Normalizes what the sidebar sends (roles as column numbers, variables). */
function od_cleanSetupInput_(p, headers) {
  const n = headers.length;
  const col = c => { c = Number(c); return c >= 1 && c <= n ? c : 0; };
  const nameRole = r => {
    if (!r) return null;
    if (r.first || r.last) { const f = col(r.first), l = col(r.last); return f || l ? { first: f, last: l } : null; }
    const c = col(r.col);
    return c ? { col: c } : null;
  };
  const roles = { osis: col(p.roles && p.roles.osis), student: nameRole(p.roles && p.roles.student), guardian: nameRole(p.roles && p.roles.guardian),
                  phone: col(p.roles && p.roles.phone), email: col(p.roles && p.roles.email), "class": col(p.roles && p.roles["class"]) };
  const roleCols = od_roleCols_(roles);
  const vars = (p.vars || []).filter(v => col(v.col) && roleCols.indexOf(col(v.col)) === -1).map(v => ({
    col: col(v.col), header: headers[col(v.col) - 1], label: od_cleanLabel_(v.label) || headers[col(v.col) - 1],
    show: !!v.show, column: !!v.column,
    blank: { mode: v.blank && v.blank.mode === "text" ? "text" : "nothing", text: toText_(v.blank && v.blank.text).slice(0, 100) },
    notOnList: { mode: v.notOnList && v.notOnList.mode === "text" ? "text" : "nothing", text: toText_(v.notOnList && v.notOnList.text).slice(0, 100) }
  }));
  return { by: p.by === "guardian" ? "guardian" : "student", roles: roles, vars: vars };
}

/**
 * Step 3 (Match): the rows that need a person (pick one / best guess / not found)
 * with their choices, the counts, and the list to pick from (Student List
 * students, or guardians for a guardian list).
 */
function od_matchPreview(p) {
  return od_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const w = od_wizardList_(ss);
    if (!w.list || !w.list.sheet) return { error: "This list's tab is gone.", gone: true };
    const table = od_readTable_(w.list.sheet, null);
    const input = od_cleanSetupInput_(p || {}, table.headers);
    const setup = Object.assign({}, w.list.setup, input);
    const idx = od_indexFromSheet_(ss);
    if (!idx.students.length) return { error: "The Student List is empty. Build the sheets once (🏗️ Build Sheets Only), then come back." };
    const results = od_matchAll_(setup, table.rows, idx);
    const by = setup.by;
    const label = (row, cells) => {
      if (cells.name) return cells.name.raw;
      if (cells.guardian) return cells.guardian.raw;
      return row.filter(Boolean).slice(0, 2).join(" · ");
    };
    const howCounts = {};
    results.forEach(r => { if (r.status === "ok") howCounts[r.how] = (howCounts[r.how] || 0) + 1; });
    const attention = [];
    results.forEach((r, i) => {
      if (r.status === "ok" || (r.status === "skip" && !(p && p.showSkipped))) return;
      const row = table.rows[i];
      const extra = [];
      if (r.cells.osis) extra.push("OSIS " + r.cells.osis);
      if (by === "student" && r.cells.guardian) extra.push("parent “" + r.cells.guardian.raw + "”");
      setup.vars.filter(v => v.show || v.column).slice(0, 1).forEach(v => { if (toText_(row[v.col - 1])) extra.push(v.label + " " + toText_(row[v.col - 1])); });
      const conf = t => (r.scores && r.scores[t] !== undefined ? r.scores[t] : null);
      attention.push({
        i: i, row: OD_HEADER_ROW + 1 + i, label: label(row, r.cells), sub: extra.join(" · "), status: r.status, reason: r.reason || "",
        target: r.target || null, targetText: r.target ? od_targetText_(by, r.target, idx) : null,
        confidence: r.confidence === undefined ? null : r.confidence, level: od_confLevel_(r.confidence),
        // pick: the choices; guess: other close ones; none: the closest few (all with their confidence)
        options: (r.options || []).map(t => Object.assign({ target: t, confidence: conf(t), level: od_confLevel_(conf(t)) }, od_targetText_(by, t, idx)))
      });
    });
    const picker = by === "student"
      ? idx.students.slice().sort(byStudentLastName_("name")).map(st => ({ target: st.key, text: st.osis + " · " + st.name, sub: st.guardianText }))
      : idx.persons.slice().sort((a, b) => compareNatural_(a.last + " " + a.first, b.last + " " + b.first))
          .map(pp => ({ target: pp.key, text: pp.name, sub: pp.kids.map(k => k.name).join(", ") }));
    return { phase: "match", by: by, stats: od_stats_(results), how: howCounts, attention: attention, picker: picker,
             levels: { high: OD_CONF_HIGH_, medium: OD_CONF_MEDIUM_, min: OD_CONF_MIN_ } };
  });
}

/**
 * Save: the setup, your decisions ({ rowIndex: target | "skip" }), the "✔ Matched"
 * columns on the tab. The Student List / Parents Divided columns and the phone
 * contacts follow at the next build.
 */
function od_saveList(p) {
  return od_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const w = od_wizardList_(ss);
    if (!w.list || !w.list.sheet) return { error: "This list's tab is gone.", gone: true };
    const list = w.list;
    const table = od_readTable_(list.sheet, null);
    const input = od_cleanSetupInput_(p || {}, table.headers);
    const name = od_cleanName_(p && p.name);
    const setup = Object.assign(list.setup, input, { ready: true, headers: table.headers.slice() });
    const idx = od_indexFromSheet_(ss);
    const before = od_matchAll_(setup, table.rows, idx);
    const decisions = (p && p.decisions) || {};
    Object.keys(decisions).forEach(i => {
      const r = before[Number(i)];
      const t = decisions[i];
      if (!r || !t) return;
      if (t === "auto") { delete setup.picks[r.key]; return; }
      if (t === "skip" || (setup.by === "student" ? idx.byOsis.has(t) : idx.pByKey.has(t))) setup.picks[r.key] = t;
    });
    const results = od_matchAll_(setup, table.rows, idx);
    setup.stats = od_stats_(results);
    if (name && name !== list.name) {
      list.name = name;
      list.sheet.setName(od_uniqueTabName_(ss, OD_TAB_PREFIX + name, list.sheet));
    }
    if (!list.updated) list.updated = od_today_();
    od_writeToolColumns_(list.sheet, list, table, results, idx);
    od_saveLists_(ss, w.lists);
    idm_noteListsChanged_();
    return { phase: "done", list: { tabId: list.tabId, name: list.name, tabName: list.sheet.getName() }, stats: setup.stats,
             inputMode: !!idm_loadState_() };
  });
}

/**
 * Migration Wizard (Menus.gs ma_migrateSimpleSheets): copies every list of the old
 * tracker (its "📥" tab + its setup, your picks included) into this one. The old sheet
 * is only read. A list with the same name already here (the wizard run again) is
 * replaced. Returns how many lists were copied.
 */
function od_migrateLists_(oldSS, newSS) {
  const oldRows = readBackendBlockFromOtherSpreadsheet_(oldSS, "OWN_LISTS").slice(1)
    .filter(r => String(r[0]).trim() !== "");
  if (!oldRows.length) return 0;
  let lists = od_loadLists_(newSS);
  let moved = 0;
  oldRows.forEach(r => {
    const oldSheet = od_sheetById_(oldSS, r[0]);
    if (!oldSheet) return; // its tab was deleted on the old sheet
    const name = toText_(r[1]) || oldSheet.getName().replace(OD_TAB_PREFIX, "");
    let setup = null;
    try { setup = JSON.parse(String(r[2] || "{}")); } catch (e) { setup = null; }

    // Same list already here: replace it (never the tab the user is on)
    lists.filter(l => l.name.toLowerCase() === name.toLowerCase()).forEach(l => {
      if (!l.sheet) return;
      const active = newSS.getActiveSheet();
      if (active && active.getSheetId() === l.sheet.getSheetId()) {
        const other = newSS.getSheets().filter(s => !s.isSheetHidden() && s.getSheetId() !== l.sheet.getSheetId())[0];
        if (other) other.activate();
      }
      newSS.deleteSheet(l.sheet);
    });
    lists = lists.filter(l => l.name.toLowerCase() !== name.toLowerCase());

    const copy = oldSheet.copyTo(newSS);
    copy.setName(od_uniqueTabName_(newSS, oldSheet.getName(), copy));
    lists.push({ tabId: copy.getSheetId(), name: name, setup: od_fillSetup_(setup), updated: toText_(r[3]), sheet: copy });
    moved++;
  });
  od_saveLists_(newSS, lists);
  return moved;
}

/** Back to Input Data Mode's sidebar (when it is on), else just close. */
function od_backToInputMode() {
  PropertiesService.getDocumentProperties().deleteProperty(OD_WIZARD_KEY);
  if (idm_loadState_()) {
    openInputDataMode();
    return { reopened: true };
  }
  return { reopened: false };
}

/** "Build now" from the setup sidebar. */
function od_build() {
  PropertiesService.getDocumentProperties().deleteProperty(OD_WIZARD_KEY);
  runBuildOnly();
  return true;
}
