// ======================================================================
// FILE: Z#InputDataMode.gs (+ Z#InputDataSidebar.html)
// PURPOSE: "📥 Input Data Mode" — walks the user through loading their three
// exports (Student Data / ATS, Parent Square, NYCSA) one at a time:
//   1. Every visible tab is hidden (and remembered).
//   2. A temporary tab ("Student Data", then "Parent Square Data", then
//      "NYCSA Data") is created; the user pastes anywhere on it, or uploads
//      the file in the sidebar instead.
//   3. The sidebar checks the tab every 2 seconds. Once data is there it finds
//      the header row, maps the columns, checks for numbers Excel cut off
//      (7.19E+09 / ####), and replaces that section of Raw Data.
//   4. The temp tab is deleted and the next one is created.
//   5. At the end (or on Exit) temp tabs are deleted and the hidden tabs come back.
// Step 4 "Site / Class" is filled in by hand instead: one row per class code
// (from the ATS data), with Teacher Name and Site (optional, dropdown that grows
// as new sites are typed). Save writes Raw Data BH:BI (MAKE OWN DATA) as one row
// per site: BH = site, BI = that site's classes, one per line ("Y25: Carl" or
// just "Y24" with no teacher); read by parseClassCell_ (Build.gs).
//
// SELF-CONTAINED ("Z#" = trying out a new feature): the only hooks into the rest
// of the app are two guarded lines in onOpen (Menus.gs):
//     if (typeof addInputDataModeMenuItem_ === "function") addInputDataModeMenuItem_(menu);
// Without this file (and its sidebar) those lines are no-ops.
// It uses shared helpers (fixRawDataNumberFormatting_, forceVibrantColors_,
// applyBandingToRawData_, runBuildOnly); shared code never relies on this file.
// ======================================================================

var IDM_STATE_KEY = 'INPUT_DATA_MODE_STATE';

// 🐞 TEMPORARY DEBUG: every sidebar call returns a step-by-step log (shown in the
// sidebar's Debug panel) and script errors are shown instead of retried silently.
// Set to false (or delete the idm_log_ lines) once Input Data Mode works.
var IDM_DEBUG = true;
var IDM_DBG = null;
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
      "Click <b>✅ Save Site / Class &amp; finish</b> below when you're done."
    ]
  }
];
var IDM_SITECLASS_UPDATED_KEY = 'INPUT_DATA_MODE_SITECLASS_UPDATED';

// ---------------------------------------------------------------------
// MENU + LAUNCHER
// ---------------------------------------------------------------------

/** Called from onOpen (guarded). Adds the menu item; reminds the user if the mode was left on. */
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
  const html = HtmlService.createHtmlOutputFromFile('Z#InputDataSidebar').setTitle('📥 Input Data Mode');
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
    idm_debugSheets_(ss, state, temp);
    if (IDM_STEPS[state.step].manual) return idm_view_(ss, state);
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
             anyImported: idm_anyImported_(results), summary: idm_summary_(results) };
  });
}

function idm_exit() {
  return idm_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const state = idm_loadState_();
    if (!state) return idm_view_(ss, null);
    const results = state.results;
    idm_cleanup_(ss, state, false);
    return { active: false, finished: true, exited: true, anyImported: idm_anyImported_(results), summary: idm_summary_(results) };
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
  if (step.manual) return idm_view_(ss, state, { type: 'error', code: 'manual', message: "This step is filled in on the tab — no file needed.", fromUpload: true });
  const fail = (code, message, extra) => idm_view_(ss, state, Object.assign({ type: 'error', code: code, message: message, fromUpload: fromUpload }, extra || {}));

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
      "Clear it and paste the " + step.zone + " export (or Skip to the " + found.step.zone + " step).");
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
        { canUseAnyway: true });
    }
  }

  const raw = idm_rawSheet_(ss);
  if (!raw) return fail('noRaw', 'The "Raw Data" tab is missing.');
  idm_log_("writing to " + raw.getName() + " (maxRows=" + raw.getMaxRows() + ", maxCols=" + raw.getMaxColumns() + ")");
  idm_writeSection_(raw, step, mapped.rows);
  idm_log_("Raw Data written");

  state.results[step.key] = { rows: mapped.rows.length };
  return idm_advance_(ss, state, { type: 'imported', message: "✅ " + step.zone + ": " + mapped.rows.length + " row(s) moved to Raw Data." });
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

  sheet.clear();
  sheet.getDataRange().clearDataValidations();
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
    return { active: false, finished: true, notice: notice, anyImported: idm_anyImported_(results), summary: idm_summary_(results) };
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
  else idm_resetTempSheet_(sheet);
}

function idm_resetTempSheet_(sheet) {
  sheet.getDataRange().clearDataValidations();
  sheet.clear();
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
  (state.hiddenIds || []).forEach(id => {
    const s = idm_sheetById_(ss, id);
    if (s) { try { s.showSheet(); } catch (e) { console.warn(e.message); } }
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
    step: { key: step.key, zone: step.zone, required: step.required, howTo: step.howTo, manual: !!step.manual },
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
  IDM_DBG = [];
  IDM_DBG.t0 = Date.now();
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
  if (IDM_DEBUG && result && typeof result === "object") {
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

/** Snapshot of what the server sees, for the Debug panel. */
function idm_debugSheets_(ss, state, temp) {
  if (!IDM_DBG) return;
  try {
    const active = ss.getActiveSheet();
    idm_log_("state: " + JSON.stringify(state));
    idm_log_("active tab (as the script sees it): " + (active ? active.getName() + " #" + active.getSheetId() : "none"));
    idm_log_("tabs: " + ss.getSheets().map(s => s.getName() + " #" + s.getSheetId() + (s.isSheetHidden() ? " (hidden)" : "")).join(" | "));
    if (temp) {
      idm_log_("temp tab: " + temp.getName() + " #" + temp.getSheetId() + " lastRow=" + temp.getLastRow() +
        " lastCol=" + temp.getLastColumn() + " maxRows=" + temp.getMaxRows() + " maxCols=" + temp.getMaxColumns() +
        " A1=" + JSON.stringify(String(temp.getRange("A1").getValue())));
    }
  } catch (e) { idm_log_("debug snapshot failed: " + e.message); }
}
