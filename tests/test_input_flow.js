// Input Data Mode / Your Lists / Welcome tab flows on a fake spreadsheet (lib/fakesheet.js):
//  - a wrong paste is cleared off the tab (and "Use it anyway" still works from the kept rows)
//  - a list can be uploaded (CSV / Excel rows) like a paste; a list with no rows is cleared
//  - step 5 builds the sheets first on a new sheet / after new data, and hides what the build showed
//  - the 👋 Welcome tab on a new copy: only tab shown, hidden by setup, deleted afterwards
//  - Reset to Display Mode: a fresh Welcome tab (permissions explained first) is the only tab, kept until setup
const fs = require('fs'), path = require('path'), assert = require('assert');
const { loadProject, ROOT } = require('./lib/load');
const { makeSpreadsheet, fakeSpreadsheetApp } = require('./lib/fakesheet');

function parseCsv(text) {
  const rows = []; let row = [], cell = '', q = false;
  text = text.replace(/^﻿/, '');
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) { if (ch === '"' && text[i + 1] === '"') { cell += '"'; i++; } else if (ch === '"') q = false; else cell += ch; }
    else if (ch === '"') q = true;
    else if (ch === ',') { row.push(cell); cell = ''; }
    else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(cell); rows.push(row); row = []; cell = ''; }
    else cell += ch;
  }
  if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
  return rows;
}
const csv = name => parseCsv(fs.readFileSync(path.join(ROOT, 'test-data', name), 'utf8'));
// What a values-only paste from Excel gives: numbers as numbers
const asPasted = rows => rows.map(r => r.map(c => (/^\d+$/.test(c) && c.length < 16) ? Number(c) : c));

const TABS = ["Contact Log", "Combined Contact Tracking", "PCAR", "Directory", "Raw Data", "Student List", "Notes", "Backend"];
function newSheet() {
  const ss = makeSpreadsheet(TABS);
  ss.getSheetByName("Backend").hideSheet();
  const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) } });
  return { ss, g };
}
/** Input Data Mode switched on at a step, other tabs hidden (like openInputDataMode). */
function startMode(ss, g, stepIndex) {
  const state = { step: stepIndex, tempId: null, hiddenIds: [], prevActiveId: null, results: {} };
  g.idm_ensureTempSheet_(ss, state);
  g.idm_hideOtherSheets_(ss, state);
  g.idm_saveState_(state);
  return g.idm_sheetById_(ss, state.tempId);
}
// An empty paste tab: "PASTE HERE" in A1 and the way back to the sidebar in A2 (neither is data)
const tabEmpty = t => t.getLastRow() === 2 && t.getLastColumn() === 1 && t.getRange("A1").getValue() === "PASTE HERE" &&
  /App Menu → 📥 Input Data Mode/.test(t.getRange("A2").getValue());

// ---------------------------------------------------------------- wrong paste is cleared
{
  const { ss, g } = newSheet();
  const temp = startMode(ss, g, 0); // Student Data step
  temp._put(asPasted(csv("2 - Parent Square.csv")), 3, 2);
  let v = g.idm_processPaste();
  assert.strictEqual(v.notice.code, 'wrongType', 'Parent Square on the ATS step');
  assert.strictEqual(v.notice.cleared, true, 'wrong paste: cleared flag');
  assert(tabEmpty(temp), 'wrong paste: tab back to PASTE HERE');
  assert.strictEqual(g.idm_check().pending, false, 'wrong paste: nothing waiting after the clear (no loop)');

  temp._put([["just", "some"], ["words", "here"]], 1, 1);
  v = g.idm_processPaste();
  assert.strictEqual(v.notice.code, 'noHeader');
  assert(v.notice.cleared && tabEmpty(temp), 'no header row: cleared');

  // Cut-off numbers: cleared too, but the rows come back for "Use it anyway"
  const broken = csv("1c - Student Data (ATS) - BROKEN numbers cut off.csv");
  temp._put(asPasted(broken), 1, 1);
  temp._set(1, 40, new Date(2026, 10, 12)); // a date cell somewhere in the paste
  v = g.idm_processPaste();
  assert.strictEqual(v.notice.code, 'cutoff');
  assert(v.notice.cleared && tabEmpty(temp), 'cut-off: cleared');
  assert(Array.isArray(v.notice.heldRows) && v.notice.heldRows.length === broken.length, 'cut-off: rows kept');
  assert(v.notice.heldRows.every(r => r.every(c => !(c instanceof Date))), 'cut-off: no Date objects (google.script.run cannot send them)');
  const sent = JSON.parse(JSON.stringify(v.notice.heldRows)); // the trip to the sidebar and back
  v = g.idm_submitRows(sent, true);
  assert.strictEqual(v.notice.type, 'imported', 'Use it anyway: imported');
  const raw = ss.getSheetByName("Raw Data");
  assert.strictEqual(String(raw.getRange(7, 3).getValue()), String(broken[1][2]), 'Use it anyway: first OSIS in Raw Data C7');
  assert.strictEqual(raw.getRange(6, 3).getValue(), "Student ID", 'Raw Data header row 6');
  assert.strictEqual(g.idm_loadState_().dataChanged, true, 'new data -> step 5 will build first');

  // An upload with the wrong export is not "cleared" (nothing was on the tab)
  temp._put([["keep"]], 5, 5);
  v = g.idm_submitRows(asPasted(csv("1 - Student Data (ATS).csv")), false); // now on the Parent Square step
  assert.strictEqual(v.notice.code, 'wrongType');
  assert.notStrictEqual(v.notice.cleared, true, 'upload: not cleared');
  assert.strictEqual(temp.getRange(5, 5).getValue(), "keep", 'upload: the tab is left alone');
  console.log("wrong paste cleared, Use it anyway from kept rows: ok");
}

// ---------------------------------------------------------------- Your Lists upload
function listsSheet() {
  const { ss, g } = newSheet();
  let lists = [];
  g.od_loadLists_ = () => lists;
  g.od_saveLists_ = (s, ls) => { lists = ls; };
  g.od_indexFromSheet_ = () => g.od_index_([], []);
  const sheet = ss.insertSheet("📥 Bus Routes");
  g.od_resetListTab_(sheet, "Bus Routes");
  lists.push({ tabId: sheet.getSheetId(), name: "Bus Routes", setup: g.od_emptySetup_(), updated: "", sheet: sheet });
  g.__props["OD_WIZARD_TAB"] = String(sheet.getSheetId());
  return { ss, g, sheet, lists: () => lists };
}
{
  const { g, sheet } = listsSheet();
  const file = csv("5a - Your Lists - Bus Routes (student list).csv");
  const v = g.od_uploadRows(asPasted(file));
  assert.strictEqual(v.phase, 'setup', 'upload -> column setup');
  const h = g.od_findHeaderRow_(file.filter(r => r.some(x => String(x).trim() !== "")));
  const dataRows = file.filter(r => r.some(x => String(x).trim() !== "")).length - h - 1;
  assert.strictEqual(v.rowCount, dataRows, 'upload: every row');
  assert.strictEqual(sheet.getRange(4, 1).getValue(), v.columns[0].header, 'upload: header on row 4');
  console.log("Your Lists upload (Excel rows): ok, " + v.rowCount + " rows");
}
{
  const { g } = listsSheet();
  const v = g.od_uploadCsv("Name,Route\nBart Simpson,12\nLisa Simpson,12\n");
  assert.strictEqual(v.phase, 'setup', 'CSV upload -> setup');
  assert.strictEqual(v.rowCount, 2);
  console.log("Your Lists upload (CSV): ok");
}
{
  const { g, sheet } = listsSheet();
  let v = g.od_uploadRows([["Name", "Route", "Stop"]]);
  assert.strictEqual(v.phase, 'paste', 'header only: back to waiting');
  assert(v.cleared && /no rows/.test(v.error), 'header only: error + cleared');
  assert.strictEqual(sheet.getRange(4, 1).getValue(), "PASTE HERE", 'header only: tab reset');
  assert.strictEqual(g.od_hasPaste_(sheet), false, 'header only: nothing waiting (no loop)');
  v = g.od_uploadRows([[""], ["", ""]]);
  assert(/empty/.test(v.error), 'empty file');
  // A pasted header without rows is cleared too (it used to stay and be re-read every 2 s)
  sheet._put([["Name", "Route"]], 4, 1);
  v = g.od_readPaste();
  assert(v.cleared && g.od_hasPaste_(sheet) === false, 'pasted header only: cleared');
  console.log("Your Lists: a list with no rows is cleared: ok");
}

// ---------------------------------------------------------------- build before step 5
{
  const { ss, g } = newSheet();
  const raw = ss.getSheetByName("Raw Data");
  raw._put([["School Year", "DBN", "Student ID"], ["2025-2026", "75M094", 123456789]], 6, 1);
  const temp = startMode(ss, g, 4); // Your Lists step
  let builds = 0;
  g.buildAllDerivedSheets = () => { builds++; ss.insertSheet("Parents Divided"); }; // a new tab shows + becomes active
  g.formatDirectoryAuto_ = () => {};
  let v = g.idm_getState();
  assert.strictEqual(v.needsBuild, true, 'new sheet (Student List empty): build first');
  v = g.idm_buildForLists();
  assert.strictEqual(builds, 1, 'built once');
  assert(v.listsBuild && v.listsBuild.ok, 'build ok');
  assert.strictEqual(v.needsBuild, false, 'not again');
  assert(ss.getSheetByName("Parents Divided").isSheetHidden(), 'the tab the build made is hidden again');
  assert.strictEqual(ss.getActiveSheet().getSheetId(), temp.getSheetId(), 'back on the temp tab');
  assert(g.idm_loadState_().hiddenIds.indexOf(ss.getSheetByName("Parents Divided").getSheetId()) !== -1, 'shown again at the end');
  assert.deepStrictEqual(ss.getSheets().filter(s => !s.isSheetHidden()).map(s => s.getName()), [temp.getName()], 'only the temp tab shows');
  assert.strictEqual(g.idm_getState().needsBuild, false, 'still not needed');

  // New data later in the same run: build again on the way back to step 5
  const st = g.idm_loadState_(); st.dataChanged = true; g.idm_saveState_(st);
  assert.strictEqual(g.idm_getState().needsBuild, true, 'new data: build again');
  g.buildAllDerivedSheets = () => { throw new Error("boom"); };
  v = g.idm_buildForLists();
  assert(v.listsBuild && !v.listsBuild.ok && /boom/.test(v.listsBuild.message), 'a failed build is reported');
  assert.strictEqual(v.needsBuild, false, 'a failed build is not retried in a loop');
  console.log("build before Your Lists: ok");
}
{
  const { ss, g } = newSheet();
  startMode(ss, g, 4);
  let builds = 0;
  g.buildAllDerivedSheets = () => { builds++; };
  const v = g.idm_buildForLists();
  assert.strictEqual(builds, 0, 'no Student Data: no build');
  assert(!v.listsBuild.ok && /Student Data/.test(v.listsBuild.message), 'no Student Data: says to load it first');
  console.log("build before Your Lists without Student Data: ok");
}
{
  const { ss, g } = newSheet();
  startMode(ss, g, 0);
  assert.strictEqual(g.idm_getState().needsBuild, false, 'only step 5 builds');
}

// ---------------------------------------------------------------- Input Data Mode resumed
{
  // Opened again while on (Getting Started, the open trigger): on its tab, every other tab hidden
  const { ss, g } = newSheet();
  g.htmlFromFile_ = () => ({ setTitle() { return this; }, setWidth() { return this; } });
  const vis = () => ss.getSheets().filter(s => !s.isSheetHidden()).map(s => s.getName());
  g.openInputDataMode();
  assert.deepStrictEqual(vis(), ["Student Data"], 'started: only the Student Data tab');
  ss.getSheetByName("Contact Log").showSheet(); ss.getSheetByName("PCAR").showSheet();
  ss.getSheetByName("Contact Log").activate();
  g.openInputDataMode();
  assert.strictEqual(ss.getActiveSheet().getName(), "Student Data", 'resumed: back on its tab');
  assert.deepStrictEqual(vis(), ["Student Data"], 'resumed: the other tabs hidden again');
  g.idm_exit();
  assert(vis().indexOf("Contact Log") !== -1 && vis().indexOf("PCAR") !== -1, 'exit: they come back');
  console.log("Input Data Mode resumed: ok");
}
// ---------------------------------------------------------------- a new user, step by step
{
  // The whole first-time path on a new copy, as the owner walked it (2026-11-12 screenshots):
  // open -> Initial Setup -> "Is this your first PC Tracker?" -> start fresh -> Input Data Mode
  const ss = makeSpreadsheet(TABS);
  ss.getSheetByName("Backend").hideSheet();
  const app = fakeSpreadsheetApp(ss);
  let dialog = "", sidebars = [];
  app._ui.showModalDialog = (h) => { dialog = h._html; };
  app._ui.showSidebar = (h) => { sidebars.push(h._name); };
  const g = loadProject({ extra: { SpreadsheetApp: app }, noTrialFiles: true }); // no z_ and no Z# files, like the owner's test copy
  assert.strictEqual(typeof g.buildContactLogMenu_, "undefined", 'loaded without Z#ContactLog');
  g.isAutoMigrateOn_ = () => false;
  g.include_ = () => "";
  g.zg_safeGuideLink_ = () => ""; g.zg_safeExampleLink_ = () => ""; g.getGuideLink_ = () => "https://example.org/guide";
  g.HtmlService = { createHtmlOutput: h => { const o = { _html: h }; o.setWidth = () => o; o.setHeight = () => o; return o; } };
  g.htmlFromFile_ = name => { const o = { _name: name }; o.setTitle = () => o; o.setWidth = () => o; return o; };
  const vis = () => ss.getSheets().filter(s => !s.isSheetHidden()).map(s => s.getName());

  // 1. The copy opens: only the Welcome tab
  g.onOpen({ source: ss });
  assert.deepStrictEqual(vis(), ["👋 Welcome"], '1. open: only the Welcome tab');

  // 2. 🚨 Initial Setup: the window opens at once and nothing else happens in that click
  let setups = 0;
  const realSteps = g.runInitialSetupSteps_;
  g.runInitialSetupSteps_ = function () { setups++; return realSteps.apply(this, arguments); };
  g.runInitialSetupWrapper();
  assert(/Setting up your tracker/.test(dialog) && /runSetup\(\);/.test(dialog), '2. the setup window, which starts the setup itself');
  assert.strictEqual(setups, 0, '2. the click itself does no setup work');
  assert(!/example\.org\/guide/.test(dialog), '2. no link look-ups before the window shows');
  assert.deepStrictEqual(vis(), ["👋 Welcome"], '2. no tab changes yet');

  // 3. The window's setup call: still ONLY the Welcome tab (no Contact Log / PCAR / Combined), now "You're set up"
  const setupReply = JSON.parse(JSON.stringify(g.zg_runSetup()));
  assert.strictEqual(setupReply.migrate, false);
  assert(setupReply.links && setupReply.links.downloadDataDoc, '3. the window gets its links from the setup call');
  assert.strictEqual(g.__props.SETUP_COMPLETE, 'true');
  assert.deepStrictEqual(vis(), ["👋 Welcome"], '3. after setup: still only the Welcome tab');
  const w = ss.getSheetByName("👋 Welcome");
  assert.strictEqual(ss.getActiveSheet(), w, '3. on the Welcome tab');
  assert.strictEqual(w.getRange("C3").getValue(), "You're set up", '3. it says what to do next');
  assert.deepStrictEqual(sidebars, ["GettingStartedSidebar"], '3. the checklist opens');
  // A second setup window (clicked twice) does not run it again
  g.zg_runSetup();
  assert.strictEqual(setups, 1, '3. setup runs once');

  // 4. "Yes, start fresh" -> "Start Input Data Mode": the pop-up only opens the sidebar...
  g.zg_action('inputData');
  assert.deepStrictEqual(sidebars.slice(-1), ["InputDataSidebar"], '4. Input Data Mode sidebar');
  assert.deepStrictEqual(vis(), ["👋 Welcome"], '4. the pop-up call itself does not touch the tabs');
  // ...whose first call switches the tabs
  const v = g.idm_open();
  assert(v.active && v.stepIndex === 0, '4. step 1');
  assert.deepStrictEqual(vis(), ["Student Data"], '4. only the Student Data tab');
  assert.strictEqual(ss.getActiveSheet().getName(), "Student Data", '4. on it');
  assert.strictEqual(g.idm_loadState_().prevActiveId, ss.getSheetByName("Contact Log").getSheetId(), '4. comes back to the Contact Log later, not the Welcome tab');
  const temp = ss.getSheetByName("Student Data");
  assert(tabEmpty(temp), '4. PASTE HERE + the way back if the panel is closed');
  assert.strictEqual(g.idm_check().pending, false, '4. the hint is not data');

  // 5. Something shows other tabs again and moves the user (that was the screenshot): the next check fixes it
  w.showSheet(); ss.getSheetByName("Contact Log").showSheet(); w.activate();
  g.idm_check();
  assert.deepStrictEqual(vis(), ["Student Data"], '5. the check hides them again');
  assert.strictEqual(ss.getActiveSheet().getName(), "Student Data", '5. and puts the user back');
  // The sidebar opened again (closed by accident, Getting Started, Your Lists' Back): same screen
  ss.getSheetByName("PCAR").showSheet(); ss.getSheetByName("PCAR").activate();
  g.idm_open();
  assert.deepStrictEqual(vis(), ["Student Data"], '5. reopened: only its tab');

  // 6. A paste below the label + hint is read; the label / hint are not
  temp._put(asPasted(csv("1 - Student Data (ATS).csv")), 4, 2);
  assert.strictEqual(g.idm_check().pending, true, '6. a paste is seen');
  const imported = g.idm_processPaste();
  assert.strictEqual(imported.notice.type, 'imported', '6. imported: ' + JSON.stringify(imported.notice));

  // 7. Exit before building: back on the Welcome tab, nothing else shown
  const tempId = temp.getSheetId();
  g.idm_exit();
  const visNow = () => ss.getSheets().filter(s => !s.isSheetHidden() && s.getSheetId() !== tempId).map(s => s.getName()); // the mode's tab goes in the next call
  assert.deepStrictEqual(visNow(), ["👋 Welcome"], '7. exit, not built: the Welcome tab, and the tabs shown by mistake stay hidden');
  assert.strictEqual(ss.getActiveSheet().getName(), "👋 Welcome");

  // 8. The first build: the everyday tabs come out, the user moves to the Contact Log, the Welcome tab goes
  g.studentListIsBuilt_ = () => true;
  g.welcomeTabAfterBuild_(ss);
  assert.deepStrictEqual(visNow(), ["Contact Log", "Combined Contact Tracking", "PCAR", "Directory"], '8. the default tabs');
  assert.strictEqual(ss.getActiveSheet().getName(), "Contact Log", '8. on the Contact Log');
  g.onOpen({ source: ss });
  assert.strictEqual(ss.getSheetByName("👋 Welcome"), null, '8. the next open deletes the Welcome tab');
  console.log("a new user, step by step: ok");
}
{
  // Step 5 builds inside Input Data Mode: the everyday tabs wait for the end of the mode
  const ss = makeSpreadsheet(TABS);
  ss.getSheetByName("Backend").hideSheet();
  const app = fakeSpreadsheetApp(ss);
  const g = loadProject({ extra: { SpreadsheetApp: app } });
  g.isAutoMigrateOn_ = () => false;
  g.zg_safeGuideLink_ = () => ""; g.zg_safeExampleLink_ = () => "";
  g.onOpen({ source: ss });
  g.runInitialSetup_();
  g.idm_open();
  const vis = () => ss.getSheets().filter(s => !s.isSheetHidden()).map(s => s.getName());
  g.studentListIsBuilt_ = () => true;
  g.welcomeTabAfterBuild_(ss); // the build at step 5 ends
  assert.deepStrictEqual(vis(), ["Student Data"], 'during the mode: still only its tab');
  assert.strictEqual(ss.getSheetByName("👋 Welcome"), null, 'the (hidden) Welcome tab is deleted');
  g.idm_exit();
  assert.deepStrictEqual(vis().filter(n => n !== "Student Data"), ["Contact Log", "Combined Contact Tracking", "PCAR", "Directory"], 'end of the mode: the everyday tabs');
  assert.strictEqual(ss.getActiveSheet().getName(), "Contact Log", 'on the Contact Log');
  console.log("first build inside Input Data Mode: ok");
}
{
  // The 10 PM rebuild on a new copy with nothing loaded yet: skipped quietly, the Welcome tab stays
  const ss = makeSpreadsheet(TABS);
  ss.getSheetByName("Backend").hideSheet();
  const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) }, noTrialFiles: true });
  g.isAutoMigrateOn_ = () => false;
  g.zg_safeGuideLink_ = () => ""; g.zg_safeExampleLink_ = () => "";
  g.onOpen({ source: ss });
  g.runInitialSetup_();
  let builds = 0; g.buildAllDerivedSheets = () => { builds++; };
  g.nightlyMasterSyncRun_();
  assert.strictEqual(builds, 0, 'no build without student data');
  assert.strictEqual(g.__props.NIGHTLY_FAILED_MSG, undefined, 'no "didn\'t finish" message');
  // A build that ends with no students (nothing loaded) leaves the guide alone
  assert.strictEqual(g.welcomeTabAfterBuild_(ss), false);
  assert.deepStrictEqual(ss.getSheets().filter(s => !s.isSheetHidden()).map(s => s.getName()), ["👋 Welcome"], 'still only the Welcome tab');
  console.log("nightly on a new copy: skipped: ok");
}
{
  // Clicking 🏗️ Build Sheets Only (menu or checklist) before loading anything: a plain note, no empty build
  const ss = makeSpreadsheet(TABS);
  ss.getSheetByName("Backend").hideSheet();
  const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) }, noTrialFiles: true });
  const cache = {};
  g.CacheService = { getScriptCache: () => ({ put: (k, v) => { cache[k] = v; }, get: k => cache[k] || null, remove: k => { delete cache[k]; } }),
    getDocumentCache: () => ({ put() {}, get: () => null, remove() {}, getAll: () => ({}), putAll() {}, removeAll() {} }), getUserCache: () => ({ put() {}, get: () => null }) };
  let builds = 0; g.buildAllDerivedSheets = () => { builds++; };
  g.startBuildOnlyProcessRun_();
  const st = JSON.parse(cache.SYNC_UI_STATE);
  assert.strictEqual(builds, 0, 'nothing built');
  assert(st.phase === 'killed' && /Input Data Mode/.test(st.notice) && !st.error, 'a note pointing at Input Data Mode, not "failed": ' + cache.SYNC_UI_STATE);
  const ps = fs.readFileSync(path.join(ROOT, 'ProgressSidebar.html'), 'utf8');
  assert(/state\.notice/.test(ps), 'the Build sidebar shows the note');
  console.log("Build before any data: a note: ok");
}
{
  // Site / Class: "Save" with the tab emptied by mistake must not erase the saved sites and classes
  const { ss, g } = newSheet();
  const raw = ss.getSheetByName("Raw Data");
  raw._maxCols = 70;
  raw._put([["Site", "Class"], ["Site 1", "Y25: Carl\nY24"]], 6, 60);
  const temp = startMode(ss, g, 3); // Site / Class step: the tab is filled from BH:BI
  assert.strictEqual(temp.getRange("A2").getValue(), "Y24", 'tab prefilled from Raw Data');
  for (let r = 2; r <= 10; r++) for (let c = 1; c <= 3; c++) temp._set(r, c, ""); // the user clears it
  let v = g.idm_saveSiteClass();
  assert(v.notice && v.notice.code === 'emptySiteClass', 'empty tab + Save: a warning, not a wipe');
  assert.strictEqual(raw.getRange(7, 60).getValue(), "Site 1", 'saved sites kept');
  v = g.idm_saveAndExit();
  assert(v.exited, 'Save and exit with nothing typed just exits');
  assert.strictEqual(raw.getRange(7, 61).getValue(), "Y25: Carl\nY24", 'still kept after the exit');
  // A new sheet, nothing typed, nothing saved before: Save moves on (like Skip)
  const n = newSheet();
  startMode(n.ss, n.g, 3);
  v = n.g.idm_saveSiteClass();
  assert(v.active && v.stepIndex === 4, 'nothing to save on a new sheet: on to Your Lists');
  console.log("Site / Class: an empty Save never erases: ok");
}
{
  // "Finish setup": the checklist stops opening by itself, a pop-up says where it is
  const ss = makeSpreadsheet(TABS);
  const app = fakeSpreadsheetApp(ss);
  const g = loadProject({ extra: { SpreadsheetApp: app }, props: { SETUP_COMPLETE: 'true' } });
  g.zg_finishSetup();
  assert.strictEqual(g.__props.ZG_AUTO_OPEN, 'off');
  assert(app._ui._alerts.some(a => /Getting Started/.test(JSON.stringify(a))), 'pop-up names ✅ Getting Started');
  let opened = 0; g.openGuideChecklist = () => opened++;
  g.guideOnOpen_();
  assert.strictEqual(opened, 0, 'no longer opens by itself');
  const html = fs.readFileSync(path.join(ROOT, 'GettingStartedSidebar.html'), 'utf8');
  assert(/zg_finishSetup/.test(html) && !/id="autoOpen"/.test(html), 'button instead of the checkbox');
  console.log("Finish setup: ok");
}

// ---------------------------------------------------------------- 👋 Welcome tab
{
  const { ss, g } = newSheet();
  g.isAutoMigrateOn_ = () => false;
  g.showWelcomeTabOnly_(ss);
  const vis = () => ss.getSheets().filter(s => !s.isSheetHidden()).map(s => s.getName());
  assert.deepStrictEqual(vis(), ["👋 Welcome"], 'new copy: only the Welcome tab');
  assert.strictEqual(ss.getSheets()[0].getName(), "👋 Welcome", 'first tab');
  g.showWelcomeTabOnly_(ss);
  assert.strictEqual(ss.getSheets().filter(s => s.getName() === "👋 Welcome").length, 1, 'made once');
  // Initial Setup: the user stays on it, now "what to do next"; the other tabs wait for the first build
  g.zg_safeGuideLink_ = () => ""; g.zg_safeExampleLink_ = () => "";
  g.welcomeTabAfterSetup_(ss);
  const w = ss.getSheetByName("👋 Welcome");
  assert.deepStrictEqual(vis(), ["👋 Welcome"], 'after setup: only the Welcome tab');
  assert.strictEqual(ss.getActiveSheet(), w, 'after setup: still on the Welcome tab');
  assert.strictEqual(w.getRange("C3").getValue(), "You're set up", 'rewritten in place as "what to do next"');
  assert(/Input Data Mode/.test(Array.from(w._cells.values()).join("\n")), 'next step named');
  assert.strictEqual(ss.getSheets().filter(s => s.getName() === "👋 Welcome").length, 1, 'same tab, not a second one');
  // Not built yet: kept (an open shows it again); the end of a build deletes it
  assert.strictEqual(g.deleteWelcomeTab_(ss), false, 'kept until the sheets are built');
  ss.getSheetByName("Contact Log").activate(); w.hideSheet();
  g.welcomeTabAfterSetup_(ss, true);
  assert.strictEqual(ss.getActiveSheet(), w, 'an open brings the user back to it');
  assert(ss._toasts.some(t => /Input Data Mode/.test(t.msg)), 'with a toast naming the next step');
  ss.getSheetByName("Contact Log").showSheet(); ss.getSheetByName("Contact Log").activate();
  assert.strictEqual(g.deleteWelcomeTab_(ss, true), true, 'deleted at the end of a build');
  assert.strictEqual(ss.getSheetByName("👋 Welcome"), null);
  console.log("Welcome tab: ok");
}

// ---------------------------------------------------------------- 👋 Welcome tab after Reset to Display Mode
{
  // Reset to Display Mode ends with a fresh Welcome tab as the only tab showing,
  // with every permission explained before the setup steps, and keeps it on the
  // sheet that was reset (setup may have run there) until Initial Setup runs again.
  const ss = makeSpreadsheet(TABS);
  ss.getSheetByName("Backend").hideSheet();
  const app = fakeSpreadsheetApp(ss);
  // rich text that remembers its text (the fake's default builder doesn't)
  app.newRichTextValue = () => { const v = {}; const b = new Proxy({}, { get: (t, k) =>
    k === 'setText' ? (x => { v.text = x; return b; }) : k === 'build' ? (() => v) : (() => b) }); return b; };
  const g = loadProject({ extra: { SpreadsheetApp: app }, props: { SETUP_COMPLETE: 'true' } });
  g.isAutoMigrateOn_ = () => false;
  g.getGuideLink_ = () => "https://example.org/guide";
  const vis = () => ss.getSheets().filter(s => !s.isSheetHidden()).map(s => s.getName());

  g.showWelcomeAfterReset_(ss);
  const w = ss.getSheetByName("👋 Welcome");
  assert.deepStrictEqual(vis(), ["👋 Welcome"], 'after the reset: only the Welcome tab');
  assert.strictEqual(ss.getSheets()[0], w, 'first tab');
  assert.strictEqual(ss.getActiveSheet(), w, 'the tab a copy opens on');
  const text = Array.from(w._cells.values()).join("\n");
  g.__get("WELCOME_PERMISSIONS_").forEach(p => assert(text.indexOf(p[0]) !== -1, 'permission explained: ' + p[0]));
  assert(text.indexOf("WHAT GOOGLE WILL ASK") < text.indexOf("SET UP YOUR COPY"), 'permissions come before the steps');
  assert(/Select all/.test(text) && /Go to … \(unsafe\)/.test(text), 'the real consent screens');

  // Next open of the reset sheet (setup ran there before): the tab stays
  assert.strictEqual(g.deleteWelcomeTab_(ss), false, 'kept after the reset');
  g.buildMenus_();
  assert.deepStrictEqual(vis(), ["👋 Welcome"], 'still the only tab after an open');

  // A second reset: the Welcome tab (the only visible one) goes, a fresh one comes back
  g.getBackendSheet_ = () => ss.getSheetByName("Backend");
  g.getStudentListSheet_ = () => ss.getSheetByName("Student List");
  g.deleteLegacySheets_(ss);
  assert.strictEqual(ss.getSheetByName("👋 Welcome"), null, 'the reset deletes the old Welcome tab');
  assert(vis().indexOf("Contact Log") !== -1, 'without leaving no visible tab');
  assert.strictEqual(g.__props.WELCOME_TAB_KEEP, undefined, 'its keep mark goes with it');
  g.showWelcomeAfterReset_(ss);
  assert.deepStrictEqual(vis(), ["👋 Welcome"]);

  // Initial Setup on that sheet: the keep mark ends, the tab is hidden, then deleted
  g.clearWelcomeKeep_(ss); g.leaveWelcomeTab_(ss); g.applyDefaultSheetVisibility_(ss);
  assert.deepStrictEqual(vis(), ["Contact Log", "Combined Contact Tracking", "PCAR", "Directory"], 'default tabs');
  assert.strictEqual(g.deleteWelcomeTab_(ss, true), true, 'deleted by the build once the keep mark is gone');
  console.log("Welcome tab after Reset to Display Mode: ok");
}
{
  // Initial Setup asks again for any permission left unticked (granular consent)
  const src = fs.readFileSync(path.join(ROOT, 'Menus.gs'), 'utf8');
  const setup = src.slice(src.indexOf('function runInitialSetupWrapper()'));
  assert(/requireAllScopes\(ScriptApp\.AuthMode\.FULL\)/.test(setup.slice(0, 600)), 'requireAllScopes first in Initial Setup');
  // The click answers at once: the menu handler only opens the setup window, which runs the setup
  const wrapper = setup.slice(0, setup.indexOf('\n}\n'));
  assert(/showGuideWelcome_\(\{ setup: true \}\)/.test(wrapper) && !/setupAllInstallableTriggers/.test(wrapper), 'Initial Setup opens the window first');
  // Getting Started: no "coming soon"; empty document links use the User Guide / data guide
  const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(makeSpreadsheet(TABS)) } });
  g.getGuideLink_ = () => "https://example.org/guide";
  // A filled-in link is used as is; an empty DOCUMENT link falls back to the User Guide, an empty VIDEO stays empty
  const L = g.__get("GUIDE_LINKS");
  const links = g.zg_links_();
  ['guideDoc', 'buildDoc', 'settingsDoc', 'contactLogDoc', 'downloadDataDoc'].forEach(k => assert.strictEqual(links[k], L[k] || "https://example.org/guide", k));
  ['setupVideo', 'downloadDataVideo'].forEach(k => assert.strictEqual(links[k], L[k], k));
  const saved = Object.assign({}, L);
  Object.keys(L).forEach(k => { L[k] = ""; });
  const empty = g.zg_links_();
  ['guideDoc', 'buildDoc', 'settingsDoc', 'contactLogDoc', 'downloadDataDoc'].forEach(k => assert.strictEqual(empty[k], "https://example.org/guide", 'empty ' + k + ' -> User Guide'));
  assert.strictEqual(empty.downloadDataVideo, "", 'no video: left out');
  Object.assign(L, saved);
  assert(!/coming soon/.test(src) && !/coming soon/.test(fs.readFileSync(path.join(ROOT, 'GettingStartedSidebar.html'), 'utf8')), 'no "coming soon" left');
  console.log("Initial Setup permissions + Getting Started links: ok");
}

{
  // 🐞 Debug → 📦 Make Release Copies: Blank (B3 = No) is a reset copy, Example the fake-data copy
  // (no Migrate copy any more: one template for everyone); 🔗 My Links lists every link
  const mine = makeSpreadsheet(TABS);
  mine.getName = () => "Mine";
  const app = fakeSpreadsheetApp(mine);
  const g = loadProject({ owner: true, extra: { SpreadsheetApp: app } });
  const made = [];
  g.ownerCopyThisSheet_ = (ss, label, name) => {
    const n = made.length, c = makeSpreadsheet(TABS); c.getId = () => "COPY_" + n; c.getUrl = () => "https://example.org/" + n;
    c.getName = () => name; made.push({ label: name, copy: c }); return c;
  };
  // today's date, like the code (the release names use new Date(); a fixed date broke this test at midnight)
  const now = new Date(), today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  g.stampVersionDate_ = c => c.getSheetByName("Backend").getRange("A2").setValue(today);
  g.Utilities.formatDate = d => (d.getMonth() + 1 + "").padStart(2, "0") + "/" + String(d.getDate()).padStart(2, "0") + "/" + String(d.getFullYear()).slice(2);
  const TODAY = g.Utilities.formatDate(today);
  g.getOrCreateVersionSheet_ = ss => ss.getSheetByName("Backend");
  const steps = [];
  const overrides = [];
  g.resetToDisplayModeSteps_ = c => { steps.push(["reset", c.getSheetByName("Backend").getRange("B3").getValue()]); overrides.push(g.WELCOME_LINKS_OVERRIDE_); };
  g.ownerRunOnCopy_ = (c, fn) => { steps.push(["fake"]); };
  const EX = "https://docs.google.com/spreadsheets/d/EX/edit?usp=sharing";
  const b = g.ownerMakeCopy("blank", EX), e = g.ownerMakeCopy("example");
  // the Blank copy (nobody runs Initial Setup on it) gets the Example made in the same run on its Welcome tab,
  // and its own saved links: its /copy link + that Example
  assert.strictEqual(JSON.stringify(overrides), JSON.stringify([{ exampleUrl: EX }]), 'the reset links the new Example copy');
  assert.strictEqual(g.WELCOME_LINKS_OVERRIDE_, null, 'only for that reset');
  const be0 = made[0].copy.getSheetByName("Backend");
  assert.strictEqual(be0.getRange("E3").getValue(), "https://docs.google.com/spreadsheets/d/COPY_0/copy", 'Blank: its own /copy link saved');
  assert.deepStrictEqual(be0.getRange("D6:E6").getValues()[0].slice(), ["EXAMPLE copy", EX], 'Blank: the Example saved');
  assert.strictEqual(mine.getSheetByName("Backend").getRange("E6").getValue(), "", 'this sheet is not changed');
  // the release names: never this sheet's name ("Mine")
  assert.deepStrictEqual(made.map(x => x.label), ["BLANK", "EXAMPLE"].map(t => "Parent Coordinator Sheet - (" + t + ") (" + TODAY + ")"));
  assert.deepStrictEqual(steps, [["reset", "No"], ["fake"]], 'B3 = No before the reset; Example loads fake data');
  assert.strictEqual(made[0].copy.getSheetByName("Backend").getRange("B3").getValue(), "No");
  assert.throws(() => g.ownerMakeCopy("migrate"), /Unknown copy/, 'no Migrate copy any more');
  // Example: every tab showing except Backend, no Welcome tab, and its open (setup not run) leaves it so
  {
    const ex = made[1].copy;
    assert.strictEqual(ex.getSheetByName(g.WELCOME_TAB_NAME), null, 'Example: no Welcome tab');
    assert(ex.getSheets().every(sh => sh.isSheetHidden() === (sh.getName() === "Backend")), 'Example: all tabs but Backend show');
    assert(g.isExampleCopy_(ex) && !g.isExampleCopy_(made[0].copy), 'only the Example is marked');
    g.showWelcomeTabOnly_(ex);
    assert.strictEqual(ex.getSheetByName(g.WELCOME_TAB_NAME), null, 'Example open: no Welcome tab made');
    assert(ex.getSheets().every(sh => sh.isSheetHidden() === (sh.getName() === "Backend")), 'Example open: nothing hidden');
    g.clearExampleMark_(ex);
    assert(!g.isExampleCopy_(ex), 'a reset removes the mark');
    console.log("Example copy: all tabs but Backend, no Welcome: ok");
  }
  // 1. every copy gets the version date, the Example one too (it has no reset), and B3 = No there
  made.forEach(x => assert(x.copy.getSheetByName("Backend").getRange("A2").getValue() instanceof Date, 'version date in ' + x.label));
  assert.strictEqual(made[1].copy.getSheetByName("Backend").getRange("B3").getValue(), "No", 'Example: not an upgrade template');
  assert.strictEqual(e.stamp, TODAY, 'the pop-up shows the version date');
  assert.strictEqual(JSON.stringify(b), JSON.stringify({ name: "Parent Coordinator Sheet - (BLANK) (" + TODAY + ")", url: "https://example.org/0", stamp: TODAY,
    trackerLink: "https://docs.google.com/spreadsheets/d/COPY_0/copy" }), 'Blank: its /copy link goes in the tracker (D)');
  assert.strictEqual(e.trackerLink, "https://docs.google.com/spreadsheets/d/COPY_1/edit?usp=sharing", 'Example: its link goes in the tracker (E)');
  assert(e.url);
  assert.throws(() => g.ownerMakeCopy("other"), /Unknown copy/);
  let dialog = "";
  app._ui.showModalDialog = (h, title) => { dialog = title + "\n" + h._html; };
  g.include_ = () => "";
  g.HtmlService = { createHtmlOutput: h => { const o = { _html: h }; o.setWidth = () => o; o.setHeight = () => o; return o; } };
  g.ownerMenuMakeAllCopies();
  assert(/ownerMakeCopy/.test(dialog) && /Parent Coordinator Sheet - \(BLANK\)/.test(dialog) && /\(EXAMPLE\)/.test(dialog), 'the release copies pop-up');
  assert(!/ORDER = \[[^\]]*migrate/.test(dialog), 'no Migrate copy in the pop-up');
  assert(/Version date in both/.test(dialog), 'the pop-up says which date goes in the Version Tracker');
  assert(/id="l_blank"/.test(dialog) && /id="l_example"/.test(dialog), 'the pop-up shows the D / E links');
  assert(/ownerMakeCopy\(k, k === 'blank' \? exampleLink : ''\)/.test(dialog), 'the pop-up passes the Example link to the Blank copy');
  g.getGuideLink_ = () => "https://example.org/guide.pdf";
  g.trackerLinks_ = () => ({ blank: "https://docs.google.com/spreadsheets/d/B/copy", example: "", guide: "", fromTracker: true,
    cols: { blank: 'D "Blank Copy"', example: "", guide: 'F "Document"' } });
  g.ownerMenuMyLinks();
  assert(/column D &quot;Blank Copy&quot;/.test(dialog) && /d\/B\/copy/.test(dialog), 'My Links: the tracker\'s Blank Copy column and link');
  assert(/no column named &quot;Example&quot;/.test(dialog), 'My Links: says when the Example column is missing');
  assert(/the guide folder&#39;s newest file|the guide folder's newest file/.test(dialog), 'My Links: an empty Document link says the folder is used');
  Object.keys(g.__get("GUIDE_LINKS")).forEach(k => assert(dialog.indexOf(k) !== -1, 'My Links lists ' + k));
  assert(/YOUR LINKS/.test(dialog) && /guide\.pdf/.test(dialog), 'My Links: where to change them + the newest guide file');
  // The guide folder: a date in the NAME ranks first, else the last change; My Links rescans and lists them
  const gfile = (name, changed) => ({ getName: () => name, getUrl: () => "https://drive.example/" + encodeURIComponent(name),
    getLastUpdated: () => changed, getId: () => name });
  const gfiles = [gfile("PC_Tracker_User_Guide.pdf", new Date(2026, 10, 12, 14, 0)),     // uploaded today, no date in the name
                  gfile("PC Tracker User Guide 11/20/26", new Date(2026, 10, 1)),        // a LATER date in its name
                  gfile("Old guide.pdf", new Date(2026, 8, 1))];
  g.DriveApp = { getFolderById: () => ({ getFiles: () => { let i = 0; return { hasNext: () => i < gfiles.length, next: () => gfiles[i++] }; } }) };
  const ranked = g.guideFolderFiles_();
  assert.strictEqual(ranked[0].name, "PC Tracker User Guide 11/20/26", 'a date in the name beats a newer upload');
  assert.strictEqual(ranked[1].name, "PC_Tracker_User_Guide.pdf");
  assert.strictEqual(g.scanGuideFolder_(), ranked[0].url);
  let rescanned = false;
  g.refreshResolvedLinks_ = () => { rescanned = true; return {}; };
  g.ownerMenuMyLinks();
  assert(rescanned, 'My Links looks again instead of using the 6-hour copy');
  assert(/PC Tracker User Guide 11\/20\/26/.test(dialog) && /date in the name/.test(dialog), 'My Links lists the guide folder in its order');
  // The YOUR LINKS box is the first thing in Menus.gs
  const src = fs.readFileSync(path.join(ROOT, 'Menus.gs'), 'utf8');
  assert(src.indexOf('YOUR LINKS') < src.indexOf('function '), 'links box at the top of Menus.gs');
  console.log("Debug → Make Release Copies + My Links: ok");
}

{
  // Backend TEMPLATE LINKS: saving the guide link (D5:E5) clears an old second "GUIDE link" row (D6:E6)
  const ss = makeSpreadsheet(TABS);
  const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) } });
  const be = ss.getSheetByName("Backend");
  g.getOrCreateVersionSheet_ = () => be;
  be.getRange("D6:E6").setValues([["GUIDE link", "https://old.example.org"]]);
  be.getRange("A6:B6").setValues([["syncToPhones", "false"]]);
  g.saveGuideLinkToVersionSheet_("https://example.org/guide");
  assert.deepStrictEqual(be.getRange("D5:E6").getValues().map(r => r.slice()), [["GUIDE link", "https://example.org/guide"], ["", ""]]);
  assert.strictEqual(be.getRange("A6").getValue(), "syncToPhones", 'the settings block next to it is untouched');
  be.getRange("D6:E6").setValues([["Something else", "x"]]);
  g.saveGuideLinkToVersionSheet_("https://example.org/guide");
  assert.strictEqual(be.getRange("D6").getValue(), "Something else", 'only a stale "GUIDE link" row is cleared');
  console.log("Backend: stale second GUIDE link row cleared: ok");
}

{
  // The Version Tracker's link columns (owner 2026-11-12): D Blank Copy, E Example, F Document, found by
  // header name; the newest dated row WITH a link counts; the folders are the backup
  const g0 = loadProject({});
  const C = h => JSON.stringify(g0.trackerColumns_(h));
  assert.strictEqual(C(["Version", "Size", "Notes", "Blank Copy", "Example", "Document"]),
    JSON.stringify({ date: 0, size: 1, notes: 2, blank: 3, example: 4, guide: 5 }));
  // today's tracker: "Link" is the blank copy; the old "Blank Link" (E) is NOT read as the example
  assert.strictEqual(C(["Version", "Size", "Notes", "Link", "Blank Link", "Column 1"]),
    JSON.stringify({ date: 0, size: 1, notes: 2, blank: 3, example: -1, guide: 5 }));
  // moved columns are found by name
  assert.strictEqual(C(["Version", "Document", "Size", "Notes", "Example", "Blank Copy"]),
    JSON.stringify({ date: 0, size: 2, notes: 3, blank: 5, example: 4, guide: 1 }));
  const row = (d, blank, example, guide) => ({ dateMs: d ? new Date(d).getTime() : NaN, link: blank, example: example, guide: guide });
  const L = g0.trackerLinksFromRows_([
    row("2026-09-01", "https://b/old/copy", "https://e/old", "https://g/old"),
    row("2026-11-12", "https://b/new/copy", "", "not a link"),
    row("", "https://b/undated/copy", "https://e/undated", "https://g/undated"),
    row("2026-10-01", "", "https://e/oct", "")]);
  assert.strictEqual(JSON.stringify(L), JSON.stringify({ blank: "https://b/new/copy", example: "https://e/oct", guide: "https://g/old", fromTracker: true }),
    'each column: the newest dated row that has a link');

  // getTemplateLinks_ / getGuideLink_: tracker first, then the folders, then the Backend's saved copy
  const ss = makeSpreadsheet(TABS);
  const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) } });
  const be = ss.getSheetByName("Backend");
  g.getOrCreateVersionSheet_ = () => be;
  let scans = 0;
  g.scanTemplateFolder_ = () => { scans++; return { blankCopyUrl: "https://folder/blank", migrateCopyUrl: "https://folder/blank" }; };
  g.scanGuideFolder_ = () => "https://folder/guide";
  const fresh = () => { const c = g.CacheService.getScriptCache(); ["TEMPLATE_LINKS_V1", "GUIDE_LINK_V1", "TRACKER_LINKS_V1"].forEach(k => c.remove(k)); };
  g.trackerLinks_ = () => ({ blank: "https://tracker/copy", example: "https://tracker/example", guide: "https://tracker/guide", fromTracker: true });
  fresh();
  let t = g.getTemplateLinks_();
  assert.strictEqual(t.blankCopyUrl, "https://tracker/copy"); assert.strictEqual(t.migrateCopyUrl, "https://tracker/copy");
  assert.strictEqual(t.exampleUrl, "https://tracker/example"); assert.strictEqual(scans, 0, 'no folder scan when the tracker has the link');
  assert.strictEqual(g.getGuideLink_(), "https://tracker/guide");
  assert.strictEqual(be.getRange("E3").getValue(), "https://tracker/copy", 'saved in the Backend for copies');
  assert.deepStrictEqual(be.getRange("D6:E6").getValues()[0].slice(), ["EXAMPLE copy", "https://tracker/example"]);
  assert.strictEqual(be.getRange("E5").getValue(), "https://tracker/guide");
  // the tracker has no links (or no Example): the folders; the Example button goes, the saved one is cleared
  g.trackerLinks_ = () => ({ blank: "", example: "", guide: "", fromTracker: true });
  fresh();
  t = g.getTemplateLinks_();
  assert.strictEqual(t.blankCopyUrl, "https://folder/blank"); assert.strictEqual(t.exampleUrl, "", 'no Example in the tracker = no button');
  assert.strictEqual(be.getRange("D6").getValue(), "", 'the saved Example is cleared');
  assert.strictEqual(g.getGuideLink_(), "https://folder/guide");
  // nothing can be read (a copy without access): the Backend's saved copy
  be.getRange("D6:E6").setValues([["EXAMPLE copy", "https://saved/example"]]);
  g.trackerLinks_ = () => ({ blank: "", example: "", guide: "", fromTracker: false });
  g.scanTemplateFolder_ = () => ({ blankCopyUrl: "", migrateCopyUrl: "" });
  g.scanGuideFolder_ = () => "";
  fresh();
  t = g.getTemplateLinks_();
  assert.strictEqual(t.blankCopyUrl, "https://folder/blank", 'saved blank link');
  assert.strictEqual(t.exampleUrl, "https://saved/example", 'saved example link');
  assert.strictEqual(g.getGuideLink_(), "https://folder/guide", 'saved guide link');
  // the help links show the Example button only when there is one
  g.getTemplateLinks_ = () => ({ exampleUrl: "https://tracker/example" });
  assert.strictEqual(g.zg_links_().example, "https://tracker/example");
  g.getTemplateLinks_ = () => ({ exampleUrl: "" });
  assert.strictEqual(g.zg_links_().example, "");
  // the Welcome tab: an Example row only with a link
  const welcomeText = exampleUrl => {
    const ss2 = makeSpreadsheet(TABS);
    const app2 = fakeSpreadsheetApp(ss2);
    app2.newRichTextValue = () => { const v = {}; const b = new Proxy({}, { get: (t, k) =>
      k === 'setText' ? (x => { v.text = x; return b; }) : k === 'build' ? (() => v) : (() => b) }); return b; };
    const g2 = loadProject({ extra: { SpreadsheetApp: app2 } });
    g2.isAutoMigrateOn_ = () => false;
    return Array.from(g2.buildWelcomeTab_(ss2, { guideUrl: "https://g", exampleUrl: exampleUrl })._cells.values()).join("\n");
  };
  assert(/See an example tracker/.test(welcomeText("https://e")), 'Welcome tab: an example row with a link');
  {
    // Reset to Display Mode with WELCOME_LINKS_OVERRIDE_ (Make Release Copies): that Example, not the tracker's
    const ss3 = makeSpreadsheet(TABS);
    const app3 = fakeSpreadsheetApp(ss3);
    app3.newRichTextValue = () => { const v = {}; const b = new Proxy({}, { get: (t, k) =>
      k === 'setText' ? (x => { v.text = x; return b; }) : k === 'setLinkUrl' ? (u => { v.links = (v.links || []).concat(u); return b; }) : k === 'build' ? (() => v) : (() => b) }); return b; };
    const seen = [];
    const ws = ss3.getSheets()[0];
    const g3 = loadProject({ extra: { SpreadsheetApp: app3 } });
    g3.isAutoMigrateOn_ = () => false;
    g3.refreshResolvedLinks_ = () => ({ guide: "https://g", links: { exampleUrl: "https://tracker/old-example" } });
    g3.buildWelcomeTab_ = (sh, o) => { seen.push(o); return ws; };
    g3.showWelcomeTabOnly_ = () => {};
    g3.WELCOME_LINKS_OVERRIDE_ = { exampleUrl: "https://new/example" };
    g3.showWelcomeAfterReset_(ss3);
    assert.strictEqual(seen[0].exampleUrl, "https://new/example", 'the override wins');
    g3.WELCOME_LINKS_OVERRIDE_ = null;
    seen.length = 0; ss3.getSheetByName = () => null;
    g3.showWelcomeAfterReset_(ss3);
    assert.strictEqual(seen[0].exampleUrl, "https://tracker/old-example", 'no override: the tracker\'s');
  }
  assert(!/See an example tracker/.test(welcomeText("")), 'Welcome tab: no example row without one');
  console.log("Version Tracker links (Blank Copy / Example / Document, folders as backup): ok");
}

{
  // 2. New/Edit Student (Backend G:AV): title row 1, header row 2, students from row 3. A section one row
  //    lower (the owner's sheet) is moved up, so a reset (clears rows 3+) keeps its header; an empty header
  //    row gets the headers back; other sections' columns are never touched.
  const ss = makeSpreadsheet(TABS);
  const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) } });
  const be = ss.getSheetByName("Backend");
  g.getBackendSheet_ = () => be;
  const H = g.__get("OVERRIDES_HEADERS_");
  const row = (osis, name) => { const r = Array(42).fill(""); r[0] = "11/01/26"; r[1] = "Edit"; r[2] = osis; r[3] = name; return r; };
  be.getRange("G1").setValue("NEW / EDIT STUDENT");
  be.getRange("G3:AV3").setValues([H]);                       // header one row too low
  be.getRange("G4:AV5").setValues([row("234567890", "Belcher"), row("234567891", "Addams")]);
  be.getRange("A2:A6").setValues([["11/12/26"], ["Auto-Migrate"], ["x"], ["SAVED SETTINGS"], ["syncToPhones"]]); // Version section
  be.getRange("AX4").setValue("event row");                    // Event Log section
  assert.strictEqual(g.fixOverridesHeaderRow_(ss, false), 1, 'moved up one row');
  assert.strictEqual(be.getRange("I2").getValue(), "OSIS", 'header on row 2');
  assert.deepStrictEqual([be.getRange("I3").getValue(), be.getRange("I4").getValue(), be.getRange("I5").getValue()], ["234567890", "234567891", ""], 'students from row 3');
  assert.strictEqual(be.getRange("A6").getValue(), "syncToPhones", 'Version section untouched');
  assert.strictEqual(be.getRange("AX4").getValue(), "event row", 'Event Log untouched');
  assert.strictEqual(g.fixOverridesHeaderRow_(ss, false), 0, 'already right: nothing moves');
  // what a reset does: clear rows 3+, then an empty header row gets the headers
  be.getRange("G3:AV6").clearContent();
  be.getRange("G2:AV2").clearContent();
  g.fixOverridesHeaderRow_(ss, true);
  assert.strictEqual(be.getRange("G2").getValue(), "Date"); assert.strictEqual(be.getRange("AV2").getValue(), "Hide From Report");
  // a student above a header-looking row is never moved over
  be.getRange("G2:AV2").setValues([row("234567892", "Parker")]); be.getRange("G3:AV3").setValues([H]);
  assert.strictEqual(g.fixOverridesHeaderRow_(ss, false), 0, 'a student on row 2 stays');
  console.log("New/Edit Student header row: ok");
}

{
  // One template for everyone: the first window after Initial Setup asks "Is this your first PC Tracker?"
  const ss = makeSpreadsheet(TABS);
  const app = fakeSpreadsheetApp(ss);
  let dialog = "";
  app._ui.showModalDialog = (h, title) => { dialog = h._html; };
  const g = loadProject({ extra: { SpreadsheetApp: app } });
  g.include_ = () => ""; g.openGuideChecklist = () => {}; g.getGuideLink_ = () => "https://example.org/guide";
  g.HtmlService = { createHtmlOutput: h => { const o = { _html: h }; o.setWidth = () => o; o.setHeight = () => o; return o; } };
  g.showGuideWelcome_();
  assert(/Is this your first PC Tracker\?/.test(dialog), 'the question');
  assert(/Yes, start fresh/.test(dialog) && /zg_action\('inputData'\)/.test(dialog), 'start fresh -> Input Data Mode');
  assert(/No, bring over last year's tracker/.test(dialog) && /zg_action\('importOld'\)/.test(dialog), 'bring over -> the Migration Wizard');

  // The wizard's list: trackers in the user's Drive, newest first, this sheet left out, at most 8
  const files = [["THIS_SHEET", "Parent Coordinator Sheet - (BLANK) (11/12/26)", 9],
    ["OLD1", "Copy of Parent Coordinator Tracker 7/22/25 (BLANK) – Retired on 06/20/26", 5],
    ["OLD2", "PC Tracker 2024", 2]].map(([id, name, day]) => ({
      getId: () => id, getName: () => name, getUrl: () => "https://docs.google.com/spreadsheets/d/" + id + "/edit",
      getLastUpdated: () => new Date(2026, 5, day) }));
  const queries = [];
  g.SpreadsheetApp.getActiveSpreadsheet = () => ({ getId: () => "THIS_SHEET" });
  g.DriveApp = { searchFiles: q => { queries.push(q); if (/'me' in owners/.test(q)) throw new Error("not allowed here");
    let i = 0; return { hasNext: () => i < files.length, next: () => files[i++] }; } };
  g.Utilities.formatDate = d => d.toISOString().slice(0, 10);
  const found = g.ma_findOldTrackers();
  assert.strictEqual(JSON.stringify(found.map(f => f.url.split("/d/")[1].split("/")[0])), JSON.stringify(["OLD1", "OLD2"]), 'newest first, not this sheet');
  assert.strictEqual(queries.length, 2, 'falls back when the owner filter is refused');
  assert(/title contains 'Retired on'/.test(queries[1]) && /trashed = false/.test(queries[1]));
  g.DriveApp = { searchFiles: () => { throw new Error("no Drive"); } };
  assert.strictEqual(g.ma_findOldTrackers().length, 0, 'never throws: the wizard shows only the paste box');

  // The template folder: the upgrade link is the (BLANK) file; an old (MIGRATE) file is not used
  const tfile = (id, name) => ({ getId: () => id, getName: () => name, getMimeType: () => "application/vnd.google-apps.spreadsheet",
    getLastUpdated: () => new Date(2026, 10, 1) });
  const tfiles = [tfile("NEW_BLANK", "Parent Coordinator Sheet - (BLANK) (11/12/26)"), tfile("OLD_MIG", "Parent Coordinator Tracker 7/22/26 (MIGRATE)")];
  g.DriveApp = { getFolderById: () => ({ getFiles: () => { let i = 0; return { hasNext: () => i < tfiles.length, next: () => tfiles[i++] }; } }) };
  const t = g.scanTemplateFolder_();
  assert(/NEW_BLANK\/copy$/.test(t.blankCopyUrl) && t.migrateCopyUrl === t.blankCopyUrl, 'one template link: ' + JSON.stringify(t));
  console.log("One template: the question, the Drive list, the template link: ok");
}

// ---------------------------------------------------------------- new-user flow pieces
{
  // A new copy opens: the Welcome tab only, a toast pointing at Initial Setup, no pop-up
  const ss = makeSpreadsheet(TABS);
  const app = fakeSpreadsheetApp(ss);
  const g = loadProject({ extra: { SpreadsheetApp: app } });
  g.buildMenus_();
  assert.deepStrictEqual(ss.getSheets().filter(s => !s.isSheetHidden()).map(s => s.getName()), ["👋 Welcome"], 'new copy: Welcome tab only');
  assert.strictEqual(app._ui._alerts.length, 0, 'new copy: no pop-up on top of the Welcome tab');
  assert(ss._toasts.some(t => /Initial Setup/.test(t.msg)), 'new copy: a toast points at Initial Setup');
  assert(app._ui._menus.some(m => m.items.indexOf("runInitialSetupWrapper") !== -1), 'new copy: the Initial Setup menu');
  console.log("new copy open: Welcome tab + toast, no pop-up: ok");
}
{
  // End of Input Data Mode: no second build when step 5 already built and nothing changed since
  const { g } = newSheet();
  const st = { results: { ats: { rows: 18 } }, listsBuild: { ok: true }, dataChanged: false };
  assert.strictEqual(g.idm_finalBuildNeeded_(st), false, 'built at step 5: no build button');
  assert.strictEqual(g.idm_builtAlready_(st), true);
  assert.strictEqual(g.idm_finalBuildNeeded_(Object.assign({}, st, { listsChanged: true })), true, 'a list saved after: build');
  assert.strictEqual(g.idm_finalBuildNeeded_(Object.assign({}, st, { dataChanged: true })), true, 'new data after: build');
  assert.strictEqual(g.idm_finalBuildNeeded_({ results: { ats: { rows: 18 } } }), true, 'no step 5 build (left early): build');
  assert.strictEqual(g.idm_finalBuildNeeded_({ results: {} }), false, 'nothing came in: no build');
  // A list saved during the mode is noted
  g.idm_saveState_({ step: 4, results: {}, listsBuild: { ok: true } });
  g.idm_noteListsChanged_();
  assert.strictEqual(g.idm_loadState_().listsChanged, true, 'list saved -> listsChanged');
  console.log("end screen build button only when needed: ok");
}
{
  // The full menus right after the first build (not during Input Data Mode, not twice)
  const { g } = newSheet();
  let opens = 0;
  g.onOpen = () => { opens++; };
  g.studentListIsBuilt_ = () => true;
  assert.strictEqual(g.refreshMenusAfterFirstBuild_(), false, 'before setup: no');
  g.__props.SETUP_COMPLETE = 'true';
  g.__props.INPUT_DATA_MODE_STATE = '{}';
  assert.strictEqual(g.refreshMenusAfterFirstBuild_(), false, 'during Input Data Mode: no');
  delete g.__props.INPUT_DATA_MODE_STATE;
  assert.strictEqual(g.refreshMenusAfterFirstBuild_(), true, 'first build: menus redrawn');
  g.__props.FULL_MENU_UNLOCKED = 'true';
  assert.strictEqual(g.refreshMenusAfterFirstBuild_(), false, 'already unlocked: no');
  assert.strictEqual(opens, 1);
  console.log("full menus after the first build: ok");
}
{
  // Closing Input Data Mode: the other tabs come back, the Your Lists tabs stay hidden
  const { ss, g, sheet, lists } = listsSheet();
  const old = ss.insertSheet("📥 PTA"); // a list from before the mode, showing when the mode starts
  lists().push({ tabId: old.getSheetId(), name: "PTA", setup: g.od_emptySetup_(), updated: "", sheet: old });
  sheet.hideSheet();                     // "Bus Routes" stands for a list made DURING the mode:
  const temp = startMode(ss, g, 4);      //   not there when the mode hid the tabs...
  sheet.showSheet(); sheet.activate();   //   ...then made by od_createList (shown and active)
  const v = g.idm_finishLists();
  assert(v.finished, 'mode finished');
  const vis = ss.getSheets().filter(s => !s.isSheetHidden()).map(s => s.getName());
  assert(vis.indexOf("📥 Bus Routes") === -1 && vis.indexOf("📥 PTA") === -1, 'list tabs hidden: ' + vis.join(", "));
  ["Contact Log", "Combined Contact Tracking", "PCAR", "Directory"].forEach(n => assert(vis.indexOf(n) !== -1, n + ' back'));
  assert.notStrictEqual(ss.getActiveSheet().getSheetId(), sheet.getSheetId(), 'not left on a hidden list tab');
  assert.notStrictEqual(ss.getActiveSheet().getSheetId(), temp.getSheetId(), 'off the temp tab (deleted next call)');
  console.log("closing Input Data Mode hides the Your Lists tabs: ok");
}

// ---------------------------------------------------------------- Your Lists: View step (how the values read)
{
  const { g } = newSheet();
  // Lists saved before the View step read exactly as before; a new list starts with bold names
  const old = g.od_fillSetup_({ ready: true, by: "student", roles: {}, vars: [] });
  assert.strictEqual(old.view.layout, "lines");
  assert.strictEqual(old.view.boldNames, false, 'saved list: no bold');
  assert.strictEqual(g.od_emptySetup_().view.boldNames, true, 'new list: bold names');
  const items = [{ label: "Bus Route", value: "R12", show: true, column: true },
                 { label: "Stop", value: "", show: true, column: true },
                 { label: "AM / PM", value: "AM", show: true, column: false }];
  assert.deepStrictEqual(JSON.parse(JSON.stringify(g.od_viewText_(old.view, "Bus Routes", items))),
    { lines: ["Bus Route: R12", "AM / PM: AM"], column: "R12" }, 'default view = "Name: value" lines');
  const one = g.od_cleanView_({ layout: "oneLine", boldNames: true, sep: "dot", columns: "one" });
  const t = g.od_viewText_(one, "Bus Routes", items);
  assert.deepStrictEqual(Array.from(t.lines), [g.boldText_("Bus Routes") + ": R12 · AM"], 'one line, bold name, dot');
  assert.strictEqual(g.od_viewText_(Object.assign({}, one, { namesBefore: true, start: "custom", startText: "Bus" }), "Bus Routes", items).lines[0],
    g.boldText_("Bus") + ": " + g.boldText_("Bus Route") + " R12 · " + g.boldText_("AM / PM") + " AM", 'names before values, own words');
  assert.strictEqual(g.od_viewText_(Object.assign({}, one, { start: "none", boldValues: true, italic: true }), "x", items).lines[0],
    g.boldText_("R12", true) + " · " + g.boldText_("AM", true), 'values bold italic, no start');
  assert.deepStrictEqual(Array.from(g.od_viewText_(one, "x", items.map(i => Object.assign({}, i, { value: "" }))).lines), [], 'nothing to show: no line');
  // Junk from the sidebar is cleaned
  const junk = g.od_cleanView_({ layout: "zigzag", sep: "<b>", start: 5, startText: "x".repeat(99), columns: "all", order: "no" });
  assert.deepStrictEqual([junk.layout, junk.sep, junk.start, junk.startText.length, junk.columns, junk.order.length], ["lines", "comma", "list", 60, "each", 0]);
  // Order by header (a new version of the list with moved columns keeps it)
  const setup = { vars: [{ col: 3, header: "Route", show: true }, { col: 4, header: "Stop", column: true }, { col: 5, header: "Driver" }],
                  view: { order: ["Stop", "Route"] } };
  assert.deepStrictEqual(g.od_viewVars_(setup).map(v => v.header), ["Stop", "Route"], 'order by header, unkept left out');

  // The sidebar's preview copy (YourListsSidebar.html, VIEW TEXT) gives the same text as the server
  const html = fs.readFileSync(path.join(ROOT, 'YourListsSidebar.html'), 'utf8');
  const block = html.slice(html.indexOf('// VIEW TEXT'), html.indexOf('// END VIEW TEXT'));
  assert(block.length > 500, 'VIEW TEXT block found in the sidebar');
  const side = require('vm').runInNewContext(block + '; ({ viewText: viewText, viewVars: viewVars, boldText: boldText })', {});
  const views = [];
  ["lines", "oneLine"].forEach(layout => [false, true].forEach(boldNames => [false, true].forEach(boldValues =>
    ["comma", "dot", "bar", "slash"].forEach(sep => ["list", "none", "custom"].forEach(start => [false, true].forEach(namesBefore =>
      views.push({ layout, boldNames, boldValues, italic: boldValues && sep === "bar", sep, start, startText: " Bus 9 ", namesBefore, columns: "each", order: [] })))))));
  views.forEach(vw => {
    const a = JSON.stringify(g.od_viewText_(vw, " Bus Routes ", items)), b = JSON.stringify(side.viewText(vw, " Bus Routes ", items));
    assert.strictEqual(b, a, 'sidebar preview = server text for ' + JSON.stringify(vw));
  });
  assert.deepStrictEqual(side.viewVars(setup).map(v => v.header), ["Stop", "Route"], 'sidebar order = server order');
  assert.strictEqual(side.boldText("Ab9 é"), g.boldText_("Ab9 é"));
  console.log("Your Lists View: default = old text, layouts, cleaning, sidebar copy matches (" + views.length + " views): ok");
}
{
  // The build with a View: one line per list on the phone, one column for the list
  const { g } = newSheet();
  const ats = csv("1 - Student Data (ATS).csv").slice(1);
  const raw = [[], [], [], [], [], []].concat(ats);
  const sm = g.buildStudentMap_(raw, new Map());
  const master = [["OSIS", "Student", "Guardian", "Site", "Class"]].concat([...sm.entries()].map(([id, s]) => [id, s.student, "", s.site || "", s.classCode || ""]));
  const guardians = g.buildGuardianObjects_(raw, new Map());
  const idx = g.od_index_(master.slice(1), guardians);
  const block = csv("5a - Your Lists - Bus Routes (student list).csv").filter(r => r.some(x => String(x).trim() !== ""));
  const h = g.od_findHeaderRow_(block), headers = block[h], rows = block.slice(h + 1);
  const guess = g.od_guessRoles_(headers, rows, idx);
  const vars = g.od_varsFor_(headers, guess.roles, []).map(v => Object.assign(v, { show: true, column: true }));
  const make = view => ({ tabId: 1, name: "Bus Routes", sheet: { getSheetId: () => 1 },
    setup: g.od_fillSetup_({ ready: true, by: guess.by, roles: guess.roles, vars: JSON.parse(JSON.stringify(vars)), picks: {}, headers: headers, view: view }) });
  let addl = null, list = null;
  g.od_loadLists_ = () => [list];
  g.od_readTable_ = () => ({ headers: headers, rows: rows });
  g.od_writeToolColumns_ = () => {}; g.od_saveLists_ = () => {}; g.backendClearBlock_ = () => {};
  g.backendWriteBlock_ = (ss, key, r) => { if (key === "ADDL_INFO") addl = r; };
  list = make(undefined);
  let out = g.od_buildOutputs_({}, master, guardians);
  const bart = () => addl.find(r => r[1] === "Bart Simpson");
  assert.strictEqual(bart()[3].split("\n").length, vars.length, 'no view: one plain line per value');
  assert(!/[\u{1D400}-\u{1D7FF}]/u.test(bart()[3]), 'no view: no bold');
  assert.strictEqual(out.studentCols.length, vars.length, 'no view: a column per value');
  list = make({ layout: "oneLine", boldNames: true, columns: "one" });
  out = g.od_buildOutputs_({}, master, guardians);
  assert.strictEqual(bart()[3].split("\n").length, 1, 'one line: a single line');
  assert(bart()[3].indexOf(g.boldText_("Bus Routes") + ": ") === 0, 'one line starts with the bold list name');
  assert.strictEqual(out.studentCols.length, 1, 'one column for the list');
  assert.strictEqual(out.studentCols[0].header, "Bus Routes");
  assert(out.studentCols[0].byOsis.get("123456789").indexOf(", ") !== -1, 'the column holds the values with commas');
  console.log("build with a View: one line + one column: ok");
}
{
  // "Save, exit and build": only offered when there is Student Data (or on its step)
  const { ss, g } = newSheet();
  startMode(ss, g, 1); // Parent Square, no Student Data yet
  assert.strictEqual(g.idm_getState().canBuild, false, 'no Student Data: no build');
  g.idm_goToStep(0);
  assert.strictEqual(g.idm_getState().canBuild, true, 'on the Student Data step: allowed (a paste may be waiting)');
  ss.getSheetByName("Raw Data")._put([["School Year", "DBN", "Student ID"], ["2025-2026", "75M094", 123456789]], 6, 1);
  g.idm_goToStep(2);
  assert.strictEqual(g.idm_getState().canBuild, true, 'Student Data there: allowed');
  const v = g.idm_saveAndExit();
  assert(v.finished && v.exited && v.canBuild === true, 'the exit result says a build can run');
  console.log("Save, exit and build availability: ok");
}

// ---------------------------------------------------------------- faster Initial Setup (owner 2026-11-13)
// A new copy skips Sheet Lock at setup (its first build applies every warning; until then only the
// Welcome tab shows); a sheet that is already built still gets it. The settings go in with ONE
// property write, and getSettings reads every property in one call; the values are as before.
{
  const setup = (built) => {
    const ss = makeSpreadsheet(TABS);
    ss.getSheetByName("Backend").hideSheet();
    const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) }, noTrialFiles: true });
    g.isAutoMigrateOn_ = () => false;
    g.zg_safeGuideLink_ = () => ""; g.zg_safeExampleLink_ = () => "";
    g.onOpen({ source: ss });
    if (built) g.studentListIsBuilt_ = () => true;
    let locks = 0;
    g.applySheetLocks_ = () => { locks++; };
    const dp = g.PropertiesService.getDocumentProperties();
    const one = dp.setProperty, many = dp.setProperties;
    const singles = [], batches = [];
    dp.setProperty = (k, v) => { singles.push(k); return one(k, v); };
    dp.setProperties = (o, d) => { batches.push(Object.keys(o)); return many(o, d); };
    g.runInitialSetup_();
    return { g, locks, singles, batches };
  };

  const fresh = setup(false);
  assert.strictEqual(fresh.locks, 0, 'new copy: no Sheet Lock at setup (the first build does it)');
  assert.strictEqual(setup(true).locks, 1, 'a built sheet: Sheet Lock at setup as before');

  const p = fresh.g.__props;
  assert.strictEqual(p.SETUP_COMPLETE, 'true');
  assert.strictEqual(p.ORIGINAL_SHEET_ID, 'FAKE_SS');
  const expected = { syncToPhones: 'false', phoneContacts: 'false', parentsDivided: 'false', notesTab: 'false', sendOut: 'false',
    siteClass: 'false', typeCol: 'false', autoNotes: 'true', disableHiddenNotes: 'false', dataReminder: 'week' };
  Object.keys(expected).forEach(k => assert.strictEqual(p[k], expected[k], 'setting ' + k));
  const settingKeys = Object.keys(expected);
  assert(fresh.batches.some(b => settingKeys.every(k => b.indexOf(k) !== -1)), 'the settings go in with one write');
  assert.deepStrictEqual(fresh.singles.filter(k => settingKeys.indexOf(k) !== -1 || k === 'SETUP_COMPLETE'), [], 'not one by one');

  // getSettings: one read, same answers (missing = default, saved values win)
  const g = fresh.g, dp = g.PropertiesService.getDocumentProperties();
  dp.setProperty('typeCol', 'true'); dp.setProperty('showRawData', 'true'); dp.deleteProperty('autoNotes');
  let reads = 0, readAll = 0;
  const get = dp.getProperty, getAll = dp.getProperties;
  dp.getProperty = k => { reads++; return get(k); };
  dp.getProperties = () => { readAll++; return getAll(); };
  const s = g.getSettings();
  assert.strictEqual(readAll, 1, 'getSettings: every property in one call');
  assert(reads <= 1, 'getSettings: no one-by-one reads (only the data reminder helper)');
  assert.strictEqual(s.typeCol, true, 'a saved value');
  assert.strictEqual(s.autoNotes, true, 'a missing value = its default');
  assert.strictEqual(s.syncToPhones, false);
  assert.strictEqual(s.showRawData, true, 'a tab shown in Settings');
  assert.strictEqual(s.showContactLog, true, 'a tab never saved = its default');
  assert.strictEqual(s.dataReminder, 'week');
  console.log("faster Initial Setup (no Sheet Lock on a new copy, one settings write / read): ok");
}

// ---------------------------------------------------------------- the 🚀 App Menu first when a sheet opens (owner 2026-11-13)
// A new copy's first open took long (or never) to show the menu: the menu is now the first thing
// (after the copy check), the Contact Log filter isn't re-made before setup, Backend B3 is read once,
// and the copy check only deletes the properties the copy has. A set-up sheet: every menu first,
// then the filter (still re-made on every open).
{
  const ss = makeSpreadsheet(TABS);
  ss.getSheetByName("Backend").hideSheet();
  const app = fakeSpreadsheetApp(ss);
  const order = [];
  const createMenu = app._ui.createMenu;
  app._ui.createMenu = name => { order.push('menu:' + name); return createMenu(name); };
  const toast = ss.toast;
  ss.toast = (m, t) => { order.push('toast'); return toast(m, t); };
  const g = loadProject({ extra: { SpreadsheetApp: app }, noTrialFiles: true });
  let migrateReads = 0;
  g.isAutoMigrateOn_ = () => { migrateReads++; return false; };
  g.applyQBlackTextFilter_ = () => { order.push('filter'); };
  const welcome = g.showWelcomeTabOnly_;
  g.showWelcomeTabOnly_ = s => { order.push('welcome'); return welcome(s); };
  const dp = g.PropertiesService.getDocumentProperties();
  // a copy of a sheet that was set up: the copy check starts it over
  dp.setProperty('ORIGINAL_SHEET_ID', 'THE_SHEET_IT_WAS_COPIED_FROM');
  dp.setProperty('SETUP_COMPLETE', 'true');
  dp.setProperty('ZG_AUTO_OPEN', 'off');
  const deleted = [];
  const del = dp.deleteProperty;
  dp.deleteProperty = k => { deleted.push(k); return del(k); };

  g.onOpen({ source: ss });
  assert.strictEqual(order[0], 'menu:🚀 App Menu', 'new copy: the menu is the first thing');
  assert.deepStrictEqual(app._ui._menus[0].items, ['runInitialSetupWrapper'], 'with Initial Setup');
  assert.strictEqual(order.indexOf('filter'), -1, 'no Contact Log filter before setup');
  assert(order.indexOf('welcome') > 0 && order.lastIndexOf('toast') > order.indexOf('welcome'), 'then the Welcome tab, then the pointer');
  assert.deepStrictEqual(deleted.slice().sort(), ['SETUP_COMPLETE', 'ZG_AUTO_OPEN'], 'only the properties the copy has are deleted');
  assert.strictEqual(g.__props.SETUP_COMPLETE, undefined, 'the copy starts over');
  assert.deepStrictEqual(ss.getSheets().filter(s => !s.isSheetHidden()).map(s => s.getName()), ["👋 Welcome"], 'only the Welcome tab');
  // the next open (the Welcome tab is there now; making it read B3 too): B3 read once, menu first again
  migrateReads = 0; order.length = 0;
  g.onOpen({ source: ss });
  assert.strictEqual(migrateReads, 1, 'Backend B3 read once');
  assert.strictEqual(order[0], 'menu:🚀 App Menu', 'menu first again');

  // a set-up, built sheet: every menu, then the filter (once)
  order.length = 0;
  dp.setProperty('SETUP_COMPLETE', 'true');
  dp.setProperty('ORIGINAL_SHEET_ID', 'FAKE_SS');
  g.studentListIsBuilt_ = () => true;
  g.onOpen();
  const f = order.indexOf('filter');
  assert(f > 0 && order.filter(x => x === 'filter').length === 1, 'set-up sheet: the filter is still re-made, once');
  ['menu:🚀 App Menu', 'menu:👤 Student Overrides', 'menu:📆 Events'].forEach(m =>
    assert(order.indexOf(m) !== -1 && order.indexOf(m) < f, m + ' before the filter'));
  console.log("the 🚀 App Menu first when a sheet opens: ok");
}
