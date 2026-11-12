// Input Data Mode / Your Lists / Welcome tab flows on a fake spreadsheet (lib/fakesheet.js):
//  - a wrong paste is cleared off the tab (and "Use it anyway" still works from the kept rows)
//  - a list can be uploaded (CSV / Excel rows) like a paste; a list with no rows is cleared
//  - step 5 builds the sheets first on a new sheet / after new data, and hides what the build showed
//  - the 👋 Welcome tab on a new copy: only tab shown, hidden by setup, deleted afterwards
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
const tabEmpty = t => t.getLastRow() === 1 && t.getLastColumn() === 1 && t.getRange("A1").getValue() === "PASTE HERE";

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
  g.leaveWelcomeTab_(ss); g.applyDefaultSheetVisibility_(ss);
  assert.deepStrictEqual(vis(), ["Contact Log", "Combined Contact Tracking", "PCAR", "Directory"], 'after setup: default tabs');
  assert.strictEqual(ss.getActiveSheet().getName(), "Contact Log");
  assert.strictEqual(g.deleteWelcomeTab_(ss), true, 'deleted at the next open / build');
  assert.strictEqual(ss.getSheetByName("👋 Welcome"), null);
  console.log("Welcome tab: ok");
}
