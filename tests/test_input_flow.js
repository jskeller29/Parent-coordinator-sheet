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
  assert.deepStrictEqual(vis(), ["Contact Log", "Combined Contact Tracking", "PCAR", "Directory"], 'after setup: default tabs');
  assert.strictEqual(g.deleteWelcomeTab_(ss), true, 'deleted at the next open');
  console.log("Welcome tab after Reset to Display Mode: ok");
}
{
  // Initial Setup asks again for any permission left unticked (granular consent)
  const src = fs.readFileSync(path.join(ROOT, 'Menus.gs'), 'utf8');
  const setup = src.slice(src.indexOf('function runInitialSetupWrapper()'));
  assert(/requireAllScopes\(ScriptApp\.AuthMode\.FULL\)/.test(setup.slice(0, 600)), 'requireAllScopes first in Initial Setup');
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
  // Owner's 🐞 Debug → 🔄 Copy Sheet and Set to Migrate: the COPY gets Backend B3 = Yes before
  // its reset, so its Welcome tab is the upgrade one; this sheet is not touched
  const mine = makeSpreadsheet(TABS);
  const copy = makeSpreadsheet(TABS);
  copy.getId = () => "COPY_SS"; copy.getUrl = () => "https://example.org/copy";
  const app = fakeSpreadsheetApp(mine);
  app.newRichTextValue = () => { const v = {}; const b = new Proxy({}, { get: (t, k) =>
    k === 'setText' ? (x => { v.text = x; return b; }) : k === 'build' ? (() => v) : (() => b) }); return b; };
  const g = loadProject({ owner: true, extra: { SpreadsheetApp: app } });
  const shown = [];
  g.ownerCopyThisSheet_ = (ss, label, name) => { assert(/^Parent Coordinator Sheet - \(MIGRATE\) \(/.test(name), name); return copy; };
  g.ownerShowCopyLink_ = (url, title) => shown.push(title);
  g.getOrCreateVersionSheet_ = ss => ss.getSheetByName("Backend");
  g.getGuideLink_ = () => "https://example.org/guide";
  let b3AtReset = null;
  g.resetToDisplayModeSteps_ = ss => { b3AtReset = ss.getSheetByName("Backend").getRange("B3").getValue(); g.showWelcomeAfterReset_(ss); };
  g.ownerMenuCopyAndSetToMigrate();
  assert.strictEqual(b3AtReset, "Yes", 'B3 = Yes before the reset');
  assert.strictEqual(copy.getSheetByName("Backend").getRange("B3").getValue(), "Yes", 'the copy is an upgrade template');
  assert.strictEqual(mine.getSheetByName("Backend").getRange("B3").getValue(), "", 'this sheet is not touched');
  const text = Array.from(copy.getSheetByName("👋 Welcome")._cells.values()).join("\n");
  assert(/Bring over your old data/.test(text), 'the copy\'s Welcome tab: the upgrade step');
  assert.deepStrictEqual(copy.getSheets().filter(s => !s.isSheetHidden()).map(s => s.getName()), ["👋 Welcome"]);
  assert.strictEqual(mine.getSheetByName("👋 Welcome"), null, 'no Welcome tab on this sheet');
  assert(/set to Migrate/.test(shown[0]), 'done message: ' + shown[0]);
  console.log("Debug → Copy Sheet and Set to Migrate: ok");
}
{
  // 🐞 Debug → 📦 Make All 3 Copies: Blank (B3 = No) and Migrate (B3 = Yes) are reset copies,
  // Example is the fake-data copy; 🔗 My Links lists every link from the YOUR LINKS box
  const mine = makeSpreadsheet(TABS);
  mine.getName = () => "Mine";
  const app = fakeSpreadsheetApp(mine);
  const g = loadProject({ owner: true, extra: { SpreadsheetApp: app } });
  const made = [];
  g.ownerCopyThisSheet_ = (ss, label, name) => {
    const n = made.length, c = makeSpreadsheet(TABS); c.getId = () => "COPY_" + n; c.getUrl = () => "https://example.org/" + n;
    c.getName = () => name; made.push({ label: name, copy: c }); return c;
  };
  g.stampVersionDate_ = c => c.getSheetByName("Backend").getRange("A2").setValue(new Date(2026, 10, 12));
  g.Utilities.formatDate = d => (d.getMonth() + 1 + "").padStart(2, "0") + "/" + String(d.getDate()).padStart(2, "0") + "/" + String(d.getFullYear()).slice(2);
  g.getOrCreateVersionSheet_ = ss => ss.getSheetByName("Backend");
  const steps = [];
  g.resetToDisplayModeSteps_ = c => steps.push(["reset", c.getSheetByName("Backend").getRange("B3").getValue()]);
  g.ownerRunOnCopy_ = (c, fn) => { steps.push(["fake"]); };
  const b = g.ownerMakeCopy("blank"), m = g.ownerMakeCopy("migrate"), e = g.ownerMakeCopy("example");
  // 3. the release names: never this sheet's name ("Mine")
  assert.deepStrictEqual(made.map(x => x.label), ["BLANK", "MIGRATE", "EXAMPLE"].map(t => "Parent Coordinator Sheet - (" + t + ") (11/12/26)"));
  assert.deepStrictEqual(steps, [["reset", "No"], ["reset", "Yes"], ["fake"]], 'B3 set before each reset; Example loads fake data');
  assert.strictEqual(made[0].copy.getSheetByName("Backend").getRange("B3").getValue(), "No");
  assert.strictEqual(made[1].copy.getSheetByName("Backend").getRange("B3").getValue(), "Yes");
  // 1. every copy gets the version date, the Example one too (it has no reset), and B3 = No there
  made.forEach(x => assert(x.copy.getSheetByName("Backend").getRange("A2").getValue() instanceof Date, 'version date in ' + x.label));
  assert.strictEqual(made[2].copy.getSheetByName("Backend").getRange("B3").getValue(), "No", 'Example: not an upgrade template');
  assert.strictEqual(e.stamp, "11/12/26", 'the pop-up shows the version date');
  assert.strictEqual(JSON.stringify(b), JSON.stringify({ name: "Parent Coordinator Sheet - (BLANK) (11/12/26)", url: "https://example.org/0", stamp: "11/12/26" }));
  assert(m.url && e.url);
  assert.throws(() => g.ownerMakeCopy("other"), /Unknown copy/);
  let dialog = "";
  app._ui.showModalDialog = (h, title) => { dialog = title + "\n" + h._html; };
  g.include_ = () => "";
  g.HtmlService = { createHtmlOutput: h => { const o = { _html: h }; o.setWidth = () => o; o.setHeight = () => o; return o; } };
  g.ownerMenuMakeAllCopies();
  assert(/ownerMakeCopy/.test(dialog) && /Parent Coordinator Sheet - \(BLANK\)/.test(dialog) && /\(MIGRATE\)/.test(dialog) && /\(EXAMPLE\)/.test(dialog), 'the 3-copies pop-up');
  assert(/Version date in all three/.test(dialog), 'the pop-up says which date goes in the Version Tracker');
  g.getGuideLink_ = () => "https://example.org/guide.pdf";
  g.ownerMenuMyLinks();
  Object.keys(g.__get("GUIDE_LINKS")).forEach(k => assert(dialog.indexOf(k) !== -1, 'My Links lists ' + k));
  assert(/YOUR LINKS/.test(dialog) && /guide\.pdf/.test(dialog), 'My Links: where to change them + the newest guide file');
  // The YOUR LINKS box is the first thing in Menus.gs
  const src = fs.readFileSync(path.join(ROOT, 'Menus.gs'), 'utf8');
  assert(src.indexOf('YOUR LINKS') < src.indexOf('function '), 'links box at the top of Menus.gs');
  console.log("Debug → Make All 3 Copies + My Links: ok");
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
