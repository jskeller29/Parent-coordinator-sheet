// Release check round 1 (owner's report, 2026-11-14): one test per problem fixed.
//  - Input Data Mode: the start page, counts, a busy sheet, tabs kept hidden, back from a list, a lost tab
//  - "build again" after Input Data Mode (lists changed after the build), Site / Class counts
//  - the Build panel: a fresh report for every run, Close + Build again
//  - Initial Setup: its steps on the progress bar, Sheet Lock outside the locked part
//  - Person Spoke With: the Backend copy, the old copy taken out / kept, an extra copy removed
//  - the END bar's line, the lookup box outline after a build
//  - Migration: a newer old tracker pauses itself (no Settings steps), the Welcome tab goes
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
const read = f => fs.readFileSync(path.join(ROOT, f), 'utf8');

const TABS = ["Contact Log", "Combined Contact Tracking", "PCAR", "Directory", "Raw Data", "Student List", "Notes", "Backend"];
function newSheet(opts) {
  const ss = makeSpreadsheet(TABS);
  ss.getSheetByName("Backend").hideSheet();
  const app = fakeSpreadsheetApp(ss);
  const g = loadProject(Object.assign({ extra: { SpreadsheetApp: app } }, opts || {}));
  g.htmlFromFile_ = name => { const o = { _name: name }; o.setTitle = () => o; o.setWidth = () => o; return o; };
  return { ss, g, app };
}
const vis = ss => ss.getSheets().filter(s => !s.isSheetHidden()).map(s => s.getName());
function atStep(g, i) { g.idm_open(); return g.idm_goToStep(i); }

// ---------------------------------------------------------------- counts: rows with no student, students counted (3.1)
{
  const { g } = newSheet();
  atStep(g, 2); // NYCSA
  const file = csv("3 - NYCSA (Students with Non-NYCSA Account).csv");
  let v = g.idm_submitRows(file, false);
  assert.strictEqual(v.notice.done.title, 'NYCSA completed!');
  assert.strictEqual(v.notice.done.sub, '25 rows moved to Raw Data (17 students). Next: Site / Class.', v.notice.done.sub);
  // A row whose student / guardian cells were emptied (only the school's code left) is not a row
  g.idm_goToStep(2);
  const edited = file.map(r => r.slice());
  for (let c = 3; c < edited[0].length; c++) edited[25][c] = "";           // the last row: only school code, district, name
  edited.splice(3, 1);                                                      // and one row deleted
  v = g.idm_submitRows(edited, false);
  assert.strictEqual(v.notice.done.sub, '23 rows moved to Raw Data (16 students). 1 row with no student or guardian was left out. Next: Site / Class.', v.notice.done.sub);
  assert.strictEqual(g.idm_loadState_().results.nycsa.rows, 23, 'the step remembers the real count');
  console.log("Input Data Mode counts (rows, students, empty rows left out): ok");
}

// ---------------------------------------------------------------- the start page (2.16)
{
  const { ss, g } = newSheet();
  const raw = ss.getSheetByName("Raw Data");
  raw.getRange("A1").setValue("ATS - DATE UPDATED: 11/01/26");
  const v = g.idm_open();
  assert(v.overview && v.start.length === 5, 'start page first');
  assert.deepStrictEqual(vis(ss), ["Start Here"]);
  assert.strictEqual(v.start[0].last.text, "11/01/26", 'last loaded date of Student Data');
  assert.strictEqual(v.start[1].last, null, 'Parent Square never loaded');
  assert(v.start.every(it => it.uses && it.where), 'what each is for + where to get it');
  const tab = ss.getSheetByName("Start Here");
  assert.strictEqual(tab.getRange(6, 1).getValue(), "Student Data (ATS) (needed)");
  assert(/11\/01\/26/.test(tab.getRange(6, 5).getValue()), 'the tab shows the date too');
  assert.strictEqual(g.idm_skip().overview, true, 'Skip does nothing on the start page');
  assert.strictEqual(g.idm_submitRows([["a"]], false).notice.code, 'overview', 'an upload asks to pick a step first');
  // Save and exit from the start page: nothing to save, the mode just closes
  const r = g.idm_saveAndExit();
  assert(r.exited && !r.notice, 'exit from the start page');
  console.log("Input Data Mode start page: ok");
}

// ---------------------------------------------------------------- the way back to the panel on every tab (2.16)
{
  const { ss, g } = newSheet();
  const hint = /App Menu → 📥 Input Data Mode/;
  g.idm_open();
  assert(hint.test(ss.getSheetByName("Start Here").getRange("A3").getValue()), 'start page');
  g.idm_goToStep(0);
  assert(hint.test(ss.getSheetByName("Student Data").getRange("A2").getValue()), 'paste tab');
  g.idm_goToStep(3);
  const site = ss.getSheets().filter(s => /Site/.test(s.getName()))[0];
  assert(hint.test(site.getRange("E1").getValue()), 'Site / Class tab (right of the table)');
  g.idm_goToStep(4);
  assert(hint.test(ss.getSheetByName("Your Lists").getRange("A3").getValue()), 'Your Lists tab');
  console.log("Input Data Mode: 'closed the side panel?' on every tab: ok");
}

// ---------------------------------------------------------------- tabs shown again are hidden by the full check (3.1)
{
  const { ss, g } = newSheet();
  atStep(g, 0);
  ss.getSheetByName("PCAR").showSheet(); ss.getSheetByName("Directory").showSheet(); // a build / Settings save showed them
  g.idm_check();                       // the user is still on the paste tab: the quick check doesn't look
  assert(vis(ss).indexOf("PCAR") !== -1, 'quick check: one call only');
  g.idm_check(true);                   // every ~10 s the sidebar asks for a full check
  assert.deepStrictEqual(vis(ss), ["Student Data"], 'full check: hidden again');
  console.log("Input Data Mode: a full check hides tabs shown again: ok");
}

// ---------------------------------------------------------------- exit brings the everyday tabs back on a built sheet (3.1)
{
  const { ss, g } = newSheet();
  ss.getSheetByName("Student List")._put([["OSIS"], ["123456789"]], 1, 1); // built
  ss.getSheetByName("Contact Log").activate();
  // A run whose state lost track of the tabs (the broken session): nothing remembered
  atStep(g, 0);
  const st = g.idm_loadState_(); st.hiddenIds = []; g.idm_saveState_(st);
  g.idm_exit();
  const shown = vis(ss);
  ["Contact Log", "Combined Contact Tracking", "PCAR", "Directory"].forEach(t => assert(shown.indexOf(t) !== -1, 'back: ' + t + ' in ' + shown));
  console.log("Input Data Mode exit: the everyday tabs come back on a built sheet: ok");
}

// ---------------------------------------------------------------- a busy sheet: wait, never "Lock timeout" (3.1)
{
  const { ss, g } = newSheet();
  let sidebars = 0;
  g.openInputDataModeSidebar_ = () => { sidebars++; };
  g.LockService = { getDocumentLock: () => ({ tryLock: () => false, waitLock() { throw new Error("Lock timeout: another process was holding the lock for too long."); }, releaseLock() {} }),
                    getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) };
  assert.doesNotThrow(() => g.openInputDataMode(), 'the menu item never throws on a busy sheet');
  assert.strictEqual(sidebars, 1, 'the side panel opens anyway (it waits by itself)');
  const v = g.idm_open();
  assert(v.busy && v.notice.type === 'busy', 'a busy reply, not an error: ' + JSON.stringify(v));
  const sb = read('InputDataSidebar.html');
  assert(/function waitBusy\(/.test(sb) && /if \(v && v\.busy\) return waitBusy\(v\)/.test(sb), 'the sidebar asks again by itself');
  console.log("Input Data Mode on a busy sheet: waits, no Lock timeout: ok");
}

// ---------------------------------------------------------------- a lost tab: "reload the page" (3.1)
{
  const { g } = newSheet();
  const v = g.idm_errorView_(new Error("Sheet 763596992 not found"));
  assert.strictEqual(v.notice.code, 'reload');
  assert(/Reload the page/.test(v.notice.message), v.notice.message);
  assert(/sheet\\s\+\\S\+\\s\+not found/.test(read('InputDataSidebar.html')), 'the sidebar says it for a failed call too');
  console.log("Input Data Mode: 'Sheet … not found' says to reload: ok");
}

// ---------------------------------------------------------------- the running mode's tab is never deleted as a leftover (3.1)
{
  const { ss, g } = newSheet();
  atStep(g, 0);
  const temp = ss.getSheetByName("Student Data");
  g.__props.INPUT_DATA_MODE_DELETE_TAB = String(temp.getSheetId());
  ss.getSheetByName("Contact Log").activate(); // the user is elsewhere for a moment
  g.idm_deleteLeftoverTab_();
  assert(ss.getSheetByName("Student Data"), 'kept: it is the running mode\'s tab');
  assert.strictEqual(g.__props.INPUT_DATA_MODE_DELETE_TAB, undefined, 'and no longer marked to delete');
  console.log("Input Data Mode: its own tab is never deleted as a leftover: ok");
}

// ---------------------------------------------------------------- back from a list's sidebar: the Your Lists tab (2.16)
{
  const { ss, g } = newSheet();
  let lists = [];
  g.od_loadLists_ = () => lists;
  g.od_saveLists_ = (s, ls) => { lists = ls; };
  g.od_showWizard_ = () => {};
  g.openInputDataModeSidebar_ = () => {};
  atStep(g, 4);
  const r = g.od_createList("Bus Routes", "🚌");
  const bus = ss.getSheets().filter(t => t.getSheetId() === r.tabId)[0];
  assert.strictEqual(ss.getActiveSheet(), bus, 'on the list tab while setting it up');
  g.od_backToInputMode();
  g.idm_open(); // the sidebar's first call
  assert.strictEqual(ss.getActiveSheet().getName(), "Your Lists", 'back on the Your Lists tab');
  assert(bus.isSheetHidden(), 'the list tab is hidden again');
  console.log("Your Lists → Back to Input Data Mode: the Your Lists tab: ok");
}

// ---------------------------------------------------------------- "build again" after the mode (2.16)
{
  const { ss, g } = newSheet();
  const s = { results: {}, step: 4 };
  assert.strictEqual(g.idm_buildNeededReason_(s), null, 'nothing came in');
  assert.strictEqual(g.idm_buildNeededReason_({ results: { lists: { saved: true, rows: 2, lists: true } } }), null, 'lists only looked at: no build');
  assert.strictEqual(g.idm_buildNeededReason_({ results: { ats: { rows: 18 } } }), 'data', 'data in, no build yet');
  assert.strictEqual(g.idm_buildNeededReason_({ results: { ats: { rows: 18 } }, listsBuild: { ok: true } }), null, 'data in, built at step 5');
  assert.strictEqual(g.idm_buildNeededReason_({ results: { ats: { rows: 18 } }, listsBuild: { ok: true }, listsChanged: true }), 'lists', 'a list saved after the build');
  assert.strictEqual(g.idm_buildNeededReason_({ results: {}, listsBuild: { ok: false }, dataChanged: false }), null);

  // The checklist: built, but a list changed after the build -> not ticked, says why; a build clears it
  ss.getSheetByName("Student List")._put([["OSIS"], ["123456789"]], 1, 1);
  g.zg_links_ = () => ({});
  atStep(g, 4);
  const st = g.idm_loadState_(); st.listsBuild = { ok: true }; st.listsChanged = true; g.idm_saveState_(st);
  const end = g.idm_finishLists();
  assert(end.finished && end.anyImported && end.buildReason === 'lists', 'end screen: build again, because of the lists: ' + JSON.stringify(end));
  assert.strictEqual(g.__props.BUILD_NEEDED, 'lists', 'remembered for the checklist');
  let step = g.zg_checklist_(ss).steps.filter(x => x.id === 'build')[0];
  assert(!step.done && /again/.test(step.title) && /lists/.test(step.text), 'checklist: ' + JSON.stringify(step));
  delete g.__props.BUILD_NEEDED; // what the end of a build does (Build.gs)
  step = g.zg_checklist_(ss).steps.filter(x => x.id === 'build')[0];
  assert(step.done && step.title === "Build your sheets", 'ticked after the build');
  assert(/deleteProperty\(BUILD_NEEDED_KEY_\)/.test(read('Build.gs')), 'every build clears it');
  const sb = read('InputDataSidebar.html');
  assert(/Build your sheets first\?/.test(sb), '"Next steps" asks to build first when the sheets lack the new data');
  console.log("Build again after Input Data Mode (end screen + checklist): ok");
}

// ---------------------------------------------------------------- Site / Class: counts for "Save and go" (2.11)
{
  const { ss, g } = newSheet();
  atStep(g, 3);
  const site = ss.getActiveSheet();
  site._put([["Y25", "Carl", ""], ["Y24", "", "Site 1"], ["X30", "Dana", "Site 2"]], 2, 1);
  let v = g.idm_check();
  assert.deepStrictEqual(JSON.parse(JSON.stringify(v.siteClass)), { classes: 3, teachers: 2, sites: 2 });
  site._put([["Y24", "Bob", "Site 1"]], 3, 1);
  v = g.idm_check();
  assert.strictEqual(v.siteClass.teachers, 3, 'every class has a teacher now');
  const sb = read('InputDataSidebar.html');
  assert(/c\.teachers === c\.classes \|\| c\.sites === c\.classes/.test(sb), 'ready when every class has a teacher OR a site');
  assert(/id="siteNextBtn"/.test(sb), 'the Save and go button');
  console.log("Site / Class counts (Save and go when filled in): ok");
}

// ---------------------------------------------------------------- the side panel: clicks never dropped, one Finish (3.1 / 2.16)
{
  const sb = read('InputDataSidebar.html');
  const fn = name => { const i = sb.indexOf('function ' + name + '('); return sb.slice(i, sb.indexOf('\n  }', i)); };
  assert(/cancelPoll\(\)/.test(fn('working')), 'a click drops the check in flight');
  assert(/if \(mySeq !== pollSeq\)/.test(fn('poll')), 'a stale check reply is ignored');
  assert(/asking = true/.test(fn('exitMode')) && /cancelPoll\(\)/.test(fn('exitMode')), 'the exit question stops the checks (they closed it)');
  assert(/if \(paused \|\| busy \|\| asking\) return;/.test(fn('schedule')), 'no checks while asking');
  assert(!/finishListsBtn/.test(sb), 'the Your Lists card has no Finish of its own any more');
  assert(/show\('exitBtn', mode !== 'lists'\)/.test(sb), 'on Your Lists only "Finish Input Data Mode" at the bottom');
  assert(/Got your paste\. Moving it to Raw Data/.test(sb) && /Got your file\. Moving it to Raw Data/.test(sb), 'the move is said at the top at once (2.9)');
  assert(/pcFollowBuild\(listsBar\)/.test(sb), 'the build before Your Lists shows the shared progress bar');
  console.log("Input Data Mode side panel: clicks, exit question, one Finish, the bar: ok");
}

// ---------------------------------------------------------------- the Build panel (3.1)
{
  const ps = read('ProgressSidebar.html');
  const onload = ps.slice(ps.indexOf('window.onload'), ps.indexOf('function pollServer'));
  assert(!/getBuildLogData/.test(onload), 'no old report when the panel opens');
  const fresh = ps.slice(ps.indexOf('function startFreshBuild'), ps.indexOf('function startFreshBuild') + 600);
  assert(/withSuccessHandler\(\(\) => \{[\s\S]*startBuildOnlyProcess\(\)[\s\S]*\.resetSyncStatus\(\)/.test(fresh), 'a build panel clears the last run\'s state first, then starts (in the reply)');
  assert(/prevPhase !== 'building'\) \{[^}]*cachedBuildLogData = null/.test(ps), 'a new build forgets the old report');
  assert(/id="btnClose"/.test(ps) && /'Build again'/.test(ps) && !/'Start new build'/.test(ps), 'Close + Build again');
  assert(/pcBar\(el\('buildBar'\)/.test(ps), 'the shared progress bar');
  // The report says when it was built
  const { g } = newSheet();
  assert(/builtAt: Date\.now\(\)/.test(read('Build.gs')), 'builtAt in the report');
  const st = g.getBuildProgress_();
  assert(st.buildSteps.length === 8 && typeof st.now === 'number', 'progress: steps + server time');
  g.reportBuildStep_(3);
  const st2 = g.getBuildProgress_();
  assert(st2.buildStep === 3 && st2.buildStepAt > 0, 'step + when');
  console.log("Build panel: fresh report each run, Close / Build again, the bar: ok");
}

// ---------------------------------------------------------------- the shared progress bar (Styles.html)
{
  const st = read('Styles.html');
  assert(/function pcBar\(/.test(st) && /function pcFollowBuild\(/.test(st), 'pcBar + pcFollowBuild');
  assert(/st\.buildStepAt >= serverStart/.test(st), 'a step left from the build before is ignored');
  // it runs: a tiny DOM stand-in
  const els = {};
  const mk = () => ({ style: {}, set innerHTML(h) { this._h = h; }, get innerHTML() { return this._h; }, setAttribute() {}, textContent: '', className: '' });
  const root = mk(); root.querySelector = c => (els[c] = els[c] || mk());
  const el = { set innerHTML(h) {}, get firstChild() { return root; } };
  const vm = require('vm');
  const ctx = { setInterval: () => 1, clearInterval() {}, setTimeout() {}, Date, Math, String, Number };
  vm.createContext(ctx);
  vm.runInContext(st.slice(st.indexOf('<script>') + 8, st.indexOf('</script>')), ctx);
  const bar = ctx.pcBar(el, { title: 'Building your sheets' });
  bar.set(2, [{ label: 'A' }, { label: 'B' }, { label: 'C', detail: 'third' }, { label: 'D' }]);
  assert.strictEqual(els['.pb-label'].textContent, 'Step 3 of 4: C');
  assert.strictEqual(els['.pb-sub'].textContent, 'third');
  assert(parseInt(els['.pb-fill'].style.width, 10) >= 50, 'half way at step 3 of 4: ' + els['.pb-fill'].style.width);
  bar.done('Sheets built');
  assert.strictEqual(els['.pb-label'].textContent, 'Sheets built');
  assert.strictEqual(els['.pb-fill'].style.width, '100%');
  ['ProgressSidebar.html', 'InputDataSidebar.html', 'MigrationUI.html'].forEach(f => assert(/pcBar\(/.test(read(f)), f + ' uses it'));
  assert(/pcBar\(document\.getElementById\('setupBar'\)/.test(read('Menus.gs')), 'the setup window uses it');
  console.log("The shared progress bar: ok");
}

// ---------------------------------------------------------------- Initial Setup: steps on the bar, Sheet Lock outside the lock (2.3 / 3.1)
{
  const { ss, g } = newSheet();
  ss.getSheetByName("Student List")._put([["OSIS"], ["123456789"]], 1, 1); // a copy of a built tracker
  g.isAutoMigrateOn_ = () => false;
  g.zg_safeGuideLink_ = () => ""; g.zg_safeExampleLink_ = () => "";
  const steps = [];
  const rep0 = g.reportSetupStep_;
  g.reportSetupStep_ = i => { steps.push(i); rep0(i); };
  let locks = 0, lockHeld = false, lockedDuringSheetLock = false;
  g.LockService = { getDocumentLock: () => ({ waitLock() { lockHeld = true; }, tryLock() { lockHeld = true; return true; }, releaseLock() { lockHeld = false; } }),
                    getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) };
  g.applySheetLocks_ = () => { locks++; if (lockHeld) lockedDuringSheetLock = true; };
  g.runInitialSetup_();
  assert.deepStrictEqual(steps, [0, 1, 2, 3, 4], 'every setup step reported');
  const p = g.zg_setupProgress();
  assert(p.step === 4 && p.at > 0 && p.steps.length === 5, 'the window reads them: ' + JSON.stringify(p));
  assert.strictEqual(locks, 0, 'no Sheet Lock inside the setup (it held the lock Input Data Mode waits for)');
  g.zg_setupFinish();
  assert.strictEqual(locks, 1, 'Sheet Lock in the window\'s second call, on a built sheet');
  assert.strictEqual(lockedDuringSheetLock, false, '...outside the document lock');
  const m = read('Menus.gs');
  assert(/zg_setupProgress\(\)/.test(m) && /Still working: a big sheet can take a few minutes/.test(m), 'the window polls and says when it is slow');
  console.log("Initial Setup: steps on the bar, Sheet Lock after the lock: ok");
}

// ---------------------------------------------------------------- the Welcome tab's first line (3.2)
{
  const { ss, g } = newSheet();
  g.isAutoMigrateOn_ = () => false;
  const t = g.buildWelcomeTab_(ss, {});
  assert(/Don't see 🚀 App Menu/.test(t.getRange("C1").getValue()) && /reload the page/.test(t.getRange("C1").getValue()), 'row 1: reload if the menu is missing');
  assert(/experienced an error/.test(t.getDataRange().getValues().map(r => r.join(" ")).join(" ")), 'step 3: what to do after "experienced an error"');
  const d = g.buildWelcomeTab_(ss, { done: true, into: t });
  assert(/Don't see 🚀 App Menu/.test(d.getRange("C1").getValue()), 'after setup too');
  assert.strictEqual(d.getRange("C3").getValue(), "You're set up", 'the title stays in C3 (read there)');
  console.log("Welcome tab: 'App Menu missing? reload' at the top: ok");
}

// ---------------------------------------------------------------- Person Spoke With: the Backend copy (4.3 / 4.4 / 2.19)
function spokeSheet() {
  const { ss, g } = newSheet();
  const log = ss.getSheetByName("Contact Log");
  log._put([["Date", "Guardian", "Person", "OSIS"]], 1, 2);
  log._put([["Morticia Addams and Gomez Addams", "", "234567891"], ["Marge Simpson and Homer Simpson", "", "123456789"]], 2, 3);
  log._set(5, 5, "END");
  return { ss, g, log, backend: ss.getSheetByName("Backend") };
}
{
  // The dropdown still reads the old place (R:AA): the R2 formula is kept; the Backend copy is written too
  const { g, log, backend } = spokeSheet();
  let r = g.ensureSpokeWithChoices_(g.SpreadsheetApp.getActiveSpreadsheet());
  assert.strictEqual(r, "added", 'the old place answers as before: ' + r);
  assert.strictEqual(log.getRange("R2").getFormula(), g.__get("SPOKE_WITH_CHOICES_FORMULA"), 'old place: R2');
  assert(/'Contact Log'!C:C/.test(backend.getRange(2, 84).getFormula()) && /SPLIT\(/.test(backend.getRange(2, 84).getFormula()), 'Backend CF2: the same list, read from the Contact Log');
  assert(/PERSON SPOKE WITH CHOICES/.test(backend.getRange(1, 84).getValue()), 'its title');
  assert(backend.getMaxRows() >= 5 + 300, 'room for every log row + 300');
  assert.strictEqual(g.__props.SPOKE_CHOICES_MODE, 'contactlog');
  // a second copy further down (two formulas, one #REF!; 2.19): the build keeps only R2's
  log.getRange("R4").setFormula(g.__get("SPOKE_WITH_CHOICES_FORMULA"));
  r = g.ensureSpokeWithChoices_(g.SpreadsheetApp.getActiveSpreadsheet());
  assert(/1 extra copy removed/.test(r) && log.getRange("R4").getFormula() === "", r);
  // a row change while the dropdown reads R:AA: R2 put back as before
  log.getRange("R2").clearContent();
  assert.strictEqual(g.restoreSpokeWithChoicesAfterRowChange_(g.SpreadsheetApp.getActiveSpreadsheet(), log), "added");
  console.log("Person Spoke With (dropdown on R:AA): kept, extra copy removed, Backend copy ready: ok");
}
{
  // The owner set column D's dropdown to 'Backend'!$CF2:$CO2: the R2 formula is taken out (frees "Delete table row")
  const { ss, g, log, backend } = spokeSheet();
  log.getRange("R2").setFormula(g.__get("SPOKE_WITH_CHOICES_FORMULA"));
  log._dv.set("2,4", { getCriteriaValues: () => [backend.getRange(2, 84, 1, 10)] });
  const r = g.ensureSpokeWithChoices_(ss);
  assert(/^Backend added; the old formula in Contact Log R taken out \(1\)/.test(r), r);
  assert.strictEqual(log.getRange("R2").getFormula(), "", 'R2 empty now');
  assert.strictEqual(g.__props.SPOKE_CHOICES_MODE, 'backend');
  // row deletes / inserts: nothing read or written (no lag)
  const before = ss._log.length;
  let reads = 0; const gr = log.getRange; log.getRange = function () { reads++; return gr.apply(log, arguments); };
  assert.strictEqual(g.restoreSpokeWithChoicesAfterRowChange_(ss, log), "backend: nothing to do");
  assert.strictEqual(reads, 0, 'no sheet call on a row change');
  log.getRange = gr;
  assert.strictEqual(ss._log.length, before);
  // the next build: kept as it is
  assert(/^Backend kept/.test(g.ensureSpokeWithChoices_(ss)));
  console.log("Person Spoke With (dropdown on the Backend): old formula out, row changes cost nothing: ok");
}

// ---------------------------------------------------------------- the END bar's line + the box outline (4.4)
{
  const e = read('Edits.gs');
  assert(!/setBorder\(true, null, true, null, null, null, "#000000"/.test(e), 'no thin line under the END bar any more');
  assert.strictEqual((e.match(/setBorder\(true, null, null, null, null, null, "#000000", SpreadsheetApp\.BorderStyle\.SOLID\)/g) || []).length, 2, 'the line above it, on an edit and after a row delete');
  assert(/redrawLookupBoxOutlineForBuild_\(ss\)/.test(read('Build.gs')), 'every build draws the box outline again');
  console.log("END bar line + lookup box outline: ok");
}

// ---------------------------------------------------------------- Migration: a newer old tracker pauses itself (3.3)
{
  const { ss, g } = newSheet();
  const old = makeSpreadsheet(["Contact Log", "Backend"]);
  g.SpreadsheetApp.openByUrl = () => old;
  // no mark, no date: can't tell -> the Settings steps stay (as before)
  let r = g.ma_markOldTracker("https://old");
  assert(r.marked && r.pausesItself === null, JSON.stringify(r));
  // a version date before 11/14/2026: older
  old.getSheetByName("Backend").getRange("A2").setValue(new Date(2026, 9, 1));
  assert.strictEqual(g.ma_markOldTracker("https://old").pausesItself, false);
  // a newer date, or the mark its builds leave: pauses itself
  old.getSheetByName("Backend").getRange("A2").setValue(new Date(2026, 10, 20));
  assert.strictEqual(g.ma_markOldTracker("https://old").pausesItself, true);
  old.getSheetByName("Backend").getRange("A2").setValue(new Date(2026, 9, 1));
  g.markTrackerCode_(old);
  assert.strictEqual(g.ma_markOldTracker("https://old").pausesItself, true, 'the mark wins over an old date (the owner\'s own sheet)');
  assert.strictEqual(g.markTrackerCode_(old), false, 'marked once');
  const ui = read('MigrationUI.html');
  assert(/id="oldPaused"/.test(ui) && /res\.marked && res\.pausesItself === true/.test(ui), 'the done screen without Settings steps');
  assert(/markTrackerCode_\(ss\)/.test(read('Build.gs')), 'every build leaves the mark');
  // the Welcome tab: moved off in step 6, deleted by the next call
  g.isAutoMigrateOn_ = () => false;
  g.buildWelcomeTab_(ss, {});
  ss.getSheetByName("Contact Log").activate();
  g.__props.SETUP_COMPLETE = 'true';
  g.ma_markOldTracker("https://old");
  assert.strictEqual(ss.getSheetByName("👋 Welcome"), null, 'the Welcome tab is gone at the end of the migration');
  console.log("Migration: newer old trackers pause themselves; the Welcome tab goes: ok");
}

console.log("release check round 1 fixes: ok");
