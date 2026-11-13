// Menu features on the fake spreadsheet (lib/fakesheet.js):
//  - Migration: refuses this tracker's own link and a sheet that isn't a tracker, makes room for a
//    bigger Raw Data, and gives the copied event tables their look and dropdowns back
//  - Rollover: the archive is named for the school year that is ending
//  - dialogs: Settings can't save before your settings have loaded; New Student checks the OSIS
const fs = require('fs'), path = require('path'), assert = require('assert');
const { loadProject, ROOT } = require('./lib/load');
const { makeSpreadsheet, fakeSpreadsheetApp } = require('./lib/fakesheet');

// ---------------------------------------------------------------- Migration: which sheet
{
  const here = makeSpreadsheet(["Contact Log", "Raw Data", "Events", "Backend"]);
  const app = fakeSpreadsheetApp(here);
  const old = makeSpreadsheet(["Contact Log", "Raw Data", "Events"]); old.getId = () => "OLD"; old.getName = () => "PC Tracker 2025";
  const gradebook = makeSpreadsheet(["Grades"]); gradebook.getId = () => "GRADES"; gradebook.getName = () => "Gradebook";
  app.openByUrl = url => url === "here" ? here : url === "old" ? old : url === "grades" ? gradebook : (() => { throw new Error("no access"); })();
  const g = loadProject({ extra: { SpreadsheetApp: app }, noTrialFiles: true });
  assert(/this tracker/.test(g.ma_verify("here").error), "this tracker's own link is refused");
  assert(/doesn't look like a PC Tracker/.test(g.ma_verify("grades").error), 'a sheet with no Contact Log is refused');
  assert(/Could not open/.test(g.ma_verify("nope").error), 'no access');
  assert.strictEqual(g.ma_verify("old").success, true, 'last year\'s tracker is fine');

  // Raw Data: last year's tab is bigger than the template's
  const oldRaw = old.getSheetByName("Raw Data"), newRaw = here.getSheetByName("Raw Data");
  newRaw._maxRows = 50; newRaw._maxCols = 26;
  for (let r = 1; r <= 120; r++) oldRaw._set(r, 3, "OSIS" + r);
  oldRaw._set(1, 61, "BI1");
  g.ma_migrateRawData("old");
  assert(newRaw.getMaxRows() >= 120 && newRaw.getMaxColumns() >= 61, 'the tab grew to fit (setValues throws outside it)');
  assert.strictEqual(newRaw.getRange(120, 3).getValue(), "OSIS120", 'every row came over');
  assert.strictEqual(newRaw.getRange(1, 61).getValue(), "BI1", 'every column came over');
  console.log("Migration: which sheet + a bigger Raw Data: ok");
}

// ---------------------------------------------------------------- Migration: the old tracker pauses
{
  const here = makeSpreadsheet(["Contact Log", "Backend"]); here.getName = () => "PC Tracker 2026";
  const old = makeSpreadsheet(["Contact Log"]); old.getId = () => "OLD"; old.getName = () => "PC Tracker 2025";
  const app = fakeSpreadsheetApp(here);
  app.openByUrl = url => url === "old" ? old : (() => { throw new Error("no access"); })();
  const g = loadProject({ extra: { SpreadsheetApp: app }, noTrialFiles: true });
  const r = g.ma_markOldTracker("old");
  assert(r.marked && r.name === "PC Tracker 2025" && /OLD/.test(r.url), 'the old tracker is marked and named');
  g.ma_markOldTracker("old");
  assert.strictEqual(old._meta.length, 1, 'one mark: a later migration replaces it');
  const mark = JSON.parse(old._meta[0].value);
  assert(mark.old === "OLD" && mark.newId === "FAKE_SS" && mark.newName === "PC Tracker 2026", 'the mark names both trackers');
  assert.strictEqual(g.ma_markOldTracker("nope").marked, false, 'no access: not marked, no error');

  // The OLD tracker (newer code): its nightly skips while the new tracker exists; nothing is deleted or changed
  let trashed = false;
  const drive = { getFileById: id => { if (id !== "FAKE_SS") throw new Error("not found"); return { isTrashed: () => trashed }; } };
  const deleted = [];
  const trig = (uid, fn) => ({ getUniqueId: () => uid, getHandlerFunction: () => fn });
  const scriptApp = { getProjectTriggers: () => [trig("T1", "nightlyMasterSync")], deleteTrigger: t => deleted.push(t.getUniqueId()) };
  const oldApp = fakeSpreadsheetApp(old);
  let answer = null; const asked = [];
  const ui = oldApp.getUi();
  ui.Button = { YES: 'YES', NO: 'NO', CLOSE: 'CLOSE' }; ui.ButtonSet = { YES_NO: 'YES_NO' };
  ui.alert = (title, msg) => { asked.push(title + " | " + msg); return answer; };
  const o = loadProject({ extra: { SpreadsheetApp: oldApp, ScriptApp: scriptApp, DriveApp: drive }, noTrialFiles: true });
  o.__props.syncToPhones = 'true';
  assert.strictEqual(o.nightlyTriggerShouldRun_({ triggerUid: "T1" }), false, 'the replaced tracker skips its nightly run');
  assert.deepStrictEqual(deleted, [], 'no trigger deleted');
  assert.strictEqual(o.__props.syncToPhones, 'true', 'settings unchanged (undo = take the mark away)');
  assert.strictEqual(o.__props.disableNightlySync, undefined);

  // Opening it asks once: closed -> asked again; Yes -> only a reminder after that
  answer = 'CLOSE';
  assert.strictEqual(o.retiredTrackerOnOpen_(), true, 'closed: stays paused');
  answer = 'YES';
  assert.strictEqual(o.retiredTrackerOnOpen_(), true);
  assert.strictEqual(asked.length, 2, 'asked again after a close');
  assert(/PC Tracker 2026/.test(asked[1]), 'the question names the new tracker');
  old._toasts.length = 0;
  assert.strictEqual(o.retiredTrackerOnOpen_(), true);
  assert.strictEqual(asked.length, 2, 'not asked again after Yes');
  assert(/paused/.test(old._toasts[0].msg), 'a reminder instead');
  // No = still in use: the mark goes, nightly runs again
  o.__props.REPLACED_BY_ACK = '';
  answer = 'NO';
  assert.strictEqual(o.retiredTrackerOnOpen_(), false, 'No: the open goes on as usual');
  assert.strictEqual(old._meta.length, 0, 'the mark is gone');
  assert.strictEqual(o.nightlyTriggerShouldRun_({ triggerUid: "T1" }), true, 'nightly runs again');

  // The new tracker thrown away (a test copy in the trash): the mark is removed by itself
  g.ma_markOldTracker("old");
  trashed = true;
  assert.strictEqual(o.nightlyTriggerShouldRun_({ triggerUid: "T1" }), true, 'new tracker trashed: runs');
  assert.strictEqual(old._meta.length, 0, 'and the mark is gone');
  // The new tracker can't be read (no access): no pause, mark kept
  g.ma_markOldTracker("old");
  const o3 = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(old), DriveApp: { getFileById: () => { throw new Error("no access"); } } }, noTrialFiles: true });
  assert.strictEqual(o3.nightlyTriggerShouldRun_({ triggerUid: "T1" }), true, 'unreadable new tracker: runs');
  assert.strictEqual(old._meta.length, 1, 'mark kept');

  // A Drive copy of the marked tracker (other ID) keeps its nightly
  const copy = makeSpreadsheet(["Contact Log"]); copy.getId = () => "COPY"; copy._meta = old._meta.slice();
  const c = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(copy), DriveApp: drive }, noTrialFiles: true });
  trashed = false;
  assert.strictEqual(c.nightlyTriggerShouldRun_({ triggerUid: "X" }), true, 'a copy of it still runs');
  // The new tracker itself is not affected
  assert.strictEqual(g.nightlyTriggerShouldRun_({ triggerUid: "N" }), true, 'the new tracker runs');
  console.log("Migration: the old tracker pauses (and can be undone): ok");
}

// ---------------------------------------------------------------- Migration: final run-through (2026-11-13)
{
  const mk = () => {
    const here = makeSpreadsheet(["Contact Log", "Raw Data", "Events", "Send Out", "Notes", "Student List", "Backend"]);
    const old = makeSpreadsheet(["Contact Log", "Raw Data", "Events", "Send Out", "Notes", "Student List", "Backend"]);
    old.getId = () => "OLD"; old.getName = () => "PC Tracker 2025";
    const app = fakeSpreadsheetApp(here);
    app.openByUrl = () => old;
    const g = loadProject({ extra: { SpreadsheetApp: app }, noTrialFiles: true });
    return { here, old, app, g };
  };

  // Never backwards: a NEWER tracker can't be brought into an older one
  {
    const { here, old, g } = mk();
    here.getSheetByName("Backend")._set(2, 1, new Date(2026, 10, 1));
    old.getSheetByName("Backend")._set(2, 1, new Date(2026, 10, 12));
    const v = g.ma_verify("old");
    assert(!v.success && /newer version/.test(v.error), 'a newer tracker is refused');
    old.getSheetByName("Backend")._set(2, 1, new Date(2025, 8, 1));
    assert.strictEqual(g.ma_verify("old").success, true, 'an older one is fine');
    old.getSheetByName("Backend")._set(2, 1, "");
    assert.strictEqual(g.ma_verify("old").success, true, 'no date (very old tracker) is fine');
  }

  // Settings: the old tracker's saved block (not the notes date range); without one, the old guess
  {
    const { old, g } = mk();
    const ob = old.getSheetByName("Backend");
    ob._set(5, 1, "SAVED SETTINGS"); ob._set(6, 1, "syncToPhones"); ob._set(6, 2, false);
    ob._set(7, 1, "disableRebuilds"); ob._set(7, 2, true); ob._set(8, 1, "notesStartDate"); ob._set(8, 2, "2025-09-01");
    ob._set(9, 1, "dataReminder"); ob._set(9, 2, "month"); ob._set(10, 1, "notAKey"); ob._set(10, 2, "x");
    const ph = old.insertSheet("Phone Contacts"); // shown: the old guess would have turned phone sync ON
    const res = g.ma_migrateSettings("old");
    assert.strictEqual(res.saved, 3, 'three saved settings copied');
    assert.strictEqual(g.__props.syncToPhones, 'false', 'phone sync as the old tracker had it');
    assert.strictEqual(g.__props.disableRebuilds, 'true');
    assert.strictEqual(g.__props.dataReminder, 'month');
    assert.strictEqual(g.__props.notesStartDate, undefined, "last year's notes date range is not copied");
    assert.strictEqual(g.__props.notAKey, undefined);
    // no saved block: the guess from the tabs, as before
    const { old: old2, g: g2 } = mk();
    old2.insertSheet("Phone Contacts");
    assert.strictEqual(g2.ma_migrateSettings("old").saved, 0);
    assert.strictEqual(g2.__props.syncToPhones, 'true', 'guessed from the shown Phone Contacts tab');
  }

  // Step 4: Notes tab, Type List, Send Out by header (a refused column is skipped, the rest still comes over)
  {
    const { here, old, g } = mk();
    old.getSheetByName("Notes")._put([["OSIS", "Guardian", "Student", "Note", "Delete"],
      [111, "Ann Lee", "Bo Lee", "Prefers calls after 5 PM", false], ["", "", "", "", false], [222, "Cy Diaz", "Di Diaz", "Spanish letters", false]]);
    here.getSheetByName("Notes")._put([["OSIS", "Guardian", "Student", "Note", "Delete"], [999, "Template", "Row", "old", false]]);
    const ob = old.getSheetByName("Backend"), nb = here.getSheetByName("Backend");
    ob._set(1, 64, "TYPE LIST"); ob._set(2, 64, "Bus"); ob._set(3, 64, "Attendance");
    nb._set(1, 64, "TYPE LIST"); nb._set(2, 64, "Attendance"); nb._set(3, 64, "Behavior");
    old.getSheetByName("Send Out")._put([["Date", "Send Out Title", "Send Out Type", "Attachments"],
      ["10/1/2025", "Welcome Letter", "Services", "flyer.pdf"], ["10/2/2025", "PTC Invite", "Events", ""]]);
    const so = here.getSheetByName("Send Out");
    so._put([["Date", "Send Out Title", "Send Out Type", "On Behalf Of", "Where Did You Send", "Population", "# Reached", "Attachments"]]);
    const origGetRange = so.getRange;
    so.getRange = (a, b, c, d) => {
      const rg = origGetRange(a, b, c, d);
      if (typeof a === 'number' && b === 8 && a === 2) return new Proxy(rg, { get(t, k) { return k === 'setValues' ? () => { throw new Error("file chips only"); } : t[k]; } });
      return rg;
    };
    old.getSheetByName("Events")._put([["Fall PTC", "", "", "", "", "", ""], ["OSIS", "Student", "Guardian", "Type", "Where", "Here?", "Notes"]]);
    const res = g.ma_migrateSimpleSheets("old");
    const notes = here.getSheetByName("Notes");
    assert.strictEqual(notes.getRange(2, 4).getValue(), "Prefers calls after 5 PM", 'Notes tab copied');
    assert.strictEqual(notes.getRange(3, 4).getValue(), "Spanish letters", 'blank rows left out');
    assert.strictEqual(notes.getRange(3, 1).getValue(), 222);
    assert.strictEqual(nb.getRange(2, 64).getValue(), "Bus", 'old types first');
    assert.deepStrictEqual([2, 3, 4].map(r => nb.getRange(r, 64).getValue()), ["Bus", "Attendance", "Behavior"], "then the template's new ones");
    assert.strictEqual(so.getRange(1, 2).getValue(), "Send Out Title", 'the new headers are kept');
    assert.strictEqual(so.getRange(2, 2).getValue(), "Welcome Letter", 'Send Out rows by header');
    assert.strictEqual(so.getRange(3, 3).getValue(), "Events");
    assert(res.warnings.some(w => /Attachments/.test(w)), 'the refused column is reported');
    assert.strictEqual(here.getSheetByName("Events").getRange(1, 1).getValue(), "Fall PTC", 'Events still came over');
  }

  // Step 5 returns its count (no pop-up), step 6 brings the Student List history and undoes it if the build fails
  {
    const { here, old, app, g } = mk();
    const nl = here.getSheetByName("Contact Log"); nl._put([["Week", "Date", "Guardian", "Person Spoke With", "OSIS", "Student", "#"]]); nl._set(6, 5, "END");
    const ol = old.getSheetByName("Contact Log"); ol._put([["Week", "Date", "Guardian", "Person Spoke With", "OSIS", "Student", "#"],
      ["", "10/1/2025", "Ann Lee", "Ann Lee", "111", "Bo Lee", 1]]); ol._set(3, 5, "END");
    const ui = app.getUi(); ui._alerts.length = 0;
    assert.strictEqual(g.ma_migrateContactLog("old").records, 1, 'step 5 returns the count');
    assert.strictEqual(ui._alerts.length, 0, 'no pop-up in the middle of the wizard');

    old.getSheetByName("Student List")._put([["OSIS", "Student", "Guardian", "Site", "Class", "Sibling", "Status"],
      ["111", "Bo Lee", "Ann Lee", "", "", "", "Current"], ["333", "Ed Gone", "Fay Gone", "", "", "", "Discharged"]]);
    let sawHistory = null;
    g.buildAllDerivedSheets = () => { sawHistory = here.getSheetByName("Student List").getRange(3, 7).getValue(); };
    g.onOpen = () => {};
    assert.strictEqual(g.ma_finish("old").success, true);
    assert.strictEqual(sawHistory, "Discharged", 'the build sees the old Student List (its Discharged students)');
    g.buildAllDerivedSheets = () => { throw new Error("Another build is running right now"); };
    assert.throws(() => g.ma_finish("old"), /Another build/);
    assert.strictEqual(here.getSheetByName("Student List").getLastRow(), 0, 'a failed build: the copy is taken out again');
  }
  console.log("Migration: newer tracker refused, settings, notes, types, Send Out, history: ok");
}

// ---------------------------------------------------------------- User Guide headings the document links open
{
  const g = loadProject({ noTrialFiles: true });
  const md = fs.readFileSync(path.join(ROOT, 'docs', 'USER_GUIDE.md'), 'utf8');
  const headings = md.split(/\r?\n/).filter(l => /^#{2,4} /.test(l)).map(l => l.replace(/^#+ /, '').trim());
  const H = g.__get('GUIDE_LINK_HEADINGS_'), L = g.__get('GUIDE_LINKS');
  const menus = fs.readFileSync(path.join(ROOT, 'Menus.gs'), 'utf8').toLowerCase();
  Object.keys(H).forEach(k => {
    assert(k in L, k + ' is a link in GUIDE_LINKS');
    assert.strictEqual(headings.filter(h => h === H[k]).length, 1, 'the guide has exactly one heading "' + H[k] + '"');
    assert(menus.includes('guide: ' + H[k].toLowerCase()), 'the Getting Started button says "' + H[k] + '"');
  });
  ['downloadDataDoc', 'buildDoc', 'settingsDoc', 'contactLogDoc'].forEach(k => assert(H[k], k + ' has a heading'));
  console.log("User Guide: a heading for every document link: ok");
}

// ---------------------------------------------------------------- Migration: renamed Contact Log headers
{
  const g = loadProject({ noTrialFiles: true });
  const low = a => a.map(h => h.toLowerCase());
  const oldH = low(["Week", "Date", "Guardian", "Person Spoke With", "OSIS", "Student", "#", "Method 1", "Method 2", "Notes", "Followup", "Followup Notes", "Hide"]);
  const newH = low(["Week", "Date", "Guardian", "Person Spoke With", "OSIS", "Student", "#", "Site", "Class", "Method 1", "Method 2", "Type", "Event Description", "Follow-up", "Additional Notes", "Hide"]);
  const m = g.migrationHeaderMap_(oldH, newH);
  assert.strictEqual(m[9], 12, 'old "Notes" -> Event Description');
  assert.strictEqual(m[11], 14, 'old "Followup Notes" -> Additional Notes');
  assert.strictEqual(m[10], 13, 'old "Followup" -> Follow-up');
  assert.strictEqual(m[6], 6, '# still matches itself');
  assert.strictEqual(m[1], 1, 'same headers as before');
  const m2 = g.migrationHeaderMap_(low(["Date", "Comments"]), newH);
  assert.strictEqual(m2[1], 12, '"Comments" -> Event Description');
  const m3 = g.migrationHeaderMap_(low(["Event Description", "Notes"]), newH);
  assert.strictEqual(m3[0], 12, 'an exact match wins');
  assert.strictEqual(m3[1], undefined, 'and the column is used once');
  console.log("Migration: renamed Contact Log headers: ok");
}

// ---------------------------------------------------------------- Migration: Person Spoke With formula
{
  const here = makeSpreadsheet(["Contact Log", "Backend"]);
  const old = makeSpreadsheet(["Contact Log"]); old.getId = () => "OLD";
  const app = fakeSpreadsheetApp(here);
  app.openByUrl = url => old;
  const g = loadProject({ extra: { SpreadsheetApp: app }, noTrialFiles: true });
  const hdr = ["Week", "Date", "Guardian", "Person Spoke With", "OSIS", "Student", "#", "Site", "Class", "Method 1", "Method 2", "Type", "Event Description", "Follow-up", "Additional Notes", "Hide"];
  const nl = here.getSheetByName("Contact Log");
  nl._put([hdr]); nl._set(10, 5, "END");
  nl.getRange(2, 18).setFormula(g.__get("SPOKE_WITH_CHOICES_FORMULA"));
  nl._set(2, 19, "spilled name"); // what the formula shows (Sheets keeps the formula in R2 only)
  const ol = old.getSheetByName("Contact Log");
  ol._put([["Week", "Date", "Guardian", "Person Spoke With", "OSIS", "Student", "#", "Method 1", "Method 2", "Notes", "Followup", "Followup Notes", "Hide"],
           ["", "10/1/2025", "Ann Lee", "Ann Lee", "111", "Bo Lee", 1, "Phone", "", "Talked about the bus", "No", "Call back", false],
           ["", "10/2/2025", "Cy Diaz", "Cy Diaz", "222", "Di Diaz", 1, "Email", "", "Trip form", "Yes", "", false]]);
  ol._set(4, 5, "END");
  g.ma_migrateContactLog("old");
  assert.strictEqual(nl.getRange(2, 18).getFormula(), g.__get("SPOKE_WITH_CHOICES_FORMULA"), 'R2 keeps the choices formula');
  assert.strictEqual(nl.getRange(2, 13).getValue(), "Talked about the bus", 'old Notes -> Event Description');
  assert.strictEqual(nl.getRange(2, 15).getValue(), "Call back", 'old Followup Notes -> Additional Notes');
  assert.strictEqual(nl.getRange(3, 14).getValue(), "Followup Needed", 'Follow-up value mapped');

  // A template without the formula gets it from the migration itself
  nl.getRange(2, 18).clearContent(); nl.getRange(2, 19).clearContent();
  g.ma_migrateContactLog("old");
  assert.strictEqual(nl.getRange(2, 18).getFormula(), g.__get("SPOKE_WITH_CHOICES_FORMULA"), 'missing formula put back');
  console.log("Migration: Person Spoke With formula kept: ok");
}

// ---------------------------------------------------------------- Migration: event tables
{
  const ss = makeSpreadsheet(["Events", "Backend"]);
  const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) }, noTrialFiles: true });
  const ev = ss.getSheetByName("Events");
  ev._put([["Fall PTC", "", "Parent Teacher Conference", "Gym", "10/01/2025", true, "", "Spring Fair", "", "Family event", "Yard", "04/01/2026", false],
           ["OSIS", "Guardian(s)", "Parent Attendee", "Student Name", "# of Indv", "Notes", "", "OSIS"],
           ["111111111", "Marge Simpson", "Marge", "Bart Simpson", 1, ""]], 1, 1);
  const calls = [];
  g.eventDropdownRules_ = () => ({ osis: "OSIS-RULE", guardian: "G-RULE", student: "S-RULE" });
  const realGetRange = ev.getRange;
  ev.getRange = function (a, b, c, d) {
    const r = realGetRange(a, b, c, d);
    return new Proxy(r, { get(t, k) {
      if (k === 'setDataValidation') return rule => { calls.push([b, rule]); return r; };
      if (k === 'insertCheckboxes') return () => { calls.push([b, 'checkbox']); return r; };
      return t[k];
    } });
  };
  assert.strictEqual(g.restyleEventTables_(ss, ev), 2, 'both tables found');
  assert(calls.some(x => x[0] === 6 && x[1] === 'checkbox') && calls.some(x => x[0] === 13 && x[1] === 'checkbox'), 'each table gets its delete box back');
  assert(calls.some(x => x[0] === 1 && x[1] === 'OSIS-RULE') && calls.some(x => x[0] === 9 && x[1] === 'G-RULE') && calls.some(x => x[0] === 11 && x[1] === 'S-RULE'), 'and its dropdowns');
  assert.strictEqual(ev.getRange(3, 1).getValue(), "111111111", 'values stay');
  console.log("Migration: event tables get their look back: ok");
}

// ---------------------------------------------------------------- Rollover: the archive's name
{
  const g = loadProject({ noTrialFiles: true });
  assert.strictEqual(g.rolloverArchiveName_(new Date(2027, 5, 20)), "PC_Tracker_Archive_2026-2027", 'June 2027 = the 2026-2027 year');
  assert.strictEqual(g.rolloverArchiveName_(new Date(2027, 7, 30)), "PC_Tracker_Archive_2026-2027", 'late August too');
  assert.strictEqual(g.rolloverArchiveName_(new Date(2026, 9, 1)), "PC_Tracker_Archive_2026-2027", 'October = the year that started');
  const src = fs.readFileSync(path.join(ROOT, 'Menus.gs'), 'utf8');
  const roll = src.slice(src.indexOf('function processRolloverBackend'));
  assert(/setProperty\('disableNightlySync', 'true'\)/.test(roll.slice(0, 3000)), 'the retired tracker stops rebuilding at night');
  console.log("Rollover: archive named for the school year that ends: ok");
}

// ---------------------------------------------------------------- dialogs
{
  const settings = fs.readFileSync(path.join(ROOT, 'SettingsDialog.html'), 'utf8');
  const load = settings.slice(settings.indexOf('// ---------- 1. Load the current settings'));
  assert(/setBusy\(true\)/.test(load.slice(0, 600)) && /withFailureHandler/.test(load.slice(0, 1200)), 'Settings: Save is off until the settings have loaded');
  const ns = fs.readFileSync(path.join(ROOT, 'NewStudentDialog.html'), 'utf8');
  assert(/replace\(\/\\D\/g, ''\)\.length < 6/.test(ns), 'New Student: an OSIS with fewer than 6 digits is caught');
  const mo = fs.readFileSync(path.join(ROOT, 'ManageOverridesDialog.html'), 'utf8');
  assert(/withFailureHandler[\s\S]*deleteOverride/.test(mo), 'Manage Overrides: a failed delete is shown');
  const ev = fs.readFileSync(path.join(ROOT, 'Edits.gs'), 'utf8');
  assert(/withFailureHandler[\s\S]{0,300}buildEventTable/.test(ev), 'New Event: a failure is shown');
  console.log("dialogs: no save before load, OSIS check, failures shown: ok");
}

// ---------------------------------------------------------------- Fake data: Site / Class in the Contact Log
// (owner 2026-11-13: the Example's Contact Log had no Site / Class; the typist that fills H:I never runs on
// rows the script writes, so the fake-data load fills them from the Directory after its build)
{
  const ss = makeSpreadsheet(["Contact Log", "Directory", "Backend"]);
  const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) }, noTrialFiles: true });
  const cl = ss.getSheetByName("Contact Log"), dir = ss.getSheetByName("Directory");
  dir._put([["Last Name", "First Name", "OSIS", "Site", "Class/ Teacher", "Parent Contact"],
    ["Simpson", "Bart", "123456789", "15", "Y41: Tessa", "Homer"],
    ["Belcher", "Gene", "967891234", "51", "Y27: Rosa", "Bob"]], 1, 1);
  cl._put([["", "Date", "Parent", "Person Spoke With", "OSIS", "Student", "#", "Site", "Class"],
    ["", "11/12/26", "", "Homer Simpson", "123456789", "Bart Simpson"],
    ["", "11/12/26", "", "Bob Belcher", "967891234", "Gene Belcher"],
    ["", "", "", "", "", ""],
    ["", "11/12/26", "", "Someone", "555555555", "Not in the Directory", "", "kept", "as is"]], 1, 1);
  cl._set(6, 5, "END");
  g.fillContactLogSiteClass_(ss);
  assert.deepStrictEqual(cl.getRange(2, 8, 4, 2).getValues(),
    [["15", "Y41: Tessa"], ["51", "Y27: Rosa"], ["", ""], ["kept", "as is"]], 'each row gets its student\'s Site / Class');
  assert.strictEqual(cl.getRange(6, 8).getValue(), "", 'the END row is not touched');
  const regen = fs.readFileSync(path.join(ROOT, 'Regenerate.gs'), 'utf8');
  const start = regen.indexOf('function loadFakeDataSteps_');
  const body = regen.slice(start, regen.indexOf('\nfunction ', start + 10));
  assert(body.indexOf('buildAllDerivedSheets();') !== -1 && body.indexOf('buildAllDerivedSheets();') < body.indexOf('fillContactLogSiteClass_(ss)'),
    'the fake-data load fills them after its build (the Directory exists then)');
  console.log("Fake data: Contact Log Site / Class from the Directory: ok");
}
