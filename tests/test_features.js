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

// ---------------------------------------------------------------- Migration: the old tracker's nightly
{
  const here = makeSpreadsheet(["Contact Log", "Backend"]);
  const old = makeSpreadsheet(["Contact Log"]); old.getId = () => "OLD"; old.getName = () => "PC Tracker 2025";
  const app = fakeSpreadsheetApp(here);
  app.openByUrl = url => url === "old" ? old : (() => { throw new Error("no access"); })();
  const g = loadProject({ extra: { SpreadsheetApp: app }, noTrialFiles: true });
  const r = g.ma_markOldTracker("old");
  assert(r.marked && r.name === "PC Tracker 2025" && /OLD/.test(r.url), 'the old tracker is marked and named');
  g.ma_markOldTracker("old");
  assert.strictEqual(old._meta.length, 1, 'marked once only');
  assert.strictEqual(g.ma_markOldTracker("nope").marked, false, 'no access: not marked, no error');

  // The OLD tracker's own nightly trigger (newer code) sees the mark and turns itself off
  const deleted = [];
  const trig = (uid, fn) => ({ getUniqueId: () => uid, getHandlerFunction: () => fn });
  const triggers = [trig("T1", "nightlyMasterSync"), trig("T2", "installedOnEdit")];
  const o = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(old),
    ScriptApp: { getProjectTriggers: () => triggers, deleteTrigger: t => deleted.push(t.getUniqueId()) } }, noTrialFiles: true });
  o.__props.syncToPhones = 'true';
  assert.strictEqual(o.nightlyTriggerShouldRun_({ triggerUid: "T1" }), false, 'the replaced tracker does not rebuild');
  assert.deepStrictEqual(deleted, ["T1"], 'only its nightly trigger is deleted');
  assert.strictEqual(o.__props.disableNightlySync, 'true');
  assert.strictEqual(o.__props.syncToPhones, 'false');

  // A Drive copy of the marked tracker (other ID) keeps its nightly
  const copy = makeSpreadsheet(["Contact Log"]); copy.getId = () => "COPY"; copy._meta = old._meta.slice();
  const c = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(copy) }, noTrialFiles: true });
  assert.strictEqual(c.nightlyTriggerShouldRun_({ triggerUid: "X" }), true, 'a copy of it still runs');
  // The new tracker itself is not affected
  assert.strictEqual(g.nightlyTriggerShouldRun_({ triggerUid: "N" }), true, 'the new tracker runs');
  console.log("Migration: the old tracker's nightly turns off: ok");
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
