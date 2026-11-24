// Events: one event at a time + 👁️ View Events (owner 2026-11-24)
//  - 📋 New Event shows only the new table: every other table's 7 columns are hidden
//  - eventTablesList_ lists every table (name, type, place, date, rows filled in, shown)
//  - showEventTables (View Events' Show button) shows the ticked ones and hides the rest,
//    one call per run of tables; finds a table by name when tables moved since the window opened
//  - deleting the only table showing shows the newest one left
//  - every table's Notes column wraps its text
//  - an Events edit only fills the row in (no Backend Event Log / notes until a build); duplicates still go;
//    an empty table can be deleted; the Events panel hides the tab and opens the build
//  - Settings: Raw Data / Events / the Override Report not in it, New Students only on the owner's sheet
//  - the Students Override Report is deleted by the nightly run (or a build when the nightly is off)
//  - the View Events window gets the list as JSON that can't close its <script>
const assert = require('assert');
const { loadProject } = require('./lib/load');
const { makeSpreadsheet, fakeSpreadsheetApp } = require('./lib/fakesheet');

const plain = x => JSON.parse(JSON.stringify(x));
let dialogHtml = "";
const ss = makeSpreadsheet(["Events", "Backend"]);
const app = fakeSpreadsheetApp(ss);
const g = loadProject({
  extra: {
    SpreadsheetApp: app,
    HtmlService: { createHtmlOutput: h => { dialogHtml = h; const o = { setWidth: () => o, setHeight: () => o }; return o; } }
  },
  noTrialFiles: true
});
g.include_ = () => "";
g.eventDropdownRules_ = () => ({ osis: "OSIS-RULE", guardian: "GUARDIAN-RULE", student: "STUDENT-RULE" });
g.applySheetLocks_ = () => {};
g.backendReadBlock_ = () => [["Date"]];
const ev = ss.getSheetByName("Events");
const hidden = () => Array.from(ev._hiddenCols).sort((a, b) => a - b);
const range = (a, n) => Array.from({ length: n }, (_, i) => a + i);

// ---------------------------------------------------------------- New Event: only the new one shows
g.buildEventTable("Fall PTC", "Gym", "10/01/2026", "Parent Teacher Conference");
assert.deepStrictEqual(hidden(), [], "the first event: nothing to hide");
ev.getRange(3, 1).setValue("111111111");
ev.getRange(4, 3).setValue("Marge Simpson");

g.buildEventTable("Graduation Workshop", "Library", "11/05/2026", "Workshop");
assert.strictEqual(ev.getRange(1, 8).getValue(), "Fall PTC", "the older table moved right");
assert.deepStrictEqual(hidden(), range(8, 7), "New Event: the older table is hidden");

g.buildEventTable("Spring Festival", "Yard", "11/20/2026", "Family event");
assert.deepStrictEqual(hidden(), range(8, 14), "New Event: both older tables hidden");
assert(!ev.isColumnHiddenByUser(1), "the new table shows");
const hides = ss._log.filter(c => c.tab === "Events" && c.fn === "hideColumns").slice(-1)[0];
assert.deepStrictEqual(plain(hides.args), [8, 14], "two tables side by side = one hideColumns call");

// Notes column wraps (header row down), on every table
const wraps = ss._log.filter(c => c.tab === "Events" && c.fn === "setWrap");
[6, 13, 20].forEach(col => assert(wraps.some(c => c.col === col && c.row === 2 && c.nc === 1 && c.args[0] === true
  && c.nr === ev.getMaxRows() - 1), "Notes column " + col + " wraps"));
assert(!wraps.some(c => c.col % 7 !== 6), "only the Notes columns wrap");

// ---------------------------------------------------------------- the list
const list = plain(g.eventTablesList_(ev));
assert.deepStrictEqual(list.map(t => [t.col, t.name, t.type, t.place, t.date, t.shown]), [
  [1, "Spring Festival", "Family event", "Yard", "11/20/2026", true],
  [8, "Graduation Workshop", "Workshop", "Library", "11/05/2026", false],
  [15, "Fall PTC", "Parent Teacher Conference", "Gym", "10/01/2026", false]
], "every table, newest first, with what shows");
assert.deepStrictEqual(list.map(t => t.rows), [0, 0, 2], "rows filled in (OSIS or a name)");

// ---------------------------------------------------------------- View Events: show some / all
let r = plain(g.showEventTables([{ col: 15, name: "Fall PTC" }, { col: 8, name: "Graduation Workshop" }]));
assert.deepStrictEqual(r, { shown: 2, total: 3 }, "two shown");
assert.deepStrictEqual(hidden(), range(1, 7), "the two ticked show, the new one is hidden");
assert(ss._toasts.some(t => /Showing 2 of 3 events/.test(t.msg)), "toast says how many");

r = plain(g.showEventTables(list.map(t => ({ col: t.col, name: t.name }))));
assert.deepStrictEqual(hidden(), [], "Show all: every table shows");
assert(ss._toasts.some(t => /Showing every event/.test(t.msg)), "toast for all");

// a table moved since the window opened (a new event was added): found by its name
g.buildEventTable("Book Fair", "Hall", "11/22/2026", "Family event");
r = plain(g.showEventTables([{ col: 15, name: "Fall PTC" }]));
assert.strictEqual(r.shown, 1, "found by name");
assert.deepStrictEqual(hidden(), range(1, 21), "Fall PTC (now at column 22) is the only one showing");

assert.throws(() => g.showEventTables([{ col: 1, name: "Gone" }]), /no longer on the Events tab/, "an event that's gone");
assert.deepStrictEqual(hidden(), range(1, 21), "nothing changed after a refusal");

// ---------------------------------------------------------------- deleting the only one showing
g.buildBackendEventLogData_ = () => [["Date"]];
g.backendWriteBlock_ = () => {};
ss._active = ev;
const box = ev.getRange(1, 27);
g.watchEventsSheet_({ range: box, value: "TRUE", source: ss });
assert.deepStrictEqual(plain(g.eventTablesList_(ev).map(t => [t.name, t.shown])),
  [["Book Fair", true], ["Spring Festival", false], ["Graduation Workshop", false]], "deleted; the newest left shows");
assert(app._ui._alerts.some(a => /Showing "Book Fair" now/.test(a)), "the alert says which one shows");

// ---------------------------------------------------------------- the window
g.openViewEventsDialog();
const m = /const EVENTS = (.*);\n/.exec(dialogHtml);
assert(m, "the list is in the window");
assert.strictEqual(JSON.parse(m[1]).length, 3, "3 events listed");
ev.getRange(1, 1).setValue("</script><b>x");
g.openViewEventsDialog();
assert(!/<\/script><b>/.test(/const EVENTS = (.*);\n/.exec(dialogHtml)[1]), "a name can't close the script tag");
assert.strictEqual(JSON.parse(/const EVENTS = (.*);\n/.exec(dialogHtml)[1])[0].name, "</script><b>x", "the name itself is kept");

// ---------------------------------------------------------------- fast edits: the typist only
{
  let pushed = 0;
  ["backendWriteBlock_", "backendAppendRows_", "backendDeleteRows_", "backendSortBlock_", "updateOsisDeltas_", "buildBackendEventLogData_", "clearNotesSourcesCache_"]
    .forEach(fn => { g[fn] = () => { pushed++; return [["Date"]]; }; });
  g.studentLookupSource_ = () => ({});
  g.findInMasterTable_ = (src, col, v) => (String(v) === "111111111" ? { osis: "111111111", student: "Bart Simpson", guardian: "Marge Simpson" } : null);
  ss._active = ev;
  const startCol = g.eventTablesList_(ev).find(t => t.shown).col;
  ev.getRange(3, startCol).setValue("111111111");
  g.installedOnEdit({ range: ev.getRange(3, startCol), value: "111111111", source: ss });
  assert.deepStrictEqual(ev.getRange(3, startCol, 1, 5).getValues()[0], ["111111111", "Marge Simpson", "", "Bart Simpson", 1], "the row fills in");
  assert.strictEqual(pushed, 0, "an Events edit writes nothing to the Backend Event Log / notes (the build does)");

  const alerts = app._ui._alerts.length;
  ev.getRange(4, startCol).setValue("111111111");
  g.watchEventsSheet_({ range: ev.getRange(4, startCol), value: "111111111", source: ss });
  assert.strictEqual(ev.getRange(4, startCol).getValue(), "", "the same family twice: the later row is cleared");
  assert.strictEqual(ev.getRange(3, startCol).getValue(), "111111111", "the first one stays");
  assert(app._ui._alerts.slice(alerts).some(a => /Duplicate Removed/.test(a)), "and the user is told");

  // a new name / date in row 1: nothing now (the next build reads it)
  g.watchEventsSheet_({ range: ev.getRange(1, startCol), value: "Book Fair 2", oldValue: "Book Fair", source: ss });
  assert.strictEqual(pushed, 0, "a rename pushes nothing");
}

// ---------------------------------------------------------------- an empty table can be deleted
{
  const ss2 = makeSpreadsheet(["Contact Log", "Events", "Backend"]);
  const g2 = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss2) }, noTrialFiles: true });
  g2.eventDropdownRules_ = () => ({}); g2.applySheetLocks_ = () => {}; g2.backendReadBlock_ = () => [["Date"]];
  const ev2 = ss2.getSheetByName("Events");
  ev2.hideSheet();
  g2.buildEventTable("Empty", "Gym", "11/24/2026", "Workshop");
  assert(!ev2.isSheetHidden(), "New Event shows the hidden Events tab");
  assert.strictEqual(ss2.getActiveSheet(), ev2, "and goes to it");
  g2.watchEventsSheet_({ range: ev2.getRange(1, 6), value: "TRUE", source: ss2 });
  assert.strictEqual(g2.eventTablesList_(ev2).length, 0, "a table with no rows yet can be deleted (it couldn't before)");

  // the Events panel
  g2.buildEventTable("Fall PTC", "Gym", "11/24/2026", "PTC");
  const panel = plain(g2.getEventsPanel());
  assert.deepStrictEqual([panel.showing.map(t => t.name), panel.total, panel.tabShown, panel.nightlyOff], [["Fall PTC"], 1, true, false], "panel data");
  let built = 0;
  g2.runBuildOnly = () => { built++; };
  assert.deepStrictEqual(plain(g2.eventsFinish(false)), { building: false }, "close, build later");
  assert(ev2.isSheetHidden(), "the Events tab is hidden");
  assert.strictEqual(ss2.getActiveSheet().getName(), "Contact Log", "the user lands on the Contact Log");
  assert(ss2._toasts.some(t => /tonight's rebuild/.test(t.msg)), "says when the notes get it");
  g2.buildEventTable("Spring", "Yard", "11/25/2026", "Family event");
  g2.eventsFinish(true);
  assert(ev2.isSheetHidden() && built === 1, "Done: hides the tab and opens the build");
}

// ---------------------------------------------------------------- Settings: tabs its tools handle are not in it
{
  const ss3 = makeSpreadsheet(["Contact Log", "Raw Data", "Events", "Students Override Report", "New Students", "Notes", "Backend"]);
  const g3 = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss3) }, noTrialFiles: true });
  g3.__props.showRawData = "true"; g3.__props.showEvents = "true"; g3.__props.showNewStudents = "true";
  let st = g3.getSettings();
  ["showRawData", "showEvents", "showOverrideReport", "showNewStudents"].forEach(k => assert.strictEqual(st.tabsPresent[k], false, k + ": no row in Settings on a school's sheet"));
  assert.strictEqual(st.tabsPresent.notesTab, true, "Notes still in Settings");
  assert.strictEqual(st.showRawData, false, "an old saved choice for Raw Data is ignored");
  g3.getOwnerHooks_ = () => ({});
  st = g3.getSettings();
  assert.strictEqual(st.tabsPresent.showNewStudents, true, "New Students: in Settings on the owner's sheet");
  assert.strictEqual(st.tabsPresent.showEvents, false, "Events: not even there");

  const raw = ss3.getSheetByName("Raw Data"), events = ss3.getSheetByName("Events"), notes = ss3.getSheetByName("Notes");
  events.hideSheet();
  g3.updateVisualSettings_({ notesTab: false });
  assert(!raw.isSheetHidden(), "a Settings save leaves Raw Data as it is");
  assert(events.isSheetHidden(), "and Events");
  assert(notes.isSheetHidden(), "a Settings tab still follows the setting");
}

// ---------------------------------------------------------------- the Override Report goes at night
{
  const mk = () => {
    const s4 = makeSpreadsheet(["Contact Log", "Students Override Report", "Backend"]);
    const g4 = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(s4) }, noTrialFiles: true });
    s4._active = s4.getSheetByName("Students Override Report");
    return { s4, g4 };
  };
  let { s4, g4 } = mk();
  assert.strictEqual(g4.deleteOverrideReport_(s4), true, "deleted");
  assert(!s4.getSheetByName("Students Override Report"), "the tab is gone");
  assert.strictEqual(s4.getActiveSheet().getName(), "Contact Log", "someone on it lands on the Contact Log");
  assert.strictEqual(g4.deleteOverrideReport_(s4), false, "no report: nothing to do");

  // the nightly run deletes it
  ({ s4, g4 } = mk());
  g4.studentListIsBuilt_ = () => true; g4.buildAllDerivedSheets = () => {}; g4.resetContactLogFormatting = () => {}; g4.refreshResolvedLinks_ = () => {};
  g4.nightlyMasterSyncLocked_();
  assert(!s4.getSheetByName("Students Override Report"), "the 10 PM run deletes it");

  // a build: only when the nightly rebuild is off
  const src = require('fs').readFileSync(require('path').join(__dirname, '..', 'Build.gs'), 'utf8');
  assert(/getProperty\('disableNightlySync'\) === 'true'\) deleteOverrideReport_\(ss\);\n\n  \/\/ "This tracker pauses itself/.test(src),
    "the end of every build deletes it when the nightly rebuild is off");
}

console.log("Events: one event at a time, 👁️ View Events, fast edits + the Events panel; Settings tabs; the report at night: ok");
