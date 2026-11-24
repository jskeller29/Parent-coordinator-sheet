// Events: one event at a time + 👁️ View Events (owner 2026-11-24)
//  - 📋 New Event shows only the new table: every other table's 7 columns are hidden
//  - eventTablesList_ lists every table (name, type, place, date, rows filled in, shown)
//  - showEventTables (View Events' Show button) shows the ticked ones and hides the rest,
//    one call per run of tables; finds a table by name when tables moved since the window opened
//  - deleting the only table showing shows the newest one left
//  - every table's Notes column wraps its text
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

console.log("Events: one event at a time, 👁️ View Events: ok");
