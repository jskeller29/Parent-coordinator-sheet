// Fixes from 2026-11-12 that are plain logic (no sheet needed).
const assert = require('assert');
const { loadProject } = require('./lib/load');
let triggers = [];
const g = loadProject({ extra: { ScriptApp: { getProjectTriggers: () => triggers, deleteTrigger: t => { triggers = triggers.filter(x => x !== t); } } } });
const P = g.__props;

// Notes date range (Settings) applies to the quick update too
const contact = [["h"], ["", "11/01/2026", "", "Marge", "123456789", "Bart", "", "", "", "Phone", "", "", "Old call", "", "", false],
                        ["", "11/10/2026", "", "Marge", "123456789", "Bart", "", "", "", "Phone", "", "", "New call", "", "", false]];
let t = g.notesTextForOsis_("123456789", contact, null, null, false);
assert(t.includes("Old call") && t.includes("New call"));
P.notesStartDate = "2026-11-05"; g.NOTES_DATE_RANGE_ = null;
t = g.notesTextForOsis_("123456789", contact, null, null, false);
assert(!t.includes("Old call") && t.includes("New call"));
delete P.notesStartDate; g.NOTES_DATE_RANGE_ = null;
console.log("notes date range: ok");

// Sheet Lock: ranges around the lookup box cells you type in
const r = g.rangesAround_(11, 1, 40, 20, [[13, 5], [12, 6]]);
const cells = new Set();
r.forEach(([r1, c1, r2, c2]) => { for (let a = r1; a <= r2; a++) for (let b = c1; b <= c2; b++) { const k = a + ":" + b; assert(!cells.has(k)); cells.add(k); } });
assert(cells.size === 30 * 20 - 2 && !cells.has("13:5") && !cells.has("12:6"));
console.log("sheet lock ranges: ok");

// Follow-up values: Migration mapping, the Contact Log tool's "still open" test
assert.strictEqual(g.migrateFollowupValue_("Yes"), "Followup Needed");
assert.strictEqual(g.migrateFollowupValue_(" yes-Important "), "Followup Needed - ASAP");
assert.strictEqual(g.migrateFollowupValue_("No"), "No Followup Needed");
assert.strictEqual(g.migrateFollowupValue_("Resolved"), "Resolved");
assert(g.cl_isOpenFollowup_("Followup Needed - ASAP") && !g.cl_isOpenFollowup_("No Followup Needed"));
console.log("follow-up values: ok");

// One nightly job per sheet
const mk = id => ({ getUniqueId: () => id, getHandlerFunction: () => "nightlyMasterSync" });
triggers = [mk("A")];
assert(g.nightlyTriggerShouldRun_({ triggerUid: "A" }) === true && P.NIGHTLY_TRIGGER_UID === "A");
triggers = [mk("B")];
assert(g.nightlyTriggerShouldRun_({ triggerUid: "B" }) === false && triggers.length === 0, "a 2nd person's copy deletes itself");
P.NIGHTLY_LAST_RUN = String(Date.now() - 5 * 86400000); P.NIGHTLY_TRIGGER_UID_SET = String(Date.now() - 10 * 86400000);
triggers = [mk("C")];
assert(g.nightlyTriggerShouldRun_({ triggerUid: "C" }) === true && P.NIGHTLY_TRIGGER_UID === "C", "takes over after 3 days");
P.disableNightlySync = "true";
assert(g.nightlyTriggerShouldRun_({ triggerUid: "C" }) === false, "off in Settings");
delete P.disableNightlySync;
console.log("nightly de-dupe: ok");

// Last night's rebuild failed -> a message for the next open; a finished build clears it
g.withBuildLock_ = () => { const e = new Error("Another \"Build\" is running right now. Try again in a few minutes."); e.buildBusy = true; throw e; };
g.nightlyMasterSyncRun_();
assert(/didn't run: Another/.test(P.NIGHTLY_FAILED_MSG));
g.withBuildLock_ = () => { throw new Error("Sheet \"Raw Data\" not found."); };
g.nightlyMasterSyncRun_();
assert(/didn't finish: Sheet "Raw Data" not found/.test(P.NIGHTLY_FAILED_MSG));
console.log("nightly failure message: ok");

// Build lock: re-entrant, and "busy" says what is running
const g2 = loadProject();
let inner = 0;
g2.withBuildLock_("outer", () => g2.withBuildLock_("inner", () => { inner++; }));
assert(inner === 1 && g2.BUILD_LOCK_DEPTH_ === 0);
g2.LockService = { getScriptLock: () => ({ tryLock: () => false, releaseLock() {} }) };
assert.throws(() => g2.withBuildLock_("x", () => {}), e => e.buildBusy && /Try again/.test(e.message));
console.log("build lock: ok");

// Quick rebuild keeps a header row (the first student keeps their details)
g2.studentLookupRows_ = () => [["111", "Ann", "G", "S"], ["222", "Ben", "H", "S"]];
const m = g2.getExistingMasterTable_({});
assert(m[0][0] === "OSIS" && m.length === 3);
console.log("quick-rebuild header: ok");

// Ghost typists ignore stray spaces
assert.strictEqual(g2.typistKey_("  Ann   Smith "), g2.typistKey_("ann smith"));
g2.studentLookupRows_ = () => [["111", "Ann  Smith", "Mary Smith "], ["222", "Ben Lee", "Joe Lee"]];
g2.studentLookupColumn_ = () => ({ createTextFinder: () => ({ matchEntireCell: () => ({ matchCase: () => ({ findNext: () => null }) }) }) });
assert.strictEqual(g2.findInMasterTable_({}, 2, "ann smith").osis, "111");
assert.strictEqual(g2.findInMasterTable_({}, 3, "Mary  Smith").osis, "111");
assert.strictEqual(g2.findInMasterTable_({}, 2, "Nobody"), null);
console.log("typist stray spaces: ok");

// Every Follow-up value the code writes is one of the real dropdown options
// (the owner's Contact Log Table, 2026-11-12; exact text, case included)
const fs = require('fs'), path = require('path');
const { ROOT } = require('./lib/load');
const FOLLOWUP_OPTIONS = ["No Followup Needed", "Followup Needed", "Followup Needed - ASAP",
  "Followed up / Resolved", "Skipped / Newer Interaction", "Waiting for Response"];
const isOption = v => FOLLOWUP_OPTIONS.indexOf(v) !== -1;
const written = [g.CLQ_FOLLOWUP_NEEDED, g.CLQ_NO_FOLLOWUP, g.CLQ_RESOLVED].concat(Object.values(g.__get("MIGRATION_FOLLOWUP_MAP_")));
const edits = fs.readFileSync(path.join(ROOT, "Edits.gs"), "utf8");
const autofill = edits.match(/getRange\(row, 14\)\.setValue\("([^"]+)"\)/);
assert(autofill, "auto-fill default for N not found");
written.push(autofill[1]);
// Fake Contact Log rows (Regenerate.gs): [days, guardian, spoke with, OSIS, student, M1, M2, type, notes, FOLLOW-UP, ...]
const regen = fs.readFileSync(path.join(ROOT, "Regenerate.gs"), "utf8");
const fakeRows = regen.split("\n").filter(l => /^\s*\[\d+,"/.test(l)).map(l => JSON.parse(l.trim().replace(/,\s*$/, "")));
assert(fakeRows.length >= 20, "fake Contact Log rows not found");
fakeRows.forEach(r => written.push(r[9]));
const bad = written.filter(v => !isOption(v));
assert.deepStrictEqual(bad, [], "Follow-up values that are not dropdown options: " + JSON.stringify(bad));
console.log("follow-up values are real options: ok (" + written.length + " checked)");

// z_NewCopy.gs: the pop-up on a new copy of the owner's sheet
const gc = loadProject({ owner: true, extra: { SpreadsheetApp: { getActiveSpreadsheet: () => ({ getId: () => gc.__sheetId }) } } });
const alerts = [];
const fakeUi = answer => ({ ButtonSet: { YES_NO: "YN", OK: "OK" }, Button: { YES: "YES", NO: "NO" },
  alert: (title, text, buttons) => { alerts.push(title); return buttons === "YN" ? answer : "OK"; } });
gc.__sheetId = "MAIN";
gc.ownerNewCopyCheck_(fakeUi("YES"));                       // first open of the main sheet: asked, answered Yes
assert.deepStrictEqual(alerts, ["📋 New copy?"]);
gc.ownerNewCopyCheck_(fakeUi("YES"));                       // never asked again there
assert.strictEqual(alerts.length, 1);
gc.__props.OWNER_HOME_SHEET_ID = undefined; delete gc.__props.OWNER_HOME_SHEET_ID; // a copy: Google didn't copy the settings
gc.__sheetId = "COPY1";
gc.ownerNewCopyCheck_(fakeUi("NO"));                        // asked, answered No -> the checklist
assert.deepStrictEqual(alerts.slice(1), ["📋 New copy?", "📋 New copy: what to change"]);
gc.ownerNewCopyCheck_(fakeUi("NO"));                        // shown once per copy
assert.strictEqual(alerts.length, 3);
assert(/z_MINE/.test(gc.newCopyChecklistText_()) && /App Scripts \(General\)/.test(gc.newCopyChecklistText_()));
console.log("new copy checklist pop-up: ok");
