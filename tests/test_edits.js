// Contact Log / Notes / Events edit handling on the fake spreadsheet (lib/fakesheet.js):
//  - a multi-column paste whose rows carry their own OSIS keeps it (siblings share a parent's name)
//  - a Notes row whose OSIS is cleared / typed over takes its note away from that student
//  - a Notes delete that can't ask "are you sure?" deletes nothing
const assert = require('assert');
const { loadProject } = require('./lib/load');
const { makeSpreadsheet, fakeSpreadsheetApp } = require('./lib/fakesheet');

const TABS = ["Contact Log", "Combined Contact Tracking", "Directory", "Student List", "Notes", "Backend"];
function sheetWithSiblings() {
  const ss = makeSpreadsheet(TABS);
  // Student List (header row 1): two children of the same parent, plus another family
  ss.getSheetByName("Student List")._put([
    ["OSIS", "Student", "Guardian", "Site", "Class", "Sibling", "Status"],
    ["111111111", "Bart Simpson", "Marge Simpson"],
    ["222222222", "Lisa Simpson", "Marge Simpson"],
    ["333333333", "Ned Flanders", "Maude Flanders"]
  ], 1, 1);
  const app = fakeSpreadsheetApp(ss);
  const g = loadProject({ extra: { SpreadsheetApp: app }, noTrialFiles: true });
  return { ss, g, app };
}

// ---------------------------------------------------------------- Contact Log: paste with OSIS
{
  const { ss, g } = sheetWithSiblings();
  const log = ss.getSheetByName("Contact Log");
  // Rows copied from another log: C Guardian, D Person, E OSIS, F Student (Lisa's rows carry Lisa's OSIS)
  log._put([["Marge Simpson", "Mom", "222222222", "Lisa Simpson"], ["Marge Simpson", "Mom", "111111111", "Bart Simpson"]], 2, 3);
  g.fillPastedContactLogRows_(ss, log, 2, 2, 3, 10, 4); // pasted C2:F3
  assert.strictEqual(String(log.getRange(2, 5).getValue()), "222222222", 'Lisa keeps her own OSIS (was overwritten with Bart\'s)');
  assert.strictEqual(log.getRange(2, 6).getValue(), "Lisa Simpson");
  assert.strictEqual(String(log.getRange(3, 5).getValue()), "111111111");
  // A paste of the Guardian column only: looked up by the name, as before
  log._put([["Maude Flanders"]], 5, 3);
  g.fillPastedContactLogRows_(ss, log, 5, 1, 3, 10, 1);
  assert.strictEqual(String(log.getRange(5, 5).getValue()), "333333333", 'name-only paste still fills the OSIS');
  console.log("Contact Log paste keeps each row's own OSIS: ok");
}

// ---------------------------------------------------------------- Notes: OSIS cleared / typed over
{
  const { ss, g } = sheetWithSiblings();
  const notes = ss.getSheetByName("Notes");
  notes._put([["OSIS", "Guardian", "Student", "Notes"], ["", "Marge Simpson", "Bart Simpson", "Call back Tuesday"]], 1, 1);
  const synced = [];
  g.updateOsisDeltas_ = list => { synced.push(list.slice()); };
  // The user clears A2, which held Bart's OSIS
  g.watchNotesSheet_({ source: ss, range: notes.getRange(2, 1), value: undefined, oldValue: "111111111" });
  assert(synced.some(l => l.indexOf("111111111") !== -1), 'Bart loses the note when the OSIS is cleared: ' + JSON.stringify(synced));
  console.log("Notes: a cleared OSIS takes the note away: ok");
}

// ---------------------------------------------------------------- Notes: delete without a yes
{
  const { ss, g, app } = sheetWithSiblings();
  const notes = ss.getSheetByName("Notes");
  notes._put([["OSIS", "Guardian", "Student", "Notes", "Delete"], ["111111111", "Marge Simpson", "Bart Simpson", "Keep me", true]], 1, 1);
  let deleted = 0;
  notes.deleteRow = () => { deleted++; };
  app._ui.alert = () => { throw new Error("no UI here"); };
  g.watchNotesSheet_({ source: ss, range: notes.getRange(2, 5), value: "TRUE" });
  assert.strictEqual(deleted, 0, 'no question shown = nothing deleted');
  console.log("Notes: no delete without a yes: ok");
}
