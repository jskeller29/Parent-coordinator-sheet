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

// ---------------------------------------------------------------- Contact Log: Person Spoke With choices (R:AA)
// ensureSpokeWithChoices_ (Edits.gs, end of every build): the formula in R2 lists each row's guardians
// for column D's dropdown ('Contact Log'!$R2:$AA2). Put back when missing, freed when blocked (#REF!),
// never written over a school's own data in R:AA, a formula changed by hand is kept, R1 stays empty.
{
  const fs = require('fs'), path = require('path');
  const { ROOT } = require('./lib/load');
  const fresh = (maxCols) => {
    const ss = makeSpreadsheet(["Contact Log", "Backend"]);
    const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) }, noTrialFiles: true });
    const log = ss.getSheetByName("Contact Log");
    if (maxCols) log._maxCols = maxCols;
    log._put([["Week", "Date", "Guardian", "Person Spoke With", "OSIS"],
              ["", "", "Morticia Addams and Gomez Addams", "", "111111111"],
              ["", "", "Marge Simpson", "", "222222222"],
              ["", "", "", "", ""],
              ["", "", "", "", "END"]], 1, 1); // END on row 5: log rows 2-4
    const calls = fn => ss._log.filter(x => x.tab === "Contact Log" && x.fn === fn);
    return { ss, g, log, calls };
  };

  // the formula and its columns
  {
    const { g } = fresh();
    const F = g.__get("SPOKE_WITH_CHOICES_FORMULA");
    assert.strictEqual(g.__get("CL_COL_SPOKE_CHOICES"), 18, 'starts in R');
    assert.strictEqual(g.__get("CL_COL_SPOKE_CHOICES") + g.__get("CL_SPOKE_CHOICES_WIDTH") - 1, 27, 'ends in AA');
    assert(F.indexOf('SPLIT(INDEX(C:C, 2):INDEX(C:C, MATCH(TRUE, TRIM(E:E)="END", 0) - 1), " and ", FALSE)') !== -1,
      'splits C on " and ", from row 2 to the row above END');
  }

  // first build: written in R2 + R:AA hidden; R1 left empty; the next build leaves it alone
  {
    const { ss, g, log, calls } = fresh();
    const F = g.__get("SPOKE_WITH_CHOICES_FORMULA");
    assert.strictEqual(g.ensureSpokeWithChoices_(ss), "added");
    assert.strictEqual(log.getRange(2, 18).getFormula(), F, 'formula in R2');
    assert.deepStrictEqual(calls("hideColumns").map(x => x.args), [[18, 10]], 'R:AA hidden');
    assert.strictEqual(log.getRange(1, 18).getFormula() + log.getRange(1, 18).getValue(), "", 'R1 stays empty');
    assert.strictEqual(g.ensureSpokeWithChoices_(ss), "kept");
    assert.strictEqual(calls("setFormula").length, 1, 'not written twice');

    // Migration clears rows 2+ in every column / row 2 deleted: the next build puts it back
    log.getRange(2, 1, 3, log.getMaxColumns()).clearContent(); // rows 2 .. END-1, like ma_ Migration
    log._put([["", "", "Morticia Addams and Gomez Addams", "", "111111111"]], 2, 1);
    assert.strictEqual(g.ensureSpokeWithChoices_(ss), "added");
    assert.strictEqual(log.getRange(2, 18).getFormula(), F, 'back after the wipe');
  }

  // blocked (#REF!): a pasted row left plain names in the list's cells -> cleared; nothing else is
  {
    const { ss, g, log } = fresh();
    const F = g.__get("SPOKE_WITH_CHOICES_FORMULA");
    log._set(2, 18, "#REF!"); log.getRange(2, 18).setFormula(F);
    log._set(2, 19, "Gomez Addams"); log._set(4, 18, "Marge Simpson");
    log._set(7, 18, "below END"); log._set(3, 3, "Marge Simpson");
    assert.strictEqual(g.ensureSpokeWithChoices_(ss), "unblocked");
    assert.strictEqual(log.getRange(2, 19).getValue() + log.getRange(4, 18).getValue(), "", 'blockers cleared');
    assert.strictEqual(log.getRange(2, 18).getFormula(), F, 'the formula stays');
    assert.strictEqual(log.getRange(7, 18).getValue(), "below END", 'nothing under END is touched');
    assert.strictEqual(log.getRange(3, 3).getValue(), "Marge Simpson", 'nor the log itself');
  }

  // a formula changed by hand (e.g. also split on " & ") is kept as it is
  {
    const { ss, g, log, calls } = fresh();
    const mine = '=ARRAYFORMULA(IFERROR(SPLIT(SUBSTITUTE(C2:C4, " & ", " and "), " and ", FALSE), ""))';
    log._set(2, 18, "Morticia Addams"); log.getRange(2, 18).setFormula(mine);
    assert.strictEqual(g.ensureSpokeWithChoices_(ss), "kept");
    assert.strictEqual(log.getRange(2, 18).getFormula(), mine);
    assert.strictEqual(calls("hideColumns").length, 0, 'columns not touched');
  }

  // a school's own data in R:AA: no formula over it, nothing hidden
  {
    const { ss, g, log, calls } = fresh();
    log._set(3, 20, "my own note");
    assert.strictEqual(g.ensureSpokeWithChoices_(ss), "skipped: R:AA hold data");
    assert.strictEqual(log.getRange(2, 18).getFormula(), "", 'no formula');
    assert.strictEqual(log.getRange(3, 20).getValue(), "my own note", 'their data stays');
    assert.strictEqual(calls("hideColumns").length, 0, 'their columns are not hidden');
  }

  // a tab narrower than W (A:P only, or A:T with data in R) / no END
  {
    const { ss, g, log } = fresh(16);
    assert.strictEqual(g.ensureSpokeWithChoices_(ss), "added");
    assert(log.getMaxColumns() >= 27, 'columns added up to AA first');
    const t = fresh(20);
    t.log._set(2, 18, "typed here");
    assert.strictEqual(t.g.ensureSpokeWithChoices_(t.ss), "skipped: R:AA hold data", 'the part of R:AA it has is checked too');
    const n = fresh();
    n.log._set(5, 5, "");
    assert.strictEqual(n.g.ensureSpokeWithChoices_(n.ss), "skipped: no log rows above END");
  }

  // every build runs it (after the lookup box rows, in a try/catch)
  const build = fs.readFileSync(path.join(ROOT, 'Build.gs'), 'utf8');
  const run = build.slice(build.indexOf('function buildAllDerivedSheetsRun_'));
  assert(/try \{ ensureSpokeWithChoices_\(ss\); \} catch/.test(run.slice(0, run.indexOf('\nfunction '))), 'called by the build');

  // A row delete (owner 2026-11-14): deleting row 2 takes the formula with it -> back right away,
  // not at the next build. Only R2 is looked at; any other delete changes nothing.
  {
    const { ss, g, log, calls } = fresh();
    const F = g.__get("SPOKE_WITH_CHOICES_FORMULA");
    const del = { changeType: "REMOVE_ROW", source: ss };
    g.ensureSpokeWithChoices_(ss);                       // set up as after a build
    const writes = () => calls("setFormula").length + calls("hideColumns").length;
    const before = writes();
    log.activate();
    g.installedOnChange(del);                            // row 5 deleted, say: R2 still has it
    assert.strictEqual(writes(), before, 'another row deleted: nothing written');
    log._set(2, 18, "");                                 // row 2 deleted: its formula is gone
    g.installedOnChange(del);
    assert.strictEqual(log.getRange(2, 18).getFormula(), F, 'row 2 deleted: the formula is back at once');
    log._set(2, 18, "");
    ss.getSheetByName("Backend").activate();             // a row deleted on another tab
    g.installedOnChange(del);
    assert.strictEqual(log.getRange(2, 18).getFormula(), "", 'another tab: the Contact Log is not touched');
    log.activate();
    g.installedOnChange({ changeType: "INSERT_ROW", source: ss });
    assert.strictEqual(log.getRange(2, 18).getFormula(), "", 'not a delete: nothing');
    log._set(3, 20, "my own note");                      // something typed into R:AA
    g.installedOnChange(del);
    assert.strictEqual(log.getRange(2, 18).getFormula(), "", 'never written over what is typed in R:AA');
    assert.strictEqual(log.getRange(3, 20).getValue(), "my own note");
    assert.strictEqual(g.restoreSpokeWithChoicesAfterDelete_(ss, log), "skipped: R:AA hold data");
  }
  console.log("Contact Log: Person Spoke With choices: ok");
}
