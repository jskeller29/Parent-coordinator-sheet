// The owner's 10 ideas of 2026-11-21 (built 2026-11-23; the "PC Tracker Ideas: Thoughts and Mockups" doc)
//  9  Contact Log: Method (J) left blank on a row with a Date is light red
const fs = require('fs'), path = require('path'), assert = require('assert');
const { loadProject, ROOT } = require('./lib/load');
const { makeSpreadsheet, fakeSpreadsheetApp } = require('./lib/fakesheet');
// Values from the project's context have that context's Array / Object: compare them as JSON
const eq = (a, b, m) => assert.deepStrictEqual(JSON.parse(JSON.stringify(a === undefined ? null : a)), JSON.parse(JSON.stringify(b === undefined ? null : b)), m);

// ---------------------------------------------------------------- 9. red Method cell
{
  const ss = makeSpreadsheet(["Contact Log", "Backend"]);
  const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) } });
  const cl = ss.getSheetByName("Contact Log");
  const L = g.__get("SHEET_LOOK_");
  const isMethodRule = r => r._st.formula === '=AND($B2<>"", $J2="")';

  // The 10 PM redraw (and 🎨 Apply Sheet Style) draws it with the others
  g.resetContactLogFormatting();
  let rules = cl._cf.filter(isMethodRule);
  assert.strictEqual(rules.length, 1, "the full redraw has the Method rule once");
  assert.strictEqual(rules[0]._st.bg, L.missingMethod, "light red");
  assert.strictEqual(L.missingMethod, "#f4cccc");
  const rg = rules[0]._st.ranges[0];
  eq([rg.getRow(), rg.getColumn(), rg.getNumColumns(), rg.getLastRow()], [2, 10, 1, cl.getMaxRows()], "on J2:J (to the bottom)");
  assert(cl._cf.some(r => r._st.formula === '=AND($B2<>"", $G2="")'), "the dark red # rule is still there");
  assert(cl._cf.some(r => r._st.formula === '=$J2="Workshop"'), "the Workshop rule is still there");

  // A build adds it to a sheet that doesn't have it (nightly run off), once, and keeps the other rules
  cl._cf = cl._cf.filter(r => !isMethodRule(r));
  const others = cl._cf.length, writes = cl._cfWrites;
  assert.strictEqual(g.ensureMethodMissingRule_(ss), true, "added when missing");
  assert.strictEqual(cl._cf.filter(isMethodRule).length, 1, "now there");
  assert.strictEqual(cl._cf.length, others + 1, "no other rule lost");
  assert.strictEqual(g.ensureMethodMissingRule_(ss), false, "the next build: nothing to do");
  assert.strictEqual(cl._cfWrites, writes + 1, "one write in all");
  // Recolored by hand (same formula): left alone
  cl._cf.filter(isMethodRule)[0]._st.bg = "#ff0000";
  assert.strictEqual(g.ensureMethodMissingRule_(ss), false, "a recolored rule is kept");
  assert.strictEqual(cl._cf.filter(isMethodRule)[0]._st.bg, "#ff0000");
  // No Contact Log: nothing happens
  const g2 = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(makeSpreadsheet(["Backend"])) } });
  assert.strictEqual(g2.ensureMethodMissingRule_(g2.SpreadsheetApp.getActiveSpreadsheet()), false);
  // The build calls it (in its own try / catch)
  const build = fs.readFileSync(path.join(ROOT, 'Build.gs'), 'utf8');
  assert(/try \{ ensureMethodMissingRule_\(ss\); \} catch/.test(build), "every build checks it");
  console.log("9 red Method cell: ok");
}

// ---------------------------------------------------------------- 10. Method 2 choices (Backend CQ)
{
  const ss = makeSpreadsheet(["Contact Log", "Backend"]);
  const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) } });
  const be = ss.getSheetByName("Backend");
  const b = g.BACKEND_BLOCKS.METHOD2;
  assert.strictEqual(g.columnLetter_(b.col), "CQ", "the section is column CQ");
  // No section overlaps another, and each has an empty spacer column before the next one
  const blocks = Object.keys(g.BACKEND_BLOCKS).map(k => g.BACKEND_BLOCKS[k]).sort((x, y) => x.col - y.col);
  blocks.slice(1).forEach((o, i) => assert(o.col > blocks[i].col + blocks[i].width, o.title + " starts after a spacer column"));
  const f = g.method2ChoicesFormula_();
  assert(/^=ARRAYFORMULA\(LET\(/.test(f) && f.includes("'Contact Log'!K:K") && f.includes("TRIM('Contact Log'!E:E)=\"END\""), "reads K above END");
  g.METHOD2_FIXED_CHOICES_.forEach(c => assert(f.includes('"' + c + '"'), "fixed choice " + c));
  assert.strictEqual((f.match(/\(/g) || []).length, (f.match(/\)/g) || []).length, "balanced parentheses");
  assert(!/\$\d|\$[A-Z]+\d/.test(f.replace(/'Contact Log'!/g, "")), "no fixed cell addresses that a row move would break");

  assert.strictEqual(g.ensureMethod2ChoicesOnBackend_(ss), "added", "a sheet without it gets it");
  assert.strictEqual(be.getRange(1, b.col).getValue(), b.title, "section title");
  assert.strictEqual(be.getRange(2, b.col).getValue(), "Method 2", "header");
  assert.strictEqual(be.getRange(3, b.col).getFormula(), f, "the formula in CQ3");
  assert.strictEqual(g.ensureMethod2ChoicesOnBackend_(ss), "kept", "the next build keeps it");

  // The owner's defaults (2026-11-24), in this order, before the typed ones
  eq(g.METHOD2_FIXED_CHOICES_, ["Call", "Email", "Text", "Voicemail", "Meeting", "In School", "Parents Square"]);
  assert(f.includes('VSTACK("Call", "Email", "Text", "Voicemail", "Meeting", "In School", "Parents Square", typed)'), "fixed choices first");

  // Changed by hand (other fixed choices): kept as it is
  const mine = f.replace('"Email", ', '"Email", "Home Visit", ');
  be.getRange(3, b.col).setFormula(mine);
  assert.strictEqual(g.ensureMethod2ChoicesOnBackend_(ss), "kept");
  assert.strictEqual(be.getRange(3, b.col).getFormula(), mine, "your own list stays");

  // Blocked by something typed under it (#REF!): the column under it is cleared, the formula kept
  be._put([["Stray"]], 6, b.col);
  be._set(3, b.col, "#REF!"); be._formulas.set("3," + b.col, mine);
  assert.strictEqual(g.ensureMethod2ChoicesOnBackend_(ss), "unblocked");
  assert.strictEqual(be.getRange(6, b.col).getValue(), "", "the blocking value cleared");
  assert.strictEqual(be.getRange(3, b.col).getFormula(), mine, "formula kept");

  // Something else there: never overwritten
  const ss2 = makeSpreadsheet(["Contact Log", "Backend"]);
  const g2 = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss2) } });
  const be2 = ss2.getSheetByName("Backend");
  g2.ensureMethod2ChoicesOnBackend_(ss2);
  be2.getRange(3, b.col).setFormula("=SUM(A1:A3)");
  assert(/^skipped/.test(g2.ensureMethod2ChoicesOnBackend_(ss2)), "another formula: skipped");
  assert.strictEqual(be2.getRange(3, b.col).getFormula(), "=SUM(A1:A3)");
  be2._set(3, b.col, ""); be2._put([["Call"], ["Text"]], 3, b.col);
  assert(/^skipped/.test(g2.ensureMethod2ChoicesOnBackend_(ss2)), "typed values: skipped");
  eq(be2.getRange(3, b.col, 2, 1).getValues(), [["Call"], ["Text"]], "the typed values stay");
  const build = fs.readFileSync(path.join(ROOT, 'Build.gs'), 'utf8');
  assert(/ensureMethod2ChoicesOnBackend_\(ss\)/.test(build), "every build checks it");
  console.log("10 Method 2 choices: ok");
}

// ---------------------------------------------------------------- 2. Insight has only two contacts
{
  const g = loadProject();
  const ats = g.IDM_STEPS[0];
  assert.strictEqual(ats.key, "ats");
  assert(/2 guardians per student/.test(ats.note) && /Parent Square, NYCSA/.test(ats.note), "the note on the Student Data step");
  assert(/2 per student/.test(ats.uses), "the start page says it too");
  assert(g.IDM_STEPS.slice(1).every(s => !s.note), "only on Student Data");
  const html = fs.readFileSync(path.join(ROOT, 'InputDataSidebar.html'), 'utf8');
  assert(/id="stepNote"/.test(html) && /show\('stepNote', !!v\.step\.note\)/.test(html), "the panel shows it when the step has one");
  const idm = fs.readFileSync(path.join(ROOT, 'Input Data.gs'), 'utf8');
  assert(/note: step\.note \|\| ''/.test(idm), "the view passes it");
  console.log("2 Insight note: ok");
}

// ---------------------------------------------------------------- 8. phone sync: whose Google Contacts
{
  const { propStore } = require('./lib/load');
  const mkDrive = email => ({ getRootFolder: () => ({ getOwner: () => (email === null ? null : { getEmail: () => email }) }) });
  const mkUi = answers => {
    const ui = { _alerts: [], ButtonSet: { YES_NO: 'YES_NO' }, Button: { YES: 'YES', NO: 'NO' },
      alert: (...a) => { ui._alerts.push(a); return answers.length ? answers.shift() : 'NO'; }, showSidebar: () => { ui._sidebars = (ui._sidebars || 0) + 1; } };
    return ui;
  };
  const make = (email, answers, triggers) => {
    const ui = mkUi(answers || []);
    const user = propStore();
    const g = loadProject({ extra: {
      DriveApp: mkDrive(email),
      SpreadsheetApp: { getUi: () => ui, getActiveSpreadsheet: () => null },
      ScriptApp: { getProjectTriggers: () => (triggers || []).map(id => ({ getUniqueId: () => id, getHandlerFunction: () => "nightlyMasterSync" })) }
    } });
    g.PropertiesService.getUserProperties = () => user;
    return { g, ui, user };
  };

  // Whose account: mine (no nightly job yet, or my own), someone else's (theirs, once remembered)
  let { g } = make("Jo.Doe@schools.nyc.gov");
  assert.strictEqual(g.myAccountEmail_(), "jo.doe@schools.nyc.gov", "read through Drive, lower case");
  eq(g.getPhoneSyncAccount(), { me: "jo.doe@schools.nyc.gov", target: "jo.doe@schools.nyc.gov", nightlyIsMine: true, doe: true, domain: "schools.nyc.gov" }, "no nightly job yet: mine");
  assert.strictEqual(g.isDoeAccount_("a@gmail.com"), false);
  assert.strictEqual(g.isDoeAccount_("a@notschools.nyc.gov"), false, "the domain itself, not a look-alike");
  assert.strictEqual(g.isDoeAccount_(""), null, "unknown");
  ({ g } = make("co.worker@schools.nyc.gov", [], ["MINE"]));
  g.__props.NIGHTLY_TRIGGER_UID = "OTHER"; g.__props.NIGHTLY_TRIGGER_UID_SET = String(Date.now());
  let a = g.getPhoneSyncAccount();
  assert.strictEqual(a.nightlyIsMine, false, "another person's nightly job that still runs");
  assert.strictEqual(a.target, "", "their address isn't known yet");
  g.__props.NIGHTLY_OWNER_EMAIL = "setup.person@gmail.com";
  a = g.getPhoneSyncAccount();
  eq([a.target, a.doe], ["setup.person@gmail.com", false], "their account (not DOE: the strong warning)");
  g.__props.NIGHTLY_TRIGGER_UID = "MINE";
  assert.strictEqual(g.getPhoneSyncAccount().target, "co.worker@schools.nyc.gov", "my own nightly job: mine");
  // Drive can't say: unknown, the generic words
  ({ g } = make(null));
  eq([g.getPhoneSyncAccount().target, g.getPhoneSyncAccount().doe], ["", null]);

  // The nightly job's account is remembered with the job, and learned at the next run on older sheets
  ({ g } = make("setup@schools.nyc.gov"));
  g.rememberNightlyTrigger_(g.PropertiesService.getDocumentProperties(), "T1");
  assert.strictEqual(g.__props.NIGHTLY_OWNER_EMAIL, "setup@schools.nyc.gov", "remembered with the trigger");
  delete g.__props.NIGHTLY_OWNER_EMAIL;
  g.nightlyPausedInTrash_ = () => false; g.nightlyPausedByReplacement_ = () => false;
  assert.strictEqual(g.nightlyTriggerShouldRun_({ triggerUid: "T1" }), true);
  assert.strictEqual(g.__props.NIGHTLY_OWNER_EMAIL, "setup@schools.nyc.gov", "learned at the next nightly run");

  // 🧱 Full Sync: asked once per person for a DOE account; NO stops it
  let ui, user;
  ({ g, ui, user } = make("jo@schools.nyc.gov", ["NO"]));
  assert.strictEqual(g.confirmFullSyncAccount_(), false, "No = no sync");
  assert(/jo@schools\.nyc\.gov/.test(ui._alerts[0][1]) && /DOE phone/.test(ui._alerts[0][1]), "names the account and the DOE rule");
  assert(!/not a DOE/.test(ui._alerts[0][1]));
  g.CacheService.getScriptCache().remove('RUN_MODE');
  g.openProgressBarSidebar();
  assert.strictEqual(g.CacheService.getScriptCache().get('RUN_MODE'), null, "the menu item stops before the panel");
  assert.strictEqual(ui._sidebars || 0, 0);
  ({ g, ui, user } = make("jo@schools.nyc.gov", ["YES"]));
  assert.strictEqual(g.confirmFullSyncAccount_(), true);
  assert.strictEqual(user.getProperty("FULL_SYNC_OK_FOR"), "jo@schools.nyc.gov", "remembered for this person");
  assert.strictEqual(g.confirmFullSyncAccount_(), true, "the next time: not asked again");
  assert.strictEqual(ui._alerts.length, 1);
  // Not a DOE account: asked every time, with the warning first
  ({ g, ui, user } = make("me@gmail.com", ["YES", "YES"]));
  assert.strictEqual(g.confirmFullSyncAccount_(), true);
  assert.strictEqual(g.confirmFullSyncAccount_(), true);
  assert.strictEqual(ui._alerts.length, 2, "asked every time");
  assert(/^⚠️ me@gmail\.com is not a DOE/.test(ui._alerts[0][1]));
  assert.strictEqual(user.getProperty("FULL_SYNC_OK_FOR"), null, "never remembered");

  // ⚙️ Settings: the switch goes on only after "I understand"
  const html = fs.readFileSync(path.join(ROOT, 'SettingsDialog.html'), 'utf8');
  assert(/id="syncWarn"/.test(html) && /I understand, turn it on/.test(html) && /getPhoneSyncAccount\(\)/.test(html), "the warning in Settings");
  assert(!/Sync contacts to my phone/.test(html), "no more \"my phone\" (a co-worker's switch fills the nightly job's account)");
  console.log("8 phone sync warning: ok");
}

// ---------------------------------------------------------------- 7. Save stores an added parent the safe way
function rawRow(osis, last, first, cls, g1, g2) {
  const r = new Array(31).fill("");
  r[0] = "2025-2026"; r[2] = osis; r[3] = last; r[4] = first; r[16] = "3"; r[17] = cls; r[18] = "English";
  const put = (base, g) => { if (!g) return; r[base] = g[0]; r[base + 1] = g[1]; r[base + 2] = g[2]; r[base + 3] = g[3]; r[base + 4] = "English"; };
  put(19, g1); put(25, g2);
  return r;
}
function overridesSheet() {
  const ss = makeSpreadsheet(["Contact Log", "Raw Data", "Student List", "Backend"]);
  const raw = ss.getSheetByName("Raw Data");
  raw._put([["ATS - DATE UPDATED: 11/01/26"]], 1, 1);
  raw._put([
    rawRow("123456789", "SIMPSON", "BART", "Y41", ["HOMER SIMPSON", "FATHER", "1111111111", "homer@x.org"], ["MARGE SIMPSON", "MOTHER", "2222222222", "marge@x.org"]),
    rawRow("567891234", "SIMPSON", "LISA", "Y25", ["HOMER SIMPSON", "FATHER", "1111111111", "homer@x.org"], ["MARGE SIMPSON", "MOTHER", "2222222222", "marge@x.org"])
  ], 7, 1);
  const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) }, props: { disableRebuilds: "true" } });
  ss.getSheetByName("Backend")._put([g.__get("OVERRIDES_HEADERS_")], 2, 7);
  g.od_loadLists_ = () => [];
  return { ss, g, raw };
}
// What New / Edit Student sends for a student picked in the search (the dialog fills the form from it)
function formFor(g, osis, change) {
  const s = JSON.parse(JSON.stringify(g.getAllStudentsForEdit().find(x => x.osis === osis)));
  const p = { entryType: "Edit Existing Student", osis: s.osis, classCode: s.classCode, firstName: s.firstName, lastName: s.lastName,
    label: s.label, hideFromReport: false, grade: s.grade, age: s.age, gender: s.gender, homeLanguage: s.homeLanguage,
    labelAtOpen: s.label || "Current", prefilled: null, guardians: s.guardians };
  if (change) change(p);
  return p;
}
const rowsOf = (g, ss) => g.backendReadBlock_(ss, "OVERRIDES").slice(1).filter(x => String(x[2]).trim());
const abe = { name: "Abe Simpson", relationship: "Grandfather", phone: "3333333333", email: "abe@x.org", language: "English" };
{
  let { ss, g, raw } = overridesSheet();
  // Only added a parent: an "Added Guardian" row with just that parent; the ATS parents stay live
  g.saveStudentOverrides([formFor(g, "123456789", p => p.guardians.push(abe))]);
  let rows = rowsOf(g, ss);
  assert.strictEqual(rows.length, 1);
  assert.strictEqual(rows[0][1], "Added Guardian (by hand)", "saved the safe way");
  eq([rows[0][3], rows[0][4], rows[0][8], rows[0][10], rows[0][16], rows[0][40]], ["SIMPSON", "BART", "Y41", "Abe Simpson", "", ""], "only the new parent, no label");
  let ov = g.getOverrideData_(ss);
  assert(!ov.students.has("123456789"), "the student still comes from ATS");
  const names = id => g.buildGuardianObjects_(raw.getDataRange().getValues(), g.getOverrideData_(ss).guardians)
    .filter(x => x.student_id === id).map(x => x.guardianOutput.replace(/ \(.*\)$/, "") + " " + x.mobile);
  eq(names("123456789"), ["Homer Simpson (111) 111-1111", "Marge Simpson (222) 222-2222", "Abe Simpson (333) 333-3333"]);
  // ...so next month's ATS phone for Homer comes in (an Edit row kept the old one forever)
  raw._put([["9999999999"]], 7, 22);
  eq(names("123456789")[0], "Homer Simpson (999) 999-9999", "ATS changes still reach the ATS parents");
  raw._put([["1111111111"]], 7, 22);
  // Edit on that card + Save without changes: stays the same kind of row (it used to turn into an Edit row)
  g.saveStudentOverrides([formFor(g, "123456789")]);
  rows = rowsOf(g, ss);
  eq([rows.length, rows[0][1], rows[0][10]], [1, "Added Guardian (by hand)", "Abe Simpson"], "re-saved: still added, one row");
  // ...and adding another keeps both, with both sources
  ({ ss, g, raw } = overridesSheet());
  ss.getSheetByName("Backend")._put([["11/18/2026", "Added Guardian (NYCSA)", "123456789", "SIMPSON", "BART", "", "", "", "Y41", "",
    "Patty Bouvier", "Aunt", "4444444444", "", "", ""]], 3, 7);
  g.saveStudentOverrides([formFor(g, "123456789", p => p.guardians.push(abe))]);
  rows = rowsOf(g, ss);
  eq([rows.length, rows[0][1], rows[0][10], rows[0][16]], [1, "Added Guardian (NYCSA, by hand)", "Patty Bouvier", "Abe Simpson"], "both, with where each came from");

  // "Check and edit before adding": the parent keeps its source
  ({ ss, g, raw } = overridesSheet());
  g.saveStudentOverrides([formFor(g, "123456789", p => { p.guardians.push(abe); p.prefilled = { name: "Abe Simpson", sources: ["Parent Square"] }; })]);
  assert.strictEqual(rowsOf(g, ss)[0][1], "Added Guardian (Parent Square)", "from Parents to add: its source");

  // Real edits stay Edit rows (with every parent), as before
  const realEdits = [
    ["an ATS parent's phone changed", p => { p.guardians[0].phone = "5550000000"; p.guardians.push(abe); }],
    ["an ATS parent taken out", p => { p.guardians = [p.guardians[0], abe]; }],
    ["the class changed", p => { p.classCode = "Y99"; p.guardians.push(abe); }],
    ["the label changed", p => { p.label = "Moved"; p.guardians.push(abe); }],
    ["nothing added", p => {}],
    ["a new 'parent' who is an ATS parent in another order", p => { p.guardians.push(Object.assign({}, abe, { name: "Simpson Homer" })); }]
  ];
  realEdits.forEach(([what, change]) => {
    const t = overridesSheet();
    t.g.saveStudentOverrides([formFor(t.g, "123456789", change)]);
    const r = rowsOf(t.g, t.ss);
    assert.strictEqual(r[0][1], "Edit Existing Student", what + ": a real edit");
  });
  // An older window (no labelAtOpen): only an untouched "Current" / blank label counts as unchanged
  const t2 = overridesSheet();
  t2.g.saveStudentOverrides([formFor(t2.g, "123456789", p => { delete p.labelAtOpen; p.guardians.push(abe); })]);
  assert(/^Added Guardian/.test(rowsOf(t2.g, t2.ss)[0][1]), "older window, label Current");
  // A new student (not in ATS) is never turned into anything else
  const t3 = overridesSheet();
  t3.g.saveStudentOverrides([{ entryType: "New Student", osis: "999888777", classCode: "Y1", firstName: "Tina", lastName: "Belcher", label: "Current",
    hideFromReport: false, grade: "", age: "", gender: "", homeLanguage: "", guardians: [abe] }]);
  assert.strictEqual(rowsOf(t3.g, t3.ss)[0][1], "New Student");
  // "Sync to all siblings": each sibling gets its own added row
  const t4 = overridesSheet();
  const main = formFor(t4.g, "123456789", p => p.guardians.push(abe));
  const sib = formFor(t4.g, "567891234", p => { p.guardians = main.guardians; });
  t4.g.saveStudentOverrides([main, sib]);
  eq(rowsOf(t4.g, t4.ss).map(r => r[1] + " " + r[2] + " " + r[10]), ["Added Guardian (by hand) 123456789 Abe Simpson", "Added Guardian (by hand) 567891234 Abe Simpson"], "siblings too");
  // The dialog sends what the server needs
  const html = fs.readFileSync(path.join(ROOT, 'NewStudentDialog.html'), 'utf8');
  assert(/labelAtOpen: labelAtOpen,/.test(html) && /prefilled: prefilled,/.test(html) && /labelAtOpen: sib\.label \|\| data\.labelAtOpen/.test(html), "the dialog sends labelAtOpen and prefilled");
  console.log("7 add a parent: ok");
}

// ---------------------------------------------------------------- 5. several phone numbers per parent
{
  const g = loadProject();
  // Every number in one cell; one number reads exactly as before
  eq(g.phonesFromCell_("(718) 555-0101"), ["(718) 555-0101"]);
  eq(g.phonesFromCell_("7185550101"), ["(718) 555-0101"]);
  eq(g.phonesFromCell_("1-718-555-0101 ext 12"), ["(718) 555-0101"], "extension dropped, as before");
  eq(g.phonesFromCell_("718-555-0101 / 917-555-0144"), ["(718) 555-0101", "(917) 555-0144"], "two numbers: both (it was none)");
  eq(g.phonesFromCell_("7185550101 9175550144"), ["(718) 555-0101", "(917) 555-0144"]);
  eq(g.phonesFromCell_("(718)555-0101, (718) 555-0101; 917.555.0144"), ["(718) 555-0101", "(917) 555-0144"], "no repeats");
  eq(g.phonesFromCell_("555-0101"), [], "7 digits: no number, as before");
  eq(g.phonesFromCell_(""), []);
  eq(g.phonesFromCell_(7185550101), ["(718) 555-0101"], "a number cell");
  const one = g.buildGuardianObject_("123456789", "MARGE SIMPSON", "MOTHER", "English", "", "m@x.org", "7182220000");
  eq([one.mobile, one.phones], ["(718) 222-0000", ["(718) 222-0000"]], "mobile = the first, as before");
  const two = g.buildGuardianObject_("123456789", "MARGE SIMPSON", "MOTHER", "English", "", "m@x.org", "718-222-0000 or 917-333-0000");
  eq([two.mobile, two.phones], ["(718) 222-0000", ["(718) 222-0000", "(917) 333-0000"]]);
  eq(g.guardianPhones_({ mobile: "(718) 222-0000" }), ["(718) 222-0000"], "an object without phones (owner hooks): its mobile");
  assert.strictEqual(g.formatGuardianDirectoryLine_(two), "Marge Simpson (Mother) - m@x.org, (718) 222-0000, (917) 333-0000", "Directory line: every number");
}
{
  // Phone Contacts: one contact per parent, every number / email, siblings together
  const ss = makeSpreadsheet(["Contact Log", "Backend", "Raw Data"]);
  const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) } });
  const G = (id, name, email, phone) => g.buildGuardianObject_(id, name, "MOTHER", "English", "", email, phone);
  const studentMap = new Map([["111", { student: "Bart Simpson", classCode: "Y41", site: "", teachers: "" }],
                              ["222", { student: "Lisa Simpson", classCode: "Y25", site: "", teachers: "" }],
                              ["333", { student: "Maggie Simpson", classCode: "Y01", site: "", teachers: "" }]]);
  const master = [["OSIS"], ["111"], ["222"], ["333"]];
  const psIdx = g.buildParentSquareIndex_([]), nyIdx = g.buildNYCSAIndex_([]);
  const build = rows => g.buildPhoneContactsData_(ss, studentMap, rows, master, [["OSIS"]], psIdx, nyIdx);
  const H = build([])[0];
  const col = name => H.indexOf(name);
  // Same parent, a different number on each child's record: ONE contact with both numbers (was two)
  let out = build([G("111", "MARGE SIMPSON", "m@x.org", "7182220000"), G("222", "MARGE SIMPSON", "m@x.org", "9173330000")]);
  assert.strictEqual(out.length, 2, "one contact");
  eq([out[1][col("Last Name")], out[1][col("Phone 1 - Value")], out[1][col("Phone 2 - Value")], out[1][col("E-mail 1 - Value")]],
    ["| Student: Bart Simpson, Lisa Simpson |", "(718) 222-0000", "(917) 333-0000", "m@x.org"]);
  assert.strictEqual(out[1][col("Phone 2 - Type")], out[1][col("Phone 1 - Type")], "the 2nd number has the same class label");
  // Same number, two emails: one contact, E-mail 2
  out = build([G("111", "MARGE SIMPSON", "m@x.org", "7182220000"), G("222", "MARGE SIMPSON", "marge@y.org", "7182220000")]);
  eq([out.length, out[1][col("E-mail 1 - Value")], out[1][col("E-mail 2 - Value")]], [2, "m@x.org", "marge@y.org"]);
  // Same name, nothing shared: two people, two contacts (as before)
  out = build([G("111", "MARIA LOPEZ", "a@x.org", "7182220000"), G("333", "MARIA LOPEZ", "b@x.org", "6465550000")]);
  assert.strictEqual(out.length, 3, "two contacts");
  // No number and no email: grouped by name, as before
  out = build([G("111", "MARGE SIMPSON", "", ""), G("222", "MARGE SIMPSON", "", "")]);
  assert.strictEqual(out.length, 2, "one bare contact");
  // A row that links two groups (one shares the number, the other the email) makes them one
  out = build([G("111", "MARGE SIMPSON", "", "7182220000"), G("222", "MARGE SIMPSON", "m@x.org", "9173330000"), G("333", "MARGE SIMPSON", "m@x.org", "7182220000")]);
  assert.strictEqual(out.length, 2, "one contact for all three");
  eq([out[1][col("Phone 1 - Value")], out[1][col("Phone 2 - Value")], out[1][col("OSIS")]], ["(718) 222-0000", "(917) 333-0000", "111, 222, 333"]);
  // A cell with two numbers: Phone 1 and Phone 2; three is the most
  out = build([G("111", "MARGE SIMPSON", "", "7182220000 / 9173330000 / 6465550000 / 3475550000")]);
  eq([out[1][col("Phone 1 - Value")], out[1][col("Phone 2 - Value")], out[1][col("Phone 3 - Value")]], ["(718) 222-0000", "(917) 333-0000", "(646) 555-0000"]);
  // Parents List: every number in its one phone column; siblings found by any number
  const siblings = g.buildSiblingMap_(studentMap, [G("111", "MARGE SIMPSON", "", "7182220000 / 9173330000"), G("222", "MARGE SIMPSON", "", "9173330000")]);
  assert(/222 - Lisa Simpson/.test(siblings.get("111")), "a sibling found by the 2nd number");
  const pl = g.buildParentsDividedData_(ss, [G("111", "MARGE SIMPSON", "", "7182220000 / 9173330000")], new Map(), psIdx, nyIdx, studentMap, [], []);
  assert.strictEqual(pl[1][pl[0].indexOf("mobile")], "(718) 222-0000, (917) 333-0000", "Parents List: both numbers");
}
{
  // The shared sync writes the 2nd / 3rd number and email; a row with one of each builds the same contact
  const saved = [];
  const People = { People: {
    Connections: { list: () => ({ connections: [] }) },
    createContact: d => { saved.push(JSON.parse(JSON.stringify(d))); return { resourceName: "people/c" + saved.length }; },
    updateContact: d => { saved.push(JSON.parse(JSON.stringify(d))); return { resourceName: "people/x" }; }
  } };
  const g = loadProject({ extra: { People } });
  const H = ["First Name", "Last Name", "E-mail 1 - Value", "E-mail 2 - Value", "E-mail 3 - Value", "Phone 1 - Type", "Phone 1 - Value",
             "Phone 2 - Type", "Phone 2 - Value", "Phone 3 - Type", "Phone 3 - Value", "Notes", "Organization Name", "Organization Department"];
  const read = g.phoneContactsRowReader_(H);
  const index = g.buildContactsIndex_("contactGroups/sd");
  read(["(Mother) Marge Simpson", "| Student: Bart |", "m@x.org", "", "", "Class: Y41", "(718) 222-0000", "", "", "", "", "n", "OSIS: 1", ""]).upsert("contactGroups/sd", index);
  read(["(Mother) Marge Simpson", "| Student: Lisa |", "m@x.org", "marge@y.org", "", "Class: Y25", "(718) 222-0000", "Class: Y25", "(917) 333-0000", "", "", "n", "OSIS: 2", ""]).upsert("contactGroups/sd", index);
  eq(saved[0].phoneNumbers, [{ value: "(718) 222-0000", type: "Class: Y41" }], "one number: exactly as before");
  eq(saved[0].emailAddresses, [{ value: "m@x.org", type: "home" }]);
  eq(saved[1].phoneNumbers, [{ value: "(718) 222-0000", type: "Class: Y25" }, { value: "(917) 333-0000", type: "Class: Y25" }], "both numbers");
  eq(saved[1].emailAddresses, [{ value: "m@x.org", type: "home" }, { value: "marge@y.org", type: "home" }], "both emails");
  // An older Phone Contacts tab (no Phone 2 / E-mail 2 columns): nothing extra, nothing read from the wrong column
  const old = g.phoneContactsRowReader_(["First Name", "Last Name", "E-mail 1 - Value", "Phone 1 - Type", "Phone 1 - Value", "Notes", "Organization Name", "Organization Department"]);
  old(["(Mother) Ann", "| Student: Kid |", "", "mobile", "2125550101", "", "", ""]).upsert("contactGroups/sd", index);
  eq(saved[2].phoneNumbers, [{ value: "2125550101", type: "mobile" }]);
  eq(saved[2].emailAddresses, []);
  console.log("5 several phone numbers: ok");
}

// ---------------------------------------------------------------- 4. Phone numbers to check
function phonesSheet() {
  const ss = makeSpreadsheet(["Contact Log", "Raw Data", "Student List", "Backend"]);
  const raw = ss.getSheetByName("Raw Data");
  raw._put([["ATS - DATE UPDATED: 11/01/26"]], 1, 1);
  raw._put([
    rawRow("123456789", "SIMPSON", "BART", "Y41", ["HOMER SIMPSON", "FATHER", "1111111111", "homer@x.org"], ["MARGE SIMPSON", "MOTHER", "2222222222", "marge@x.org"]),
    rawRow("567891234", "SIMPSON", "LISA", "Y25", ["HOMER SIMPSON", "FATHER", "1111111111", "homer@x.org"], ["MARGE SIMPSON", "MOTHER", "2222222222", "marge@x.org"]),
    rawRow("111222333", "ADDAMS", "WEDNESDAY", "Y12", ["MORTICIA ADDAMS", "MOTHER", "5555555555", ""], ["GOMEZ ADDAMS", "FATHER", "6666666666", ""])
  ], 7, 1);
  // Parent Square (AG:AN): Name, Students, Email, Phone, Secondary Phone, ...
  raw._put([
    ["Simpson, Homer", "Bart Simpson, Lisa Simpson", "homer@x.org", "(999) 999-0000", "", "", "Yes", "Yes"], // another number, both kids
    ["Simpson, Marjorie", "Lisa Simpson", "marge@x.org", "(777) 777-0000", "", "", "Yes", "Yes"],          // Marge by email, another spelling
    ["Addams, Gomez", "Wednesday Addams", "", "(666) 666-6666", "", "", "Yes", "Yes"]                      // the same number: nothing
  ], 7, 33);
  const ny = (osis, sf, sl, gf, gl, home, work, cell, rel) => ["M094", "75", "PS", "Y41", "3", osis, sf, sl, gf, gl, "No", home, work, cell, rel, "Custodial User", "C1"];
  raw._put([
    ny("123456789", "BART", "SIMPSON", "HOMER", "SIMPSON", "3333333333", "", "1111111111", "FATHER"),   // cell agrees: the home number is noise
    ny("111222333", "WEDNESDAY", "ADDAMS", "MORTICIA", "ADDAMS", "", "", "8888888888", "MOTHER"),     // another cell
    ny("111222333", "WEDNESDAY", "ADDAMS", "GRANDMAMA", "ADDAMS", "", "", "7777777777", "GRANDMOTHER") // a new parent: Parents to add
  ], 7, 42);
  const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) } });
  ss.getSheetByName("Backend")._put([g.__get("OVERRIDES_HEADERS_")], 2, 7);
  let builds = 0;
  g.runBuildOnly = () => { builds++; };
  g.od_loadLists_ = () => [];
  return { ss, g, raw, builds: () => builds };
}
{
  const { ss, g, raw, builds } = phonesSheet();
  let found = g.findPhonesToCheck_(ss);
  const show = f => f.items.map(it => it.guardian + " " + it.phone + " [" + it.kinds.join(", ") + "] for " + it.students.map(s => s.name).join(" & "));
  eq(show(found), [
    "Morticia Addams (888) 888-8888 [NYCSA (cell)] for Wednesday Addams",
    "Homer Simpson (999) 999-0000 [Parent Square (phone)] for Bart Simpson & Lisa Simpson",
    "Marge Simpson (777) 777-0000 [Parent Square (phone)] for Lisa Simpson"
  ], "one card per parent + number, siblings together; NYCSA's extra home number next to the cell we have is left out");
  const homer = found.items.find(it => it.guardian === "Homer Simpson");
  eq(homer.current, ["(111) 111-1111"], "the number they have now");
  const marge = found.items.find(it => it.guardian === "Marge Simpson");
  assert(marge.alike && marge.alike.name === "Marjorie Simpson" && marge.alike.pct >= 50, "another spelling: how alike (" + JSON.stringify(marge.alike) + ")");
  assert.strictEqual(homer.alike, null, "the same name: no % line");
  // Parents to add still finds only the new parent (the refactor changed nothing there)
  eq(g.findParentsToAdd_(ss).items.map(it => it.name), ["Grandmama Addams"]);
  // The page gets both sections from one read
  const home = g.getOverridesHome();
  eq([home.toAdd.length, home.phones.length, home.phonesSkipped], [1, 3, 0]);

  // Add as a 2nd number: one row per student, then a build; the cards go
  let r = g.savePhoneNumbers([homer.key], "add");
  eq([r.saved, r.building, builds()], [1, true, 1]);
  const rows = g.backendReadBlock_(ss, "PHONES");
  eq(rows[0], g.PHONES_HEADERS_, "header written");
  eq(rows.slice(1).map(x => [x[0], x[1], x[2], x[3], x[4]]), [
    ["123456789", "Homer Simpson", "(999) 999-0000", "add", "Parent Square"],
    ["567891234", "Homer Simpson", "(999) 999-0000", "add", "Parent Square"]]);
  const numbers = (id, name) => {
    const ov = g.getOverrideData_(ss);
    const gr = g.buildGuardianObjects_(raw.getDataRange().getValues(), ov.guardians, ov.phones).find(x => x.student_id === id && x.guardianOutput.indexOf(name) === 0);
    return [gr.mobile, gr.phones];
  };
  eq(numbers("123456789", "Homer"), ["(111) 111-1111", ["(111) 111-1111", "(999) 999-0000"]], "the build: a 2nd number, the ATS one first");
  eq(numbers("567891234", "Homer"), ["(111) 111-1111", ["(111) 111-1111", "(999) 999-0000"]], "for the sibling too");
  eq(numbers("123456789", "Marge"), ["(222) 222-2222", ["(222) 222-2222"]], "nobody else changes");
  assert(!g.findPhonesToCheck_(ss).items.some(it => it.guardian === "Homer Simpson"), "not offered again");
  // The Directory stays ATS only
  eq(g.buildGuardianObjects_(raw.getDataRange().getValues(), new Map()).find(x => x.guardianOutput.indexOf("Homer") === 0).phones, ["(111) 111-1111"]);

  // Use it instead: that number only
  const mort = g.findPhonesToCheck_(ss).items.find(it => it.guardian === "Morticia Addams");
  g.savePhoneNumbers([mort.key], "instead");
  eq(numbers("111222333", "Morticia"), ["(888) 888-8888", ["(888) 888-8888"]]);

  // Not needed: hidden, counted, and back with Show them again
  const skippedHome = g.skipPhoneNumbers([marge.key]);
  eq([skippedHome.phones.length, skippedHome.phonesSkipped], [0, 1]);
  eq(g.unskipPhoneNumbers().phones.length, 1);

  // ATS gets the number: the row goes by itself at the next build's cleanup
  raw._put([["1111111111 / 9999990000"]], 7, 22); // Bart's Homer now has both in ATS
  g.cleanRedundantOverrides();
  eq(g.backendReadBlock_(ss, "PHONES").slice(1).map(x => x[0] + " " + x[3]), ["567891234 add", "111222333 instead of (555) 555-5555"], "Bart's row went; Lisa's ATS doesn't have it yet (\"instead\" notes the number it replaced)");
  raw._put([["8888888888"]], 9, 22);
  g.cleanRedundantOverrides();
  eq(g.backendReadBlock_(ss, "PHONES").slice(1).map(x => x[0]), ["567891234"], "\"instead\": goes once it is ATS's number");

  // Rebuilds paused: saved, no build, a toast
  const p2 = phonesSheet();
  p2.g.__props.disableRebuilds = "true";
  const h2 = p2.g.findPhonesToCheck_(p2.ss).items[0];
  r = p2.g.savePhoneNumbers([h2.key], "add");
  eq([r.saved, r.building, p2.builds()], [1, false, 0]);
  // A reset clears them; Migration copies them
  const regen = fs.readFileSync(path.join(ROOT, 'Regenerate.gs'), 'utf8');
  assert(/backendClearBlock_\(ss, "PHONES", BACKEND_HEADER_ROW \+ 1\)/.test(regen), "resets clear them");
  const menus = fs.readFileSync(path.join(ROOT, 'Menus.gs'), 'utf8');
  assert(/readBackendBlockFromOtherSpreadsheet_\(oldSS, "PHONES"\)/.test(menus), "Migration brings them");
  const html = fs.readFileSync(path.join(ROOT, 'ManageOverridesDialog.html'), 'utf8');
  ['savePhoneNumbers', 'skipPhoneNumbers', 'unskipPhoneNumbers', 'deletePhoneChanges'].forEach(fn => assert(html.indexOf('.' + fn + '(') !== -1, "the page calls " + fn));
  console.log("4 phone numbers to check: ok");
}
{
  // Your changes lists the saved numbers (one card per parent + number + use, siblings together);
  // Delete undoes one: its rows go, a build runs, the number is offered again, the others stay
  const { ss, g, raw, builds } = phonesSheet();
  ss.getSheetByName("Backend")._put([["OSIS", "Student"], ["123456789", "Bart Simpson"], ["567891234", "Lisa Simpson"]], 2, 76); // BX:BY copy
  eq(g.getOverridesHome().phoneChanges, [], "nothing saved yet");
  const found = g.findPhonesToCheck_(ss).items;
  const homer = found.find(it => it.guardian === "Homer Simpson"), mort = found.find(it => it.guardian === "Morticia Addams");
  g.savePhoneNumbers([homer.key], "add");
  g.savePhoneNumbers([mort.key], "instead");
  const saved = g.getOverridesHome().phoneChanges;
  eq(saved.map(c => [c.guardian, c.phone, c.use, c.from, c.students.map(s => s.name)]), [
    ["Homer Simpson", "(999) 999-0000", "add", "Parent Square", ["Bart Simpson", "Lisa Simpson"]],
    ["Morticia Addams", "(888) 888-8888", "instead", "NYCSA", ["OSIS 111222333"]]], "one card each; a student the copy lacks shows its OSIS");
  const numbers = (id, name) => {
    const ov = g.getOverrideData_(ss);
    return g.buildGuardianObjects_(raw.getDataRange().getValues(), ov.guardians, ov.phones).find(x => x.student_id === id && x.guardianOutput.indexOf(name) === 0).phones;
  };
  eq(numbers("111222333", "Morticia"), ["(888) 888-8888"], "\"instead\" in effect");
  const b0 = builds();
  let r = g.deletePhoneChanges([saved[1].key]);
  eq([r.removed, r.building, builds() - b0], [1, true, 1], "deleted, then a build");
  eq(numbers("111222333", "Morticia"), ["(555) 555-5555"], "ATS's number is back");
  eq(numbers("123456789", "Homer"), ["(111) 111-1111", "(999) 999-0000"], "the other saved number stays");
  assert(r.home.phones.some(it => it.guardian === "Morticia Addams" && it.phone === "(888) 888-8888"), "offered again to check");
  eq(r.home.phoneChanges.map(c => c.guardian), ["Homer Simpson"]);
  // Both siblings' rows go with one Delete; an unknown key does nothing
  r = g.deletePhoneChanges([saved[0].key]);
  eq([r.removed, r.home.phoneChanges.length], [2, 0]);
  eq(g.deletePhoneChanges(["nobody|0|add"]).removed, 0);
  // Rebuilds paused: deleted, no build
  const p2 = phonesSheet();
  p2.g.savePhoneNumbers([p2.g.findPhonesToCheck_(p2.ss).items[0].key], "add");
  p2.g.__props.disableRebuilds = "true";
  const before = p2.builds();
  r = p2.g.deletePhoneChanges([p2.g.getOverridesHome().phoneChanges[0].key]);
  eq([r.removed > 0, r.building, p2.builds() - before], [true, false, 0]);
  console.log("4 saved numbers can be deleted: ok");
}
{
  // The end of Input Data Mode counts both (the Student Data step's note says they're asked about there)
  const { g } = phonesSheet();
  const r = g.idm_parentsToAddSummary();
  eq([r.count, r.phones], [1, 3], "1 parent to add, 3 numbers to check");
  const html = fs.readFileSync(path.join(ROOT, 'InputDataSidebar.html'), 'utf8');
  assert(/r\.phones/.test(html) && /phone numbers to check/.test(html), "the end screen card says so");
  console.log("4 end of Input Data Mode: ok");
}
{
  // The test file 2c (test-data/README.md and the release check say what it shows): exactly these two cards
  const parseCsv = text => {
    const rows = []; let row = [], cell = '', q = false;
    text = text.replace(/^﻿/, '');
    for (let i = 0; i < text.length; i++) { const ch = text[i];
      if (q) { if (ch === '"' && text[i + 1] === '"') { cell += '"'; i++; } else if (ch === '"') q = false; else cell += ch; }
      else if (ch === '"') q = true; else if (ch === ',') { row.push(cell); cell = ''; }
      else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(cell); rows.push(row); row = []; cell = ''; }
      else cell += ch; }
    if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
    return rows;
  };
  const csv = n => parseCsv(fs.readFileSync(path.join(ROOT, 'test-data', n), 'utf8')).filter(r => r.some(c => String(c).trim())).slice(1);
  const ss = makeSpreadsheet(["Contact Log", "Raw Data", "Student List", "Backend"]);
  const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) } });
  g.od_loadLists_ = () => [];
  g.runBuildOnly = () => {};
  ss.getSheetByName("Backend")._put([g.__get("OVERRIDES_HEADERS_")], 2, 7);
  const raw = ss.getSheetByName("Raw Data");
  raw._put(csv("1 - Student Data (ATS).csv"), 7, 1);
  raw._put(csv("2c - Parent Square - new phone numbers.csv"), 7, 33);
  raw._put(csv("3 - NYCSA (Students with Non-NYCSA Account).csv"), 7, 42);
  const found = g.findPhonesToCheck_(ss);
  eq(found.items.map(it => it.guardian + " " + it.phone + " for " + it.students.map(s => s.name).join(" & ") + " (has " + it.current.join(", ") + ")" +
      (it.alike ? ", as " + it.alike.name + " " + it.alike.pct + "%" : "")), [
    "Morticia Addams (333) 555-0303 for Wednesday Addams (has (333) 333-3333), as Morticia Adams 99%",
    "Homer Simpson (111) 555-0101 for Bart Simpson & Lisa Simpson (has (111) 111-1111)"], "the two cards of the test file");
  eq(g.findParentsToAdd_(ss).items, [], "nobody to add from it");
  // Both as 2nd numbers: Homer has both for Bart and Lisa; Morticia both for Wednesday, one for Pugsley
  g.savePhoneNumbers(found.items.map(it => it.key), "add");
  const ov = g.getOverrideData_(ss);
  const gs = g.buildGuardianObjects_(raw.getDataRange().getValues(), ov.guardians, ov.phones);
  const nums = (id, first) => gs.find(x => x.student_id === id && x.guardianOutput.indexOf(first) === 0).phones;
  eq([nums("123456789", "Homer"), nums("567891234", "Homer"), nums("234567891", "Morticia"), nums("245678912", "Morticia")], [
    ["(111) 111-1111", "(111) 555-0101"], ["(111) 111-1111", "(111) 555-0101"], ["(333) 333-3333", "(333) 555-0303"], ["(333) 333-3333"]]);
  eq(g.findPhonesToCheck_(ss).items, [], "nothing left to check");
  // Phone Contacts: Morticia is ONE contact with both numbers (the same email links Wednesday's and Pugsley's rows)
  const ids = ["123456789", "567891234", "234567891", "245678912"];
  const studentMap = new Map(ids.map(id => [id, { student: id, classCode: "Y41", site: "", teachers: "" }]));
  const pc = g.buildPhoneContactsData_(ss, studentMap, gs.filter(x => ids.indexOf(x.student_id) !== -1), [["OSIS"]].concat(ids.map(id => [id])), [["OSIS"]],
    g.buildParentSquareIndex_([]), g.buildNYCSAIndex_([]));
  const col = name => pc[0].indexOf(name);
  const contact = first => pc.filter(r => String(r[col("First Name")]).indexOf(first) !== -1);
  eq(contact("Morticia").map(r => [r[col("Phone 1 - Value")], r[col("Phone 2 - Value")], r[col("OSIS")]]),
    [["(333) 333-3333", "(333) 555-0303", "234567891, 245678912"]], "one Morticia with both numbers");
  eq(contact("Homer").map(r => [r[col("Phone 1 - Value")], r[col("Phone 2 - Value")]]), [["(111) 111-1111", "(111) 555-0101"]], "one Homer with both numbers");
  console.log("4 the test file 2c: ok");
}

// ---------------------------------------------------------------- 3. "What is this list for?"
{
  const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(makeSpreadsheet(["Backend"])) } });
  // The setup: "info" unless said; an older sidebar (no purpose sent) keeps the saved one
  assert.strictEqual(g.od_emptySetup_().purpose, "info");
  assert.strictEqual(g.od_fillSetup_({ ready: true }).purpose, "info", "a list saved before: shows its info, as before");
  const clean = p => g.od_cleanSetupInput_(Object.assign({ roles: {}, vars: [] }, p), ["Parent", "Child", "Phone"]).purpose;
  eq([clean({ purpose: "contacts" }), clean({ purpose: "info" }), clean({}), clean({ purpose: "x" })], ["contacts", "info", null, null]);
  // Saved from the first screen (before the paste)
  let saved = null;
  const list = { tabId: 7, name: "PTA", sheet: {}, setup: g.od_fillSetup_({}) };
  g.od_withLock_ = fn => fn();
  g.od_wizardList_ = () => ({ lists: [list], list: list });
  g.od_saveLists_ = (ss, lists) => { saved = lists[0].setup.purpose; };
  eq(g.od_setListPurpose("contacts"), { ok: true, purpose: "contacts" });
  assert.strictEqual(saved, "contacts");
  eq(g.od_setListPurpose("nonsense").purpose, "info", "anything else = info");
}
{
  // A contacts list puts nothing on the sheets, but its tab and counts still update at the build
  const { g } = (function () {
    const ss = makeSpreadsheet(["Contact Log", "Backend"]);
    return { g: loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) } }) };
  })();
  const raw = [[], [], [], [], [], [], rawRow("123456789", "SIMPSON", "BART", "Y41", ["HOMER SIMPSON", "FATHER", "1111111111", ""], null)];
  const sm = g.buildStudentMap_(raw, new Map());
  const master = [["OSIS", "Student", "Guardian", "Site", "Class"], ["123456789", "Bart Simpson", "", "", "Y41"]];
  const guardians = g.buildGuardianObjects_(raw, new Map());
  const headers = ["Parent", "Child", "Phone", "Role"], rows = [["Homer Simpson", "Bart Simpson", "5551234567", "Chair"]];
  let addl = null, toolCols = 0, list = null;
  g.od_loadLists_ = () => [list];
  g.od_readTable_ = () => ({ headers: headers, rows: rows });
  g.od_writeToolColumns_ = () => { toolCols++; }; g.od_saveLists_ = () => {}; g.backendClearBlock_ = () => {};
  g.backendWriteBlock_ = (ss, key, r) => { if (key === "ADDL_INFO") addl = r; };
  const make = purpose => ({ tabId: 1, name: "PTA", sheet: { getSheetId: () => 1 },
    setup: g.od_fillSetup_({ ready: true, by: "guardian", purpose: purpose, roles: { guardian: { col: 1 }, student: { col: 2 }, phone: 3 },
      vars: [{ col: 4, header: "Role", label: "Role", show: true, column: true, blank: { mode: "nothing" }, notOnList: { mode: "nothing" } }], picks: {}, headers: headers }) });
  list = make("info");
  let out = g.od_buildOutputs_({}, master, guardians);
  assert(addl.length > 1 && out.guardianCols.length === 1, "info: a line and a column, as before");
  list = make("contacts");
  toolCols = 0;
  out = g.od_buildOutputs_({}, master, guardians);
  eq([addl.length, out.studentCols.length, out.guardianCols.length], [1, 0, 0], "contacts: no lines (only the header), no columns");
  assert.strictEqual(toolCols, 1, "its tab's Matched column still updates");
  assert(list.setup.stats && list.setup.stats.rows === 1, "and its counts");
}
{
  // ...and it feeds Parents to add and Phone numbers to check
  const { ss, g } = phonesSheet();
  const tab = ss.insertSheet("👪 Contact cards");
  const headers = ["Parent", "Child", "Phone"];
  tab._put([headers, ["Selma Bouvier", "Wednesday Addams", "4445556666"], ["Homer Simpson", "Bart Simpson", "(555) 123-4567"]], 4, 1);
  g.od_loadLists_ = () => [{ tabId: tab.getSheetId(), name: "Contact cards", sheet: tab,
    setup: g.od_fillSetup_({ by: "guardian", ready: true, purpose: "contacts", icon: "👪", headers: headers,
      roles: { guardian: { col: 1 }, student: { col: 2 }, phone: 3 }, vars: [], picks: {} }) }];
  const toAdd = g.findParentsToAdd_(ss).items.find(it => it.name === "Selma Bouvier");
  assert(toAdd && toAdd.sources[0] === "👪 Contact cards" && toAdd.phone === "(444) 555-6666", "a new parent from the list");
  const num = g.findPhonesToCheck_(ss).items.find(it => it.phone === "(555) 123-4567");
  assert(num && num.guardian === "Homer Simpson" && num.sources[0] === "👪 Contact cards", "another number for a parent we have: " + JSON.stringify(g.findPhonesToCheck_(ss).items.map(i => i.guardian + " " + i.phone)));
  console.log("3 list purpose: ok");
}
{
  // The test file 5e as "Only find parents and phone numbers to add", with the standard exports (the release check):
  // the column guess, 3 matched + 1 not found, then exactly Selma (to add) and Homer's other number (to check)
  const parseCsv = text => {
    const rows = []; let row = [], cell = '', q = false;
    text = text.replace(/^﻿/, '');
    for (let i = 0; i < text.length; i++) { const ch = text[i];
      if (q) { if (ch === '"' && text[i + 1] === '"') { cell += '"'; i++; } else if (ch === '"') q = false; else cell += ch; }
      else if (ch === '"') q = true; else if (ch === ',') { row.push(cell); cell = ''; }
      else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(cell); rows.push(row); row = []; cell = ''; }
      else cell += ch; }
    if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
    return rows;
  };
  const all = n => parseCsv(fs.readFileSync(path.join(ROOT, 'test-data', n), 'utf8')).filter(r => r.some(c => String(c).trim()));
  const ss = makeSpreadsheet(["Contact Log", "Raw Data", "Student List", "Backend"]);
  const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) } });
  const raw = ss.getSheetByName("Raw Data");
  const ats = all("1 - Student Data (ATS).csv").slice(1);
  raw._put(ats, 7, 1);
  raw._put(all("2 - Parent Square.csv").slice(1), 7, 33);
  raw._put(all("3 - NYCSA (Students with Non-NYCSA Account).csv").slice(1), 7, 42);
  const rawVals = [[], [], [], [], [], []].concat(ats);
  const sm = g.buildStudentMap_(rawVals, new Map());
  const idx = g.od_index_([...sm.entries()].map(([id, s]) => [id, s.student, "", s.site || "", s.classCode || ""]), g.buildGuardianObjects_(rawVals, new Map()));
  const block = all("5e - Your Lists - Parent Contact Cards (only find parents and numbers).csv");
  const headers = block[0], rows = block.slice(1);
  const guess = g.od_guessRoles_(headers, rows, idx);
  eq([guess.by, guess.roles.student, guess.roles.guardian, guess.roles.phone, guess.roles.email], ["student", { col: 2 }, { col: 1 }, 3, 4],
    "guessed: a row per child, Parent = the guardian column (so the contacts list can go straight on)");
  const tab = ss.insertSheet("📥 Parent Contact Cards");
  tab._put(block, 4, 1);
  g.od_loadLists_ = () => [{ tabId: tab.getSheetId(), name: "Parent Contact Cards", sheet: tab,
    setup: g.od_fillSetup_({ by: guess.by, ready: true, purpose: "contacts", headers: headers, roles: guess.roles, vars: [], picks: {} }) }];
  eq(g.findParentsToAdd_(ss).items.map(it => it.student + ": " + it.name + " " + it.phone + " [" + it.sources.join(", ") + "]"),
    ["Wednesday Addams: Selma Bouvier (444) 555-6666 [📥 Parent Contact Cards]"]);
  eq(g.findPhonesToCheck_(ss).items.map(it => it.guardian + " " + it.phone + " for " + it.students.map(s => s.name).join(" & ") + " [" + it.sources.join(", ") + "]"),
    ["Homer Simpson (111) 222-3333 for Bart Simpson [📥 Parent Contact Cards]"]);
  eq(g.idm_parentsToAddSummary(), { count: 1, sample: ["Selma Bouvier (Wednesday Addams)"], phones: 1 }, "the end of Input Data Mode");
  console.log("3 the test file 5e: ok");
}

// ---------------------------------------------------------------- 6. "Tell me what changed" (Your Lists)
{
  const g = loadProject();
  const m = obj => new Map(Object.keys(obj).map(k => [k, obj[k]]));
  const before = m({ a: { who: "Bart Simpson", values: { "Bus Route": "R12", Stop: "Elm" } }, b: { who: "Lisa Simpson", values: { "Bus Route": "R3" } } });
  const now = m({ a: { who: "Bart Simpson", values: { "Bus Route": "R14", Stop: "Elm", New: "x" } }, c: { who: "Gene <b>Belcher</b>", values: { "Bus Route": "R7" } } });
  eq(g.od_watchChanges_("🚌 Bus Routes", before, now), { list: "🚌 Bus Routes", added: ["Gene &lt;b&gt;Belcher&lt;/b&gt;"], removed: ["Lisa Simpson"],
    changed: ["Bart Simpson: Bus Route R12 → R14"] }, "new, gone, changed (only values in both copies; text escaped)");
  assert.strictEqual(g.od_watchChanges_("x", before, before), null, "nothing changed: nothing reported");
}
{
  // Two builds: the first only takes the copy; the second reports
  const ss = makeSpreadsheet(["Contact Log", "Backend"]);
  const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) } });
  const raw = [[], [], [], [], [], [],
    rawRow("123456789", "SIMPSON", "BART", "Y41", ["HOMER SIMPSON", "FATHER", "1111111111", ""], null),
    rawRow("567891234", "SIMPSON", "LISA", "Y25", ["HOMER SIMPSON", "FATHER", "1111111111", ""], null),
    rawRow("111222333", "BELCHER", "GENE", "Y12", ["LINDA BELCHER", "MOTHER", "2222222222", ""], null)];
  const sm = g.buildStudentMap_(raw, new Map());
  const master = [["OSIS", "Student", "Guardian", "Site", "Class"]].concat([...sm.entries()].map(([id, s]) => [id, s.student, "", "", s.classCode]));
  const guardians = g.buildGuardianObjects_(raw, new Map());
  const headers = ["OSIS", "Student", "Bus Route"];
  let rows = [["123456789", "Bart Simpson", "R12"], ["567891234", "Lisa Simpson", "R3"]];
  const list = { tabId: 42, name: "Bus Routes", sheet: { getSheetId: () => 42 },
    setup: g.od_fillSetup_({ ready: true, by: "student", watch: true, icon: "🚌", roles: { osis: 1, student: { col: 2 } },
      vars: [{ col: 3, header: "Bus Route", label: "Bus Route", show: true, column: true, blank: { mode: "nothing" }, notOnList: { mode: "nothing" } }], picks: {}, headers: headers }) };
  g.od_loadLists_ = () => [list];
  g.od_readTable_ = () => ({ headers: headers, rows: rows });
  g.od_writeToolColumns_ = () => {}; g.od_saveLists_ = () => {};
  let out = g.od_buildOutputs_(ss, master, guardians);
  eq(out.listChanges, [], "first build: only the copy");
  assert.strictEqual(g.backendReadBlock_(ss, "LIST_WATCH").length, 4, "header + the copied mark + 2 students");
  rows = [["123456789", "Bart Simpson", "R14"], ["111222333", "Gene Belcher", "R7"]];
  out = g.od_buildOutputs_(ss, master, guardians);
  eq(out.listChanges, [{ list: "🚌 Bus Routes", added: ["Gene Belcher"], removed: ["Lisa Simpson"], changed: ["Bart Simpson: Bus Route R12 → R14"] }]);
  out = g.od_buildOutputs_(ss, master, guardians);
  eq(out.listChanges, [], "the build after: nothing new");
  // Untick it: its copy goes at the next build
  list.setup.watch = false;
  g.od_buildOutputs_(ss, master, guardians);
  assert.strictEqual(g.backendReadBlock_(ss, "LIST_WATCH").length, 1, "only the header left");
  // A list that never watched writes nothing at all
  const ss2 = makeSpreadsheet(["Contact Log", "Backend"]);
  const g2 = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss2) } });
  Object.assign(g2, { od_loadLists_: () => [Object.assign({}, list, { setup: g2.od_fillSetup_(Object.assign({}, list.setup, { watch: false })) })],
    od_readTable_: g.od_readTable_, od_writeToolColumns_: () => {}, od_saveLists_: () => {} });
  g2.od_buildOutputs_(ss2, master, guardians);
  assert.strictEqual(g2.backendReadBlock_(ss2, "LIST_WATCH").length, 0, "no section written");
  // The report carries them, and the panel shows them
  const build = fs.readFileSync(path.join(ROOT, 'Build.gs'), 'utf8');
  assert(/listChanges: \(listOutputs && listOutputs\.listChanges\) \|\| \[\]/.test(build), "in the build report");
  const panel = fs.readFileSync(path.join(ROOT, 'ProgressSidebar.html'), 'utf8');
  assert(/openBuildDetails\('list:\$\{i\}'\)/.test(panel) && /New on the list:/.test(panel), "the Build panel lists them");
  const sidebar = fs.readFileSync(path.join(ROOT, 'YourListsSidebar.html'), 'utf8');
  assert(/id="watchBox"/.test(sidebar) && /watch: purpose === 'contacts' \? false : !!setup\.watch/.test(sidebar), "the tick in step 2");
  console.log("6 tell me what changed: ok");
}

// ---------------------------------------------------------------- 1. 🔍 Find a Family
function familySheet() {
  const ss = makeSpreadsheet(["Contact Log", "Student List", "Parents List", "Combined Contact Tracking", "Backend"]);
  const app = fakeSpreadsheetApp(ss);
  const ui = app.getUi();
  ui._shown = [];
  ui.showSidebar = h => ui._shown.push(["panel", h._title]);
  ui.showModelessDialog = (h, t) => ui._shown.push(["popup", t]);
  const g = loadProject({ extra: { SpreadsheetApp: app } });
  const user = require('./lib/load').propStore();
  g.PropertiesService.getUserProperties = () => user;
  g.htmlFromFile_ = name => { const h = { _name: name, setTitle: t => { h._title = t; return h; }, setWidth: () => h, setHeight: () => h }; return h; };
  return { ss, g, ui, user };
}
{
  let { ss, g, ui, user } = familySheet();
  // Before the first build: a note, no data
  eq([g.getFindFamilyData().notBuilt, g.getFindFamilyData().as], [true, "popup"]);
  ss.getSheetByName("Student List")._put([["OSIS", "Student", "Guardian", "Site", "Class", "Sibling", "Status", "Last Interaction", "Bus Route"],
    ["123456789", "Bart Simpson", "Homer Simpson and Marge Simpson", "Main", "Y41", "567891234 - Lisa Simpson", "Current", "11/16/2026", "R12"],
    ["567891234", "Lisa Simpson", "Homer Simpson and Marge Simpson", "Main", "Y25", "123456789 - Bart Simpson", "Current", "None logged", ""],
    ["111, 222", "Group row", "", "", "", "", "", "", ""]], 1, 1);
  ss.getSheetByName("Parents List")._put([["student_id", "Guardian output", "Student", "Class / Site", "relationship", "language", "email", "mobile", "Sibling", "Parents Square", "NYCSA", "Last Interaction"],
    ["123456789", "Marge Simpson", "Bart Simpson", "Y41", "Mother", "English", "marge@x.org", "(718) 555-0102, (917) 555-0144", "", "Yes", "Yes/Unknown", "11/16/2026"],
    ["123456789", "Homer Simpson", "Bart Simpson", "Y41", "Father", "English", "", "(718) 555-0101", "", "No", "No (NG) K7Q2", ""]], 1, 1);
  let d = g.getFindFamilyData();
  eq(d.students.map(s => s.osis + " " + s.name + " " + s.cls + " " + s.last), ["123456789 Bart Simpson Y41 11/16/2026", "567891234 Lisa Simpson Y25 None logged"], "students (no sibling-group rows)");
  eq(d.parents[0].phones, ["(718) 555-0102", "(917) 555-0144"], "every number of a parent");
  eq([d.parents[1].ps, d.parents[1].nycsa], ["No", "No (NG) K7Q2"]);

  // Popup by default; the switch opens the other one with the same search, remembered per person
  g.openFindFamily();
  eq(ui._shown.pop(), ["popup", "Find a Family"]);
  assert.strictEqual(g.findFamilySwitch("panel", { q: "simpsn", osis: "123456789" }), "panel");
  eq(ui._shown.pop(), ["panel", "Find a Family"]);
  d = g.getFindFamilyData();
  eq([d.as, d.keep], ["panel", { q: "simpsn", osis: "123456789" }], "the side panel gets the search");
  assert.strictEqual(g.getFindFamilyData().keep, null, "read once");
  g.openFindFamily();
  eq(ui._shown.pop()[0], "panel", "the next time it opens as you left it");

  // The family's details: notes (no divider lines), open follow-ups (newest first), Student Info
  const cct = ss.getSheetByName("Combined Contact Tracking");
  cct._put([["OSIS", "Student", "Guardian", "Communications"], ["567891234", "Lisa", "", "x"],
            ["123456789", "Bart", "", "────\nMon 11/16 - {Call - Homer} start of year\n────\nThu 11/12 - {Email - Marge} teacher email\n────"]], 1, 1);
  const realGetRange = cct.getRange;
  cct.getRange = function (a, b, c, d2) {
    const r = realGetRange.apply(this, arguments);
    if (a === 1 && b === 1 && d2 === 1) {
      r.createTextFinder = text => { const f = { matchEntireCell: () => f, findNext: () => {
        for (let i = 1; i <= c; i++) if (String(cct._get(i, 1)) === String(text)) return realGetRange.call(cct, i, 1);
        return null; } }; return f; };
    }
    return r;
  };
  const log = ss.getSheetByName("Contact Log");
  const L = (b, osis, n, m, o) => ["", b, "Homer Simpson and Marge Simpson", "Marge Simpson", osis, "Bart Simpson", 1, "", "", "Phone", "Call", "Type", m, n, o, false];
  log._put([L("11/10/2026", "123456789", "Followup Needed", "first open", ""), L("11/11/2026", "123456789", "No Followup Needed", "closed", ""),
            L("11/12/2026", "123456789", "Followup Needed - ASAP", "teacher email", "call back"), L("11/12/2026", "567891234", "Followup Needed", "Lisa's", "")], 2, 1);
  log._put([["END"]], 8, 5);
  g.backendWriteBlock_(ss, "ADDL_INFO", [["OSIS", "Student", "Guardian", "Additional Information"], ["123456789", "Bart Simpson", "", "Bus Route: R12"], ["123456789", "Bart Simpson", "Marge Simpson", "PTA: Chair"]]);
  const det = g.getFindFamilyDetails("123456789");
  assert(!/─/.test(det.notes) && /start of year/.test(det.notes) && /teacher email/.test(det.notes), "notes without the divider lines: " + JSON.stringify(det.notes));
  eq(det.followUps.map(u => u.status + " " + u.text), ["Followup Needed - ASAP teacher email", "Followup Needed first open"], "open ones, newest first, this student only");
  eq(det.info, [{ who: "", line: "Bus Route: R12" }, { who: "Marge Simpson", line: "PTA: Chair" }]);
  eq(g.getFindFamilyDetails("nonsense"), { notes: "", followUps: [], info: [] });

  // No "Log a contact" yet: logging a contact from the card comes with the logging UI (a future update, owner 2026-11-23)
  assert.strictEqual(typeof g.findFamilyLogContact, "undefined", "no Log a contact endpoint");
  assert(!/Log a contact/.test(fs.readFileSync(path.join(ROOT, 'FindFamily.html'), 'utf8')), "no Log a contact button");

  // Edit student: New / Edit Student with the student filled in (not "back to Student Overrides")
  let opened = 0;
  g.openNewStudentDialog = () => { opened++; };
  user.setProperty("OVERRIDES_FROM_HOME", "1");
  g.findFamilyEditStudent("123456789");
  eq([opened, user.getProperty("PREFILL_OSIS"), user.getProperty("OVERRIDES_FROM_HOME")], [1, "123456789", null]);
  // In the full menu
  const menus = fs.readFileSync(path.join(ROOT, 'Menus.gs'), 'utf8');
  assert(/menu\.addItem\("🔍 Find a Family", "openFindFamily"\)/.test(menus), "🚀 App Menu → 🔍 Find a Family");
  console.log("1 find a family (server): ok");
}

// ---------------------------------------------------------------- Student List: Home Language from ATS (owner 2026-11-23)
{
  const { newTracker, parseCsv: readCsv, csvText } = require('./lib/tracker');
  const t = newTracker();
  t.idmUpload("1 - Student Data (ATS).csv");
  t.g.idm_exit(); t.build();
  const g = t.g;
  const sl = () => t.tab("Student List");
  const langOf = osis => { const tab = sl(); const c = tab[0].indexOf("Home Language"); const row = tab.find(r => String(r[0]) === osis); return row ? row[c] : undefined; };
  const headers = sl()[0];
  assert.deepStrictEqual(headers.slice(0, 9), ["OSIS", "Student", "Guardian", "Site", "Class", "Sibling", "Status", "Last Interaction", "Home Language"],
    "Home Language right after Last Interaction; no other column moved: " + JSON.stringify(headers));
  const ats = readCsv(csvText("1 - Student Data (ATS).csv")).slice(1).filter(r => r[2]);
  ats.forEach(r => assert.strictEqual(langOf(r[2]), r[18], r[4] + " " + r[3] + ": ATS's Home Language"));
  assert.strictEqual(langOf("989123456"), "Chinese");
  // The autofill's Backend copy (BX:CD) is still OSIS..Status
  assert.deepStrictEqual(g.backendReadBlock_(t.ss, "STUDENT_LIST")[0], ["OSIS", "Student", "Guardian", "Site", "Class", "Sibling", "Status"]);
  // An all-caps export reads like the names do
  assert.strictEqual(g.buildStudentMap_([[ "", "", "111222333", "SMITH", "ANN", "", "", "", "", "", "", "", "", "", "", "", "", "X01", "HAITIAN CREOLE" ]], new Map()).get("111222333").homeLanguage, "Haitian Creole");

  // New / Edit Student: a language typed there wins; a New Student (not in ATS) shows its own
  const pay = g.getAllStudentsForEdit().find(x => x.osis === "123456789");
  g.saveStudentOverrides([{ entryType: "Edit Existing Student", osis: pay.osis, classCode: pay.classCode, firstName: pay.firstName, lastName: pay.lastName,
    label: pay.label || "Current", hideFromReport: false, grade: pay.grade, age: pay.age, gender: pay.gender, homeLanguage: "Spanish",
    labelAtOpen: pay.label || "Current", prefilled: null, guardians: pay.guardians.map(x => ({ name: x.name, relationship: x.relationship, phone: x.phone, email: x.email, language: x.language })) },
    { entryType: "New Student", osis: "111222333", classCode: "Y41", firstName: "Ann", lastName: "Smith", label: "", hideFromReport: false,
      grade: "", age: "", gender: "", homeLanguage: "haitian creole", labelAtOpen: "", prefilled: null,
      guardians: [{ name: "Joe Smith", relationship: "Father", phone: "7185550199", email: "", language: "" }] }]);
  t.build();
  assert.strictEqual(langOf("123456789"), "Spanish", "Bart: the language typed in New / Edit Student");
  assert.strictEqual(langOf("111222333"), "Haitian Creole", "a New Student shows the one typed for them");
  assert.strictEqual(langOf("567891234"), "English", "Lisa: still ATS's");

  // Next month: a student who left ATS keeps the language the Student List had
  const before = new Set(ats.map(r => r[2]));
  const next = readCsv(csvText("1b - Student Data (ATS) - next month (2 new, 1 left).csv")).slice(1).filter(r => r[2]);
  const gone = Array.from(before).filter(o => !next.some(r => r[2] === o));
  assert.strictEqual(gone.length, 1, "one student left");
  const goneLang = langOf(gone[0]);
  assert(goneLang, "had a language");
  t.idmUpload("1b - Student Data (ATS) - next month (2 new, 1 left).csv");
  t.g.idm_exit(); t.build();
  const goneRow = sl().find(r => String(r[0]) === gone[0]);
  assert(/discharged/i.test(goneRow[6]), "the student who left is Discharged");
  assert.strictEqual(langOf(gone[0]), goneLang, "and keeps their Home Language");
  next.filter(r => !before.has(r[2])).forEach(r => assert.strictEqual(langOf(r[2]), r[18], "a new student: ATS's"));

  // Your Lists columns come after it
  const r = g.od_createList("Bus Routes", "🚌");
  const v = g.od_uploadCsv(csvText("5a - Your Lists - Bus Routes (student list).csv"));
  g.od_saveList({ by: v.by, roles: v.roles, vars: v.vars.map(x => Object.assign({}, x, { show: true, column: x.header === "Route" })), view: v.view, purpose: 'info', watch: false, decisions: {} });
  t.build();
  const h2 = sl()[0];
  assert.deepStrictEqual(h2.slice(7, 10), ["Last Interaction", "Home Language", "Route"], "a list column follows Home Language: " + JSON.stringify(h2));
  console.log("Student List: Home Language from ATS (New / Edit Student wins, kept after leaving): ok");
}

// ---------------------------------------------------------------- Contact Log: an open Follow-up with a newer contact is orange (owner 2026-11-24)
{
  const ss = makeSpreadsheet(["Contact Log", "Backend"]);
  const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) } });
  const cl = ss.getSheetByName("Contact Log");
  const L = g.__get("SHEET_LOOK_");
  // The formula as checked with a formula engine on sample rows (newer contact below / above, an older one logged
  // later, same-day pairs, a long time fraction, text OSIS, the lookup box under END): change it only with a new check
  const VERIFIED = '=IF(OR(TRIM($N2)="Followup Needed", TRIM($N2)="Followup Needed - ASAP", TRIM($N2)="Waiting for Response", ' +
    'TRIM($N2)="Need to respond", TRIM($N2)="Waiting - Check Up"), IF(AND($E2<>"", ISNUMBER($B2)), COUNTIFS($E$2:$E, $E2, $B$2:$B, ' +
    '">"&($B2+1/86400)) + COUNTIFS($E3:$E, $E2, $B3:$B, ">="&($B2-1/86400)) > 0))';
  assert.strictEqual(g.__get("FOLLOWUP_NEWER_FORMULA_"), VERIFIED, "the checked formula");
  // Open = what the hide rule keeps showing (Regenerate.gs getContactLogHideFormula_)
  const hide = g.getContactLogHideFormula_();
  g.__get("FOLLOWUP_OPEN_VALUES_").forEach(v => assert(hide.toLowerCase().indexOf('trim($n2)<>"' + v.toLowerCase() + '"') !== -1, v + " is open for the hide rule too"));
  const isOrange = r => r._st.formula === VERIFIED;

  // The 10 PM redraw draws it once, on N2:N, before the Workshop color (so it wins on a Workshop row)
  g.resetContactLogFormatting();
  const all = cl._cf;
  const rules = all.filter(isOrange);
  assert.strictEqual(rules.length, 1, "drawn once");
  assert.strictEqual(rules[0]._st.bg, L.followupNewer, "orange");
  assert.strictEqual(L.followupNewer, "#f6b26b");
  const rg = rules[0]._st.ranges[0];
  eq([rg.getRow(), rg.getColumn(), rg.getNumColumns(), rg.getLastRow()], [2, 14, 1, cl.getMaxRows()], "on N2:N (to the bottom)");
  const workshop = all.findIndex(r => r._st.formula === '=$J2="Workshop"');
  assert(all.indexOf(rules[0]) < workshop, "before the Workshop color");
  assert(all.some(r => r._st.formula === '=AND($B2<>"", $J2="")'), "the red Method rule is still there");

  // A build adds it when missing (first), once; a recolored one is kept; no Contact Log = nothing
  cl._cf = cl._cf.filter(r => !isOrange(r));
  const others = cl._cf.length, writes = cl._cfWrites;
  assert.strictEqual(g.ensureFollowupNewerRule_(ss), true, "added when missing");
  assert(isOrange(cl._cf[0]), "first in the list");
  assert.strictEqual(cl._cf.length, others + 1, "no other rule lost");
  assert.strictEqual(g.ensureFollowupNewerRule_(ss), false, "the next build: nothing to do");
  assert.strictEqual(cl._cfWrites, writes + 1, "one write in all");
  cl._cf[0]._st.bg = "#ff9900";
  assert.strictEqual(g.ensureFollowupNewerRule_(ss), false, "a recolored rule is kept");
  const g2 = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(makeSpreadsheet(["Backend"])) } });
  assert.strictEqual(g2.ensureFollowupNewerRule_(g2.SpreadsheetApp.getActiveSpreadsheet()), false);
  const build = fs.readFileSync(path.join(ROOT, 'Build.gs'), 'utf8');
  assert(/try \{ ensureFollowupNewerRule_\(ss\); \} catch/.test(build), "every build checks it");
  console.log("Contact Log: open Follow-up with a newer contact is orange: ok");
}
