// The three testers' findings (release check round 3, 2026-11-23): each one reproduced here and fixed.
const fs = require('fs'), path = require('path'), assert = require('assert');
const { loadProject, ROOT } = require('./lib/load');
const { makeSpreadsheet, fakeSpreadsheetApp } = require('./lib/fakesheet');
const eq = (a, b, m) => assert.deepStrictEqual(JSON.parse(JSON.stringify(a === undefined ? null : a)), JSON.parse(JSON.stringify(b === undefined ? null : b)), m);

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

// A school with the test files: ATS + Parent Square 2c (new numbers) + the standard NYCSA (which still has ATS's numbers)
function tracker(psFile) {
  const ss = makeSpreadsheet(["Contact Log", "Raw Data", "Student List", "Backend"]);
  const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) } });
  g.od_loadLists_ = () => [];
  let builds = 0;
  g.runBuildOnly = () => { builds++; };
  ss.getSheetByName("Backend")._put([g.__get("OVERRIDES_HEADERS_")], 2, 7);
  const raw = ss.getSheetByName("Raw Data");
  raw._put(csv("1 - Student Data (ATS).csv"), 7, 1);
  raw._put(csv(psFile || "2c - Parent Square - new phone numbers.csv"), 7, 33);
  raw._put(csv("3 - NYCSA (Students with Non-NYCSA Account).csv"), 7, 42);
  const numbers = (id, first) => {
    const ov = g.getOverrideData_(ss);
    const gr = g.buildGuardianObjects_(raw.getDataRange().getValues(), ov.guardians, ov.phones).find(x => x.student_id === id && x.guardianOutput.indexOf(first) === 0);
    return gr ? gr.phones : null;
  };
  const cards = () => g.findPhonesToCheck_(ss).items.map(it => it.guardian + " " + it.phone + " for " + it.students.map(s => s.name).join(" & "));
  const atsRowOf = osis => raw.getDataRange().getValues().findIndex((r, i) => i >= 6 && String(r[2]) === osis) + 1; // sheet row
  return { ss, g, raw, numbers, cards, atsRowOf, builds: () => builds };
}
const homerKey = t => t.g.findPhonesToCheck_(t.ss).items.find(it => it.guardian === "Homer Simpson").key;

{
  // "Use it instead" brought the number it replaced right back (NYCSA still has ATS's number): never again
  const t = tracker();
  eq(t.cards(), ["Morticia Addams (333) 555-0303 for Wednesday Addams", "Homer Simpson (111) 555-0101 for Bart Simpson & Lisa Simpson"]);
  t.g.savePhoneNumbers([homerKey(t)], "instead");
  eq(t.numbers("123456789", "Homer"), ["(111) 555-0101"], "used instead");
  eq(t.cards(), ["Morticia Addams (333) 555-0303 for Wednesday Addams"], "Homer's old number is not offered back (NYCSA still has it)");
  eq(t.g.backendReadBlock_(t.ss, "PHONES").slice(1).map(r => r[3]), ["instead of (111) 111-1111", "instead of (111) 111-1111"],
    "the row notes the ATS number it replaced, per student");
  // An edited phone in New / Edit Student: the number it replaced isn't offered either (ATS's own number is known)
  const t2 = tracker("2 - Parent Square.csv");
  const pay = JSON.parse(JSON.stringify(t2.g.getAllStudentsForEdit().find(x => x.osis === "234567891")));
  const form = { entryType: "Edit Existing Student", osis: pay.osis, classCode: pay.classCode, firstName: pay.firstName, lastName: pay.lastName,
    label: pay.label, hideFromReport: false, grade: pay.grade, age: pay.age, gender: pay.gender, homeLanguage: pay.homeLanguage,
    labelAtOpen: pay.label || "Current", prefilled: null, guardians: pay.guardians };
  form.guardians[0].phone = "(646) 555-1212"; // Morticia's phone corrected by hand
  t2.g.saveStudentOverrides([form]);
  assert(!t2.cards().some(c => /Morticia Addams \(333\) 333-3333/.test(c)), "the number just replaced by hand is not offered back: " + JSON.stringify(t2.cards()));
  console.log("Use it instead / an edited phone: the old number isn't offered back: ok");
}
{
  // Ticking two numbers of one parent and "Use the 2 instead" kept only the last one
  const t = tracker();
  const r = t.raw.getDataRange().getValues();
  r.forEach((row, i) => { if (i >= 6 && row[32] === "Simpson, Homer") t.raw._put([["(111) 555-0202"]], i + 1, 37); }); // a Secondary Phone too
  const homer = t.g.findPhonesToCheck_(t.ss).items.filter(it => it.guardian === "Homer Simpson");
  eq(homer.map(it => it.phone), ["(111) 555-0101", "(111) 555-0202"], "two cards for Homer");
  t.g.savePhoneNumbers(homer.map(it => it.key), "instead");
  eq(t.numbers("123456789", "Homer"), ["(111) 555-0101", "(111) 555-0202"], "both kept");
  eq(t.numbers("567891234", "Homer"), ["(111) 555-0101", "(111) 555-0202"], "for the sibling too");
  assert(!t.cards().some(c => /Homer/.test(c)), "nothing left to check for Homer");
  // A later "Use it instead" replaces what was saved before for that parent (the earlier "add" row goes)
  const t2 = tracker();
  const mort = () => t2.g.findPhonesToCheck_(t2.ss, true).items.find(it => it.guardian === "Morticia Addams");
  t2.g.savePhoneNumbers([mort().key], "add");
  eq(t2.numbers("234567891", "Morticia"), ["(333) 333-3333", "(333) 555-0303"]);
  assert.strictEqual(mort(), undefined, "a saved number is not offered again");
  const r2 = t2.raw.getDataRange().getValues();
  r2.forEach((row, i) => { if (i >= 6 && row[32] === "Adams, Morticia") t2.raw._put([["(333) 555-0303"], ], i + 1, 36); });
  r2.forEach((row, i) => { if (i >= 6 && row[32] === "Adams, Morticia") t2.raw._put([["(333) 555-0404"]], i + 1, 37); });
  // Parent Square has 555-0303 (saved, known) and 555-0404: a source with a number we have is not offered ...
  assert.strictEqual(mort(), undefined, "...so 555-0404 alone would need its own source");
  r2.forEach((row, i) => { if (i >= 6 && row[32] === "Adams, Morticia") t2.raw._put([["(333) 555-0404"]], i + 1, 36); });
  r2.forEach((row, i) => { if (i >= 6 && row[32] === "Adams, Morticia") t2.raw._put([[""]], i + 1, 37); });
  t2.g.savePhoneNumbers([mort().key], "instead");
  eq(t2.g.backendReadBlock_(t2.ss, "PHONES").slice(1).map(r => [r[2], String(r[3]).split(" ")[0]]), [["(333) 555-0404", "instead"]],
    "the earlier \"add\" row went with the new \"instead\"");
  eq(t2.numbers("234567891", "Morticia"), ["(333) 555-0404"]);
  console.log("several numbers instead / a later save: ok");
}
{
  // ATS changes the number after "Use it instead": ATS's new number shows, the stale row goes at the next cleanup
  const t = tracker();
  t.g.savePhoneNumbers([homerKey(t)], "instead");
  [t.atsRowOf("123456789"), t.atsRowOf("567891234")].forEach(row => t.raw._put([["9295557777"]], row, 22));
  eq(t.numbers("123456789", "Homer"), ["(929) 555-7777"], "ATS changed since: its new number shows (the saved one no longer applies)");
  t.g.cleanRedundantOverrides();
  eq(t.g.backendReadBlock_(t.ss, "PHONES").slice(1).length, 0, "the stale rows are gone");
  assert(t.cards().some(c => /Homer Simpson \(111\) 555-0101/.test(c)), "Parent Square's number can be checked again");
  // ATS lists the saved number among others: the row goes (ATS has it), wherever it is in the cell
  const t2 = tracker();
  t2.g.savePhoneNumbers([homerKey(t2)], "instead");
  [t2.atsRowOf("123456789"), t2.atsRowOf("567891234")].forEach(row => t2.raw._put([["929-555-7777 / 111-555-0101"]], row, 22));
  t2.g.cleanRedundantOverrides();
  eq(t2.g.backendReadBlock_(t2.ss, "PHONES").slice(1).length, 0, "ATS has it now (2nd in the cell)");
  eq(t2.numbers("123456789", "Homer"), ["(929) 555-7777", "(111) 555-0101"]);
  console.log("ATS changes or catches up after Use it instead: ok");
}
{
  // ATS spells the parent another way next month ("HOMER J SIMPSON"): the saved number still applies
  const t = tracker();
  t.g.savePhoneNumbers([homerKey(t)], "add");
  [t.atsRowOf("123456789"), t.atsRowOf("567891234")].forEach(row => t.raw._put([["HOMER J SIMPSON"]], row, 20));
  eq(t.numbers("123456789", "Homer"), ["(111) 111-1111", "(111) 555-0101"], "still both");
  assert(!t.cards().some(c => /Homer/.test(c)), "and it isn't offered again");
  [t.atsRowOf("123456789"), t.atsRowOf("567891234")].forEach(row => t.raw._put([["1111111111 / 1115550101"]], row, 22));
  t.g.cleanRedundantOverrides();
  eq(t.g.backendReadBlock_(t.ss, "PHONES").slice(1).length, 0, "cleaned up once ATS has it, whatever the spelling");
  console.log("a parent spelled another way: ok");
}
{
  // Adding a parent later wiped the student's Status label (the Edit row became an Added Guardian row)
  const t = tracker("2 - Parent Square.csv");
  const formOf = osis => {
    const s = JSON.parse(JSON.stringify(t.g.getAllStudentsForEdit().find(x => x.osis === osis)));
    return { entryType: "Edit Existing Student", osis: s.osis, classCode: s.classCode, firstName: s.firstName, lastName: s.lastName,
      label: s.label, hideFromReport: false, grade: s.grade, age: s.age, gender: s.gender, homeLanguage: s.homeLanguage,
      labelAtOpen: s.label || "Current", prefilled: null, guardians: s.guardians };
  };
  let f = formOf("123456789");
  f.label = "Moved - new address pending";
  f.guardians.push({ name: "Abe Simpson", relationship: "Grandfather", phone: "3434343434", email: "", language: "English" });
  t.g.saveStudentOverrides([f]);
  f = formOf("123456789");
  assert.strictEqual(f.label, "Moved - new address pending", "the label shows when the student is opened again");
  f.guardians.push({ name: "Mona Simpson", relationship: "Grandmother", phone: "", email: "", language: "English" });
  t.g.saveStudentOverrides([f]);
  const rows = t.g.backendReadBlock_(t.ss, "OVERRIDES").slice(1).filter(r => String(r[2]).trim());
  eq(rows.map(r => [r[1], r[40]]), [["Edit Existing Student", "Moved - new address pending"]], "still an Edit row, the label kept");
  console.log("adding a parent keeps an Edit row's label: ok");
}
{
  // "Not new" / "Not needed" at a big school: 500 keys all stick (one property held ~8.5 KB, the oldest came back)
  const t = tracker();
  const keys = Array.from({ length: 500 }, (_, i) => "234567891|garcia maria" + i);
  t.g.skipParentsToAdd(keys.slice(0, 250));
  t.g.skipParentsToAdd(keys.slice(250));
  eq(t.g.parentsSkipped_().length, 500, "all 500 kept");
  assert(Object.keys(t.g.__props).filter(k => /^PARENTS_TO_ADD_SKIPPED/.test(k)).length >= 2, "over several properties");
  assert(Object.keys(t.g.__props).filter(k => /^PARENTS_TO_ADD_SKIPPED/.test(k)).every(k => t.g.__props[k].length <= 8500), "each under the limit");
  t.g.skipParentsToAdd([keys[0]]); // again: moved to the end, not doubled
  eq(t.g.parentsSkipped_().length, 500);
  t.g.unskipParentsToAdd();
  eq([t.g.parentsSkipped_().length, Object.keys(t.g.__props).filter(k => /^PARENTS_TO_ADD_SKIPPED/.test(k)).length], [0, 0], "Show them again clears every piece");
  // An older sheet's single property is still read
  t.g.__props.PHONES_TO_CHECK_SKIPPED = JSON.stringify(["homer simpson|1115550101"]);
  eq(t.g.findPhonesToCheck_(t.ss).skipped, 1, "the old one-property list still counts");
  console.log("skip lists over several properties: ok");
}
{
  // "Check and edit before adding" with the phone changed: the source's number isn't offered right back
  const t = tracker("2 - Parent Square.csv");
  t.raw._put(csv("3b - NYCSA - a third guardian.csv"), 7, 42);
  const patty = t.g.findParentsToAdd_(t.ss).items.find(it => it.name === "Patty Bouvier");
  assert(patty && patty.phone === "(454) 545-4545", "Patty to add, with NYCSA's number");
  const s0 = JSON.parse(JSON.stringify(t.g.getAllStudentsForEdit().find(x => x.osis === "123456789")));
  const form = { entryType: "Edit Existing Student", osis: s0.osis, classCode: s0.classCode, firstName: s0.firstName, lastName: s0.lastName,
    label: s0.label, hideFromReport: false, grade: s0.grade, age: s0.age, gender: s0.gender, homeLanguage: s0.homeLanguage,
    labelAtOpen: s0.label || "Current", guardians: s0.guardians,
    prefilled: { name: "Patty Bouvier", sources: ["NYCSA"], phone: patty.phone } };
  form.guardians.push({ name: "Patty Bouvier", relationship: "Aunt", phone: "(646) 555-7000", email: "", language: "English" });
  t.g.saveStudentOverrides([form]);
  const rows = t.g.backendReadBlock_(t.ss, "OVERRIDES").slice(1).filter(r => String(r[2]).trim());
  eq(rows.map(r => r[1]), ["Added Guardian (NYCSA)"]);
  assert(!t.cards().some(c => /Patty/.test(c)), "NYCSA's number for Patty is not offered: " + JSON.stringify(t.cards()));
  eq(t.g.findPhonesToCheck_(t.ss).skipped, 1, "it is counted as \"not needed\" (Show them again brings it back)");
  console.log("check and edit with another phone: ok");
}
{
  // The page before any data: says to load it (it said everyone matches)
  const ss = makeSpreadsheet(["Contact Log", "Raw Data", "Student List", "Backend"]);
  const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) } });
  g.od_loadLists_ = () => [];
  const home = g.getOverridesHome();
  eq([home.noData, home.toAdd.length, home.phones.length], [true, 0, 0]);
  const html = fs.readFileSync(path.join(ROOT, 'ManageOverridesDialog.html'), 'utf8');
  assert(/home\.noData \? 'Load your data first/.test(html), "the page says so");
  assert(/if \(!r\.saved\) \{ say\('Nothing to save/.test(html) && /if \(!r\.removed\) \{ say\('Already deleted/.test(html), "another window's work is said plainly");
  // "% alike": a one-letter typo is a close spelling; a real other last name keeps its words
  const G = n => ({ guardianOutput: n, phones: [], mobile: "", email: "" });
  eq(g.parentMaybeSame_("Adams, Morticia", [], "", [G("Morticia Addams")]).why, "a close spelling");
  eq(g.parentMaybeSame_("Patty Simpson", [], "", [G("Patty Bouvier")]).why, "same first name, another last name");
  console.log("page messages and wording: ok");
}
{
  // Phone numbers are whole runs of digits: no made-up number from a typo or an extension (the 24 messy formats
  // the careless tester tried read as before)
  const g = loadProject();
  const read = v => g.phonesFromCell_(v);
  eq(read("27185550104"), [], "11 digits not starting with 1: left out, as before");
  eq(read("718-555-0107 ext 1234 9175550108"), ["(718) 555-0107", "(917) 555-0108"], "the extension doesn't eat the next number");
  eq(read("71855501019175550144"), ["(718) 555-0101", "(917) 555-0144"], "two typed together (20 digits) are still two");
  eq(read("7185550101917555014"), [], "19 digits: none");
  eq(read("OSIS 123456789 718-555-0101"), ["(718) 555-0101"]);
  eq(read("+1 (718) 555-0101"), ["(718) 555-0101"]);
  eq(read("1-718-555-0101 x12"), ["(718) 555-0101"]);
  eq(read("718-555-0101 / 917-555-0144 / 646-555-0111"), ["(718) 555-0101", "(917) 555-0144", "(646) 555-0111"]);
  eq(read(17185550101), ["(718) 555-0101"]);
  console.log("phone numbers as whole digit runs: ok");
}
{
  // Input Data Mode: a list with only "Email" and "Phone" is not Parent Square (it replaced the real data)
  const g = loadProject();
  const pta = parseCsv(fs.readFileSync(path.join(ROOT, 'test-data', '5b - Your Lists - PTA Members (guardian list).csv'), 'utf8'));
  const found = g.idm_findHeader_(pta);
  assert(!found || found.step.key !== "ps", "the PTA list is not taken as Parent Square: " + (found && found.step.key));
  const ps = parseCsv(fs.readFileSync(path.join(ROOT, 'test-data', '2 - Parent Square.csv'), 'utf8'));
  eq(g.idm_findHeader_(ps).step.key, "ps", "the real Parent Square export still is");
  console.log("Parent Square needs one of its own columns: ok");
}
{
  // The clean-up kept only rows that differ from ATS in names / phones / emails: a label, the class or "hide from
  // the report" alone was deleted at the next build
  const t = tracker("2 - Parent Square.csv");
  const formOf = osis => {
    const s0 = JSON.parse(JSON.stringify(t.g.getAllStudentsForEdit().find(x => x.osis === osis)));
    return { entryType: "Edit Existing Student", osis: s0.osis, classCode: s0.classCode, firstName: s0.firstName, lastName: s0.lastName,
      label: s0.label, hideFromReport: false, grade: s0.grade, age: s0.age, gender: s0.gender, homeLanguage: s0.homeLanguage,
      labelAtOpen: s0.label || "Current", prefilled: null, guardians: s0.guardians };
  };
  let f = formOf("123456789"); f.label = "Moved"; t.g.saveStudentOverrides([f]);
  f = formOf("567891234"); f.classCode = "Y99"; t.g.saveStudentOverrides([f]);
  f = formOf("234567891"); f.hideFromReport = true; t.g.saveStudentOverrides([f]);
  f = formOf("245678912"); t.g.saveStudentOverrides([f]); // saved with nothing changed: ATS has it all
  t.g.cleanRedundantOverrides();
  eq(t.g.backendReadBlock_(t.ss, "OVERRIDES").slice(1).filter(r => String(r[2]).trim()).map(r => String(r[2])).sort(),
    ["123456789", "234567891", "567891234"], "the label, the class and the hidden row stay; only the unchanged one goes");
  console.log("the clean-up keeps a label or a field change: ok");
}
{
  // Find a Family: a phone with +1 / 1- or an extension is found; more than 20 results say how many more; the OSIS
  // goes through data-osis (an OSIS with ' broke the click), in New / Edit Student's list too
  const ff = fs.readFileSync(path.join(ROOT, 'FindFamily.html'), 'utf8');
  assert(/digitsOnly && qd\.length === 11 && qd\.charAt\(0\) === '1'/.test(ff) && /function dropExtension/.test(ff), "country code and extension");
  assert(/var SHOW_RESULTS = 20/.test(ff) && /more families'\) \+ ' match/.test(ff), "how many more");
  assert(/onclick="pick\(this\.dataset\.osis\)"/.test(ff) && !/pick\(\\'' \+ esc/.test(ff), "pick by data-osis");
  assert(/busy = true;\s*google\.script\.run\.withSuccessHandler/.test(ff), "Edit student once per click");
  const ns = fs.readFileSync(path.join(ROOT, 'NewStudentDialog.html'), 'utf8');
  assert(/onclick="selectSearchResult\(this\.dataset\.osis\)"/.test(ns) && /\$\{esc\(s\.lastName\)\}/.test(ns), "New / Edit Student's list");
  console.log("Find a Family / New Student search fixes: ok");
}
console.log("tester findings: ok");

// ---------------------------------------------------------------- Your Lists: "Paste new version" keeps the saved one until Save
{
  const { newTracker, parseCsv: readCsv } = require('./lib/tracker');
  const t = newTracker();
  t.idmUpload("1 - Student Data (ATS).csv", "2 - Parent Square.csv", "3 - NYCSA (Students with Non-NYCSA Account).csv");
  t.g.idm_exit(); t.build();
  const g = t.g, ss = t.ss;
  const busCsv = t.csvText("5a - Your Lists - Bus Routes (student list).csv");
  const r = g.od_createList("Bus Routes", "🚌");
  const v = g.od_uploadCsv(busCsv);
  const vars = v.vars.map(x => Object.assign({}, x, { show: x.header !== 'AM Pickup', column: x.header !== 'AM Pickup',
    notOnList: x.header === 'Route' ? { mode: 'text', text: 'No bus' } : { mode: 'nothing', text: '' } }));
  g.od_saveList({ by: v.by, roles: v.roles, vars: vars, view: v.view, purpose: 'info', watch: true, decisions: {} });
  t.build(); t.build(); // the first build copies the list for "Tell me what changed", the second compares
  const report = () => JSON.parse(g.CacheService.getScriptCache().get('BUILD_REPORT_LOGS') || '{}').listChanges || [];
  const routes = () => { const sl = t.tab("Student List"); const c = sl[0].indexOf("Route"); return sl.slice(1).map(x => x[1] + "=" + x[c]).join("; "); };
  const tab = () => ss.getSheets().find(s => s.getSheetId() === r.tabId);
  const oldTabs = () => ss.getSheets().filter(s => / \(old version\)$/.test(s.getName()));
  const cells = () => JSON.stringify(tab().getDataRange().getValues());
  const setupOf = () => g.od_loadLists_(ss).find(l => l.tabId === r.tabId).setup;
  const routesBefore = routes(), tabBefore = cells(), addlBefore = JSON.stringify(g.backendReadBlock_(ss, "ADDL_INFO"));
  assert(/Bart Simpson=R12/.test(routesBefore) && /Wednesday Addams=R7/.test(routesBefore), routesBefore);

  // The click: the saved version goes to a hidden tab, the list's tab waits for the paste
  g.od_repasteList(r.tabId);
  eq(oldTabs().map(s => [s.getName(), s.isSheetHidden()]), [["🚌 Bus Routes (old version)", true]], "the saved version is kept, hidden");
  assert.strictEqual(tab().getRange(4, 1).getValue(), "PASTE HERE", "the list's tab waits for the paste");
  assert.strictEqual(g.od_listSummaries_(ss)[0].waiting, true, "the step 5 card says a new version isn't saved");
  assert.strictEqual(g.od_getWizard().list.keptOld, true, "the paste view says the old one stays in use");
  // A build before the paste (10 PM, an override save) uses the saved version and writes nothing on the tab
  t.build();
  assert.strictEqual(routes(), routesBefore, "the Student List keeps every route (it said \"No bus\" for everyone)");
  assert.strictEqual(JSON.stringify(g.backendReadBlock_(ss, "ADDL_INFO")), addlBefore, "Additional Information (phone notes) unchanged");
  eq(report(), [], "\"Tell me what changed\" reports nothing (it said all 10 dropped off)");
  assert.strictEqual(tab().getRange(4, 1).getValue(), "PASTE HERE", "no ✔ Matched columns written on the waiting tab");
  assert(!JSON.stringify(tab().getDataRange().getValues()).includes("✔ Matched"));
  // Back without pasting: the saved version is back, as it was
  g.od_backToInputMode();
  assert.strictEqual(cells(), tabBefore, "Back puts the list back on its tab, cell for cell");
  eq(oldTabs().length, 0, "the hidden copy is gone");
  assert.strictEqual(setupOf().oldVersion, undefined);
  assert.strictEqual(g.od_listSummaries_(ss)[0].waiting, false);

  // Two clicks keep the FIRST copy (not the emptied tab)
  g.od_repasteList(r.tabId); g.od_repasteList(r.tabId);
  eq(oldTabs().length, 1, "one hidden copy");
  g.od_backToInputMode();
  assert.strictEqual(cells(), tabBefore, "the list comes back after two clicks too");

  // A new version read but not saved: builds still use the saved one; Edit setup drops the unsaved one
  const newCsv = busCsv.replace("Bart Simpson,123456789,R12", "Bart Simpson,123456789,R14");
  g.od_repasteList(r.tabId);
  assert.strictEqual(g.od_uploadCsv(newCsv).phase, "setup");
  t.build();
  assert(/Bart Simpson=R12/.test(routes()), "not saved yet: the build uses the saved version");
  g.od_openList(r.tabId);
  assert.strictEqual(cells(), tabBefore, "Edit setup puts the saved version back");
  eq(oldTabs().length, 0);

  // ... and so does the end of Input Data Mode
  g.idm_open(); g.idm_goToStep(4);
  g.od_repasteList(r.tabId);
  g.od_uploadCsv(newCsv);
  g.idm_exit();
  assert.strictEqual(cells(), tabBefore, "leaving Input Data Mode puts the saved version back");
  eq(oldTabs().length, 0);

  // Saved: the new version is the list now (the copy goes), and the build reports the real change
  g.od_repasteList(r.tabId);
  const nv = g.od_uploadCsv(newCsv);
  assert.strictEqual(nv.kept, true, "same columns: the setup was kept");
  eq(nv.missing, [], "no kept column is missing");
  g.od_saveList({ by: nv.by, roles: nv.roles, vars: nv.vars, view: nv.view, purpose: 'info', watch: true, decisions: {} });
  eq(oldTabs().length, 0, "Save drops the hidden copy");
  t.build();
  assert(/Bart Simpson=R14/.test(routes()), "the saved new version is used");
  const ch = report();
  assert(ch.length === 1 && ch[0].changed.some(c => /Bart Simpson: Route R12 → R14/.test(c)) && !ch[0].removed.length,
    "only the real change is reported: " + JSON.stringify(ch));

  // Deleting a list while a new version waits deletes its hidden copy too
  g.od_repasteList(r.tabId);
  g.od_deleteList(r.tabId);
  eq(oldTabs().length, 0, "no hidden copy left behind");
  assert(!tab(), "the list's tab is gone");

  // A renamed kept column is named (it said "Same columns as last time" and dropped it without a word)
  const r2 = g.od_createList("Bus Routes 2", "🚌");
  const v2 = g.od_uploadCsv(busCsv);
  const vars2 = v2.vars.map(x => Object.assign({}, x, { show: true, column: true }));
  g.od_saveList({ by: v2.by, roles: v2.roles, vars: vars2, view: v2.view, purpose: 'info', watch: false, decisions: {} });
  g.od_repasteList(r2.tabId);
  let nv2 = g.od_uploadCsv(busCsv.replace("Student,OSIS,Route,Stop", "Student,OSIS,Route #,Stop"));
  assert.strictEqual(nv2.kept, true, "the columns that find the student are the same");
  eq(nv2.missing, ["Route"], "Route is named as missing");
  assert.strictEqual(nv2.changed, false);
  g.od_repasteList(r2.tabId);
  nv2 = g.od_uploadCsv(busCsv.replace("Student,OSIS,Route,Stop", "Student,Student ID,Route,Stop"));
  assert.strictEqual(nv2.kept, false, "the OSIS column is gone: guesses again");
  assert.strictEqual(nv2.changed, true, "... and the sidebar says the columns changed");
  const ys = fs.readFileSync(path.join(ROOT, 'YourListsSidebar.html'), 'utf8');
  assert(/id="missingNote"/.test(ys) && /id="changedNote"/.test(ys) && /show\('keptNote', !!v\.kept && !missing\.length\)/.test(ys),
    "the page shows the missing columns instead of \"Same columns as last time\"");
  assert(/id="oldNotePaste"/.test(ys) && /show\('oldNotePaste', !!v\.list\.keptOld\)/.test(ys), "the paste view says the old version stays in use");
  const ids = fs.readFileSync(path.join(ROOT, 'InputDataSidebar.html'), 'utf8');
  assert(/l\.waiting \?/.test(ids), "the step 5 card shows a list waiting for its new version");
  g.od_backToInputMode();
  console.log("Your Lists: a new version doesn't replace the saved one until Save; renamed columns are named: ok");
}

// ---------------------------------------------------------------- Migration / Rollover: an unsaved new version stays behind
{
  const { newTracker } = require('./lib/tracker');
  const t = newTracker();
  t.idmUpload("1 - Student Data (ATS).csv");
  t.g.idm_exit(); t.build();
  const g = t.g, old = t.ss;
  const busCsv = t.csvText("5a - Your Lists - Bus Routes (student list).csv");
  const r = g.od_createList("Bus Routes", "🚌");
  const v = g.od_uploadCsv(busCsv);
  g.od_saveList({ by: v.by, roles: v.roles, vars: v.vars.map(x => Object.assign({}, x, { show: true, column: true })), view: v.view, purpose: 'info', watch: false, decisions: {} });
  const listTab = () => old.getSheets().find(s => s.getSheetId() === r.tabId);
  const saved = JSON.stringify(listTab().getDataRange().getValues());
  g.od_repasteList(r.tabId); // a new version started on last year's tracker, never saved
  g.od_uploadCsv(busCsv.replace("R12", "R99"));
  // Migration into a new tracker: the saved version comes over, without the unsaved one or its hidden copy
  const t2 = newTracker();
  const moved = t2.g.od_migrateLists_(old, t2.ss);
  assert.strictEqual(moved, 1);
  const lists2 = t2.g.od_loadLists_(t2.ss);
  assert.strictEqual(lists2.length, 1);
  assert.strictEqual(JSON.stringify(lists2[0].sheet.getDataRange().getValues()), saved, "the saved version came over");
  assert.strictEqual(lists2[0].sheet.getName(), "🚌 Bus Routes", "under the list's own tab name");
  assert.strictEqual(lists2[0].setup.oldVersion, undefined, "no hidden copy to look for in the new tracker");
  assert(!t2.ss.getSheets().some(s => / \(old version\)$/.test(s.getName())));
  // The Rollover archive and Parents to add read the saved version too
  const menus = fs.readFileSync(path.join(ROOT, 'Menus.gs'), 'utf8');
  const roll = menus.slice(menus.indexOf('function processRolloverBackend'));
  assert(/od_savedTab_\(ss, l\)\.copyTo\(tempSs\)/.test(roll) && /if \(copy\.isSheetHidden\(\)\) copy\.showSheet\(\)/.test(roll),
    "Rollover archives the saved version, shown (a PDF leaves hidden tabs out)");
  assert(/od_readTable_\(od_savedTab_\(ss, list\), list\.setup\)/.test(menus), "Parents to add / Phone numbers to check read the saved version");
  console.log("Your Lists: Migration / Rollover / Parents to add use the saved version: ok");
}

// ---------------------------------------------------------------- Contacts sync: a 2nd number / email change counts as Edited
{
  const g = loadProject({});
  let writes = 0;
  g.People = { People: { updateContact: (data, rn) => { writes++; return { resourceName: rn, etag: "e2" }; }, createContact: () => ({ resourceName: "people/new" }) } };
  const contact = (phones, emails) => ({ resourceName: "people/1", etag: "e1",
    names: [{ givenName: "(Mother) Martha Kent", familyName: "| Student: Clark Kent |" }],
    phoneNumbers: phones.map(v => ({ value: v, type: "Class: Y50" })), emailAddresses: emails.map(v => ({ value: v, type: "home" })),
    biographies: [{ value: "notes" }], organizations: [], memberships: [] });
  const index = { add() {} };
  const upsert = (existing, phone, extraPhones, email, extraEmails) => {
    g.findIndexedContact_ = () => existing;
    return g.searchAndUpsertSingleContact("(Mother) Martha Kent", "| Student: Clark Kent |", email, phone, "Class: Y50", "notes", "", "", null, index,
      { phones: (extraPhones || []).map(v => ({ value: v, type: "Class: Y50" })), emails: extraEmails || [] });
  };
  let res = upsert(contact(["(718) 555-0105", "(212) 555-0106"], ["MARTHA.KENT@gmail.com"]), "(718) 555-0105", ["(212) 555-0999"], "MARTHA.KENT@gmail.com");
  assert(res.wrote && res.isEdited, "only Phone 2 changed: written AND counted as edited");
  assert.strictEqual(res.logStr, "Phone: 7185550105, 2125550106 ➔ 7185550105, 2125550999");
  res = upsert(contact(["(718) 555-0105"], ["MARTHA.KENT@gmail.com"]), "(718) 555-0105", [], "MARTHA.KENT@gmail.com", ["work@x.org"]);
  assert(res.isEdited && res.logStr === "Email: martha.kent@gmail.com ➔ martha.kent@gmail.com, work@x.org", "a 2nd email: " + res.logStr);
  // One number and one email: the same log text as before the change
  res = upsert(contact(["(718) 555-0105"], ["MARTHA.KENT@gmail.com"]), "(917) 555-0144", [], "MARTHA.KENT@gmail.com");
  assert.strictEqual(res.logStr, "Phone: 7185550105 ➔ 9175550144");
  // Nothing changed: nothing written, nothing counted
  const before = writes;
  res = upsert(contact(["(718) 555-0105", "(212) 555-0106"], ["MARTHA.KENT@gmail.com"]), "(718) 555-0105", ["(212) 555-0106"], "MARTHA.KENT@gmail.com");
  assert(!res.wrote && !res.isEdited && writes === before, "no change: no write");
  console.log("Contacts sync: every number and email counts for Edited: ok");
}

// ---------------------------------------------------------------- Added parents: the siblings' source, the report's sources and phones
{
  const { newTracker } = require('./lib/tracker');
  const t = newTracker();
  t.idmUpload("1 - Student Data (ATS).csv", "2b - Parent Square - a parent not in ATS.csv", "3 - NYCSA (Students with Non-NYCSA Account).csv");
  t.g.idm_exit(); t.build();
  const g = t.g;
  const rows = () => g.backendReadBlock_(t.ss, 'OVERRIDES').slice(1).filter(r => r[2]).map(r => r[2] + " " + r[1]);
  const formFor = osis => {
    const s = g.getAllStudentsForEdit().find(x => x.osis === osis);
    return { entryType: 'Edit Existing Student', osis: s.osis, classCode: s.classCode, firstName: s.firstName, lastName: s.lastName, label: s.label || 'Current',
      hideFromReport: !!s.hideFromReport, grade: s.grade || '', age: s.age || '', gender: s.gender || '', homeLanguage: s.homeLanguage || '',
      labelAtOpen: s.label || 'Current', prefilled: null,
      guardians: s.guardians.map(x => ({ name: x.name, relationship: x.relationship, phone: x.phone, email: x.email, language: x.language })) };
  };
  // "Check and edit before adding" on Abe (Parent Square), then "Sync to all siblings": both rows name Parent Square
  const abe = g.getOverridesHome().toAdd.find(p => p.name === "Abe Simpson" && p.osis === "123456789");
  g.openNewStudentFromHome(abe.osis, { name: abe.name, relationship: abe.relationship, phone: abe.phone, email: abe.email, sources: abe.sources });
  const pre = g.getNewStudentPrefill();
  const bart = formFor(pre.osis);
  bart.guardians.push({ name: pre.guardian.name, relationship: pre.guardian.relationship, phone: pre.guardian.phone, email: pre.guardian.email, language: '' });
  bart.prefilled = { name: pre.guardian.name, sources: pre.guardian.sources, phone: pre.guardian.phone };
  const lisa = formFor("567891234");
  // what the dialog sends for a sibling now (NewStudentDialog submitForm): the same guardians AND the same prefilled
  const sib = Object.assign({}, lisa, { label: lisa.label || bart.label, hideFromReport: bart.hideFromReport, prefilled: bart.prefilled, guardians: bart.guardians });
  g.saveStudentOverrides([bart, sib]); t.build();
  eq(rows(), ["123456789 Added Guardian (Parent Square)", "567891234 Added Guardian (Parent Square)"], "the sibling's row names Parent Square too (it said \"by hand\")");
  const dlg = fs.readFileSync(path.join(ROOT, 'NewStudentDialog.html'), 'utf8');
  const sibCode = dlg.slice(dlg.indexOf('currentSiblings.forEach(sib =>'), dlg.indexOf('payloads.push(sibData)'));
  assert(/prefilled: data\.prefilled/.test(sibCode), "the dialog's sibling payload carries prefilled");
  // Mona typed by hand for Bart: one "added from" line for the row, phones as the sheets show them
  const bart2 = formFor("123456789");
  bart2.guardians.push({ name: 'Mona Simpson', relationship: 'Grandmother', phone: '7185559999', email: '', language: '' });
  g.saveStudentOverrides([bart2]); t.build();
  g.buildOverrideReport();
  const rep = t.tab('Students Override Report').find(r => r[0] === "123456789" || r[0] === 123456789);
  assert(rep, "Bart is in the report");
  const text = String(rep[2]);
  assert(/Phone: \(718\) 555-9999/.test(text), "Mona's phone formatted (it showed 7185559999): " + text);
  assert.strictEqual((text.match(/added from/gi) || []).length, 1, "the sources once for the row, not under each parent: " + text);
  assert(/These parents were added from: Parent Square, by hand$/.test(text), text);
  console.log("Added parents: siblings keep the source; the report's sources once, phones formatted: ok");
}

// ---------------------------------------------------------------- the typical user's findings (2026-11-23)
{
  const g = loadProject({});
  const G = n => ({ guardianOutput: n, email: '', phones: [], mobile: '' });
  const maybe = (raw, names) => g.parentMaybeSame_(raw, [], '', names.map(G));
  // "Maybe the same person" only with a reason: a shared last name alone is another person (Abe was "Marge, a close spelling")
  assert.strictEqual(maybe('Abe Simpson', ['Homer Simpson', 'Marge Simpson']), null, "Abe Simpson is not maybe Marge");
  assert.strictEqual(maybe('Patty Bouvier', ['Homer Simpson', 'Marge Simpson']), null);
  eq(maybe('Morticia Adams', ['Morticia Addams', 'Gomez Addams']).why, "a close spelling");
  eq(maybe('Marje Simpson', ['Homer Simpson', 'Marge Simpson']).why, "a close spelling");
  eq(maybe('Marge Bouvier', ['Homer Simpson', 'Marge Simpson']).why, "same first name, another last name", "the s4-14 variant still shows");
  assert(/has no last name in your sheets/.test(maybe('Hippolyta Prince', ['Hippolyta']).why), "a one-name guardian has no last name to differ");
  // A guardian saved with one name is the source's parent with that FIRST name (same student), never a last name alone
  const gs = [G('Hippolyta')];
  assert.strictEqual(g.knownGuardianFor_(gs, ['hippolyta', 'prince'], [], '', 'HIPPOLYTA PRINCE'), gs[0], "Hippolyta = HIPPOLYTA PRINCE");
  assert.strictEqual(g.knownGuardianFor_(gs, ['prince', 'hippolyta'], [], '', 'Prince, Hippolyta'), gs[0], "\"Last, First\" too");
  assert.strictEqual(g.knownGuardianFor_([G('Prince')], ['hippolyta', 'prince'], [], '', 'HIPPOLYTA PRINCE'), null, "a last name alone never matches");
  assert.strictEqual(g.knownGuardianFor_(gs, ['hippolyta', 'prince'], [], ''), null, "without the raw name: as before");
  const menus = fs.readFileSync(path.join(ROOT, 'Menus.gs'), 'utf8');
  assert(/tokens, phones, email, rawName\)\) return;/.test(menus) && /tokens, digits, email, x\.rawName\);/.test(menus), "both callers pass the raw name");

  // Rows in Raw Data but no date stamp (pasted by hand, the fake data): loaded, date not known
  const { newTracker } = require('./lib/tracker');
  const t = newTracker();
  t.idmUpload("1 - Student Data (ATS).csv", "2 - Parent Square.csv"); t.g.idm_exit();
  const raw = t.ss.getSheetByName("Raw Data");
  raw.getRange("A1").setValue("ATS - DATE UPDATED: 00/00/00");
  raw.getRange("AG1").setValue("Parent Square - DATE UPDATED: 00/00/00");
  const items = t.g.idm_startItems_(t.ss, null);
  eq(items[0].last, { text: "(date not known)", days: null }, "Student Data: loaded, date not known (it said Never loaded)");
  eq(items[1].last, { text: "(date not known)", days: null });
  assert.strictEqual(items[2].last, null, "NYCSA was never loaded");
  eq(items.map(i => i.picked), [false, false, false, false, false], "nothing ticked: no age to go by (every card was ticked)");
  const ids = fs.readFileSync(path.join(ROOT, 'InputDataSidebar.html'), 'utf8');
  assert(/ats && ats\.required && !ats\.last/.test(ids), "the \"Go on without it?\" question needs no Student Data at all");

  // Wording
  const idm = fs.readFileSync(path.join(ROOT, 'Input Data.gs'), 'utf8');
  assert(/click " \+ found\.step\.short \+ " in the bar at the top/.test(idm) && !/or Skip to the/.test(idm), "a wrong paste says to click its step in the bar (Skip only goes forward)");
  eq(g.idm_summary_({ nycsa: { skipped: true } })[0], "— Student Data (ATS / Insight): not loaded (nothing changed)");
  assert(/skipped \(nothing changed\)/.test(g.idm_summary_({ nycsa: { skipped: true } })[2]), "no \"old data kept\" on a copy that never had any");
  assert(/has a parent who isn\\'t/.test(ids) && /parents who aren\\'t/.test(ids), "the end card's grammar");
  assert(/skipped \(nothing changed\)<\/div>/.test(ids));
  const ys = fs.readFileSync(path.join(ROOT, 'YourListsSidebar.html'), 'utf8');
  assert(/s\.review === 1 \? ' still needs you'/.test(ys), "1 still needs you");
  assert(/purpose === 'contacts'\s*\?\s*' Parents are found through their child/.test(ys), "a parents-only list's advice is about the child on the row");
  assert(!/Delta Report/.test(fs.readFileSync(path.join(ROOT, 'NewStudentDialog.html'), 'utf8')), "Override Report, as the page calls it");
  const mo = fs.readFileSync(path.join(ROOT, 'ManageOverridesDialog.html'), 'utf8');
  assert(/for the students on the card\?/.test(mo) && !/leaves their contact\. If your Student Data/.test(mo), "\"Use it instead\" says what happens for siblings");
  const ps = fs.readFileSync(path.join(ROOT, 'ProgressSidebar.html'), 'utf8');
  assert(/if \(state\.notice\) el\('btnStart'\)\.style\.display = 'none'/.test(ps), "no \"Build again\" when there's nothing to build yet");
  console.log("the typical user's findings (maybe the same, one-name guardians, undated data, wording): ok");
}
