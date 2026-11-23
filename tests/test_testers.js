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
console.log("tester findings: ok");
