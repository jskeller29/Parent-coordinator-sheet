// The files in test-data/ go through the real code the way the sheet would use them:
// Input Data Mode finds each export's header (and refuses the broken one), and each
// Your Lists file gets the expected mix of matches. Keeps test-data/README.md honest.
const fs = require('fs'), path = require('path'), assert = require('assert');
const { loadProject, ROOT } = require('./lib/load');
const DATA = path.join(ROOT, 'test-data');
const g = loadProject();

function parseCsv(text) {
  const rows = []; let row = [], cell = '', q = false;
  text = text.replace(/^﻿/, '');
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) { if (ch === '"' && text[i + 1] === '"') { cell += '"'; i++; } else if (ch === '"') q = false; else cell += ch; }
    else if (ch === '"') q = true;
    else if (ch === ',') { row.push(cell); cell = ''; }
    else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(cell); rows.push(row); row = []; cell = ''; }
    else cell += ch;
  }
  if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
  return rows;
}
const csv = name => parseCsv(fs.readFileSync(path.join(DATA, name), 'utf8'));
const disp = rows => rows.map(r => r.map(String));

// Input Data Mode: header, right export, cut-off numbers
const exportsExpected = [
  ["1 - Student Data (ATS).csv", "ats", 18, false],
  ["1b - Student Data (ATS) - next month (2 new, 1 left).csv", "ats", 19, false],
  ["1c - Student Data (ATS) - BROKEN numbers cut off.csv", "ats", 18, true],
  ["2 - Parent Square.csv", "ps", 28, false],
  ["3 - NYCSA (Students with Non-NYCSA Account).csv", "nycsa", 25, false]
];
exportsExpected.forEach(([file, key, n, cut]) => {
  const rows = csv(file);
  const found = g.idm_findHeader_(rows);
  assert(found && found.step.key === key, file + ": header / export");
  const mapped = g.idm_mapRows_(found.step, rows, disp(rows), found.row);
  assert.strictEqual(mapped.rows.length, n, file + ": rows");
  assert.strictEqual(g.idm_findCutoffNumbers_(found.step, mapped.rows, mapped.displays).length > 0, cut, file + ": cut-off check");
});
console.log("test data, Input Data Mode exports: ok");

// Your Lists against the Student List the ATS file builds
const ats = csv("1 - Student Data (ATS).csv").slice(1);
const raw = [[], [], [], [], [], []].concat(ats);
const sm = g.buildStudentMap_(raw, new Map());
const idx = g.od_index_([...sm.entries()].map(([id, s]) => [id, s.student, "", s.site || "", s.classCode || ""]), g.buildGuardianObjects_(raw, new Map()));
const listsExpected = [
  ["5a - Your Lists - Bus Routes (student list).csv", "student", 10, 6, 2],
  ["5b - Your Lists - PTA Members (guardian list).csv", "guardian", 11, 1, 1],
  ["5c - Your Lists - Field Trip Forms (first + last name columns).csv", "student", 9, 0, 1],
  ["5d - Your Lists - After School Sign-up (LAST, FIRST).csv", "student", 6, 2, 0]
];
listsExpected.forEach(([file, by, ok, guesses, none]) => {
  const block = csv(file).filter(r => r.some(x => String(x).trim() !== ""));
  const h = g.od_findHeaderRow_(block);
  const headers = block[h], rows = block.slice(h + 1);
  const guess = g.od_guessRoles_(headers, rows, idx);
  assert.strictEqual(guess.by, by, file + ": student / guardian list");
  const res = g.od_matchAll_(g.od_fillSetup_({ by: guess.by, roles: guess.roles, picks: {} }), rows, idx);
  const count = s => res.filter(r => r.status === s).length;
  assert.deepStrictEqual([count("ok"), count("guess"), count("none")], [ok, guesses, none], file + ": matched / best guesses / not found");
});
console.log("test data, Your Lists: ok");
