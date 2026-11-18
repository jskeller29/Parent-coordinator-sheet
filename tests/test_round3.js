// Release check round 2: the owner's notes of 11/17 and the three new ideas (2026-11-18)
//  - build report: Parent Square / NYCSA sign-ups read by header, per student + guardian (2.18)
//  - big cache values split over several keys (the ghost typist on a big sheet; 7.2 speed)
//  - Contact Log: a paste that reaches END keeps the new row's checkbox, and no second blank row (7.2)
//  - Additional Notes as bullets under the date line (idea C)
//  - Input Data Mode: pick what to update on the start page, then only those steps (idea A); the
//    Your Lists tab shows a list added from the panel (2.12 / 2.13)
//  - Parents to add: parents Parent Square / NYCSA have that ATS doesn't, added on top of ATS (idea B)
//  - ⚙️ Settings ticks the checklist however it is opened (2.20); no "Not now" in the welcome (2.5)
const fs = require('fs'), path = require('path'), assert = require('assert');
const { loadProject, ROOT, memStore } = require('./lib/load');
const { makeSpreadsheet, fakeSpreadsheetApp } = require('./lib/fakesheet');
// Values from the project's context have that context's Array / Object: compare them as JSON
const eq = (a, b, m) => assert.deepStrictEqual(JSON.parse(JSON.stringify(a === undefined ? null : a)), JSON.parse(JSON.stringify(b === undefined ? null : b)), m);

// ---------------------------------------------------------------- build report: sign-ups (2.18)
{
  const g = loadProject();
  const H = ["student_id", "Guardian output", "Student", "Class / Site", "relationship", "language", "email", "mobile",
             "Sibling", "Parents Square", "NYCSA", "Last Interaction"];
  const row = (osis, name, student, sib, ps, ny) => [osis, name, student, "", "", "", "", "", sib, ps, ny, ""];
  const before = [H,
    row("123456789", "Marge Simpson", "Bart Simpson", "Lisa Simpson", "No", "No (NG) B1"),
    row("567891234", "Marge Simpson", "Lisa Simpson", "Bart Simpson", "Yes", "No (NG) L1"), // same parent, other child
    row("111222333", "Morticia Addams", "Wednesday Addams", "", "", "No (OG) W1"),
    row("444555666", "Ned Flanders", "Rod Flanders", "", "Yes", "Yes/Unknown")];
  const after = [H,
    row("123456789", "Marge Simpson", "Bart Simpson", "Lisa Simpson", "Yes", "Yes/Unknown"),   // both signed up
    row("567891234", "Marge Simpson", "Lisa Simpson", "Bart Simpson", "Yes", "No (NG) L1"),   // nothing new
    row("111222333", "Morticia Addams", "Wednesday Addams", "", "Yes", "No (OG) W1"),        // joined Parent Square
    row("444555666", "Ned Flanders", "Rod Flanders", "", "Yes", "Yes/Unknown"),
    row("777888999", "New Parent", "New Kid", "", "Yes", "Yes/Unknown")];                    // a new student: not a sign-up
  let r = g.signupReport_(before, after);
  eq(r.nycsa, ["Marge Simpson (Student: Bart Simpson)"], 'NYCSA: read from the NYCSA column, per child');
  eq(r.ps, ["Marge Simpson (Student: Bart Simpson)", "Morticia Addams (Student: Wednesday Addams)"],
    'Parent Square: No -> Yes, and not in Parent Square -> registered once Parent Square was loaded');
  r = g.signupReport_(after, after);
  eq([r.nycsa, r.ps], [[], []], 'the next build with nothing new: 0 and 0 (it said "2 new NYCSA" again)');
  // The first Parent Square load: nobody had a status, so nobody is a "sign-up"
  const noPs = before.map((x, i) => i ? x.slice(0, 9).concat(["", x[10], ""]) : x);
  eq(g.signupReport_(noPs, after).ps, [], 'first Parent Square load: no sign-ups');
  // The old layout without the Sibling column (positions differ): still read by header
  const old = before.map(x => x.slice(0, 8).concat(x.slice(9)));
  eq(g.signupReport_(old, after).nycsa, ["Marge Simpson (Student: Bart Simpson)"], 'older Parents List layout');
  const build = fs.readFileSync(path.join(ROOT, 'Build.gs'), 'utf8');
  assert(!/oldPdData\[i\]\[8\]|parentsDivided\[i\]\[8\]/.test(build), 'no fixed column numbers left in the report');
  console.log("build report sign-ups: ok");
}

// ---------------------------------------------------------------- big cache values (7.2 speed)
{
  const g = loadProject();
  const cache = memStore();
  cache.putAll = (vals) => Object.keys(vals).forEach(k => cache.put(k, vals[k]));
  const piece = g.BIG_CACHE_PIECE_;
  assert(g.bigCachePut_(cache, "K", "small", 60), 'small value stored');
  assert.strictEqual(cache.get("K"), "small", 'small value stored as before (no marker)');
  assert.strictEqual(g.bigCacheGet_(cache.getAll(g.bigCacheKeys_("K")), "K"), "small");
  const big = "x".repeat(piece * 2) + "end";
  assert(g.bigCachePut_(cache, "K", big, 60), 'a big value is split');
  assert(cache.get("K_1") && cache.get("K_2"), 'over three keys');
  assert.strictEqual(g.bigCacheGet_(cache.getAll(g.bigCacheKeys_("K")), "K"), big, 'read back whole');
  cache.remove("K_1");
  assert.strictEqual(g.bigCacheGet_(cache.getAll(g.bigCacheKeys_("K")), "K"), null, 'a lost piece = not cached');
  g.bigCacheRemove_(cache, "K");
  assert.strictEqual(cache.get("K"), null);
  assert.strictEqual(g.bigCachePut_(cache, "K", "y".repeat(piece * g.BIG_CACHE_PIECES_ + 1), 60), false, 'too big even in pieces');
  // The typist's data: kept even when it is bigger than one key (it was read from the tabs on every edit)
  const ss = makeSpreadsheet(["Directory", "Backend"]);
  const g2 = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) } });
  const dc = g2.CacheService.getDocumentCache();
  dc.putAll = (vals) => Object.keys(vals).forEach(k => dc.put(k, vals[k]));
  const master = Array.from({ length: 3000 }, (_, i) => [String(100000000 + i), "Student Number " + i, "Guardian Name " + i + " and Other " + i]);
  g2.studentLookupRows_ = () => master;
  let reads = 0;
  const realGet = g2.studentLookupRows_;
  g2.studentLookupRows_ = function () { reads++; return realGet.apply(this, arguments); };
  g2.ghostTypistData_(ss);
  assert(JSON.stringify(master).length > 95000, 'the test data is bigger than one cache key');
  const pre = g2.contactLogPrefetch_({ getSheetId: () => 1, getLastRow: () => 0, getRange: () => ({ getDisplayValues: () => [[""]], getDisplayValue: () => "" }) }, 2);
  assert(pre.ghostRaw && JSON.parse(pre.ghostRaw).master.length === 3000, 'the next edit gets it from the cache');
  g2.ghostTypistData_(ss, pre.ghostRaw);
  assert.strictEqual(reads, 1, 'the tabs are read once, not on every edit');
  g2.clearGhostTypistCache_();
  assert.strictEqual(dc.get("GHOST_TYPIST_DATA_1"), null, 'a build clears every piece');
  console.log("big cache values: ok");
}

// ---------------------------------------------------------------- Contact Log: the new row's checkbox (7.2)
{
  const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(makeSpreadsheet(["Contact Log"])) } });
  let inserted = 0, cleared = [];
  const fakeLog = {
    insertRowAfter: () => { inserted++; },
    getRange: (r, c, nr, nc) => {
      const rg = new Proxy({}, { get: (t, k) => (k === 'clearDataValidations' ? () => { cleared.push(r); return rg; } : () => rg) });
      return rg;
    }
  };
  let insertedRows = 0;
  fakeLog.insertRowsAfter = (r, n) => { inserted++; insertedRows += n; };
  assert.strictEqual(g.autoAddBlankRow_(fakeLog, 9, 10), 3, 'the last row before END: 3 empty rows are added');
  assert.strictEqual(g.autoAddBlankRow_(fakeLog, 5, 10), 0, 'not the last row: none');
  assert.strictEqual(g.autoAddBlankRow_(fakeLog, 9, null), 0, 'END unknown: none');
  assert.strictEqual(inserted, 1, 'one insert call'); assert.strictEqual(insertedRows, 3);
  const edits = fs.readFileSync(path.join(ROOT, 'Edits.gs'), 'utf8');
  const handler = edits.slice(edits.indexOf('function handleEdit_'), edits.indexOf('// CONTACT LOG SEARCH BAR'));
  assert(/pasteAddedRows \|\| autoAddBlankRow_\(sheet, row, endRow\)/.test(handler) && /endRow \+ addedRows/.test(handler), 'a paste that added the row: END known to have moved, no second row');
  assert(/rowLock\.tryLock\(10000\)/.test(handler) && handler.indexOf('releaseRowLock(); // the new row') > handler.indexOf('autoAddBlankRow_(sheet, row, endRow)'),
    'one Contact Log edit at a time until the new row is on the sheet');
  assert(handler.indexOf('releaseRowLock(); // the new row') < handler.indexOf('syncNotesForRowOsis_(mainOsis'), 'let go before the notes (they lock too)');
  assert(/if \(numRows > 1 \|\| range\.getNumColumns\(\) > 1\) \{\s*let maxRows/.test(handler), 'a one-cell edit never clears checkboxes under END');
  console.log("Contact Log new row: ok");
}

// ---------------------------------------------------------------- Additional Notes as bullets (idea C)
{
  const g = loadProject();
  const row = (desc, extra) => ["", "11/17/2026", "Marge Simpson", "Marge", "123456789", "Bart Simpson", 1, "", "", "Phone", "", "", desc, "", extra, false];
  let e = g.contactLogNoteLine_(row("Called about the trip", "- Needs the form by Friday\n- Will bring the money"), false);
  const lines = e.text.split("\n");
  assert.strictEqual(lines.length, 3, 'date line + 2 bullets: ' + JSON.stringify(e.text));
  assert(/Called about the trip$/.test(lines[0]) && lines[0].indexOf("[") === -1, 'no [ ] any more');
  assert.strictEqual(lines[1], "   • Needs the form by Friday", 'the "-" is left out');
  assert.strictEqual(lines[2], "   • Will bring the money", 'a second "-" line = a second bullet');
  e = g.contactLogNoteLine_(row("Called", "Just one note"), false);
  assert(/\n   • Just one note$/.test(e.text), 'no dash: still one bullet under the date');
  e = g.contactLogNoteLine_(row("", "- only the note"), false);
  assert(/\} *\n   • only the note$/.test(e.text) || /-\n   • only the note$/.test(e.text) || /\}\n   • only the note$/.test(e.text), 'no description: the bullet under the date line: ' + JSON.stringify(e.text));
  assert.strictEqual(g.contactLogNoteLine_(row("", "-"), false), null, 'a lone "-" is nothing');
  assert.strictEqual(g.contactLogNoteLine_(row("", "  \n "), false), null, 'blank lines are nothing');
  eq(g.additionalNoteBullets_("- - a\n\n• b\n* c\n–d"), ["   • a", "   • b", "   • c", "   • d"], 'any bullet mark typed is left out');
  // Combined Contact Tracking's divider lines still wrap each whole entry; Phone Contacts strips only them
  const joined = g.joinNoteLines_("", [{ dateObj: new Date(2026, 10, 17), text: "A\n   • b" }, { dateObj: new Date(2026, 10, 18), text: "C" }]);
  eq(g.stripNoteDividers_(joined).split("\n"), ["C", "A", "   • b"], 'newest first, bullets kept with their entry');
  console.log("Additional Notes as bullets: ok");
}

// ---------------------------------------------------------------- Input Data Mode: pick what to update (idea A)
{
  const TABS = ["Contact Log", "Combined Contact Tracking", "PCAR", "Directory", "Raw Data", "Student List", "Notes", "Backend"];
  const ss = makeSpreadsheet(TABS);
  ss.getSheetByName("Backend").hideSheet();
  const raw = ss.getSheetByName("Raw Data");
  raw._put([["ATS - DATE UPDATED: 10/01/26"]], 1, 1);           // loaded before, long ago: due
  raw._put([["Parent Square - DATE UPDATED: 00/00/00"]], 1, 33);  // never loaded
  const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) } });
  g.htmlFromFile_ = () => ({ setTitle() { return this; }, setWidth() { return this; } });
  let v = g.idm_open();
  assert(v.overview, 'starts on the start page');
  const picked = v.start.filter(it => it.picked).map(it => it.key);
  eq(picked, ["ats"], 'suggested: the export that is due (Student Data), not the never-loaded optional ones');
  assert(v.start[0].due && !v.start[1].due, 'due flag');
  assert.strictEqual(g.idm_startPicked([]).notice.code, 'nothingPicked', 'nothing ticked: a message, no step');
  v = g.idm_startPicked([2, 0]); // NYCSA + Student Data
  assert(!v.overview && v.stepIndex === 0, 'Start: the first ticked step');
  eq(v.next, { index: 2, short: "NYCSA" }, 'next = the next ticked step');
  eq(v.progress.map(p => p.picked), [true, false, true, false, false], 'the bar knows what is ticked');
  v = g.idm_skip();
  assert.strictEqual(v.stepIndex, 2, 'Skip goes to the next ticked step (Parent Square left out)');
  assert.strictEqual(v.next, null, 'NYCSA is the last ticked one');
  v = g.idm_skip();
  assert(v.finished && !v.active, 'after the last ticked step the mode ends (Build / Done)');
  // The "completed!" banner names the next TICKED step
  const st = { step: 0, picked: [0, 3], results: {} };
  const n = g.idm_doneNotice_(ss, st, g.IDM_STEPS[0], 5, "row", "rows", "moved", "x");
  assert(/Next: Site \/ Class\./.test(n.done.sub), 'banner: Next: the next ticked step: ' + n.done.sub);
  // An older run (no picks): one step after the other, as before
  assert.strictEqual(g.idm_nextStep_({ step: 1, results: {} }, 1), 2);
  assert.strictEqual(g.idm_nextStep_({ step: 4, results: {} }, 4), null);
  // A ticked step passed over by a click in the bar comes back after the later ones
  assert.strictEqual(g.idm_nextStep_({ picked: [0, 1, 3], results: { ats: { rows: 1 } } }, 3), 1);
  // A step clicked in the bar from the start page: ticked + that one, and it goes there first
  const ss2 = makeSpreadsheet(TABS);
  const g2 = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss2) } });
  g2.idm_open();
  v = g2.idm_startPicked([0], 3);
  assert(v.stepIndex === 3 && v.next && v.next.index === 0, 'bar click: that step first, then the ticked one');
  const sidebar = fs.readFileSync(path.join(ROOT, 'InputDataSidebar.html'), 'utf8');
  assert(/call\('idm_startPicked', picks/.test(sidebar) && /togglePick\(/.test(sidebar), 'sidebar: tick cards, Start sends the ticks');
  assert(/<button class="primary" id="buildBtn"[^>]*>Build my sheets<\/button>\s*<button id="doneBtn"[^>]*>Done<\/button>/.test(sidebar), 'end: Build my sheets + Done');
  console.log("Input Data Mode: pick what to update: ok");
}

// ---------------------------------------------------------------- Input Data Mode panel layout (2.9 / 2.11 / 2.15)
{
  const sidebar = fs.readFileSync(path.join(ROOT, 'InputDataSidebar.html'), 'utf8');
  const body = sidebar.slice(sidebar.indexOf('<body>'));
  assert(body.indexOf('id="errorBox"') < body.indexOf('id="startView"') && body.indexOf('id="errorBox"') > body.indexOf('id="progress"'),
    '2.9: the error box sits right under the title and the bar of steps, outside the step');
  const foot = body.slice(body.indexOf('<div class="footer" id="footer">'));
  assert(foot.indexOf('id="siteNextBtn"') !== -1 && foot.indexOf('id="siteNextBtn"') < foot.indexOf('id="skipBtn"'), '2.11: Save above Skip, in the footer');
  assert(!/id="saveSiteBtn"/.test(sidebar), '2.11: one Save (the card\'s is gone)');
  assert(/show\('footTop', mode === 'site'\)/.test(sidebar), '2.11: shown on the Site / Class step');
  assert(/button\.ico::after \{ content: "Edit"/.test(sidebar), '2.15: "Edit" on the icon when the mouse is on it');
  console.log("Input Data Mode panel layout: ok");
}

// ---------------------------------------------------------------- the Your Lists tab follows the lists (2.12 / 2.13)
{
  const ss = makeSpreadsheet(["Contact Log", "Raw Data", "Student List", "Backend"]);
  const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) } });
  let lists = [];
  g.od_loadLists_ = () => lists;
  g.od_saveLists_ = (s, ls) => { lists = ls; };
  const state = { step: 4, tempId: null, hiddenIds: [], prevActiveId: null, results: {} };
  g.idm_ensureTempSheet_(ss, state);
  g.idm_saveState_(state);
  const tab = g.idm_sheetById_(ss, state.tempId);
  assert.strictEqual(tab.getName(), "Your Lists");
  assert.strictEqual(tab.getLastRow(), 4, 'no lists yet: title, help, hint, header');
  const listTab = ss.insertSheet("🚌 Bus Routes");
  lists.push({ tabId: listTab.getSheetId(), name: "Bus Routes", setup: g.od_emptySetup_(), updated: "11/17/26", sheet: listTab });
  g.__props[g.IDM_BACK_FROM_LIST_KEY_] = "1"; // back from the list's own panel
  g.idm_open();
  assert.strictEqual(tab.getRange(5, 1).getValue(), "Bus Routes", 'the list added in its own panel shows on the tab');
  const drawn = tab._log ? tab._log.length : 0;
  g.idm_getState();
  assert.strictEqual(tab.getRange(5, 1).getValue(), "Bus Routes", 'still there');
  lists = [];
  g.idm_getState(); // after a delete
  assert.strictEqual(tab.getLastRow(), 4, 'a deleted list leaves the tab');
  void drawn;
  console.log("Your Lists tab follows the lists: ok");
}

// ---------------------------------------------------------------- Parents to add (idea B)
function rawRow(osis, last, first, cls, g1, g2) {
  const r = new Array(31).fill("");
  r[0] = "2025-2026"; r[2] = osis; r[3] = last; r[4] = first; r[16] = "3"; r[17] = cls; r[18] = "English";
  const put = (base, g) => { if (!g) return; r[base] = g[0]; r[base + 1] = g[1]; r[base + 2] = g[2]; r[base + 3] = g[3]; r[base + 4] = "English"; };
  put(19, g1); put(25, g2);
  return r;
}
function parentsSheet() {
  const ss = makeSpreadsheet(["Contact Log", "Raw Data", "Student List", "Backend"]);
  const raw = ss.getSheetByName("Raw Data");
  raw._put([["ATS - DATE UPDATED: 11/01/26"]], 1, 1);
  raw._put([
    rawRow("123456789", "SIMPSON", "BART", "Y41", ["HOMER SIMPSON", "FATHER", "1111111111", "homer@x.org"], ["MARGE SIMPSON", "MOTHER", "2222222222", "marge@x.org"]),
    rawRow("567891234", "SIMPSON", "LISA", "Y25", ["HOMER SIMPSON", "FATHER", "1111111111", "homer@x.org"], ["MARGE SIMPSON", "MOTHER", "2222222222", "marge@x.org"]),
    rawRow("111222333", "ADDAMS", "WEDNESDAY", "Y12", ["MORTICIA ADDAMS", "MOTHER", "5555555555", ""], ["GOMEZ ADDAMS", "FATHER", "6666666666", ""])
  ], 7, 1);
  // Parent Square (AG:AN): Name, Students, Email, Phone, Secondary Phone, Record Created, Registered?, SIS Synced
  raw._put([
    ["Simpson, Homer", "Bart Simpson", "homer@x.org", "(111) 111-1111", "", "", "Yes", "Yes"],      // known
    ["Simpson, Abe", "Bart Simpson, Lisa Simpson", "abe@x.org", "(333) 333-3333", "", "", "Yes", "No"], // new, two children
    ["Simpson, Marjorie", "Lisa Simpson", "marge@x.org", "", "", "", "Yes", "Yes"],                   // same email + last name: known
    ["Flanders, Ned", "Rod Flanders", "ned@x.org", "", "", "", "Yes", "Yes"]                          // no such student
  ], 7, 33);
  // NYCSA (AP:BF)
  const ny = (osis, sf, sl, gf, gl, cell, rel) => ["M094", "75", "PS", "Y41", "3", osis, sf, sl, gf, gl, "No", "", "", cell, rel, "Custodial User", "C1"];
  raw._put([
    ny("123456789", "BART", "SIMPSON", "HOMER", "SIMPSON", "1111111111", "FATHER"),   // known
    ny("123456789", "BART", "SIMPSON", "PATTY", "BOUVIER", "4444444444", "AUNT"),     // new
    ny("111222333", "WEDNESDAY", "ADDAMS", "GRANDMAMA", "ADDAMS", "7777777777", "GRANDMOTHER"), // new
    ny("111222333", "WEDNESDAY", "ADDAMS", "MORTICIA", "ADDAMS", "5555555555", "MOTHER"),       // known
    ny("111222333", "WEDNESDAY", "ADDAMS", "", "ADDAMS", "", "")                                  // a last name alone: never offered
  ], 7, 42);
  const app = fakeSpreadsheetApp(ss);
  const g = loadProject({ extra: { SpreadsheetApp: app } });
  ss.getSheetByName("Backend")._put([g.__get("OVERRIDES_HEADERS_")], 2, 7);
  let builds = 0;
  g.runBuildOnly = () => { builds++; };
  g.od_loadLists_ = () => [];
  return { ss, g, raw, builds: () => builds };
}
{
  const { ss, g, raw, builds } = parentsSheet();
  let found = g.findParentsToAdd_(ss);
  const show = f => f.items.map(it => it.student + ": " + it.name + " [" + it.sources.join(", ") + "]");
  eq(show(found), [
    "Wednesday Addams: Grandmama Addams [NYCSA]",
    "Bart Simpson: Abe Simpson [Parent Square]",
    "Bart Simpson: Patty Bouvier [NYCSA]",
    "Lisa Simpson: Abe Simpson [Parent Square]"
  ], 'the parents ATS doesn\'t have, nobody it has (name order, case, same email + last name)');
  const patty = found.items.find(it => it.name === "Patty Bouvier");
  eq([patty.relationship, patty.phone, patty.cls, patty.studentFirst, patty.studentLast],
    ["Aunt", "(444) 444-4444", "Y41", "Bart", "Simpson"], 'details for the card and the override row');

  // Add two for Bart: ONE "Added Guardian" row, on top of ATS (Homer and Marge stay live)
  const abeBart = found.items.find(it => it.name === "Abe Simpson" && it.student === "Bart Simpson");
  const r = g.addParentsAsGuardians([abeBart.key, patty.key]);
  assert.strictEqual(r.added, 2);
  assert.strictEqual(builds(), 1, 'one build after adding, like every override change');
  const rows = g.backendReadBlock_(ss, "OVERRIDES").slice(1).filter(x => String(x[2]).trim());
  assert.strictEqual(rows.length, 1, 'one row for the student');
  assert.strictEqual(rows[0][1], "Added Guardian (Parent Square, NYCSA)", 'the row says where they came from');
  eq([rows[0][3], rows[0][4], rows[0][10], rows[0][16]], ["Simpson", "Bart", "Abe Simpson", "Patty Bouvier"]);
  let ov = g.getOverrideData_(ss);
  assert(!ov.students.has("123456789"), 'the student itself still comes from ATS');
  const names = id => g.buildGuardianObjects_(raw.getDataRange().getValues(), ov.guardians)
    .filter(x => x.student_id === id).map(x => x.guardianOutput.replace(/ \(.*\)$/, ""));
  eq(names("123456789"), ["Homer Simpson", "Marge Simpson", "Abe Simpson", "Patty Bouvier"], 'ATS guardians + the added ones');
  eq(names("567891234"), ["Homer Simpson", "Marge Simpson"], 'the sibling is not changed');
  found = g.findParentsToAdd_(ss);
  eq(show(found), ["Wednesday Addams: Grandmama Addams [NYCSA]", "Lisa Simpson: Abe Simpson [Parent Square]"], 'added ones are not offered again');
  // In New / Edit Student the student shows with all four; the Override Report-style list knows the kind
  const bart = g.getAllStudentsForEdit().find(s => s.osis === "123456789");
  eq(bart.guardians.map(x => x.name), ["HOMER SIMPSON", "MARGE SIMPSON", "Abe Simpson", "Patty Bouvier"]);
  assert.strictEqual(bart.firstName, "BART", 'the student is the ATS one');
  const list = g.getOverridesList();
  eq([list[0].kind, list[0].source, list[0].guardians], ["added", "Parent Square, NYCSA", ["Abe Simpson", "Patty Bouvier"]]);

  // "Not new": not offered again; "Show them again" brings it back
  const gm = found.items.find(it => it.name === "Grandmama Addams");
  let home = g.skipParentsToAdd([gm.key]);
  eq([home.toAdd.length, home.skipped], [1, 1], 'skipped one is hidden and counted');
  home = g.unskipParentsToAdd();
  eq([home.toAdd.length, home.skipped], [2, 0], 'back again');

  // A student with a New / Edit row: the parent goes into that row (it replaces the ATS guardians)
  g.saveStudentOverrides([{ entryType: "Edit Existing Student", osis: "567891234", classCode: "Y25", firstName: "Lisa", lastName: "Simpson",
    label: "", hideFromReport: false, grade: "", age: "", gender: "", homeLanguage: "",
    guardians: [{ name: "Marge Simpson", relationship: "Mother", phone: "2222222222", email: "marge@x.org", language: "English" }] }]);
  const abeLisa = g.findParentsToAdd_(ss).items.find(it => it.student === "Lisa Simpson");
  g.addParentsAsGuardians([abeLisa.key]);
  const lisaRow = g.backendReadBlock_(ss, "OVERRIDES").slice(1).find(x => String(x[2]) === "567891234");
  eq([lisaRow[1], lisaRow[10], lisaRow[16]], ["Edit Existing Student", "Marge Simpson", "Abe Simpson"], 'added into the Edit row');

  // ATS catches up: the build stops adding the one ATS has; once ATS has them all, the row goes
  const sheetRow = 7; // Bart
  raw._put([["PATTY BOUVIER", "AUNT", "4444444444", "", "English", "English"]], sheetRow, 26);
  ov = g.getOverrideData_(ss);
  eq(names("123456789"), ["Homer Simpson", "Patty Bouvier", "Abe Simpson"], 'no double Patty once ATS has her');
  g.cleanRedundantOverrides();
  assert(g.backendReadBlock_(ss, "OVERRIDES").slice(1).some(x => String(x[2]) === "123456789"), 'Abe still only added: the row stays');
  raw._put([["ABE SIMPSON", "GRANDFATHER", "3333333333", "abe@x.org", "English", "English"]], sheetRow, 20);
  g.cleanRedundantOverrides();
  assert(!g.backendReadBlock_(ss, "OVERRIDES").slice(1).some(x => String(x[2]) === "123456789"), 'ATS has both now: the row goes');
  console.log("Parents to add: ok");
}
{
  // Your Lists: a guardian list's row whose guardian isn't known but whose child is
  const { ss, g } = parentsSheet();
  const listTab = ss.insertSheet("👪 PTA");
  listTab._put([["Parent", "Child", "Role"], ["Selma Bouvier", "Wednesday Addams", "Volunteer"], ["Marge Simpson", "Bart Simpson", "Chair"]], 4, 1);
  g.od_loadLists_ = () => [{ tabId: listTab.getSheetId(), name: "PTA", sheet: listTab,
    setup: g.od_fillSetup_({ by: "guardian", ready: true, icon: "👪", headers: ["Parent", "Child", "Role"],
      roles: { guardian: { col: 1 }, student: { col: 2 } }, vars: [], picks: {} }) }];
  const found = g.findParentsToAdd_(ss);
  const selma = found.items.find(it => it.name === "Selma Bouvier");
  assert(selma && selma.student === "Wednesday Addams" && selma.sources[0] === "👪 PTA", 'from a list, with its icon and name: ' + JSON.stringify(found.items.map(i => i.name)));
  assert(!found.items.some(it => it.name === "Marge Simpson"), 'a matched guardian is not offered');
  console.log("Parents to add from your lists: ok");
}
{
  // The test files 2b / 3b (test-data/README.md): exactly Abe (Bart, Lisa) and Patty (Bart); the standard files: nobody
  const parseCsv = text => {
    const rows = []; let row = [], cell = '', q = false;
    text = text.replace(/^\uFEFF/, '');
    for (let i = 0; i < text.length; i++) { const ch = text[i];
      if (q) { if (ch === '"' && text[i + 1] === '"') { cell += '"'; i++; } else if (ch === '"') q = false; else cell += ch; }
      else if (ch === '"') q = true; else if (ch === ',') { row.push(cell); cell = ''; }
      else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(cell); rows.push(row); row = []; cell = ''; }
      else cell += ch; }
    if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
    return rows;
  };
  const csv = n => parseCsv(fs.readFileSync(path.join(ROOT, 'test-data', n), 'utf8')).filter(r => r.some(c => String(c).trim())).slice(1);
  const withFiles = (psFile, nyFile) => {
    const ss = makeSpreadsheet(["Contact Log", "Raw Data", "Student List", "Backend"]);
    const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) } });
    g.od_loadLists_ = () => [];
    const raw = ss.getSheetByName("Raw Data");
    raw._put(csv("1 - Student Data (ATS).csv"), 7, 1);
    raw._put(csv(psFile), 7, 33);
    raw._put(csv(nyFile), 7, 42);
    return g.findParentsToAdd_(ss).items.map(it => it.student + ": " + it.name + " [" + it.sources.join(", ") + "]");
  };
  eq(withFiles("2 - Parent Square.csv", "3 - NYCSA (Students with Non-NYCSA Account).csv"), [], 'the standard test files: nobody to add');
  eq(withFiles("2b - Parent Square - a parent not in ATS.csv", "3b - NYCSA - a third guardian.csv"),
    ["Bart Simpson: Abe Simpson [Parent Square]", "Bart Simpson: Patty Bouvier [NYCSA]", "Lisa Simpson: Abe Simpson [Parent Square]"],
    'test files 2b + 3b: the two parents (README / release check)');
  console.log("Parents to add on the test files: ok");
}
{
  // The main page, New / Edit Student and the end of Input Data Mode are wired together
  const home = fs.readFileSync(path.join(ROOT, 'ManageOverridesDialog.html'), 'utf8');
  ['getOverridesHome', 'addParentsAsGuardians', 'skipParentsToAdd', 'unskipParentsToAdd', 'openNewStudentFromHome', 'buildOverrideReport', 'deleteOverride']
    .forEach(fn => assert(home.indexOf(fn + '(') !== -1, 'main page calls ' + fn));
  const ns = fs.readFileSync(path.join(ROOT, 'NewStudentDialog.html'), 'utf8');
  assert(/getNewStudentPrefill\(\)/.test(ns) && /openManageOverridesDialog\(\)/.test(ns) && /fromHome \? backHome\(\)/.test(ns), 'New / Edit Student: back and forth');
  const idm = fs.readFileSync(path.join(ROOT, 'InputDataSidebar.html'), 'utf8');
  assert(/idm_parentsToAddSummary\(\)/.test(idm) && /openManageOverridesDialog\(\)/.test(idm), 'end of Input Data Mode asks');
  const menus = fs.readFileSync(path.join(ROOT, 'Menus.gs'), 'utf8');
  assert(/addItem\('🗂️ Student Overrides \(parents to add, changes, report\)', 'openManageOverridesDialog'\)/.test(menus), 'menu: the main page first');
  // prefill: read once
  const g = loadProject();
  const up = { m: {}, getProperty(k) { return k in this.m ? this.m[k] : null; }, setProperty(k, v) { this.m[k] = String(v); }, deleteProperty(k) { delete this.m[k]; } };
  g.PropertiesService.getUserProperties = () => up;
  g.openNewStudentDialog = () => {};
  g.openNewStudentFromHome("123456789", { name: "Abe Simpson", relationship: "", phone: "(333) 333-3333", email: "" });
  const p = g.getNewStudentPrefill();
  eq([p.osis, p.guardian.name, p.fromHome], ["123456789", "Abe Simpson", true]);
  eq(g.getNewStudentPrefill(), { osis: "", guardian: null, fromHome: false }, 'read once');
  console.log("Student Overrides main page wiring: ok");
}

// ---------------------------------------------------------------- Settings tick + no "Not now" (2.20 / 2.5)
{
  const ss = makeSpreadsheet(["Contact Log"]);
  const app = fakeSpreadsheetApp(ss);
  const g = loadProject({ extra: { SpreadsheetApp: app } });
  g.htmlFromFile_ = () => ({ setTitle() { return this; }, setWidth() { return this; }, setHeight() { return this; } });
  g.openSettingsDialog(); // from 🚀 App Menu → ⚙️ Settings
  assert.strictEqual(g.__props.ZG_SETTINGS_SEEN, 'true', 'Settings opened from the menu ticks "Check your settings"');
  const menus = fs.readFileSync(path.join(ROOT, 'Menus.gs'), 'utf8');
  assert(!/>Not now</.test(menus) && !/zg_notNow/.test(menus), 'no "Not now" in the welcome window');
  console.log("Settings tick + no Not now: ok");
}

// ---------------------------------------------------------------- 2026-11-18 later: Parents List, Select all, Back to the start
{
  const ss = makeSpreadsheet(["Contact Log", "Parents Divided", "Backend"]);
  const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) } });
  const old = ss.getSheetByName("Parents Divided");
  assert.strictEqual(g.getParentsListSheet_(ss), old, 'the old tab is the one used');
  assert.strictEqual(old.getName(), "Parents List", 'renamed on first use (formulas pointing at it follow in Sheets)');
  assert.strictEqual(ss.getSheetByName("Parents Divided"), null, 'no second tab');
  assert(g.KEEP_TABS_.indexOf("Parents List") !== -1 && g.KEEP_TABS_.indexOf("Parents Divided") !== -1, 'a reset keeps either name');
  const tv = g.__get("TAB_VISIBILITY").find(t => t.key === 'parentsDivided');
  eq(tv.sheets, ["Parents List", "Parents Divided"], 'Settings: same key, both names');
  const code = fs.readdirSync(ROOT).filter(f => /\.(gs|html)$/.test(f)).map(f => fs.readFileSync(path.join(ROOT, f), 'utf8')).join("\n");
  const left = (code.match(/getSheetByName\("Parents Divided"\)/g) || []).length;
  assert.strictEqual(left, 2, 'the old name is looked up only as a fallback (an older tracker in Migration, Z#ContactLog)');
  assert(!/writeSheet_\(ss, "Parents Divided"/.test(code), 'the build writes "Parents List"');
  const sb = fs.readFileSync(path.join(ROOT, 'InputDataSidebar.html'), 'utf8');
  assert(/id="pickAllBtn" onclick="pickAll\(\)">Select all</.test(sb) && /show\('footPick', mode === 'start'\)/.test(sb), 'start page: Select all at the bottom');
  assert(/id="restartBtn" onclick="backToStart\(\)"/.test(sb) && /call\('idm_open'\); \/\/ starts a new run/.test(sb), 'All done: Back to the start');
  console.log("Parents List rename, Select all, Back to the start: ok");
}

// ---------------------------------------------------------------- 2026-11-18 later: maybe the same parent, tabs by hand
{
  const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(makeSpreadsheet(["Contact Log"])) } });
  const gs = [g.buildGuardianObject_("1", "MARGE SIMPSON", "MOTHER", "English", "", "marge@x.org", "2222222222"),
              g.buildGuardianObject_("1", "HOMER SIMPSON", "FATHER", "English", "", "", "1111111111")];
  let m = g.parentMaybeSame_("Bouvier, Marge", [], "", gs);
  assert(m && m.name === "Marge Simpson" && m.pct >= 60 && /same first name/.test(m.why), 'a maiden name: ' + JSON.stringify(m));
  m = g.parentMaybeSame_("Marge Bouvier", [], "marge@x.org", gs);
  assert(m && m.pct >= 90 && /email/.test(m.why), 'same email too: ' + JSON.stringify(m));
  assert.strictEqual(g.parentMaybeSame_("Patty Bouvier", ["4444444444"], "", gs), null, 'someone else: nothing');
  const { ss, g: g2 } = parentsSheet();
  const patty = g2.findParentsToAdd_(ss).items.find(it => it.name === "Patty Bouvier");
  eq(patty.current, ["Homer Simpson", "Marge Simpson"], 'the card lists the student\'s parents now');
  const home = fs.readFileSync(path.join(ROOT, 'ManageOverridesDialog.html'), 'utf8');
  assert(/Parents now: /.test(home) && /Maybe the same person as/.test(home) && /fold\('addSec'\)/.test(home) && /fold\('chgSec'\)/.test(home),
    'page: parents now, maybe the same, both sections fold');
  console.log("Parents to add: maybe the same person, parents now, folding: ok");
}
{
  const ss = makeSpreadsheet(["Contact Log", "Student List", "Phone Contacts", "Directory", "Backend"]);
  ss.getSheetByName("Student List")._put([["OSIS", "Student"], ["123456789", "Bart Simpson"]], 1, 1);
  const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) }, props: { SETUP_COMPLETE: 'true' } });
  ss.getSheetByName("Directory").hideSheet(); // hidden by hand (default: shown)
  eq(g.rememberTabChoices_(ss).sort(), ["phoneContacts", "showDirectory", "showMasterTable"], 'shown / hidden by hand -> the settings');
  assert.strictEqual(g.__props.phoneContacts, 'true'); assert.strictEqual(g.__props.showDirectory, 'false');
  eq(g.rememberTabChoices_(ss), [], 'nothing new: nothing written');
  ss.getSheetByName("Directory").showSheet();
  g.__props.INPUT_DATA_MODE_STATE = '{"step":0}';
  eq(g.rememberTabChoices_(ss), [], 'Input Data Mode on: tabs hidden on purpose, not choices');
  delete g.__props.INPUT_DATA_MODE_STATE; delete g.__props.SETUP_COMPLETE;
  eq(g.rememberTabChoices_(ss), [], 'before setup: nothing');
  const menus = fs.readFileSync(path.join(ROOT, 'Menus.gs'), 'utf8'), idm = fs.readFileSync(path.join(ROOT, 'Input Data.gs'), 'utf8');
  assert(/function openSettingsDialog\(\) \{[\s\S]{0,200}rememberTabChoices_\(\)/.test(menus), 'Settings opens with the tabs as they are');
  assert(/rememberTabChoices_\(ss\)/.test(idm), 'Input Data Mode reads them before hiding the tabs');
  console.log("tabs shown / hidden by hand become the settings: ok");
}
