// The sheet look (owner 2026-11-12) on the fake spreadsheet (lib/fakesheet.js):
//  - SHEET_LOOK_ (Core.gs): the Type colors read on white and cream, enough site colors
//  - Contact Log Type colors: the same after every reset / night (Type List order), nothing
//    written when already right, the other rules kept, a type taken off the list loses its rule
//  - Events: a new event and Migration's restyle get the brown title row, the tan header row,
//    Montserrat and none of the old pastels; the look-only restyle leaves the dropdowns alone
//  - Override Report: navy header + Montserrat. Raw Data: no Roboto, the date titles on the
//    left, a tan Parent Square header. Student List: a reset leaves A:H (twice in a row too);
//    Your Lists headers after Last Interaction get the Status header's look
//  - 🐞 Debug → 🎨 Apply Sheet Style (z_MINE.gs) with a fake Sheets API: table colors and the
//    Send Out table's name, header rows, PCAR, the empty "Column 9", RAW Data -> Raw Data, the
//    Contact Log's fonts left alone except column D, its color rules redrawn (pale green
//    Workshop rows), Send Out's dropdown ranges reported;
//    without the Sheets API the table colors are listed to do by hand
const assert = require('assert');
const { loadProject } = require('./lib/load');
const { makeSpreadsheet, fakeSpreadsheetApp } = require('./lib/fakesheet');

const lum = h => [1, 3, 5].map(i => parseInt(h.slice(i, i + 2), 16) / 255)
  .map(v => (v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4)))
  .reduce((s, v, i) => s + v * [0.2126, 0.7152, 0.0722][i], 0);
const contrast = (a, b) => { const [x, y] = [lum(a), lum(b)].sort((p, q) => q - p); return (x + 0.05) / (y + 0.05); };
const calls = (ss, tab, fn) => ss._log.filter(c => c.tab === tab && c.fn === fn);
const plain = x => JSON.parse(JSON.stringify(x)); // arrays made inside the project's context, for deepStrictEqual

// ---------------------------------------------------------------- the colors themselves
{
  const g = loadProject({ noTrialFiles: true });
  const L = g.__get("SHEET_LOOK_");
  L.typeColors.forEach(c => {
    assert(contrast(c, "#ffffff") >= 4.5, c + " reads on white");
    assert(contrast(c, L.typing.band) >= 4.5, c + " reads on cream");
  });
  assert.strictEqual(new Set(L.typeColors).size, L.typeColors.length, "every Type color is different");
  assert(L.siteColors.length >= 12 && new Set(L.siteColors).size === L.siteColors.length, "12 different site colors");
  L.siteColors.forEach(c => assert(contrast(c, L.ink) >= 7, "text reads on site color " + c));
  assert(contrast("#ffffff", L.typing.header) >= 4.5 && contrast("#ffffff", L.built.header) >= 4.5, "white header text reads");
  assert(contrast(L.built.header, L.built.soft) >= 4.5 && contrast(L.typing.header, L.typing.subHeader) >= 4.5, "section rows read");
  assert(contrast("#ffffff", L.pcar.header) >= 4.5 && contrast("#ffffff", L.pcar.headerDark) >= 4.5
    && contrast(L.pcar.sectionText, L.pcar.section) >= 4.5, "PCAR's browns read");
  assert(lum(L.pcar.header) < lum(L.typing.header), "PCAR's brown is darker than the Contact Log's");
  console.log("SHEET_LOOK_: readable colors: ok");
}

// ---------------------------------------------------------------- Contact Log Type colors
{
  const ss = makeSpreadsheet(["Contact Log", "Backend"]);
  const app = fakeSpreadsheetApp(ss);
  const g = loadProject({ extra: { SpreadsheetApp: app }, noTrialFiles: true });
  const L = g.__get("SHEET_LOOK_");
  let types = [["Call"], ["Email"], ["Meeting"]];
  g.backendReadBlock_ = (s, key) => (key === "TYPES" ? types : []);
  const cl = ss.getSheetByName("Contact Log");
  const today = app.newConditionalFormatRule().whenFormulaSatisfied("=$B2 = TODAY()").setBackground("#d9ead3")
    .setRanges([cl.getRange("B2:B")]).build();
  const typeRules = () => plain(cl._cf.filter(r => r._st.text !== undefined).map(r => r._st.text + "=" + r._st.font));
  const want3 = ["Call=" + L.typeColors[0], "Email=" + L.typeColors[1], "Meeting=" + L.typeColors[2]];

  cl._cf = [today];
  g.applyTypeDropdownColors();
  assert.deepStrictEqual(typeRules(), want3, "each type gets its place's color");
  assert(cl._cf.includes(today), "the other rules are kept");
  const writes = cl._cfWrites;
  g.applyTypeDropdownColors();
  assert.strictEqual(cl._cfWrites, writes, "already right: nothing written");

  cl._cf = [today]; // resetContactLogFormatting (every reset / 10 PM run) clears them ...
  g.applyTypeDropdownColors();
  assert.deepStrictEqual(typeRules(), want3, "... and they come back the same, not new random ones");

  types = [["Call"], ["Meeting"]];
  g.applyTypeDropdownColors();
  assert.deepStrictEqual(typeRules(), ["Call=" + L.typeColors[0], "Meeting=" + L.typeColors[1]], "a type taken off the list loses its rule");
  console.log("Contact Log Type colors: fixed, kept, written only when needed: ok");
}

// ---------------------------------------------------------------- Events tables
{
  const ss = makeSpreadsheet(["Events", "Backend"]);
  const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) }, noTrialFiles: true });
  const L = g.__get("SHEET_LOOK_");
  g.eventDropdownRules_ = () => ({ osis: "OSIS-RULE", guardian: "GUARDIAN-RULE", student: "STUDENT-RULE" });
  g.applySheetLocks_ = () => {};
  g.backendReadBlock_ = () => [["Date"]];
  const ev = ss.getSheetByName("Events");

  g.buildEventTable("Fall PTC", "Gym", "11/12/2026", "PTC");
  assert.deepStrictEqual([1, 3, 4, 5].map(c => ev.getRange(1, c).getValue()), ["Fall PTC", "PTC", "Gym", "11/12/2026"], "row 1 values");
  assert.deepStrictEqual(ev.getRange(2, 1, 1, 6).getValues()[0], Array.from(g.__get("EVENT_TABLE_HEADERS")), "header row");
  [1, 2, 3, 4, 5].forEach(c => assert.strictEqual(ev.getRange(1, c).getBackground(), L.typing.header, "title row brown, column " + c));
  assert.strictEqual(ev.getRange(1, 6).getBackground(), L.danger.fill, "delete box soft red");
  assert.strictEqual(ev.getRange(2, 4).getBackground(), L.typing.subHeader, "header row light tan");
  assert(ss._log.some(c => c.tab === "Events" && c.fn === "merge" && c.row === 1 && c.col === 1 && c.nc === 2), "name over 2 columns");
  assert(calls(ss, "Events", "setFontFamily").some(c => c.row === 1 && c.col === 1 && c.nc === 6 && c.nr === ev.getMaxRows()
    && c.args[0] === "Montserrat"), "Montserrat on the whole table");
  const pastels = ["#d9ead3", "#DAA06D", "#fff2cc", "#c9daf8", "#ea9999", "#efefef"];
  assert(!ss._log.some(c => c.fn === "setBackground" && pastels.includes(c.args[0])), "none of the old pastel colors");
  assert(!ss._log.some(c => c.fn === "setBorder" && c.args.includes("#000000")), "no black grid");
  assert.deepStrictEqual(calls(ss, "Events", "setDataValidation").map(c => c.args[0]), ["OSIS-RULE", "GUARDIAN-RULE", "STUDENT-RULE"], "dropdowns as before");

  // 🎨 Apply Sheet Style's restyle: the look only
  let asked = 0;
  g.eventDropdownRules_ = () => { asked++; return {}; };
  ss._log.length = 0;
  assert.strictEqual(g.restyleEventTables_(ss, ev, true), 1, "one table found");
  assert.strictEqual(asked, 0, "look only: the dropdown lists are not even read");
  assert.strictEqual(calls(ss, "Events", "setDataValidation").length, 0, "look only: no dropdowns written");
  assert.strictEqual(ev.getRange(1, 1).getValue(), "Fall PTC", "values kept");
  console.log("Events tables: one look for new events, Migration and 🎨: ok");
}

// ---------------------------------------------------------------- Override Report, Raw Data, Student List
{
  const ss = makeSpreadsheet(["Students Override Report", "Raw Data", "Student List", "Backend"]);
  const g = loadProject({ extra: { SpreadsheetApp: fakeSpreadsheetApp(ss) }, noTrialFiles: true });
  const L = g.__get("SHEET_LOOK_");

  const rep = ss.getSheetByName("Students Override Report");
  g.styleOverrideReport_(rep, 3);
  assert.strictEqual(rep.getRange(1, 2).getBackground(), L.built.header, "report header navy");
  assert(calls(ss, "Students Override Report", "setFontFamily").some(c => c.nr === 3 && c.nc === 3 && c.args[0] === "Montserrat"), "report in Montserrat");

  const raw = ss.getSheetByName("Raw Data");
  raw._maxCols = 61;
  g.forceVibrantColors_(raw);
  assert(!ss._log.some(c => c.fn === "setFontFamily" && c.args[0] === "Roboto"), "no Roboto left on Raw Data");
  assert(calls(ss, "Raw Data", "setFontFamily").some(c => c.row === 1 && c.col === 1 && c.nr === raw.getMaxRows() && c.nc === 61
    && c.args[0] === "Montserrat"), "Montserrat on the whole tab, pasted rows too");
  [1, 33, 42].forEach(col => assert(calls(ss, "Raw Data", "setHorizontalAlignment")
    .some(c => c.row === 1 && c.col === col && c.args[0] === "left"), "DATE UPDATED title at the left, column " + col));
  assert.strictEqual(raw.getRange(6, 34).getBackground(), "#f6dfc4", "Parent Square header soft tan (AH6; AG6 is its yellow paste box)");
  assert.strictEqual(raw.getRange(6, 33).getBackground(), "#FFF2CC", "AG6 stays the yellow paste box");

  const sl = ss.getSheetByName("Student List");
  const head = ["OSIS", "Student", "Guardian", "Site", "Class", "Sibling", "Status", "Last Interaction"];
  sl._put([head.concat("Column 9"), ["123456789", "Bart Simpson", "Homer Simpson", "15", "Y41", "", "Current", "None logged", ""]], 1, 1);
  sl._maxCols = 9;
  g.resetMasterTable_(ss);
  assert.strictEqual(sl.getMaxColumns(), 8, "a reset leaves A:H");
  assert.deepStrictEqual(sl.getRange(1, 1, 1, 8).getValues()[0], head, "headers kept");
  assert.strictEqual(sl.getRange(2, 1).getValue(), "", "rows cleared");
  g.resetMasterTable_(ss);
  assert.strictEqual(sl.getMaxColumns(), 8, "a second reset is fine");
  assert.strictEqual(calls(ss, "Student List", "clearContent").slice(-1)[0].nc, 8,
    "the second reset clears A:H only (a range past the last column throws in Sheets)");
  sl._maxCols = 26;
  g.resetMasterTable_(ss);
  assert.strictEqual(sl.getMaxColumns(), 8, "26 columns -> 8");

  // Your Lists headers after Last Interaction look like the Status header
  const withLists = head.concat(["Bus Route", "PTA"]);
  sl._maxCols = 10;
  sl._put([withLists], 1, 1);
  sl.getRange(1, 7).setBackground("#183e69");
  g.styleLastInteractionHeader_(sl, withLists);
  [8, 9, 10].forEach(c => assert.strictEqual(sl.getRange(1, c).getBackground(), "#183e69", "list header " + c + " like Status"));
  sl.getRange(1, 1, 1, 10).setBackground(null);
  g.styleLastInteractionHeader_(sl, withLists);
  [8, 9, 10].forEach(c => assert.strictEqual(sl.getRange(1, c).getBackground(), "#ffffff", "Status with no fill: no fill (the table's color shows)"));
  assert.strictEqual(calls(ss, "Student List", "setBackground").slice(-1)[0].args[0], null, "never a white fill");
  console.log("Override Report, Raw Data, Student List A:H + list headers: ok");
}

// ---------------------------------------------------------------- 🐞 Debug → 🎨 Apply Sheet Style
const CL_BROWN = "#6a4d38";
function styleTestSheet(opts) {
  opts = opts || {};
  const ss = makeSpreadsheet(["Contact Log", "RAW Data", "Events", "Send Out", "Notes", "PCAR", "Combined Contact Tracking",
    "Directory", "Students Override Report", "Student List", "Parents Divided", "Phone Contacts", "Backend"]);
  const t = n => ss.getSheetByName(n);
  t("Contact Log")._put([["", "Date", "Parent", "Person Spoke With", "OSIS"]], 1, 1);
  t("Contact Log")._set(5, 5, "END");
  t("RAW Data")._maxCols = 61;
  t("RAW Data")._set(1, 1, "ATS - DATE UPDATED: 11/12/26");
  t("Events")._put([["Fall PTC", "", "PTC", "Gym", "11/12/2026", false]], 1, 1);
  t("Send Out")._maxCols = 8;
  t("Send Out")._put([["Date", "Send Out Title", "Send Out Type", "On Behalf Of", "Where Did You Send", "Population", "# Reached", "Attachments"]], 1, 1);
  t("Send Out").getRange(1, 4).setBackground("#b45f06");
  t("Send Out").getRange(1, 5).setBackground("#e06666");
  t("Notes")._maxCols = 5;
  t("Notes")._put([["OSIS", "Guardian(s)", "Student(s)", "Note", "Remove"]], 1, 1);
  const sl = t("Student List");
  sl._maxCols = opts.listColumn ? 10 : 9;
  sl._put([["OSIS", "Student", "Guardian", "Site", "Class", "Sibling", "Status", "Last Interaction", opts.listColumn ? "Bus Route" : "Column 9"],
    ["123456789", "Bart Simpson", "Homer Simpson", "15", "Y41: Tessa", "", "Current", "November 5, 2026", opts.listColumn ? "Route 7" : ""]], 1, 1);
  t("Parents Divided")._put([["student_id", "Guardian output", "Student", "Class / Site", "relationship", "language", "email", "mobile",
    "Sibling", "Parents Square", "NYCSA", "Last Interaction"],
    ["123456789", "Homer Simpson", "Bart Simpson", "Y41 (Site: 15)", "Father", "English", "H@x.com", "(111) 111-1111", "", "Yes", "No", "None logged"]], 1, 1);
  t("Directory")._put([["Last Name", "First Name", "OSIS", "Site", "Class/ Teacher", "Parent Contact"], ["Simpson", "Bart", "123456789", "15", "Y41: Tessa", "Homer"]], 1, 1);
  t("Combined Contact Tracking")._put([["OSIS", "Guardians", "Student", "Notes"], ["123456789", "Homer", "Bart", "a note"]], 1, 1);
  t("Phone Contacts")._put([["Name Prefix", "First Name"], ["", "(Father) Homer"]], 1, 1);
  t("Students Override Report")._put([["Student OSIS", "Old Information", "New Information"]], 1, 1);
  // PCAR: 3 boxes; the third box's first section title is missing (and not merged)
  const pc = t("PCAR");
  pc._put([["", "This Week", "", "", "Last Week", "", "", "PUT YOUR OWN INPUT", ""],
    ["", "November 8, 2026", "November 14, 2026", "", "November 1, 2026", "November 7, 2026", "", "October 11, 2026", "October 17, 2026"],
    ["", "Parent/Family Consultations", "", "", "Parent/Family Consultations", "", "", "", ""],
    ["", "Total number of phone consultations:", 9, "", "Total number of phone consultations:", 8, "", "Total number of phone consultations:", 0]], 1, 1);
  [2, 5, 8].forEach(c => pc.getRange(1, c, 1, 2).merge());
  [2, 5].forEach(c => pc.getRange(3, c, 1, 2).merge());
  // Send Out dropdowns: C reads Backend BH4 (one row too low), D is a typed list without its first choice, E is right
  const backend = t("Backend");
  const rangeRule = r => ({ getCriteriaType: () => "VALUE_IN_RANGE", getCriteriaValues: () => [r, true] });
  t("Send Out")._dv.set("2,3", rangeRule(backend.getRange("BH4:BH")));
  t("Send Out")._dv.set("2,4", { getCriteriaType: () => "VALUE_IN_LIST", getCriteriaValues: () => [["Principal", "PTA", "School"], true] });
  t("Send Out")._dv.set("2,5", rangeRule(backend.getRange("BJ3:BJ")));
  ss._log.length = 0;
  return ss;
}
function fakeSheetsApi() {
  const requests = [], c = h => ({ rgbColor: { red: parseInt(h.slice(1, 3), 16) / 255, green: parseInt(h.slice(3, 5), 16) / 255, blue: parseInt(h.slice(5, 7), 16) / 255 } });
  const table = (id, name, cols) => ({ tableId: id, name: name, range: { sheetId: 1, startRowIndex: 0, endRowIndex: 40, startColumnIndex: 0, endColumnIndex: cols } });
  const sheets = [
    { properties: { title: "Contact Log" }, tables: [Object.assign(table("t0", "Contact Log", 16),
      { rowsProperties: { headerColorStyle: c(CL_BROWN), firstBandColorStyle: c("#ffffff"), secondBandColorStyle: c("#fff6e5") } })] },
    { properties: { title: "Notes" }, tables: [table("t1", "Notes", 5)] },
    { properties: { title: "Send Out" }, tables: [table("t2", "Table1", 8)] },
    { properties: { title: "Student List" }, tables: [table("t3", "Student List", 9)] }, // before "Column 9" went
    { properties: { title: "Parents Divided" }, tables: [table("t4", "Parents Divided", 12)] },
    { properties: { title: "Directory" }, tables: [table("t5", "Directory", 6)] },
    { properties: { title: "Combined Contact Tracking" }, tables: [table("t6", "Contact Note Tracking", 4)] },
    { properties: { title: "Raw Data" } }, { properties: { title: "PCAR" } }
  ];
  return { _requests: requests, Spreadsheets: {
    get: (id, o) => { assert(/tables/.test(o.fields), "asks for the tables"); return { sheets: sheets }; },
    batchUpdate: (body, id) => { requests.push(...body.requests); return {}; } } };
}
const hexOf = cs => "#" + ["red", "green", "blue"].map(k => ("0" + Math.round(cs.rgbColor[k] * 255).toString(16)).slice(-2)).join("");

{
  const ss = styleTestSheet();
  const sheetsApi = fakeSheetsApi();
  const g = loadProject({ owner: true, extra: { SpreadsheetApp: fakeSpreadsheetApp(ss), Sheets: sheetsApi }, noTrialFiles: true });
  const L = g.__get("SHEET_LOOK_");
  const t = n => ss.getSheetByName(n);
  const rep = g.ownerApplySheetStyle_(ss);

  assert.deepStrictEqual(plain(rep.failed), [], "no step failed");
  assert(t("Raw Data") && !ss.getSheets().some(s => s.getName() === "RAW Data"), 'RAW Data -> Raw Data');
  assert.strictEqual(t("Student List").getMaxColumns(), 8, '"Column 9" deleted');
  assert.strictEqual(t("Student List").getRange(2, 8).getValue(), "November 5, 2026", "Student List values kept");

  // Sheets Tables: colors (typing = the Contact Log table's own brown) + Send Out's name
  const colorReqs = sheetsApi._requests.filter(r => r.updateTable.table.rowsProperties);
  const byId = {};
  colorReqs.forEach(r => { byId[r.updateTable.table.tableId] = r.updateTable.table.rowsProperties; });
  assert.deepStrictEqual(Object.keys(byId).sort(), ["t1", "t2", "t3", "t4", "t5", "t6"], "6 tables recolored, not the Contact Log");
  ["t1", "t2"].forEach(id => {
    assert.strictEqual(hexOf(byId[id].headerColorStyle), CL_BROWN, "typing tab " + id + ": the Contact Log's brown");
    assert.strictEqual(hexOf(byId[id].secondBandColorStyle), "#fff6e5", "typing tab " + id + ": the Contact Log's cream");
  });
  ["t3", "t4", "t5", "t6"].forEach(id => {
    assert.strictEqual(hexOf(byId[id].headerColorStyle), L.built.header, "built tab " + id + ": navy");
    assert.strictEqual(hexOf(byId[id].secondBandColorStyle), L.built.band, "built tab " + id + ": light grey");
  });
  colorReqs.forEach(r => assert(/rowsProperties\.headerColorStyle/.test(r.updateTable.fields), "field mask"));
  const rename = sheetsApi._requests.filter(r => r.updateTable.table.name);
  assert.deepStrictEqual(rename.map(r => [r.updateTable.table.tableId, r.updateTable.table.name, r.updateTable.fields]), [["t2", "Send Out", "name"]], 'Table1 -> "Send Out"');

  // Header rows: the table's color with white bold text, over Send Out's old orange / coral
  [4, 5].forEach(c => assert.strictEqual(t("Send Out").getRange(1, c).getBackground(), CL_BROWN, "Send Out header " + c));
  [1, 5].forEach(c => assert.strictEqual(t("Notes").getRange(1, c).getBackground(), CL_BROWN, "Notes header " + c));
  [[ "Student List", 8 ], ["Parents Divided", 12], ["Directory", 6], ["Combined Contact Tracking", 4]].forEach(([name, n]) => {
    for (let c = 1; c <= n; c++) assert.strictEqual(t(name).getRange(1, c).getBackground(), L.built.header, name + " header " + c);
    assert.strictEqual(t(name).getRange(1, n + 1).getBackground(), "#ffffff", name + ": nothing painted past the table");
  });
  assert.strictEqual(t("Phone Contacts").getRange(1, 2).getBackground(), L.built.header, "Phone Contacts header navy");

  // Fonts: every tab, the Contact Log only in column D (rows 2..END-1)
  ["Raw Data", "Events", "Notes", "Send Out", "PCAR", "Student List", "Directory", "Phone Contacts"].forEach(name =>
    assert(calls(ss, name, "setFontFamily").some(c => c.row === 1 && c.col === 1 && c.nr === t(name).getMaxRows() && c.args[0] === "Montserrat"), name + " in Montserrat"));
  const clFonts = calls(ss, "Contact Log", "setFontFamily");
  assert.deepStrictEqual(clFonts.map(c => [c.row, c.col, c.nr, c.nc, c.args[0]]), [[2, 4, 3, 1, "Montserrat"]], "Contact Log: column D only");

  // PCAR: one look, the missing title copied (and merged), the own-dates box brown with yellow date cells
  const pc = t("PCAR");
  assert.strictEqual(pc.getRange(3, 8).getValue(), "Parent/Family Consultations", "missing section title added");
  assert(pc._merges.some(m => m.r === 3 && m.c === 8 && m.nc === 2), "and merged over the box");
  [2, 5, 8].forEach(c => assert.strictEqual(pc.getRange(1, c).getBackground(), L.pcar.header, "box title darker brown, column " + c));
  assert.strictEqual(pc.getRange(2, 2).getBackground(), L.pcar.headerDark, "This Week dates darkest brown");
  assert.strictEqual(pc.getRange(2, 9).getBackground(), L.input, "own date cells yellow");
  [2, 5, 8].forEach(c => assert.strictEqual(pc.getRange(3, c).getBackground(), L.pcar.section, "section row light tan, column " + c));
  assert(!pc._bg.size || ![...pc._bg.values()].some(v => v === L.built.header), "no navy left on PCAR");
  assert.strictEqual(pc.getRange(4, 9).getValue(), 0, "values kept");
  assert(calls(ss, "PCAR", "setFontStyle").some(c => c.row === 4 && c.col === 2 && c.args[0] === "normal"), "labels not italic");

  // The parts the code paints anyway
  assert.strictEqual(t("Events").getRange(1, 3).getBackground(), L.typing.header, "event table repainted");
  assert(calls(ss, "Raw Data", "setHorizontalAlignment").some(c => c.row === 1 && c.col === 1 && c.args[0] === "left"), "Raw Data repainted");
  assert.strictEqual(t("Students Override Report").getRange(1, 1).getBackground(), L.built.header, "Override Report repainted");
  assert(L.siteColors.includes(t("Directory").getRange(2, 1).getBackground()), "Directory site color from the new set");
  const clRules = t("Contact Log")._cf.map(r => r._st);
  assert(clRules.some(r => r.formula === '=$J2="Workshop"' && r.bg === L.workshopRow), "Contact Log rules redrawn: pale green Workshop rows");
  assert(!clRules.some(r => r.bg === "#b6d7a8"), "the old solid Workshop green is gone");
  assert(clRules.some(r => /TODAY\(\)/.test(r.formula || "") && r.bg === "#d9ead3"), "the week / today colors stay as they were");

  // Reported: Send Out's dropdown ranges, the chip colors, the Contact Log widths
  assert(rep.byHand.some(l => /Send Out Type/.test(l) && /BH4/.test(l) && /Backend!BH3:BH/.test(l)), "C's list one row too low");
  assert(rep.byHand.some(l => /On Behalf Of/.test(l) && /Parent Coordinator/.test(l)), "D's typed list without its first choice");
  assert(!rep.byHand.some(l => /Where Did You Send/.test(l)), "E is right: nothing said");
  assert(rep.byHand.some(l => /chip colors/.test(l)) && rep.byHand.some(l => /Contact Log widths/.test(l)), "the by-hand rest");
  assert(!rep.byHand.some(l => /Table colors/.test(l)), "table colors were done, not listed");
  const text = g.ownerSheetStyleReport_(rep);
  assert(/^Done:/.test(text) && /Still by hand:/.test(text) && !/Didn't work/.test(text), "report text");

  // Running it again changes nothing important and fails nowhere
  const again = g.ownerApplySheetStyle_(ss);
  assert.deepStrictEqual(plain(again.failed), [], "safe to run again");
  assert.strictEqual(t("Student List").getMaxColumns(), 8, "still A:H");
  assert.strictEqual(sheetsApi._requests.filter(r => r.updateTable.table.name).length, 2, "the fake keeps saying Table1, so it is asked again (harmless)");
  console.log("🎨 Apply Sheet Style with the Sheets API: ok");
}

{
  // A Your Lists column after H stays; no Sheets API: table colors listed to do by hand
  const ss = styleTestSheet({ listColumn: true });
  const fetched = [];
  const g = loadProject({ owner: true, noTrialFiles: true, extra: {
    SpreadsheetApp: fakeSpreadsheetApp(ss),
    ScriptApp: { getOAuthToken: () => "token", getProjectTriggers: () => [] },
    UrlFetchApp: { fetch: (url, o) => { fetched.push(url); return { getResponseCode: () => 403,
      getContentText: () => JSON.stringify({ error: { message: "Google Sheets API has not been used in project 123 before or it is disabled." } }) }; } } } });
  const L = g.__get("SHEET_LOOK_");
  const t = n => ss.getSheetByName(n);
  const rep = g.ownerApplySheetStyle_(ss);
  assert.deepStrictEqual(plain(rep.failed), [], "no step failed");
  assert(fetched.length === 1 && /sheets\.googleapis\.com\/v4\/spreadsheets\/FAKE_SS\?fields=/.test(fetched[0]), "tried the web once");
  assert(rep.byHand.some(l => /^Table colors: Google didn't let the script change them \(Google Sheets API has not been used/.test(l)
    && l.includes(L.built.header) && l.includes(L.typing.header)), "table colors listed with the colors to use");
  assert.strictEqual(t("Student List").getMaxColumns(), 9, "the Your Lists column (with a header) stays, the empty one after it goes");
  assert.strictEqual(t("Student List").getRange(2, 9).getValue(), "Route 7", "with its values");
  assert.strictEqual(t("Student List").getRange(1, 9).getBackground(), L.built.header, "and gets the header look");
  assert.strictEqual(t("Notes").getRange(1, 1).getBackground(), L.typing.header, "headers painted with the house brown");
  console.log("🎨 Apply Sheet Style without the Sheets API + a Your Lists column: ok");
}

{
  // Google refuses the Phone Contacts banding change (owner's sheet 2026-11-13): the header is still painted,
  // the row colors are listed to do by hand, and no step counts as failed
  const ss = styleTestSheet();
  const refusal = "Service Spreadsheets failed while accessing document with id X.";
  ss.getSheetByName("Phone Contacts")._bandings = [{
    setHeaderRowColor() { throw new Error(refusal); }, setFirstRowColor() { return this; }, setSecondRowColor() { return this; } }];
  const g = loadProject({ owner: true, extra: { SpreadsheetApp: fakeSpreadsheetApp(ss), Sheets: fakeSheetsApi() }, noTrialFiles: true });
  const L = g.__get("SHEET_LOOK_");
  const rep = g.ownerApplySheetStyle_(ss);
  assert.deepStrictEqual(plain(rep.failed), [], "a refused banding is not a failed step");
  assert.strictEqual(ss.getSheetByName("Phone Contacts").getRange(1, 1).getBackground(), L.built.header, "the header is still navy");
  assert(rep.byHand.some(l => /^Phone Contacts row colors/.test(l) && l.includes(refusal) && l.includes(L.built.band)), "row colors listed by hand with the values");
  assert(rep.done.some(l => l === "Phone Contacts: navy header"), "reported as header only");
  console.log("🎨 Apply Sheet Style: a refused Phone Contacts banding: ok");
}

{
  // The menu item is there and the Mine manifest has the Sheets service
  const fs = require('fs'), path = require('path');
  const { ROOT } = require('./lib/load');
  const mine = fs.readFileSync(path.join(ROOT, 'z_MINE.gs'), 'utf8');
  assert(/addItem\("🎨 Apply Sheet Style[^"]*", "ownerMenuApplySheetStyle"\)/.test(mine), "🐞 Debug menu item");
  const manifest = JSON.parse(fs.readFileSync(path.join(ROOT, 'App Scripts (Mine).Json'), 'utf8'));
  assert(manifest.dependencies.enabledAdvancedServices.some(s => s.userSymbol === "Sheets" && s.serviceId === "sheets"), "Mine manifest: Sheets service");
  const general = JSON.parse(fs.readFileSync(path.join(ROOT, 'App Scripts (General).Json'), 'utf8'));
  assert.deepStrictEqual(manifest.oauthScopes.filter(s => !general.oauthScopes.includes(s)).sort(),
    ["https://www.googleapis.com/auth/script.send_mail", "https://www.googleapis.com/auth/userinfo.email"], "no new permission");
  console.log("menu item + manifest: ok");
}
