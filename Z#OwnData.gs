// ======================================================================
// FILE: Z#OwnData.gs (+ Z#OwnDataSidebar.html)
// PURPOSE: "Your Lists", step 5 of 📥 Input Data Mode. Any outside list (bus
// routes, trip forms, after school, PTA ...) is pasted on its own tab
// ("📥 Bus Routes"). The tool finds the header row, guesses what each column
// is, matches every row to a student (student lists) or a guardian (guardian
// lists), and the columns you keep ("variables") go to:
//   - the Backend tab's ADDITIONAL INFORMATION section (OSIS / Student /
//     Guardian / lines), which the phone contacts show as "Additional
//     Information:" above Communications (readAdditionalInfo_, Build.gs);
//   - the Student List (student lists) or Parents Divided (guardian lists),
//     one column per variable.
// Everyone on a list gets it (owner's pick "A"); Combined Contact Tracking is
// not changed.
//
// List tab layout: row 1 title band, row 2 status band, row 3 what each column
// is, row 4 the list's header row (a paste is moved there, title rows above the
// headers dropped), data from row 5, then a spacer and the "✔ Matched" / "How"
// columns the tool fills. The setup (columns, variables, your picks) is one JSON
// cell per list in the Backend's YOUR LISTS (SETUP) section.
//
// Every build (od_buildOutputs_, one guarded line in buildAllDerivedSheets)
// matches every list again, so new students, sorting or a fixed name are picked
// up; rows only a person can decide (two possible students, a best guess you
// haven't accepted, not found) stay unmatched until you decide in the sidebar.
// Your decisions are kept per row (by what the row says) for the next paste.
//
// SELF-CONTAINED ("Z#" = trying out a new feature): shared code only reaches it
// through the guarded od_buildOutputs_ line; Z#InputDataMode.gs shows the lists
// in its step 5 (guarded too). Without this file nothing changes.
// ======================================================================

var OD_TAB_PREFIX = "📥 ";
var OD_PLACEHOLDER = "PASTE HERE";
var OD_HEADER_ROW = 4;               // a list's header row after a paste; data from row 5
var OD_WIZARD_KEY = "OD_WIZARD_TAB"; // tab id the setup sidebar works on
var OD_BAND_TITLE = "📥 OWN DATA · ";
var OD_TOOL_HEADERS = ["✔ Matched", "How"];
var OD_COLORS = {
  band1: "#134f5c", band2: "#2f6b77", chips: "#fbfcfe", header: "#d0e0e3", paste: "#FFF2CC",
  ok: "#e6f4ea", guess: "#fef3d6", none: "#fce8e6", skip: "#f1f3f4"
};
var OD_HEADER_WORDS_ = /osis|student|name|first|last|parent|guardian|mother|father|phone|cell|mobile|e-?mail|class|grade|\bid\b|route|bus|stop|date|address|site|room|teacher/i;

// ======================================================================
// SETUP STORAGE (Backend "YOUR LISTS (SETUP)": Tab ID | List | Setup | Updated)
// ======================================================================

function od_emptySetup_() {
  return { v: 1, ready: false, by: "student", roles: {}, vars: [], picks: {}, headers: [], title: "", stats: null };
}

function od_fillSetup_(s) {
  const e = od_emptySetup_();
  s = s && typeof s === "object" ? s : {};
  Object.keys(e).forEach(k => { if (s[k] === undefined) s[k] = e[k]; });
  if (!s.picks || typeof s.picks !== "object") s.picks = {};
  if (!Array.isArray(s.vars)) s.vars = [];
  return s;
}

/** Every list: { tabId, name, setup, updated, sheet } (sheet = null when its tab is gone). */
function od_loadLists_(ss) {
  return backendReadBlock_(ss, "OWN_LISTS", true).slice(1)
    .filter(r => String(r[0]).trim() !== "")
    .map(r => {
      let setup = null;
      try { setup = JSON.parse(String(r[2] || "{}")); } catch (e) { setup = null; }
      const tabId = Number(r[0]);
      return { tabId: tabId, name: toText_(r[1]), setup: od_fillSetup_(setup), updated: toText_(r[3]), sheet: od_sheetById_(ss, tabId) };
    });
}

/** Writes the lists back (a list whose tab was deleted by hand is dropped). */
function od_saveLists_(ss, lists) {
  const rows = lists.filter(l => l.sheet).map(l => {
    let json = JSON.stringify(l.setup);
    // A cell holds 50,000 characters: forget the oldest decisions first if a list ever gets there
    while (json.length > 45000 && Object.keys(l.setup.picks).length) {
      delete l.setup.picks[Object.keys(l.setup.picks)[0]];
      json = JSON.stringify(l.setup);
    }
    return [l.tabId, l.name, json, l.updated || ""];
  });
  backendWriteBlock_(ss, "OWN_LISTS", [["Tab ID", "List", "Setup", "Updated"]].concat(rows));
}

function od_findList_(lists, tabId) {
  return lists.filter(l => l.tabId === Number(tabId))[0] || null;
}

// ======================================================================
// LIST TABS
// ======================================================================

function od_sheetById_(ss, id) {
  if (id === null || id === undefined || id === "" || isNaN(Number(id))) return null;
  return ss.getSheets().filter(s => s.getSheetId() === Number(id))[0] || null;
}

function od_uniqueTabName_(ss, base, self) {
  let name = base, n = 2;
  for (;;) {
    const other = ss.getSheetByName(name);
    if (!other || (self && other.getSheetId() === self.getSheetId())) return name;
    name = base + " (" + (n++) + ")";
  }
}

function od_cleanName_(name) {
  return String(name || "").replace(/[\[\]:*?\/\\'"]/g, " ").replace(/\s+/g, " ").trim().slice(0, 60);
}

/** A variable's name ("Bus Route"): one line, no ":" (the line reads "Name: value"); quotes are fine. */
function od_cleanLabel_(label) {
  return String(label || "").replace(/[:\r\n\t]/g, " ").replace(/\s+/g, " ").trim().slice(0, 60);
}

function od_colLetter_(c) {
  let s = "";
  while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
  return s;
}

function od_today_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yy");
}

/** Rows 1-3 (title, status, what each column is). */
function od_writeBands_(sheet, name, updated, status, chips) {
  const width = Math.max(chips.length, 8);
  if (sheet.getMaxColumns() < width) sheet.insertColumnsAfter(sheet.getMaxColumns(), width - sheet.getMaxColumns());
  const row1 = new Array(width).fill(""), row2 = new Array(width).fill(""), row3 = new Array(width).fill("");
  row1[0] = OD_BAND_TITLE + String(name).toUpperCase() + " — DATE UPDATED: " + (updated || "00/00/00");
  row2[0] = status;
  chips.forEach((c, i) => { row3[i] = c || ""; });
  sheet.getRange(1, 1, 3, width).setValues([row1, row2, row3]);
  sheet.getRange(1, 1, 1, width).setBackground(OD_COLORS.band1).setFontColor("#ffffff").setFontWeight("bold");
  sheet.getRange(2, 1, 1, width).setBackground(OD_COLORS.band2).setFontColor("#ffffff").setFontWeight("normal");
  sheet.getRange(3, 1, 1, width).setBackground(OD_COLORS.chips).setFontColor("#174ea6").setFontWeight("bold").setFontSize(9);
}

/** An empty list tab: bands + the yellow PASTE HERE box in A4. */
function od_resetListTab_(sheet, name) {
  try { sheet.setFrozenRows(0); } catch (e) { /* none */ }
  sheet.getDataRange().clearDataValidations();
  sheet.clear();
  od_writeBands_(sheet, name, "", "Paste your whole list (headers too) in the yellow box (A4) as values: Ctrl+Shift+V. The sidebar sets it up.", []);
  sheet.getRange(OD_HEADER_ROW, 1).setValue(OD_PLACEHOLDER).setBackground(OD_COLORS.paste).setFontWeight("bold")
    .setNote("Paste your whole list here (headers too). Title rows above the headers are fine.");
  sheet.setColumnWidth(1, 160);
  try { sheet.setTabColor(OD_COLORS.band1); } catch (e) { /* cosmetic */ }
}

/** Index of the first row that is part of the paste (rows 1-3 are ours while the title band is still in A1). */
function od_pasteStart_(values) {
  const bandIntact = values.length > 0 && String(values[0][0]).indexOf(OD_BAND_TITLE) === 0;
  return bandIntact ? OD_HEADER_ROW - 1 : 0;
}

function od_hasPaste_(sheet) {
  const values = sheet.getDataRange().getValues();
  const start = od_pasteStart_(values);
  for (let r = start; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const v = values[r][c];
      if (v === "" || v === null) continue;
      if (r === OD_HEADER_ROW - 1 && c === 0 && String(v).trim() === OD_PLACEHOLDER) continue;
      return true;
    }
  }
  return false;
}

/**
 * Header row of a pasted block (display values): among the first 15 rows, the
 * first clearly header-like one (2+ cells, mostly words, about as wide as the
 * table, column-name words like "Name" / "OSIS" / "Phone" help). Title rows
 * ("District 75 ... Fall Routes", "Printed 11/03/26") have one cell and are skipped.
 */
function od_findHeaderRow_(d) {
  const width = r => (d[r] || []).filter(x => String(x).trim() !== "").length;
  let maxWidth = 0;
  for (let r = 0; r < Math.min(30, d.length); r++) maxWidth = Math.max(maxWidth, width(r));
  let best = 0, bestScore = -1;
  for (let r = 0; r < Math.min(15, d.length); r++) {
    const cells = d[r].map(x => String(x).trim()).filter(Boolean);
    if (!cells.length || (cells.length < 2 && maxWidth >= 2)) continue;
    const words = cells.filter(x => !/^[\d\s.,:\/()+#$%-]+$/.test(x)).length / cells.length;
    const wide = cells.length / maxWidth;
    if (words < 0.6 || wide < 0.5) continue;
    const keywords = Math.min(4, cells.filter(x => OD_HEADER_WORDS_.test(x)).length);
    const score = wide + words + keywords * 0.5 + (width(r + 1) ? 0.5 : 0);
    if (score > bestScore + 0.25) { best = r; bestScore = score; }
  }
  return best;
}

/**
 * Moves what was pasted (anywhere on the tab) into the fixed layout: header row
 * on row 4, data rows under it (empty rows dropped), title rows above the
 * header remembered as the list's title. Cells are written as the text you saw
 * (plain text), so a leading zero or a phone format is kept.
 * Returns { headers, rows, title } or null when the tab is empty.
 */
function od_takePaste_(sheet, name) {
  const range = sheet.getDataRange();
  const values = range.getValues(), displays = range.getDisplayValues();
  const start = od_pasteStart_(values);
  const block = [];
  for (let r = start; r < displays.length; r++) {
    const row = displays[r].map(x => String(x));
    if (r === OD_HEADER_ROW - 1 && row[0].trim() === OD_PLACEHOLDER) row[0] = "";
    block.push(row);
  }
  let minC = Infinity, maxC = -1, minR = Infinity, maxR = -1;
  block.forEach((row, r) => row.forEach((v, c) => {
    if (v.trim() === "") return;
    minC = Math.min(minC, c); maxC = Math.max(maxC, c); minR = Math.min(minR, r); maxR = Math.max(maxR, r);
  }));
  if (maxC < 0) return null;
  const d = block.slice(minR, maxR + 1).map(r => r.slice(minC, maxC + 1).map(x => x.trim()));
  const h = od_findHeaderRow_(d);
  const title = d.slice(0, h).map(r => r.filter(Boolean).join(" ")).filter(Boolean).join(" · ");
  const headers = d[h].map((x, i) => x || ("Column " + od_colLetter_(i + 1)));
  const rows = d.slice(h + 1).filter(r => r.some(x => x !== ""));

  sheet.getDataRange().clearDataValidations();
  sheet.clear();
  const width = headers.length;
  const needRows = OD_HEADER_ROW + rows.length + 1, needCols = width + 3;
  if (sheet.getMaxRows() < needRows) sheet.insertRowsAfter(sheet.getMaxRows(), needRows - sheet.getMaxRows());
  if (sheet.getMaxColumns() < needCols) sheet.insertColumnsAfter(sheet.getMaxColumns(), needCols - sheet.getMaxColumns());
  const table = [headers].concat(rows);
  sheet.getRange(OD_HEADER_ROW, 1, table.length, width).setNumberFormat("@").setValues(table);
  sheet.getRange(OD_HEADER_ROW, 1, 1, width).setFontWeight("bold").setBackground(OD_COLORS.header);
  try { sheet.setFrozenRows(OD_HEADER_ROW); } catch (e) { /* cosmetic */ }
  od_writeBands_(sheet, name, od_today_(), "Pasted: " + rows.length + " row(s). Finish the setup in the sidebar.", headers.map(() => ""));
  return { headers: headers, rows: rows, title: title };
}

/** The list's table as it is on the tab now: { headers, rows } (display text; tool columns left out). */
function od_readTable_(sheet, setup) {
  const lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (lastRow < OD_HEADER_ROW || lastCol < 1) return { headers: [], rows: [] };
  const width = Math.min(lastCol, setup && setup.headers && setup.headers.length ? setup.headers.length : lastCol);
  const d = sheet.getRange(OD_HEADER_ROW, 1, lastRow - OD_HEADER_ROW + 1, width).getDisplayValues();
  let headers = d[0].map(x => String(x).trim());
  const tool = headers.indexOf(OD_TOOL_HEADERS[0]);
  if (tool !== -1) headers = headers.slice(0, Math.max(0, tool - 1)); // spacer before the tool columns
  const rows = d.slice(1).map(r => r.slice(0, headers.length).map(x => String(x).trim()));
  return { headers: headers, rows: rows };
}

// ======================================================================
// STUDENTS / GUARDIANS INDEX (what a row can be matched to)
// ======================================================================

/** "SIMPSON, BART" / "Bart Simpson" -> { raw, tokens, first, last } (lowercase parts), or null. */
function od_nameParts_(text) {
  const s = toText_(text);
  if (!s) return null;
  let firstT, lastT;
  if (s.indexOf(",") !== -1) {
    const i = s.indexOf(",");
    lastT = nameTokens_(s.slice(0, i));
    firstT = nameTokens_(s.slice(i + 1));
  } else {
    const t = nameTokens_(s);
    firstT = t.slice(0, 1);
    lastT = t.slice(1);
  }
  const tokens = firstT.concat(lastT);
  if (!tokens.length) return null;
  return { raw: s, tokens: tokens, first: firstT[0] || "", last: lastT.length ? lastT[lastT.length - 1] : "" };
}

/** Name from two columns (first name, last name). */
function od_namePartsTwo_(firstText, lastText) {
  const firstT = nameTokens_(firstText), lastT = nameTokens_(lastText);
  const tokens = firstT.concat(lastT);
  if (!tokens.length) return null;
  return { raw: (toText_(firstText) + " " + toText_(lastText)).trim(), tokens: tokens, first: firstT[0] || "", last: lastT.length ? lastT[lastT.length - 1] : "" };
}

function od_tokenKey_(tokens) {
  return tokens.slice().sort().join(" ");
}

function od_addTo_(map, key, item) {
  if (!key) return;
  if (!map.has(key)) map.set(key, []);
  const list = map.get(key);
  if (list.indexOf(item) === -1) list.push(item);
}

/**
 * studentRows: Student List rows (OSIS, Student, Guardian, Site, Class, ...);
 * guardianRows: buildGuardianObjects_ output (Raw Data + overrides) for phones
 * and emails. Students who aren't in Raw Data (Discharged, the owner's Old
 * Master Table) get their guardians from the Student List's Guardian text.
 */
function od_index_(studentRows, guardianRows) {
  const idx = {
    byOsis: new Map(), students: [], byNameKey: new Map(), byToken: new Map(), byFirst: new Map(), byLast: new Map(),
    persons: [], pByKey: new Map(), pByNameKey: new Map(), pByToken: new Map(), pByLast: new Map(), byPhone: new Map(), byEmail: new Map(), classes: new Set()
  };
  studentRows.forEach(r => {
    const osis = toText_(r[0]);
    if (!osis || osis.indexOf(",") !== -1 || !isRealStudentId_(osis)) return;
    const key = normalizeOsis_(osis);
    if (idx.byOsis.has(key)) return;
    const parts = od_nameParts_(r[1]) || { tokens: [], first: "", last: "" };
    const st = {
      osis: osis, key: key, name: toText_(r[1]), tokens: parts.tokens, first: parts.first, last: parts.last,
      cls: toText_(r[4]).split(/[:\/]/)[0].trim().toUpperCase(), guardianText: toText_(r[2]), persons: []
    };
    idx.byOsis.set(key, st);
    idx.students.push(st);
    od_addTo_(idx.byNameKey, od_tokenKey_(st.tokens), st);
    st.tokens.forEach(t => od_addTo_(idx.byToken, t, st));
    od_addTo_(idx.byFirst, st.first, st);
    od_addTo_(idx.byLast, st.last, st);
    if (st.cls) idx.classes.add(st.cls);
  });

  const addPerson = (st, name, phone, email) => {
    const parts = od_nameParts_(name);
    if (!parts) return;
    const nameKey = od_tokenKey_(parts.tokens);
    // Same name = same person, unless both have phones / emails and none of them agree
    let person = (idx.pByNameKey.get(nameKey) || []).filter(p => {
      const phonesDiffer = phone && p.phones.size && !p.phones.has(phone);
      const emailsDiffer = email && p.emails.size && !p.emails.has(email);
      const phonesSame = phone && p.phones.has(phone), emailsSame = email && p.emails.has(email);
      return phonesSame || emailsSame || !(phonesDiffer || emailsDiffer);
    })[0];
    if (!person) {
      const n = (idx.pByNameKey.get(nameKey) || []).length;
      person = { key: nameKey + (n ? "#" + (n + 1) : ""), name: name, tokens: parts.tokens, first: parts.first, last: parts.last,
                 phones: new Set(), emails: new Set(), kids: [] };
      idx.persons.push(person);
      idx.pByKey.set(person.key, person);
      od_addTo_(idx.pByNameKey, nameKey, person);
      od_addTo_(idx.pByLast, person.last, person);
      parts.tokens.forEach(t => od_addTo_(idx.pByToken, t, person));
    }
    if (phone) { person.phones.add(phone); od_addTo_(idx.byPhone, phone, person); }
    if (email) { person.emails.add(email); od_addTo_(idx.byEmail, email, person); }
    if (person.kids.indexOf(st) === -1) person.kids.push(st);
    if (st.persons.indexOf(person) === -1) st.persons.push(person);
  };

  guardianRows.forEach(g => {
    const st = idx.byOsis.get(normalizeOsis_(g.student_id));
    if (!st) return;
    addPerson(st, stripTrailingParenthetical_(g.guardianOutput), normalizePhoneForMatch_(g.mobile), normalizeEmail_(g.email));
  });
  idx.students.forEach(st => {
    if (st.persons.length || !st.guardianText) return;
    st.guardianText.split(/\s+and\s+|,/i).map(x => stripTrailingParenthetical_(x).trim()).filter(Boolean)
      .forEach(name => addPerson(st, name, "", ""));
  });
  return idx;
}

/** Student List rows (row 5 down) + Raw Data guardians, for the sidebar (outside a build). */
function od_indexFromSheet_(ss) {
  const sl = getStudentListSheet_(ss);
  const last = sl ? sl.getLastRow() : 0;
  const studentRows = last >= 5 ? sl.getRange(5, 1, last - 4, 5).getValues() : [];
  const raw = ss.getSheetByName("Raw Data") || ss.getSheetByName("RAW Data");
  let guardianRows = [];
  if (raw && raw.getLastRow() > 1) {
    const overrides = getOverrideData_(ss);
    guardianRows = buildGuardianObjects_(raw.getDataRange().getValues(), overrides.guardians);
  }
  return od_index_(studentRows, guardianRows);
}

// ======================================================================
// MATCHING
// ======================================================================

// Common short names (both ways). "Sam" / "Mira" style short names that START
// the full name are caught without this list (od_firstLike_).
var OD_NICKNAMES_ = {
  william: "will bill billy liam", robert: "rob bob bobby robbie", richard: "rich rick ricky dick", james: "jim jimmy jamie",
  john: "jack johnny", jonathan: "jon jonny", joseph: "joe joey", michael: "mike mikey mick", christopher: "chris topher",
  christina: "chris tina", christine: "chris tina", elizabeth: "liz lizzie beth betsy eliza lisa", katherine: "kate katie kathy kat",
  catherine: "cathy cat kate", margaret: "maggie meg peggy", jennifer: "jen jenny", jessica: "jess jessie", alexander: "alex xander",
  alexandra: "alex lexi sandra", anthony: "tony", benjamin: "ben benny", daniel: "dan danny", david: "dave davey", edward: "ed eddie ted",
  matthew: "matt", nicholas: "nick nicky", patrick: "pat", samuel: "sam sammy", samantha: "sam sammy", stephen: "steve", steven: "steve",
  thomas: "tom tommy", timothy: "tim timmy", victoria: "vicky tori", zachary: "zach zack", gabriel: "gabe", abigail: "abby",
  rebecca: "becky", deborah: "debbie deb", susan: "sue susie", jacqueline: "jackie", francisco: "paco pancho", jose: "pepe",
  guillermo: "memo", alejandro: "alex", eduardo: "lalo eddie", gerardo: "jerry"
};
var OD_NICK_LOOKUP_ = null;
function od_nicknames_() {
  if (OD_NICK_LOOKUP_) return OD_NICK_LOOKUP_;
  OD_NICK_LOOKUP_ = {};
  Object.keys(OD_NICKNAMES_).forEach(full => {
    const group = [full].concat(OD_NICKNAMES_[full].split(" "));
    group.forEach(a => { OD_NICK_LOOKUP_[a] = OD_NICK_LOOKUP_[a] || {}; group.forEach(b => { if (a !== b) OD_NICK_LOOKUP_[a][b] = full; }); });
  });
  return OD_NICK_LOOKUP_;
}

/** How two first names are alike: "same", "start" (Sam / Samuel), "nickname" (Bill / William), "initial" (D / Diana) or "". */
function od_firstLike_(a, b) {
  if (!a || !b) return "";
  if (a === b) return "same";
  if ((a.length === 1 && b.indexOf(a) === 0) || (b.length === 1 && a.indexOf(b) === 0)) return "initial";
  const short = a.length <= b.length ? a : b, long = short === a ? b : a;
  if (short.length >= 3 && long.indexOf(short) === 0) return "start";
  const nick = od_nicknames_();
  if (nick[a] && nick[a][b]) return "nickname";
  return "";
}

function od_lastLike_(a, b) {
  if (!a || !b) return "";
  if (a === b) return "same";
  if ((a.length === 1 && b.indexOf(a) === 0) || (b.length === 1 && a.indexOf(b) === 0)) return "initial";
  return "";
}

/** Why a short first name is a best guess, in words: "“Mira” is how “Mirabel” starts". */
function od_firstWhy_(how, rowFirst, name) {
  const full = String(name || "").split(/\s+/)[0];
  const cap = s => s ? s.charAt(0).toUpperCase() + s.slice(1) : s;
  if (how === "start") return "“" + cap(rowFirst) + "” is how “" + full + "” starts";
  if (how === "nickname") return "“" + cap(rowFirst) + "” is a short name for “" + full + "”";
  if (how === "initial") return "“" + cap(rowFirst) + ".” could be “" + full + "”";
  return "";
}

/** The identity cells of one list row, already split into name parts / phone digits. */
function od_rowCells_(row, roles) {
  const at = c => (c ? toText_(row[c - 1]) : "");
  const name = r => {
    if (!r) return null;
    if (r.first || r.last) return od_namePartsTwo_(at(r.first), at(r.last));
    return od_nameParts_(at(r.col));
  };
  return {
    osis: at(roles.osis),
    name: name(roles.student),
    guardian: name(roles.guardian),
    phone: normalizePhoneForMatch_(at(roles.phone)),
    email: normalizeEmail_(at(roles.email)),
    cls: at(roles["class"]).toUpperCase()
  };
}

/** What a row says about who it is, short (for the remembered decisions). */
function od_rowKey_(cells) {
  const key = JSON.stringify([normalizeOsis_(cells.osis), cells.name ? od_tokenKey_(cells.name.tokens) : "",
    cells.guardian ? od_tokenKey_(cells.guardian.tokens) : "", cells.phone, cells.email]);
  // Two different 32-bit hashes -> a short, practically unique key
  let h1 = 0x811c9dc5, h2 = 0x01000193;
  for (let i = 0; i < key.length; i++) {
    const c = key.charCodeAt(i);
    h1 = Math.imul(h1 ^ c, 0x01000193) >>> 0;
    h2 = Math.imul(h2 + c, 0x5bd1e995) >>> 0;
  }
  return h1.toString(36) + h2.toString(36);
}

function od_ok_(target, how) { return { status: "ok", target: target, how: how }; }
function od_guess_(target, why) { return { status: "guess", target: target, how: "Best guess", reason: why }; }
function od_pick_(targets, why) { return { status: "pick", options: targets, how: "Check", reason: why }; }
function od_none_(why) { return { status: "none", how: "—", reason: why }; }

/** Guardian name fits a person: same name, one name inside the other, or same last name with a short / initial first name. */
function od_personFits_(parts, p) {
  if (!parts) return "";
  if (sameNameExactly_(parts.tokens, p.tokens)) return "same";
  if (namesMatch_(parts.tokens, p.tokens)) return "inside";
  const f = od_firstLike_(parts.first, p.first), l = od_lastLike_(parts.last, p.last);
  return f && l && !(f === "initial" && l === "initial") ? f + "/" + l : "";
}

/** Guardians worth comparing with a name: sharing a name part, the first name or the last name (not everyone). */
function od_personCandidates_(parts, idx) {
  const seen = [];
  const add = p => { if (seen.indexOf(p) === -1) seen.push(p); };
  parts.tokens.forEach(t => (idx.pByToken.get(t) || []).forEach(add));
  if (parts.last) (idx.pByLast.get(parts.last) || []).forEach(add);
  if (parts.last && parts.last.length === 1) idx.persons.forEach(p => { if (p.first === parts.first) add(p); }); // "Martha K."
  return seen;
}

/** Students a row's phone / email belongs to (through their guardians). */
function od_studentsByContact_(cells, idx) {
  const persons = [];
  if (cells.phone) (idx.byPhone.get(cells.phone) || []).forEach(p => { if (persons.indexOf(p) === -1) persons.push(p); });
  if (cells.email) (idx.byEmail.get(cells.email) || []).forEach(p => { if (persons.indexOf(p) === -1) persons.push(p); });
  const kids = [];
  persons.forEach(p => p.kids.forEach(k => { if (kids.indexOf(k) === -1) kids.push(k); }));
  return { persons: persons, kids: kids };
}

/** Best-guess students for a name (never used without a person saying yes). */
function od_studentGuesses_(cells, idx) {
  const out = [];
  const add = (st, why) => { if (!out.some(o => o.st === st)) out.push({ st: st, why: why }); };
  const n = cells.name;
  if (n) {
    const seen = new Set();
    n.tokens.forEach(t => (idx.byToken.get(t) || []).forEach(st => seen.add(st)));
    if (n.first) (idx.byFirst.get(n.first) || []).forEach(st => seen.add(st));
    // a short first name ("Mira") with no shared token yet: students with that last name
    if (n.last) (idx.byLast.get(n.last) || []).forEach(st => seen.add(st));
    seen.forEach(st => {
      if (namesMatch_(n.tokens, st.tokens) && !sameNameExactly_(n.tokens, st.tokens)) { add(st, "“" + n.raw + "” is part of " + st.name); return; }
      if (n.last) {
        const f = od_firstLike_(n.first, st.first), l = od_lastLike_(n.last, st.last);
        if (f && l && !(f === "initial" && l === "initial")) {
          add(st, l === "initial" ? "Same first name, “" + n.last.toUpperCase() + ".” could be " + st.last.charAt(0).toUpperCase() + st.last.slice(1)
                                  : (f === "same" ? "Same name" : od_firstWhy_(f, n.first, st.name) + ", same last name"));
        }
      } else if (n.tokens.length === 1 && st.first === n.first) {
        add(st, "First name only");
      }
    });
  }
  if (!out.length && cells.guardian) {
    // No student name to go on: the parent's name ("Martha K.") -> their children
    od_personCandidates_(cells.guardian, idx).forEach(p => {
      const fit = od_personFits_(cells.guardian, p);
      if (fit) p.kids.forEach(st => add(st, "Parent “" + cells.guardian.raw + "” looks like " + p.name));
    });
  }
  return out;
}

/** One row of a STUDENT list -> ok / guess / pick / none (targets = normalized OSIS). */
function od_matchStudentRow_(cells, idx) {
  let osisUnknown = false;
  if (cells.osis && normalizeOsis_(cells.osis)) {
    const st = idx.byOsis.get(normalizeOsis_(cells.osis));
    if (st) return od_ok_(st.key, "OSIS");
    osisUnknown = true;
  }
  const exact = cells.name ? (idx.byNameKey.get(od_tokenKey_(cells.name.tokens)) || []) : [];
  const contact = od_studentsByContact_(cells, idx);
  const narrow = list => list.filter(st =>
    (cells.guardian && st.persons.some(p => od_personFits_(cells.guardian, p))) ||
    contact.kids.indexOf(st) !== -1 || (cells.cls && st.cls === cells.cls));

  if (exact.length === 1) {
    if (!osisUnknown) return od_ok_(exact[0].key, "Name");
    return od_guess_(exact[0].key, "Same name, but OSIS " + cells.osis + " isn't on the Student List");
  }
  if (exact.length > 1) {
    const n = narrow(exact);
    if (n.length === 1 && !osisUnknown) return od_ok_(n[0].key, "Name + parent");
    return od_pick_((n.length ? n : exact).map(st => st.key), exact.length + " students named " + exact[0].name);
  }
  if (contact.kids.length === 1 && !osisUnknown) {
    // One family with one child: the phone / email says who it is (a name, if any, must not say otherwise)
    const st = contact.kids[0];
    if (!cells.name || od_studentGuesses_(cells, idx).some(g => g.st === st) || !cells.name.last) return od_ok_(st.key, cells.phone ? "Phone" : "Email");
  }
  if (contact.kids.length > 1) {
    const byName = cells.name ? od_studentGuesses_(cells, idx).filter(g => contact.kids.indexOf(g.st) !== -1) : [];
    if (byName.length === 1 && !osisUnknown) return od_ok_(byName[0].st.key, (cells.phone ? "Phone" : "Email") + " + name");
  }
  let guesses = od_studentGuesses_(cells, idx);
  if (guesses.length > 1) {
    const n = guesses.filter(g => narrow([g.st]).length);
    if (n.length) guesses = n;
  }
  if (guesses.length === 1) return od_guess_(guesses[0].st.key, guesses[0].why + (osisUnknown ? " (OSIS " + cells.osis + " isn't on the Student List)" : ""));
  if (guesses.length > 1) {
    const first = cells.name ? cells.name.raw : (cells.guardian ? cells.guardian.raw : "this row");
    return od_pick_(guesses.map(g => g.st.key), guesses.length + " possible students for “" + first + "”");
  }
  if (contact.kids.length > 1) return od_pick_(contact.kids.map(st => st.key), "That phone / email belongs to " + contact.kids.length + " students");
  if (osisUnknown) return od_none_("OSIS " + cells.osis + " isn't on the Student List");
  if (!cells.name && !cells.guardian && !cells.phone && !cells.email) return od_none_("This row has no OSIS, name, phone or email");
  return od_none_("No student on the Student List matches this row");
}

/** One row of a GUARDIAN list -> ok / guess / pick / none (targets = person keys). */
function od_matchGuardianRow_(cells, idx) {
  if (cells.osis && normalizeOsis_(cells.osis)) {
    const st = idx.byOsis.get(normalizeOsis_(cells.osis));
    if (st) {
      if (cells.guardian) {
        const m = st.persons.filter(p => od_personFits_(cells.guardian, p));
        if (m.length === 1) return od_ok_(m[0].key, "OSIS + name");
      } else if (st.persons.length === 1) {
        return od_ok_(st.persons[0].key, "OSIS");
      }
      if (st.persons.length) return od_pick_(st.persons.map(p => p.key), "Which guardian of " + st.name + "?");
    }
  }
  const exact = cells.guardian ? (idx.pByNameKey.get(od_tokenKey_(cells.guardian.tokens)) || []) : [];
  const contact = od_studentsByContact_(cells, idx).persons;
  if (exact.length === 1) return od_ok_(exact[0].key, "Name");
  if (exact.length > 1) {
    const n = exact.filter(p => contact.indexOf(p) !== -1);
    if (n.length === 1) return od_ok_(n[0].key, "Name + phone");
    return od_pick_(exact.map(p => p.key), exact.length + " guardians named " + exact[0].name);
  }
  if (contact.length === 1) return od_ok_(contact[0].key, cells.phone ? "Phone" : "Email");
  const guesses = [];
  if (cells.guardian) od_personCandidates_(cells.guardian, idx).forEach(p => { if (od_personFits_(cells.guardian, p)) guesses.push(p); });
  if (!guesses.length && cells.name) {
    // A guardian list that also names the child: that child's guardians
    od_studentGuesses_({ name: cells.name }, idx).forEach(g => g.st.persons.forEach(p => { if (guesses.indexOf(p) === -1) guesses.push(p); }));
  }
  if (guesses.length === 1) return od_guess_(guesses[0].key, "“" + (cells.guardian ? cells.guardian.raw : cells.name.raw) + "” looks like " + guesses[0].name);
  if (guesses.length > 1) return od_pick_(guesses.map(p => p.key), guesses.length + " possible guardians");
  if (contact.length > 1) return od_pick_(contact.map(p => p.key), "That phone / email belongs to " + contact.length + " guardians");
  return od_none_("No guardian matches this row");
}

/** Every data row of a list -> its result, your decisions (setup.picks) first. */
function od_matchAll_(setup, rows, idx) {
  const by = setup.by === "guardian" ? "guardian" : "student";
  const exists = t => (by === "student" ? idx.byOsis.has(t) : idx.pByKey.has(t));
  return rows.map(row => {
    const cells = od_rowCells_(row, setup.roles || {});
    const key = od_rowKey_(cells);
    const decided = setup.picks[key];
    let res;
    if (decided === "skip") res = { status: "skip", how: "Skipped" };
    else if (decided && exists(decided)) res = od_ok_(decided, "Your pick");
    else res = by === "student" ? od_matchStudentRow_(cells, idx) : od_matchGuardianRow_(cells, idx);
    res.key = key;
    res.cells = cells;
    return res;
  });
}

function od_targetText_(by, target, idx) {
  if (by === "student") {
    const st = idx.byOsis.get(target);
    return st ? { text: st.osis + " · " + st.name, sub: st.guardianText } : { text: target, sub: "" };
  }
  const p = idx.pByKey.get(target);
  return p ? { text: p.name, sub: p.kids.map(k => k.name).join(", ") } : { text: target, sub: "" };
}

function od_stats_(results) {
  const s = { rows: results.length, matched: 0, review: 0, skipped: 0, none: 0 };
  results.forEach(r => {
    if (r.status === "ok") s.matched++;
    else if (r.status === "skip") s.skipped++;
    else if (r.status === "none") { s.none++; s.review++; }
    else s.review++;
  });
  return s;
}

// ======================================================================
// OUTPUTS (every build)
// ======================================================================

/** A variable's value for one row: the cell, or the blank text, or null (nothing to show). */
function od_varValue_(v, row) {
  const cell = toText_(row[v.col - 1]);
  if (cell !== "") return cell;
  return v.blank && v.blank.mode === "text" && toText_(v.blank.text) ? toText_(v.blank.text) : null;
}

/**
 * Called by every build (guarded line in buildAllDerivedSheets), after the
 * Student List is built in memory. Matches every list, refreshes each list tab's
 * "✔ Matched" columns and bands, writes the Backend's Additional Information and
 * returns { studentCols, guardianCols } for the Student List / Parents Divided.
 */
function od_buildOutputs_(ss, masterTable, guardianRows) {
  const lists = od_loadLists_(ss);
  const ready = lists.filter(l => l.sheet && l.setup.ready);
  const out = { studentCols: [], guardianCols: [] };
  if (!lists.length) {
    backendClearBlock_(ss, "ADDL_INFO", BACKEND_HEADER_ROW);
    return out;
  }
  const idx = od_index_(masterTable.slice(1), guardianRows);
  const lines = new Map(); // OSIS -> { student: [], byGuardian: Map(lower name -> { name, lines }) }
  const entry = osis => {
    if (!lines.has(osis)) lines.set(osis, { student: [], byGuardian: new Map() });
    return lines.get(osis);
  };
  const addLine = (osis, guardianName, text) => {
    const e = entry(osis);
    if (!guardianName) { if (e.student.indexOf(text) === -1) e.student.push(text); return; }
    const k = guardianName.toLowerCase();
    if (!e.byGuardian.has(k)) e.byGuardian.set(k, { name: guardianName, lines: [] });
    const g = e.byGuardian.get(k);
    if (g.lines.indexOf(text) === -1) g.lines.push(text);
  };

  ready.forEach(list => {
    try {
      const setup = list.setup;
      const table = od_readTable_(list.sheet, setup);
      const results = od_matchAll_(setup, table.rows, idx);
      od_writeToolColumns_(list.sheet, list, table, results, idx);
      setup.stats = od_stats_(results);
      const by = setup.by === "guardian" ? "guardian" : "student";
      setup.vars.filter(v => v.col && (v.show || v.column)).forEach(v => {
        const label = toText_(v.label) || toText_(v.header) || ("Column " + od_colLetter_(v.col));
        const values = new Map(); // target -> [values]
        results.forEach((res, i) => {
          if (res.status !== "ok") return;
          const val = od_varValue_(v, table.rows[i]);
          if (val === null) return;
          if (!values.has(res.target)) values.set(res.target, []);
          if (values.get(res.target).indexOf(val) === -1) values.get(res.target).push(val);
        });
        const notText = v.notOnList && v.notOnList.mode === "text" ? toText_(v.notOnList.text) : "";
        if (by === "student") {
          const byOsis = new Map();
          idx.students.forEach(st => {
            const vals = values.get(st.key);
            const value = vals ? vals.join(" / ") : notText;
            if (!value) return;
            byOsis.set(st.osis, value);
            if (v.show) addLine(st.osis, "", label + ": " + value);
          });
          if (v.column) out.studentCols.push({ header: label, byOsis: byOsis });
        } else {
          const byGuardian = new Map();
          idx.persons.forEach(p => {
            const vals = values.get(p.key);
            const value = vals ? vals.join(" / ") : notText;
            if (!value) return;
            p.kids.forEach(st => {
              byGuardian.set(st.osis + "|" + p.name.toLowerCase(), value);
              if (v.show) addLine(st.osis, p.name, label + ": " + value);
            });
          });
          if (v.column) out.guardianCols.push({ header: label, byGuardian: byGuardian });
        }
      });
    } catch (e) {
      console.warn("Your Lists: \"" + list.name + "\" skipped: " + e.message);
    }
  });

  // Backend Additional Information: student lines (Guardian blank), then each guardian's
  const rows = [];
  idx.students.forEach(st => {
    const e = lines.get(st.osis);
    if (!e) return;
    if (e.student.length) rows.push([st.osis, st.name, "", e.student.join("\n")]);
    e.byGuardian.forEach(g => rows.push([st.osis, st.name, g.name, g.lines.join("\n")]));
  });
  rows.sort(byStudentLastName_(1));
  backendWriteBlock_(ss, "ADDL_INFO", [["OSIS", "Student", "Guardian", "Additional Information"]].concat(rows));
  od_saveLists_(ss, lists); // keeps each list's counts for the sidebar
  return out;
}

/** Rows 1-3 + the "✔ Matched" / "How" columns of a list tab. */
function od_writeToolColumns_(sheet, list, table, results, idx) {
  const setup = list.setup;
  const width = table.headers.length;
  if (!width) return;
  const toolCol = width + 2;
  if (sheet.getMaxColumns() < toolCol + 1) sheet.insertColumnsAfter(sheet.getMaxColumns(), toolCol + 1 - sheet.getMaxColumns());
  const by = setup.by === "guardian" ? "guardian" : "student";
  const cells = [OD_TOOL_HEADERS.slice()], colors = [[OD_COLORS.header, OD_COLORS.header]];
  results.forEach(r => {
    let text, color;
    if (r.status === "ok") { text = od_targetText_(by, r.target, idx).text; color = OD_COLORS.ok; }
    else if (r.status === "guess") { text = "? " + od_targetText_(by, r.target, idx).text + " (best guess)"; color = OD_COLORS.guess; }
    else if (r.status === "pick") { text = "? " + r.reason; color = OD_COLORS.guess; }
    else if (r.status === "skip") { text = "— skipped"; color = OD_COLORS.skip; }
    else { text = "✗ " + r.reason; color = OD_COLORS.none; }
    cells.push([text, r.how || ""]);
    colors.push([color, color]);
  });
  const last = sheet.getMaxRows();
  if (last >= OD_HEADER_ROW) sheet.getRange(OD_HEADER_ROW, toolCol - 1, last - OD_HEADER_ROW + 1, 3).clearContent().setBackground(null);
  sheet.getRange(OD_HEADER_ROW, toolCol, cells.length, 2).setValues(cells).setBackgrounds(colors);
  sheet.getRange(OD_HEADER_ROW, toolCol, 1, 2).setFontWeight("bold");

  const stats = od_stats_(results);
  const status = stats.review
    ? stats.matched + " of " + stats.rows + " rows matched · " + stats.review + " need you (🚀 App Menu → 📥 Input Data Mode → Your Lists)"
    : "All " + stats.matched + " row(s) matched" + (stats.skipped ? " (" + stats.skipped + " skipped)" : "") + " ✓";
  const roleName = { osis: "OSIS", "class": "CLASS", phone: "PHONE", email: "EMAIL" };
  const chips = table.headers.map((h, i) => {
    const c = i + 1, roles = setup.roles || {};
    for (const k of Object.keys(roleName)) if (roles[k] === c) return roleName[k];
    const nameRole = (r, label) => r && (r.col === c || r.first === c || r.last === c) ? label + (r.first === c ? " (FIRST)" : r.last === c ? " (LAST)" : "") : "";
    const nr = nameRole(roles.student, "STUDENT") || nameRole(roles.guardian, "GUARDIAN");
    if (nr) return nr;
    const v = setup.vars.filter(x => x.col === c)[0];
    if (v && (v.show || v.column)) return "SHOWN: " + (toText_(v.label) || h);
    return "TAB ONLY";
  });
  while (chips.length < toolCol - 1) chips.push("");
  chips.push("FILLED BY MATCHING");
  od_writeBands_(sheet, list.name, list.updated || od_today_(), status, chips);
}

// ======================================================================
// COLUMN GUESSES (what each column is)
// ======================================================================

function od_columnStats_(headers, rows, idx) {
  return headers.map((h, i) => {
    const vals = rows.map(r => toText_(r[i])).filter(Boolean);
    const n = vals.length || 1;
    let osisLike = 0, osisKnown = 0, phoneLike = 0, emailLike = 0, studentKnown = 0, guardianKnown = 0, classKnown = 0, comma = 0;
    vals.forEach(v => {
      const digits = v.replace(/\D/g, "");
      if (/^[\d\s-]+$/.test(v) && digits.length >= 6 && digits.length <= 9) { osisLike++; if (idx.byOsis.has(normalizeOsis_(v))) osisKnown++; }
      if (digits.length >= 10 && digits.length <= 11 && normalizePhoneForMatch_(v).length === 10) phoneLike++;
      if (v.indexOf("@") !== -1) emailLike++;
      if (v.indexOf(",") !== -1) comma++;
      const parts = od_nameParts_(v);
      if (parts && /[a-z]/i.test(v) && v.indexOf("@") === -1) {
        const key = od_tokenKey_(parts.tokens);
        if (idx.byNameKey.has(key) || (parts.last && (idx.byLast.get(parts.last) || []).some(st => od_firstLike_(parts.first, st.first)))) studentKnown++;
        if (idx.pByNameKey.has(key) || (parts.last && (idx.pByLast.get(parts.last) || []).some(p => od_firstLike_(parts.first, p.first)))) guardianKnown++;
      }
      if (idx.classes.has(v.toUpperCase())) classKnown++;
    });
    return { col: i + 1, header: h, n: vals.length, osisLike: osisLike / n, osisKnown: osisKnown, osisKnownShare: osisKnown / n,
             phone: phoneLike / n, email: emailLike / n, student: studentKnown / n, guardian: guardianKnown / n, cls: classKnown / n, comma: comma / n,
             samples: vals.slice(0, 3) };
  });
}

/** Best guess of what each column is + student or guardian list. */
function od_guessRoles_(headers, rows, idx) {
  const stats = od_columnStats_(headers, rows, idx);
  const used = new Set();
  const roles = { osis: 0, student: null, guardian: null, phone: 0, email: 0, "class": 0 };
  const H = s => String(s.header).toLowerCase();
  const pick = (test, score) => {
    let best = null, bestScore = 0;
    stats.forEach(s => { if (used.has(s.col) || !test(s)) return; const sc = score(s); if (sc > bestScore) { best = s; bestScore = sc; } });
    if (best) used.add(best.col);
    return best;
  };
  const guardianWord = /parent|guardian|mother|father|\bmom\b|\bdad\b|family|adult|contact/;
  const studentWord = /student|child|kid|scholar/;

  const osis = pick(s => s.osisKnownShare >= 0.3 || (/osis|student\s*(id|#|no|num)|\bid\b|id\s*#/.test(H(s)) && s.osisLike >= 0.5),
                    s => s.osisKnownShare * 2 + s.osisLike + (/osis/.test(H(s)) ? 1 : 0));
  if (osis) roles.osis = osis.col;
  const email = pick(s => s.email >= 0.6 || (/e-?mail/.test(H(s)) && s.email >= 0.3), s => s.email + (/e-?mail/.test(H(s)) ? 1 : 0));
  if (email) roles.email = email.col;
  const phone = pick(s => s.phone >= 0.6 || (/phone|cell|mobile|\btel/.test(H(s)) && s.phone >= 0.3), s => s.phone + (/phone|cell|mobile/.test(H(s)) ? 1 : 0));
  if (phone) roles.phone = phone.col;

  // Two-column names ("First Name" / "Last Name", "Parent First" / "Parent Last")
  const firstCols = stats.filter(s => !used.has(s.col) && /first/.test(H(s)));
  const lastCols = stats.filter(s => !used.has(s.col) && /last|surname/.test(H(s)));
  const pairFor = isGuardian => {
    const f = firstCols.filter(s => guardianWord.test(H(s)) === isGuardian)[0];
    const l = lastCols.filter(s => guardianWord.test(H(s)) === isGuardian)[0];
    if (f && l) { used.add(f.col); used.add(l.col); return { first: f.col, last: l.col }; }
    return null;
  };
  roles.guardian = pairFor(true);
  roles.student = pairFor(false);

  if (!roles.student) {
    const st = pick(s => !guardianWord.test(H(s)) && (studentWord.test(H(s)) || /name/.test(H(s)) || s.student >= 0.3) && s.osisLike < 0.5 && s.phone < 0.5,
                    s => s.student * 2 + (studentWord.test(H(s)) ? 1 : 0) + (/name/.test(H(s)) ? 0.5 : 0) - (s.guardian > s.student ? 1 : 0));
    if (st && (st.student >= 0.2 || studentWord.test(H(st)) || /^(full\s*)?name$/.test(H(st).trim()))) roles.student = { col: st.col };
    else if (st) used.delete(st.col);
  }
  if (!roles.guardian) {
    const g = pick(s => (guardianWord.test(H(s)) || s.guardian >= 0.3) && s.osisLike < 0.5 && s.phone < 0.5 && s.email < 0.5,
                   s => s.guardian * 2 + (guardianWord.test(H(s)) ? 1 : 0));
    if (g) roles.guardian = { col: g.col };
  }
  const cls = pick(s => s.cls >= 0.3 && /class|homeroom|official|section|room/.test(H(s)), s => s.cls);
  if (cls) roles["class"] = cls.col;

  const by = (roles.osis || roles.student) ? "student" : (roles.guardian || roles.phone || roles.email) ? "guardian" : "student";
  return { roles: roles, by: by, stats: stats };
}

/** Columns that aren't used to find who the row is (the variables you can keep). */
function od_roleCols_(roles) {
  const cols = [];
  ["osis", "phone", "email", "class"].forEach(k => { if (roles[k]) cols.push(roles[k]); });
  ["student", "guardian"].forEach(k => {
    const r = roles[k];
    if (!r) return;
    ["col", "first", "last"].forEach(x => { if (r[x]) cols.push(r[x]); });
  });
  return cols;
}

/**
 * Variables for the headers: kept from the saved setup by header name (a new
 * version of the same list keeps its settings), new columns start "tab only"
 * (a list with just one other column starts shown).
 */
function od_varsFor_(headers, roles, savedVars) {
  const roleCols = od_roleCols_(roles);
  const byHeader = {};
  (savedVars || []).forEach(v => { byHeader[String(v.header).toLowerCase()] = v; });
  const vars = [];
  headers.forEach((h, i) => {
    const col = i + 1;
    if (roleCols.indexOf(col) !== -1) return;
    const old = byHeader[String(h).toLowerCase()];
    vars.push(old ? Object.assign({}, old, { col: col, header: h }) :
      { col: col, header: h, label: h, show: false, column: false, blank: { mode: "nothing", text: "" }, notOnList: { mode: "nothing", text: "" } });
  });
  if (vars.length === 1 && !(savedVars || []).length) { vars[0].show = true; vars[0].column = true; }
  return vars;
}

/** Saved roles moved to where those headers are now (null when one is missing). */
function od_rolesByHeader_(savedSetup, headers) {
  const old = savedSetup.headers || [];
  if (!old.length) return null;
  const pos = {};
  headers.forEach((h, i) => { pos[String(h).toLowerCase()] = i + 1; });
  const move = c => (c ? pos[String(old[c - 1] || "").toLowerCase()] || -1 : 0);
  const r = savedSetup.roles || {}, out = {};
  let missing = false;
  ["osis", "phone", "email", "class"].forEach(k => { const m = move(r[k]); if (m === -1) missing = true; out[k] = m > 0 ? m : 0; });
  ["student", "guardian"].forEach(k => {
    if (!r[k]) { out[k] = null; return; }
    const o = {};
    ["col", "first", "last"].forEach(x => { if (r[k][x]) { const m = move(r[k][x]); if (m === -1) missing = true; else o[x] = m; } });
    out[k] = Object.keys(o).length ? o : null;
  });
  return missing ? null : out;
}

// ======================================================================
// SIDEBAR ENDPOINTS (Z#OwnDataSidebar.html; the list buttons in Input Data Mode)
// ======================================================================

function od_withLock_(fn) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) return { error: "The sheet is busy. Try again in a moment." };
  try {
    return fn();
  } catch (e) {
    console.error("Your Lists: " + (e && e.stack ? e.stack : e));
    return { error: "❌ Script error: " + (e && e.message ? e.message : e) };
  } finally {
    lock.releaseLock();
  }
}

/** For Input Data Mode's step 5: one line per list (from the Backend only, no tab reads). */
function od_listSummaries_(ss) {
  return od_loadLists_(ss).filter(l => l.sheet).map(l => ({
    tabId: l.tabId, name: l.name, tabName: l.sheet.getName(), by: l.setup.by, ready: !!l.setup.ready,
    updated: l.updated, stats: l.setup.stats || null
  }));
}

/** "Add another list": makes the "📥 Name" tab and opens the setup sidebar on it. */
function od_createList(name) {
  const res = od_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    name = od_cleanName_(name);
    if (!name) return { error: "Give the list a name first." };
    const sheet = ss.insertSheet(od_uniqueTabName_(ss, OD_TAB_PREFIX + name));
    if (sheet.getMaxColumns() < 26) sheet.insertColumnsAfter(sheet.getMaxColumns(), 26 - sheet.getMaxColumns());
    od_resetListTab_(sheet, name);
    const lists = od_loadLists_(ss);
    lists.push({ tabId: sheet.getSheetId(), name: name, setup: od_emptySetup_(), updated: "", sheet: sheet });
    od_saveLists_(ss, lists);
    sheet.activate();
    return { tabId: sheet.getSheetId() };
  });
  if (res && res.tabId !== undefined) od_showWizard_(res.tabId);
  return res;
}

/** "Edit setup": the setup sidebar on an existing list. */
function od_openList(tabId) {
  const res = od_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const list = od_findList_(od_loadLists_(ss), tabId);
    if (!list || !list.sheet) return { error: "That list's tab is gone." };
    if (list.sheet.isSheetHidden()) list.sheet.showSheet();
    list.sheet.activate();
    return { tabId: list.tabId };
  });
  if (res && res.tabId !== undefined) od_showWizard_(res.tabId);
  return res;
}

/** "Paste new version": empties the list's tab (its setup and your decisions are kept) and waits for the paste. */
function od_repasteList(tabId) {
  const res = od_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const list = od_findList_(od_loadLists_(ss), tabId);
    if (!list || !list.sheet) return { error: "That list's tab is gone." };
    if (list.sheet.isSheetHidden()) list.sheet.showSheet();
    od_resetListTab_(list.sheet, list.name);
    list.sheet.activate();
    return { tabId: list.tabId };
  });
  if (res && res.tabId !== undefined) od_showWizard_(res.tabId);
  return res;
}

function od_renameList(tabId, name) {
  return od_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const lists = od_loadLists_(ss);
    const list = od_findList_(lists, tabId);
    name = od_cleanName_(name);
    if (!list || !list.sheet) return { error: "That list's tab is gone." };
    if (!name) return { error: "Give the list a name." };
    list.name = name;
    list.sheet.setName(od_uniqueTabName_(ss, OD_TAB_PREFIX + name, list.sheet));
    const title = String(list.sheet.getRange(1, 1).getValue());
    if (title.indexOf(OD_BAND_TITLE) === 0) list.sheet.getRange(1, 1).setValue(title.replace(/^📥 OWN DATA · .*? — /, OD_BAND_TITLE + name.toUpperCase() + " — "));
    od_saveLists_(ss, lists);
    return { ok: true };
  });
}

/** Deletes the list's tab and its setup (never the tab the user is standing on: they are moved off it first). */
function od_deleteList(tabId) {
  return od_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const lists = od_loadLists_(ss);
    const list = od_findList_(lists, tabId);
    if (!list) return { ok: true };
    if (list.sheet) {
      const active = ss.getActiveSheet();
      if (active && active.getSheetId() === list.sheet.getSheetId()) {
        const other = ss.getSheets().filter(s => !s.isSheetHidden() && s.getSheetId() !== list.sheet.getSheetId())[0];
        if (other) other.activate();
      }
      ss.deleteSheet(list.sheet);
    }
    od_saveLists_(ss, lists.filter(l => l !== list));
    return { ok: true };
  });
}

function od_showWizard_(tabId) {
  PropertiesService.getDocumentProperties().setProperty(OD_WIZARD_KEY, String(tabId));
  const html = HtmlService.createHtmlOutputFromFile("Z#OwnDataSidebar").setTitle("📥 Your Lists");
  SpreadsheetApp.getUi().showSidebar(html);
}

function od_wizardList_(ss) {
  const tabId = PropertiesService.getDocumentProperties().getProperty(OD_WIZARD_KEY);
  const lists = od_loadLists_(ss);
  return { lists: lists, list: od_findList_(lists, tabId) };
}

/** The setup sidebar's view: "paste" (waiting for the list) or "setup" (columns / variables). */
function od_getWizard() {
  return od_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const w = od_wizardList_(ss);
    if (!w.list || !w.list.sheet) return { error: "This list's tab is gone. Go back to Input Data Mode." , gone: true };
    const table = od_readTable_(w.list.sheet, null);
    if (!table.headers.length || !od_hasPaste_(w.list.sheet)) return od_pasteView_(w.list);
    return od_setupView_(ss, w.list, table, null);
  });
}

function od_pasteView_(list) {
  return { phase: "paste", list: { tabId: list.tabId, name: list.name, tabName: list.sheet.getName(), ready: !!list.setup.ready } };
}

/** Polled every 2 s while waiting: has something been pasted on the list's tab? */
function od_wizardCheck() {
  return od_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const w = od_wizardList_(ss);
    if (!w.list || !w.list.sheet) return { error: "This list's tab is gone.", gone: true };
    return { pasted: od_hasPaste_(w.list.sheet) };
  });
}

/** Moves the paste into place, then the column setup (a new version of a known list keeps its setup). */
function od_readPaste() {
  return od_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const w = od_wizardList_(ss);
    if (!w.list || !w.list.sheet) return { error: "This list's tab is gone.", gone: true };
    const taken = od_takePaste_(w.list.sheet, w.list.name);
    if (!taken) return od_pasteView_(w.list);
    if (!taken.rows.length) return Object.assign(od_pasteView_(w.list), { error: "I found the header row, but no rows under it. Paste the whole list, headers and rows." });
    w.list.updated = od_today_();
    w.list.setup.title = taken.title;
    od_saveLists_(ss, w.lists);
    return od_setupView_(ss, w.list, { headers: taken.headers, rows: taken.rows }, taken);
  });
}

/** Columns + variables for the sidebar (saved setup when it still fits, else the guesses). */
function od_setupView_(ss, list, table, taken) {
  const idx = od_indexFromSheet_(ss);
  const guess = od_guessRoles_(table.headers, table.rows, idx);
  const saved = list.setup;
  const savedRoles = saved.ready ? (taken ? od_rolesByHeader_(saved, table.headers) : saved.roles) : null;
  const roles = savedRoles || guess.roles;
  const by = savedRoles ? saved.by : guess.by;
  return {
    phase: "setup",
    list: { tabId: list.tabId, name: list.name, tabName: list.sheet.getName(), ready: !!saved.ready, title: saved.title || "" },
    kept: !!savedRoles,                 // a new version of the list: the old setup was kept
    headerRow: OD_HEADER_ROW,
    rowCount: table.rows.length,
    columns: guess.stats.map(s => ({ col: s.col, letter: od_colLetter_(s.col), header: s.header, samples: s.samples,
      hint: { osisKnown: s.osisKnown, n: s.n, comma: s.comma >= 0.5 } })),
    roles: roles,
    by: by,
    vars: od_varsFor_(table.headers, roles, saved.vars)
  };
}

/** Normalizes what the sidebar sends (roles as column numbers, variables). */
function od_cleanSetupInput_(p, headers) {
  const n = headers.length;
  const col = c => { c = Number(c); return c >= 1 && c <= n ? c : 0; };
  const nameRole = r => {
    if (!r) return null;
    if (r.first || r.last) { const f = col(r.first), l = col(r.last); return f || l ? { first: f, last: l } : null; }
    const c = col(r.col);
    return c ? { col: c } : null;
  };
  const roles = { osis: col(p.roles && p.roles.osis), student: nameRole(p.roles && p.roles.student), guardian: nameRole(p.roles && p.roles.guardian),
                  phone: col(p.roles && p.roles.phone), email: col(p.roles && p.roles.email), "class": col(p.roles && p.roles["class"]) };
  const roleCols = od_roleCols_(roles);
  const vars = (p.vars || []).filter(v => col(v.col) && roleCols.indexOf(col(v.col)) === -1).map(v => ({
    col: col(v.col), header: headers[col(v.col) - 1], label: od_cleanLabel_(v.label) || headers[col(v.col) - 1],
    show: !!v.show, column: !!v.column,
    blank: { mode: v.blank && v.blank.mode === "text" ? "text" : "nothing", text: toText_(v.blank && v.blank.text).slice(0, 100) },
    notOnList: { mode: v.notOnList && v.notOnList.mode === "text" ? "text" : "nothing", text: toText_(v.notOnList && v.notOnList.text).slice(0, 100) }
  }));
  return { by: p.by === "guardian" ? "guardian" : "student", roles: roles, vars: vars };
}

/**
 * Step 3 (Match): the rows that need a person (pick one / best guess / not found)
 * with their choices, the counts, and the list to pick from (Student List
 * students, or guardians for a guardian list).
 */
function od_matchPreview(p) {
  return od_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const w = od_wizardList_(ss);
    if (!w.list || !w.list.sheet) return { error: "This list's tab is gone.", gone: true };
    const table = od_readTable_(w.list.sheet, null);
    const input = od_cleanSetupInput_(p || {}, table.headers);
    const setup = Object.assign({}, w.list.setup, input);
    const idx = od_indexFromSheet_(ss);
    if (!idx.students.length) return { error: "The Student List is empty. Build the sheets once (🏗️ Build Sheets Only), then come back." };
    const results = od_matchAll_(setup, table.rows, idx);
    const by = setup.by;
    const label = (row, cells) => {
      if (cells.name) return cells.name.raw;
      if (cells.guardian) return cells.guardian.raw;
      return row.filter(Boolean).slice(0, 2).join(" · ");
    };
    const howCounts = {};
    results.forEach(r => { if (r.status === "ok") howCounts[r.how] = (howCounts[r.how] || 0) + 1; });
    const attention = [];
    results.forEach((r, i) => {
      if (r.status === "ok" || (r.status === "skip" && !(p && p.showSkipped))) return;
      const row = table.rows[i];
      const extra = [];
      if (r.cells.osis) extra.push("OSIS " + r.cells.osis);
      if (by === "student" && r.cells.guardian) extra.push("parent “" + r.cells.guardian.raw + "”");
      setup.vars.filter(v => v.show || v.column).slice(0, 1).forEach(v => { if (toText_(row[v.col - 1])) extra.push(v.label + " " + toText_(row[v.col - 1])); });
      attention.push({
        i: i, row: OD_HEADER_ROW + 1 + i, label: label(row, r.cells), sub: extra.join(" · "), status: r.status, reason: r.reason || "",
        target: r.target || null, targetText: r.target ? od_targetText_(by, r.target, idx) : null,
        options: (r.options || []).map(t => Object.assign({ target: t }, od_targetText_(by, t, idx)))
      });
    });
    const picker = by === "student"
      ? idx.students.slice().sort(byStudentLastName_("name")).map(st => ({ target: st.key, text: st.osis + " · " + st.name, sub: st.guardianText }))
      : idx.persons.slice().sort((a, b) => compareNatural_(a.last + " " + a.first, b.last + " " + b.first))
          .map(pp => ({ target: pp.key, text: pp.name, sub: pp.kids.map(k => k.name).join(", ") }));
    return { phase: "match", by: by, stats: od_stats_(results), how: howCounts, attention: attention, picker: picker };
  });
}

/**
 * Save: the setup, your decisions ({ rowIndex: target | "skip" }), the "✔ Matched"
 * columns on the tab. The Student List / Parents Divided columns and the phone
 * contacts follow at the next build.
 */
function od_saveList(p) {
  return od_withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const w = od_wizardList_(ss);
    if (!w.list || !w.list.sheet) return { error: "This list's tab is gone.", gone: true };
    const list = w.list;
    const table = od_readTable_(list.sheet, null);
    const input = od_cleanSetupInput_(p || {}, table.headers);
    const name = od_cleanName_(p && p.name);
    const setup = Object.assign(list.setup, input, { ready: true, headers: table.headers.slice() });
    const idx = od_indexFromSheet_(ss);
    const before = od_matchAll_(setup, table.rows, idx);
    const decisions = (p && p.decisions) || {};
    Object.keys(decisions).forEach(i => {
      const r = before[Number(i)];
      const t = decisions[i];
      if (!r || !t) return;
      if (t === "auto") { delete setup.picks[r.key]; return; }
      if (t === "skip" || (setup.by === "student" ? idx.byOsis.has(t) : idx.pByKey.has(t))) setup.picks[r.key] = t;
    });
    const results = od_matchAll_(setup, table.rows, idx);
    setup.stats = od_stats_(results);
    if (name && name !== list.name) {
      list.name = name;
      list.sheet.setName(od_uniqueTabName_(ss, OD_TAB_PREFIX + name, list.sheet));
    }
    if (!list.updated) list.updated = od_today_();
    od_writeToolColumns_(list.sheet, list, table, results, idx);
    od_saveLists_(ss, w.lists);
    return { phase: "done", list: { tabId: list.tabId, name: list.name, tabName: list.sheet.getName() }, stats: setup.stats,
             inputMode: typeof idm_loadState_ === "function" && !!idm_loadState_() };
  });
}

/** Back to Input Data Mode's sidebar (when it is on), else just close. */
function od_backToInputMode() {
  PropertiesService.getDocumentProperties().deleteProperty(OD_WIZARD_KEY);
  if (typeof openInputDataMode === "function" && typeof idm_loadState_ === "function" && idm_loadState_()) {
    openInputDataMode();
    return { reopened: true };
  }
  return { reopened: false };
}

/** "Build now" from the setup sidebar. */
function od_build() {
  PropertiesService.getDocumentProperties().deleteProperty(OD_WIZARD_KEY);
  runBuildOnly();
  return true;
}
