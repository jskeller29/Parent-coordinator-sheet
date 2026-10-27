// ======================================================================
// FILE: NewStudents.gs
// PURPOSE: Keeps a running "New Students" log. Every build (including the
// nightly sync) compares the fresh Master Table with the one it is about to
// replace and records students who just appeared or came back from Discharged.
// The tab is created once and only ever added to (a row is re-dated when its
// student re-enters or its Source changes, e.g. Directory -> ATS/Directory).
// ======================================================================

const NEW_STUDENTS_SHEET = "New Students";

/**
 * "75M094" -> "P94M" (D75 school naming). Falls back to "" when the DBN
 * doesn't look like a D75 code.
 */
function schoolShortName_(rawData) {
  for (let i = 0; i < rawData.length; i++) {
    const m = toText_(rawData[i][1]).toUpperCase().match(/^75([MXKQR])(\d{3})$/);
    if (m) return `P${Number(m[2])}${m[1]}`;
  }
  return "";
}

/**
 * Reads the Status of every student in the current Master Table (before it is
 * rewritten). Status is found by header, so column shifts don't break it.
 */
function readMasterStatusSnapshot_(ss) {
  const map = new Map();
  const sheet = ss.getSheetByName("Master Table");
  if (!sheet) return map;
  const data = sheet.getDataRange().getValues();
  const headerIdx = data.findIndex(r => String(r[0]).trim().toUpperCase() === "OSIS");
  if (headerIdx === -1) return map;
  let statusIdx = data[headerIdx].indexOf("Status");
  if (statusIdx === -1) statusIdx = 6;
  for (let i = headerIdx + 1; i < data.length; i++) {
    const osis = String(data[i][0]).trim();
    if (osis && !osis.includes(",")) map.set(osis, String(data[i][statusIdx]).trim());
  }
  return map;
}

/**
 * @param {Array[]} masterTable        Fresh Master Table (header row first)
 * @param {Map}     oldMasterStatusMap OSIS -> Status from the Master Table before this build
 * @param {Array[]} rawData            Raw Data values
 * @param {Set}     directoryOsis      OSIS that come from a non-ATS source (New/Edit Student, Old Master Table, ...)
 * @param {Array}   guardianRows       Guardian objects (for Language)
 * @param {Set}     [alsoNewOsis]      Optional: OSIS to log even though they were already in the Master Table
 */
function logNewStudents_(ss, masterTable, oldMasterStatusMap, rawData, directoryOsis, guardianRows, alsoNewOsis) {
  let sheet = ss.getSheetByName(NEW_STUDENTS_SHEET);
  if (!sheet) sheet = ss.insertSheet(NEW_STUDENTS_SHEET);
  if (!toText_(sheet.getRange(1, 1).getValue())) {
    const school = schoolShortName_(rawData);
    const headers = [school ? `Date Entered ${school}` : "Date Entered", "Source", "OSIS", "Student", "Guardian", "Site", "Class", "Language"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  // First build has nothing to compare against, so it would list every student
  if (oldMasterStatusMap.size === 0) return;

  const atsOsis = new Set();
  const homeLanguage = new Map();
  for (let i = 0; i < rawData.length; i++) {
    const id = toText_(rawData[i][2]);
    if (!isRealStudentId_(id)) continue;
    atsOsis.add(id);
    if (!homeLanguage.has(id) && toText_(rawData[i][18])) homeLanguage.set(id, properCase_(firstWord_(rawData[i][18])));
  }

  const languagesById = new Map();
  for (const g of guardianRows) {
    if (!g.language) continue;
    if (!languagesById.has(g.student_id)) languagesById.set(g.student_id, []);
    const list = languagesById.get(g.student_id);
    if (list.indexOf(g.language) === -1) list.push(g.language);
  }

  // Existing log rows by OSIS (row number + Source)
  const lastRow = sheet.getLastRow();
  const logged = new Map();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 3).getValues().forEach((r, i) => {
      const osis = String(r[2]).trim();
      if (osis) logged.set(osis, { row: i + 2, source: String(r[1]).trim() });
    });
  }

  const today = new Date();
  const toAppend = [];
  for (let i = 1; i < masterTable.length; i++) {
    const m = masterTable[i];
    const osis = String(m[0]).trim();
    if (!osis || osis.includes(",")) continue;
    if (/discharge/i.test(String(m[6]))) continue;

    const inAts = atsOsis.has(osis);
    const inDirectory = directoryOsis.has(osis);
    const source = inAts && inDirectory ? "ATS/Directory" : inAts ? "ATS" : "Directory";

    const wasListed = oldMasterStatusMap.has(osis);
    const reEntered = wasListed && /discharge/i.test(oldMasterStatusMap.get(osis));
    const prior = logged.get(osis);
    const forcedNew = !!(alsoNewOsis && alsoNewOsis.has(osis));
    if (wasListed && !reEntered && !forcedNew && !(prior && prior.source !== source)) continue;

    const language = (languagesById.get(osis) || []).join(", ") || homeLanguage.get(osis) || "";
    const row = [today, source, m[0], m[1], m[2], m[3], m[4], language];

    if (prior) sheet.getRange(prior.row, 1, 1, row.length).setValues([row]);
    else toAppend.push(row);
  }

  if (toAppend.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, toAppend.length, toAppend[0].length).setValues(toAppend);
  }
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).setNumberFormat("M/d/yyyy");
}
