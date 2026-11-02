// ======================================================================
// FILE: NewStudents.gs
// PURPOSE: Keeps a running "New Students" log. Every build (including the
// nightly sync) compares the fresh Master Table with the one it is about to
// replace and records students who just appeared in Raw Data (ATS) or came
// back from Discharged. Only Raw Data counts; Source is always "ATS".
// The tab is created once and only ever added to (a row is re-dated when its
// student re-enters).
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

const ATS_SNAPSHOT_PROP = "NEW_STUDENTS_ATS_OSIS";
const ATS_SNAPSHOT_CHUNK = 8000; // Document properties cap each value at ~9KB

/**
 * OSIS that were in Raw Data at the previous build, or null when there is no
 * snapshot for this spreadsheet yet (first run, or a copied sheet's leftovers).
 */
function loadAtsSnapshot_(ss) {
  const props = PropertiesService.getDocumentProperties();
  if (props.getProperty(ATS_SNAPSHOT_PROP + "_SHEET") !== ss.getId()) return null;
  const count = Number(props.getProperty(ATS_SNAPSHOT_PROP + "_N")) || 0;
  let text = "";
  for (let i = 0; i < count; i++) text += props.getProperty(ATS_SNAPSHOT_PROP + "_" + i) || "";
  return new Set(text.split(",").filter(Boolean));
}

function saveAtsSnapshot_(ss, osisSet) {
  const props = PropertiesService.getDocumentProperties();
  const text = Array.from(osisSet).join(",");
  const oldCount = Number(props.getProperty(ATS_SNAPSHOT_PROP + "_N")) || 0;
  const values = {};
  let n = 0;
  for (let i = 0; i < text.length; i += ATS_SNAPSHOT_CHUNK) values[ATS_SNAPSHOT_PROP + "_" + n++] = text.slice(i, i + ATS_SNAPSHOT_CHUNK);
  values[ATS_SNAPSHOT_PROP + "_N"] = String(n);
  values[ATS_SNAPSHOT_PROP + "_SHEET"] = ss.getId();
  props.setProperties(values);
  for (let i = n; i < oldCount; i++) props.deleteProperty(ATS_SNAPSHOT_PROP + "_" + i);
}

/**
 * Logs students who are new to Raw Data (ATS). Only Raw Data counts: New/Edit
 * Student and other Directory sources are ignored, and Source is always "ATS".
 * A student is new when their OSIS was not in Raw Data at the previous build
 * (tracked in a snapshot), or when they come back from Discharged.
 *
 * @param {Array[]} masterTable        Fresh Master Table (header row first)
 * @param {Map}     oldMasterStatusMap OSIS -> Status from the Master Table before this build
 * @param {Array[]} rawData            Raw Data values (A:AE)
 * @param {Array}   guardianRows       Guardian objects (for Language)
 */
function logNewStudents_(ss, masterTable, oldMasterStatusMap, rawData, guardianRows) {
  let sheet = ss.getSheetByName(NEW_STUDENTS_SHEET);
  if (!sheet) sheet = ss.insertSheet(NEW_STUDENTS_SHEET);
  if (!toText_(sheet.getRange(1, 1).getValue())) {
    const school = schoolShortName_(rawData);
    const headers = [school ? `Date Entered ${school}` : "Date Entered", "Source", "OSIS", "Student", "Guardian", "Site", "Class", "Language"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  const atsOsis = new Set();
  const homeLanguage = new Map();
  for (let i = 0; i < rawData.length; i++) {
    const id = toText_(rawData[i][2]);
    if (!isRealStudentId_(id)) continue;
    atsOsis.add(id);
    if (!homeLanguage.has(id) && toText_(rawData[i][18])) homeLanguage.set(id, properCase_(firstWord_(rawData[i][18])));
  }

  // Previous Raw Data roster. Without a snapshot, fall back to the old Master Table.
  const prevAts = loadAtsSnapshot_(ss);
  const hadBefore = osis => prevAts ? prevAts.has(osis) : oldMasterStatusMap.has(osis);

  // First build has nothing to compare against, so it would list every student
  if (!prevAts && oldMasterStatusMap.size === 0) { saveAtsSnapshot_(ss, atsOsis); return; }

  const languagesById = new Map();
  for (const g of guardianRows) {
    if (!g.language) continue;
    if (!languagesById.has(g.student_id)) languagesById.set(g.student_id, []);
    const list = languagesById.get(g.student_id);
    if (list.indexOf(g.language) === -1) list.push(g.language);
  }

  // Existing log rows by OSIS
  const lastRow = sheet.getLastRow();
  const logged = new Map();
  if (lastRow > 1) {
    sheet.getRange(2, 3, lastRow - 1, 1).getValues().forEach((r, i) => {
      const osis = String(r[0]).trim();
      if (osis) logged.set(osis, i + 2);
    });
  }

  const today = new Date();
  const toAppend = [];
  for (let i = 1; i < masterTable.length; i++) {
    const m = masterTable[i];
    const osis = String(m[0]).trim();
    if (!osis || osis.includes(",") || !atsOsis.has(osis)) continue;
    if (/discharge/i.test(String(m[6]))) continue;

    const reEntered = oldMasterStatusMap.has(osis) && /discharge/i.test(oldMasterStatusMap.get(osis));
    if (hadBefore(osis) && !reEntered) continue;

    const language = (languagesById.get(osis) || []).join(", ") || homeLanguage.get(osis) || "";
    const row = [today, "ATS", m[0], m[1], m[2], m[3], m[4], language];

    const priorRow = logged.get(osis);
    if (priorRow) sheet.getRange(priorRow, 1, 1, row.length).setValues([row]);
    else toAppend.push(row);
  }

  if (toAppend.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, toAppend.length, toAppend[0].length).setValues(toAppend);
  }
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).setNumberFormat("M/d/yyyy");

  saveAtsSnapshot_(ss, atsOsis);
}
