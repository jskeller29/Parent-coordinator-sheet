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
 * Compares Raw Data (ATS, A:AE) with the previous build. Shared by the New
 * Students tab and the build report (sidebar / email) so they always agree.
 *   newOsis:        in Raw Data now but not at the last build, or back from Discharged
 *   dischargedOsis: "Current" in the old Master Table, in Raw Data at the last
 *                   build, and no longer in Raw Data
 * Call saveAtsSnapshot_(ss, changes.atsOsis) once the build has used them.
 *
 * @param {Array[]} masterTable        Fresh Master Table (header row first)
 * @param {Map}     oldMasterStatusMap OSIS -> Status from the Master Table before this build
 * @param {Array[]} rawData            Raw Data values
 */
function detectAtsChanges_(ss, masterTable, oldMasterStatusMap, rawData) {
  const atsOsis = new Set();
  for (let i = 0; i < rawData.length; i++) {
    const id = toText_(rawData[i][2]);
    if (isRealStudentId_(id)) atsOsis.add(id);
  }
  const newOsis = new Set();
  const dischargedOsis = new Set();

  // Previous Raw Data roster. Without a snapshot, fall back to the old Master Table.
  const prevAts = loadAtsSnapshot_(ss);
  // First build has nothing to compare against, so it would list every student
  if (!prevAts && oldMasterStatusMap.size === 0) return { atsOsis, newOsis, dischargedOsis };

  for (let i = 1; i < masterTable.length; i++) {
    const m = masterTable[i];
    const osis = String(m[0]).trim();
    if (!osis || osis.includes(",")) continue;
    const oldStatus = String(oldMasterStatusMap.has(osis) ? oldMasterStatusMap.get(osis) : "").trim();

    if (atsOsis.has(osis)) {
      if (/discharge/i.test(String(m[6]))) continue;
      const hadBefore = prevAts ? prevAts.has(osis) : oldMasterStatusMap.has(osis);
      if (!hadBefore || /discharge/i.test(oldStatus)) newOsis.add(osis);
    } else if (/^current$/i.test(oldStatus)) {
      // Without a snapshot, only count students the Master Table itself now marks Discharged
      const wasInAts = prevAts ? prevAts.has(osis) : /discharge/i.test(String(m[6]));
      if (wasInAts) dischargedOsis.add(osis);
    }
  }
  return { atsOsis, newOsis, dischargedOsis };
}

/**
 * Logs the students detectAtsChanges_ found new to Raw Data. Only Raw Data
 * counts: New/Edit Student and other Directory sources are ignored, and Source
 * is always "ATS".
 *
 * @param {Array[]} masterTable  Fresh Master Table (header row first)
 * @param {Object}  changes      Result of detectAtsChanges_
 * @param {Array[]} rawData      Raw Data values (A:AE)
 * @param {Array}   guardianRows Guardian objects (for Language)
 */
function logNewStudents_(ss, masterTable, changes, rawData, guardianRows) {
  let sheet = ss.getSheetByName(NEW_STUDENTS_SHEET);
  if (!sheet) sheet = ss.insertSheet(NEW_STUDENTS_SHEET);
  if (!toText_(sheet.getRange(1, 1).getValue())) {
    const school = schoolShortName_(rawData);
    const headers = [school ? `Date Entered ${school}` : "Date Entered", "Source", "OSIS", "Student", "Guardian", "Site", "Class", "Language"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  if (!changes.newOsis.size) return;

  const homeLanguage = new Map();
  for (let i = 0; i < rawData.length; i++) {
    const id = toText_(rawData[i][2]);
    if (changes.newOsis.has(id) && !homeLanguage.has(id) && toText_(rawData[i][18])) homeLanguage.set(id, properCase_(firstWord_(rawData[i][18])));
  }

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
    if (!changes.newOsis.has(osis)) continue;

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
}

/** Build report lines ("Name (OSIS)") for the sidebar / email. */
function atsChangeReportLines_(masterTable, changes) {
  const newStudents = [];
  const discharged = [];
  for (let i = 1; i < masterTable.length; i++) {
    const osis = String(masterTable[i][0]).trim();
    const line = `${masterTable[i][1]} (${osis})`;
    if (changes.newOsis.has(osis)) newStudents.push(line);
    if (changes.dischargedOsis.has(osis)) discharged.push(line);
  }
  return { newStudents, discharged };
}
