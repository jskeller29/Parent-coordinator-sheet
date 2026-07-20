/* =========================
   PARENT SQUARE INDEX
========================= */

function buildParentSquareIndex_(rawData) {
  const emailStatus = new Map();
  const phoneStatus = new Map();

  for (let i = 0; i < rawData.length; i++) {
    const r = rawData[i];

    const email = normalizeEmail_(r[34]);
    const phone = normalizePhoneForMatch_(r[35]);
    const secondary = normalizePhoneForMatch_(r[36]);
    const registered = normalizeYesNo_(r[38]);

    if (email) mergeStatusIntoMap_(emailStatus, email, registered);
    if (phone) mergeStatusIntoMap_(phoneStatus, phone, registered);
    if (secondary) mergeStatusIntoMap_(phoneStatus, secondary, registered);
  }

  return { emailStatus, phoneStatus };
}

function mergeStatusIntoMap_(map, key, newStatus) {
  if (!key) return;
  const current = map.get(key) || "";
  map.set(key, mergeYesNoStatus_(current, newStatus));
}

function mergeYesNoStatus_(a, b) {
  const aa = normalizeYesNo_(a);
  const bb = normalizeYesNo_(b);

  if (aa === "Yes" || bb === "Yes") return "Yes";
  if (aa === "No" || bb === "No") return "No";
  return "";
}

function normalizeYesNo_(value) {
  const v = toText_(value).toLowerCase();
  if (v === "yes") return "Yes";
  if (v === "no") return "No";
  return "";
}

function getParentsSquareStatus_(guardian, parentSquareIndex) {
  let status = "";

  const email = normalizeEmail_(guardian.email);
  const phone = normalizePhoneForMatch_(guardian.mobile);

  if (email && parentSquareIndex.emailStatus.has(email)) {
    status = mergeYesNoStatus_(status, parentSquareIndex.emailStatus.get(email));
  }

  if (phone && parentSquareIndex.phoneStatus.has(phone)) {
    status = mergeYesNoStatus_(status, parentSquareIndex.phoneStatus.get(phone));
  }

  return status;
}

/* =========================
   NYCSA INDEX (UPGRADED)
========================= */

function buildNYCSAIndex_(rawData) {
  const byStudentAndName = new Map();

  for (let i = 0; i < rawData.length; i++) {
    const r = rawData[i];

    const osis = String(r[46] || "").trim(); // NYCSA Student ID is index 46
    const first = properCase_(toText_(r[49]));
    const last = properCase_(toText_(r[50]));
    const fullName = `${first} ${last}`.trim();
    if (!fullName || !osis) continue;

    const anotherCustodial = normalizeYesNo_(r[51]);
    const accountCode = toText_(r[57]); // NYCSA Account Creation Code (Column BF)

    // ✨ NEW: Unique Key using OSIS + Name to prevent sibling bleed
    const compositeKey = `${osis}_${normalizeName_(fullName)}`;

    byStudentAndName.set(compositeKey, {
      anotherCustodial,
      accountCode
    });
  }

  return { byStudentAndName };
}

function getNYCSAStatus_(currentGuardian, otherGuardian, index) {
  const targetOsis = String(currentGuardian.student_id).trim();

  const getMatch = (g) => {
    if (!g || !g.guardianOutput) return null;
    const name = normalizeName_(stripTrailingParenthetical_(g.guardianOutput));
    
    // ✨ NEW: Look up using the strict OSIS + Name key
    const compositeKey = `${targetOsis}_${name}`;
    return index.byStudentAndName.get(compositeKey);
  };

  const currentMatch = getMatch(currentGuardian);
  const otherMatch = getMatch(otherGuardian);

  if (currentMatch) {
    const code = currentMatch.accountCode ? ` ${currentMatch.accountCode}` : "";
    const type = currentMatch.anotherCustodial === "Yes" ? "OG" : "NG";
    return `No (${type})${code}`;
  }

  if (otherMatch) {
    if (otherMatch.anotherCustodial === "No") return "-";
    if (otherMatch.anotherCustodial === "Yes") return "Yes/Unknown";
  }

  return "Yes/Unknown";
}

function normalizeName_(text) {
  return toText_(text).toLowerCase().replace(/\s+/g, " ").trim();
}

/* ========================= MASTER TABLE ========================= */
/**
 * Compiles the Master Table.
 * Preserves "Discharged" students by comparing 
 * current raw data with previously generated Master Table data.
 */
function buildMasterTableData_(ss, studentMap, guardianRows, siblingMap) {
  const out = [[
    "OSIS",
    "Student",
    "Guardian",
    "Site",
    "Class",
    "Sibling",
    "Status"
  ]];
  
  const guardiansByStudent = new Map();
  for (const g of guardianRows) {
    if (!guardiansByStudent.has(g.student_id)) guardiansByStudent.set(g.student_id, []);
    guardiansByStudent.get(g.student_id).push(stripTrailingParenthetical_(g.guardianOutput));
  }
  
  const oldMasterMap = new Map();
  const oldMasterSheet = ss.getSheetByName("Master Table");
  
  if (oldMasterSheet) {
    const oldData = oldMasterSheet.getDataRange().getValues();
    
    let headerRowIdx = -1;
    let oldHeaderMap = {};
    for (let i = 0; i < Math.min(10, oldData.length); i++) {
       if (String(oldData[i][0]).trim().toUpperCase() === "OSIS") {
           headerRowIdx = i;
           for (let c = 0; c < oldData[i].length; c++) {
               oldHeaderMap[String(oldData[i][c]).trim()] = c;
           }
           break;
       }
    }
    
    const startRow = headerRowIdx !== -1 ? headerRowIdx + 1 : 4;
    const hStudent = oldHeaderMap["Student"] !== undefined ? oldHeaderMap["Student"] : 1;
    const hGuardian = oldHeaderMap["Guardian"] !== undefined ? oldHeaderMap["Guardian"] : 2;
    const hSite = oldHeaderMap["Site"] !== undefined ? oldHeaderMap["Site"] : -1;
    const hClass = oldHeaderMap["Class"] !== undefined ? oldHeaderMap["Class"] : 3;
    const hSibling = oldHeaderMap["Sibling"] !== undefined ? oldHeaderMap["Sibling"] : 4;
    const hStatus = oldHeaderMap["Status"] !== undefined ? oldHeaderMap["Status"] : 5;

    for (let i = startRow; i < oldData.length; i++) {
      const row = oldData[i];
      const osis = toText_(row[0]);
      
      if (osis.includes(",")) continue;
      if (!isRealStudentId_(osis)) continue;
      
      if (!oldMasterMap.has(osis)) {
        oldMasterMap.set(osis, {
          student: toText_(row[hStudent]),
          guardian: toText_(row[hGuardian]),
          site: hSite !== -1 ? toText_(row[hSite]) : "",
          classCode: toText_(row[hClass]),
          sibling: toText_(row[hSibling]),
          status: toText_(row[hStatus])
        });
      }
    }
  }
  
  const currentIds = new Set(Array.from(studentMap.keys()));
  const allIds = new Set([ ...Array.from(oldMasterMap.keys()), ...Array.from(studentMap.keys()) ]);
  const ids = Array.from(allIds).sort();
  
  for (const id of ids) {
    const currentStudent = studentMap.get(id);
    const oldStudent = oldMasterMap.get(id);
    const isCurrent = currentIds.has(id);
    
    let student = "";
    let guardian = "";
    let site = "";
    let classCode = "";
    let sibling = "";
    let status = isCurrent ? "Current" : "Discharged";

    if (isCurrent) {
      student = currentStudent.student;
      site = currentStudent.site || "";
      
      classCode = currentStudent.teachers && currentStudent.teachers.trim() !== "" 
          ? `${currentStudent.classCode}: ${currentStudent.teachers}` 
          : currentStudent.classCode;
          
      guardian = Array.from(new Set(guardiansByStudent.get(id) || []))
        .filter(Boolean)
        .join(" and ");
      sibling = siblingMap.get(id) || "";
      
      // 🚀 INJECTS THE CUSTOM LABEL INSTEAD OF FORCING "CURRENT"
      if (currentStudent.label && currentStudent.label.trim() !== "") {
         status = currentStudent.label;
      }
      
    } else if (oldStudent) {
      student = oldStudent.student;
      guardian = oldStudent.guardian;
      site = oldStudent.site || "";
      classCode = oldStudent.classCode;
      sibling = oldStudent.sibling;
      
      const rawStatus = String(oldStudent.status).trim();
      const statLower = rawStatus.toLowerCase();
      if (statLower !== "current" && statLower !== "master" && rawStatus !== "") {
          status = rawStatus;
      }
    }
    
    out.push([ id, student, guardian, site, classCode, sibling, status ]);
  }
  
  // ==========================================
  // --- UPGRADED: FAMILY ALIAS GENERATOR ---
  // ==========================================
  const familyMap = {};
  
  for (let i = 1; i < out.length; i++) {
    const row = out[i];
    const osis = String(row[0]).trim();
    const student = row[1];
    const guardian = row[2];
    const site = row[3];
    const classCode = row[4];
    const siblingString = row[5];
    const status = row[6];
    
    let siblingOsisArray = [];
    if (siblingString && String(siblingString).trim() !== "") {
       const matches = String(siblingString).match(/\b\d{6,}\b/g);
       if (matches) siblingOsisArray = matches;
    }
    
    if (siblingOsisArray.length > 0) {
      const fullFamily = [osis, ...siblingOsisArray].sort();
      const familyKey = fullFamily.join("_");
      
      if (!familyMap[familyKey]) familyMap[familyKey] = [];
      familyMap[familyKey].push({
         osis: osis,
         student: student,
         site: site,
         pClass: classCode,
         status: status,
         originalRow: row
      });
    }
  }
  
  for (const key in familyMap) {
    const kids = familyMap[key];
    if (kids.length > 1) {
      const combinedOsis = kids.map(k => k.osis).join(", ");
      const combinedStudent = kids.map(k => k.student).join(" & ");
      const combinedSite = Array.from(new Set(kids.map(k => k.site).filter(Boolean))).join(" / ");
      const combinedClass = kids.map(k => k.pClass).join(" / ");
      
      const isFamilyCurrent = kids.some(k => k.status === "Current");
      const familyStatus = isFamilyCurrent ? "Current" : "Discharged";
      
      let comboRow = [...kids[0].originalRow];
      comboRow[0] = combinedOsis;
      comboRow[1] = combinedStudent;
      comboRow[3] = combinedSite;
      comboRow[4] = combinedClass;
      comboRow[6] = familyStatus;
      
      out.push(comboRow);
    }
  }
  
  // ==========================================
  // --- ALPHABETIZE BY LAST NAME ---
  // ==========================================
  const headers = out[0];
  const body = out.slice(1);
  
  body.sort((a, b) => {
    const extractLastName = (name) => {
      const firstStudent = String(name || "").split("&")[0].trim();
      const parts = firstStudent.split(/\s+/);
      return parts.length > 1 ? parts.pop().toLowerCase() : firstStudent.toLowerCase();
    };
    
    const lastNameA = extractLastName(a[1]);
    const lastNameB = extractLastName(b[1]);
    
    const lastNameCompare = lastNameA.localeCompare(lastNameB, undefined, { numeric: true, sensitivity: 'base' });
    
    if (lastNameCompare === 0) {
      return String(a[1] || "").localeCompare(String(b[1] || ""), undefined, { numeric: true, sensitivity: 'base' });
    }
    return lastNameCompare;
  });

  return [headers, ...body];
}

/* =========================
   PARENTS DIVIDED
========================= */

function buildParentsDividedData_(ss, guardianRows, siblingMap, parentSquareIndex, nycsaIndex, studentMap, contactLogData, eventLogData) {
  const out = [[
    "student_id",
    "Guardian output",
    "Student",
    "Class / Site",
    "relationship",
    "language",
    "email",
    "mobile",
    "Sibling",
    "Parents Square",
    "NYCSA",
    "Last Interaction"
  ]];

  const guardiansByStudentId = new Map();
  for (const g of guardianRows) {
    if (!guardiansByStudentId.has(g.student_id)) guardiansByStudentId.set(g.student_id, []);
    guardiansByStudentId.get(g.student_id).push(g);
  }

  const allInteractions = [];

  for (let i = 1; i < contactLogData.length; i++) {
    const row = contactLogData[i];
    
    // ✨ FIX: Contact Log shifted columns! Read the correct indexes:
    // OSIS is now index 4 (Col E), Hidden Checkbox is 15 (Col P), Person Spoke With is 3 (Col D)
    const rowMarker = typeof toText_ === "function" ? toText_(row[4]) : String(row[4]);
    if (rowMarker.toUpperCase() === "END") break;
    
    const isHidden = row[15]; 
    if (isHidden === true || String(isHidden).toLowerCase() === "true") continue;
    
    const dateVal = row[1];
    const personStr = String(row[3] || row[2] || "").toLowerCase().replace("▶", "").trim();
    const osisStr = String(row[4] || "").trim();
    
    if (dateVal && osisStr) {
       let d = new Date(dateVal);
       if (!isNaN(d)) allInteractions.push({ osisStr, personStr, dateObj: d, dateRaw: dateVal });
    }
  }

  for (let i = 1; i < eventLogData.length; i++) {
    const row = eventLogData[i];
    const dateVal = row[0];
    const osisStr = String(row[3] || "").trim();
    const personStr = String(row[5] || row[4] || "").toLowerCase().trim();
    
    if (dateVal && osisStr) {
       let d = new Date(dateVal);
       if (!isNaN(d)) allInteractions.push({ osisStr, personStr, dateObj: d, dateRaw: dateVal });
    }
  }

  allInteractions.sort((a, b) => a.dateObj.getTime() - b.dateObj.getTime());

  const interactionsByOsis = new Map();
  for (const interaction of allInteractions) {
    // ✨ UPGRADED: Smart number extraction handles siblings even if you separated them with "&" or "/"
    const matches = interaction.osisStr.match(/\b\d{6,}\b/g) || interaction.osisStr.split(",").map(p => p.trim()).filter(Boolean);
    for (const osis of matches) {
      if (!interactionsByOsis.has(osis)) interactionsByOsis.set(osis, []);
      interactionsByOsis.get(osis).push(interaction);
    }
  }

  for (const g of guardianRows) {
    const student = studentMap.get(g.student_id) || { student: "", classCode: "", site: "" };

    let classSiteStr = String(student.classCode || "");
    if (student.site && String(student.site).trim() !== "") {
      classSiteStr += classSiteStr ? ` (Site: ${student.site})` : `Site: ${student.site}`;
    }

    let lastInteraction = "None logged";
    const cleanGuardianName = stripTrailingParenthetical_(g.guardianOutput).toLowerCase().trim();
    const gFirstName = cleanGuardianName.split(" ")[0];
    const targetOsis = String(g.student_id).trim();

    const guardianInteractions = interactionsByOsis.get(targetOsis) || [];
    for (let i = guardianInteractions.length - 1; i >= 0; i--) {
      const interaction = guardianInteractions[i];
      
      if (interaction.personStr) {
        const logName = interaction.personStr;
        const logFirstName = logName.split(" ")[0];

        if (cleanGuardianName.includes(logName) || logName.includes(cleanGuardianName) || logFirstName === gFirstName) {
          if (typeof formatDateFull_ === "function") {
            lastInteraction = formatDateFull_(interaction.dateRaw);
          } else {
            lastInteraction = Utilities.formatDate(interaction.dateObj, Session.getScriptTimeZone(), "MMM d, yyyy");
          }
          break;
        }
      }
    }

    const otherGuardian = (guardiansByStudentId.get(g.student_id) || []).find(row => row !== g) || null;

    out.push([
      g.student_id,
      stripTrailingParenthetical_(g.guardianOutput), 
      student.student,
      classSiteStr,
      g.relationship,
      g.language,
      g.email,
      g.mobile,
      siblingMap.get(g.student_id) || "",
      getParentsSquareStatus_(g, parentSquareIndex),
      getNYCSAStatus_(g, otherGuardian, nycsaIndex),
      lastInteraction
    ]);
  }

  const headers = out[0];
  const body = out.slice(1);

  body.sort((a, b) => {
    const extractLastName = (name) => {
      const parts = String(name || "").trim().split(/\s+/);
      return parts.length > 1 ? parts.pop().toLowerCase() : String(name || "").toLowerCase();
    };
    
    const lastNameA = extractLastName(a[2]);
    const lastNameB = extractLastName(b[2]);
    
    const lastNameCompare = lastNameA.localeCompare(lastNameB, undefined, { numeric: true, sensitivity: 'base' });
    
    if (lastNameCompare === 0) {
      const studentCompare = String(a[2] || "").localeCompare(String(b[2] || ""), undefined, { numeric: true, sensitivity: 'base' });
      if (studentCompare === 0) {
        return String(a[1] || "").localeCompare(String(b[1] || ""), undefined, { numeric: true, sensitivity: 'base' });
      }
      return studentCompare;
    }
    return lastNameCompare;
  });

  return [headers, ...body];
}

/* =========================
   COMBINED CONTACT TRACKING
========================= */

function buildCombinedContactTrackingData_(ss, masterTable, contactLogData, eventLogData) {
  
  const props = PropertiesService.getDocumentProperties();
  const disableHiddenNotes = props.getProperty('disableHiddenNotes') === 'true'; 
  
  // --- Retrieve Custom Date Range Filter ---
  const notesStartDateStr = props.getProperty('notesStartDate');
  const notesEndDateStr = props.getProperty('notesEndDate');
  
  let filterStartDate = null;
  let filterEndDate = null;
  
  if (notesStartDateStr) {
    const parts = notesStartDateStr.split("-");
    if (parts.length === 3) filterStartDate = new Date(parts[0], parts[1] - 1, parts[2], 0, 0, 0);
  }
  if (notesEndDateStr) {
    const parts = notesEndDateStr.split("-");
    if (parts.length === 3) filterEndDate = new Date(parts[0], parts[1] - 1, parts[2], 23, 59, 59);
  }

  const out = [[ "OSIS", "Guardians", "Student", "Notes" ]];
  
  const startIndex = (masterTable.length > 0 && String(masterTable[0][0]).toUpperCase() === "OSIS") ? 1 : 0;
  
  const masterMap = new Map(
    masterTable
      .slice(startIndex)
      .filter(r => typeof isRealStudentId_ === "function" && isRealStudentId_(r[0]) && !String(r[0]).includes(","))
      .map(r => [
        typeof toText_ === "function" ? toText_(r[0]) : String(r[0]),
        { guardians: typeof toText_ === "function" ? toText_(r[2]) : String(r[2]), student: typeof toText_ === "function" ? toText_(r[1]) : String(r[1]) }
      ])
  );

  const oldCombinedSheet = ss.getSheetByName("Combined Contact Tracking");
  const oldCombinedMap = new Map();
  if (oldCombinedSheet) {
    const oldData = oldCombinedSheet.getDataRange().getValues();
    for (let i = 1; i < oldData.length; i++) {
      const row = oldData[i];
      const osis = typeof toText_ === "function" ? toText_(row[0]) : String(row[0]);
      if (typeof isRealStudentId_ === "function" && !isRealStudentId_(osis)) continue;
      
      if (!oldCombinedMap.has(osis)) {
        oldCombinedMap.set(osis, {
          guardians: typeof toText_ === "function" ? toText_(row[1]) : String(row[1]),
          student: typeof toText_ === "function" ? toText_(row[2]) : String(row[2])
        });
      }
    }
  }

  // --- NEW: Retrieve Manual Notes (Matches Contact Log Array Logic) ---
  const notesSheet = ss.getSheetByName("Notes");
  const usableNotesRows = [];
  if (notesSheet) {
    const nData = notesSheet.getDataRange().getValues();
    for (let i = 1; i < nData.length; i++) {
      if (!nData[i]) continue;
      const rawOsis = typeof toText_ === "function" ? toText_(nData[i][0]) : String(nData[i][0]);
      if (rawOsis.trim() !== "") {
        usableNotesRows.push(nData[i]);
      }
    }
  }

  const usableContactRows = [];
  for (let i = 1; i < contactLogData.length; i++) {
    const row = contactLogData[i];
    const rowMarker = typeof toText_ === "function" ? toText_(row[4]) : String(row[4]);
    if (rowMarker.toUpperCase() === "END") break;
    usableContactRows.push(row);
  }

  const usableEventRows = eventLogData && eventLogData.length > 1 ? eventLogData.slice(1) : [];

  const osisSet = new Set();
  
  // Tokenize into whole OSIS numbers so siblings formatted as "222222222, 333333333"
  // (or with & / separators) each register independently.
  const addOsisToSet = (osisString) => {
    extractOsisTokens_(osisString).forEach(cleanOsis => {
      if (isRealStudentId_(cleanOsis)) osisSet.add(cleanOsis);
    });
  };

  usableContactRows.forEach(r => addOsisToSet(r[4])); 
  usableEventRows.forEach(r => addOsisToSet(r[3])); 
  usableNotesRows.forEach(r => addOsisToSet(r[0])); 

  for (const osis of osisSet) {
    const combinedInteractions = [];

    // --- Process Contact Log ---
    for (const row of usableContactRows) {
      if (!osisCellIncludes_(row[4], osis)) continue;
      
      const isHidden = row[15]; 
      const hiddenBool = (isHidden === true || String(isHidden).toUpperCase() === "TRUE");
      
      if (!disableHiddenNotes && hiddenBool) continue;

      let dateVal = row[1];
      if (!dateVal) dateVal = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");

      let dateObj = new Date(dateVal);
      if (isNaN(dateObj)) dateObj = new Date(0);
      
      if (filterStartDate && dateObj < filterStartDate) continue;
      if (filterEndDate && dateObj > filterEndDate) continue;

      const eVal = typeof toText_ === "function" ? toText_(row[3]) : String(row[3]); 
      const mVal = typeof toText_ === "function" ? toText_(row[9]) : String(row[9]);  
      const nVal = typeof toText_ === "function" ? toText_(row[10]) : String(row[10]); 
      const pVal = typeof toText_ === "function" ? toText_(row[12]) : String(row[12]); 
      const rVal = typeof toText_ === "function" ? toText_(row[14]) : String(row[14]); 

      if (pVal === "" && rVal === "") continue;

      const methodLabel = nVal !== "" ? nVal : mVal;
      const dateText = typeof formatDateFull_ === "function" ? formatDateFull_(dateVal) : dateVal;
      
      let line = `• ${dateText} - {${methodLabel} - ${eVal}} ${pVal}`;
      if (rVal !== "") line += ` [${rVal}]`;

      combinedInteractions.push({
        dateObj: dateObj,
        text: line.trim()
      });
    }

    // --- Process Event Log ---
    for (const row of usableEventRows) {
      if (!osisCellIncludes_(row[3], osis)) continue;

      const dateVal = row[0];
      
      let dateObj = new Date(dateVal);
      if (isNaN(dateObj)) dateObj = new Date(0);
      
      if (filterStartDate && dateObj < filterStartDate) continue;
      if (filterEndDate && dateObj > filterEndDate) continue;

      const eventName = typeof toText_ === "function" ? toText_(row[1]) : String(row[1]);
      const guardian = typeof toText_ === "function" ? toText_(row[5]) : String(row[5]);
      const attendee = typeof toText_ === "function" ? toText_(row[6]) : String(row[6]);
      const notes = typeof toText_ === "function" ? toText_(row[8]) : String(row[8]);

      let person = attendee ? attendee : guardian;
      if (!person) person = "Parent";

      const dateText = typeof formatDateFull_ === "function" ? formatDateFull_(dateVal) : dateVal;
      let line = `• ${dateText} - {Event - [${person}]} |${eventName}|`;
      if (notes) line += ` ${notes}`;

      combinedInteractions.push({
        dateObj: dateObj,
        text: line.trim()
      });
    }

    // --- Process Manual Notes ---
    let mergedManualNotes = "";
    for (const nRow of usableNotesRows) {
      if (!osisCellIncludes_(nRow[0], osis)) continue;
      
      const nText = typeof toText_ === "function" ? toText_(nRow[3]) : String(nRow[3]);
      if (nText.trim() !== "") {
        mergedManualNotes = mergedManualNotes ? mergedManualNotes + "\n\n" + nText.trim() : nText.trim();
      }
    }

    let baseNotes = "";
    if (combinedInteractions.length > 0) {
      combinedInteractions.sort((a, b) => a.dateObj.getTime() - b.dateObj.getTime());
      baseNotes = combinedInteractions.map(interaction => interaction.text).join("\n");
    }
    
    const manualNotes = mergedManualNotes !== "" ? `𝑵𝒐𝒕𝒆𝒔: ${mergedManualNotes}` : "";
    
    let mergedNotes = "";
    if (manualNotes !== "" && baseNotes !== "") {
      mergedNotes = manualNotes + "\n\n" + baseNotes;
    } else if (manualNotes !== "") {
      mergedNotes = manualNotes;
    } else if (baseNotes !== "") {
      mergedNotes = baseNotes;
    }

    if (mergedNotes !== "") {
      const master = masterMap.get(osis);
      const old = oldCombinedMap.get(osis);
      
      const guardians = master?.guardians || old?.guardians || "";
      const student = master?.student || old?.student || "";

      out.push([osis, guardians, student, mergedNotes.trim()]);
    }
  }

  const headers = out[0];
  const body = out.slice(1);
  
  body.sort((a, b) => {
    const extractLastName = (name) => {
      const firstStudent = String(name || "").split("&")[0].trim();
      const parts = firstStudent.split(/\s+/);
      return parts.length > 1 ? parts.pop().toLowerCase() : firstStudent.toLowerCase();
    };
    
    const lastNameA = extractLastName(a[2]);
    const lastNameB = extractLastName(b[2]);
    
    const lastNameCompare = lastNameA.localeCompare(lastNameB, undefined, { numeric: true, sensitivity: 'base' });
    
    if (lastNameCompare === 0) {
      return String(a[2] || "").localeCompare(String(b[2] || ""), undefined, { numeric: true, sensitivity: 'base' });
    }
    return lastNameCompare;
  });

  return [headers, ...body];
}

function isTruthy_(val) {
  return val === true || String(val).toLowerCase() === "true";
}

/* =========================
   DIRECTORY
========================= */

function buildDirectoryData_(studentMap, guardianRows) {
  const headers = ["Last Name", "First Name", "OSIS", "Site", "Class/ Teacher", "Parent Contact"];
  const out = [headers];
  const grouped = groupGuardiansByStudent_(guardianRows);

  for (const [id, student] of studentMap) {
    const guardians = grouped.get(id) || [];
    const parentContact = guardians.map(formatGuardianDirectoryLine_).filter(Boolean).join("\n");

    const parsedStudent = splitFullName_(student.student);
    const classTeacherInfo = student.teachers 
      ? `${student.classCode}: ${student.teachers}` 
      : student.classCode;

    const row = [parsedStudent.last, parsedStudent.first, id, student.site || "", classTeacherInfo, parentContact];
    out.push(row);
  }

  const body = out.slice(1).sort((a, b) => {
    const siteA = String(a[3] || "").trim();
    const siteB = String(b[3] || "").trim();
    const primarySort = siteA.localeCompare(siteB, undefined, { numeric: true, sensitivity: "base" });
    if (primarySort !== 0) return primarySort;

    const classA = String(a[4] || "").trim();
    const classB = String(b[4] || "").trim();
    const secondarySort = classA.localeCompare(classB, undefined, { numeric: true, sensitivity: "base" });
    if (secondarySort !== 0) return secondarySort;

    const lastNameA = String(a[0] || "").trim();
    const lastNameB = String(b[0] || "").trim();
    return lastNameA.localeCompare(lastNameB, undefined, { numeric: true, sensitivity: "base" });
  });

  return [out[0], ...body];
}

function formatDirectoryAuto_(sheet) {
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;

  const siteValues = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
  const hasSiteData = siteValues.some(r => String(r[0]).trim() !== "");

  const targetCol = hasSiteData ? 4 : 5; 
  const targetValues = sheet.getRange(2, targetCol, lastRow - 1, 1).getValues();

  const colors = [
    "#FFB3BA", "#BAE1FF", "#BAFFC9", "#FFFFBA", "#CBAACB", "#FFDFBA", 
    "#A9DEF9", "#D0F0C0", "#FFFACD", "#E2B2E8", "#FFDAB3", "#F6D2E6",
    "#B6CCFE", "#FCF6BD", "#FF99C8", "#D7E3FC", "#D0F4DE", "#FFE5D9"
  ];

  const map = new Map();
  let idx = 0;
  const columnAColors = [];

  for (let i = 0; i < targetValues.length; i++) {
    const val = String(targetValues[i][0] || "").trim();
    if (!val) {
      columnAColors.push([null]); 
      continue;
    }
    if (!map.has(val)) {
      map.set(val, colors[idx % colors.length]);
      idx++;
    }
    columnAColors.push([map.get(val)]);
  }

  sheet.getRange(2, 1, lastRow - 1, sheet.getMaxColumns()).setBackground(null);
  sheet.getRange(2, 1, columnAColors.length, 1).setBackgrounds(columnAColors);
  sheet.autoResizeColumns(1, sheet.getLastColumn());
}

function formatGuardianDirectoryLine_(g) {
  const left = g.guardianOutput;
  const parts = [toText_(g.email), toText_(g.mobile)].filter(Boolean);
  if (!left) return "";
  if (parts.length === 0) return left;
  return `${left} - ${parts.join(", ")}`;
}

/* =========================
   PHONE CONTACTS
========================= */

function buildPhoneContactsData_(ss, studentMap, guardianRows, masterTable, combinedContactTrackingArray, parentSquareIndex, nycsaIndex) {
  const out = [[
    "Name Prefix", "First Name", "Middle Name", "Last Name", "Name Suffix", 
    "Phonetic First Name", "Phonetic Middle Name", "Phonetic Last Name", "Nickname", "File As", 
    "E-mail 1 - Label", "E-mail 1 - Value", "E-mail 2 - Label", "E-mail 2 - Value", "E-mail 3 - Label", 
    "E-mail 3 - Value", "Phone 1 - Type", "Phone 1 - Value", "Phone 2 - Type", "Phone 2 - Value", 
    "Phone 3 - Type", "Phone 3 - Value", "Organization Name", "Organization Department", "Birthday", 
    "Event 1 - Label", "Event 1 - Value", "Relation 1 - Label", "Relation 1 - Value", "Website 1 - Label", 
    "Website 1 - Value", "Custom Field 1 - Label", "Custom Field 1 - Value", "Notes", "Labels", "OSIS"
  ]];

  const masterMap = new Map(
    masterTable.slice(1).filter(r => isRealStudentId_(r[0])).map(r => [toText_(r[0]), r])
  );
  
  const tracking = buildTrackingNotesMap_(ss, combinedContactTrackingArray); 

  const guardiansByStudentId = new Map();
  for (const g of guardianRows) {
    if (!guardiansByStudentId.has(g.student_id)) guardiansByStudentId.set(g.student_id, []);
    guardiansByStudentId.get(g.student_id).push(g);
  }

  const groupedGuardians = new Map();

  for (const g of guardianRows) {
    const id = g.student_id;
    const master = masterMap.get(id) || [];
    const student = studentMap.get(id) || { student: "", classCode: "", site: "", teachers: "" };
    const guardianName = stripTrailingParenthetical_(g.guardianOutput);

    const key = `${guardianName.toLowerCase()}|${toText_(g.email).toLowerCase()}|${toText_(g.mobile)}`;

    if (!groupedGuardians.has(key)) {
      groupedGuardians.set(key, {
        guardianName: guardianName,
        relationship: toText_(g.relationship),
        email: toText_(g.email),
        phone: toText_(g.mobile),
        students: []
      });
    }

    const group = groupedGuardians.get(key);

    let classDisplay = toText_(student.classCode);
    if (student.teachers && toText_(student.teachers).trim() !== "") {
      classDisplay += ` - ${toText_(student.teachers)}`;
    } else if (master[4] && toText_(master[4]).trim() !== "") { 
      classDisplay = toText_(master[4]); 
    }

    const otherGuardian = (guardiansByStudentId.get(id) || []).find(row => row !== g) || null;

    const psStatus = getParentsSquareStatus_(g, parentSquareIndex);
    const nycsaStatus = getNYCSAStatus_(g, otherGuardian, nycsaIndex); 

    let studentBlock = `STUDENT NAME: ${toText_(student.student)}\n` +
                       `RELATIONSHIP: ${toText_(g.relationship)}\n` +
                       `LANGUAGE: ${toText_(g.language)}\n` +
                       `OSIS: ${id}\n` +
                       `Parent Name: ${toText_(master[2] || guardianName)}\n` +
                       `Class: ${classDisplay}`;
    
    if (student.site && toText_(student.site) !== "") studentBlock += `\nSite: ${toText_(student.site)}`;
    
    studentBlock += `\nParents Square: ${psStatus || "No"}`;
    studentBlock += `\nNYCSA: ${nycsaStatus || "No"}`;

    const studentNotes = tracking.get(id);
    const cleanedNotes = (studentNotes && studentNotes.trim() !== "") ? studentNotes.trim() : "No Notes";
    
    studentBlock += `\n\nCommunications:\n${cleanedNotes}`;

    group.students.push({
      studentName: toText_(student.student),
      osis: id,
      block: studentBlock,
      site: toText_(student.site),
      classCode: classDisplay
    });
  }

  for (const group of groupedGuardians.values()) {
    const studentNames = group.students.map(s => s.studentName).join(", ");
    const osisList = group.students.map(s => s.osis).join(", ");
    const finalNotes = group.students.map(s => s.block).join("\n\n---\n");

    const hasSiblings = group.students.length > 1;

    const phoneLabels = group.students.map(s => {
      let label = "";
      if (hasSiblings) {
        const firstName = String(s.studentName).split(" ")[0];
        label += `{${firstName}} `;
      }
      if (s.site && s.site !== "") label += `Site: ${s.site}, `;
      label += `Class: ${s.classCode}`;
      return label.trim();
    }).join(" & ");

    const orgNames = group.students.map(s => {
      if (hasSiblings) {
        const firstName = String(s.studentName).split(" ")[0];
        return `{${firstName}} OSIS: ${s.osis}`;
      } else {
        return `OSIS: ${s.osis}`;
      }
    }).join(", ");

    const row = new Array(36).fill("");
    
    row[1]  = `(${group.relationship}) ${group.guardianName}`.trim();
    row[3]  = `| Student: ${studentNames} |`;
    row[11] = group.email;
    row[16] = phoneLabels;
    row[17] = group.phone;
    row[22] = orgNames;
    row[33] = finalNotes;
    row[34] = "School Directory";
    row[35] = osisList;

    out.push(row);
  }
  
  return out;
}

function buildTrackingNotesMap_(ss, combinedContactTrackingArray) { 
  const map = new Map();
  
  if (combinedContactTrackingArray && combinedContactTrackingArray.length > 0) {
    for (let i = 1; i < combinedContactTrackingArray.length; i++) {
      const osis = toText_(combinedContactTrackingArray[i][0]);
      if (!isRealStudentId_(osis)) continue;
      if (!map.has(osis)) {
        map.set(osis, toText_(combinedContactTrackingArray[i][3])); 
      }
    }
    return map;
  }

  const sheet = ss.getSheetByName("Combined Contact Tracking");
  if (!sheet) return map;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const osis = toText_(data[i][0]);
    if (!isRealStudentId_(osis)) continue;
    if (!map.has(osis)) {
      map.set(osis, toText_(data[i][3]));
    }
  }
  return map;
}

/* =========================
   OVERRIDE MAP GENERATOR (Build.gs)
========================= */

function getOverrideData_(ss) {
  const sheet = ss.getSheetByName("New/Edit Student");
  const overrideStudents = new Map();
  const overrideGuardians = new Map();

  if (!sheet) return { students: overrideStudents, guardians: overrideGuardians };

  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const id = toText_(r[2]);
    if (!isRealStudentId_(id)) continue;
    
    const lastName = toText_(r[3]);
    const firstName = toText_(r[4]);
    const classCode = toText_(r[8]).toUpperCase();
    
    overrideStudents.set(id, {
      student: properCase_(`${firstName} ${lastName}`.trim()),
      classCode: classCode,
      site: "",
      teachers: "",
      label: toText_(r[40]) // <-- Retrieves the new Custom Label
    });

    const guardianList = [];
    for (let g = 0; g < 5; g++) {
      const baseCol = 10 + (g * 6); 
      if (toText_(r[baseCol])) {
        const gInfo = buildGuardianObject_(
          id, 
          r[baseCol],
          r[baseCol + 1],
          r[baseCol + 4],
          r[baseCol + 5], // Passed twice in UI mapping
          r[baseCol + 3],
          r[baseCol + 2]
        );
        if (gInfo) guardianList.push(gInfo);
      }
    }
    overrideGuardians.set(id, guardianList);
  }

  return { students: overrideStudents, guardians: overrideGuardians };
}

/* =========================
   MASTER TABLE PIPELINE
========================= */

function buildStudentMap_(rawData, overrideStudents) {
  const map = new Map();
  const classRefMap = buildClassReferenceMap_(rawData);

  for (let i = 0; i < rawData.length; i++) {
    const r = rawData[i];
    const id = toText_(r[2]);
    if (!isRealStudentId_(id)) continue;

    if (overrideStudents.has(id)) continue; 

    const studentName = properCase_(`${toText_(r[4])} ${toText_(r[3])}`.trim());
    const classCode = toText_(r[17]).toUpperCase();

    let assignedSite = "";
    let assignedTeacher = "";
    if (classRefMap.has(classCode)) {
      const ref = classRefMap.get(classCode);
      assignedSite = ref.site;
      assignedTeacher = ref.teacher;
    }

    if (!map.has(id)) {
      map.set(id, {
        student: studentName,
        classCode: classCode,
        site: assignedSite,
        teachers: assignedTeacher,
        label: "Current" // Default ATS Status
      });
    }
  }

  for (const [id, studentObj] of overrideStudents.entries()) {
    if (classRefMap.has(studentObj.classCode)) {
      const ref = classRefMap.get(studentObj.classCode);
      studentObj.site = ref.site;
      studentObj.teachers = ref.teacher;
    }
    map.set(id, studentObj);
  }

  return map;
}

/* =========================
   BASE BUILDERS
========================= */

function buildGuardianObjects_(rawData, overrideGuardians) {
  const out = [];

  for (let i = 0; i < rawData.length; i++) {
    const r = rawData[i];
    const id = toText_(r[2]);
    if (!isRealStudentId_(id)) continue;

    if (overrideGuardians.has(id)) continue;

    const g1 = buildGuardianObject_(id, r[19], r[20], r[23], r[24], r[22], r[21]);
    if (g1) out.push(g1);

    const g2 = buildGuardianObject_(id, r[25], r[26], r[29], r[30], r[28], r[27]);
    if (g2) out.push(g2);
  }

  for (const guardianList of overrideGuardians.values()) {
    out.push(...guardianList);
  }

  return out;
}

function buildGuardianObject_(studentId, fullName, relationship, spokenLanguage, writtenLanguage, email, phone) {
  const name = toText_(fullName);
  if (!name) return null;

  let language = "";
  if (toText_(spokenLanguage)) {
    language = properCase_(firstWord_(spokenLanguage));
  } else if (toText_(writtenLanguage)) {
    language = properCase_(firstWord_(writtenLanguage));
  }

  let relTag = properCase_(relationship);
  if (language && language.toLowerCase() !== "english") {
    relTag = relTag ? `${relTag} - ${language}` : language;
  }

  return {
    student_id: studentId,
    guardianOutput: properCase_(name) + (relTag ? ` (${relTag})` : ""),
    relationship: properCase_(relationship),
    language: language,
    email: toText_(email),
    mobile: formatPhone_(phone)
  };
}

function buildClassReferenceMap_(rawData) {
  const map = new Map();
  const regex = /([A-Za-z0-9]+):([^:]*?)(?=\s+[A-Za-z0-9]+:|$)/g;

  for (let i = 0; i < rawData.length; i++) {
    const siteVal = toText_(rawData[i][59]);
    const classesVal = toText_(rawData[i][60]);

    if (!classesVal) continue; 

    const matches = [...classesVal.matchAll(regex)];
    
    for (const m of matches) {
      const classCode = m[1].trim().toUpperCase();
      const teacher = m[2].trim();
      
      map.set(classCode, { 
        site: siteVal, 
        teacher: teacher 
      });
    }
  }
  return map;
}

/* =========================
   SIBLINGS
========================= */

function buildSiblingMap_(studentMap, guardianRows) {
  const nameEmailMap = new Map();
  const namePhoneMap = new Map();

  for (const g of guardianRows) {
    const id = g.student_id;
    if (!studentMap.has(id)) continue;

    const guardianName = stripTrailingParenthetical_(toText_(g.guardianOutput)).toLowerCase();
    if (!guardianName) continue; 

    const email = normalizeEmail_(g.email);
    const phone = normalizePhoneForMatch_(g.mobile);

    if (email) {
      const emailKey = `${guardianName}|${email}`;
      if (!nameEmailMap.has(emailKey)) nameEmailMap.set(emailKey, new Set());
      nameEmailMap.get(emailKey).add(id);
    }

    if (phone) {
      const phoneKey = `${guardianName}|${phone}`;
      if (!namePhoneMap.has(phoneKey)) namePhoneMap.set(phoneKey, new Set());
      namePhoneMap.get(phoneKey).add(id);
    }
  }

  const rowsByStudent = new Map();
  for (const g of guardianRows) {
    if (!rowsByStudent.has(g.student_id)) rowsByStudent.set(g.student_id, []);
    rowsByStudent.get(g.student_id).push(g);
  }

  const siblingMap = new Map();

  for (const [id] of studentMap) {
    const siblings = new Map();
    const rows = rowsByStudent.get(id) || [];

    for (const g of rows) {
      const guardianName = stripTrailingParenthetical_(toText_(g.guardianOutput)).toLowerCase();
      if (!guardianName) continue;

      const email = normalizeEmail_(g.email);
      const phone = normalizePhoneForMatch_(g.mobile);

      if (email) {
        const emailKey = `${guardianName}|${email}`;
        if (nameEmailMap.has(emailKey)) {
          nameEmailMap.get(emailKey).forEach(otherId => {
            if (otherId !== id && studentMap.has(otherId)) {
              siblings.set(otherId, `${otherId} - ${studentMap.get(otherId).student}`);
            }
          });
        }
      }

      if (phone) {
        const phoneKey = `${guardianName}|${phone}`;
        if (namePhoneMap.has(phoneKey)) {
          namePhoneMap.get(phoneKey).forEach(otherId => {
            if (otherId !== id && studentMap.has(otherId)) {
              siblings.set(otherId, `${otherId} - ${studentMap.get(otherId).student}`);
            }
          });
        }
      }
    }

    siblingMap.set(id, Array.from(siblings.values()).sort().join(", "));
  }

  return siblingMap;
}
