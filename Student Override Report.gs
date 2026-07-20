/**
 * Generates a visual side-by-side Delta report comparing Raw Data vs. Overridden Data.
 * Skips any records that have "Stealth Mode / Hide from Report" checked!
 */
function buildOverrideReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportSheetName = "Students Override Report";
  
  const rawSheet = ss.getSheetByName("RAW Data") || ss.getSheetByName("Raw Data");
  const overrideSheet = ss.getSheetByName("New/Edit Student");
  
  if (!rawSheet || !overrideSheet) {
    SpreadsheetApp.getUi().alert("Error: Could not find 'Raw Data' or 'New/Edit Student' sheets.");
    return;
  }

  // 1. Read Data
  const rawData = rawSheet.getDataRange().getValues().slice(6); 
  const overrideData = overrideSheet.getDataRange().getValues().slice(1);

  // 2. Map RAW Data by OSIS
  const rawMap = new Map();
  rawData.forEach(row => {
    const osis = String(row[2]).trim();
    if (osis) rawMap.set(osis, row);
  });

  // 3. Helper to Extract Guardians
  function extractGuardians(row, startCol, maxGuardians) {
    const guardians = [];
    for (let i = 0; i < maxGuardians; i++) {
      const base = startCol + (i * 6);
      const name = String(row[base] || "").trim();
      if (!name) continue; 
      
      // Pulls Spoken Lang (base+4) since Spoken and Written are synced now
      guardians.push({
        name: name,
        phone: String(row[base + 2] || "").trim(),
        email: String(row[base + 3] || "").trim(),
        lang: String(row[base + 4] || "").trim() 
      });
    }
    return guardians;
  }

  // 4. Helper to format a full parent block
  function formatFullParent(g) {
    return `Guardian: ${g.name}\n  Phone: ${g.phone || "None"}\n  Email: ${g.email || "None"}\n  Lang: ${g.lang || "None"}`;
  }

  const reportOutput = [];
  const richTextNewInfo = []; 
  const boldStyle = SpreadsheetApp.newTextStyle().setBold(true).build();

  // Add Headers
  reportOutput.push(["Student OSIS", "Old Information", "New Information"]);

  // 5. Main Comparison Engine
  overrideData.forEach(oRow => {
    const osis = String(oRow[2]).trim();
    if (!osis) return; 

    // THE GATEKEEPER: If Stealth Mode is on, skip this student entirely!
    const hideFromReport = String(oRow[41]).trim().toLowerCase() === "true";
    if (hideFromReport) return;

    const newStudentName = `${String(oRow[4] || "").trim()} ${String(oRow[3] || "").trim()}`.trim();
    const newGuardians = extractGuardians(oRow, 10, 5);

    // --- CASE A: BRAND NEW STUDENT ---
    if (!rawMap.has(osis)) {
      let newBlocks = [`Student Name: ${newStudentName}`];
      let namesToBold = [];

      newGuardians.forEach(ng => {
        newBlocks.push(formatFullParent(ng));
        namesToBold.push(ng.name);
      });

      const newInfoString = newBlocks.join("\n\n");
      const rtv = SpreadsheetApp.newRichTextValue().setText(newInfoString);
      
      namesToBold.forEach(name => {
        let startIdx = newInfoString.indexOf(`Guardian: ${name}`);
        if (startIdx !== -1) rtv.setTextStyle(startIdx + 10, startIdx + 10 + name.length, boldStyle);
      });

      reportOutput.push([osis, "[BRAND NEW STUDENT - NOT IN ATS YET]", newInfoString]);
      richTextNewInfo.push([rtv.build()]);
      return; 
    }

    // --- CASE B: EXISTING STUDENT COMPARISON ---
    const rRow = rawMap.get(osis);
    const oldStudentName = `${String(rRow[4] || "").trim()} ${String(rRow[3] || "").trim()}`.trim();
    const oldGuardians = extractGuardians(rRow, 19, 2);

    let oldBlocks = [];
    let newBlocks = [];
    let namesToBold = [];
    let hasChanges = false;

    // 1. Check Student Name
    if (newStudentName.toLowerCase() !== oldStudentName.toLowerCase()) {
      oldBlocks.push(`Student Name: ${oldStudentName}`);
      newBlocks.push(`Student Name: ${newStudentName}`);
      hasChanges = true;
    }

    // 2. Check Added & Modified Guardians
    newGuardians.forEach(ng => {
      const og = oldGuardians.find(g => g.name.toLowerCase() === ng.name.toLowerCase());
      
      if (!og) {
        // ADDED NEW PARENT
        oldBlocks.push(`Guardian: [Not in ATS]`);
        newBlocks.push(formatFullParent(ng));
        namesToBold.push(ng.name);
        hasChanges = true;
      } else {
        // MODIFIED EXISTING PARENT
        let oldF = [];
        let newF = [];
        
        if (og.phone !== ng.phone) { oldF.push(`Phone: ${og.phone || "None"}`); newF.push(`Phone: ${ng.phone || "None"}`); }
        if (og.email !== ng.email) { oldF.push(`Email: ${og.email || "None"}`); newF.push(`Email: ${ng.email || "None"}`); }
        if (og.lang !== ng.lang) { oldF.push(`Lang: ${og.lang || "None"}`); newF.push(`Lang: ${ng.lang || "None"}`); }

        if (newF.length > 0) {
          oldBlocks.push(`Guardian: ${og.name}\n  ${oldF.join("\n  ")}`);
          newBlocks.push(`Guardian: ${ng.name}\n  ${newF.join("\n  ")}`);
          namesToBold.push(ng.name);
          hasChanges = true;
        }
      }
    });

    // 3. Check Removed Guardians
    oldGuardians.forEach(og => {
      const ng = newGuardians.find(g => g.name.toLowerCase() === og.name.toLowerCase());
      if (!ng) {
        oldBlocks.push(formatFullParent(og));
        newBlocks.push(`Guardian: ${og.name} [REMOVED]`);
        namesToBold.push(og.name);
        hasChanges = true;
      }
    });

    // 4. Push to Output if differences exist
    if (hasChanges) {
      const newInfoString = newBlocks.join("\n\n");
      const rtv = SpreadsheetApp.newRichTextValue().setText(newInfoString);
      
      namesToBold.forEach(name => {
        let startIdx = newInfoString.indexOf(`Guardian: ${name}`);
        if (startIdx !== -1) rtv.setTextStyle(startIdx + 10, startIdx + 10 + name.length, boldStyle);
      });

      reportOutput.push([osis, oldBlocks.join("\n\n"), newInfoString]);
      richTextNewInfo.push([rtv.build()]);
    }
  });

  // 6. Write to Sheet
  let reportSheet = ss.getSheetByName(reportSheetName);
  if (!reportSheet) {
    reportSheet = ss.insertSheet(reportSheetName);
  } else {
    reportSheet.clear();
  }

  // Handle case where NO overrides differ from Raw Data
  if (reportOutput.length === 1) {
    reportOutput.push(["-", "No differences found between Overrides and ATS Data.", "-"]);
    richTextNewInfo.push([SpreadsheetApp.newRichTextValue().setText("-").build()]);
  }

  const targetRange = reportSheet.getRange(1, 1, reportOutput.length, 3);
  targetRange.setValues(reportOutput);

  // Overwrite ONLY Column C starting at Row 2 with the Bolded Rich Text
  if (richTextNewInfo.length > 0) {
    reportSheet.getRange(2, 3, richTextNewInfo.length, 1).setRichTextValues(richTextNewInfo);
  }

  // 7. Visual Formatting
  reportSheet.setColumnWidth(1, 100); 
  reportSheet.setColumnWidth(2, 400); 
  reportSheet.setColumnWidth(3, 400); 
  
  targetRange.setWrap(true).setVerticalAlignment("middle");
  
  reportSheet.getRange("A1:C1")
    .setFontWeight("bold")
    .setBackground("#f3f3f3")
    .setHorizontalAlignment("center");

  reportSheet.setFrozenRows(1);
  reportSheet.showSheet(); 
  reportSheet.activate(); 

  ss.toast("Delta Override Report generated successfully!", "Report Ready", 5);
}
