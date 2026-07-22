/* =========================
MAIN EXECUTION FUNCTIONS
========================= */

function resetToDisplayMode() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    "⚠️ Wipe ALL data?",
    "This permanently deletes Raw Data, the Contact Log, Events, and all derived sheets, leaving an empty display-mode template. This cannot be undone.\n\nAre you sure?",
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  deleteLegacySheets_(ss); // NEW: Clean up old tabs

  resetRawData_(ss);
  writeRawDataHeaders_(ss);
  resetDerivedSheets_(ss);
  resetEventsSystem_(ss); // NEW: Clears Events & Backend Log
  
  // Shrink the Contact Log to exactly 3 rows (plus the END row)
  resizeContactLog_(ss, 3);
  
  resetContactLog_(ss);
  clearTypeDropdownColors_(ss); // Wipe the dropdown colors!
  resetMasterTable_(ss);
  fillContactLogHstackFormulas_(ss);

  // Stamp this template's build date so the Version Checker knows how old
  // this copy is (see VersionCheck.gs — comparison is now date-based).
  stampVersionDate_(ss);

  SpreadsheetApp.getUi().alert("Display Mode Reset complete.");
}

/**
 * Writes today's date into the hidden "Version" tab, cell A1. This is the
 * build date of this template copy; the Version Checker compares it against
 * the dated rows in the master tracker. Creates (and hides) the tab if it
 * doesn't exist yet. The value travels with every copy of the sheet.
 */
function stampVersionDate_(ss) {
  let sheet = ss.getSheetByName("Version");
  if (!sheet) {
    sheet = ss.insertSheet("Version");
    sheet.hideSheet();
  }
  const today = new Date();
  today.setHours(0, 0, 0, 0); // midnight, so date-only comparisons are clean
  sheet.getRange("A1").setValue(today).setNumberFormat("MM/dd/yy");

  // B2 = Auto-Migrate toggle ("Yes"/"No"). Label it (A2, pointing right at
  // B2), and default it to "No" only when blank so we never overwrite a
  // template author's existing choice.
  sheet.getRange("A1").setNote("Build date — set automatically on Reset to Display Mode.");
  sheet.getRange("A2").setValue("Auto-Migrate on upgrade? (Yes/No) →");
  const toggleCell = sheet.getRange("B2");
  if (String(toggleCell.getValue()).trim() === "") toggleCell.setValue("No");
}

function resetAndLoadFakeData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    "⚠️ Wipe ALL data and load FAKE data?",
    "This permanently deletes Raw Data, the Contact Log, Events, and all derived sheets, then loads fake demo data. This cannot be undone.\n\nAre you sure?",
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  deleteLegacySheets_(ss); // NEW: Clean up old tabs

  const sheet = ss.getSheetByName("Raw Data");

  resetRawData_(ss);
  writeRawDataHeaders_(ss);

  // --- DATA MODE FORMATTING ---
  sheet.getRange("A5").setValue("No effective students filters. Data selection: Guardian contact information.")
    .setFontColor("#000000").setFontWeight("normal").setFontFamily("Roboto");

  sheet.getRange("AG6").setValue("Name")
    .setFontColor("#000000").setFontWeight("bold").setFontFamily("Roboto");

  sheet.getRange("AP2").setValue("");
  sheet.getRange("BH2").setValue("");

  resetDerivedSheets_(ss);
  resetEventsSystem_(ss); // NEW: Clears Events & Backend Log before loading new ones
  
  // Expand the Contact Log to exactly 23 rows (plus the END row) BEFORE loading data
  resizeContactLog_(ss, 23);
  
  resetContactLog_(ss);
  clearTypeDropdownColors_(ss); // Wipe the old colors before loading new data!
  resetMasterTable_(ss);

  // Load All Data (Row 7 Start)
  loadFakeATS_(ss);
  loadFakeParentsSquare_(ss); // RESTORED: Parent Square Fake Data
  loadFakeNYCSA_(ss);
  loadFakeClassSiteData_(ss);
  loadFakeContactLog_(ss);
  loadFakeEvents_(ss); // NEW: Builds event tables and populates backend

  fillContactLogHstackFormulas_(ss);

  if (typeof buildAllDerivedSheets === "function") {
    buildAllDerivedSheets();
  }

  applyQBlackTextFilter_();

  // Apply the alternating row colors to the newly loaded data!
  if (sheet) applyBandingToRawData_(sheet);

  // Apply the vibrant text colors to the newly generated Type Dropdowns!
  if (typeof applyTypeDropdownColors === "function") {
    applyTypeDropdownColors();
  }

  SpreadsheetApp.getUi().alert("Reset + Full Data Load (including Events & Parents Square) complete.");
}

/* =========================
RESET & HEADER FUNCTIONS
========================= */

// NEW: Helper function to delete legacy override sheets
function deleteLegacySheets_(ss) {
  const oldPhone = ss.getSheetByName("Old Phone Contacts");
  if (oldPhone) ss.deleteSheet(oldPhone);

  const oldMaster = ss.getSheetByName("Old Master Table");
  if (oldMaster) ss.deleteSheet(oldMaster);
}

function resetRawData_(ss) {
  const sheet = ss.getSheetByName("Raw Data");
  if (!sheet) throw new Error('Sheet "Raw Data" not found.');
  
  sheet.clear();

  // Remove all banding safely across the whole sheet first
  try {
    sheet.getBandings().forEach(b => b.remove());
  } catch(e) { console.error(e); }
  
  const maxCols = sheet.getMaxColumns();
  if (maxCols > 0) {
    sheet.getRange(1, 1, 1, maxCols).breakApart();
  }
}

function writeRawDataHeaders_(ss) {
  const sheet = ss.getSheetByName("Raw Data");
  const totalCols = columnLetterToNumber_("BI");

  // Force Google to build the missing columns so it doesn't crash
  let currentCols = sheet.getMaxColumns();
  if (currentCols < totalCols) {
    sheet.insertColumnsAfter(currentCols, totalCols - currentCols);
  }

  // Ensure we have enough rows to avoid out-of-bounds errors
  let maxRows = sheet.getMaxRows();
  if (maxRows < 10) {
    sheet.insertRowsAfter(maxRows, 10 - maxRows);
    maxRows = 10;
  }

  // --- Row 6: Column Headers ---
  const row6 = new Array(totalCols).fill("");
  const atsHeaders = [
    "School Year","DBN","Student ID","Last Name","First Name","Race","Gender","Age","IEP Flag","ELL Flag",
    "504 Flag","NYSAA Eligible","Housing Status","Temporary Housing Flag","Economic Disadvantage Flag","Cohort Year",
    "Grade","Official Class","Home Language","Guardian Name 1","Guardian Relationship 1","Guardian Phone Number 1",
    "Guardian Email 1","Guardian Spoken Language 1","Guardian Written Language 1","Guardian Name 2",
    "Guardian Relationship 2","Guardian Phone Number 2","Guardian Email 2","Guardian Spoken Language 2",
    "Guardian Written Language 2",""
  ];
  const psHeaders = ["Name","Students","Email","Phone","Secondary Phone","Record Created","Registered?","SIS Synced"];
  const nycsaHeaders = [
    "Location Code","School District","School Name","Class Code","Grade Level","Student ID","Student First Name",
    "Student Last Name","Guardian First Name","Guardian Last Name","Another Custodial NYCSA Account Exist",
    "Guardian Home Number","Guardian Work Number","Guardian Cell Number","Relationship","Role","Account Creation Code"
  ];

  for (let i = 0; i < atsHeaders.length; i++) row6[i] = atsHeaders[i];

  const psStart = columnLetterToNumber_("AG") - 1;
  for (let i = 1; i < psHeaders.length; i++) row6[psStart + i] = psHeaders[i];

  const nycsaStart = columnLetterToNumber_("AP") - 1;
  for (let i = 0; i < nycsaHeaders.length; i++) row6[nycsaStart + i] = nycsaHeaders[i];

  row6[columnLetterToNumber_("BH") - 1] = "Site";
  row6[columnLetterToNumber_("BI") - 1] = "Class";

  sheet.getRange(6, 1, 1, totalCols).setValues([row6]);

  // Write the default 00/00/00 dates before painting
  sheet.getRange("A1").setValue("ATS - DATE UPDATED: 00/00/00");
  sheet.getRange("AG1").setValue("Parent Square - DATE UPDATED: 00/00/00");
  sheet.getRange("AP1").setValue("NYCSA - DATE UPDATED: 00/00/00");
  sheet.getRange("BH3").setValue("MAKE OWN DATA");

  // Lock in the vibrant paint job and ENFORCE CLIPPING!
  forceVibrantColors_(sheet);

  // --- PASTE Placeholders ---
  sheet.getRange("A5").setValue("PASTE").setHorizontalAlignment("center");
  sheet.getRange("AG6").setValue("PASTE").setHorizontalAlignment("center");
  sheet.getRange("AP2").setValue("PASTE").setHorizontalAlignment("center");

  // Apply initial alternating row colors for the empty sheet
  applyBandingToRawData_(sheet);
}

/* =========================
RAW DATA COLOR HELPERS
========================= */

function forceVibrantColors_(sheet) {
  const totalCols = columnLetterToNumber_("BI");
  const maxRows = sheet.getMaxRows();
  
  // 0. ENFORCE CLIPPING & Remove Gridlines
  sheet.setHiddenGridlines(true);
  sheet.getRange(1, 1, maxRows, totalCols).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  
  // 1. Lock in Rows 1-3 Dark Blue Theme
  sheet.getRange("A1:AE3").setBackground("#073763").setFontColor("#FFFFFF").setFontWeight("bold").setFontFamily("Roboto");
  sheet.getRange("AG1:AN3").setBackground("#783f04").setFontColor("#FFFFFF").setFontWeight("bold").setFontFamily("Roboto");
  sheet.getRange("AP1:BF3").setBackground("#274e13").setFontColor("#FFFFFF").setFontWeight("bold").setFontFamily("Roboto");
  sheet.getRange("BH1:BI2").setBackground("#4c1130").setFontColor("#FFFFFF").setFontWeight("bold").setFontFamily("Roboto");
  
  // Override ONLY Row 1 back to Montserrat
  sheet.getRange("A1:BI1").setFontFamily("Montserrat");

  // Keep the title text centered (without overwriting the actual dates!)
  sheet.getRange("A1:AE1").merge().setHorizontalAlignment("center");
  sheet.getRange("AG1:AN1").merge().setHorizontalAlignment("center");
  sheet.getRange("AP1:BF1").merge().setHorizontalAlignment("center");
  sheet.getRange("BH3:BI3").merge().setHorizontalAlignment("center");

  // Make Row 4 and 5 a clean white divider
  sheet.getRange("A4:BI5").setBackground("#FFFFFF").setFontColor("#000000").setFontFamily("Roboto");

  // 2. Lock in Row 6 Headers Theme (Pastel Colors)
  sheet.getRange("A6:AE6").setBackground("#bcd9ea").setFontColor("#000000").setFontWeight("bold").setFontFamily("Roboto");
  sheet.getRange("AG6:AN6").setBackground("#ff9900").setFontColor("#000000").setFontWeight("bold").setFontFamily("Roboto");
  sheet.getRange("AP6:BF6").setBackground("#d6ecd2").setFontColor("#000000").setFontWeight("bold").setFontFamily("Roboto");
  sheet.getRange("BH6:BI6").setBackground("#eddbf4").setFontColor("#000000").setFontWeight("bold").setFontFamily("Roboto");
  
  // 3. Lock in the Yellow Paste Zones & Borders
  const pasteRanges = ["A5", "AG6", "AP2", "BH3"];
  pasteRanges.forEach(rangeA1 => {
    const r = sheet.getRange(rangeA1);
    r.setBackground("#FFF2CC"); // Light yellow
    r.setFontColor("#000000"); // Black text
    r.setFontFamily("Montserrat"); // FORCE Montserrat exception!
    r.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  });

  // 4. Heavy Section Dividers (Columns AF, AO, BG)
  const spacers = ["AF", "AO", "BG"];
  spacers.forEach(col => {
    const colRange = sheet.getRange(`${col}1:${col}${maxRows}`);
    colRange.clearContent();
    colRange.setBackground("#FFFFFF");
    colRange.setBorder(null, true, null, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });
}

function applyBandingToRawData_(sheet) {
  const maxRows = sheet.getMaxRows();
  const dataRows = maxRows - 6; // Strictly Rows 7 and down!
  if (dataRows <= 0) return;

  try {
    sheet.getBandings().forEach(b => b.remove());
  } catch(e) { console.error(e); }

  // Apply alternating colors starting at Row 7
  applyBandingSafe_(sheet.getRange(7, 1, dataRows, 31), SpreadsheetApp.BandingTheme.BLUE); // ATS
  applyBandingSafe_(sheet.getRange(7, 33, dataRows, 8), SpreadsheetApp.BandingTheme.BROWN); // RESTORED: Parent Square
  applyBandingSafe_(sheet.getRange(7, 42, dataRows, 17), SpreadsheetApp.BandingTheme.GREEN); // NYCSA
  applyBandingSafe_(sheet.getRange(7, 60, dataRows, 2), SpreadsheetApp.BandingTheme.PINK); // SITE/CLASS
}

function applyBandingSafe_(range, theme) {
  try {
    range.getBandings().forEach(b => b.remove());
    range.applyRowBanding(theme, false, false);
  } catch (e) { console.error(e); }
}

/* =========================
LOAD RAW DATA (ROW 7+)
========================= */

function loadFakeATS_(ss) {
  const sheet = ss.getSheetByName("Raw Data");
  const data = [
    ["2025-2026","75M094","123456789","SIMPSON","BART","NA","M",11,1,0,0,1,"P - Permanent",0,1,"",6,"Y41","English","HOMER SIMPSON","FATHER","1111111111","HOMERSIMPSON@gmail.com","English","English","MARGE SIMPSON","MOTHER","2222222222","MARGE.SIMPSON@gmail.com","English","English"],
    ["2025-2026","75M094","567891234","SIMPSON","LISA","NA","F",9,1,0,0,1,"P - Permanent",0,1,"",3,"Y25","English","HOMER SIMPSON","FATHER","1111111111","HOMERSIMPSON@gmail.com","English","English","MARGE SIMPSON","MOTHER","2222222222","MARGE.SIMPSON@gmail.com","English","English"],
    ["2025-2026","75M094","234567891","ADDAMS","WEDNESDAY","NA","F",7,1,1,0,1,"P - Permanent",0,1,"",2,"Y66","Spanish","MORTICIA ADDAMS","MOTHER","3333333333","MORTICIA.ADDAMS@gmail.com","Spanish","Spanish","GOMEZ ADDAMS","FATHER","3333334444","GOMEZ.ADDAMS@gmail.com","English","English"],
    ["2025-2026","75M094","345678912","INCREDIBLE","DASH","NA","M",6,1,0,0,0,"P - Permanent",0,1,"",1,"X30","English","MR. INCREDIBLE","FATHER","4444444444","MR.INCREDIBLE@gmail.com","English","English","MRS. INCREDIBLE","MOTHER","5555555555","","English","English"],
    ["2025-2026","75M094","456789123","INCREDIBLE","VIOLET","NA","F",8,1,0,0,0,"P - Permanent",0,1,"",1,"X30","English","MR. INCREDIBLE","FATHER","4444444444","MR.INCREDIBLE@gmail.com","English","English","MRS. INCREDIBLE","MOTHER","5555555555","","English","English"],
    ["2025-2026","75M094","678912345","PARKER","PETER","NA","M",10,1,0,0,0,"P - Permanent",0,1,"",4,"Y11","English","MAY PARKER","AUNT","6666666666","MAY.PARKER@gmail.com","English","English","","","","","",""],
    ["2025-2026","75M094","789123456","MORALES","MILES","NA","M",11,1,1,0,1,"P - Permanent",0,1,"",5,"Y24","Spanish","RIO MORALES","MOTHER","7777777777","RIO.MORALES@gmail.com","Spanish","Spanish","JEFFERSON MORALES","FATHER","8888888888","JEFFERSON.MORALES@gmail.com","English","English"],
    ["2025-2026","75M094","891234567","KENT","CLARK","NA","M",12,1,0,0,0,"P - Permanent",0,1,"",6,"Y50","English","MARTHA KENT","MOTHER","9999999999","MARTHA.KENT@gmail.com","English","English","JONATHAN KENT","FATHER","1010101010","JONATHAN.KENT@gmail.com","English","English"],
    ["2025-2026","75M094","912345678","WAYNE","BRUCE","NA","M",12,1,0,1,0,"P - Permanent",0,1,"",6,"Y50","English","ALFRED PENNYWORTH","GUARDIAN","1212121212","ALFRED@WAYNE.COM","English","English","","","","","",""],
    ["2025-2026","75M094","923456789","STARK","TONY","NA","M",12,0,0,0,0,"P - Permanent",0,1,"",6,"Y50","English","PEPPER POTTS","GUARDIAN","1313131313","PEPPER@STARK.COM","English","English","","","","","",""],
    ["2025-2026","75M094","934567891","BANNER","BRUCE","NA","M",11,1,0,0,0,"T - Temporary",1,1,"",5,"Y24","English","BETTY ROSS","GUARDIAN","1414141414","BETTY.ROSS@gmail.com","English","English","","","","","",""],
    ["2025-2026","75M094","945678912","PRINCE","DIANA","NA","F",10,1,0,0,0,"P - Permanent",0,1,"",4,"Y11","English","HIPPOLYTA","MOTHER","1515151515","HIPPOLYTA@AMAZON.COM","English","English","","","","","",""]
  ];
  sheet.getRange(7, 1, data.length, data[0].length).setValues(data);
}

// RESTORED: Parent Square Fake Data Loader
function loadFakeParentsSquare_(ss) {
  const sheet = ss.getSheetByName("Raw Data");
  const data = [
    ["Simpson, Homer","Bart Simpson","HOMERSIMPSON@gmail.com","(111) 111-1111","","Oct 30, 2025","Yes","Yes"],
    ["Simpson, Marge","Bart Simpson","MARGE.SIMPSON@gmail.com","(222) 222-2222","","Oct 2, 2025","Yes","Yes"],
    ["Simpson, Homer","Lisa Simpson","HOMERSIMPSON@gmail.com","(111) 111-1111","","Nov 5, 2025","Yes","Yes"],
    ["Simpson, Marge","Lisa Simpson","MARGE.SIMPSON@gmail.com","(222) 222-2222","","Nov 5, 2025","Yes","Yes"],
    ["Addams, Morticia","Wednesday Addams","MORTICIA.ADDAMS@gmail.com","(333) 333-3333","","Sep 18, 2025","No","Yes"],
    ["Addams, Gomez","Wednesday Addams","GOMEZ.ADDAMS@gmail.com","(333) 333-4444","","Sep 18, 2025","Yes","Yes"],
    ["Incredible, Mr.","Dash Incredible","MR.INCREDIBLE@gmail.com","(444) 444-4444","","Oct 12, 2025","Yes","Yes"],
    ["Incredible, Mrs.","Dash Incredible","","(555) 555-5555","","Oct 12, 2025","Yes","Yes"],
    ["Incredible, Mr.","Violet Incredible","MR.INCREDIBLE@gmail.com","(444) 444-4444","","Oct 12, 2025","Yes","Yes"],
    ["Incredible, Mrs.","Violet Incredible","","(555) 555-5555","","Oct 12, 2025","Yes","Yes"],
    ["Parker, May","Peter Parker","MAY.PARKER@gmail.com","(666) 666-6666","","Nov 1, 2025","No","Yes"],
    ["Morales, Rio","Miles Morales","RIO.MORALES@gmail.com","(777) 777-7777","","Oct 25, 2025","Yes","Yes"],
    ["Morales, Jefferson","Miles Morales","JEFFERSON.MORALES@gmail.com","(888) 888-8888","","Oct 25, 2025","No","Yes"],
    ["Kent, Martha","Clark Kent","MARTHA.KENT@gmail.com","(999) 999-9999","","Oct 10, 2025","Yes","Yes"],
    ["Kent, Jonathan","Clark Kent","JONATHAN.KENT@gmail.com","(101) 010-1010","","Oct 10, 2025","Yes","Yes"],
    ["Pennyworth, Alfred","Bruce Wayne","ALFRED@WAYNE.COM","(121) 212-1212","","Oct 15, 2025","Yes","No"],
    ["Potts, Pepper","Tony Stark","PEPPER@STARK.COM","(131) 313-1313","","Oct 20, 2025","Yes","Yes"],
    ["Ross, Betty","Bruce Banner","BETTY.ROSS@gmail.com","(141) 414-1414","","Oct 22, 2025","No","No"],
    ["Prince, Hippolyta","Diana Prince","HIPPOLYTA@AMAZON.COM","(151) 515-1515","","Oct 18, 2025","Yes","Yes"]
  ];
  sheet.getRange(7, columnLetterToNumber_("AG"), data.length, data[0].length).setValues(data);
}

function loadFakeNYCSA_(ss) {
  const sheet = ss.getSheetByName("Raw Data");
  const data = [
    ["M094",75,"P.S. M094","Y41",6,"123456789","BART","SIMPSON","HOMER","SIMPSON","No","","","1111111111","FATHER","Custodial User","B11111C2"],
    ["M094",75,"P.S. M094","Y41",6,"123456789","BART","SIMPSON","MARGE","SIMPSON","Yes","","","2222222222","MOTHER","Custodial User","B11111C3"],
    ["M094",75,"P.S. M094","Y25",3,"567891234","LISA","SIMPSON","HOMER","SIMPSON","No","","","1111111111","FATHER","Custodial User","L22222A1"],
    ["M094",75,"P.S. M094","Y25",3,"567891234","LISA","SIMPSON","MARGE","SIMPSON","Yes","","","2222222222","MOTHER","Custodial User","L22222A2"],
    ["M094",75,"P.S. M094","Y66",2,"234567891","WEDNESDAY","ADDAMS","MORTICIA","ADDAMS","No","","","3333333333","MOTHER","Custodial User","W33333B1"],
    ["M094",75,"P.S. M094","Y66",2,"234567891","WEDNESDAY","ADDAMS","GOMEZ","ADDAMS","Yes","","","3333334444","FATHER","Non-Custodial User","W33333B2"],
    ["M094",75,"P.S. M094","X30",1,"345678912","DASH","INCREDIBLE","MR.","INCREDIBLE","No","","","4444444444","FATHER","Custodial User","D44444C1"],
    ["M094",75,"P.S. M094","X30",1,"345678912","DASH","INCREDIBLE","MRS.","INCREDIBLE","Yes","","","5555555555","MOTHER","Custodial User","D44444C2"],
    ["M094",75,"P.S. M094","X30",1,"456789123","VIOLET","INCREDIBLE","MR.","INCREDIBLE","No","","","4444444444","FATHER","Custodial User","V55555D1"],
    ["M094",75,"P.S. M094","X30",1,"456789123","VIOLET","INCREDIBLE","MRS.","INCREDIBLE","Yes","","","5555555555","MOTHER","Custodial User","V55555D2"],
    ["M094",75,"P.S. M094","Y11",4,"678912345","PETER","PARKER","MAY","PARKER","No","","","6666666666","AUNT","Custodial User","P66666E1"],
    ["M094",75,"P.S. M094","Y24",5,"789123456","MILES","MORALES","RIO","MORALES","No","","","7777777777","MOTHER","Custodial User","M77777F1"],
    ["M094",75,"P.S. M094","Y24",5,"789123456","MILES","MORALES","JEFFERSON","MORALES","Yes","","","8888888888","FATHER","Non-Custodial User","M77777F2"],
    ["M094",75,"P.S. M094","Y50",6,"891234567","CLARK","KENT","MARTHA","KENT","No","","","9999999999","MOTHER","Custodial User","C88888G1"],
    ["M094",75,"P.S. M094","Y50",6,"891234567","CLARK","KENT","JONATHAN","KENT","Yes","","","1010101010","FATHER","Custodial User","C88888G2"],
    ["M094",75,"P.S. M094","Y50",6,"912345678","BRUCE","WAYNE","ALFRED","PENNYWORTH","No","","","1212121212","GUARDIAN","Custodial User","B99999H1"],
    ["M094",75,"P.S. M094","Y50",6,"923456789","TONY","STARK","PEPPER","POTTS","No","","","1313131313","GUARDIAN","Custodial User","T10101J1"],
    ["M094",75,"P.S. M094","Y24",5,"934567891","BRUCE","BANNER","BETTY","ROSS","No","","","1414141414","GUARDIAN","Non-Custodial User","B20202K1"],
    ["M094",75,"P.S. M094","Y11",4,"945678912","DIana","PRINCE","HIPPOLYTA","PRINCE","No","","","1515151515","MOTHER","Custodial User","D30303L1"]
  ];
  sheet.getRange(7, columnLetterToNumber_("AP"), data.length, data[0].length).setValues(data);
}

function loadFakeClassSiteData_(ss) {
  const sheet = ss.getSheetByName("Raw Data");
  const data = [
    ["15", "Y41: Tessa"],
    ["51", "Y25: Sam"],
    ["51", "Y66: Carl"],
    ["188", "X30: Ham"],
    ["276", "Y11: Bam"],
    ["281", "Y24: LAM"],
    ["340", "Y50:"]
  ];
  sheet.getRange(7, columnLetterToNumber_("BH"), data.length, 2).setValues(data);
}

/* =========================
CONTACT LOG (16 Column Version)
========================= */

function loadFakeContactLog_(ss) {
  const sheet = ss.getSheetByName("Contact Log");
  if (!sheet) throw new Error('Sheet "Contact Log" not found.');

  const endRow = findEndRowInColumnE_(sheet); // ⚡ SHIFTED
  if (!endRow) throw new Error('Could not find "END" in column E of "Contact Log".');

  const startRow = 2;
  const maxRows = endRow - startRow; 

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const fmtDate = (d) => Utilities.formatDate(d, Session.getScriptTimeZone(), "EEEE, M/d/yy");

  const daysAgo = (n) => {
    const d = new Date(today);
    d.setDate(d.getDate() - n);
    return fmtDate(d);
  };

  const rows = [
    [daysAgo(17),"Homer Simpson and Marge Simpson","Homer Simpson",123456789,"Bart Simpson","Phone","Call","Homer asked about start of school year",false],
    [daysAgo(15),"Morticia Addams and Gomez Addams","Morticia Addams",234567891,"Wednesday Addams","Walk In","Intake","Submitted documents for enrollment",true],
    [daysAgo(14),"Mr. Incredible and Mrs. Incredible","Mrs. Incredible",345678912,"Dash Incredible","Parent-Teacher","In School","Attended parent-teacher conference",false],
    [daysAgo(13),"Mr. Incredible and Mrs. Incredible","Mr. Incredible",456789123,"Violet Incredible","Phone","Parent Meeting","Asked about academic progress",false],
    [daysAgo(12),"Homer Simpson and Marge Simpson","Marge Simpson",123456789,"Bart Simpson","Phone","Email","Requesting teacher contact info",true],
    [daysAgo(9),"Homer Simpson and Marge Simpson","Homer Simpson",123456789,"Bart Simpson","Phone","Email","Requested teacher email again",true],
    [daysAgo(8),"Morticia Addams and Gomez Addams","Morticia Addams",234567891,"Wednesday Addams","Phone","Call","Discussed student progress",false],
    [daysAgo(7),"Rio Morales and Jefferson Morales","Rio Morales",789123456,"Miles Morales","Phone","Call","Discussed attendance concerns",true],
    [daysAgo(6),"Rio Morales and Jefferson Morales","Jefferson Morales",789123456,"Miles Morales","Phone","Email","Requested IEP documents",false],
    [daysAgo(5),"May Parker","May Parker",678912345,"Peter Parker","Phone","Call","Asked about tutoring options",false],
    [daysAgo(4),"Martha Kent and Jonathan Kent","Martha Kent",891234567,"Clark Kent","Phone","Call","Discussed college readiness",true],
    [daysAgo(3),"Martha Kent and Jonathan Kent","Jonathan Kent",891234567,"Clark Kent","Phone","Call","Confirmed attendance improvement",false],
    [daysAgo(2),"Alfred Pennyworth","Alfred Pennyworth",912345678,"Bruce Wayne","Walk In","Meeting","Discussed behavioral concerns",true],
    [daysAgo(2),"Pepper Potts","Pepper Potts",923456789,"Tony Stark","Phone","Email","Requested report card copy",false],
    [daysAgo(1),"Betty Ross","Betty Ross",934567891,"Bruce Banner","Phone","Call","Discussed housing situation",true],
    [daysAgo(1),"Hippolyta","Hippolyta",945678912,"Diana Prince","Phone","Call","Asked about extracurriculars",false],
    [daysAgo(0),"Homer Simpson and Marge Simpson","Homer Simpson",567891234,"Lisa Simpson","Phone","Call","Discussed academic performance",false],
    [daysAgo(0),"Homer Simpson and Marge Simpson","Marge Simpson",567891234,"Lisa Simpson","Phone","Email","Requested meeting with teacher",false]
  ];

  const writeRows = rows.slice(0, maxRows);
  if (!writeRows.length) return;

  // ⚡ ALL COLUMNS SHIFTED 
  writeSingleColumn_(sheet, startRow, 2, writeRows.map(r => [r[0]]));  // Col B (Date)
  writeSingleColumn_(sheet, startRow, 3, writeRows.map(r => [r[1]]));  // Col C (Guardian)
  writeSingleColumn_(sheet, startRow, 4, writeRows.map(r => [r[2]]));  // Col D (Person Spoke With)
  writeSingleColumn_(sheet, startRow, 5, writeRows.map(r => [r[3]]));  // Col E (OSIS)
  writeSingleColumn_(sheet, startRow, 6, writeRows.map(r => [r[4]]));  // Col F (Student)
  
  writeSingleColumn_(sheet, startRow, 10, writeRows.map(r => [r[5]])); // Col J (Method 1)
  writeSingleColumn_(sheet, startRow, 11, writeRows.map(r => [r[6]])); // Col K (Method 2)
  
  const fakeTypes = [
    "School Schedule", "Enrollment Documents", "School Meeting", "Academic Inquiry",
    "Contact Request", "Email Request", "Progress Review", "Attendance Meeting", 
    "Records Request", "Academic Support", "College Planning", "Attendance Progress", 
    "Behavioral Consultation", "Student Records", "Housing Situation", "Student Activities", 
    "Academic Advising", "Meeting Request"
  ];
  writeSingleColumn_(sheet, startRow, 12, fakeTypes.slice(0, writeRows.length).map(t => [t])); // Col L (Type)

  writeSingleColumn_(sheet, startRow, 13, writeRows.map(r => [r[7]])); // Col M (Notes)

  const numberValues = new Array(writeRows.length).fill([1]);
  sheet.getRange(startRow, 7, numberValues.length, 1).setValues(numberValues); // Col G (# interactions)

  const sisStatus = [
    "No","No","No","Yes","No","Yes - Important","No",
    "No","Yes","Newer Interaction","Yes - Important",
    "No","Yes","No","No","No","No", "No"
  ].slice(0, writeRows.length).map(v => [v]);
  sheet.getRange(startRow, 14, sisStatus.length, 1).setValues(sisStatus); // Col N (SIS)

  if (maxRows >= 7) sheet.getRange("O8").setValue("Called principal"); // Col O
  if (maxRows >= 18) sheet.getRange("O19").setValue("Teacher will reach out"); // Col O

  const checkboxValues = writeRows.map(r => [r[8]]);
  const sRange = sheet.getRange(startRow, 16, checkboxValues.length, 1); // Col P (16)
  sRange.clearContent().insertCheckboxes().setValues(checkboxValues).setFontColor("#000000");

  fillContactLogHstackFormulas_(ss);
}

/* =========================
NEW: EVENTS DATA BUILDER
========================= */

function resetEventsSystem_(ss) {
  // 1. Wipe visual Events sheet by deleting all columns and rebuilding A-Z
  let eventsSheet = ss.getSheetByName("Events");
  if (eventsSheet) {
    const maxCols = eventsSheet.getMaxColumns();
    if (maxCols > 0) {
      // First clear everything
      eventsSheet.clear();
      
      // Insert 26 fresh, unformatted columns
      eventsSheet.insertColumnsAfter(maxCols, 26);
      
      // Delete all the old columns that might have lingering data validations or formatting
      eventsSheet.deleteColumns(1, maxCols);
      
      // FIX: Strip checkboxes AND clear the underlying "FALSE" text they leave behind
      eventsSheet.getRange(1, 1, eventsSheet.getMaxRows(), 26).clearDataValidations().clearContent();
      
      // Set the width of the 26 new columns to the default 100 pixels
      for (let i = 1; i <= 26; i++) {
        eventsSheet.setColumnWidth(i, 100);
      }
    }
  }

  // 2. Safely clear Backend Data but keep headers
  let backendSheet = ss.getSheetByName("Backend_Event_Log");
  if (backendSheet && backendSheet.getLastRow() > 1) {
    backendSheet.getRange(2, 1, backendSheet.getLastRow() - 1, backendSheet.getLastColumn()).clearContent();
  }
}

function loadFakeEvents_(ss) {
  // Ensure the table builder function exists before trying to run it
  if (typeof buildEventTable !== "function") return;

  const today = new Date();
  
  // Fake Event 1: Graduation Workshop (10 Days Ago)
  let d1 = new Date(today);
  d1.setDate(d1.getDate() - 7);
  const dateStr1 = Utilities.formatDate(d1, Session.getScriptTimeZone(), "MM/dd/yyyy");
  buildEventTable("Graduation Workshop", "Auditorium", dateStr1,"Workshop");

  // Fake Event 2: Spring Festival (15 Days Ago)
  let d2 = new Date(today);
  d2.setDate(d2.getDate() - 0);
  const dateStr2 = Utilities.formatDate(d2, Session.getScriptTimeZone(), "MM/dd/yyyy");
  buildEventTable("Spring Festival", "School Gym", dateStr2, "Other");

  const eventsSheet = ss.getSheetByName("Events");
  const backendSheet = ss.getSheetByName("Backend_Event_Log");
  
  // --- INJECT SPRING FESTIVAL DATA (Cols 1-6) ---
  const springData = [
    [456789123, "Mr. Incredible and Mrs. Incredible", "Mrs. Incredible", "Violet Incredible", 2, "Bought tickets early"],
    [678912345, "May Parker", "May Parker", "Peter Parker", 1, ""],
    [923456789, "Pepper Potts", "Pepper Potts", "Tony Stark", 3, "Donated to bake sale"]
  ];
  eventsSheet.getRange(3, 1, springData.length, 6).setValues(springData);

  springData.forEach(row => {
    backendSheet.appendRow([d2, "Spring Festival", "Other", row[0], row[1], row[2], row[3], row[4], row[5]]);
  });

  // --- INJECT GRADUATION WORKSHOP DATA (Cols 8-13) ---
  const gradData = [
    [123456789, "Homer Simpson and Marge Simpson", "Homer Simpson", "Bart Simpson", 1, "Wants to sit in on class"],
    [234567891, "Morticia Addams and Gomez Addams", "Morticia Addams", "Wednesday Addams", 1, ""],
    [789123456, "Rio Morales and Jefferson Morales", "Jefferson Morales", "Miles Morales", 1, "Asked about diploma tracks"]
  ];
  eventsSheet.getRange(3, 8, gradData.length, 6).setValues(gradData);

  gradData.forEach(row => {
    backendSheet.appendRow([d1, "Graduation Workshop", "Workshop", row[0], row[1], row[2], row[3], row[4], row[5]]);
  });
}

// --------------------------------------------------------------------------------------

function resetContactLog_(ss) {
  const sheet = ss.getSheetByName("Contact Log");
  if (!sheet) return;
  const endRow = findEndRowInColumnE_(sheet); // ⚡ SHIFTED
  if (!endRow || endRow <= 2) return;

  const numRows = endRow - 2;

  // 1. Clear text in Col B through O (14 columns)
  sheet.getRange(2, 2, numRows, 14).clearContent();
  
  // 2. NUKE old ghost checkboxes from Column O (15) just in case they are stuck!
  sheet.getRange(2, 15, numRows, 1).clearDataValidations();
  
  // 3. Insert fresh checkboxes into Column P (16) for standard log rows and force UNCHECK
  const sRange = sheet.getRange(2, 16, numRows, 1);
  sRange.clearContent().insertCheckboxes().uncheck().setFontColor("#000000");

  // 3b. FORCE the Master Filter Checkbox on the exact "END" row to be checked (TRUE)
  sheet.getRange(endRow, 15).insertCheckboxes().check().setFontColor("#000000"); // ⚡ SHIFTED Col 15

  // 4. Forcefully redraw the correct divider lines so the table looks perfect
  formatContactLogRows_(sheet, 2, numRows);

  // NEW: 5. Delete rows 9 through 19 below the "END" row to wipe out old UI elements/overrides
  const startDeleteRow = endRow + 9;
  const maxRows = sheet.getMaxRows();
  if (maxRows >= startDeleteRow) {
    // Determine how many rows we can actually delete (up to 11 rows to cover 9-19)
    const rowsToDelete = Math.min(11, maxRows - startDeleteRow + 1);
    if (rowsToDelete > 0) {
      sheet.deleteRows(startDeleteRow, rowsToDelete);
    }
  }
}

function resetMasterTable_(ss) {
  const sheet = ss.getSheetByName("Master Table");
  if (sheet && sheet.getMaxRows() >= 5) {
    sheet.getRange(5, 1, sheet.getMaxRows() - 4, 9).clearContent();
  }
}

function fillContactLogHstackFormulas_(ss) {
  const sheet = ss.getSheetByName("Contact Log");
  if (!sheet) return;

  const endRow = findEndRowInColumnE_(sheet); // ⚡ SHIFTED
  if (!endRow || endRow <= 2) return;

  const startRow = 2;
  const numRows = endRow - startRow; 

  // 1. Generate row-specific formulas for Column A
  const colAFormulas = [];
  for (let i = 0; i < numRows; i++) {
    const rowNum = startRow + i;
    colAFormulas.push([`=IF(B${rowNum}="", "", B${rowNum} - WEEKDAY(B${rowNum}, 1) + 1)`]);
  }
  
  sheet.getRange(startRow, 1, numRows, 1).setFormulas(colAFormulas);
  // (We removed fC, fF, fH, fK, fL because the Ghost Typist handles them!)
}

// findEndRowInColumnE_ is defined once in QuickUpdates.gs (identical body) and
// shared across the project, so the duplicate that used to live here was removed
// to avoid load-order ambiguity.

function columnLetterToNumber_(letter) {
  let col = 0, s = letter.toUpperCase();
  for (let i = 0; i < s.length; i++) col = col * 26 + (s.charCodeAt(i) - 64);
  return col;
}

function writeSingleColumn_(sheet, startRow, col, values) {
  if (values.length) sheet.getRange(startRow, col, values.length, 1).setValues(values);
}

function applyQBlackTextFilter_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Contact Log");
  const endRow = findEndRowInColumnE_(sheet); // ⚡ SHIFTED
  if (!sheet || !endRow) return;
  if (sheet.getFilter()) sheet.getFilter().remove();

  const criteria = SpreadsheetApp.newFilterCriteria()
    .setVisibleForegroundColor(
      SpreadsheetApp.newColor().setRgbColor("#000000").build()
    )
    .build();

  // Filter now spans 16 columns, and looks at Column 16 (P) for the checkboxes
  sheet.getRange(1, 1, endRow - 1, 16).createFilter().setColumnFilterCriteria(16, criteria);
}

function resizeContactLog_(ss, desiredRows) {
  const sheet = ss.getSheetByName("Contact Log");
  if (!sheet) return;

  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }

  const endRow = findEndRowInColumnE_(sheet); // ⚡ SHIFTED
  if (!endRow) return;

  const currentDataRows = endRow - 2; 

  if (currentDataRows > desiredRows) {
    const rowsToDelete = currentDataRows - desiredRows;
    sheet.deleteRows(2 + desiredRows, rowsToDelete);
  } else if (currentDataRows < desiredRows) {
    sheet.insertRowsAfter(2, desiredRows - currentDataRows);
    const maxCols = sheet.getMaxColumns();
    const templateRange = sheet.getRange(2, 1, 1, maxCols);
    const targetRange = sheet.getRange(3, 1, desiredRows - 1, maxCols);
    templateRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  }

  const newEndRow = findEndRowInColumnE_(sheet); // ⚡ SHIFTED
  if (newEndRow) {
    const maxRows = sheet.getMaxRows();
    if (maxRows > newEndRow) {
       // Just wipe checkboxes from P (16) below the table.
       sheet.getRange(newEndRow + 1, 16, maxRows - newEndRow, 1).clearDataValidations().clearContent();
    }
  }
}

function clearTypeDropdownColors_(ss) {
  const contactLog = ss.getSheetByName('Contact Log');
  if (!contactLog) return;
  
  let existingRules = contactLog.getConditionalFormatRules();
  
  // Keep everything EXCEPT the rules strictly on Column L (12)
  let rulesToKeep = existingRules.filter(rule => {
    const ranges = rule.getRanges();
    if (ranges.length === 0) return true;
    
    const firstCol = ranges[0].getColumn();
    const lastCol = ranges[0].getLastColumn();
    
    // ⚡ SHIFTED to 12
    if (firstCol === 12 && lastCol === 12) {
      return false; // Destroy the dropdown color rule
    }
    return true; 
  });
  
  contactLog.setConditionalFormatRules(rulesToKeep);
}

function resetDerivedSheets_(ss) {
  ["Phone Contacts", "Combined Contact Tracking", "Parents Divided","New/Edit Student","Students Override Report"].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet && sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }
  });

  const dirSheet = ss.getSheetByName("Directory");
  if (dirSheet && dirSheet.getMaxRows() > 1) {
    dirSheet.getRange(2, 1, dirSheet.getMaxRows() - 1, dirSheet.getMaxColumns()).clear();
  }
}

function resetContactLogFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Contact Log");
  if (!sheet) return;

  const maxRows = sheet.getMaxRows();

  // ⚡ ALL RANGES SHIFTED 
  const rangeB2B = sheet.getRange("B2:B" + maxRows);
  const rangeA2B = sheet.getRange("A2:B" + maxRows);
  const rangeC2P = sheet.getRange("C2:P" + maxRows);  // Was D2:S
  const rangeC2O = sheet.getRange("C2:O" + maxRows);  // Was D2:R
  const rangeG2G = sheet.getRange("G2:G" + maxRows);  // Was J2:J
  const rangeN2N = sheet.getRange("N2:N" + maxRows);  // Was Q2:Q
  const rangeP2P = sheet.getRange("P2:P" + maxRows);  // Was S2:S

  const rules = [];

   // 1. Last Week (A2:B)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$B2 = TODAY()")
    .setBackground("#d9ead3")
    .setRanges([rangeB2B]).build());
 
  // 1. Last Week (A2:B)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=AND($A2 >= TODAY() - WEEKDAY(TODAY(), 1) - 6, $A2 <= TODAY() - WEEKDAY(TODAY(), 1))")
    .setBackground("#d0e0e3")
    .setRanges([rangeA2B]).build());

  // 2. This Week (A2:B)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=AND($A2 >= TODAY() - WEEKDAY(TODAY(), 1) + 1, $A2 <= TODAY() - WEEKDAY(TODAY(), 1) + 7)")
    .setBackground("#fce5cd")
    .setRanges([rangeA2B]).build());

  // 3. Workshop (C2:O) -> ⚡ Uses Col J (Method 1)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$J2="Workshop"')
    .setBackground("#b6d7a8")
    .setRanges([rangeC2O]).build());

  // 4. Missing Interactions (G2:G) -> ⚡ Checks Col G
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($B2<>"", $G2="")')
    .setBackground("#990000")
    .setRanges([rangeG2G]).build());

  // 5. Yes - Important (N2:N) -> ⚡ Checks Col N
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("Yes - Important")
    .setBold(true)
    .setRanges([rangeN2N]).build());

 // 6. Grey Checkboxes (P2:P) -> ⚡ Regex checks Col N
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=NOT(OR(AND($A2 >= TODAY() - 7, $A2 <= TODAY()), $B2 = "", REGEXMATCH($N2, "(?i)\\bYes\\b(\\s*-\\s*Important)?")))')
    .setFontColor("#999999")
    .setRanges([rangeP2P]).build());

  sheet.setConditionalFormatRules(rules);

  if (typeof applyTypeDropdownColors === "function") {
    applyTypeDropdownColors();
  }
}
