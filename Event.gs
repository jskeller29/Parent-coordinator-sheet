// ==============================================================
// 1. ONE-BOX EVENT FORM (Custom UI with Calendar Picker)
// ==============================================================
function createNewEventTable() {
  const html = `
    <div style="font-family: Arial, sans-serif; padding: 10px;">
      <label><b>Event Name</b> <span style="color:red;">*</span> (e.g., Spring Festival):</label><br>
      <input type="text" id="name" style="width: 100%; margin-bottom: 15px; padding: 5px;" autofocus><br>
      
      <label><b>Event Type / Category</b> <span style="color:red;">*</span> (e.g., Workshop):</label>
      <button type="button" onclick="fillPTC()" style="margin-left: 5px; font-size: 11px; padding: 3px 6px; background: #e0e0e0; border: 1px solid #ccc; border-radius: 4px; cursor: pointer;">PTC Quick Fill</button><br>
      <input type="text" id="type" style="width: 100%; margin-bottom: 15px; padding: 5px;"><br>

      <label><b>Location</b> (e.g., Gym, Zoom):</label><br>
      <input type="text" id="loc" style="width: 100%; margin-bottom: 15px; padding: 5px;"><br>
      
      <label><b>Date</b> <span style="color:red;">*</span>:</label><br>
      <input type="date" id="date" style="width: 100%; margin-bottom: 20px; padding: 5px; font-family: Arial, sans-serif;"><br>
      
      <button id="createBtn" onclick="submit()" style="padding: 8px 15px; background: #0f9d58; color: white; border: none; border-radius: 4px; cursor: pointer; font-weight: bold;">Create Table</button>
      <button onclick="google.script.host.close()" style="padding: 8px 15px; background: #ddd; border: none; border-radius: 4px; cursor: pointer; margin-left: 10px;">Cancel</button>
      
      <script>
        function fillPTC() {
          document.getElementById('type').value = "Parent Teacher Conference";
        }

        function submit() {
          const name = document.getElementById('name').value.trim();
          const type = document.getElementById('type').value.trim();
          const loc = document.getElementById('loc').value.trim();
          const rawDate = document.getElementById('date').value; // Comes out as YYYY-MM-DD
          
          if(!name) {
            alert("⚠️ Event Name is required!");
            return;
          }
          if(!type) {
            alert("⚠️ Event Type is required!");
            return;
          }
          if(!rawDate) {
            alert("⚠️ Date is required! Please select a date from the calendar.");
            return;
          }
          
          // Prevent double-clicking by instantly disabling the button
          const btn = document.getElementById('createBtn');
          btn.disabled = true;
          btn.innerText = 'Creating...';
          btn.style.background = '#888';
          btn.style.cursor = 'not-allowed';
          
          // Reformat the calendar output from YYYY-MM-DD to MM/DD/YYYY
          const parts = rawDate.split('-');
          const formattedDate = parts[1] + '/' + parts[2] + '/' + parts[0];
          
          google.script.run.withSuccessHandler(google.script.host.close).buildEventTable(name, loc, formattedDate, type);
        }
      </script>
    </div>
  `;
  const ui = HtmlService.createHtmlOutput(html).setWidth(300).setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(ui, '🚀 Create New Event');
}

// ==============================================================
// 2. THE TABLE BUILDER & DROPDOWN GENERATOR
// ==============================================================
function buildEventTable(eventName, eventLoc, eventDate, eventType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Events");
  
  if (!sheet) {
    sheet = ss.insertSheet("Events");
  }
  
  sheet.activate(); 

  // Insert 7 columns on the far left (6 data + 1 spacer)
  sheet.insertColumns(1, 7);
  
  // Clear any residual formatting/validation from neighboring columns
  sheet.getRange(1, 1, sheet.getMaxRows(), 7).clearDataValidations().setBackground(null);

  // Set Row 1 Layout: Merged Title (Cols 1, 2), Category (Col 3), Location (Col 4), Date (Col 5), Delete Box (Col 6)
  sheet.getRange(1, 1, 1, 2).merge().setValue(eventName).setFontWeight("bold").setBackground("#d9ead3").setHorizontalAlignment("center");
  sheet.getRange(1, 3).setValue(eventType).setFontWeight("bold").setBackground("#DAA06D").setHorizontalAlignment("center"); // NEW: Event Type
  sheet.getRange(1, 4).setValue(eventLoc).setFontWeight("bold").setBackground("#fff2cc").setHorizontalAlignment("center");
  sheet.getRange(1, 5).setValue(eventDate).setFontWeight("bold").setBackground("#c9daf8").setHorizontalAlignment("center");
  sheet.getRange(1, 6).insertCheckboxes().setBackground("#ea9999");
  
  // Set Row 2 Headers
  const headers = ["OSIS", "Guardian(s)", "Parent Attendee", "Student Name", "# of Indv", "Notes"];
  sheet.getRange("A2:F2").setValues([headers]).setFontWeight("bold").setBackground("#efefef");

  // Format Column Widths
  sheet.setColumnWidth(1, 100); // OSIS
  sheet.setColumnWidth(2, 150); // Guardian(s)
  sheet.setColumnWidth(3, 150); // Parent Attendee
  sheet.setColumnWidth(4, 150); // Student
  sheet.setColumnWidth(5, 80);  // # of Indv
  sheet.setColumnWidth(6, 200); // Notes
  sheet.setColumnWidth(7, 30);  // Spacer

  // Apply Borders
  sheet.getRange("A1:F").setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  
  // =======================================================
  // STRICT DROPDOWN MAPPING
  // =======================================================
  const masterSheet = ss.getSheetByName("Master Table"); 
  if (masterSheet) {
    const osisRule = SpreadsheetApp.newDataValidation().requireValueInRange(masterSheet.getRange("A5:A"), true).build();
    sheet.getRange("A3:A").setDataValidation(osisRule);

    const guardianRule = SpreadsheetApp.newDataValidation().requireValueInRange(masterSheet.getRange("C5:C"), true).build();
    sheet.getRange("B3:B").setDataValidation(guardianRule);
    
    // Parent Attendee is manual entry, no strict validation needed.

    const studentRule = SpreadsheetApp.newDataValidation().requireValueInRange(masterSheet.getRange("B5:B"), true).build();
    sheet.getRange("D3:D").setDataValidation(studentRule);
  }

  // Ensure Backend exists and has updated headers
  let backendSheet = ss.getSheetByName("Backend_Event_Log");
  if (!backendSheet) {
    backendSheet = ss.insertSheet("Backend_Event_Log");
    backendSheet.hideSheet();
    backendSheet.getRange("A1:H1").setValues([["Date", "Event Name", "Category", "OSIS", "Guardian(s)", "Parent Attendee", "# of Indv", "Notes"]]).setFontWeight("bold");
  } else {
    // Migration: Update existing backend sheet if missing new columns
    const currentHeaders = backendSheet.getRange("A1:H1").getValues()[0];
    if (currentHeaders[5] !== "Parent Attendee") {
       backendSheet.insertColumnAfter(4);
       backendSheet.getRange("E1").setValue("Guardian(s)").setFontWeight("bold");
       backendSheet.getRange("F1").setValue("Parent Attendee").setFontWeight("bold");
    }
  }
}

// ==============================================================
// 3. EVENTS SHEET WATCHER: FULL RECONCILIATION SWEEP & SORT
// ==============================================================
function watchEventsSheet_(e) {
  if (!e || !e.range) return;

  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== "Events") return; 

  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();

  const maxRows = sheet.getLastRow();
  const maxCols = sheet.getLastColumn();
  if (maxRows < 3 || maxCols < 6) return;

  // Tables are 7 columns wide (6 data + 1 spacer)
  const tableStartCol = Math.floor((col - 1) / 7) * 7 + 1;
  const relativeCol = col - tableStartCol; // 0=OSIS, 1=Guardian, 2=Attendee, 3=Student, 4=#Indv, 5=Notes/Box
  
  // 1. GHOST TYPIST LOGIC (Runs instantly before anything else)
  if (row >= 3) {
    const masterSheet = e.source.getSheetByName("Master Table");
    if (masterSheet) {
      
      // ⚡ CASE A: VERTICAL PASTE (Copying & Pasting a column of OSIS numbers)
      if (range.getNumRows() > 1 && range.getNumColumns() === 1 && relativeCol === 0) {
        const pastedValues = range.getDisplayValues();
        const masterData = masterSheet.getDataRange().getValues();
        const updateData = [];
        let hasUpdates = false;

        // Fetch existing data for these rows to preserve manual entries (like attendee/notes)
        const existingData = sheet.getRange(row, tableStartCol, pastedValues.length, 6).getValues();

        for (let i = 0; i < pastedValues.length; i++) {
          let targetOsis = String(pastedValues[i][0]).trim();
          
          if (!targetOsis) {
             updateData.push([existingData[i][1], existingData[i][2], existingData[i][3], existingData[i][4]]); 
             continue;
          }
          
          let foundGuardian = "";
          let foundStudent = "";
          
          // Fast memory sweep
          for(let r = 4; r < masterData.length; r++) { 
            if(String(masterData[r][0]).trim() === targetOsis) {
              foundStudent = masterData[r][1];
              foundGuardian = masterData[r][2];
              break;
            }
          }
          
          updateData.push([
            foundGuardian || existingData[i][1], 
            existingData[i][2], 
            foundStudent || existingData[i][3], 
            (String(existingData[i][4]).trim() === "") ? 1 : existingData[i][4]
          ]);
          hasUpdates = true;
        }
        
        if (hasUpdates) {
           sheet.getRange(row, tableStartCol + 1, updateData.length, 4).setValues(updateData);
        }
      }
      
      // ⚡ CASE B: SINGLE CELL EDIT OR COMMA-SEPARATED PASTE
      else if (range.getNumRows() === 1 && range.getNumColumns() === 1) {
        const editedValue = String(e.value || range.getDisplayValue() || "").trim();
        
        if ((relativeCol === 0 || relativeCol === 1 || relativeCol === 3) && editedValue && editedValue !== "undefined") {
          
          // ⚡ NEW: MULTI-OSIS BATCH PASTE LOGIC (Comma or Line-break separated in ONE cell)
          if (relativeCol === 0 && (editedValue.includes(",") || editedValue.includes("\n"))) {
             // Split by comma or newline, remove whitespace, and filter out empties
             const osisArray = editedValue.split(/[,\n]+/).map(o => o.trim()).filter(Boolean);
             
             if (osisArray.length > 0) {
               const masterData = masterSheet.getDataRange().getValues();
               const numNewRows = osisArray.length - 1;
               
               if (numNewRows > 0) {
                 // Safely shift ONLY these 6 columns down to make room (Protects side-by-side tables!)
                 sheet.getRange(row + 1, tableStartCol, numNewRows, 6).insertCells(SpreadsheetApp.Dimension.ROWS);
                 
                 // Copy Data Validation and Formatting into the newly created gap
                 const templateRange = sheet.getRange(row, tableStartCol, 1, 6);
                 templateRange.copyTo(sheet.getRange(row + 1, tableStartCol, numNewRows, 6), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
                 templateRange.copyTo(sheet.getRange(row + 1, tableStartCol, numNewRows, 6), SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
               }
               
               const updateData = [];
               for (let i = 0; i < osisArray.length; i++) {
                 let targetOsis = osisArray[i];
                 let foundGuardian = "";
                 let foundStudent = "";
                 
                 // Fast memory sweep of the Master Table
                 for(let r = 4; r < masterData.length; r++) { 
                   if(String(masterData[r][0]).trim() === targetOsis) {
                     foundStudent = masterData[r][1];
                     foundGuardian = masterData[r][2];
                     break;
                   }
                 }
                 
                 // Preserve any manual notes typed into the original row, but keep subsequent rows clean
                 if (i === 0) {
                   const existingRowData = sheet.getRange(row, tableStartCol, 1, 6).getValues()[0];
                   updateData.push([
                     targetOsis, 
                     foundGuardian || existingRowData[1], 
                     existingRowData[2], 
                     foundStudent || existingRowData[3], 
                     (String(existingRowData[4]).trim() === "") ? 1 : existingRowData[4], 
                     existingRowData[5]
                   ]);
                 } else {
                   updateData.push([targetOsis, foundGuardian, "", foundStudent, 1, ""]);
                 }
               }
               
               // Blast the array back onto the spreadsheet instantly
               sheet.getRange(row, tableStartCol, osisArray.length, 6).setValues(updateData);
             }
          } 
          // ⚡ ORIGINAL LOGIC: Single typed entry
          else {
            let searchRange;
            if (relativeCol === 0) searchRange = masterSheet.getRange("A5:A");      
            else if (relativeCol === 1) searchRange = masterSheet.getRange("C5:C"); 
            else if (relativeCol === 3) searchRange = masterSheet.getRange("B5:B"); 

            const found = searchRange.createTextFinder(editedValue).matchEntireCell(true).findNext();
            if (found) {
              const foundRow = found.getRow();
              const correctOsis = masterSheet.getRange(foundRow, 1).getValue();
              const correctStudent = masterSheet.getRange(foundRow, 2).getValue();
              const correctGuardian = masterSheet.getRange(foundRow, 3).getValue();

              if (relativeCol !== 0) sheet.getRange(row, tableStartCol).setValue(correctOsis);
              if (relativeCol !== 1) sheet.getRange(row, tableStartCol + 1).setValue(correctGuardian);
              if (relativeCol !== 3) sheet.getRange(row, tableStartCol + 3).setValue(correctStudent);
              
              // Checks strictly for a completely empty string before forcing a 1
              if (String(sheet.getRange(row, tableStartCol + 4).getValue()).trim() === "") {
                sheet.getRange(row, tableStartCol + 4).setValue(1); 
              }
            }
          }
        }
      }
    }
  }

  SpreadsheetApp.flush(); 

  const eventName = String(sheet.getRange(1, tableStartCol).getValue()).trim();
  if (!eventName) return; 
  const category = String(sheet.getRange(1, tableStartCol + 2).getValue()).trim() || "Other";
  const dateString = sheet.getRange(1, tableStartCol + 4).getValue();
  let eventDate = new Date(dateString);
  if (isNaN(eventDate)) eventDate = new Date(); 
  const fDateString = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

  // =======================================================
  // 2. DELETE EVENT LOGIC (Red Checkbox)
  // =======================================================
  if (row === 1 && relativeCol === 5) {
    if (String(e.value).toUpperCase() === "TRUE") {
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert("⚠️ Delete Event", `Are you sure you want to completely delete the event "${eventName}"?\n\nThis will permanently delete all logs associated with it from the backend.`, ui.ButtonSet.YES_NO);
      
      if (response === ui.Button.YES) {
        sheet.deleteColumns(tableStartCol, 7);
        SpreadsheetApp.flush();
        
        const backendSheet = e.source.getSheetByName("Backend_Event_Log");
        if (backendSheet && typeof buildBackendEventLogData_ === "function") {
           const freshData = buildBackendEventLogData_(e.source);
           backendSheet.clearContents();
           if (freshData && freshData.length > 0) {
              backendSheet.getRange(1, 1, freshData.length, freshData[0].length).setValues(freshData);
           }
        }
        ui.alert("✅ Event Removed", "The event has been deleted. To push this deletion to the Contact Tracker, click 'Build Sheets Only' in your App Menu.", ui.ButtonSet.OK);
        return; 
      } else {
        range.uncheck(); 
        return; 
      }
    } else {
      return; 
    }
  }

  // =======================================================
  // 3. RENAME / DATE CHANGE INTERCEPTOR 
  // =======================================================
  if (row === 1) {
      const oldValue = String(e.oldValue || "").trim();
      if (oldValue && oldValue !== "undefined") {
          e.source.toast("Rebuilding event data...", "⚙️ Syncing", 3);
          const backendSheet = e.source.getSheetByName("Backend_Event_Log");
          if (backendSheet && typeof buildBackendEventLogData_ === "function") {
              const freshData = buildBackendEventLogData_(e.source);
              backendSheet.clearContents();
              if (freshData && freshData.length > 0) {
                 backendSheet.getRange(1, 1, freshData.length, freshData[0].length).setValues(freshData);
              }
              let affected = new Set();
              for(let r = 1; r < freshData.length; r++) {
                 if (String(freshData[r][1]).trim() === eventName) affected.add(String(freshData[r][3]).trim());
              }
              SpreadsheetApp.flush();
              if (typeof updateSingleOsisDelta_ === "function") {
                  affected.forEach(osis => { if (osis) updateSingleOsisDelta_(osis); });
              }
          }
          return; 
      }
  }

  // =======================================================
  // 4. THE ABSOLUTE RECONCILIATION SWEEP
  // =======================================================
  const allTableData = sheet.getRange(1, 1, maxRows, maxCols).getValues();
  const frontendRecords = new Map(); 
  const rowsToClear = []; 

  for (let c = 0; c < maxCols; c += 7) {
      const tName = String(allTableData[0][c] || "").trim();
      if (!tName) continue;

      const tDateStr = allTableData[0][c + 4];
      let tDate = new Date(tDateStr);
      if (isNaN(tDate)) tDate = new Date();
      const tDateString = Utilities.formatDate(tDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

      if (tName === eventName && tDateString === fDateString) {
          for (let r = 2; r < maxRows; r++) { 
              let rOsisRaw = String(allTableData[r][c] || "").trim();
              if (!rOsisRaw) continue; 

              let rGuardian = String(allTableData[r][c + 1] || "").trim();
              let rAttendee = String(allTableData[r][c + 2] || "").trim();
              let rStudentRaw = String(allTableData[r][c + 3] || "").trim(); 
              
              // ⚡ THE ZERO FIX: Evaluates the literal raw cell value!
              let rawNumCell = allTableData[r][c + 4];
              let rNum = 1; 
              if (String(rawNumCell).trim() !== "") {
                  rNum = Number(rawNumCell);
              }

              let rNotes = String(allTableData[r][c + 5] || "").trim();

              const osisArray = rOsisRaw.split(',').map(o => o.trim()).filter(o => o !== "");
              const nameArray = rStudentRaw.split(',').map(n => n.trim()).filter(n => n !== "");

              for (let j = 0; j < osisArray.length; j++) {
                  let singleOsis = osisArray[j];
                  let singleName = nameArray[j] !== undefined ? nameArray[j] : rStudentRaw;
                  let assignedNum = (j === 0) ? rNum : ""; 

                  if (frontendRecords.has(singleOsis)) {
                      rowsToClear.push({ row: r + 1, col: c + 1 }); 
                  } else {
                      frontendRecords.set(singleOsis, {
                         guardian: rGuardian,
                         attendee: rAttendee,
                         student: singleName,
                         numIndv: assignedNum,
                         notes: rNotes,
                         synced: false 
                      });
                  }
              }
          }
      }
  }

  if (rowsToClear.length > 0) {
      SpreadsheetApp.getUi().alert("⚠️ Duplicate Removed", "This family is already logged for this event. The duplicate has been removed.", SpreadsheetApp.getUi().ButtonSet.OK);
      rowsToClear.forEach(target => sheet.getRange(target.row, target.col, 1, 6).clearContent());
  }

  const backendSheet = e.source.getSheetByName("Backend_Event_Log");
  if (!backendSheet) return;
  const backendData = backendSheet.getDataRange().getValues();
  const affectedOsisNumbers = new Set(); 

  for (let i = backendData.length - 1; i > 0; i--) {
      let bEventName = String(backendData[i][1]).trim();
      let bDate = backendData[i][0];
      
      let bDateString = "";
      if (bDate instanceof Date) {
        bDateString = Utilities.formatDate(bDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else if (bDate) {
        let parsedBDate = new Date(bDate);
        if (!isNaN(parsedBDate)) bDateString = Utilities.formatDate(parsedBDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }

      if (bEventName === eventName && bDateString === fDateString) {
         let bOsis = String(backendData[i][3]).trim();

         if (frontendRecords.has(bOsis)) {
           let fData = frontendRecords.get(bOsis);
           
           if (backendData[i][4] != fData.student || backendData[i][5] != fData.guardian || backendData[i][6] != fData.attendee || backendData[i][7] != fData.numIndv || backendData[i][8] != fData.notes) {
              backendSheet.getRange(i + 1, 5, 1, 5).setValues([[fData.student, fData.guardian, fData.attendee, fData.numIndv, fData.notes]]);
              affectedOsisNumbers.add(bOsis);
           }
           fData.synced = true; 
         } else {
           backendSheet.deleteRow(i + 1);
           affectedOsisNumbers.add(bOsis);
         }
      }
  }

  for (let [fOsis, fData] of frontendRecords.entries()) {
     if (!fData.synced) {
        backendSheet.appendRow([eventDate, eventName, category, fOsis, fData.student, fData.guardian, fData.attendee, fData.numIndv, fData.notes]);
        affectedOsisNumbers.add(fOsis);
     }
  }
  
  const lastRow = backendSheet.getLastRow();
  if (lastRow > 1) {
    const rangeToSort = backendSheet.getRange(2, 1, lastRow - 1, backendSheet.getLastColumn());
    rangeToSort.sort([{column: 1, ascending: false}, {column: 2, ascending: true}]);
  }

  SpreadsheetApp.flush(); 
  if (typeof updateSingleOsisDelta_ === "function") {
    affectedOsisNumbers.forEach(osis => {
      if (osis) updateSingleOsisDelta_(osis);
    });
  }
}
