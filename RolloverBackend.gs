// ======================================================================
// FILE: RolloverBackend.gs
// PURPOSE: Handles End-of-Year archiving by freezing triggers, renaming, and exporting PDF
// ======================================================================

/**
 * Opens the Rollover Wizard UI.
 */
function openRolloverWizard() {
  const html = HtmlService.createHtmlOutputFromFile('RolloverWizard')
      .setTitle('End of Year Rollover')
      .setWidth(450)
      .setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, 'End of Year Rollover');
}

/**
 * Backend function called by the Wizard.
 * 1. Deletes sync triggers.
 * 2. Renames the current sheet to "Retired on [Date]".
 * 3. Copies tabs to a temp sheet.
 * 4. Unhides all filtered rows.
 * 5. Crops Contact Log 1 row above "END" bar.
 * 6. Exports temp sheet to PDF (with native fallback).
 */
function processRolloverBackend() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const driveFile = DriveApp.getFileById(ss.getId());
    
    // Find the folder where the current spreadsheet lives so we save the PDF in the same place
    const parentFolder = driveFile.getParents().hasNext() ? driveFile.getParents().next() : DriveApp.getRootFolder();
    
    const year = new Date().getFullYear();
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");
    const pdfName = `PC_Tracker_Archive_${year}-${year+1}`;

    // ---------------------------------------------------------
    // 1. FREEZE THE SHEET (DELETE TRIGGERS)
    // ---------------------------------------------------------
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
      const funcName = triggers[i].getHandlerFunction().toLowerCase();
      // Delete any trigger associated with the Phone Sync
      if (funcName.includes('sync') || funcName.includes('nightly')) {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
    
    // Update settings memory so the UI knows it's off
    const props = PropertiesService.getDocumentProperties();
    props.setProperty('syncToPhones', 'false');

    // ---------------------------------------------------------
    // 2. RENAME MAIN SPREADSHEET TO "RETIRED"
    // ---------------------------------------------------------
    const currentName = ss.getName();
    if (!currentName.includes("- Retired on")) {
      ss.rename(currentName + " - Retired on " + dateStr);
    }

    // ---------------------------------------------------------
    // 3. CREATE TEMPORARY SPREADSHEET FOR EXPORT
    // ---------------------------------------------------------
    const tempSs = SpreadsheetApp.create("TEMP_" + pdfName);
    const tempSsId = tempSs.getId();
    
    const contactLog = ss.getSheetByName("Contact Log");
    const combinedLog = ss.getSheetByName("Combined Contact Tracking");
    
    let hasData = false;
    
    // Copy the sheets over
    if (contactLog) {
      const tempCl = contactLog.copyTo(tempSs).setName("Contact Log");
      tempCl.setHiddenGridlines(true); // Ensure clean PDF look
      hasData = true;
      
      // --- UNHIDE ALL FILTERED ROWS ---
      // 1. Remove any active filters so everything shows up in the PDF
      if (tempCl.getFilter()) {
        tempCl.getFilter().remove();
      }
      // 2. Unhide all rows just in case they were manually hidden
      tempCl.showRows(1, tempCl.getMaxRows());
      
      // --- CROP THE CONTACT LOG AT THE "END" BAR ---
      // This guarantees the PDF ends exactly one row above the END bar.
      const endCell = tempCl.createTextFinder("END").matchEntireCell(true).matchCase(false).findNext();
      if (endCell) {
        const endRow = endCell.getRow();
        const maxRows = tempCl.getMaxRows();
        
        // Start deleting 1 row ABOVE the "END" row (the blank spacer)
        const startDeleteRow = endRow > 1 ? endRow - 1 : endRow; 
        
        if (maxRows >= startDeleteRow) {
          tempCl.deleteRows(startDeleteRow, maxRows - startDeleteRow + 1);
        }
      }
    }
    
    if (combinedLog) {
      combinedLog.copyTo(tempSs).setName("Combined Contact Tracking");
      hasData = true;
    }
    
    if (!hasData) {
      DriveApp.getFileById(tempSsId).setTrashed(true);
      throw new Error("Could not find Contact Log or Combined Contact Tracking to archive.");
    }
    
    // Delete the default "Sheet1" that is created with new spreadsheets
    const sheet1 = tempSs.getSheetByName("Sheet1");
    if (sheet1) tempSs.deleteSheet(sheet1);
    
    SpreadsheetApp.flush(); // Ensure everything is fully saved to Google's servers

    // ⚠️ CRITICAL PAUSE: Give Google's export servers time to register the new temp sheet
    Utilities.sleep(5000);

    // ---------------------------------------------------------
    // 4. GENERATE PDF (BULLETPROOF DUAL-LAYER METHOD)
    // ---------------------------------------------------------
    let pdfBlob;
    
    try {
      // ATTEMPT 1: Try the formatted URL export
      const url = "https://docs.google.com/spreadsheets/d/" + tempSsId + "/export?exportFormat=pdf&format=pdf&size=LETTER&portrait=false&fitw=true&gridlines=false";
      const token = ScriptApp.getOAuthToken();
      const response = UrlFetchApp.fetch(url, {
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
      });
      
      // If Google Security blocks it, it returns HTML. We MUST catch this!
      if (response.getResponseCode() === 200 && !response.getContentText().includes("<html")) {
         pdfBlob = response.getBlob().setName(pdfName + ".pdf");
      } else {
         throw new Error("URL fetch returned an HTML login page due to Workspace security.");
      }
      
    } catch (urlError) {
      // ATTEMPT 2: Fallback to Native Drive Conversion (Guaranteed to work)
      pdfBlob = DriveApp.getFileById(tempSsId).getAs('application/pdf').setName(pdfName + ".pdf");
    }

    // Save PDF to Drive
    const savedPdf = parentFolder.createFile(pdfBlob);
    
    // Cleanup: Trash the temporary spreadsheet
    DriveApp.getFileById(tempSsId).setTrashed(true);
    
    return { success: true, fileUrl: savedPdf.getUrl(), isFallback: false };

  } catch (e) {
    // Cleanup on fatal error so we don't leave mess behind
    if (typeof tempSsId !== 'undefined') DriveApp.getFileById(tempSsId).setTrashed(true);
    return { success: false, error: e.message };
  }
}
