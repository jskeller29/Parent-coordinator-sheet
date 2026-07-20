// ======================================================================
// FILE: Constants.gs
// PURPOSE: Single source of truth for the Contact Log schema, plus the
// shared data-row formatter. All .gs files share one global scope, so a
// future column shift only needs to be updated here.
// ======================================================================

const CL_WIDTH = 16;                // Full row width, A:P
const CL_COL_GUARDIAN = 3;          // C — Guardian name
const CL_COL_SPOKE_WITH = 4;        // D — Person spoken with
const CL_COL_OSIS = 5;              // E — OSIS (also holds the "END" marker on the end bar row)
const CL_COL_STUDENT = 6;           // F — Student name
const CL_COL_TYPE = 12;             // L — Type dropdown
const CL_COL_DIVIDER_NOTES = 13;    // M — divider line before Notes
const CL_COL_FILTER_TOGGLE = 15;    // O — master hide-filter checkbox on the end bar row
const CL_COL_DIVIDER_FOLLOWUP = 15; // O — divider line before Followup Notes
const CL_COL_CHECKBOX = 16;         // P — row checkbox

/**
 * Applies the standard Contact Log data-row formatting:
 * - no borders on the Guardian/Spoke With pair (C:D)
 * - no horizontal borders anywhere (only the dark brown end bar carries one)
 * - divider lines before Notes (M) and Followup Notes (O)
 * - checkboxes planted in P
 */
function formatContactLogRows_(sheet, startRow, numRows) {
  if (!numRows || numRows < 1) return;
  sheet.getRange(startRow, CL_COL_GUARDIAN, numRows, 2).setBorder(false, false, false, false, false, false);
  sheet.getRange(startRow, 1, numRows, CL_WIDTH).setBorder(false, null, false, null, null, false);
  sheet.getRange(startRow, CL_COL_DIVIDER_NOTES, numRows, 1).setBorder(null, null, null, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(startRow, CL_COL_DIVIDER_FOLLOWUP, numRows, 1).setBorder(null, null, null, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(startRow, CL_COL_CHECKBOX, numRows, 1).insertCheckboxes().setFontColor("#000000");
}
