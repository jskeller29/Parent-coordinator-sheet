// ======================================================================
// FILE: z_InputDataDebug.gs  (owner's testing only: DELETE IT on the copies that
// go to schools / the master template; nothing else changes without it)
// PURPOSE: the Input Data Mode debug log (moved out of Input Data.gs 2026-11-12,
// owner's request). While this file is in the project:
//   - every Input Data Mode sidebar call returns its step-by-step log
//     (idm_log_ lines in Input Data.gs), shown in the sidebar's Debug panel
//     with "Check now" / "Copy debug info";
//   - each 2-second check also logs a snapshot of the tabs (idm_debugSheets_).
// Without it: no log is kept, the replies carry no debug info, and the panel
// stays hidden. Input Data.gs only reaches this file through typeof checks.
// ======================================================================

/** Input Data.gs asks this before every sidebar call (idm_debugOn_). Return false to pause it. */
function idm_debugEnabled_() {
  return true;
}

/** Snapshot of what the server sees, for the Debug panel. */
function idm_debugSheets_(ss, state, temp) {
  if (!IDM_DBG) return;
  try {
    const active = ss.getActiveSheet();
    idm_log_("state: " + JSON.stringify(state));
    idm_log_("active tab (as the script sees it): " + (active ? active.getName() + " #" + active.getSheetId() : "none"));
    idm_log_("tabs: " + ss.getSheets().map(s => s.getName() + " #" + s.getSheetId() + (s.isSheetHidden() ? " (hidden)" : "")).join(" | "));
    if (temp) {
      idm_log_("temp tab: " + temp.getName() + " #" + temp.getSheetId() + " lastRow=" + temp.getLastRow() +
        " lastCol=" + temp.getLastColumn() + " maxRows=" + temp.getMaxRows() + " maxCols=" + temp.getMaxColumns() +
        " A1=" + JSON.stringify(String(temp.getRange("A1").getValue())));
    }
  } catch (e) { idm_log_("debug snapshot failed: " + e.message); }
}
