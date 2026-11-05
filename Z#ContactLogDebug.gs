// ======================================================================
// FILE: Z#ContactLogDebug.gs  (+ Z#ContactLogDebugSidebar.html)
// PURPOSE: TEMPORARY "🐞 Contact Log Debug" sidebar (owner's request 2026-11-05):
// shows, live, what every Contact Log edit / row delete did and how long each
// step took (ms since the script started), plus a snapshot of the search bar's
// Communications box and the notes Combined Contact Tracking holds for it.
// "Copy report" gives one block of text to send.
//
// How it records: while this sidebar is open it keeps the document cache key
// CL_DEBUG_ON alive (every poll), and the shared logger in Core.gs
// (clDbgStart_ / clDbg_ / clDbgEnd_) writes each run to CL_DEBUG_LOG. Closing
// the sidebar stops recording by itself within 5 minutes ("Stop" stops it now).
// Note: the time Google takes to START the trigger after you type is not
// visible to the script; the times shown start when the script starts.
//
// Z# = trial file: nothing relies on it. Its only link is the guarded
//   if (typeof addContactLogDebugMenuItem_ === "function") addContactLogDebugMenuItem_(menu);
// line in Menus.gs buildMenus_. Delete both Z#ContactLogDebug files to remove it.
// ======================================================================

var CLD_ON_SECONDS = 300;

function addContactLogDebugMenuItem_(menu) {
  menu.addItem("🐞 Contact Log Debug", "openContactLogDebug");
}

function openContactLogDebug() {
  CacheService.getDocumentCache().put(CL_DEBUG_ON_KEY, "1", CLD_ON_SECONDS);
  const html = htmlFromFile_("Z#ContactLogDebugSidebar").setTitle("Contact Log Debug");
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Sidebar poll: keeps recording on and returns the newest runs. */
function cld_poll() {
  const cache = CacheService.getDocumentCache();
  cache.put(CL_DEBUG_ON_KEY, "1", CLD_ON_SECONDS);
  let log = [];
  try { log = JSON.parse(cache.get(CL_DEBUG_LOG_KEY) || "[]"); } catch (e) { log = []; }
  return { log: log, now: new Date().toLocaleTimeString() };
}

function cld_stop() {
  CacheService.getDocumentCache().remove(CL_DEBUG_ON_KEY);
  return true;
}

function cld_clear() {
  CacheService.getDocumentCache().remove(CL_DEBUG_LOG_KEY);
  return cld_poll();
}

/** Puts the Communications formula back (same as running restoreSearchBarCommunicationsFormula). */
function cld_fixFormula() {
  restoreSearchBarCommunicationsFormula();
  return cld_state();
}

/** A snapshot of everything the search bar / notes depend on, with the time each read took. */
function cld_state() {
  const out = [];
  const t0 = Date.now();
  const add = (k, v) => out.push([k, String(v)]);
  const timed = (k, fn) => {
    const t = Date.now();
    let v;
    try { v = fn(); } catch (e) { v = "❌ " + e.message; }
    add(k + " (" + (Date.now() - t) + " ms)", v);
    return v;
  };
  const short = (text, n) => { text = String(text); return text.length > n ? text.slice(0, n) + "… (" + text.length + " chars)" : text; };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getDocumentProperties();
  add("Owner sheet (z_MINE.gs)", getOwnerHooks_() ? "yes" : "no");
  add("Update notes as I type (autoNotes)", props.getProperty("autoNotes") === "false" ? "OFF" : "on");
  add("Hidden notes included (disableHiddenNotes)", props.getProperty("disableHiddenNotes") === "true" ? "yes" : "no");
  add("Notes date range", (props.getProperty("notesStartDate") || "-") + " to " + (props.getProperty("notesEndDate") || "-"));
  add("Communications mode (SEARCH_BAR_COMMS_MODE)", props.getProperty(SEARCH_BAR_COMMS_PROP) || "(not set = formula, checked once)");
  timed("Triggers (this user)", () => ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction() + ":" + t.getEventType()).join(", ") || "none");
  add("Typist data cached", CacheService.getDocumentCache().get(GHOST_CACHE_KEY_) ? "yes" : "no");

  const sheet = ss.getSheetByName("Contact Log");
  if (!sheet) { add("Contact Log", "missing"); return { rows: out, ms: Date.now() - t0 }; }
  const endRow = timed("END row (scan column E)", () => findEndRowInColumnE_(sheet));
  const at = timed("Communications cell", () => {
    const a = findSearchBarComms_(sheet, endRow);
    return a ? sheet.getRange(a.row, a.col).getA1Notation() : "NOT FOUND in the 15 rows below END";
  });
  const found = endRow ? findSearchBarComms_(sheet, endRow) : null;
  if (found) {
    const label = sheet.getRange(found.row, found.col);
    const formula = label.getFormula();
    add("Communications cell has the formula", formula ? "yes" : "NO (plain text)");
    const osisCell = sheet.getRange(found.row - 1, found.col + 1);
    const nameCell = sheet.getRange(found.row - 2, found.col + 2);
    const osisBox = osisCell.getDisplayValue();
    add("OSIS box " + osisCell.getA1Notation(), JSON.stringify(osisBox) + (osisCell.getFormula() ? "  formula: " + short(osisCell.getFormula(), 200) : "  (typed, no formula)"));
    add("Name switch " + nameCell.getA1Notation(), JSON.stringify(nameCell.getDisplayValue()));
    const notesCell = sheet.getRange(found.row, found.col + 1);
    add("Notes shown " + notesCell.getA1Notation(), JSON.stringify(short(notesCell.getDisplayValue(), 300)));

    // What Combined holds for the first OSIS in the box, and what the logs give now
    const first = searchBarOsisList_(osisBox)[0];
    if (first) {
      const fromCombined = timed("Combined notes for " + first, () => {
        const v = combinedNotesReader_(ss)(first);
        return v === "" ? "(no row / empty)" : short(v, 300);
      });
      const fromLogs = timed("Notes built from the logs for " + first, () => {
        const v = logsNotesReader_(ss, sheet, endRow)(first);
        return v === "" ? "(none)" : short(v, 300);
      });
      add("Combined matches the logs", fromCombined === fromLogs ? "yes" : "NO (Combined is out of date for this student)");
    }
  }

  // The newest Contact Log rows (what the search bar may be pointing at)
  if (endRow && endRow > 2) {
    const from = Math.max(2, endRow - 3);
    const rows = sheet.getRange(from, 1, endRow - from, 16).getDisplayValues();
    rows.forEach((r, i) => add("Contact Log row " + (from + i), "B " + r[1] + " | C " + r[2] + " | E " + r[4] + " | F " + r[5] + " | M " + short(r[12], 60)));
  }
  return { rows: out, ms: Date.now() - t0 };
}
