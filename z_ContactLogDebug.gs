// ======================================================================
// FILE: z_ContactLogDebug.gs (+ z_ContactLogDebugSidebar.html)
// Owner's testing only (2026-11-12, owner's request): DELETE BOTH z_ContactLogDebug
// files on the copies that go to schools / the master template; then the
// "🐞 Contact Log Debug" menu item is simply not there (the two lines in
// buildMenus_, Menus.gs, check typeof first) and nothing else changes.
// (Was the trial file Z#ContactLogDebug.gs, then the end of Edits.gs.)
//
// TEMPORARY "🐞 Contact Log Debug" sidebar (owner's request 2026-11-05): shows,
// live, what every Contact Log edit / row delete did and how long each step took
// (ms since the script started), plus a snapshot of the search bar's
// Communications box and the notes Combined Contact Tracking holds for it.
// "Copy report" gives one block of text to send.
//
// How it records: while this sidebar is open it keeps the document cache key
// CL_DEBUG_ON alive (every poll), and the shared logger in Core.gs
// (clDbgStart_ / clDbg_ / clDbgEnd_) writes each run to CL_DEBUG_LOG. Without
// this file the key is never set, so the logger costs one cache read per edit.
// It uses the search bar helpers in Edits.gs (findSearchBarComms_,
// combinedNotesReader_, logsNotesReader_, restoreSearchBarCommunicationsFormula).
// ======================================================================

var CLD_ON_SECONDS = 300;

function addContactLogDebugMenuItem_(menu) {
  menu.addItem("🐞 Contact Log Debug", "openContactLogDebug");
}

function openContactLogDebug() {
  CacheService.getDocumentCache().put(CL_DEBUG_ON_KEY, "1", CLD_ON_SECONDS);
  const html = htmlFromFile_("z_ContactLogDebugSidebar").setTitle("Contact Log Debug");
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

/**
 * A snapshot of everything the search bar / notes depend on, with the time each read took.
 * anon: note texts become "[N chars, K lines, #fingerprint]" (same text = same fingerprint);
 * names / OSIS are replaced later by cld_report's anonymizer.
 */
function cld_state(anon) {
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
  const short = (text, n) => {
    text = String(text);
    if (anon) return cldNoteSummary_(text);
    return text.length > n ? text.slice(0, n) + "… (" + text.length + " chars)" : text;
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getDocumentProperties();
  add("Owner sheet (z_MINE.gs)", getOwnerHooks_() ? "yes" : "no");
  add("Update notes as I type (autoNotes)", props.getProperty("autoNotes") === "false" ? "OFF" : "on");
  add("Hidden notes included (disableHiddenNotes)", props.getProperty("disableHiddenNotes") === "true" ? "yes" : "no");
  add("Notes date range", (props.getProperty("notesStartDate") || "-") + " to " + (props.getProperty("notesEndDate") || "-"));
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
    const up = searchBarOsisRowsAbove_(sheet, endRow, found); // 1 in the old layout
    const osisCell = sheet.getRange(found.row - up, found.col + 1);
    const nameCell = sheet.getRange(found.row - up - 1, found.col + 2);
    const osisBox = osisCell.getDisplayValue();
    const osisFormula = osisCell.getFormula();
    add("OSIS box " + osisCell.getA1Notation(), JSON.stringify(osisBox) + (osisFormula ? "  formula: " + (osisFormula.length > 300 ? osisFormula.slice(0, 300) + "…" : osisFormula) : "  (typed, no formula)"));
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
      // What each source of the notes costs to read
      timed("  read Contact Log to END", () => sheet.getRange(1, 1, endRow, CL_WIDTH).getDisplayValues().length + " rows");
      timed("  read Backend Event Log", () => backendReadBlock_(ss, "EVENT_LOG", true).length + " rows");
      timed("  read Notes tab", () => { const n = ss.getSheetByName("Notes"); return n ? n.getDataRange().getValues().length + " rows x " + n.getLastColumn() + " cols" : "no tab"; });
      timed("  Backend tab size", () => { const b = ss.getSheetByName(BACKEND_SHEET_NAME); return b ? b.getLastRow() + " rows x " + b.getLastColumn() + " cols" : "no tab"; });
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
    rows.forEach((r, i) => add("Contact Log row " + (from + i), "B " + r[1] + " | C " + r[2] + " | D " + (anon ? (r[3] ? "(filled)" : "(blank)") : r[3]) + " | E " + r[4] + " | F " + r[5] +
      " | J/K " + r[9] + "/" + r[10] + " | M " + short(r[12], 60) + " | O " + short(r[14], 60) + " | P " + r[15]));
  }
  return { rows: out, ms: Date.now() - t0 };
}

// ======================================================================
// REPORT (optionally anonymized, to share)
// ======================================================================

/** "[N chars, K lines, #abcd1234]": the same text always gives the same fingerprint. */
function cldNoteSummary_(text) {
  text = String(text);
  if (text === "") return "(empty)";
  let h = 5381;
  for (let i = 0; i < text.length; i++) h = ((h * 33) ^ text.charCodeAt(i)) >>> 0;
  return "[" + text.length + " chars, " + text.split("\n").length + " lines, #" + h.toString(16) + "]";
}

/**
 * Replaces every student / guardian name (Student List B / C, whole names and
 * their single words), every OSIS, phone number and email with placeholders.
 * The same name always gets the same placeholder within one report.
 */
function cldAnonymizer_(ss) {
  const pairs = [];   // [text, placeholder]
  const seen = new Map();
  const counters = { Student: 0, Guardian: 0, Name: 0, OSIS: 0 };
  const addName = (raw, kind) => {
    const name = String(raw).replace(/\([^)]*\)/g, " ").replace(/\s+/g, " ").trim();
    if (name.length < 2 || seen.has(name.toLowerCase())) return;
    const ph = kind + "-" + (++counters[kind]);
    seen.set(name.toLowerCase(), ph);
    pairs.push([name, ph]);
  };
  const rows = studentLookupRows_(ss, 3).map(r => r.map(String)); // Student List names (Backend copy)
  const splitNames = v => String(v).split(/,|&|\band\b|\n/i).map(x => x.trim()).filter(Boolean);
  rows.forEach(r => {
    String(r[0]).split(/[^0-9]+/).filter(x => x.length >= 6).forEach(o => {
      if (!seen.has(o)) { const ph = "OSIS-" + (++counters.OSIS); seen.set(o, ph); pairs.push([o, ph]); }
    });
    splitNames(r[1]).forEach(n => addName(n, "Student"));
    splitNames(r[2]).forEach(n => addName(n, "Guardian"));
  });
  // Single words of every name (first / last names on their own), after the whole names
  pairs.slice().forEach(([text, ph]) => {
    if (/^OSIS-/.test(ph)) return;
    text.split(" ").forEach(w => { if (/^[A-Za-zÀ-ÿ'’-]{2,}$/.test(w)) addName(w, "Name"); });
  });
  pairs.sort((a, b) => b[0].length - a[0].length);
  const esc = t => t.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const rules = pairs.map(([text, ph]) => [new RegExp("(^|[^A-Za-z0-9À-ÿ])" + esc(text) + "(?=$|[^A-Za-z0-9À-ÿ])", "gi"), ph]);

  return text => {
    let out = String(text);
    rules.forEach(([re, ph]) => { out = out.replace(re, (m, pre) => pre + ph); });
    return out
      .replace(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g, "[email]")
      .replace(/\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}/g, "[phone]")
      .replace(/\b\d{9}\b/g, "OSIS-?")
      // Typed text that is not a known name (run titles ' = "…"' and "'…' not found")
      .replace(/ = "([^"]*)"/g, (m, v) => ' = "' + cldMaskTyped_(v) + '"')
      .replace(/'([^']*)' not found/g, (m, v) => "'" + cldMaskTyped_(v) + "' not found");
  };
}

/** Keeps placeholders / numbers / the "(paste …)" marker; anything else with letters becomes [typed text]. */
function cldMaskTyped_(v) {
  if (!/[A-Za-zÀ-ÿ]/.test(v) || /^\(paste/.test(v)) return v;
  return /^((Student|Guardian|Name|OSIS)-[0-9?]+[ ,&]*)+$/.test(v.trim()) ? v : "[typed text]";
}

/** One block of text to send: the search bar snapshot + the newest runs. */
function cld_report(anon) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const state = cld_state(anon);
  let log = [];
  try { log = JSON.parse(CacheService.getDocumentCache().get(CL_DEBUG_LOG_KEY) || "[]"); } catch (e) { log = []; }

  const lines = ["CONTACT LOG DEBUG REPORT " + new Date().toLocaleString() + (anon ? " (anonymized)" : ""), ""];
  lines.push("SEARCH BAR & NOTES (checked in " + state.ms + " ms)");
  state.rows.forEach(r => lines.push("  " + r[0] + ": " + r[1]));
  lines.push("", "RUNS (newest first)");
  log.forEach(r => {
    lines.push("- " + r.label + "  [" + (r.at || "") + ", total " + (r.total || 0) + " ms]");
    let prev = 0;
    (r.steps || []).forEach(st => { lines.push("    " + st[0] + " ms (+" + (st[0] - prev) + "): " + st[1]); prev = st[0]; });
  });
  const text = lines.join("\n");
  return anon ? cldAnonymizer_(ss)(text) : text;
}
