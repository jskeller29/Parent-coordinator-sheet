// ======================================================================
// FILE: VersionCheck.gs
// PURPOSE: On open, checks the master "Parent Coordinator Version Tracker"
// sheet and pops a changelog dialog when a NEWER-DATED template exists.
//
// HOW THE BASELINE WORKS (date-based):
//   • Each template copy carries its own build date in the hidden "Version"
//     tab, cell A1. That date is stamped by resetToDisplayMode() in
//     Regenerate.gs and travels with every copy of the sheet.
//   • The checker collects every tracker row DATED AFTER that build date
//     (or after the newest date the user has already dismissed).
//
// TRACKER LAYOUT (first tab, newest release ALWAYS inserted at Row 2):
//   Row 1 Headers:  Date | Size | Notes | Link
//   Row 2+:         07/22/26 | Small | Fixed XYZ bug | https://.../copy
//
// RELEASE CHECKLIST (every time you ship a new template):
//   1. Make your code changes in the template.
//   2. Run resetToDisplayMode() so the template's "Version"!A1 build date is
//      stamped to today (or set it by hand).
//   3. Insert a new row at the TOP of the tracker (Row 2) with today's Date,
//      its Size (Major/Small), Notes, and the /copy link.
//      (Only the newest filled-in Link matters — older links are ignored.)
// ======================================================================

// The master tracker spreadsheet (shared: Anyone with the link -> Viewer)
const VERSION_TRACKER_ID = "1nnJXnYEGyISvQtvsp2M4D1EQT2cHhNWkBXMG1Ubbljc";

// How often (in hours) the sheet actually reaches out to the tracker on open
const VERSION_CHECK_HOURS = 6;


// ==========================================
// TRIGGER + MENU ENTRY POINTS
// ==========================================

/**
 * Installable OnOpen trigger handler (created in setupAllInstallableTriggers).
 * NOTE: This MUST be an installable trigger — the simple onOpen() runs in
 * limited-auth mode and is not allowed to open other spreadsheets.
 */
function checkVersionOnOpen() {
  try {
    checkForUpdates_(false);
  } catch (e) {
    console.error("Version check failed: " + e.message);
  }
}

/**
 * Menu item: 🚀 App Menu -> 🔔 Check for Updates
 * Manual checks bypass the throttle AND "Dismiss Forever", and ALWAYS open
 * the dialog (even when caught up) so the Major-only / all-updates toggle is
 * always reachable.
 */
function manualCheckForUpdates() {
  checkForUpdates_(true);
}


// ==========================================
// CORE FLOW
// ==========================================
function checkForUpdates_(isManual) {
  const props = PropertiesService.getDocumentProperties();

  // Never run before the user has authorized / completed setup
  if (props.getProperty('SETUP_COMPLETE') !== 'true') return;

  if (!isManual) {
    // Respect "Dismiss Forever"
    if (props.getProperty('VERSION_DISMISS_FOREVER') === 'true') return;

    // Throttle: only actually contact the tracker every VERSION_CHECK_HOURS
    const lastCheck = Number(props.getProperty('VERSION_LAST_CHECK') || 0);
    if (Date.now() - lastCheck < VERSION_CHECK_HOURS * 60 * 60 * 1000) return;
  }

  // Stamp the attempt FIRST so a broken tracker doesn't retry on every open
  props.setProperty('VERSION_LAST_CHECK', String(Date.now()));

  // Manual checks show BOTH Major and Small updates (showAll = true).
  const payload = computeUpdatePayload_(isManual);

  if (!payload) {
    // Tracker unreachable — stay silent on auto-checks
    if (isManual) {
      SpreadsheetApp.getUi().alert(
        "⚠️ Update Check Failed",
        "Could not reach the Version Tracker sheet. Check your internet connection, or the tracker may have been moved or unshared.",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return;
  }

  // Auto checks stay quiet when there's nothing new. A MANUAL check always
  // opens the dialog so the user can flip Major-only / all updates even when
  // already caught up (the dialog renders a friendly "caught up" state).
  if (!isManual && payload.entries.length === 0) return;

  // Stash the payload so the dialog can grab it instantly when it loads.
  // (5 minute lifetime; the dialog recomputes automatically if it expires.)
  try {
    CacheService.getUserCache().put('VC_PAYLOAD', JSON.stringify(payload), 300);
  } catch (e) { console.error(e); }

  // Plain static HTML — the dialog fetches its data via google.script.run
  const title = payload.entries.length ? "🚀 New Version Available!" : "🔔 Update Preferences";
  const html = HtmlService.createHtmlOutputFromFile('VersionDialog')
    .setWidth(540)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, title);
}

/**
 * Reads the tracker and builds everything the dialog needs.
 * Returns null only if the tracker can't be reached at all.
 *
 * @param {boolean} showAll  When true (manual checks), surface BOTH Major and
 *   Small updates regardless of the "Major only" preference. Automatic checks
 *   pass false so that preference still governs the pop-up.
 */
function computeUpdatePayload_(showAll) {
  const props = PropertiesService.getDocumentProperties();

  let rows;
  try {
    const trackerSS = SpreadsheetApp.openById(VERSION_TRACKER_ID);
    const tab = trackerSS.getSheets()[0]; // Changelog must stay the FIRST tab
    const lastRow = tab.getLastRow();
    // getValues (not display) so date-formatted cells come back as real Dates.
    rows = (lastRow < 2) ? [] : tab.getRange(2, 1, lastRow - 1, 4).getValues();
  } catch (e) {
    console.error("Could not read Version Tracker: " + e.message);
    return null;
  }

  const baselineMs = getLocalBaselineDateMs_();
  const majorsOnly = props.getProperty('VERSION_MAJORS_ONLY') === 'true';

  // Collect every tracker row DATED AFTER our baseline. Order-independent, so
  // a deleted/rearranged tracker row never breaks the walk.
  const missed = [];
  for (let i = 0; i < rows.length; i++) {
    const ms = toDateMs_(rows[i][0]);
    if (isNaN(ms)) continue;        // Skip blank / junk / header-ish rows
    if (ms <= baselineMs) continue; // Already have it / already dismissed it
    missed.push({
      dateMs: ms,
      dateStr: formatTrackerDate_(rows[i][0]),
      size:  String(rows[i][1] || "").trim(),
      notes: String(rows[i][2] || "").trim(),
      link:  String(rows[i][3] || "").trim()
    });
  }
  missed.sort((a, b) => b.dateMs - a.dateMs); // Newest first

  // "Majors only" still walks the full list, then surfaces ONLY Major rows —
  // but a manual check (showAll) always shows everything, Major and Small.
  const entries = (majorsOnly && !showAll) ? missed.filter(e => isMajor_(e.size)) : missed;

  // The ONE link shown in the dialog: the newest filled-in Link among ALL
  // missed releases (scans the unfiltered list, so even a Majors-only user
  // gets pointed at the true latest template).
  let latestLink = "";
  for (let i = 0; i < missed.length; i++) {
    if (missed[i].link && /^https?:\/\//i.test(missed[i].link)) {
      latestLink = missed[i].link;
      break;
    }
  }

  return {
    entries: entries.map(e => ({ dateStr: e.dateStr, size: e.size, notes: e.notes, link: e.link })),
    latestLink: latestLink,
    newestDateMs: entries.length ? entries[0].dateMs : baselineMs,
    newestDateStr: entries.length ? entries[0].dateStr : "",
    installedDateStr: formatMsDate_(baselineMs),
    majorsOnly: majorsOnly,
    dismissForever: props.getProperty('VERSION_DISMISS_FOREVER') === 'true'
  };
}

/**
 * The baseline the tracker is compared against: the LATER of this copy's
 * build date ("Version"!A1) and the newest date the user has dismissed.
 * Returns 0 (compare against epoch => surface everything) when neither is set,
 * so an un-stamped sheet still discovers updates instead of going dark.
 */
function getLocalBaselineDateMs_() {
  let buildMs = NaN;
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Version");
    if (sheet) buildMs = toDateMs_(sheet.getRange("A1").getValue());
  } catch (e) { console.error(e); }

  const seenMs = Number(PropertiesService.getDocumentProperties().getProperty('VERSION_LAST_SEEN_DATE') || NaN);

  const candidates = [buildMs, seenMs].filter(n => !isNaN(n));
  return candidates.length ? Math.max.apply(null, candidates) : 0;
}


// ==========================================
// DIALOG SERVER CALLS (from VersionDialog.html)
// ==========================================

/**
 * The dialog asks for its data right after it loads.
 * Returns a JSON string (cached copy if fresh, otherwise recomputed live).
 */
function vc_getPayload() {
  try {
    const cached = CacheService.getUserCache().get('VC_PAYLOAD');
    if (cached) return cached;
  } catch (e) { console.error(e); }

  const payload = computeUpdatePayload_();
  return payload ? JSON.stringify(payload) : null;
}

/** "Dismiss": remember the newest date shown so it won't pop again until something newer ships. */
function vc_dismiss(newestDateMs) {
  const ms = Number(newestDateMs);
  if (!isNaN(ms) && ms > 0) {
    PropertiesService.getDocumentProperties().setProperty('VERSION_LAST_SEEN_DATE', String(ms));
  }
  return true;
}

/**
 * "Stop automatic update checks" toggle in the dialog. When ON, automatic
 * (on-open) checks are suppressed; the manual menu check always still works.
 */
function vc_setDismissForever(enabled) {
  PropertiesService.getDocumentProperties()
    .setProperty('VERSION_DISMISS_FOREVER', String(enabled === true));
  return true;
}

/** "Major updates only" checkbox in the dialog. Governs AUTOMATIC checks only. */
function vc_setMajorsOnly(enabled) {
  PropertiesService.getDocumentProperties()
    .setProperty('VERSION_MAJORS_ONLY', String(enabled === true));
  return true;
}


// ==========================================
// HELPERS
// ==========================================

/**
 * Converts a tracker/Version cell into a comparable date timestamp (midnight).
 * Handles real Date objects (from getValues) and common typed strings like
 * "7/22/26", "2026-07-22", and "Monday, 7/20/26". Returns NaN for non-dates.
 */
function toDateMs_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate()).getTime();
  }
  const s = String(value == null ? "" : value).trim();
  if (!s) return NaN;
  const cleaned = s.replace(/^[A-Za-z]+,\s*/, ""); // drop a leading "Monday, "
  const parsed = new Date(cleaned);
  if (!isNaN(parsed.getTime())) {
    return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate()).getTime();
  }
  return NaN;
}

/** Formats a timestamp as MM/dd/yy (or "unknown" for an unset baseline). */
function formatMsDate_(ms) {
  if (!ms || isNaN(ms)) return "unknown";
  return Utilities.formatDate(new Date(ms), Session.getScriptTimeZone(), "MM/dd/yy");
}

/** Formats a raw tracker cell for display: real Dates -> MM/dd/yy, else as-typed. */
function formatTrackerDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "MM/dd/yy");
  }
  return String(value == null ? "" : value).trim();
}

/** "Major", "MAJOR", "major update" all count as Major. Everything else = Small. */
function isMajor_(size) {
  return String(size || "").trim().toLowerCase().indexOf("major") === 0;
}


// ======================================================================
// 🩺 DIAGNOSTIC: WHY IS IT SAYING "UP TO DATE"?
// Run this from the Script Editor (or add it to a menu) when the checker
// shows the wrong result. It pops an alert showing the local build date,
// the stored baseline, and exactly how each tracker row's Date cell parses —
// so you can see whether a row is being skipped (NaN) or judged "not newer".
// ======================================================================
function debugVersionCheck() {
  const props = PropertiesService.getDocumentProperties();
  const lines = [];

  // --- Local build date (Version!A1) ---
  const vsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Version");
  if (!vsheet) {
    lines.push("Version tab: MISSING (build date unknown → baseline falls back to epoch).");
  } else {
    const v = vsheet.getRange("A1").getValue();
    lines.push("Version!A1 = \"" + v + "\"  [" + ((v instanceof Date) ? "Date" : typeof v) +
               "] → " + formatMsDate_(toDateMs_(v)));
  }

  const seenRaw = props.getProperty('VERSION_LAST_SEEN_DATE');
  lines.push("Dismissed-date memory = " + (seenRaw || "(none)") +
             (seenRaw ? " → " + formatMsDate_(Number(seenRaw)) : ""));

  const baseline = getLocalBaselineDateMs_();
  lines.push("BASELINE compared against = " + formatMsDate_(baseline) +
             "  (later of the two above)");
  lines.push("majorsOnly=" + (props.getProperty('VERSION_MAJORS_ONLY') === 'true') +
             ", autoChecksOff=" + (props.getProperty('VERSION_DISMISS_FOREVER') === 'true'));
  lines.push("");

  // --- Tracker rows ---
  try {
    const tab = SpreadsheetApp.openById(VERSION_TRACKER_ID).getSheets()[0];
    const lastRow = tab.getLastRow();
    lines.push("Tracker tab read = \"" + tab.getName() + "\"  (lastRow " + lastRow + ")");
    const rows = (lastRow < 2) ? [] : tab.getRange(2, 1, lastRow - 1, 4).getValues();
    if (!rows.length) lines.push("  (no data rows below the header)");
    rows.slice(0, 10).forEach(function (r, i) {
      const ms = toDateMs_(r[0]);
      const tag = isNaN(ms) ? "⚠️ NOT A DATE (skipped)" : (ms > baseline ? "✅ NEWER (would show)" : "— not newer");
      lines.push("  Row " + (i + 2) + ": \"" + r[0] + "\" [" +
                 ((r[0] instanceof Date) ? "Date" : typeof r[0]) + "] → " + formatMsDate_(ms) + "  " + tag);
    });
  } catch (e) {
    lines.push("TRACKER READ ERROR: " + e.message);
  }

  const report = lines.join("\n");
  console.log(report); // Also in the execution log as a backup
  try {
    SpreadsheetApp.getUi().alert("🩺 Version Check Diagnostic", report, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    // No UI context (pure Script Editor run) — the console.log above still has it.
  }
}


// ======================================================================
// 🛠️ DEV TOOL #1: AUTO-MIGRATE POPUP TOGGLE
// Run this from the Script Editor on a TEMPLATE before sharing its link.
//
//   ON  = every fresh copy auto-opens the Migration Wizard right after
//         Initial Setup (use for the link you give EXISTING users upgrading)
//   OFF = fresh copies get the normal Welcome guide
//         (use for the link you give BRAND NEW users)
//
// The setting lives in the hidden "Version" tab, cell B2 ("Yes"/"No"), so it
// travels with every copy and can also be edited by hand. To offer BOTH links
// at once, keep two template files: your master (B2 = No) and a copy with
// B2 = Yes.
// ======================================================================
function toggleAutoMigratePopup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Version");
  if (!sheet) {
    sheet = ss.insertSheet("Version");
    sheet.hideSheet();
  }
  const cell = sheet.getRange("B2");
  const isOn = String(cell.getValue()).trim().toLowerCase() === "yes";
  cell.setValue(isOn ? "No" : "Yes");
  ss.toast(
    "Auto-Migrate Popup is now " + (!isOn ? "ON ✅ (upgrade template)" : "OFF ❌ (new-user template)"),
    "Dev Setting Updated", 8
  );
}


// ======================================================================
// 🛠️ DEV TOOL #2: RESET VERSION MEMORY (for testing)
// Wipes Dismiss/Dismiss-Forever/Majors-Only/throttle on THIS sheet so you
// can re-test the popup. Pair with manualCheckForUpdates() to skip the
// 6-hour throttle entirely.
// ======================================================================
function resetVersionCheckState() {
  const props = PropertiesService.getDocumentProperties();
  ['VERSION_LAST_SEEN', 'VERSION_LAST_SEEN_DATE', 'VERSION_DISMISS_FOREVER', 'VERSION_MAJORS_ONLY', 'VERSION_LAST_CHECK']
    .forEach(k => props.deleteProperty(k));
  try { CacheService.getUserCache().remove('VC_PAYLOAD'); } catch (e) { console.error(e); }
  SpreadsheetApp.getActiveSpreadsheet().toast("Version check memory wiped for this sheet.", "Dev Reset", 5);
}
