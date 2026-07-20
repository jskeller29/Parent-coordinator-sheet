// ======================================================================
// FILE: VersionCheck.gs
// PURPOSE: On open, checks the master "Parent Coordinator Version Tracker"
// sheet and pops a changelog dialog when a newer template version exists.
//
// TRACKER LAYOUT (first tab, newest release ALWAYS inserted at Row 2):
//   Row 1 Headers:  Version | Size | Notes | Link
//   Row 2+:         5.1     | Small | Fixed XYZ bug | https://.../copy
//
// RELEASE CHECKLIST (every time you ship a new template):
//   1. Make your code changes in the template.
//   2. Bump SCRIPT_VERSION below (e.g. "5.0" -> "5.1").
//   3. Insert a new row at the TOP of the tracker (Row 2) with the SAME
//      version number, its Size (Major/Small), Notes, and the /copy link.
//      (Only the newest filled-in Link matters — older links are ignored.)
// ======================================================================

// 🚨 BUMP THIS EVERY RELEASE — must match the tracker row you add!
const SCRIPT_VERSION = "5.1"; // Format: MAJOR.SMALL  (5 major, 0 small)

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
 * Manual checks bypass the throttle AND "Dismiss Forever".
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

  const payload = computeUpdatePayload_();

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

  if (payload.entries.length === 0) {
    if (isManual) showUpToDateAlert_(payload.hiddenSmallCount);
    return; // Nothing new (for this user's preference) — stay quiet
  }

  // Stash the payload so the dialog can grab it instantly when it loads.
  // (5 minute lifetime; the dialog recomputes automatically if it expires.)
  try {
    CacheService.getUserCache().put('VC_PAYLOAD', JSON.stringify(payload), 300);
  } catch (e) {}

  // Plain static HTML — the dialog fetches its data via google.script.run
  const html = HtmlService.createHtmlOutputFromFile('VersionDialog')
    .setWidth(540)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "🚀 New Version Available!");
}

/**
 * Reads the tracker and builds everything the dialog needs.
 * Returns null only if the tracker can't be reached at all.
 */
function computeUpdatePayload_() {
  const props = PropertiesService.getDocumentProperties();

  let rows;
  try {
    const trackerSS = SpreadsheetApp.openById(VERSION_TRACKER_ID);
    const tab = trackerSS.getSheets()[0]; // Changelog must stay the FIRST tab
    const lastRow = tab.getLastRow();
    rows = (lastRow < 2) ? [] : tab.getRange(2, 1, lastRow - 1, 4).getDisplayValues();
  } catch (e) {
    console.error("Could not read Version Tracker: " + e.message);
    return null;
  }

  const localVersion = normVersion_(props.getProperty('VERSION_LAST_SEEN') || SCRIPT_VERSION);
  const majorsOnly   = props.getProperty('VERSION_MAJORS_ONLY') === 'true';

  // WALK DOWN the tracker (newest on top), collecting every release
  // ABOVE the version this sheet has already seen.
  const missed = [];
  for (let i = 0; i < rows.length; i++) {
    const ver = normVersion_(rows[i][0]);
    if (!ver) continue;                 // Skip blank / junk rows
    if (ver === localVersion) break;    // Reached what we already have — stop walking
    missed.push({
      version: ver,
      size:  String(rows[i][1] || "").trim(),
      notes: String(rows[i][2] || "").trim(),
      link:  String(rows[i][3] || "").trim()
    });
  }
  // Safety net: if the local version's row was ever deleted from the
  // tracker, the loop simply collects everything instead of crashing.

  // "Majors only" still walks the full list, then surfaces ONLY Major rows
  const entries = majorsOnly ? missed.filter(e => isMajor_(e.size)) : missed;

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
    entries: entries,
    latestLink: latestLink,
    newestVersion: entries.length ? entries[0].version : "",
    installedVersion: normVersion_(SCRIPT_VERSION),
    majorsOnly: majorsOnly,
    hiddenSmallCount: missed.length - entries.length
  };
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
  } catch (e) {}

  const payload = computeUpdatePayload_();
  return payload ? JSON.stringify(payload) : null;
}

/** "Dismiss": remember the newest version shown so it won't pop again until something newer ships. */
function vc_dismiss(newestVersion) {
  const v = normVersion_(newestVersion);
  if (v) PropertiesService.getDocumentProperties().setProperty('VERSION_LAST_SEEN', v);
  return true;
}

/** "Dismiss Forever": stop all automatic checks. (The manual menu check still works.) */
function vc_dismissForever(newestVersion) {
  const props = PropertiesService.getDocumentProperties();
  const v = normVersion_(newestVersion);
  if (v) props.setProperty('VERSION_LAST_SEEN', v);
  props.setProperty('VERSION_DISMISS_FOREVER', 'true');
  return true;
}

/** "Major updates only" checkbox in the dialog. */
function vc_setMajorsOnly(enabled) {
  PropertiesService.getDocumentProperties()
    .setProperty('VERSION_MAJORS_ONLY', String(enabled === true));
  return true;
}


// ==========================================
// HELPERS
// ==========================================

/**
 * Normalizes versions so "5", "5.0", and " 5.0 " all compare as "5.0".
 * (Sheets displays a numeric 5.0 cell as just "5", so this is required.)
 * Returns "" for junk that isn't a version at all.
 */
function normVersion_(v) {
  const parts = String(v == null ? "" : v).trim().split(".");
  const major = parseInt(parts[0], 10);
  if (isNaN(major)) return "";
  const small = parseInt(parts[1], 10);
  return major + "." + (isNaN(small) ? 0 : small);
}

/** "Major", "MAJOR", "major update" all count as Major. Everything else = Small. */
function isMajor_(size) {
  return String(size || "").trim().toLowerCase().indexOf("major") === 0;
}

/** Friendly alert for manual checks when there's nothing (visible) to show. */
function showUpToDateAlert_(hiddenCount) {
  const props = PropertiesService.getDocumentProperties();
  const seen = normVersion_(props.getProperty('VERSION_LAST_SEEN') || SCRIPT_VERSION);

  let msg = "You're caught up! Latest version you've reviewed: v" + seen + ".";
  if (hiddenCount > 0) {
    msg += "\n\nHeads up: there " +
      (hiddenCount === 1 ? "is 1 Small update" : "are " + hiddenCount + " Small updates") +
      " you haven't seen, hidden by your 'Major updates only' setting.";
  }
  SpreadsheetApp.getUi().alert("✅ Up to Date", msg, SpreadsheetApp.getUi().ButtonSet.OK);
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
// Because it's a Document Property, the setting travels with every copy.
// To offer BOTH links at once, keep two template files: your master
// (toggle OFF) and a copy of it with the toggle flipped ON.
// ======================================================================
function toggleAutoMigratePopup() {
  const props = PropertiesService.getDocumentProperties();
  const isOn = props.getProperty('AUTO_MIGRATE_POPUP') === 'true';
  props.setProperty('AUTO_MIGRATE_POPUP', String(!isOn));
  SpreadsheetApp.getActiveSpreadsheet().toast(
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
  ['VERSION_LAST_SEEN', 'VERSION_DISMISS_FOREVER', 'VERSION_MAJORS_ONLY', 'VERSION_LAST_CHECK']
    .forEach(k => props.deleteProperty(k));
  try { CacheService.getUserCache().remove('VC_PAYLOAD'); } catch (e) {}
  SpreadsheetApp.getActiveSpreadsheet().toast("Version check memory wiped for this sheet.", "Dev Reset", 5);
}
