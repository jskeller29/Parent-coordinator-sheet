// ======================================================================
// FILE: LinkResolver.gs
// PURPOSE: One source of truth for the shareable template links, so you
// never hand-paste a file ID again.
//
// HOW IT WORKS
//   Drop your template files into the Drive folder below and label them:
//     • a file whose NAME contains "(BLANK)"   -> the fresh new-user template
//     • a file whose NAME contains "(MIGRATE)" -> the upgrade template that
//       auto-opens the Migration Wizard
//   The resolver finds each file by that label and builds a "Make a copy"
//   link from its live file ID. Next release, just swap the files in the
//   folder (keeping the labels) — no code edits anywhere.
//
// RESILIENCE (why copies still work)
//   getTemplateLinks_() tries a live Drive scan first. On success it mirrors
//   the result into the hidden "Version" tab (cols D:E), which travels with
//   every copy of the sheet — so a plain copy that can't reach the folder
//   still inherits the last-known links, exactly like saved settings do.
//
// SHARING REQUIREMENT
//   For the live scan to work in OTHER coordinators' copies, the folder and
//   the (BLANK)/(MIGRATE) files must be shared "Anyone with the link -> Viewer".
//   (The copy links themselves already need that sharing to be usable.)
//   If you keep the folder private, run refreshTemplateLinks() on your master
//   template before sharing it and copies will inherit the links.
// ======================================================================

// The Drive folder that holds the labeled template files.
const TEMPLATE_FOLDER_ID = "1goKFrxDJOSAPxp1HUXETX2dns5c6HXk2";

// How long a successful folder scan is cached (script cache) before re-scanning.
const TEMPLATE_LINKS_CACHE_SECONDS = 6 * 60 * 60; // 6 hours

// Fallback link if a specific label can't be resolved — lands the user in the
// folder rather than on a dead link.
const TEMPLATE_FOLDER_URL = "https://drive.google.com/drive/folders/" + TEMPLATE_FOLDER_ID;


/**
 * Public entry point for the dialogs (called via google.script.run).
 * Returns { blankCopyUrl, migrateCopyUrl, blankName, migrateName, folderUrl }.
 * Never throws — always returns usable (folder-fallback) URLs.
 */
function getTemplateLinks() {
  return getTemplateLinks_();
}

/**
 * Menu/editor helper: force a fresh folder scan and show the resolved links.
 * Run this on your MASTER template after swapping the folder files so copies
 * inherit the newest links (it also writes them into the Version tab).
 */
function refreshTemplateLinks() {
  try { CacheService.getScriptCache().remove('TEMPLATE_LINKS_V1'); } catch (e) {}
  const links = getTemplateLinks_();
  const msg =
    "BLANK  (" + (links.blankName || "not found") + "):\n" + (links.blankCopyUrl || "(unresolved)") +
    "\n\nMIGRATE (" + (links.migrateName || "not found") + "):\n" + (links.migrateCopyUrl || "(unresolved)");
  try {
    SpreadsheetApp.getUi().alert("🔗 Template Links", msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    console.log(msg); // No UI context (pure editor run) — the log still has it.
  }
  return links;
}


// ==========================================
// CORE
// ==========================================

/**
 * Cached, resilient resolver. Tries a live scan; on success caches it (script
 * cache + Version tab). On failure, falls back to the Version-tab copy that
 * traveled with this spreadsheet, then to the folder URL.
 */
function getTemplateLinks_() {
  const cache = CacheService.getScriptCache();
  try {
    const cached = cache.get('TEMPLATE_LINKS_V1');
    if (cached) return JSON.parse(cached);
  } catch (e) { console.error(e); }

  let links = scanTemplateFolder_();

  if (links.blankCopyUrl || links.migrateCopyUrl) {
    // Live scan worked — persist so plain copies inherit it.
    try { saveTemplateLinksToVersionSheet_(links); } catch (e) { console.error(e); }
  } else {
    // Live scan failed (no access / offline) — use the inherited copy if any.
    const inherited = loadTemplateLinksFromVersionSheet_();
    if (inherited) links = inherited;
  }

  // Guarantee non-empty, non-dead URLs.
  if (!links.blankCopyUrl)   links.blankCopyUrl   = TEMPLATE_FOLDER_URL;
  if (!links.migrateCopyUrl) links.migrateCopyUrl = TEMPLATE_FOLDER_URL;
  links.folderUrl = TEMPLATE_FOLDER_URL;

  try { cache.put('TEMPLATE_LINKS_V1', JSON.stringify(links), TEMPLATE_LINKS_CACHE_SECONDS); } catch (e) {}
  return links;
}

/**
 * Live-scans the template folder for the (BLANK) and (MIGRATE) files and
 * builds "Make a copy" URLs. Picks the most-recently-updated file for each
 * label so a lingering older file never wins. Returns empty URLs on failure.
 */
function scanTemplateFolder_() {
  const out = { blankCopyUrl: "", migrateCopyUrl: "", blankName: "", migrateName: "" };
  try {
    const folder = DriveApp.getFolderById(TEMPLATE_FOLDER_ID);
    const files = folder.getFiles();
    let blank = null, migrate = null, blankTime = 0, migrateTime = 0;

    while (files.hasNext()) {
      const f = files.next();
      const upper = String(f.getName()).toUpperCase();
      const updated = f.getLastUpdated().getTime();
      if (upper.indexOf('(BLANK)') !== -1 && updated >= blankTime) { blank = f; blankTime = updated; }
      if (upper.indexOf('(MIGRATE)') !== -1 && updated >= migrateTime) { migrate = f; migrateTime = updated; }
    }

    if (blank)   { out.blankCopyUrl   = copyUrlFor_(blank);   out.blankName   = blank.getName(); }
    if (migrate) { out.migrateCopyUrl = copyUrlFor_(migrate); out.migrateName = migrate.getName(); }
  } catch (e) {
    console.error("scanTemplateFolder_ failed (folder access?): " + e.message);
  }
  return out;
}

/**
 * Builds a "Make a copy" URL for a Google editors file, matching its type so
 * the copy dialog opens correctly (Sheets/Docs/Slides). Defaults to Sheets.
 */
function copyUrlFor_(file) {
  const id = file.getId();
  let host = "spreadsheets";
  try {
    const mt = file.getMimeType();
    if (mt === "application/vnd.google-apps.document") host = "document";
    else if (mt === "application/vnd.google-apps.presentation") host = "presentation";
  } catch (e) { console.error(e); }
  return "https://docs.google.com/" + host + "/d/" + id + "/copy";
}


// ==========================================
// VERSION-TAB STORAGE (cols D:E, copy-surviving)
// ==========================================

/**
 * Mirrors the resolved links into the hidden "Version" tab, cells D1:E3.
 * Uses columns D:E so it never collides with the build date (A1), the
 * Auto-Migrate toggle (B2), or the saved-settings block (rows 4+, cols A:B).
 */
function saveTemplateLinksToVersionSheet_(links) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = (typeof getOrCreateVersionSheet_ === "function")
    ? getOrCreateVersionSheet_(ss) : ss.getSheetByName("Version");
  if (!sheet) return;
  sheet.getRange(1, 4, 3, 2).setValues([
    ["TEMPLATE LINKS (auto-managed — do not edit)", ""],
    ["BLANK copy",   links.blankCopyUrl   || ""],
    ["MIGRATE copy", links.migrateCopyUrl || ""]
  ]);
}

/**
 * Reads the inherited links back out of the Version tab (E2/E3). Returns null
 * when nothing usable is stored.
 */
function loadTemplateLinksFromVersionSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Version");
  if (!sheet) return null;
  try {
    const vals = sheet.getRange(2, 5, 2, 1).getValues(); // E2:E3
    const blank = String(vals[0][0] || "").trim();
    const migrate = String(vals[1][0] || "").trim();
    if (!blank && !migrate) return null;
    return { blankCopyUrl: blank, migrateCopyUrl: migrate, blankName: "", migrateName: "" };
  } catch (e) {
    console.error(e);
    return null;
  }
}
