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

// The Drive folder that holds the User Guide. Whatever the NEWEST file in here
// is (a PDF or a Google Doc) becomes the "Open User Guide" link — just replace
// the file to publish a new guide; no code edits.
const GUIDE_FOLDER_ID = "1zhH2cD-ecIPHKCLhWPEptXW0LQ2T12MB";

// How long a successful folder scan is cached (script cache) before re-scanning.
const TEMPLATE_LINKS_CACHE_SECONDS = 6 * 60 * 60; // 6 hours

// Fallback links if a folder can't be resolved — land the user in the folder
// rather than on a dead link.
const TEMPLATE_FOLDER_URL = "https://drive.google.com/drive/folders/" + TEMPLATE_FOLDER_ID;
const GUIDE_FOLDER_URL = "https://drive.google.com/drive/folders/" + GUIDE_FOLDER_ID;


/**
 * Public entry point for the dialogs (called via google.script.run).
 * Returns { blankCopyUrl, migrateCopyUrl, blankName, migrateName, folderUrl }.
 * Never throws — always returns usable (folder-fallback) URLs.
 */
function getTemplateLinks() {
  return getTemplateLinks_();
}

/**
 * Public entry point for the User Guide link (called via google.script.run).
 * Returns the URL of the newest file in the guide folder (PDF or Doc), or the
 * guide-folder URL as a fallback. Never throws.
 */
function getGuideLink() {
  return getGuideLink_();
}

/**
 * Force a fresh scan of BOTH folders (bypassing the cache) and re-store the
 * results. No UI — safe to call from time-based triggers (the nightly sync).
 */
function refreshResolvedLinks_() {
  const cache = CacheService.getScriptCache();
  try { cache.remove('TEMPLATE_LINKS_V1'); } catch (e) {}
  try { cache.remove('GUIDE_LINK_V1'); } catch (e) {}
  const links = getTemplateLinks_();
  const guide = getGuideLink_();
  return { links: links, guide: guide };
}

/**
 * Editor helper: force a fresh scan and show every resolved link in an alert.
 * (Not needed for normal use — the dialogs resolve on open and the nightly
 * sync refreshes the stored copies. Handy to verify links after swapping files.)
 */
function refreshTemplateLinks() {
  const r = refreshResolvedLinks_();
  const links = r.links;
  const msg =
    "BLANK  (" + (links.blankName || "not found") + "):\n" + (links.blankCopyUrl || "(unresolved)") +
    "\n\nMIGRATE (" + (links.migrateName || "not found") + "):\n" + (links.migrateCopyUrl || "(unresolved)") +
    "\n\nGUIDE:\n" + (r.guide || "(unresolved)");
  try {
    SpreadsheetApp.getUi().alert("🔗 Resolved Links", msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    console.log(msg); // No UI context (pure editor run) — the log still has it.
  }
  return r;
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
    let blank = null, migrate = null, blankKey = -1, migrateKey = -1;

    while (files.hasNext()) {
      const f = files.next();
      const name = f.getName();
      const upper = String(name).toUpperCase();
      const key = fileSortKey_(name, f); // newest-dated (or newest-updated) wins
      if (upper.indexOf('(BLANK)') !== -1 && key >= blankKey) { blank = f; blankKey = key; }
      if (upper.indexOf('(MIGRATE)') !== -1 && key >= migrateKey) { migrate = f; migrateKey = key; }
    }

    if (blank)   { out.blankCopyUrl   = copyUrlFor_(blank);   out.blankName   = blank.getName(); }
    if (migrate) { out.migrateCopyUrl = copyUrlFor_(migrate); out.migrateName = migrate.getName(); }
  } catch (e) {
    console.error("scanTemplateFolder_ failed (folder access?): " + e.message);
  }
  return out;
}

/**
 * Ranking key used to pick the newest template when several share a label.
 * Files are named like "Parent Coordinator Tracker 7/22/26 (BLANK)", so prefer
 * a MM/DD/YY (or M/D/YYYY) date parsed from the NAME; fall back to the file's
 * last-updated time when the name has no date.
 */
function fileSortKey_(name, file) {
  const m = String(name).match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (m) {
    let y = parseInt(m[3], 10);
    if (y < 100) y += 2000;
    const ms = new Date(y, parseInt(m[1], 10) - 1, parseInt(m[2], 10)).getTime();
    if (!isNaN(ms)) return ms;
  }
  try { return file.getLastUpdated().getTime(); } catch (e) { return 0; }
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


// ==========================================
// USER GUIDE LINK (newest file in the guide folder)
// ==========================================

/**
 * Cached, resilient resolver for the User Guide link. Tries a live scan of the
 * guide folder (newest file wins, PDF or Doc); caches it and mirrors it into
 * the Version tab (E4) so plain copies inherit it. Falls back to the inherited
 * value, then the guide-folder URL, so the link is never dead. Never throws —
 * safe even in the limited-auth simple onOpen (the scan just no-ops there and
 * the inherited/fallback value is used).
 */
function getGuideLink_() {
  const cache = CacheService.getScriptCache();
  try {
    const cached = cache.get('GUIDE_LINK_V1');
    if (cached) return cached;
  } catch (e) { console.error(e); }

  let url = scanGuideFolder_();
  if (url) {
    try { saveGuideLinkToVersionSheet_(url); } catch (e) { console.error(e); }
  } else {
    url = loadGuideLinkFromVersionSheet_() || GUIDE_FOLDER_URL;
  }

  try { cache.put('GUIDE_LINK_V1', url, TEMPLATE_LINKS_CACHE_SECONDS); } catch (e) {}
  return url;
}

/**
 * Live-scans the guide folder and returns the open URL of the newest file
 * (a PDF's Drive view URL or a Doc's edit URL, whatever Drive reports). Returns
 * "" on any failure (no access / offline / empty folder).
 */
function scanGuideFolder_() {
  try {
    const folder = DriveApp.getFolderById(GUIDE_FOLDER_ID);
    const files = folder.getFiles();
    let newest = null, newestKey = -1;
    while (files.hasNext()) {
      const f = files.next();
      const key = fileSortKey_(f.getName(), f);
      if (key >= newestKey) { newest = f; newestKey = key; }
    }
    return newest ? newest.getUrl() : "";
  } catch (e) {
    console.error("scanGuideFolder_ failed (folder access?): " + e.message);
    return "";
  }
}

/** Mirrors the guide link into the Version tab (D4 label / E4 value). */
function saveGuideLinkToVersionSheet_(url) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = (typeof getOrCreateVersionSheet_ === "function")
    ? getOrCreateVersionSheet_(ss) : ss.getSheetByName("Version");
  if (!sheet) return;
  sheet.getRange(4, 4, 1, 2).setValues([["GUIDE link", url || ""]]);
}

/** Reads the inherited guide link back out of the Version tab (E4). */
function loadGuideLinkFromVersionSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Version");
  if (!sheet) return "";
  try {
    return String(sheet.getRange(4, 5).getValue() || "").trim(); // E4
  } catch (e) {
    console.error(e);
    return "";
  }
}
