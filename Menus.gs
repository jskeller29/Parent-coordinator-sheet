// ======================================================================
// FILE: Menus.gs
// PURPOSE: Menus, first-time setup, and the server side of every dialog:
// Getting Started (welcome + checklist), Settings, Student Overrides (+ report),
// End of Year Rollover, Migration Wizard, Version Check, the data update
// reminder, and the template / User Guide links.
// ======================================================================

// ######################################################################
// ######################################################################
// ##                                                                  ##
// ##   📌  YOUR LINKS  -  paste document and video links HERE         ##
// ##                                                                  ##
// ######################################################################
// ######################################################################
//
// HOW TO ADD OR CHANGE A LINK
//   1. Open the video or document and copy its link
//      (YouTube: Share → Copy.  Google Doc / Drive file: Share → Copy link).
//      Set it to "Anyone with the link can view", or other schools can't open it.
//   2. Paste it between the two quote marks "" on the right line below.
//      Keep the quote marks and the comma at the end of the line.
//   3. Save (Ctrl+S, Mac ⌘+S).
//   4. Paste this Menus file into the master template too (and any sheet that
//      should get the new link), then run 🐞 Debug → Reset to Display Mode on the
//      master template: the 👋 Welcome tab is only made again by a reset.
//
// LEFT EMPTY ("")?  A video button is simply not shown.
//                   A document button opens the User Guide instead.
// To see every link at once (set or not, and where it shows):
//   🐞 Debug → 🔗 My Links
//
// THREE LINKS ARE NOT IN THIS FILE: they go in the VERSION TRACKER (spreadsheet below),
// one row per release, so changing them needs no code and reaches every sheet:
//   column D "Blank Copy"  the blank template's /copy link (Make a new copy, Rollover,
//                          "Get the latest version" in 🔔 Check for Updates)
//   column E "Example"     a filled-in example tracker (made-up students): a "See an
//                          example tracker" button on the Welcome tab, the welcome
//                          pop-up, 📖 Open User Guide and Getting Started (empty = no button)
//   column F "Document"    the User Guide
// The newest dated row with a link in that column counts (leave it blank in a release
// that keeps the same link). Columns are found by their header name. Not filled in /
// the tracker can't be read = the folders below are used instead (the backup).
// ----------------------------------------------------------------------

// 🎬 VIDEOS ------------------------------------------------------------
// 📄 DOCUMENTS (empty = the User Guide opens instead) ------------------
const GUIDE_LINKS = {
  // 🎬 Setting up a new copy: the Welcome tab, the permissions, Initial Setup.
  //    Shows on: the 👋 Welcome tab ("Watch the setup video") and the
  //    Getting Started checklist (Help at the bottom).
  setupVideo: "https://www.youtube.com/watch?v=4WGkEIJrlL8",

  // 🎬 How to download your data from ATS, Parent Square and NYCSA.
  //    Shows on: the welcome pop-up after Initial Setup ("Watch the video") and
  //    the Getting Started checklist (step 1 and Help).
  downloadDataVideo: "https://www.youtube.com/watch?v=4WGkEIJrlL8",

  // 📄 How to download / update your data, step by step (a Google Doc).
  //    Shows on: the welcome pop-up ("How to download your data") and
  //    Getting Started step 1.
  downloadDataDoc: "https://docs.google.com/document/d/12GJCUgiyQean7P3qb8GhulXDy73OLGH3/edit?usp=sharing",

  // 📄 A short getting-started guide, if you write one apart from the User Guide.
  //    Shows on: the welcome pop-up ("Read the User Guide") and Getting Started Help.
  guideDoc: "https://docs.google.com/document/d/12GJCUgiyQean7P3qb8GhulXDy73OLGH3/edit?usp=sharing",

  // 📄 What Build Sheets does.           Shows on: Getting Started step 2.
  buildDoc: "https://docs.google.com/document/d/12GJCUgiyQean7P3qb8GhulXDy73OLGH3/edit#bookmark=id.lxd64bcidr3o",

  // 📄 What each setting does.           Shows on: Getting Started step 3.
  settingsDoc: "https://docs.google.com/document/d/12GJCUgiyQean7P3qb8GhulXDy73OLGH3/edit#bookmark=id.suv42utnjrnr",

  // 📄 How to log a parent contact.       Shows on: Getting Started step 4.
  contactLogDoc: "https://docs.google.com/document/d/12GJCUgiyQean7P3qb8GhulXDy73OLGH3/edit#bookmark=id.gymia0d491d6"
};

// 🗂️ FOLDERS, FORM AND TRACKER (already set; change only if you move them) ----
//    A Drive folder ID is the long part of the folder's link after /folders/
//    (https://drive.google.com/drive/folders/THIS_PART).

// 📖 The User Guide FOLDER (the BACKUP): when the Version Tracker has no
//    "Document" link, the NEWEST file in it is the User Guide link everywhere
//    (📖 Open User Guide, the Welcome tab, Getting Started).
const GUIDE_FOLDER_ID = "1zhH2cD-ecIPHKCLhWPEptXW0LQ2T12MB";

// 📋 The template FOLDER (the BACKUP): when the Version Tracker has no "Blank
//    Copy" link, the newest file with "(BLANK)" in its name is the "make a copy"
//    link (End of Year Rollover, 📖 Open User Guide).
const TEMPLATE_FOLDER_ID = "1goKFrxDJOSAPxp1HUXETX2dns5c6HXk2";

// 🐞 The bug report form (🐞 Report a Bug, the Welcome tab, Getting Started Help).
var BUG_REPORT_FORM_URL = "https://forms.gle/5dXR2me3jt37aGhQ9";

// 🔔 The Version Tracker spreadsheet ID. 🔔 Check for Updates and the three links
//    above (Blank Copy / Example / Document) read its FIRST tab.
const VERSION_TRACKER_ID = "1nnJXnYEGyISvQtvsp2M4D1EQT2cHhNWkBXMG1Ubbljc";

// (The User Guide's own "Make a copy" buttons are in the docs folder of the
//  repository, not here: ask Claude to change them.)
//
// The hidden Backend tab's "LINKS" cells (D2:E6) are the tracker's OWN saved copy of
// the links it found (Version Tracker, else the folders), used only when a sheet can't
// read either. The tracker rewrites them by itself: don't paste links there (deleting
// them is harmless).

// ######################################################################
// ##   END OF YOUR LINKS                                              ##
// ######################################################################

// When runInitialSetupWrapper() calls onOpen() to refresh the menu, we don't
// want onOpen's own "auto-open Migration Wizard" block to fire too — the
// wrapper opens it authoritatively right after. This global (shared across all
// .gs files in the one Apps Script scope) suppresses the duplicate for that
// single synchronous call. onOpen's block still fires on genuine reopens.
var IS_RUNNING_SETUP = false;

// Every tab the ⚙️ Settings "Show these tabs" list can show / hide: setting key,
// the tab name(s), and whether it is shown by default (owner's pick 2026-11-05:
// Contact Log, Combined Contact Tracking, PCAR and Directory). A key that was never
// saved uses its default. The first 4 keys are the old ones (phoneContacts ...
// sendOut); their saved values carry over. Backend is never listed (always hidden).
var TAB_VISIBILITY = [
  { key: 'showContactLog',     sheets: ["Contact Log"],               def: true  },
  { key: 'showCombined',       sheets: ["Combined Contact Tracking"], def: true  },
  { key: 'showPcar',           sheets: ["PCAR"],                      def: true  },
  { key: 'showDirectory',      sheets: ["Directory"],                 def: true  },
  { key: 'showRawData',        sheets: ["Raw Data", "RAW Data"],      def: false },
  { key: 'showMasterTable',    sheets: ["Student List", "Master Table"], def: false }, // was "Master Table" (renamed on first use)
  { key: 'phoneContacts',      sheets: ["Phone Contacts"],            def: false },
  { key: 'parentsDivided',     sheets: ["Parents Divided"],           def: false },
  { key: 'notesTab',           sheets: ["Notes"],                     def: false },
  { key: 'showEvents',         sheets: ["Events"],                    def: false },
  { key: 'sendOut',            sheets: ["Send Out"],                  def: false },
  { key: 'showOverrideReport', sheets: ["Students Override Report"],  def: false },
  { key: 'showNewStudents',    sheets: ["New Students"],              def: false }
];

// Hour of the nightly rebuild / phone sync trigger (22 = 10 PM in the script's time
// zone; Apps Script runs it some time in that hour). Initial Setup and every
// Settings save make the trigger with this hour (owner's pick 2026-11-05).
var NIGHTLY_SYNC_HOUR = 22;

// The tabs shown by default. Everything else starts hidden; the user's saved
// settings reveal the optional tabs later via updateVisualSettings_.
var DEFAULT_VISIBLE_SHEETS = TAB_VISIBILITY
  .filter(function(t) { return t.def; })
  .reduce(function(all, t) { return all.concat(t.sheets); }, []);

/**
 * Shows only the default-visible tabs and hides every other sheet. Shown tabs
 * first: on a new copy the 👋 Welcome tab is the only visible one, and Sheets
 * refuses to hide the last visible tab.
 */
function applyDefaultSheetVisibility_(ss) {
  const sheets = ss.getSheets();
  const isDefault = function(sheet) { return DEFAULT_VISIBLE_SHEETS.indexOf(sheet.getName()) !== -1; };
  sheets.forEach(function(sheet) { if (isDefault(sheet)) sheet.showSheet(); });
  sheets.forEach(function(sheet) { if (!isDefault(sheet)) sheet.hideSheet(); });
}

/**
 * Auto-Migrate mode is controlled by the hidden Backend tab, cell B3:
 * "Yes" turns it on (upgrade-template copies), anything else is off
 * (new-user copies). The cell travels with every copy of the sheet, and
 * reading the bound spreadsheet's own cell is allowed even in the simple
 * onOpen trigger. Set it with the toggleAutoMigratePopup() dev tool or by
 * typing directly into B3. Read-only (runs in the simple onOpen): falls back to
 * an older copy's separate "Version" tab (B2) without migrating anything.
 */
function isAutoMigrateOn_(ss) {
  try {
    ss = ss || SpreadsheetApp.getActiveSpreadsheet(); // ss: another spreadsheet (a copy being made into a template)
    const backend = ss.getSheetByName(BACKEND_SHEET_NAME);
    const legacy = ss.getSheetByName("Version");
    const cell = backend ? backend.getRange(BACKEND_VERSION.MIGRATE_TOGGLE)
               : legacy ? legacy.getRange("B2") : null;
    if (!cell) return false;
    return String(cell.getValue()).trim().toLowerCase() === "yes";
  } catch (e) {
    console.error(e);
    return false;
  }
}

/**
 * Automatically creates the necessary installable triggers.
 * It deletes old triggers first to prevent duplicate firings.
 */
function setupAllInstallableTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!ss) {
    console.error("Could not find active spreadsheet. Ensure you are running this from the bound script.");
    return;
  }

  // 1. Delete any existing triggers to prevent accidental duplicates
  try {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(t => {
      const handlerName = t.getHandlerFunction();
      // 🆕 "checkVersionOnOpen" added to the cleanup list
      if (handlerName === "installedOnEdit" || handlerName === "nightlyMasterSync" || handlerName === "checkVersionOnOpen" ||
          handlerName === "installedOnChange") {
        ScriptApp.deleteTrigger(t);
      }
    });
  } catch (e) {
    console.error("Error clearing old triggers: " + e.message);
  }

  // 2. Create the Installable OnEdit Trigger (for instant Contact Log updates) 
  try {
    ScriptApp.newTrigger("installedOnEdit") 
      .forSpreadsheet(ss) 
      .onEdit() 
      .create();
  } catch (e) {
    console.error("Error creating OnEdit trigger: " + e.message);
  }

  // 3. Create the Nightly Sync Trigger unconditionally (10 PM every day, NIGHTLY_SYNC_HOUR);
  //    the sheet remembers it as its one nightly trigger (Build.gs)
  try {
    const nightly = ScriptApp.newTrigger("nightlyMasterSync")
      .timeBased()
      .atHour(NIGHTLY_SYNC_HOUR)
      .everyDays(1)
      .create();
    rememberNightlyTrigger_(PropertiesService.getDocumentProperties(), nightly.getUniqueId());
  } catch (e) {
    console.error("Error creating Nightly Sync trigger: " + e.message);
  }

  // =======================================================
  // 🆕 4. VERSION UPDATE CHECKER (Installable OnOpen)
  // Must be INSTALLABLE: the simple onOpen() runs in limited-auth
  // mode and is not allowed to open the external tracker sheet.
  // =======================================================
  try {
    ScriptApp.newTrigger("checkVersionOnOpen")
      .forSpreadsheet(ss)
      .onOpen()
      .create();
  } catch (e) {
    console.error("Error creating Version Check trigger: " + e.message);
  }

  // 5. Change trigger: a deleted Contact Log row takes its note out of the notes (Edits.gs)
  ensureChangeTrigger_(ss);

  ss.toast("Setup 100% Complete! The script is fully authorized and all automatic triggers are running.", "Success", 8);
}

/**
 * Adds the installedOnChange trigger when it is missing. Runs at setup and on
 * the first open after this update (checkVersionOnOpen), so sheets set up
 * before it existed get it without running setup again.
 * fromOpen: only once per sheet (doc property CHANGE_TRIGGER_ADDED), because
 * triggers belong to whoever made them: a second person opening the sheet would
 * otherwise add a second one and every delete would run twice.
 */
function ensureChangeTrigger_(ss, fromOpen) {
  try {
    const props = PropertiesService.getDocumentProperties();
    if (fromOpen && props.getProperty('CHANGE_TRIGGER_ADDED') === 'true') return;
    ss = ss || SpreadsheetApp.getActiveSpreadsheet();
    const has = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === "installedOnChange");
    if (!has) ScriptApp.newTrigger("installedOnChange").forSpreadsheet(ss).onChange().create();
    props.setProperty('CHANGE_TRIGGER_ADDED', 'true');
  } catch (e) {
    console.error("Error creating OnChange trigger: " + e.message);
  }
}

// ==========================================
// 1. MENUS (simple onOpen trigger; also re-run after settings / setup changes)
// ==========================================
function onOpen() {
  return runAction_("Open (menus)", () => buildMenus_());
}

function buildMenus_() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();

  try {
    applyQBlackTextFilter_();
  } catch(e) { console.error(e); }

  // CLONE DETECTOR
  const currentSheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const savedSheetId = props.getProperty('ORIGINAL_SHEET_ID');

  if (savedSheetId && savedSheetId !== currentSheetId) {
    props.deleteProperty('SETUP_COMPLETE');
    props.deleteProperty('WELCOME_DIALOG_SHOWN');
    props.deleteProperty('MISSED_SYNC_MSG');
    props.deleteProperty(NIGHTLY_FAILED_KEY_);
    
    // 🆕 Fresh copies start with clean version-check memory so any
    // dismissals/preferences from the template never leak into copies.
    props.deleteProperty('VERSION_LAST_SEEN');
    props.deleteProperty('VERSION_LAST_SEEN_DATE');
    props.deleteProperty('VERSION_DISMISS_FOREVER');
    props.deleteProperty('VERSION_MAJORS_ONLY');
    props.deleteProperty('VERSION_LAST_CHECK');
    // ℹ️ The Auto-Migrate toggle now lives in the Backend tab cell B3, not
    // a document property, so it travels with copies automatically — nothing
    // to preserve here.
  }

  // LAST NIGHT'S REBUILD FAILED (Build.gs nightlyMasterSyncRun_): said first; a sync
  // message then waits for the next open (only one toast shows at a time)
  const nightlyFailed = props.getProperty(NIGHTLY_FAILED_KEY_);
  if (nightlyFailed) {
    SpreadsheetApp.getActiveSpreadsheet().toast(nightlyFailed, "⚠️ Last night's rebuild", 20);
    props.deleteProperty(NIGHTLY_FAILED_KEY_);
  }

  // OFFLINE SYNC DETECTOR
  const missedSync = nightlyFailed ? null : props.getProperty('MISSED_SYNC_MSG');
  if (missedSync) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      missedSync, 
      "🔄 Synced Since Last Opened!", 
      15
    );
    props.deleteProperty('MISSED_SYNC_MSG'); 
  }

  const isSetupDone = props.getProperty('SETUP_COMPLETE');

  if (isSetupDone !== 'true') {
    const initMenu = ui.createMenu("🚀 App Menu");
    initMenu.addItem('🚨 Initial Setup (Run Once)', 'runInitialSetupWrapper'); 
    initMenu.addToUi(); 

    // Owner's sheet only: a fresh copy of it (setup not run there yet) still gets the 🐞 Debug menu.
    // No-op on every other sheet.
    const ownerBeforeSetup = getOwnerHooks_();
    if (ownerBeforeSetup) {
      try { ownerBeforeSetup.addMenus(ui); } catch (e) { console.error(e); }
    }

    // New copy: the 👋 Welcome tab is the only tab shown until Initial Setup has run
    try { showWelcomeTabOnly_(SpreadsheetApp.getActiveSpreadsheet()); } catch (e) { console.error("Welcome tab: " + e.message); }

    // A short pointer only: the Welcome tab has the steps (it was a pop-up on top of it; simplified 2026-11-12).
    // Upgrade-template copies: Initial Setup then opens the Migration Wizard (see runInitialSetupWrapper).
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "Click 🚀 App Menu (top of the screen) → 🚨 Initial Setup (Run Once)" +
          (isAutoMigrateOn_() ? " to bring over your old data." : " to get started."),
        isAutoMigrateOn_() ? "🎉 Welcome to the newest version" : "👋 Welcome to the PC Tracker", 30);
    } catch(e) { console.error(e); }

    return;
  }

  // ==============================================================
  // WHAT SHOWS AFTER SETUP IS COMPLETE:
  // ==============================================================
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // The 👋 Welcome tab Initial Setup hid is deleted now (not in the setup call itself: see WELCOME TAB)
  if (!IS_RUNNING_SETUP) {
    try { deleteWelcomeTab_(ss); } catch (e) { console.error("Welcome tab: " + e.message); }
  }

  // Check if the sheets have ACTUALLY been built (a student under the Student List header)
  const isBuilt = studentListIsBuilt_(ss);

 if (!isBuilt) {
    // 🛑 INTERMEDIATE STATE: Setup done, but sheets haven't been built yet
    props.setProperty('FULL_MENU_UNLOCKED', 'false'); // Mark as locked
    // Using the EXACT SAME "🚀 App Menu" name allows us to overwrite the Setup button 
    // instantly without forcing the user to refresh the page!
    // New-user order (2026-11-12): the checklist, its first two steps, then the rest
    const getStartedMenu = ui.createMenu("🚀 App Menu"); 
    addGuideMenuItems_(getStartedMenu);          // ✅ Getting Started (below)
    addInputDataModeMenuItem_(getStartedMenu);   // Input Data.gs
    getStartedMenu.addItem("🏗️ Build Sheets Only", "runBuildOnly"); 
    getStartedMenu.addSeparator();
    getStartedMenu.addItem("🔄 Import Old Data", "openMigrationWizard");
    getStartedMenu.addItem("📖 Open User Guide", "showWelcomeDialog");
    getStartedMenu.addItem("🐞 Report a Bug", "reportABug");
    getStartedMenu.addItem("🔔 Check for Updates", "manualCheckForUpdates"); // 🆕
    if (typeof addContactLogDebugMenuItem_ === "function") addContactLogDebugMenuItem_(getStartedMenu); // owner test file z_ContactLogDebug.gs
    getStartedMenu.addToUi();

    // =======================================================
    // 🆕 AUTO-OPEN MIGRATION WIZARD (Upgrade-template copies)
    // Fires on every open until the sheets are built. Controlled by the
    // hidden Backend tab, cell B3 = "Yes" (see isAutoMigrateOn_).
    // =======================================================
    if (isAutoMigrateOn_() && !IS_RUNNING_SETUP) {
      try { openMigrationWizard(); } catch(e) { console.error(e); }
    }
  } else {
    // ✅ FULLY POPULATED STATE: Data exists and is built, unlock all menus!
    props.setProperty('FULL_MENU_UNLOCKED', 'true'); // Mark as explicitly unlocked
    const syncToPhones = props.getProperty('syncToPhones') === 'true';

    const menu = ui.createMenu("🚀 App Menu");
    // Only show Full Sync if they have Phone Syncing enabled!
    if (syncToPhones) {
      menu.addItem("🧱 Full Sync (Build + Push to Google Contacts)", "openProgressBarSidebar");
    }
    
    menu.addItem("🏗️ Build Sheets Only (No Push to Google Contacts)", "runBuildOnly"); 
    addInputDataModeMenuItem_(menu);   // Input Data.gs
    menu.addSeparator(); 
    addGuideMenuItems_(menu);          // ✅ Getting Started (below)
    menu.addItem("⚙️ Settings", "openSettingsDialog"); 
    menu.addItem("📖 Open User Guide", "showWelcomeDialog");
    menu.addItem("🐞 Report a Bug", "reportABug");
    menu.addItem("🔔 Check for Updates", "manualCheckForUpdates"); // 🆕
    if (typeof addContactLogDebugMenuItem_ === "function") addContactLogDebugMenuItem_(menu); // owner test file z_ContactLogDebug.gs
    // Sheet Lock.gs: turn the baby-proofing off for layout changes, then back on
    menu.addItem(isSheetLockOn_() ? "🔓 Unlock Sheet (change headers / layout)" : "🔒 Lock Sheet", "toggleSheetLock");
    
    // --- NEW: END OF YEAR ROLLOVER ---
    menu.addSeparator();
    menu.addItem("🔄 End of Year Rollover", "openRolloverWizard");
    // ---------------------------------
    
    menu.addToUi();

    // --- NEW: STUDENT OVERRIDES MENU --- 
    ui.createMenu('👤 Student Overrides') 
    .addItem('✨ New / Edit Student', 'openNewStudentDialog') 
    .addItem('🗂️ Manage Overrides', 'openManageOverridesDialog') 
    .addItem('📄 Override Report', 'buildOverrideReport')
    .addToUi();

    // 2. The Events Menu
    ui.createMenu('📆 Events')
      .addItem('📋 New Event', 'createNewEventTable')
      .addToUi();

    // 3. Contact Log menu — optional module (Z#ContactLog.gs). Guarded so the
    //    app works unchanged if that file isn't included.
    if (typeof buildContactLogMenu_ === "function") buildContactLogMenu_();
  }

  // Owner's sheet only (z_MINE.gs): the 🐞 Debug menu
  const owner = getOwnerHooks_();
  if (owner) owner.addMenus(ui);

  // --- NEW: AUTO-REPAIR CONTACT LOG FILTER CHECKBOX ---
  // Shifted to the bottom so it doesn't block the menus from appearing instantly
  try {
    const contactLog = ss.getSheetByName("Contact Log");
    if (contactLog) {
      // Find END in Column E (5). MUST use getValues() on Open because getDisplayValues 
      // fails if the spreadsheet format UI hasn't fully rendered yet!
      const eValues = contactLog.getRange("E:E").getValues();
      for (let i = 0; i < eValues.length; i++) {
        if (String(eValues[i][0]).trim().toUpperCase() === "END") {
          const endRow = i + 1;
          // Force Column O (15) to TRUE and flush it to the screen instantly
          contactLog.getRange(endRow, 15).setValue(true);
          SpreadsheetApp.flush(); 
          break;
        }
      }
    }
  } catch(err) {
    // Fail silently if there's an issue so it doesn't break the rest of the startup
  }
  // ----------------------------------------------------

  // Background Sync Lockout Check
  try {
    const cache = CacheService.getScriptCache();
    const stateStr = cache.get('SYNC_UI_STATE');
    
    if (stateStr) {
      const state = JSON.parse(stateStr);
      if (state.phase === 'building' || state.phase === 'syncing' || state.phase === 'paused' || state.phase === 'quota') {
        SpreadsheetApp.getUi().alert(
          "📡 Background Sync in Progress", 
          "The system is currently running a background sync.\n\nPlease do not enter new data until it finishes to prevent saving errors. Click '🚀 App Menu' -> '🧱 Full Sync' to view live progress.", 
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      }
    }
  } catch(e) { console.error(e); }
}

// ==========================================
// WELCOME / TUTORIAL DIALOG HTML
// ==========================================
function showWelcomeDialog() {
  // Resolve the shareable template links + the User Guide link from their Drive
  // folders (see TEMPLATE & GUIDE LINKS below). Never throws; falls back to folder URLs.
  const links = getTemplateLinks_();
  const guideUrl = getGuideLink_();

  const htmlContent = include_('Styles') + `
    <style>
      .note-dlg { padding: 8px 12px 12px; display: flex; flex-direction: column; gap: 14px; }
      .note-dlg p { line-height: 1.5; }
      .note-dlg .foot { display: flex; justify-content: flex-end; gap: 8px; flex-wrap: wrap; }
      .note-dlg .sub { border-top: 1px solid var(--line); padding-top: 14px; display: flex; flex-direction: column; gap: 10px; }
    </style>
    <div class="note-dlg">
      <h1>Welcome to the PC Tracker</h1>
      <p>This system builds your directory, tracks your parent interactions, and compiles your PCAR data for you.</p>
      <p><b>Next steps:</b> open <b>🚀 App Menu → ✅ Getting Started</b>. Its checklist walks you through loading your student data (📥 Input Data Mode), building your sheets, your settings and your first parent contact.</p>
      <p>Moving from last year's tracker? Use <b>🚀 App Menu → 🔄 Import Old Data</b>.</p>
      <div class="foot">
        ${links.exampleUrl ? `<a class="btn sec" href="${links.exampleUrl}" target="_blank" rel="noopener">See an example tracker</a>` : ``}
        <a class="btn pri" href="${guideUrl}" target="_blank" rel="noopener">Read the User Guide</a>
      </div>
      <div class="sub">
        <p class="help" style="margin:0">Need a fresh copy (for a colleague, or next year)? Its setup asks whether to start fresh or bring over an old tracker.</p>
        <div class="foot" style="justify-content:flex-start">
          <a class="btn sec" href="${links.blankCopyUrl}" target="_blank" rel="noopener">Make a new copy</a>
        </div>
      </div>
      <p class="hint">You can open this again any time from <b>🚀 App Menu → 📖 Open User Guide</b>.</p>
    </div>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(460)
    .setHeight(480);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "PC Tracker User Guide");
}

// ======================================================================
// GETTING STARTED (+ GettingStartedSidebar.html; was the trial file Z#Guide.gs)
// "✅ Getting Started" guide for new parent coordinators:
//   - Welcome pop-up right after Initial Setup (new-user copies): "Start with
//     Input Data Mode", plus links to the guide document / video.
//   - Getting Started checklist (sidebar): load data (Input Data Mode) -> build ->
//     settings -> first parent contact. Each step ticks itself from the sheet.
//     Opens by itself on every open until all four are done, unless
//     "Open this by itself" is unticked. Closing it is always fine.
//   - "This tab" panel in the same sidebar: what the tab you're on is for, and
//     which tabs are hidden and where to turn them on.
// Called from: buildMenus_ (addGuideMenuItems_, both 🚀 App Menus),
// runInitialSetupWrapper (showGuideWelcome_, instead of showWelcomeDialog;
// "📖 Open User Guide" still opens showWelcomeDialog) and checkVersionOnOpen
// (guideOnOpen_, installable on open).
// State (document properties; they don't travel with copies, so every new copy
// starts fresh): ZG_AUTO_OPEN ("off"), ZG_SETTINGS_SEEN ("true").
// ======================================================================

// LINKS: GUIDE_LINKS is in the 📌 YOUR LINKS box at the top of this file.
// An empty DOCUMENT link uses the User Guide instead (the newest file in the guide
// folder, getGuideLink_), so every document button works. An empty VIDEO link is not shown.

/**
 * Every link in the 📌 YOUR LINKS box, what it is, where it shows and its value now
 * ("" = not set). For 🐞 Debug → 🔗 My Links (z_MINE.gs).
 */
function linkList_() {
  const L = GUIDE_LINKS;
  let guide = "";
  try { guide = getGuideLink_(); } catch (e) { guide = ""; }
  const t = trackerLinks_();
  const col = (key, name) => t.fromTracker
    ? (t.cols && t.cols[key] ? "Version Tracker, column " + t.cols[key] : 'Version Tracker: no column named "' + name + '"')
    : "Version Tracker (couldn't read it" + (t.error ? ": " + t.error : "") + ")";
  const trackerRow = (key, name, what, where, empty) =>
    ({ kind: "tracker", name: col(key, name), what: what, where: where, url: t[key] || "", empty: empty });
  return [
    trackerRow("blank", "Blank Copy", "Blank copy (the /copy link people make their tracker from)",
      "📖 Open User Guide (Make a new copy), 🔔 Check for Updates, End of Year Rollover", "Not set: the template folder's newest (BLANK) file"),
    trackerRow("example", "Example", "Example tracker (made-up students)",
      "📖 Open User Guide, Welcome pop-up, 👋 Welcome tab, Getting Started (Help)", "Not set: no button"),
    trackerRow("guide", "Document", "User Guide",
      "📖 Open User Guide, 👋 Welcome tab, Getting Started", "Not set: the guide folder's newest file"),
    { kind: "video", name: "setupVideo", what: "Setting up a new copy", where: "👋 Welcome tab, Getting Started (Help)", url: L.setupVideo },
    { kind: "video", name: "downloadDataVideo", what: "How to download your data", where: "Welcome pop-up, Getting Started (step 1, Help)", url: L.downloadDataVideo },
    { kind: "doc", name: "downloadDataDoc", what: "How to download / update your data", where: "Welcome pop-up, Getting Started (step 1)", url: L.downloadDataDoc },
    { kind: "doc", name: "guideDoc", what: "Getting-started guide", where: "Welcome pop-up, Getting Started (Help)", url: L.guideDoc },
    { kind: "doc", name: "buildDoc", what: "What Build Sheets does", where: "Getting Started (step 2)", url: L.buildDoc },
    { kind: "doc", name: "settingsDoc", what: "What each setting does", where: "Getting Started (step 3)", url: L.settingsDoc },
    { kind: "doc", name: "contactLogDoc", what: "How to log a parent contact", where: "Getting Started (step 4)", url: L.contactLogDoc },
    { kind: "folder", name: "GUIDE_FOLDER_ID", what: "User Guide folder (backup: its newest file, when the Version Tracker has no Document link)",
      where: "Same places as the User Guide", url: GUIDE_FOLDER_URL, now: t.guide ? "" : guide, files: guideFolderFiles_().slice(0, 5) },
    { kind: "folder", name: "TEMPLATE_FOLDER_ID", what: "Template folder (backup: its newest (BLANK) file, when the Version Tracker has no Blank Copy link)",
      where: "Same places as the blank copy", url: TEMPLATE_FOLDER_URL },
    { kind: "form", name: "BUG_REPORT_FORM_URL", what: "Bug report form", where: "🐞 Report a Bug, Welcome tab, Getting Started", url: BUG_REPORT_FORM_URL },
    { kind: "sheet", name: "VERSION_TRACKER_ID", what: "Version Tracker (versions + the Blank Copy / Example / Document links)", where: "🔔 Check for Updates, the 3 links above", url: "https://docs.google.com/spreadsheets/d/" + VERSION_TRACKER_ID }
  ];
}

/** GUIDE_LINKS with the fallbacks filled in (videos stay empty when not set). */
function zg_links_() {
  const guide = zg_safeGuideLink_();
  return {
    setupVideo: GUIDE_LINKS.setupVideo,
    downloadDataVideo: GUIDE_LINKS.downloadDataVideo,
    guideDoc: GUIDE_LINKS.guideDoc || guide,
    downloadDataDoc: GUIDE_LINKS.downloadDataDoc || guide,
    buildDoc: GUIDE_LINKS.buildDoc || guide,
    settingsDoc: GUIDE_LINKS.settingsDoc || guide,
    contactLogDoc: GUIDE_LINKS.contactLogDoc || guide,
    example: zg_safeExampleLink_(),
    bugReport: BUG_REPORT_FORM_URL
  };
}

/**
 * What each tab is for, for the "This tab" panel. type: "data" | "type" (you
 * type here) | "report" | "built" (filled in for you). settings = shown or
 * hidden in ⚙️ Settings → Tabs & columns.
 */
const ZG_TABS = [
  { name: "Raw Data", alt: "RAW Data", type: "data", settings: true,
    what: "Where your student data lives: ATS (columns A to AE), Parent Square (AG to AN), NYCSA (AP to BF), and your sites and classes (BH to BI). Row 1 of each part shows when it was last updated. Input Data Mode puts the data in for you." },
  { name: "Contact Log", type: "type", settings: true,
    what: "Your daily log of every parent contact. Type the parent's name in column C and the student and OSIS fill in, along with the date, the # and the follow-up." },
  { name: "Notes", type: "type", settings: true,
    what: "Notes about a student that you always want to see first. They show at the top of that student's notes in Combined Contact Tracking." },
  { name: "Events", type: "type", settings: true,
    what: "Attendance for your events and workshops. Make a table with 📆 Events → 📋 New Event, then add the parents who came." },
  { name: "Send Out", type: "type", settings: true,
    what: "Messages and flyers you sent to families and how many you reached. They count toward your PCAR." },
  { name: "PCAR", type: "report", settings: true,
    what: "Your PCAR report, filled in for you from the other tabs. Only H2:I2 are yours to change." },
  { name: "Student List", type: "built", settings: true,
    what: "One row per student with guardians, siblings, site, class and status. Rebuilt on every build, so don't type here." },
  { name: "Combined Contact Tracking", type: "built", settings: true,
    what: "Every student's notes in one place: Contact Log entries, events and Notes-tab notes, newest first." },
  { name: "Parents Divided", type: "built", settings: true,
    what: "One row per guardian, with contact details and Parents Square / NYCSA sign-up." },
  { name: "Directory", type: "built", settings: true,
    what: "Your family directory, from ATS only: each student, OSIS, site, class and how to reach their parents." },
  { name: "Phone Contacts", type: "built", settings: true,
    what: "The contact list your phone sync sends to Google Contacts (⚙️ Settings → Contacts & sync)." },
  { name: "Students Override Report", type: "built", settings: true,
    what: "The student changes you made with 👤 Student Overrides." }
];

// ----------------------------------------------------------------------
// CALLED FROM THE MENUS / SETUP / OPEN TRIGGER
// ----------------------------------------------------------------------

/** 🚀 App Menu item (both the getting-started and the full menu). */
function addGuideMenuItems_(menu) {
  menu.addItem("✅ Getting Started", "openGuideChecklist");
}

/**
 * Right after Initial Setup: the checklist, then the welcome pop-up on top. Its first
 * screen asks "Is this your first PC Tracker?" (owner's request 2026-11-12: one template
 * for everyone instead of a separate MIGRATE copy): "Start fresh" shows the four steps,
 * "Bring over last year's tracker" opens the Migration Wizard (which lists the trackers
 * it finds in the user's Drive, ma_findOldTrackers). An older upgrade copy (Backend B3 =
 * Yes) still goes straight to the wizard (runInitialSetupWrapper).
 */
function showGuideWelcome_() {
  openGuideChecklist();
  const links = zg_links_();
  const link = (url, label) => url ? `<a class="btn sec" href="${url}" target="_blank" rel="noopener">${label}</a>` : ``;

  const html = include_('Styles') + `
    <style>
      .note-dlg { padding: 8px 12px 12px; display: flex; flex-direction: column; gap: 14px; }
      .note-dlg p { line-height: 1.5; }
      .note-dlg ol { margin: 0; padding-left: 20px; display: flex; flex-direction: column; gap: 4px; }
      .note-dlg .foot { display: flex; justify-content: flex-end; gap: 8px; flex-wrap: wrap; }
      .note-dlg .sub { border-top: 1px solid var(--line); padding-top: 14px; display: flex; flex-direction: column; gap: 10px; }
      .note-dlg .sub .foot { justify-content: flex-start; }
      .choice { display: flex; flex-direction: column; gap: 10px; }
      .opt { text-align: left; border: 1px solid var(--line); border-radius: var(--radius); background: #fff; padding: 12px 14px;
             cursor: pointer; font: inherit; color: var(--ink); display: flex; flex-direction: column; gap: 3px; }
      .opt:hover { border-color: var(--accent); background: var(--panel); }
      .opt b { font-size: 15px; color: var(--accent); }
      .opt span { font-size: 13px; color: var(--muted); line-height: 1.45; }
      .opt:disabled { opacity: .6; cursor: default; }
    </style>
    <div class="note-dlg" id="ask">
      <h1>Welcome to the PC Tracker</h1>
      <p>Setup is done. One question before you start:</p>
      <h2>Is this your first PC Tracker?</h2>
      <div class="choice">
        <button type="button" class="opt" id="freshBtn" onclick="fresh()">
          <b>Yes, start fresh</b>
          <span>Load your student data and build your sheets, one step at a time.</span>
        </button>
        <button type="button" class="opt" id="oldBtn" onclick="bringOver()">
          <b>No, bring over last year's tracker</b>
          <span>Copies your Contact Log, events, student changes, lists and settings from your old tracker. Your old tracker isn't changed.</span>
        </button>
      </div>
      <p class="hint">You can do the other one later: 🚀 App Menu → 🔄 Import Old Data, or ✅ Getting Started.</p>
    </div>
    <div class="note-dlg hidden" id="fresh">
      <h1>Welcome to the PC Tracker</h1>
      <p>This tracker keeps your parent contacts, events and family directory in one place, and fills in your PCAR for you. Getting started takes four steps:</p>
      <ol>
        <li><b>Load your student data</b> (Input Data Mode walks you through it)</li>
        <li><b>Build your sheets</b> (Input Data Mode does this for you)</li>
        <li><b>Check your settings</b></li>
        <li><b>Log your first parent contact</b></li>
      </ol>
      <p class="help" style="margin:0">Your checklist is open on the right. It opens by itself until you've done all four.</p>
      <div class="foot">
        <button type="button" class="btn sec" onclick="google.script.host.close()">Not now</button>
        <button type="button" class="btn pri" id="goBtn" onclick="go()">Start with Input Data Mode</button>
      </div>
      <div class="sub">
        <p class="help" style="margin:0">Before you start: how to download your data from ATS, Parent Square and NYCSA.</p>
        <div class="foot">
          ${link(links.downloadDataVideo, "Watch the video")}
          ${link(links.downloadDataDoc, "How to download your data")}
          ${link(links.guideDoc, "Read the User Guide")}
          ${link(links.example, "See an example tracker")}
        </div>
      </div>
    </div>
    <script>
      function fresh() {
        document.getElementById('ask').classList.add('hidden');
        document.getElementById('fresh').classList.remove('hidden');
      }
      function bringOver() {
        // The Migration Wizard opens in place of this pop-up
        ['freshBtn', 'oldBtn'].forEach(function (id) { document.getElementById(id).disabled = true; });
        document.querySelector('#oldBtn b').textContent = 'Opening…';
        google.script.run
          .withFailureHandler(function (e) {
            alert('Could not open the Migration Wizard: ' + e.message + ' Use 🚀 App Menu → 🔄 Import Old Data.');
            ['freshBtn', 'oldBtn'].forEach(function (id) { document.getElementById(id).disabled = false; });
            document.querySelector('#oldBtn b').textContent = "No, bring over last year's tracker";
          })
          .zg_action('importOld');
      }
      function go() {
        const b = document.getElementById('goBtn');
        b.disabled = true; b.textContent = 'Opening…';
        google.script.run
          .withSuccessHandler(function () { google.script.host.close(); })
          .withFailureHandler(function (e) { alert('Could not open Input Data Mode: ' + e.message); b.disabled = false; b.textContent = 'Start with Input Data Mode'; })
          .zg_action('inputData');
      }
    </script>`;
  const out = HtmlService.createHtmlOutput(html).setWidth(500).setHeight(470);
  SpreadsheetApp.getUi().showModalDialog(out, "Welcome to the PC Tracker");
}

/** Installable on open (checkVersionOnOpen): opens the checklist until its four steps are done. */
function guideOnOpen_() {
  const props = PropertiesService.getDocumentProperties();
  if (props.getProperty('SETUP_COMPLETE') !== 'true') return;
  if (props.getProperty('INPUT_DATA_MODE_STATE')) return; // Input Data Mode has the sidebar
  if (props.getProperty('ZG_AUTO_OPEN') === 'off') return;
  if (zg_checklist_(SpreadsheetApp.getActiveSpreadsheet()).allDone) return;
  openGuideChecklist();
}

// ----------------------------------------------------------------------
// MENU HANDLER + SIDEBAR ENDPOINTS
// ----------------------------------------------------------------------

function openGuideChecklist() {
  const html = htmlFromFile_('GettingStartedSidebar').setTitle('Getting Started');
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Everything the sidebar shows. */
function zg_getState() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getDocumentProperties();
  return {
    checklist: zg_checklist_(ss),
    autoOpen: props.getProperty('ZG_AUTO_OPEN') !== 'off',
    tab: zg_activeTab(),
    hiddenTabs: zg_hiddenTabs_(ss),
    links: zg_links_(),
    userGuideUrl: zg_safeGuideLink_()
  };
}

/** The tab you're on (polled by the sidebar every few seconds; one cheap call). */
function zg_activeTab() {
  const name = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  const info = ZG_TABS.find(t => t.name === name || t.alt === name);
  return info ? { name: name, known: true, type: info.type, what: info.what } : { name: name, known: false };
}

/** Checklist buttons. Some open another sidebar or a pop-up. */
function zg_action(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  switch (name) {
    case "inputData":
      openInputDataMode();
      return true;
    case "importOld":
      openMigrationWizard();
      return true;
    case "build":
      runBuildOnly();
      return true;
    case "settings":
      PropertiesService.getDocumentProperties().setProperty('ZG_SETTINGS_SEEN', 'true');
      openSettingsDialog();
      return true;
    case "contactLog": {
      const sheet = ss.getSheetByName("Contact Log");
      if (!sheet) throw new Error("There's no Contact Log tab.");
      if (sheet.isSheetHidden()) sheet.showSheet();
      sheet.activate();
      const end = findEndRowInColumnE_(sheet);
      let row = 2;
      if (end && end > 2) {
        const c = sheet.getRange(2, 3, end - 2, 1).getDisplayValues();
        const blank = c.findIndex(r => String(r[0]).trim() === "");
        row = blank === -1 ? end - 1 : 2 + blank;
      }
      sheet.getRange(row, 3).activate();
      return true;
    }
  }
  throw new Error("Unknown step: " + name);
}

/** "Open this by itself when the sheet opens". */
function zg_setAutoOpen(on) {
  const props = PropertiesService.getDocumentProperties();
  if (on) props.deleteProperty('ZG_AUTO_OPEN'); else props.setProperty('ZG_AUTO_OPEN', 'off');
  return !!on;
}

// ----------------------------------------------------------------------
// CHECKLIST
// ----------------------------------------------------------------------

function zg_checklist_(ss) {
  const props = PropertiesService.getDocumentProperties();
  const data = zg_dataStatus_(ss);
  const links = zg_links_();
  const steps = [
    { id: "data", title: "Load your student data", done: data.ats.loaded,
      text: "Input Data Mode walks you through it: paste or upload your ATS, Parent Square and NYCSA exports, then your sites and classes, then your own lists.",
      parts: [
        { label: "Student data (ATS)", done: data.ats.loaded, date: data.ats.date },
        { label: "Parent Square", done: data.ps.loaded, date: data.ps.date },
        { label: "NYCSA", done: data.nycsa.loaded, date: data.nycsa.date },
        { label: "Sites and classes", done: data.siteClass }
      ],
      buttons: [{ action: "inputData", label: "Open Input Data Mode", primary: true },
                { action: "importOld", label: "Import last year's tracker" }],
      links: [{ label: "Video: how to download your data", url: links.downloadDataVideo },
              { label: "Guide: how to download your data", url: links.downloadDataDoc }] },
    { id: "build", title: "Build your sheets", done: zg_isBuilt_(ss),
      text: "Input Data Mode builds them for you before Your Lists (step 5). Fills the Student List, Directory, Combined Contact Tracking and the other tabs from your data. The rest of the menus appear after the first build.",
      buttons: [{ action: "build", label: "Build Sheets Only", primary: true }],
      links: [{ label: "Guide: what Build Sheets does", url: links.buildDoc }] },
    { id: "settings", title: "Check your settings", done: props.getProperty('ZG_SETTINGS_SEEN') === 'true',
      text: "Choose which tabs to show, when sheets rebuild, phone sync, and how the Contact Log behaves.",
      buttons: [{ action: "settings", label: "Open Settings" }],
      links: [{ label: "Guide: what each setting does", url: links.settingsDoc }] },
    { id: "contact", title: "Log your first parent contact", done: zg_hasContact_(ss),
      text: "Type the parent's name in column C of the Contact Log. The student and OSIS fill in for you.",
      buttons: [{ action: "contactLog", label: "Go to the Contact Log" }],
      links: [{ label: "Guide: how to log a parent contact", url: links.contactLogDoc }] }
  ];
  return { steps: steps, doneCount: steps.filter(s => s.done).length, allDone: steps.every(s => s.done) };
}

/** Raw Data row 1 stamps ("ATS - DATE UPDATED: 11/04/26"; "00/00/00" = never loaded). */
function zg_dataStatus_(ss) {
  const raw = ss.getSheetByName("Raw Data") || ss.getSheetByName("RAW Data");
  const read = cell => {
    if (!raw) return { loaded: false, date: "" };
    const text = String(raw.getRange(cell).getDisplayValue());
    const m = text.match(/(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/);
    const loaded = !!m && !/00\/00\/00/.test(text);
    return { loaded: loaded, date: loaded ? m[1] : "" };
  };
  let siteClass = !!PropertiesService.getDocumentProperties().getProperty('INPUT_DATA_MODE_SITECLASS_UPDATED');
  if (!siteClass && raw && raw.getMaxColumns() >= 61 && raw.getMaxRows() >= 7) {
    siteClass = raw.getRange(7, 60, 1, 2).getDisplayValues()[0].join("").trim() !== "";
  }
  return { ats: read("A1"), ps: read("AG1"), nycsa: read("AP1"), siteClass: siteClass };
}

/** Same test the menus use (a student under the Student List header). */
function zg_isBuilt_(ss) {
  return studentListIsBuilt_(ss);
}

/** Any parent name in Contact Log column C above the END bar. */
function zg_hasContact_(ss) {
  const sheet = ss.getSheetByName("Contact Log");
  if (!sheet) return false;
  const end = findEndRowInColumnE_(sheet);
  const last = end ? end - 1 : sheet.getLastRow();
  if (last < 2) return false;
  return sheet.getRange(2, 3, last - 1, 1).getDisplayValues().some(r => String(r[0]).trim() !== "");
}

/** The guide's tabs that are hidden right now, and where to turn each one on. */
function zg_hiddenTabs_(ss) {
  return ZG_TABS.map(info => ({ info: info, sheet: ss.getSheetByName(info.name) || (info.alt ? ss.getSheetByName(info.alt) : null) }))
    .filter(t => t.sheet && t.sheet.isSheetHidden())
    .map(t => ({
      name: t.sheet.getName(),
      what: t.info.what,
      where: t.info.settings ? "⚙️ Settings → Tabs & columns" : "View → Hidden sheets → " + t.sheet.getName()
    }));
}

function zg_safeGuideLink_() {
  try { return getGuideLink_(); } catch (e) { return ""; }
}

/** The example tracker (Version Tracker → Example), "" when there is none: no button then. */
function zg_safeExampleLink_() {
  try { return getTemplateLinks_().exampleUrl || ""; } catch (e) { return ""; }
}

// ==========================================
// 🐞 REPORT A BUG
// A menu click can't open a URL directly, so we pop a tiny dialog that
// auto-opens the bug-report form (with a clickable button as a fallback
// for browsers that block the automatic pop-up).
// ==========================================
// Bug report form: BUG_REPORT_FORM_URL, in the 📌 YOUR LINKS box at the top of this file

function reportABug() {
  const formUrl = BUG_REPORT_FORM_URL;
  const htmlContent = include_('Styles') + `
    <style>
      .note-dlg { padding: 8px 12px 12px; display: flex; flex-direction: column; gap: 14px; }
      .note-dlg p { line-height: 1.5; }
      .note-dlg .foot { display: flex; justify-content: flex-end; gap: 8px; flex-wrap: wrap; }
      .note-dlg .sub { border-top: 1px solid var(--line); padding-top: 14px; display: flex; flex-direction: column; gap: 10px; }
    </style>
    <div class="note-dlg">
      <h1>Report a bug</h1>
      <p>Found something broken or confusing? The report form should open in a new tab. If it didn't, use the button below.</p>
      <div class="foot">
        <button type="button" class="btn sec" onclick="google.script.host.close()">Close</button>
        <a class="btn pri" href="${formUrl}" target="_blank" rel="noopener">Open the bug report form</a>
      </div>
    </div>
    <script>
      // Try to open the form immediately; the button above covers pop-up blockers.
      window.open(${JSON.stringify(formUrl)}, '_blank');
    </script>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(420)
    .setHeight(220);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Report a Bug");
}

// ==========================================
// 1.2. SIDEBAR LAUNCHERS & RUN MODES
// ==========================================
function runBuildOnly() {
  CacheService.getScriptCache().put('RUN_MODE', 'BUILD_ONLY', 600);
  const html = htmlFromFile_('ProgressSidebar')
      .setTitle('Build Sheets Manager')
      .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

function openProgressBarSidebar() {
  CacheService.getScriptCache().put('RUN_MODE', 'FULL_SYNC', 600);
  const html = htmlFromFile_('ProgressSidebar')
      .setTitle('Build & Sync Manager')
      .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ==========================================
// 1.5. SETUP WRAPPER 
// ==========================================
function runInitialSetupWrapper() {
  // Google lets people untick some permissions on its "allow" page. Setup needs all of
  // them (triggers, contacts, Drive, the update check): if one is missing, Google stops
  // here and asks again for the missing ones, before anything is set up.
  if (typeof ScriptApp.requireAllScopes === "function") ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);

  setupAllInstallableTriggers();
  const props = PropertiesService.getDocumentProperties();
  
  props.setProperty('SETUP_COMPLETE', 'true');
  
  // Save this exact sheet's ID so the Clone Detector knows this is the original!
  props.setProperty('ORIGINAL_SHEET_ID', SpreadsheetApp.getActiveSpreadsheet().getId());
  
// =======================================================
  // SETTINGS: inherit the template's saved settings, fall back to defaults.
  // A fresh copy carries its settings in the Backend Version section (cell values
  // survive copying; document properties do not), so we read those here and
  // only use the hardcoded defaults for anything not saved.
  // =======================================================
  const defaultSettings = {
    syncToPhones: 'false',
    phoneContacts: 'false',  // FALSE: Hidden by default
    parentsDivided: 'false', // FALSE: Hidden by default
    notesTab: 'false',       // FALSE: Hidden by default
    sendOut: 'false',        // FALSE: Hidden by default
    siteClass: 'false',      // FALSE: Hidden by default
    typeCol: 'false',        // FALSE: Hidden by default
    autoNotes: 'true',
    disableHiddenNotes: 'false',
    // (autoOverride / ignoreHideCheckboxes / emailNotifications removed 2026-11-12: nothing read them)
    dataReminder: 'week'     // data update reminder: week / 2weeks / month / off
  };

  const savedSettings = loadSettingsFromVersionSheet_();
  const finalSettings = Object.assign({}, defaultSettings, savedSettings || {});

  // Write variables straight into memory to eliminate any TypeError possibilities.
  for (const [key, value] of Object.entries(finalSettings)) {
    props.setProperty(key, String(value));
  }

  // =======================================================
  // HIDE UNNECESSARY SHEETS — start from the default-visible set...
  // =======================================================
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // New copy: move to the Contact Log and hide the 👋 Welcome tab (deleted at the next open / build;
  // a Reset to Display Mode on this very sheet had kept it, and that ends here)
  try { clearWelcomeKeep_(ss); leaveWelcomeTab_(ss); } catch (e) { console.error("Welcome tab: " + e.message); }
  applyDefaultSheetVisibility_(ss);

  // ...then reveal the optional tabs/columns the inherited settings ask for.
  try { updateVisualSettings_(getSettings()); } catch (e) { console.error(e); }

  // Baby-proof the tabs (headers, Contact Log END bar, PCAR, built tabs) — Sheet Lock.gs
  try { applySheetLocks_(ss); } catch (e) { console.error("Sheet lock skipped: " + e.message); }

  // Instant UI swap without refresh! Suppress onOpen's own auto-migrate open
  // for this synchronous call so the wizard isn't opened twice — the
  // authoritative open happens in the if/else right below.
  IS_RUNNING_SETUP = true;
  onOpen();
  IS_RUNNING_SETUP = false;

  // =======================================================
  // 🆕 UPGRADE TEMPLATE vs NEW-USER TEMPLATE
  // Only one modal can be open at a time, so we pick the right one:
  //   Backend!B3 = "Yes" -> jump straight into the Migration Wizard
  //   Backend!B3 = "No"  -> show the normal Welcome guide
  // =======================================================
  if (isAutoMigrateOn_()) {
    openMigrationWizard();
  } else {
    // POP UP THE WELCOME DIALOG NOW THAT THE USER HAS CLICKED!
    // Getting Started: its welcome + the checklist (📖 Open User Guide still opens showWelcomeDialog)
    showGuideWelcome_();
  }
}

/**
 * The first build unlocks the full menus (⚙️ Settings, 👤 Student Overrides, 📆 Events):
 * repaint them right away instead of at the next reload. Called after Build Sheets
 * Only and when Input Data Mode has ended (in its separate tab-delete call, because
 * buildMenus_ also deletes the mode's leftover tab). True when the menus were redrawn.
 */
function refreshMenusAfterFirstBuild_() {
  const props = PropertiesService.getDocumentProperties();
  if (props.getProperty('SETUP_COMPLETE') !== 'true' || props.getProperty('FULL_MENU_UNLOCKED') === 'true') return false;
  if (props.getProperty('INPUT_DATA_MODE_STATE')) return false; // still in Input Data Mode: at its end
  if (!studentListIsBuilt_(SpreadsheetApp.getActiveSpreadsheet())) return false;
  onOpen();
  return true;
}

// ==========================================
// 1.6. 👋 WELCOME TAB (new copies, before Initial Setup)
// Owner's request 2026-11-12. Until Initial Setup has run, every open shows this
// tab and hides all the others (buildMenus_; the simple onOpen may change its own
// sheet). The code makes it when it is missing, so the template doesn't need one.
// Initial Setup moves the user to the Contact Log and hides it; it is deleted at
// the next open or build, never in the setup call itself: deleting the tab the user
// is on breaks the rest of the session ("Sheet <id> not found", see Z#InputDataMode.gs).
// ==========================================
var WELCOME_TAB_NAME = "👋 Welcome";

// Reset to Display Mode (Regenerate.gs) leaves a fresh Welcome tab as the only tab
// showing, so a copy of the template opens straight on it. On the sheet that was
// reset, setup may already have run, and the next open would delete the tab: this
// document property (value = that sheet's ID) keeps it until Initial Setup runs there
// again (or the next reset / fake-data load). Copies don't get it (document
// properties don't travel), and their setup hasn't run, so they keep the tab anyway.
var WELCOME_KEEP_KEY_ = "WELCOME_TAB_KEEP";

/**
 * Every permission Google lists on its "wants to access your Google Account" page
 * (the oauthScopes in the manifest), in Google's words, and what the tool uses it
 * for. On the Welcome tab BEFORE Initial Setup: a menu click can't show anything of
 * ours before Google's own screens (the script may not run until it is allowed), so
 * the explanation has to be on the tab. Same list as the User Guide (Part 1, Step 2).
 */
var WELCOME_PERMISSIONS_ = [
  ["See, edit, create and delete all your Google Sheets spreadsheets",
   "builds the tabs in this tracker, checks the shared update list for a newer version, and opens last year's tracker when you import it."],
  ["See, edit, download and permanently delete your contacts",
   "only for the phone sync, which stays off until you turn it on in ⚙️ Settings. It only adds, changes and removes contacts in its own \"School Directory\" label."],
  ["See, edit, create and delete all of your Google Drive files",
   "saves your end-of-year PDF archive next to this tracker, finds the newest User Guide and template links, and finds your old tracker when you bring it over."],
  ["Connect to an external service",
   "only for that end-of-year PDF, which it downloads from Google's own export page."],
  ["Allow this application to run when you are not present",
   "the instant updates when you type in the Contact Log, the nightly rebuild at 10 PM, and the update check when you open the sheet."],
  ["Display and run third-party web content in prompts and sidebars",
   "the tool's own pop-ups and side panels (Settings, Input Data Mode, Getting Started). \"Third-party\" just means not made by Google."]
];

/**
 * End of Reset to Display Mode: a fresh Welcome tab (the reset deleted the old one,
 * see deleteLegacySheets_) as the only tab showing, kept on this sheet until
 * Initial Setup runs here. Works on a copy that isn't the open sheet too
 * (owner's "Copy Sheet, Then Reset the Copy").
 */
var WELCOME_LINKS_OVERRIDE_ = null; // { exampleUrl, guideUrl } for one reset (🐞 Make Release Copies, z_MINE.gs)

function showWelcomeAfterReset_(ss) {
  const old = ss.getSheetByName(WELCOME_TAB_NAME);
  if (old && old.isSheetHidden()) ss.deleteSheet(old); // an older one the reset didn't delete
  let guideUrl = GUIDE_FOLDER_URL, exampleUrl = "";
  try {
    const r = refreshResolvedLinks_();
    guideUrl = r.guide || GUIDE_FOLDER_URL;
    exampleUrl = (r.links && r.links.exampleUrl) || "";
  } catch (e) { console.warn("Guide link: " + e.message); }
  // The tab's links are written now and stay in every copy of this file until the next reset:
  // a release's Blank copy gets the Example copy made in the same run (not in the tracker yet)
  const o = WELCOME_LINKS_OVERRIDE_ || {};
  if (o.exampleUrl) exampleUrl = o.exampleUrl;
  if (o.guideUrl) guideUrl = o.guideUrl;
  if (!ss.getSheetByName(WELCOME_TAB_NAME)) buildWelcomeTab_(ss, { guideUrl: guideUrl, exampleUrl: exampleUrl });
  try {
    showWelcomeTabOnly_(ss);
  } catch (e) {
    // a copy opened by ID can't be "activated": hide the other tabs without it
    const tab = ss.getSheetByName(WELCOME_TAB_NAME);
    tab.showSheet();
    ss.getSheets().forEach(function(sheet) {
      if (sheet.getSheetId() !== tab.getSheetId() && !sheet.isSheetHidden()) sheet.hideSheet();
    });
  }
  if (SpreadsheetApp.getActiveSpreadsheet().getId() === ss.getId()) {
    PropertiesService.getDocumentProperties().setProperty(WELCOME_KEEP_KEY_, ss.getId());
  }
}

/** The "keep the Welcome tab after setup" mark goes (Initial Setup, every reset). */
function clearWelcomeKeep_(ss) {
  const props = PropertiesService.getDocumentProperties();
  if (props.getProperty(WELCOME_KEEP_KEY_) === ss.getId()) props.deleteProperty(WELCOME_KEEP_KEY_);
}

/** Before setup: makes the Welcome tab if needed, shows it and hides every other tab. */
function showWelcomeTabOnly_(ss) {
  const tab = ss.getSheetByName(WELCOME_TAB_NAME) || buildWelcomeTab_(ss);
  if (tab.isSheetHidden()) tab.showSheet();
  tab.activate();
  ss.getSheets().forEach(function(sheet) {
    if (sheet.getSheetId() !== tab.getSheetId() && !sheet.isSheetHidden()) sheet.hideSheet();
  });
}

/** Initial Setup: moves the user to the Contact Log (else the first default tab) and hides the Welcome tab. */
function leaveWelcomeTab_(ss) {
  const tab = ss.getSheetByName(WELCOME_TAB_NAME);
  if (!tab || tab.isSheetHidden()) return;
  let landing = null;
  for (let i = 0; i < DEFAULT_VISIBLE_SHEETS.length && !landing; i++) landing = ss.getSheetByName(DEFAULT_VISIBLE_SHEETS[i]);
  if (!landing) landing = ss.getSheets().filter(function(s) { return s.getSheetId() !== tab.getSheetId(); })[0];
  if (!landing) return; // the only tab: keep it
  landing.showSheet();
  landing.activate();
  tab.hideSheet();
}

/**
 * After setup (next open / build): deletes the Welcome tab. If the user is on it
 * (setup stopped half way), the default tabs come back and it is only hidden now,
 * then deleted the next time.
 */
function deleteWelcomeTab_(ss) {
  const tab = ss.getSheetByName(WELCOME_TAB_NAME);
  if (!tab) return false;
  // Reset to Display Mode on this sheet: the tab stays until Initial Setup runs here
  if (PropertiesService.getDocumentProperties().getProperty(WELCOME_KEEP_KEY_) === ss.getId()) return false;
  if (!tab.isSheetHidden() && ss.getActiveSheet().getSheetId() === tab.getSheetId()) {
    leaveWelcomeTab_(ss);
    applyDefaultSheetVisibility_(ss);
    return false;
  }
  ss.deleteSheet(tab);
  return true;
}

/**
 * Makes the Welcome tab (first tab, 4 columns, no gridlines) in the "District" colors.
 * opts.guideUrl: the User Guide link (Reset to Display Mode resolves it: Version Tracker,
 * else the guide folder's newest file; the simple onOpen can't read Drive, so it uses the
 * guide folder). opts.exampleUrl: the example tracker (no link = no row).
 */
function buildWelcomeTab_(ss, opts) {
  opts = opts || {};
  const NAVY = "#183e69", INK = "#1b2430", MUTED = "#566173", FAINT = "#8a94a3", LINE = "#d3d9e1";
  const upgrade = isAutoMigrateOn_(ss); // upgrade-template copy: setup opens the Migration Wizard

  const steps = [
    ["Open 🚀 App Menu",
     "It's in the menu bar at the top of the screen, after Help. Don't see it? Wait a few seconds, or reload the page."],
    ["Click 🚨 Initial Setup (Run Once)",
     "It turns on the sheet's automatic updates and brings back your tabs."],
    ["Allow access (the permissions above)",
     "\"Authorization required\": click OK, then choose your Google account. \"Google hasn't verified this app\": " +
     "click Advanced, then \"Go to … (unsafe)\" (Google shows this for tools made by a person instead of a company). " +
     "On the list of permissions, tick Select all, then click Continue. If setup doesn't start after that, " +
     "click 🚨 Initial Setup (Run Once) again."],
    upgrade
      ? ["Bring over your old data",
         "When setup is done, this tab is removed, your tabs come back and the Migration Wizard opens to bring your data over from your old sheet."]
      : ["Start fresh, or bring over last year's tracker",
         "When setup is done, this tab is removed and your tabs come back. A window asks whether this is your first PC Tracker: " +
         "start fresh and load your student data, or bring over your old tracker (it finds it in your Drive)."]
  ];

  // One entry per row: h = fixed height (spacers; other rows grow to fit), band =
  // navy background, b = step number (column B), c = text (column C), style = how
  // it looks, link = c is a link. No merged cells: Sheets doesn't grow a merged row.
  const rows = [
    { h: 20 },
    { h: 16, band: true },
    { h: 44, band: true, c: "Welcome to the PC Tracker", style: "title" },
    { band: true, c: "Your parent contacts, events and family directory in one place, with your PCAR filled in for you.", style: "subtitle" },
    { h: 18, band: true },
    { h: 26 },
    // First what Google will ask and why, then the steps that make it ask
    { c: "BEFORE YOU START: WHAT GOOGLE WILL ASK YOU TO ALLOW", style: "label" },
    { h: 8 },
    { c: "Setting up asks Google for permission once. Its screens can look scary, so here is what each permission is for, in Google's own words:", style: "body" },
    { h: 6 }
  ];
  WELCOME_PERMISSIONS_.forEach(function(perm) {
    rows.push({ c: "•  " + perm[0] + ": " + perm[1], boldTo: perm[0].length + 4, style: "perm" });
  });
  rows.push({ h: 6 });
  rows.push({ c: "Everything runs inside your own Google account. Nothing about your families is sent to the person who made this tool, or to anyone else.", style: "body" });
  rows.push({ h: 26 });
  rows.push({ c: "SET UP YOUR COPY", style: "label" });
  rows.push({ h: 10 });
  steps.forEach(function(step, i) {
    if (i) rows.push({ h: 18 });
    rows.push({ b: i + 1, c: step[0], style: "step" });
    rows.push({ c: step[1], style: "body" });
  });
  rows.push({ h: 26 });
  rows.push({ h: 16, style: "divider" });
  rows.push({ c: "Need help?", style: "help" });
  if (GUIDE_LINKS.setupVideo) rows.push({ c: "🎬 Watch the setup video", link: GUIDE_LINKS.setupVideo });
  rows.push({ c: "📖 Open the User Guide", link: opts.guideUrl || GUIDE_FOLDER_URL });
  if (opts.exampleUrl) rows.push({ c: "👀 See an example tracker (made-up students)", link: opts.exampleUrl });
  rows.push({ c: "🐞 Report a problem", link: BUG_REPORT_FORM_URL });
  rows.push({ h: 12 });
  rows.push({ c: "This tab is only here until Initial Setup has run. It is then deleted for you.", style: "hint" });
  rows.push({ h: 24 });

  const sh = ss.insertSheet(WELCOME_TAB_NAME, 0);
  const n = rows.length;
  if (sh.getMaxColumns() > 4) sh.deleteColumns(5, sh.getMaxColumns() - 4);
  if (sh.getMaxRows() > n) sh.deleteRows(n + 1, sh.getMaxRows() - n);
  sh.setHiddenGridlines(true);
  sh.setTabColor(NAVY);
  [32, 56, 600, 32].forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });

  sh.getRange(1, 1, n, 4)
    .setFontFamily("Montserrat").setFontSize(11).setFontColor(INK)
    .setVerticalAlignment("top").setWrap(true);
  sh.getRange(1, 2, n, 2).setValues(rows.map(function(r) {
    return [r.b || "", r.link ? "" : (r.c || "")];
  }));

  const byStyle = {};
  const add = function(style, a1) { (byStyle[style] = byStyle[style] || []).push(a1); };
  rows.forEach(function(r, i) {
    const row = i + 1;
    if (r.h) sh.setRowHeight(row, r.h);
    if (r.band) add("band", "A" + row + ":D" + row);
    if (r.style === "step") add("number", "B" + row);
    if (r.style === "divider") add("divider", "B" + row + ":C" + row);
    else if (r.style) add(r.style, "C" + row);
    if (r.link) {
      sh.getRange(row, 3).setRichTextValue(SpreadsheetApp.newRichTextValue().setText(r.c).setLinkUrl(r.link)
        .setTextStyle(SpreadsheetApp.newTextStyle().setForegroundColor(NAVY).setUnderline(true).build()).build());
    }
    if (r.boldTo) { // the permission's name in bold, its reason plain
      sh.getRange(row, 3).setRichTextValue(SpreadsheetApp.newRichTextValue().setText(r.c)
        .setTextStyle(0, r.boldTo, SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor(INK).build()).build());
    }
  });
  const look = function(style, apply) { if (byStyle[style]) apply(sh.getRangeList(byStyle[style])); };
  look("band", function(l) { l.setBackground(NAVY); });
  look("title", function(l) { l.setFontSize(20).setFontWeight("bold").setFontColor("#ffffff").setVerticalAlignment("middle"); });
  look("subtitle", function(l) { l.setFontColor("#d5deea"); });
  look("label", function(l) { l.setFontSize(10).setFontWeight("bold").setFontColor(NAVY); });
  look("number", function(l) { l.setFontSize(16).setFontWeight("bold").setFontColor(NAVY).setHorizontalAlignment("center"); });
  look("step", function(l) { l.setFontSize(13).setFontWeight("bold"); });
  look("body", function(l) { l.setFontColor(MUTED); });
  look("perm", function(l) { l.setFontColor(MUTED); });
  look("divider", function(l) { l.setBorder(true, null, null, null, null, null, LINE, SpreadsheetApp.BorderStyle.SOLID); });
  look("help", function(l) { l.setFontSize(12).setFontWeight("bold"); });
  look("hint", function(l) { l.setFontSize(10).setFontColor(FAINT); });
  return sh;
}

// ==========================================
// 4. SHEET LOCK & CUSTOM WARNING
// Before Initial Setup, an edit is undone with "🔒 Action Blocked". Not while any
// owner test file (z_...) is in the project (owner's request 2026-11-12): the owner's
// sheet and test copies can be typed in before setup; a school's copy (no z_ files)
// is still blocked.
// ==========================================
function ownerTestFilesPresent_() {
  return typeof ownerHooks_ === "function"                  // z_MINE
      || typeof addContactLogDebugMenuItem_ === "function"  // z_ContactLogDebug
      || typeof idm_debugEnabled_ === "function"            // z_InputDataDebug
      || typeof ownerNewCopyCheck_ === "function";          // z_NewCopy
}

function onEdit(e) {
  const props = PropertiesService.getDocumentProperties();
  
  if (props.getProperty('SETUP_COMPLETE') === 'true') {
    return; // Normal edits completely bypass the setup lock block below
  }
  if (ownerTestFilesPresent_()) return; // owner's sheet / test copies: no block

  if (e && e.range) {
    if (e.oldValue !== undefined) e.range.setValue(e.oldValue); 
    else e.range.clearContent();       
  }
  SpreadsheetApp.getUi().alert("🔒 Action Blocked", "Initialization process needs to occur first.\n\nPlease go to '🚀 App Menu' at the top of the screen and click '🚨 Initial Setup'.", SpreadsheetApp.getUi().ButtonSet.OK);
}

// ==========================================
// ⚙️ SETTINGS BACKEND LOGIC (was Settings Backend.gs)
// ==========================================
function openSettingsDialog() {
  const html = htmlFromFile_('SettingsDialog')
      .setTitle('Dashboard Settings')
      .setWidth(780)
      .setHeight(580);
  SpreadsheetApp.getUi().showModalDialog(html, 'Settings');
}

// =========================================
//  MAIN SETTINGS MEMORY
// =========================================

// Every setting key that should survive into a fresh copy. Persisted as
// key/value rows in the Backend Version section (A6:B), because document
// properties do NOT travel when a spreadsheet is copied but cell values do.
const SETTING_KEYS = [
  'syncToPhones', 'phoneContacts', 'parentsDivided', 'notesTab', 'sendOut',
  'showContactLog', 'showCombined', 'showPcar', 'showDirectory', 'showRawData',
  'showMasterTable', 'showEvents', 'showOverrideReport', 'showNewStudents',
  'siteClass', 'typeCol', 'autoNotes', 'disableHiddenNotes',
  'disableRebuilds',
  'disableNightlySync', 'disableContactAutofill', 'notesStartDate',
  'notesEndDate', 'hideDischarged', 'contactLogTodayOnly', 'dataReminder'
];

/**
 * Snapshots the current settings (from document properties) into the Backend
 * Version section so a future copy of this template inherits them. Writes a
 * header at A5 and key/value pairs from row 6 down (cols A:B only), leaving
 * A2/A3/B3 (build date + Auto-Migrate toggle) untouched.
 */
function saveSettingsToVersionSheet_(ss) {
  const props = PropertiesService.getDocumentProperties();
  ss = ss || SpreadsheetApp.getActiveSpreadsheet(); // ss: another spreadsheet (owner's copy + reset tool)
  const sheet = getOrCreateVersionSheet_(ss);
  if (!sheet) return;

  const rows = SETTING_KEYS.map(function(k) {
    const v = props.getProperty(k);
    return [k, v == null ? '' : v];
  });

  // Clear the old settings block (rows 5+ in cols A:B) before rewriting.
  const headerRow = BACKEND_VERSION.SETTINGS_HEADER_ROW;
  ensureBackendRows_(sheet, BACKEND_VERSION.SETTINGS_FIRST_ROW + rows.length);
  sheet.getRange(headerRow, 1, sheet.getMaxRows() - headerRow + 1, 2).clearContent();
  sheet.getRange(headerRow, 1, 1, 2).setValues([['SAVED SETTINGS (auto-managed — do not edit)', '']]);
  sheet.getRange(BACKEND_VERSION.SETTINGS_FIRST_ROW, 1, rows.length, 2).setValues(rows);
}

/**
 * Reads the saved settings block back out of the Backend Version section into a plain
 * {key: "value"} object (string values, matching document-property storage).
 * Returns null when there's nothing saved. Only recognized keys are returned.
 */
function loadSettingsFromVersionSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateVersionSheet_(ss);
  if (!sheet) return null;
  const firstRow = BACKEND_VERSION.SETTINGS_FIRST_ROW;
  const lastRow = sheet.getLastRow();
  if (lastRow < firstRow) return null;

  const data = sheet.getRange(firstRow, 1, lastRow - firstRow + 1, 2).getValues();
  const out = {};
  data.forEach(function(r) {
    const key = String(r[0]).trim();
    if (key && SETTING_KEYS.indexOf(key) !== -1) out[key] = String(r[1]);
  });
  return Object.keys(out).length ? out : null;
}

function getSettings() {
  const props = PropertiesService.getDocumentProperties();
  const settings = {
    syncToPhones: props.getProperty('syncToPhones') === 'true',
    
    // --- CHANGED: Columns now strictly default to FALSE (Hidden) ---
    siteClass: props.getProperty('siteClass') === 'true',
    typeCol: props.getProperty('typeCol') === 'true',
    
    // --- LIVE UPDATE SETTINGS ---
    disableRebuilds: props.getProperty('disableRebuilds') === 'true',
    disableNightlySync: props.getProperty('disableNightlySync') === 'true',
    autoNotes: props.getProperty('autoNotes') !== 'false',

    // Contact Log auto-fill of Date / # / Follow-up (default OFF = auto-fill ON).
    // Does NOT affect the parent/student/OSIS lookup.
    disableContactAutofill: props.getProperty('disableContactAutofill') === 'true',

    // Contact Log filter window: today's notes only (true) or this/last week's (default)
    contactLogTodayOnly: props.getProperty('contactLogTodayOnly') === 'true',
    
    // --- BYPASS SWITCH & DATES ---
    disableHiddenNotes: props.getProperty('disableHiddenNotes') === 'true', 
    notesStartDate: props.getProperty('notesStartDate') || '',
    notesEndDate: props.getProperty('notesEndDate') || '',

    // --- DATA MANAGEMENT ---
    hideDischarged: props.getProperty('hideDischarged') === 'true',

    // Data update reminder pop-up: 'week' (default), '2weeks', 'month', 'off'
    dataReminder: getDataReminderSetting_()
  };

  // --- Tab visibility (TAB_VISIBILITY): saved value, else the tab's default ---
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  settings.tabsPresent = {};
  TAB_VISIBILITY.forEach(function(t) {
    const v = props.getProperty(t.key);
    settings[t.key] = (v === null || v === '') ? t.def : v === 'true';
    settings.tabsPresent[t.key] = t.sheets.some(function(name) { return !!ss.getSheetByName(name); });
  });
  return settings;
}

function updateVisualSettings_(settings) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Tab Visibility (every tab in TAB_VISIBILITY; a key left out of the
  // payload keeps the tab's default). Shown tabs first, so hiding never runs
  // into "can't hide every sheet" while a tab to show is still hidden. ---
  const wantShown = function(t) { return (settings[t.key] === undefined) ? t.def : settings[t.key] === true; };
  const ordered = TAB_VISIBILITY.filter(wantShown).concat(TAB_VISIBILITY.filter(function(t) { return !wantShown(t); }));
  ordered.forEach(function(t) {
    t.sheets.forEach(function(sheetName) {
      try {
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) return;
        if (wantShown(t)) sheet.showSheet();
        else sheet.hideSheet();
      } catch (e) {
        console.warn("Could not hide/show sheet: " + sheetName);
      }
    });
  });

  // --- Contact Log Column Visibility ---
  try {
    const contactLog = ss.getSheetByName("Contact Log");
    if (contactLog) {
      if (settings.siteClass) {
        contactLog.showColumns(8, 2);
      } else {
        contactLog.hideColumns(8, 2);
      }

      if (settings.typeCol) {
        contactLog.showColumns(12);
      } else {
        contactLog.hideColumns(12);
      }
    }
  } catch (e) {
    console.warn("Could not update Contact Log columns.");
  }

  // --- PCAR Visibility Logic ---
  try {
    const pcarSheet = ss.getSheetByName("PCAR");
    if (pcarSheet && pcarSheet.getMaxRows() >= 12) {
      if (settings.sendOut) {
        pcarSheet.showRows(11, 2); // Shows row 11 and 12
      } else {
        pcarSheet.hideRows(11, 2); // Hides row 11 and 12
      }
    }
  } catch (e) {
    console.warn("Could not update PCAR rows.");
  }

  // --- Student List Filter (Hide Discharged) ---
  // The filter starts at the header row (studentListHeaderRow_: row 1 since the
  // 2026-11-11 build, row 4 on a tab not rebuilt since)
  try {
    const masterSheet = getStudentListSheet_(ss);
    if (masterSheet) {
      const headerRow = studentListHeaderRow_(masterSheet);
      const lastRow = Math.max(masterSheet.getLastRow(), headerRow);
      const range = masterSheet.getRange(headerRow, 1, lastRow - headerRow + 1, Math.max(1, masterSheet.getLastColumn()));
      let filter = masterSheet.getFilter();
      if (filter && filter.getRange().getRow() !== headerRow) { filter.remove(); filter = null; } // an old filter on the wrong row
      
      if (settings.hideDischarged) {
        if (!filter) filter = range.createFilter();
        
        const headers = range.getValues()[0];
        if (headers) {
          const statusCol = headers.indexOf("Status") + 1; 
          
          if (statusCol > 0) {
            const criteria = SpreadsheetApp.newFilterCriteria().setHiddenValues(["Discharged"]).build();
            filter.setColumnFilterCriteria(statusCol, criteria);
          }
        }
      } else {
        if (filter) filter.remove();
      }
    }
  } catch (e) {
    console.warn("Could not update Student List filter.");
  }
}

// --- TRIGGER MANAGEMENT ---
// The sheet has ONE nightly trigger (see "ONE NIGHTLY JOB PER SHEET" in Build.gs):
// a Settings save by someone who isn't its maker makes no second one (the
// trigger reads the "never at night" choice itself when it fires), unless the
// remembered one has stopped running.
function manageNightlyTrigger_(disableSync) {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const syncFunctionName = 'nightlyMasterSync'; 
    const props = PropertiesService.getDocumentProperties();
    const owned = props.getProperty(NIGHTLY_UID_KEY_);
    const mine = triggers.filter(t => t.getHandlerFunction() === syncFunctionName);
    const someoneElses = !!owned && !mine.some(t => t.getUniqueId() === owned) && nightlyOwnerAlive_(props);

    for (let i = 0; i < mine.length; i++) ScriptApp.deleteTrigger(mine[i]);

    if (someoneElses) return; // another person's trigger does the job (and reads disableNightlySync)
    if (disableSync) {
      props.deleteProperty(NIGHTLY_UID_KEY_);
      props.deleteProperty(NIGHTLY_UID_SET_KEY_);
      return;
    }
    const t = ScriptApp.newTrigger(syncFunctionName)
      .timeBased()
      .atHour(NIGHTLY_SYNC_HOUR)
      .everyDays(1)
      .create();
    rememberNightlyTrigger_(props, t.getUniqueId());
  } catch (e) {
    console.warn("Trigger management failed: " + e.message);
  }
}

// --- PERMANENT PURGE FUNCTION ---
function purgeDischargedStudents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = getStudentListSheet_(ss);
  if (!masterSheet) return false;
  
  // From the header row down (row 1 since the 2026-11-11 build; row 4 on a tab not rebuilt since)
  const headerRow = studentListHeaderRow_(masterSheet);
  const lastRow = masterSheet.getLastRow(), lastCol = masterSheet.getLastColumn();
  if (lastRow <= headerRow || lastCol < 1) return 0;
  const data = masterSheet.getRange(headerRow, 1, lastRow - headerRow + 1, lastCol).getValues();
  
  const headers = data[0];
  const statusColIdx = headers.indexOf("Status");
  if (statusColIdx === -1) return 0;
  
  const currentData = [headers];
  let purgedCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    const status = String(data[i][statusColIdx]).trim();
    if (status.toLowerCase() !== "discharged") {
      currentData.push(data[i]);
    } else {
      purgedCount++;
    }
  }
  
  if (purgedCount > 0) {
    const filter = masterSheet.getFilter();
    if (filter) filter.remove(); 
    
    masterSheet.getRange(headerRow, 1, lastRow - headerRow + 1, lastCol).clearContent(); // rows above the header stay
    masterSheet.getRange(headerRow, 1, currentData.length, currentData[0].length).setValues(currentData);
    
    const props = PropertiesService.getDocumentProperties();
    if (props.getProperty('hideDischarged') === 'true') {
      const newFilter = masterSheet.getRange(headerRow, 1, currentData.length, currentData[0].length).createFilter();
      const criteria = SpreadsheetApp.newFilterCriteria().setHiddenValues(["Discharged"]).build();
      newFilter.setColumnFilterCriteria(statusColIdx + 1, criteria);
    }
  }
  
  return purgedCount;
}

function applySettings(settings) {
  const props = PropertiesService.getDocumentProperties();
  
  const oldDisableHidden = props.getProperty('disableHiddenNotes') === 'true';
  const oldStartDate = props.getProperty('notesStartDate') || '';
  const oldEndDate = props.getProperty('notesEndDate') || '';

  const newDisableHidden = settings.disableHiddenNotes === true;
  const newTodayOnly = settings.contactLogTodayOnly === true;

  // 2. Save settings - Using String() safely handles undefined payload data instead of crashing!
  props.setProperty('syncToPhones', String(settings.syncToPhones));
  props.setProperty('phoneContacts', String(settings.phoneContacts));
  props.setProperty('parentsDivided', String(settings.parentsDivided));
  props.setProperty('notesTab', String(settings.notesTab)); 
  props.setProperty('sendOut', String(settings.sendOut));
  TAB_VISIBILITY.forEach(function(t) {
    if (settings[t.key] !== undefined) props.setProperty(t.key, String(settings[t.key] === true));
  });
  props.setProperty('siteClass', String(settings.siteClass));
  props.setProperty('typeCol', String(settings.typeCol));
  props.setProperty('disableRebuilds', String(settings.disableRebuilds));
  props.setProperty('disableNightlySync', String(settings.disableNightlySync));
  props.setProperty('autoNotes', String(settings.autoNotes));
  props.setProperty('disableContactAutofill', String(settings.disableContactAutofill));
  props.setProperty('disableHiddenNotes', String(newDisableHidden));
  props.setProperty('notesStartDate', settings.notesStartDate || '');
  props.setProperty('notesEndDate', settings.notesEndDate || '');
  props.setProperty('hideDischarged', String(settings.hideDischarged));
  props.setProperty('contactLogTodayOnly', String(newTodayOnly));
  // Data update reminder; a new choice here also ends any "Remind After ..." wait
  const reminder = (settings.dataReminder === 'off' || DATA_REMINDER_DAYS[settings.dataReminder]) ? settings.dataReminder : 'week';
  if (reminder !== getDataReminderSetting_()) props.deleteProperty(DATA_REMINDER_SNOOZE_KEY);
  props.setProperty('dataReminder', reminder);

  // 2b. Persist to the Version tab so a future copy inherits these settings.
  try { saveSettingsToVersionSheet_(); } catch (e) { console.error(e); }

  // 3. Update Visuals & Student List Filter safely
  updateVisualSettings_(settings);
  
  // 3b. Contact Log filter window: make the Column P rule match the setting (no writes
  // when it already does) and refresh the filter
  try { refreshContactLogFilter_(); } catch (e) { console.warn("Could not refresh Contact Log filter: " + e.message); }

  // 4. Update Triggers Programmatically safely
  manageNightlyTrigger_(settings.disableNightlySync);
  
  // 5. Rebuild Menu Instantly (Calling the true onOpen directly)
  try {
    onOpen();
  } catch (e) {
    console.warn("Could not instantly rebuild UI menus: " + e.message);
  }

  // 6. Check if we need to rebuild and show the appropriate toast
  let needRebuild = (oldDisableHidden !== newDisableHidden) || 
                    (oldStartDate !== settings.notesStartDate) || 
                    (oldEndDate !== settings.notesEndDate);
  
  if (settings.disableRebuilds) {
    needRebuild = false;
  }
  
  if (needRebuild) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Settings saved. Updating notes across sheets in the background...", 
      "Processing", 
      10
    );
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Settings applied successfully.", 
      "Settings Saved", 
      5
    );
  }

  return needRebuild; 
}

/**
 * Settings save: puts the "Contact Log shows" window into the Contact Log's Column P
 * grey-out rule (getContactLogHideFormula_) and re-sets the Column P filter so rows
 * re-hide/re-show right away. Only the P rule changes (Type colors and any other rule
 * stay); nothing is written when the rule already matches the setting.
 */
function refreshContactLogFilter_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Contact Log");
  if (!sheet) return;
  const wanted = getContactLogHideFormula_();
  const rules = sheet.getConditionalFormatRules();
  let found = false, changed = false;
  const newRules = rules.map(function(rule) {
    const ranges = rule.getRanges();
    const onlyP = ranges.length > 0 && ranges.every(r => r.getColumn() === 16 && r.getLastColumn() === 16);
    const cond = onlyP ? rule.getBooleanCondition() : null;
    if (!cond || cond.getCriteriaType() !== SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) return rule;
    found = true;
    if (String(cond.getCriteriaValues()[0]) === wanted) return rule;
    changed = true;
    return rule.copy().whenFormulaSatisfied(wanted).build(); // same range and grey font
  });
  if (!found) {
    // No P rule (removed by hand): add it back as resetContactLogFormatting draws it
    newRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(wanted)
      .setFontColor("#999999")
      .setRanges([sheet.getRange("P2:P" + sheet.getMaxRows())]).build());
    changed = true;
  }
  if (!changed) return;
  sheet.setConditionalFormatRules(newRules);

  const filter = sheet.getFilter();
  if (!filter) return;
  const criteria = filter.getColumnFilterCriteria(16);
  if (criteria) filter.setColumnFilterCriteria(16, criteria);
}

/**
 * ==============================================================================
 * STUDENT OVERRIDES BACKEND (was Student Overrides Backend.gs + Student Override Report.gs)
 * Handles saving overrides, memory settings, and surgical spreadsheet updates.
 * ==============================================================================
 */
function openNewStudentDialog() {
  const html = htmlFromFile_('NewStudentDialog')
      .setTitle('New / Edit Student Override')
      .setWidth(600) 
      .setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, 'New / Edit Student');
}

function openManageOverridesDialog() {
  const html = htmlFromFile_('ManageOverridesDialog')
      .setTitle('Manage Student Overrides')
      .setWidth(500)
      .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Manage Overrides');
}

/**
 * Triggered when the user clicks 'Save' in the HTML pop-up.
 * Bulk saves an array of student payloads, then hands off the heavy processing to a background trigger.
 */
function saveStudentOverrides(payloads) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");
  // New/Edit Student section of the Backend tab (index 0 = header row)
  const existingData = backendReadBlock_(ss, "OVERRIDES");
  const osisSet = new Set(payloads.map(p => toText_(p.osis)));

  // 1. DUPLICATE REMOVAL (only this section's cells shift up; other Backend sections stay put)
  const dupIdxs = [];
  for (let i = existingData.length - 1; i > 0; i--) {
    const rowOsis = toText_(existingData[i][2]); 
    if (rowOsis !== "" && osisSet.has(rowOsis)) dupIdxs.push(i);
  }
  backendDeleteRows_(ss, "OVERRIDES", dupIdxs);

  // 2. COMPILE ALL NEW ROWS IN MEMORY
  const rowsToAppend = payloads.map(data => {
      // Expanded to 42 columns to support Label and Hide From Report
      const row = new Array(42).fill("");
      row[0] = dateStr;           
      row[1] = data.entryType;    
      row[2] = data.osis;         
      row[3] = data.lastName;     
      row[4] = data.firstName;    
      row[5] = data.gender;       
      row[6] = data.age;          
      row[7] = data.grade;        
      row[8] = data.classCode;    
      row[9] = data.homeLanguage; 

      // Guardians Section (Up to 5)
      for (let i = 0; i < data.guardians.length; i++) {
        if (i >= 5) break;
        const g = data.guardians[i];
        const baseCol = 10 + (i * 6);
        row[baseCol] = g.name;
        row[baseCol + 1] = g.relationship;
        row[baseCol + 2] = g.phone;
        row[baseCol + 3] = g.email;
        // Map the single language UI input perfectly to BOTH legacy columns
        row[baseCol + 4] = g.language;
        row[baseCol + 5] = g.language;
      }
      
      // Save Label & Stealth Mode
      row[40] = data.label;
      row[41] = data.hideFromReport;

      return row;
  });

  // 3. BULK PASTE INSTANTLY
  backendAppendRows_(ss, "OVERRIDES", rowsToAppend);
  SpreadsheetApp.flush();

 // 4. CHECK SETTINGS & REBUILD
  const settings = getSettings();
  const isPaused = (settings.disableRebuilds === true || settings.disableRebuilds === 'true');

  if (!isPaused) {
      // Build is ON: open the same progress sidebar the "Build Sheets Only"
      // menu item uses. Its build-only process runs the full rebuild, which
      // integrates the new/edited override. The dialog closes on success
      // (see NewStudentDialog.html), revealing the sidebar behind it.
      runBuildOnly();
      return "Rebuilding via sidebar...";
  } else {
      SpreadsheetApp.getActiveSpreadsheet().toast("Overrides saved locally. Background sync is paused via Settings.", "⚡ Fast Entry Mode", 4);
      return "Saved locally.";
  }
}

// =========================================
//  DATA FETCHERS & MANAGE UI LOGIC
// =========================================
function getAllStudentsForEdit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName("Raw Data");
  const masterSheet = getStudentListSheet_(ss);

  // Dynamically map current Statuses from the Student List
  const statusMap = new Map();
  if (masterSheet) {
    const mData = masterSheet.getDataRange().getValues();
    let osisIdx = 0, statIdx = 6, headerIdx = -1;
    for(let i=0; i < Math.min(10, mData.length); i++) {
        for(let c=0; c<mData[i].length; c++) {
            if(String(mData[i][c]).toUpperCase() === "OSIS") { osisIdx = c; if (headerIdx === -1) headerIdx = i; }
            if(String(mData[i][c]).toUpperCase() === "STATUS") statIdx = c;
        }
    }
    // Students start under the header (row 2 now; row 5 on a tab not rebuilt since 2026-11-11)
    for(let i = headerIdx !== -1 ? headerIdx + 1 : 4; i < mData.length; i++) {
        statusMap.set(toText_(mData[i][osisIdx]), toText_(mData[i][statIdx]));
    }
  }

  const studentMap = new Map();

  // 1. Pull ATS Base Data
  if (rawSheet) {
    const rawData = rawSheet.getDataRange().getValues();
    for (let i = 1; i < rawData.length; i++) {
      const r = rawData[i];
      const osis = toText_(r[2]);
      if (osis.replace(/\D/g, "").length < 6) continue;

      const g1Name = toText_(r[19]);
      const g2Name = toText_(r[25]);
      const guardians = [];

      if (g1Name) guardians.push({ name: g1Name, relationship: toText_(r[20]), phone: toText_(r[21]), email: toText_(r[22]), language: toText_(r[23]) || toText_(r[24]) });
      if (g2Name) guardians.push({ name: g2Name, relationship: toText_(r[26]), phone: toText_(r[27]), email: toText_(r[28]), language: toText_(r[29]) || toText_(r[30]) });

      studentMap.set(osis, {
        osis: osis,
        lastName: toText_(r[3]),
        firstName: toText_(r[4]),
        gender: toText_(r[6]),
        age: toText_(r[7]),
        grade: toText_(r[16]),
        classCode: toText_(r[17]),
        homeLanguage: toText_(r[18]),
        guardians: guardians,
        label: statusMap.get(osis) || "Current",
        hideFromReport: false
      });
    }
  }

  // 2. Pull Overrides (And overwrite any base data)
  {
    const overData = backendReadBlock_(ss, "OVERRIDES");
    for (let i = 1; i < overData.length; i++) {
      const r = overData[i];
      const osis = toText_(r[2]);
      if (osis.replace(/\D/g, "").length < 6) continue;

      const guardians = [];
      for (let g = 0; g < 5; g++) {
        const baseCol = 10 + (g * 6);
        const name = toText_(r[baseCol]);
        if (name) {
          guardians.push({
            name: name,
            relationship: toText_(r[baseCol+1]),
            phone: toText_(r[baseCol+2]),
            email: toText_(r[baseCol+3]),
            language: toText_(r[baseCol+4]) || toText_(r[baseCol+5])
          });
        }
      }

      // Keep the current Raw Data class (matches buildStudentMap_); fall back to the override's
      const rawBase = studentMap.get(osis);
      studentMap.set(osis, {
        osis: osis,
        lastName: toText_(r[3]),
        firstName: toText_(r[4]),
        gender: toText_(r[5]),
        age: toText_(r[6]),
        grade: toText_(r[7]),
        classCode: (rawBase && rawBase.classCode) || toText_(r[8]),
        homeLanguage: toText_(r[9]),
        guardians: guardians,
        label: toText_(r[40]) || statusMap.get(osis) || "Current",
        hideFromReport: String(r[41]).toLowerCase() === 'true'
      });
    }
  }

  return Array.from(studentMap.values()).sort((a, b) => a.lastName.localeCompare(b.lastName));
}

function getOverridesList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = backendReadBlock_(ss, "OVERRIDES");
  const list = [];

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const osis = toText_(r[2]);
    if (osis.replace(/\D/g, "").length < 6) continue;

    let dateStr = r[0];
    if (Object.prototype.toString.call(dateStr) === "[object Date]" && !isNaN(dateStr)) {
       dateStr = Utilities.formatDate(dateStr, Session.getScriptTimeZone(), "MM/dd/yyyy");
    }

    list.push({
      date: dateStr,
      type: toText_(r[1]),
      osis: osis,
      lastName: toText_(r[3]),
      firstName: toText_(r[4]),
      classCode: toText_(r[8])
    });
  }
  return list.reverse();
}

function deleteOverride(osis) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = backendReadBlock_(ss, "OVERRIDES");
  const idxs = [];

  for (let i = data.length - 1; i > 0; i--) {
    if (toText_(data[i][2]) === toText_(osis)) idxs.push(i);
  }
  backendDeleteRows_(ss, "OVERRIDES", idxs);
  const deleted = idxs.length > 0;

  if (deleted) {
    const settings = getSettings();
    const isPaused = (settings.disableRebuilds === true || settings.disableRebuilds === 'true');

    if (!isPaused) {
      // Build is ON: open the build progress sidebar (same as the menu build).
      // The Manage Overrides dialog stays open and refreshes its list, so the
      // sidebar runs behind it and is visible once that dialog is closed.
      runBuildOnly();
    } else {
      SpreadsheetApp.getActiveSpreadsheet().toast("Override deleted locally. Sheets will update on next manual build.", "⚡ Fast Mode", 4);
    }
  }
  return true;
}

function setPrefillAndOpenEdit(osis) {
  PropertiesService.getUserProperties().setProperty('PREFILL_OSIS', osis);
  openNewStudentDialog();
}

function getPrefillOsis() {
  const props = PropertiesService.getUserProperties();
  const osis = props.getProperty('PREFILL_OSIS');
  props.deleteProperty('PREFILL_OSIS');
  return osis || "";
}

/**
 * Generates a visual side-by-side Delta report comparing Raw Data vs. Overridden Data.
 * Skips any records that have "Stealth Mode / Hide from Report" checked!
 */
function buildOverrideReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportSheetName = "Students Override Report";
  
  const rawSheet = ss.getSheetByName("RAW Data") || ss.getSheetByName("Raw Data");
  
  if (!rawSheet) {
    SpreadsheetApp.getUi().alert("Error: Could not find the 'Raw Data' sheet.");
    return;
  }

  // 1. Read Data (overrides live in the Backend tab's New/Edit Student section)
  const rawData = rawSheet.getDataRange().getValues().slice(6); 
  const overrideData = backendReadBlock_(ss, "OVERRIDES").slice(1);

  // 2. Map RAW Data by OSIS
  const rawMap = new Map();
  rawData.forEach(row => {
    const osis = String(row[2]).trim();
    if (osis) rawMap.set(osis, row);
  });

  // 3. Helper to Extract Guardians
  function extractGuardians(row, startCol, maxGuardians) {
    const guardians = [];
    for (let i = 0; i < maxGuardians; i++) {
      const base = startCol + (i * 6);
      const name = String(row[base] || "").trim();
      if (!name) continue; 
      
      // Pulls Spoken Lang (base+4) since Spoken and Written are synced now
      guardians.push({
        name: name,
        phone: String(row[base + 2] || "").trim(),
        email: String(row[base + 3] || "").trim(),
        lang: String(row[base + 4] || "").trim() 
      });
    }
    return guardians;
  }

  // 4. Helper to format a full parent block
  function formatFullParent(g) {
    return `Guardian: ${g.name}\n  Phone: ${g.phone || "None"}\n  Email: ${g.email || "None"}\n  Lang: ${g.lang || "None"}`;
  }

  const reportOutput = [];
  const richTextNewInfo = []; 
  const boldStyle = SpreadsheetApp.newTextStyle().setBold(true).build();

  // Add Headers
  reportOutput.push(["Student OSIS", "Old Information", "New Information"]);

  // 5. Main Comparison Engine
  overrideData.forEach(oRow => {
    const osis = String(oRow[2]).trim();
    if (!osis) return; 

    // THE GATEKEEPER: If Stealth Mode is on, skip this student entirely!
    const hideFromReport = String(oRow[41]).trim().toLowerCase() === "true";
    if (hideFromReport) return;

    const newStudentName = `${String(oRow[4] || "").trim()} ${String(oRow[3] || "").trim()}`.trim();
    const newGuardians = extractGuardians(oRow, 10, 5);

    // --- CASE A: BRAND NEW STUDENT ---
    if (!rawMap.has(osis)) {
      let newBlocks = [`Student Name: ${newStudentName}`];
      let namesToBold = [];

      newGuardians.forEach(ng => {
        newBlocks.push(formatFullParent(ng));
        namesToBold.push(ng.name);
      });

      const newInfoString = newBlocks.join("\n\n");
      const rtv = SpreadsheetApp.newRichTextValue().setText(newInfoString);
      
      namesToBold.forEach(name => {
        let startIdx = newInfoString.indexOf(`Guardian: ${name}`);
        if (startIdx !== -1) rtv.setTextStyle(startIdx + 10, startIdx + 10 + name.length, boldStyle);
      });

      reportOutput.push([osis, "[BRAND NEW STUDENT - NOT IN ATS YET]", newInfoString]);
      richTextNewInfo.push([rtv.build()]);
      return; 
    }

    // --- CASE B: EXISTING STUDENT COMPARISON ---
    const rRow = rawMap.get(osis);
    const oldStudentName = `${String(rRow[4] || "").trim()} ${String(rRow[3] || "").trim()}`.trim();
    const oldGuardians = extractGuardians(rRow, 19, 2);

    let oldBlocks = [];
    let newBlocks = [];
    let namesToBold = [];
    let hasChanges = false;

    // 1. Check Student Name
    if (newStudentName.toLowerCase() !== oldStudentName.toLowerCase()) {
      oldBlocks.push(`Student Name: ${oldStudentName}`);
      newBlocks.push(`Student Name: ${newStudentName}`);
      hasChanges = true;
    }

    // 2. Check Added & Modified Guardians
    newGuardians.forEach(ng => {
      const og = oldGuardians.find(g => g.name.toLowerCase() === ng.name.toLowerCase());
      
      if (!og) {
        // ADDED NEW PARENT
        oldBlocks.push(`Guardian: [Not in ATS]`);
        newBlocks.push(formatFullParent(ng));
        namesToBold.push(ng.name);
        hasChanges = true;
      } else {
        // MODIFIED EXISTING PARENT
        let oldF = [];
        let newF = [];
        
        if (og.phone !== ng.phone) { oldF.push(`Phone: ${og.phone || "None"}`); newF.push(`Phone: ${ng.phone || "None"}`); }
        if (og.email !== ng.email) { oldF.push(`Email: ${og.email || "None"}`); newF.push(`Email: ${ng.email || "None"}`); }
        if (og.lang !== ng.lang) { oldF.push(`Lang: ${og.lang || "None"}`); newF.push(`Lang: ${ng.lang || "None"}`); }

        if (newF.length > 0) {
          oldBlocks.push(`Guardian: ${og.name}\n  ${oldF.join("\n  ")}`);
          newBlocks.push(`Guardian: ${ng.name}\n  ${newF.join("\n  ")}`);
          namesToBold.push(ng.name);
          hasChanges = true;
        }
      }
    });

    // 3. Check Removed Guardians
    oldGuardians.forEach(og => {
      const ng = newGuardians.find(g => g.name.toLowerCase() === og.name.toLowerCase());
      if (!ng) {
        oldBlocks.push(formatFullParent(og));
        newBlocks.push(`Guardian: ${og.name} [REMOVED]`);
        namesToBold.push(og.name);
        hasChanges = true;
      }
    });

    // 4. Push to Output if differences exist
    if (hasChanges) {
      const newInfoString = newBlocks.join("\n\n");
      const rtv = SpreadsheetApp.newRichTextValue().setText(newInfoString);
      
      namesToBold.forEach(name => {
        let startIdx = newInfoString.indexOf(`Guardian: ${name}`);
        if (startIdx !== -1) rtv.setTextStyle(startIdx + 10, startIdx + 10 + name.length, boldStyle);
      });

      reportOutput.push([osis, oldBlocks.join("\n\n"), newInfoString]);
      richTextNewInfo.push([rtv.build()]);
    }
  });

  // 6. Write to Sheet
  let reportSheet = ss.getSheetByName(reportSheetName);
  if (!reportSheet) {
    reportSheet = ss.insertSheet(reportSheetName);
  } else {
    reportSheet.clear();
  }

  // Handle case where NO overrides differ from Raw Data
  if (reportOutput.length === 1) {
    reportOutput.push(["-", "No differences found between Overrides and ATS Data.", "-"]);
    richTextNewInfo.push([SpreadsheetApp.newRichTextValue().setText("-").build()]);
  }

  const targetRange = reportSheet.getRange(1, 1, reportOutput.length, 3);
  targetRange.setValues(reportOutput);

  // Overwrite ONLY Column C starting at Row 2 with the Bolded Rich Text
  if (richTextNewInfo.length > 0) {
    reportSheet.getRange(2, 3, richTextNewInfo.length, 1).setRichTextValues(richTextNewInfo);
  }

  // 7. Visual Formatting
  reportSheet.setColumnWidth(1, 100); 
  reportSheet.setColumnWidth(2, 400); 
  reportSheet.setColumnWidth(3, 400); 
  
  targetRange.setWrap(true).setVerticalAlignment("middle");
  
  reportSheet.getRange("A1:C1")
    .setFontWeight("bold")
    .setBackground("#f3f3f3")
    .setHorizontalAlignment("center");

  reportSheet.setFrozenRows(1);
  reportSheet.showSheet(); 
  reportSheet.activate(); 

  // Lock the report like the other built tabs (Sheet Lock.gs)
  try { applySheetLocks_(ss, [reportSheetName]); } catch (e) { console.warn("Sheet lock skipped: " + e.message); }

  ss.toast("Delta Override Report generated successfully!", "Report Ready", 5);
}

// ======================================================================
// SECTION: END OF YEAR ROLLOVER (was RolloverBackend.gs)
// Handles End-of-Year archiving by freezing triggers, renaming, and exporting PDF
// ======================================================================

/**
 * Opens the Rollover Wizard UI.
 */
function openRolloverWizard() {
  const html = htmlFromFile_('RolloverWizard')
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
  // Declared outside the try so the catch block can clean up the temp
  // spreadsheet if anything fails after it's created.
  let tempSsId = null;
  let usedFallback = false;
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
    tempSsId = tempSs.getId();
    
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

    // Your Lists (owner 2026-11-12): each list tab goes into the archive PDF too. This sheet
    // keeps them, and "Upgrade and bring over last year's data" (Migration) copies them into
    // next year's sheet (od_migrateLists_).
    try {
      od_listSummaries_(ss).forEach(l => {
        const tab = ss.getSheets().filter(x => x.getSheetId() === l.tabId)[0];
        if (tab) tab.copyTo(tempSs).setName(l.tabName);
      });
    } catch (e) { console.warn("Rollover: Your Lists not archived: " + e.message); }
    
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
      usedFallback = true;
      pdfBlob = DriveApp.getFileById(tempSsId).getAs('application/pdf').setName(pdfName + ".pdf");
    }

    // Save PDF to Drive
    const savedPdf = parentFolder.createFile(pdfBlob);
    
    // Cleanup: Trash the temporary spreadsheet
    DriveApp.getFileById(tempSsId).setTrashed(true);
    
    return { success: true, fileUrl: savedPdf.getUrl(), isFallback: usedFallback };

  } catch (e) {
    // Cleanup on fatal error so we don't leave mess behind
    if (tempSsId) {
      try { DriveApp.getFileById(tempSsId).setTrashed(true); } catch (cleanupErr) { console.error(cleanupErr); }
    }
    return { success: false, error: e.message };
  }
}

// ======================================================================
// SECTION: MIGRATION WIZARD (was MigrationAssistant.gs)
// Safely migrates data from an older version of the CRM to the
// newest version, using Smart Mapping to handle column schema changes.
// ======================================================================
function openMigrationWizard() {
  const html = htmlFromFile_('MigrationUI')
    .setTitle("Bring Over Last Year's Tracker")
    .setWidth(550)
    .setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, "Bring Over Last Year's Tracker");
}

/**
 * The wizard's list: spreadsheets in the user's Drive that look like an older PC Tracker,
 * newest first (owner's request 2026-11-12, so nobody has to find and paste a link).
 * By name (the template names, "PC Tracker", Rollover's "Retired on"), owned by the user
 * when Drive allows that filter, this sheet left out. Never throws: [] = nothing found,
 * and the wizard just shows the paste box. Returns [{ name, url, updated }], at most 8.
 */
var MA_TRACKER_NAME_WORDS_ = ["Parent Coordinator", "PC Tracker", "PC Tracking", "Coordinator Tracker", "Retired on"];

function ma_findOldTrackers() {
  try {
    const here = SpreadsheetApp.getActiveSpreadsheet().getId();
    const names = MA_TRACKER_NAME_WORDS_.map(w => "title contains '" + w.replace(/'/g, "\\'") + "'").join(" or ");
    const query = "mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false and (" + names + ")";
    let files;
    try { files = DriveApp.searchFiles(query + " and 'me' in owners"); }
    catch (e) { files = DriveApp.searchFiles(query); } // an account where the owner filter isn't allowed
    const found = [];
    for (let n = 0; files.hasNext() && n < 60; n++) {
      const f = files.next();
      if (f.getId() === here) continue;
      found.push({ name: f.getName(), url: f.getUrl(), ms: f.getLastUpdated().getTime() });
    }
    found.sort((a, b) => b.ms - a.ms);
    const tz = Session.getScriptTimeZone();
    return found.slice(0, 8).map(f => ({ name: f.name, url: f.url, updated: Utilities.formatDate(new Date(f.ms), tz, "MMM d, yyyy") }));
  } catch (e) {
    console.warn("Looking for old trackers failed: " + e.message);
    return [];
  }
}

/**
 * STEP 1: Verification
 * Ensures the URL is valid and the user has access.
 */
function ma_verify(url) {
  try {
    const oldSS = SpreadsheetApp.openByUrl(url);
    return { success: true, name: oldSS.getName() };
  } catch (e) {
    return { success: false, error: "Could not open sheet. Ensure the link is correct and you have access." };
  }
}

/**
 * STEP 2: Recover Custom Settings (Smart Inference)
 * Because background properties are locked, we infer settings based on old sheet visibility!
 */
function ma_migrateSettings(url) {
  const oldSS = SpreadsheetApp.openByUrl(url);
  const props = PropertiesService.getDocumentProperties();
  
  const pPhone = oldSS.getSheetByName("Phone Contacts");
  const pDiv = oldSS.getSheetByName("Parents Divided");
  
  if (pPhone) {
      const isPhoneEnabled = !pPhone.isSheetHidden();
      props.setProperty('phoneContacts', isPhoneEnabled.toString());
      props.setProperty('syncToPhones', isPhoneEnabled.toString()); // Tie sync setting to sheet visibility
  }
  if (pDiv) props.setProperty('parentsDivided', (!pDiv.isSheetHidden()).toString());
  
  // Trigger a visual update on the new sheet so tabs hide/show correctly
  updateVisualSettings_(getSettings());
}

/**
 * STEP 3: Migrate Raw Data
 */
function ma_migrateRawData(url) {
  const oldSS = SpreadsheetApp.openByUrl(url);
  const newSS = SpreadsheetApp.getActiveSpreadsheet();
  
  const oldRaw = oldSS.getSheetByName("Raw Data");
  const newRaw = newSS.getSheetByName("Raw Data");
  if (!oldRaw || !newRaw) return;
  
  const data = oldRaw.getDataRange().getValues();
  newRaw.clearContents();
  if (data.length > 0) {
    newRaw.getRange(1, 1, data.length, data[0].length).setValues(data);
  }
}

/**
 * STEP 4: Migrate Standard "Simple" Sheets
 * (New/Edit Student, Send Out, Events)
 */
function ma_migrateSimpleSheets(url) {
  const oldSS = SpreadsheetApp.openByUrl(url);
  const newSS = SpreadsheetApp.getActiveSpreadsheet();

  // New/Edit Student lives in the Backend tab now; the old copy may still have
  // it as its own tab, and the reader handles both.
  const overrides = readBackendBlockFromOtherSpreadsheet_(oldSS, "OVERRIDES");
  if (overrides.length > 1) backendWriteBlock_(newSS, "OVERRIDES", overrides);
  
  const sheetsToCopy = ["Send Out", "Events"];
  
  sheetsToCopy.forEach(sheetName => {
    const oldSheet = oldSS.getSheetByName(sheetName);
    const newSheet = newSS.getSheetByName(sheetName);
    
    if (oldSheet && newSheet) {
      const data = oldSheet.getDataRange().getValues();
      if (data.length > 1) { 
        const maxRows = newSheet.getMaxRows();
        if (maxRows > 1) {
          newSheet.getRange(2, 1, maxRows - 1, newSheet.getMaxColumns()).clearContent();
        }
        
        if (newSheet.getMaxRows() < data.length) newSheet.insertRowsAfter(newSheet.getMaxRows(), data.length - newSheet.getMaxRows());
        if (newSheet.getMaxColumns() < data[0].length) newSheet.insertColumnsAfter(newSheet.getMaxColumns(), data[0].length - newSheet.getMaxColumns());
        
        newSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
      }
    }
  });

  // Your Lists (owner's pick 2026-11-12: they come along): each 📥 list tab + its setup.
  // The final rebuild (step 6) then fills their columns / Additional Information.
  try {
    const moved = od_migrateLists_(oldSS, newSS);
    if (moved) console.log("Migration: " + moved + " of Your Lists copied.");
  } catch (e) { console.warn("Migration: Your Lists not copied: " + e.message); }
}

// Contact Log Follow-up column (N) and the old values (trackers before 2026-11-04) -> the new options
const MIGRATION_FOLLOWUP_COL_ = 14;
const MIGRATION_FOLLOWUP_MAP_ = { "yes": "Followup Needed", "yes - important": "Followup Needed - ASAP", "no": "No Followup Needed" };

function migrateFollowupValue_(val) {
  const key = String(val == null ? "" : val).trim().toLowerCase().replace(/\s*-\s*/g, " - ");
  return MIGRATION_FOLLOWUP_MAP_[key] || val;
}

/**
 * STEP 5: The Smart Mapper for Contact Log
 * Handles dynamic column shifts and strips blank rows
 */
function ma_migrateContactLog(url) {
  const oldSS = SpreadsheetApp.openByUrl(url);
  const newSS = SpreadsheetApp.getActiveSpreadsheet();
  
  const oldLog = oldSS.getSheetByName("Contact Log");
  const newLog = newSS.getSheetByName("Contact Log");
  if (!oldLog || !newLog) throw new Error("Contact Log missing");
  
  // 1. Get Old Data & Find END
  const oldData = oldLog.getDataRange().getValues();
  let oldEndIdx = oldData.length;
  
  for (let i = 0; i < oldData.length; i++) {
    const colA = String(oldData[i][0]).trim().toUpperCase();
    const colE = String(oldData[i][4]).trim().toUpperCase(); 
    if (colA === "END" || colE === "END" || String(oldData[i][6]).trim().toUpperCase() === "END") {
      oldEndIdx = i;
      break;
    }
  }
  
  const oldHeaders = oldData[0].map(h => String(h).trim().toLowerCase());
  
  // SCRUBBER: Strip headers, END row, AND perfectly empty/ghost rows
  const isGhostValue = (val) => {
    if (val instanceof Date && val.getFullYear() < 1970) return true; // Destroy 1899 phantom dates
    const s = String(val).trim().toUpperCase();
    return s === "" || s.startsWith("#") || s === "FALSE" || s === "-";
  };

  const oldRecords = oldData.slice(1, oldEndIdx).filter(row => {
    // Check the first 10 core columns; if ANY of them contain real data, keep the row
    return row.slice(0, 10).some(val => !isGhostValue(val));
  });
  
  if (oldRecords.length === 0) return; // Nothing to migrate
  
  // 2. Get New Headers & Find END
  const newData = newLog.getDataRange().getValues();
  const newHeaders = newData[0].map(h => String(h).trim().toLowerCase());
  let newEndRow = 2; // Default
  
  for (let i = 0; i < newData.length; i++) {
    if (String(newData[i][0]).trim().toUpperCase() === "END" || 
        String(newData[i][4]).trim().toUpperCase() === "END" ||
        String(newData[i][6]).trim().toUpperCase() === "END") {
      newEndRow = i + 1; 
      break;
    }
  }
  
  // 3. Build Smart Mapping Dictionary
  const headerMap = {}; 
  oldHeaders.forEach((oldH, oldIdx) => {
    if (!oldH) return;
    const newIdx = newHeaders.indexOf(oldH);
    if (newIdx !== -1) {
      headerMap[oldIdx] = newIdx;
    }
  });
  
  // 4. Construct Perfectly Mapped Array
  const mappedRecords = [];
  oldRecords.forEach(oldRow => {
    const newRow = new Array(newHeaders.length).fill(""); 
    Object.keys(headerMap).forEach(oldIdx => {
      const newIdx = headerMap[oldIdx];
      let val = oldRow[oldIdx];
      
      // Fix stringified dates on the fly if this maps to a Date column
      if (newHeaders[newIdx] === "date" && typeof val === 'string' && val.includes(',')) {
        val = val.split(',')[1].trim();
        let parsedDate = new Date(val);
        if (!isNaN(parsedDate)) val = parsedDate;
      }
      
      // Final sweep to prevent stray 1899 ghost dates from entering the array
      if (val instanceof Date && val.getFullYear() < 1970) val = "";

      // Follow-up (N): last year's Yes / Yes - Important / No become today's options
      if (newIdx === MIGRATION_FOLLOWUP_COL_ - 1) val = migrateFollowupValue_(val);
      
      newRow[newIdx] = val;
    });
    mappedRecords.push(newRow);
  });
  
  // 5. Inject into New Sheet (UPGRADED: Starts exactly at Row 2, Leaves 1 Blank Row)
  const currentSpace = newEndRow - 2;
  const neededSpace = mappedRecords.length + 1; // 💡 +1 explicitly ensures a blank row is ALWAYS left above the END row

  if (currentSpace > 0) {
    // Erase the blank spacer rows
    newLog.getRange(2, 1, currentSpace, newLog.getMaxColumns()).clearContent();
  }

  // Adjust rows dynamically so the END row perfectly hugs the bottom (leaving 1 spacer)
  if (neededSpace > currentSpace) {
    const rowsToAdd = neededSpace - currentSpace;
    newLog.insertRowsBefore(newEndRow, rowsToAdd);
    newEndRow += rowsToAdd;
  } else if (neededSpace < currentSpace) {
    const rowsToRemove = currentSpace - neededSpace;
    newLog.deleteRows(2 + neededSpace, rowsToRemove);
    newEndRow -= rowsToRemove;
  }

  // Paste the migrated data starting exactly at row 2
  if (mappedRecords.length > 0) {
    newLog.getRange(2, 1, mappedRecords.length, newHeaders.length).setValues(mappedRecords);
  }
  
  // 💡 CRITICAL FORMATTING OVERRIDE:
  // Instead of letting Google Sheets guess the formatting (which might pull the brown 'END' bar upwards), 
  // we take the clean formatting from Row 2 and physically paint it down over everything, 
  // ensuring the new blank row perfectly matches the data rows above it.
  if (neededSpace > 1) {
    const formatSource = newLog.getRange(2, 1, 1, newLog.getMaxColumns());
    const formatTarget = newLog.getRange(3, 1, neededSpace - 1, newLog.getMaxColumns());
    formatSource.copyTo(formatTarget, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    formatSource.copyTo(formatTarget, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  }
  
  // ==========================================
  // --- SPACER ROW SANITIZATION ---
  // ==========================================
  const spacerRow = newEndRow - 1;
  
  // 1. Wipe Columns A through O completely clean of any cloned phantom text
  newLog.getRange(spacerRow, 1, 1, 15).clearContent().clearDataValidations();
  
  // 2. 💡 THE FIX: Target ONLY Column P (16). Wipe any "FALSE" text, then cleanly plant the checkbox.
  // Using the exact syntax provided to ensure a pure, working checkbox.
  newLog.getRange(spacerRow, 16).clearContent().clearDataValidations();
  newLog.getRange(spacerRow, 16).insertCheckboxes().setFontColor("#000000");

  // Repaint all background formulas instantly (Like the Week formula in A)
  // 🚨 FIX: Changed 'ss' to 'newSS' so it uses the correct defined variable!
  fillContactLogHstackFormulas_(newSS);
  
  // ==============================================================
  // --- NEW: AUTO-APPLY MASTER FILTER VISUALLY ---
  // ==============================================================
  try {
    // Ensure the END row itself is clean in Column P (16)
    newLog.getRange(newEndRow, 16).clearContent();
    
    // Clear any old tangled filters
    let filter = newLog.getFilter();
    if (filter) filter.remove();
    
    // Re-cast the fresh filter box encompassing the data up to Column P (16), showing
    // only rows whose checkbox font is black
    // 💡 newEndRow - 1 safely encapsulates the data AND the guaranteed blank spacer row!
    createContactLogHideFilter_(newLog, newEndRow - 1);
    SpreadsheetApp.flush(); // Lock the filter in first
    
  } catch(e) {
    console.error("Filter reset failed:", e);
  }
  
  SpreadsheetApp.getUi().alert(`Success! Migrated ${oldRecords.length} records into your Typed Table and applied your active Master Filter.`);
}

/**
 * STEP 6: Final Rebuild
 */
function ma_finish() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("Migration complete! Rebuilding master directories...", "Almost Done", 5);
  
  buildAllDerivedSheets();

  // 💡 NEW: Instantly Unlock and Repaint the App Menu without requiring a refresh!
  PropertiesService.getDocumentProperties().setProperty('FULL_MENU_UNLOCKED', 'true');
  try {
    onOpen();
  } catch(e) { console.error(e); }
  
  return { success: true };
}

// ======================================================================
// SECTION: VERSION CHECK (was VersionCheck.gs)
// On open, checks the master "Parent Coordinator Version Tracker"
// sheet and pops a changelog dialog when a NEWER-DATED template exists.
//
// HOW THE BASELINE WORKS (date-based):
//   • Each template copy carries its own build date in the hidden Backend
//     tab, cell A2. That date is stamped by resetToDisplayMode() in
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
//   2. Run resetToDisplayMode() so the template's Backend!A2 build date is
//      stamped to today (or set it by hand).
//   3. Insert a new row at the TOP of the tracker (Row 2) with today's Date,
//      its Size (Major/Small), Notes, and the /copy link.
//      (Only the newest filled-in Link matters — older links are ignored.)
// ======================================================================

// The master tracker spreadsheet (shared: Anyone with the link -> Viewer):
// VERSION_TRACKER_ID, in the 📌 YOUR LINKS box at the top of this file

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
  // Sheets set up before the change trigger existed get it now (only after setup)
  if (PropertiesService.getDocumentProperties().getProperty('SETUP_COMPLETE') === 'true') {
    ensureChangeTrigger_(null, true);
    claimNightlyTrigger_(); // runs as the person who ran setup: their nightly trigger is the sheet's one (Build.gs)
  }
  try {
    const shownVersionDialog = checkForUpdates_(false);
    // Only nag about stale ATS data when we didn't just open the version
    // dialog — a spreadsheet can show only one modal at a time.
    if (!shownVersionDialog) {
      checkRawDataFreshness_();
    }
    // Getting Started: open the checklist until it's done
    try { guideOnOpen_(); } catch (e) { console.warn("Getting Started checklist not opened: " + e.message); }
  } catch (e) {
    console.error("On-open checks failed: " + e.message);
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
// 📅 DATA UPDATE REMINDER
// Raw Data stamps each section's date when it is pasted (updateRawDataTimestamp_,
// Edits.gs; Input Data Mode stamps the same cells):
//   A1  "ATS - DATE UPDATED: MM/dd/yy"            (Student Data)
//   AG1 "Parent Square - DATE UPDATED: MM/dd/yy"
//   AP1 "NYCSA - DATE UPDATED: MM/dd/yy"
// If any of them (that has ever been loaded) is older than the reminder setting
// ("dataReminder": week / 2weeks / month / off, default week), a centered pop-up
// offers Input Data Mode. Checked on open (checkVersionOnOpen) and while the sheet
// is open (installedOnEdit, at most once an hour). Shown at most once a day;
// "Skip" then offers Remind After 2 Weeks / 1 Month / Remove Reminders, which
// change the setting (also in ⚙️ Settings) and wait that long before asking again.
// ==========================================
// "How to update your data" guide (Getting Started): GUIDE_LINKS.downloadDataDoc, top of this file.
// The old name still works for anything that used it.
const DATA_GUIDE_URL = GUIDE_LINKS.downloadDataDoc;

const DATA_REMINDER_DAYS = { week: 7, '2weeks': 14, month: 30 }; // 'off' = never
const DATA_REMINDER_SNOOZE_KEY = 'DATA_REMINDER_SNOOZE_UNTIL';     // ms; doc property only (not copied)
const DATA_REMINDER_SECTIONS = [
  { name: "Student Data (ATS)", cell: "A1" },
  { name: "Parent Square", cell: "AG1" },
  { name: "NYCSA", cell: "AP1" }
];

/** Current reminder setting: 'week' (default), '2weeks', 'month' or 'off'. */
function getDataReminderSetting_() {
  const v = PropertiesService.getDocumentProperties().getProperty('dataReminder');
  return (v === 'off' || DATA_REMINDER_DAYS[v]) ? v : 'week';
}

/** Sections whose last update is older than the setting: [{ name, date, days }]. */
function getStaleDataSections_(ss) {
  const setting = getDataReminderSetting_();
  if (setting === 'off') return [];
  const limit = DATA_REMINDER_DAYS[setting];
  const rawSheet = ss.getSheetByName("Raw Data") || ss.getSheetByName("RAW Data");
  if (!rawSheet) return [];
  const stale = [];
  DATA_REMINDER_SECTIONS.forEach(sec => {
    const text = String(rawSheet.getRange(sec.cell).getDisplayValue());
    if (/00\/00\/00/.test(text)) return;                 // never loaded: nothing to remind about
    const match = text.match(/(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/);
    const ms = match ? toDateMs_(match[1]) : NaN;
    if (isNaN(ms)) return;
    const days = Math.floor((Date.now() - ms) / 86400000);
    if (days >= limit) stale.push({ name: sec.name, date: match[1], days: days });
  });
  return stale;
}

/**
 * On open (and while open, see checkDataReminderWhileOpen_): shows the reminder
 * if a section is stale, reminders are on, not snoozed, not already shown today,
 * and Input Data Mode isn't running. Returns true if the pop-up was shown.
 */
function checkRawDataFreshness_() {
  const props = PropertiesService.getDocumentProperties();
  if (props.getProperty('SETUP_COMPLETE') !== 'true') return false;
  if (getDataReminderSetting_() === 'off') return false;
  if (props.getProperty('INPUT_DATA_MODE_STATE')) return false; // already updating data

  const snooze = Number(props.getProperty(DATA_REMINDER_SNOOZE_KEY) || 0);
  if (snooze && Date.now() < snooze) return false;

  const todayKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  if (props.getProperty('ATS_REMINDER_LAST_SHOWN') === todayKey) return false;

  const stale = getStaleDataSections_(SpreadsheetApp.getActiveSpreadsheet());
  if (!stale.length) return false;

  // Throttle to once per calendar day so we don't nag on every open.
  props.setProperty('ATS_REMINDER_LAST_SHOWN', todayKey);
  showDataReminderDialog_(stale);
  return true;
}

/** Called from installedOnEdit: re-checks at most once an hour while the sheet is open. */
function checkDataReminderWhileOpen_() {
  const cache = CacheService.getUserCache();
  if (cache.get('DATA_REMINDER_CHECKED')) return;
  cache.put('DATA_REMINDER_CHECKED', '1', 3600);
  checkRawDataFreshness_();
}

function showDataReminderDialog_(stale) {
  const esc = v => String(v).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const limitText = { week: "a week", '2weeks': "2 weeks", month: "a month" }[getDataReminderSetting_()];
  const rows = stale.map(s =>
    `<tr><td>${esc(s.name)}</td><td class="age">${s.days} days ago</td></tr>`).join("");

  const html = include_('Styles') + `
    <style>
      .note-dlg { padding: 8px 12px 12px; display: flex; flex-direction: column; gap: 12px; }
      .note-dlg p { line-height: 1.5; }
      .stale { width: 100%; border-collapse: collapse; font-size: 14px; }
      .stale td { padding: 7px 2px; border-bottom: 1px solid var(--line); }
      .stale td.age { text-align: right; color: var(--danger); font-weight: 600; font-variant-numeric: tabular-nums; }
      .foot { display: flex; justify-content: flex-end; gap: 8px; flex-wrap: wrap; margin-top: 4px; }
      .choices { display: flex; flex-direction: column; gap: 8px; }
      .working { display: flex; align-items: center; gap: 8px; color: var(--muted); }
    </style>
    <div id="ask" class="note-dlg">
      <h1>Time to update your data</h1>
      <p>This data hasn't been updated in over ${limitText}:</p>
      <table class="stale">${rows}</table>
      <div class="foot">
        <button type="button" class="btn sec" onclick="skip()">Skip</button>
        <button type="button" class="btn pri" onclick="go()">Go to Input Data Mode</button>
      </div>
    </div>
    <div id="later" class="note-dlg hidden">
      <h1>Skipped</h1>
      <p>When should we remind you next?</p>
      <div class="choices">
        <button type="button" class="btn sec wide" onclick="choose('2weeks')">Remind me after 2 weeks</button>
        <button type="button" class="btn sec wide" onclick="choose('month')">Remind me after 1 month</button>
        <button type="button" class="btn sec wide" onclick="choose('off')">Turn off reminders</button>
        <button type="button" class="btn pri wide" onclick="google.script.host.close()">Just skip for today</button>
      </div>
      <p class="hint">You can change this any time in ⚙️ Settings.</p>
    </div>
    <div id="busy" class="note-dlg hidden"><div class="working"><div class="spinner"></div>Working…</div></div>
    <script>
      function show(id) { ['ask','later','busy'].forEach(function (x) { document.getElementById(x).classList.toggle('hidden', x !== id); }); }
      function fail(e) { alert('Something went wrong: ' + (e && e.message ? e.message : e)); google.script.host.close(); }
      function go() { show('busy'); google.script.run.withSuccessHandler(function () { google.script.host.close(); }).withFailureHandler(fail).dataReminderOpenInputMode(); }
      function skip() { show('later'); }
      function choose(c) { show('busy'); google.script.run.withSuccessHandler(function () { google.script.host.close(); }).withFailureHandler(fail).dataReminderChoose(c); }
    </script>`;
  const height = Math.max(330, 250 + stale.length * 36); // the Skip screen needs ~330
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(440).setHeight(height), "Update Your Data");
}

/** Reminder pop-up: "📥 Go to Input Data Mode". */
function dataReminderOpenInputMode() {
  openInputDataMode();
  return true;
}

/**
 * Reminder pop-up after Skip: '2weeks' / 'month' = remind when data is that old
 * and not before that long from now; 'off' = no reminders. Saved like any setting
 * (document property + Backend settings block) so ⚙️ Settings shows it.
 */
function dataReminderChoose(choice) {
  if (choice !== 'off' && !DATA_REMINDER_DAYS[choice]) throw new Error("Unknown reminder choice: " + choice);
  const props = PropertiesService.getDocumentProperties();
  props.setProperty('dataReminder', choice);
  if (choice === 'off') props.deleteProperty(DATA_REMINDER_SNOOZE_KEY);
  else props.setProperty(DATA_REMINDER_SNOOZE_KEY, String(Date.now() + DATA_REMINDER_DAYS[choice] * 86400000));
  try { saveSettingsToVersionSheet_(); } catch (e) { console.error(e); }
  return true;
}

// ==========================================
// CORE FLOW
// ==========================================
function checkForUpdates_(isManual) {
  const props = PropertiesService.getDocumentProperties();

  // Never run before the user has authorized / completed setup
  if (props.getProperty('SETUP_COMPLETE') !== 'true') return false;

  if (!isManual) {
    // Respect "Dismiss Forever"
    if (props.getProperty('VERSION_DISMISS_FOREVER') === 'true') return false;

    // Throttle: only actually contact the tracker every VERSION_CHECK_HOURS
    const lastCheck = Number(props.getProperty('VERSION_LAST_CHECK') || 0);
    if (Date.now() - lastCheck < VERSION_CHECK_HOURS * 60 * 60 * 1000) return false;
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
    return false;
  }

  // Auto checks stay quiet when there's nothing new. A MANUAL check always
  // opens the dialog so the user can flip Major-only / all updates even when
  // already caught up (the dialog renders a friendly "caught up" state).
  if (!isManual && payload.entries.length === 0) return false;

  // Stash the payload so the dialog can grab it instantly when it loads.
  // (5 minute lifetime; the dialog recomputes automatically if it expires.)
  try {
    CacheService.getUserCache().put('VC_PAYLOAD', JSON.stringify(payload), 300);
  } catch (e) { console.error(e); }

  // Plain static HTML — the dialog fetches its data via google.script.run
  const title = payload.entries.length ? "New Version Available" : "Update Preferences";
  const html = htmlFromFile_('VersionDialog')
    .setWidth(540)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, title);
  return true;
}

// ----------------------------------------------------------------------
// THE VERSION TRACKER'S COLUMNS (owner 2026-11-12). First tab, header on row 1, one row
// per release: Version (date) | Size | Notes | Blank Copy | Example | Document. Columns are
// found by their header names, so moving or adding one in the Table breaks nothing; a
// header that isn't found falls back to that layout (A..F), but only onto an empty or
// "Column N" header (so the old "Blank Link" column is never read as the Example).
// Blank Copy = the template's /copy link ("Get the latest version", Rollover, "Make a new
// copy"); Example = a filled-in example tracker; Document = the User Guide. For each link
// the NEWEST dated row where it is filled in counts (trackerLinksFromRows_), so a release
// that keeps the same guide leaves Document blank. The template / guide folders are the
// backup (getTemplateLinks_ / getGuideLink_), then the Backend's saved copy.
// ----------------------------------------------------------------------
function trackerColumns_(header) {
  const h = (header || []).map(v => String(v).trim().toLowerCase());
  const generic = i => i < h.length ? (h[i] === "" || /^column \d+$/.test(h[i])) : true;
  const find = (test, fallback, exclude) => {
    const i = h.findIndex((x, k) => test(x) && (exclude || []).indexOf(k) === -1);
    return i >= 0 ? i : (generic(fallback) ? fallback : -1);
  };
  const example = find(x => /example/.test(x), 4);
  const guide = find(x => /document|guide/.test(x), 5);
  const blank = find(x => /blank|template|copy/.test(x) || x === "link", 3, [example, guide]);
  return { date: find(x => /version|date/.test(x), 0), size: find(x => /size/.test(x), 1),
           notes: find(x => /note/.test(x), 2), blank: blank, example: example, guide: guide };
}

/** Reads the Version Tracker once: { rows: [{ dateMs, rawDate, size, notes, link, example, guide }], cols }. Throws if it can't. */
function readVersionTracker_() {
  const tab = SpreadsheetApp.openById(VERSION_TRACKER_ID).getSheets()[0]; // Changelog must stay the FIRST tab
  const lastRow = tab.getLastRow();
  if (lastRow < 1) return { rows: [], cols: trackerColumns_([]) };
  // getValues (not display) so date-formatted cells come back as real Dates.
  const all = tab.getRange(1, 1, lastRow, Math.max(6, tab.getLastColumn())).getValues();
  const cols = trackerColumns_(all[0]);
  const cell = (r, c) => c >= 0 && c < r.length ? String(r[c] == null ? "" : r[c]).trim() : "";
  return {
    cols: cols,
    header: all[0].map(v => String(v).trim()),
    rows: all.slice(1).map(r => ({ dateMs: toDateMs_(r[cols.date]), rawDate: r[cols.date], size: cell(r, cols.size), notes: cell(r, cols.notes),
                                   link: cell(r, cols.blank), example: cell(r, cols.example), guide: cell(r, cols.guide) }))
  };
}

/** The newest dated row's link in each link column (rows in any order). */
function trackerLinksFromRows_(rows) {
  const isUrl = v => /^https?:\/\//i.test(v);
  const dated = rows.filter(r => !isNaN(r.dateMs)).sort((a, b) => b.dateMs - a.dateMs);
  const newest = key => { const r = dated.find(x => isUrl(x[key])); return r ? r[key] : ""; };
  return { blank: newest("link"), example: newest("example"), guide: newest("guide"), fromTracker: true };
}

/** trackerLinksFromRows_ + where each link column is ({ blank: 'D "Blank Copy"', ... }, "" = not found), for 🔗 My Links. */
function trackerLinksFrom_(tracker) {
  const links = trackerLinksFromRows_(tracker.rows);
  const where = i => i < 0 ? "" : String.fromCharCode(65 + i) + ((tracker.header || [])[i] ? ' "' + tracker.header[i] + '"' : "");
  links.cols = { blank: where(tracker.cols.blank), example: where(tracker.cols.example), guide: where(tracker.cols.guide) };
  return links;
}

/**
 * The three links from the Version Tracker, remembered 6 h (script cache). fresh = read it now.
 * Never throws: { blank: "", example: "", guide: "", fromTracker: false } when it can't be read
 * (no access, offline, or the simple onOpen, which may not open another spreadsheet).
 */
var TRACKER_LINKS_KEY_ = 'TRACKER_LINKS_V1';
function trackerLinks_(fresh) {
  const cache = CacheService.getScriptCache();
  if (!fresh) {
    try { const c = cache.get(TRACKER_LINKS_KEY_); if (c) return JSON.parse(c); } catch (e) { console.error(e); }
  }
  try {
    const links = trackerLinksFrom_(readVersionTracker_());
    try { cache.put(TRACKER_LINKS_KEY_, JSON.stringify(links), TEMPLATE_LINKS_CACHE_SECONDS); } catch (e) {}
    return links;
  } catch (e) {
    console.warn("Version Tracker links not read: " + e.message);
    return { blank: "", example: "", guide: "", fromTracker: false, error: e.message };
  }
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
    const tracker = readVersionTracker_();
    rows = tracker.rows;
    // The same read gives the links: remember them (trackerLinks_)
    try { CacheService.getScriptCache().put(TRACKER_LINKS_KEY_, JSON.stringify(trackerLinksFrom_(tracker)), TEMPLATE_LINKS_CACHE_SECONDS); } catch (e) {}
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
    const ms = rows[i].dateMs;
    if (isNaN(ms)) continue;        // Skip blank / junk / header-ish rows
    if (ms <= baselineMs) continue; // Already have it / already dismissed it
    missed.push({
      dateMs: ms,
      dateStr: formatTrackerDate_(rows[i].rawDate),
      size:  rows[i].size,
      notes: rows[i].notes,
      link:  rows[i].link   // the Blank Copy column: "Get the latest version"
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
 * build date (Backend!A2) and the newest date the user has dismissed.
 * Returns 0 (compare against epoch => surface everything) when neither is set,
 * so an un-stamped sheet still discovers updates instead of going dark.
 */
function getLocalBaselineDateMs_() {
  let buildMs = NaN;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const backend = ss.getSheetByName(BACKEND_SHEET_NAME);
    const legacy = ss.getSheetByName("Version");
    if (backend) buildMs = toDateMs_(backend.getRange(BACKEND_VERSION.BUILD_DATE).getValue());
    else if (legacy) buildMs = toDateMs_(legacy.getRange("A1").getValue());
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

  // --- Local build date (Backend!A2) ---
  const vsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BACKEND_SHEET_NAME);
  if (!vsheet) {
    lines.push("Backend tab: MISSING (build date unknown → baseline falls back to epoch).");
  } else {
    const v = vsheet.getRange(BACKEND_VERSION.BUILD_DATE).getValue();
    lines.push("Backend!A2 = \"" + v + "\"  [" + ((v instanceof Date) ? "Date" : typeof v) +
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
// The setting lives in the hidden Backend tab, cell B3 ("Yes"/"No"), so it
// travels with every copy and can also be edited by hand. To offer BOTH links
// at once, keep two template files: your master (B3 = No) and a copy with
// B3 = Yes.
// ======================================================================
function toggleAutoMigratePopup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateVersionSheet_(ss);
  const cell = sheet.getRange(BACKEND_VERSION.MIGRATE_TOGGLE);
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

// ======================================================================
// SECTION: TEMPLATE & GUIDE LINKS (was LinkResolver.gs)
// One source of truth for the shareable links, so you never hand-paste a file ID.
//
// WHERE EACH LINK COMES FROM (first one found wins; owner 2026-11-12)
//   Blank copy  1. Version Tracker column "Blank Copy" (D)   2. the template folder's
//               newest file with "(BLANK)" in its name ("(MIGRATE)" only when there is
//               no BLANK one)   3. the Backend's saved copy (E3)   4. the folder itself
//   Example     1. Version Tracker column "Example" (E)      2. the Backend's saved copy
//               (E6)   (no folder backup: no link = no "See an example tracker" button)
//   User Guide  1. Version Tracker column "Document" (F)     2. the guide folder's newest
//               file   3. the Backend's saved copy (E5)   4. the folder itself
//   (Version Tracker: trackerLinks_, the newest dated row with a link in that column.)
//
// WHEN THEY ARE LOOKED UP
//   Only when something needs one (📖 Open User Guide, Getting Started, the welcome
//   pop-up, Rollover), then remembered 6 hours (script cache). Looked up fresh by the
//   nightly run, 🔔 Check for Updates (it reads the tracker anyway), 🔗 My Links and
//   Reset to Display Mode. Never in the simple onOpen (it can't read other files).
//
// RESILIENCE (why copies still work)
//   Whatever is found is mirrored into the Backend Version section (D2:E6), which
//   travels with every copy of the sheet, so a copy that can't reach the tracker or
//   the folders still inherits the last-known links, exactly like saved settings do.
//
// SHARING REQUIREMENT
//   Other coordinators' copies need view access: share the Version Tracker, the
//   folders and the files "Anyone with the link -> Viewer". If you keep them private,
//   run refreshTemplateLinks() on your master template before sharing it and copies
//   will inherit the links.
// ======================================================================

// The Drive folders that hold the labeled template files (TEMPLATE_FOLDER_ID) and
// the User Guide (GUIDE_FOLDER_ID: its NEWEST file, a PDF or a Google Doc, is the
// "Open User Guide" link): both in the 📌 YOUR LINKS box at the top of this file.

// How long a successful folder scan is cached (script cache) before re-scanning.
const TEMPLATE_LINKS_CACHE_SECONDS = 6 * 60 * 60; // 6 hours

// Fallback links if a folder can't be resolved — land the user in the folder
// rather than on a dead link.
const TEMPLATE_FOLDER_URL = "https://drive.google.com/drive/folders/" + TEMPLATE_FOLDER_ID;

const GUIDE_FOLDER_URL = "https://drive.google.com/drive/folders/" + GUIDE_FOLDER_ID;

/**
 * Public entry point for the dialogs (called via google.script.run).
 * Returns { blankCopyUrl, migrateCopyUrl, exampleUrl, blankName, migrateName, folderUrl }.
 * Never throws — always returns usable (folder-fallback) URLs.
 */
function getTemplateLinks() {
  return getTemplateLinks_();
}

/**
 * Force a fresh look (Version Tracker, then both folders; bypassing the cache) and
 * re-store the results. No UI — safe to call from time-based triggers (the nightly sync).
 */
function refreshResolvedLinks_() {
  const cache = CacheService.getScriptCache();
  try { cache.remove('TEMPLATE_LINKS_V1'); } catch (e) {}
  try { cache.remove('GUIDE_LINK_V1'); } catch (e) {}
  try { cache.remove(TRACKER_LINKS_KEY_); } catch (e) {}
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
    "\n\nEXAMPLE:\n" + (links.exampleUrl || "(none)") +
    "\n\nGUIDE:\n" + (r.guide || "(unresolved)");
  console.log("Resolved links:\n" + msg); // Always in the execution log for testing.
  try {
    SpreadsheetApp.getUi().alert("🔗 Resolved Links", msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    // No UI context (pure editor run) — the console.log above already has it.
  }
  return r;
}

// ==========================================
// CORE
// ==========================================

/**
 * Cached, resilient resolver. Version Tracker first, then a live folder scan; on
 * success caches it (script cache + Backend D2:E6). On failure, falls back to the
 * Backend copy that traveled with this spreadsheet, then to the folder URL.
 */
function getTemplateLinks_() {
  const cache = CacheService.getScriptCache();
  try {
    const cached = cache.get('TEMPLATE_LINKS_V1');
    if (cached) return JSON.parse(cached);
  } catch (e) { console.error(e); }

  // 1. The Version Tracker's Blank Copy / Example columns (owner 2026-11-12); 2. the template
  // folder's newest (BLANK) file (backup); 3. the copy saved in the Backend; 4. the folder itself.
  const tracker = trackerLinks_();
  let links = tracker.blank
    ? { blankCopyUrl: tracker.blank, migrateCopyUrl: tracker.blank, blankName: "Version Tracker", migrateName: "Version Tracker" }
    : scanTemplateFolder_();
  links.exampleUrl = tracker.example || "";

  const inherited = loadTemplateLinksFromVersionSheet_();
  if (links.blankCopyUrl || links.migrateCopyUrl) {
    // The tracker read fine but has no Example = no button (don't bring back an old saved one)
    if (!links.exampleUrl && inherited && !tracker.fromTracker) links.exampleUrl = inherited.exampleUrl || "";
    // Found it — persist so plain copies inherit it.
    try { saveTemplateLinksToVersionSheet_(links); } catch (e) { console.error(e); }
  } else if (inherited) {
    // Tracker and folder unreachable (no access / offline) — use the inherited copy.
    links = inherited;
  }

  // Guarantee non-empty, non-dead URLs (the example has no folder fallback: no link = no button).
  if (!links.blankCopyUrl)   links.blankCopyUrl   = TEMPLATE_FOLDER_URL;
  if (!links.migrateCopyUrl) links.migrateCopyUrl = TEMPLATE_FOLDER_URL;
  links.exampleUrl = links.exampleUrl || "";
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
    // One template for everyone since 2026-11-12: its Initial Setup asks "Is this your first PC
    // Tracker?", so the upgrade link is the same (BLANK) file. A (MIGRATE) file still in the folder
    // is only used when there is no (BLANK) one.
    if (blank) { out.migrateCopyUrl = out.blankCopyUrl; out.migrateName = out.blankName; }
    else if (migrate) { out.migrateCopyUrl = copyUrlFor_(migrate); out.migrateName = migrate.getName(); }
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
 * Mirrors the resolved links into the Backend Version section, cells D2:E4.
 * Uses columns D:E so it never collides with the build date (A2), the
 * Auto-Migrate toggle (B3), or the saved-settings block (rows 5+, cols A:B).
 */
function saveTemplateLinksToVersionSheet_(links, ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet(); // ss: another spreadsheet (🐞 Make Release Copies)
  const sheet = getOrCreateVersionSheet_(ss);
  if (!sheet) return;
  sheet.getRange(BACKEND_VERSION.LINKS_ROW, 4, 3, 2).setValues([
    ["LINKS (automatic: from the Version Tracker, else the folders; don't edit)", ""],
    ["BLANK copy",   links.blankCopyUrl   || ""],
    ["UPGRADE copy (same file as BLANK)", links.migrateCopyUrl || ""]
  ]);
  const ex = sheet.getRange(BACKEND_VERSION.EXAMPLE_ROW, 4, 1, 2);
  if (links.exampleUrl) ex.setValues([["EXAMPLE copy", links.exampleUrl]]);
  else if (String(ex.getValues()[0][0]).trim() === "EXAMPLE copy") ex.clearContent();
}

/**
 * Reads the inherited links back out of the Backend Version section (E3/E4).
 * Returns null when nothing usable is stored.
 */
function loadTemplateLinksFromVersionSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const sheet = getOrCreateVersionSheet_(ss);
    const vals = sheet.getRange(BACKEND_VERSION.LINKS_ROW + 1, 5, 2, 1).getValues(); // E3:E4
    const blank = String(vals[0][0] || "").trim();
    const migrate = String(vals[1][0] || "").trim();
    const ex = sheet.getRange(BACKEND_VERSION.EXAMPLE_ROW, 4, 1, 2).getValues()[0]; // D6:E6
    const example = String(ex[0]).trim() === "EXAMPLE copy" ? String(ex[1] || "").trim() : "";
    if (!blank && !migrate && !example) return null;
    return { blankCopyUrl: blank, migrateCopyUrl: migrate, exampleUrl: example, blankName: "", migrateName: "" };
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
 * the Backend Version section (E5) so plain copies inherit it. Falls back to the inherited
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

  // 1. The Version Tracker's Document column (owner 2026-11-12); 2. the guide folder's newest
  // file (backup); 3. the copy saved in the Backend; 4. the folder itself.
  let url = trackerLinks_().guide || scanGuideFolder_();
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
  const files = guideFolderFiles_();
  return files.length ? files[0].url : "";
}

/**
 * The files directly in the guide folder, newest first, as the tracker ranks them:
 * a date in the NAME (MM/DD/YY, e.g. "User Guide 11/12/26") counts first, else the
 * file's last change (fileSortKey_). On a tie the file Drive lists last wins, as before.
 * [{ name, url, ms, byName }]; [] when the folder can't be read.
 */
function guideFolderFiles_() {
  try {
    const files = DriveApp.getFolderById(GUIDE_FOLDER_ID).getFiles();
    const out = [];
    for (let i = 0; files.hasNext(); i++) {
      const f = files.next();
      const name = f.getName();
      out.push({ name: name, url: f.getUrl(), ms: fileSortKey_(name, f), byName: /\d{1,2}\/\d{1,2}\/\d{2,4}/.test(name), i: i });
    }
    return out.sort((a, b) => (b.ms - a.ms) || (b.i - a.i));
  } catch (e) {
    console.error("scanGuideFolder_ failed (folder access?): " + e.message);
    return [];
  }
}

/** Mirrors the guide link into the Backend Version section (D5 label / E5 value). */
function saveGuideLinkToVersionSheet_(url, ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateVersionSheet_(ss);
  if (!sheet) return;
  sheet.getRange(BACKEND_VERSION.GUIDE_ROW, 4, 1, 2).setValues([["GUIDE link", url || ""]]);
  // A second "GUIDE link" row right under it is left from an older layout and nothing reads it: clear it
  const below = sheet.getRange(BACKEND_VERSION.GUIDE_ROW + 1, 4, 1, 2);
  if (String(below.getValues()[0][0]).trim() === "GUIDE link") below.clearContent();
}

/** Reads the inherited guide link back out of the Backend Version section (E5). */
function loadGuideLinkFromVersionSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const sheet = getOrCreateVersionSheet_(ss);
    return String(sheet.getRange(BACKEND_VERSION.GUIDE_ROW, 5).getValue() || "").trim(); // E5
  } catch (e) {
    console.error(e);
    return "";
  }
}
