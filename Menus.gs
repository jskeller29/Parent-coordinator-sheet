// ======================================================================
// FILE: Menus.gs
// PURPOSE: Menus, first-time setup, and the server side of every dialog:
// Settings, Student Overrides (+ report), End of Year Rollover, Migration
// Wizard, Version Check, and the template / User Guide links.
// ======================================================================

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
  { key: 'showMasterTable',    sheets: ["Master Table"],              def: false },
  { key: 'phoneContacts',      sheets: ["Phone Contacts"],            def: false },
  { key: 'parentsDivided',     sheets: ["Parents Divided"],           def: false },
  { key: 'notesTab',           sheets: ["Notes"],                     def: false },
  { key: 'showEvents',         sheets: ["Events"],                    def: false },
  { key: 'sendOut',            sheets: ["Send Out"],                  def: false },
  { key: 'showOverrideReport', sheets: ["Students Override Report"],  def: false },
  { key: 'showNewStudents',    sheets: ["New Students"],              def: false }
];

// The tabs shown by default. Everything else starts hidden; the user's saved
// settings reveal the optional tabs later via updateVisualSettings_.
var DEFAULT_VISIBLE_SHEETS = TAB_VISIBILITY
  .filter(function(t) { return t.def; })
  .reduce(function(all, t) { return all.concat(t.sheets); }, []);

/** Shows only the default-visible tabs and hides every other sheet. */
function applyDefaultSheetVisibility_(ss) {
  ss.getSheets().forEach(function(sheet) {
    if (DEFAULT_VISIBLE_SHEETS.indexOf(sheet.getName()) !== -1) sheet.showSheet();
    else sheet.hideSheet();
  });
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
function isAutoMigrateOn_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
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
      if (handlerName === "installedOnEdit" || handlerName === "nightlyMasterSync" || handlerName === "checkVersionOnOpen") {
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

  // 3. Create the Nightly Sync Trigger unconditionally (Runs at 11 PM every day)
  try {
    ScriptApp.newTrigger("nightlyMasterSync")
      .timeBased()
      .atHour(23)
      .everyDays(1)
      .create();
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

  ss.toast("Setup 100% Complete! The script is fully authorized and all automatic triggers are running.", "Success", 8);
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

  // OFFLINE SYNC DETECTOR
  const missedSync = props.getProperty('MISSED_SYNC_MSG');
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
    
    try {
      if (isAutoMigrateOn_()) {
        // Upgrade-template copy: point the user at Initial Setup, which kicks
        // off the migration flow (see runInitialSetupWrapper).
        SpreadsheetApp.getUi().alert(
          "🎉 Welcome to the newest version!",
          "Click '🚀 App Menu' at the top of the screen and select '🚨 Initial Setup (Run Once)' to start the migration process.",
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      } else {
        SpreadsheetApp.getUi().alert(
          "👋 Welcome to the PC Tracker!",
          "To get started, please click '🚀 App Menu' at the top of the screen and select '🚨 Initial Setup (Run Once)'. This will authorize the system and open your clickable User Guide!\n\nIf you prefer to view the guide right now, you can copy and paste this link into a new tab:\n" + GUIDE_FOLDER_URL,
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      }
    } catch(e) { console.error(e); }

    return;
  }

  // ==============================================================
  // WHAT SHOWS AFTER SETUP IS COMPLETE:
  // ==============================================================
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("Master Table");
  
  // Check if the sheets have ACTUALLY been built (Master Table data starts on Row 5)
  const isBuilt = masterSheet && masterSheet.getLastRow() > 4;

 if (!isBuilt) {
    // 🛑 INTERMEDIATE STATE: Setup done, but sheets haven't been built yet
    props.setProperty('FULL_MENU_UNLOCKED', 'false'); // Mark as locked
    // Using the EXACT SAME "🚀 App Menu" name allows us to overwrite the Setup button 
    // instantly without forcing the user to refresh the page!
    const getStartedMenu = ui.createMenu("🚀 App Menu"); 
    getStartedMenu.addItem("🔄 Import Old Data", "openMigrationWizard");
    getStartedMenu.addItem("📖 Open User Guide", "showWelcomeDialog");
    getStartedMenu.addItem("🐞 Report a Bug", "reportABug");
    getStartedMenu.addItem("🔔 Check for Updates", "manualCheckForUpdates"); // 🆕
    getStartedMenu.addSeparator();
    // Optional module (Z#InputDataMode.gs); a no-op without that file
    if (typeof addInputDataModeMenuItem_ === "function") addInputDataModeMenuItem_(getStartedMenu);
    getStartedMenu.addItem("🏗️ Build Sheets Only", "runBuildOnly"); 
    // Optional module (Z#Guide.gs, Getting Started checklist); a no-op without that file
    if (typeof addGuideMenuItems_ === "function") addGuideMenuItems_(getStartedMenu);
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
    // Optional module (Z#InputDataMode.gs); a no-op without that file
    if (typeof addInputDataModeMenuItem_ === "function") addInputDataModeMenuItem_(menu);
    menu.addSeparator(); 
    // Optional module (Z#Guide.gs, Getting Started checklist); a no-op without that file
    if (typeof addGuideMenuItems_ === "function") addGuideMenuItems_(menu);
    menu.addItem("⚙️ Settings", "openSettingsDialog"); 
    menu.addItem("📖 Open User Guide", "showWelcomeDialog");
    menu.addItem("🐞 Report a Bug", "reportABug");
    menu.addItem("🔔 Check for Updates", "manualCheckForUpdates"); // 🆕
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
      <p><b>Next steps:</b> click <b>🚀 App Menu</b> at the top of your screen to <b>Import Old Data</b> from a previous version, or close this window and paste your new student data into the <b>RAW Data</b> tab.</p>
      <div class="foot">
        <a class="btn pri" href="${guideUrl}" target="_blank" rel="noopener">Read the User Guide</a>
      </div>
      <div class="sub">
        <p class="help" style="margin:0">Need to share a fresh copy?</p>
        <div class="foot" style="justify-content:flex-start">
          <a class="btn sec" href="${links.blankCopyUrl}" target="_blank" rel="noopener">New copy (blank)</a>
          <a class="btn sec" href="${links.migrateCopyUrl}" target="_blank" rel="noopener">Upgrade copy (migrate)</a>
        </div>
      </div>
      <p class="hint">You can open this again any time from <b>🚀 App Menu ➔ 📖 Open User Guide</b>.</p>
    </div>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(460)
    .setHeight(420);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "PC Tracker User Guide");
}

// ==========================================
// 🐞 REPORT A BUG
// A menu click can't open a URL directly, so we pop a tiny dialog that
// auto-opens the bug-report form (with a clickable button as a fallback
// for browsers that block the automatic pop-up).
// ==========================================
function reportABug() {
  const formUrl = "https://forms.gle/5dXR2me3jt37aGhQ9";
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
    autoOverride: 'true',
    autoNotes: 'true',
    disableHiddenNotes: 'false',
    ignoreHideCheckboxes: 'false',
    emailNotifications: 'false',
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
    // Optional module (Z#Guide.gs): its welcome + Getting Started checklist instead
    if (typeof showGuideWelcome_ === "function") showGuideWelcome_();
    else showWelcomeDialog();
  }
}

// ==========================================
// 4. SHEET LOCK & CUSTOM WARNING
// ==========================================
function onEdit(e) {
  const props = PropertiesService.getDocumentProperties();
  
  if (props.getProperty('SETUP_COMPLETE') === 'true') {
    return; // Normal edits completely bypass the setup lock block below
  }

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
  'siteClass', 'typeCol', 'autoOverride', 'autoNotes', 'disableHiddenNotes',
  'ignoreHideCheckboxes', 'emailNotifications', 'disableRebuilds',
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

  // --- Master Table Filter (Hide Discharged) ---
  try {
    const masterSheet = ss.getSheetByName("Master Table");
    if (masterSheet) {
      const range = masterSheet.getDataRange();
      let filter = range.getFilter();
      
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
    console.warn("Could not update Master Table filter.");
  }
}

// --- TRIGGER MANAGEMENT ---
function manageNightlyTrigger_(disableSync) {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const syncFunctionName = 'nightlyMasterSync'; 

    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === syncFunctionName) {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }

    if (!disableSync) {
      ScriptApp.newTrigger(syncFunctionName)
        .timeBased()
        .atHour(2) 
        .everyDays(1)
        .create();
    }
  } catch (e) {
    console.warn("Trigger management failed: " + e.message);
  }
}

// --- PERMANENT PURGE FUNCTION ---
function purgeDischargedStudents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("Master Table");
  if (!masterSheet) return false;
  
  const data = masterSheet.getDataRange().getValues();
  if (data.length < 2) return 0; 
  
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
    
    masterSheet.clearContents();
    masterSheet.getRange(1, 1, currentData.length, currentData[0].length).setValues(currentData);
    
    const props = PropertiesService.getDocumentProperties();
    if (props.getProperty('hideDischarged') === 'true') {
      const newFilter = masterSheet.getDataRange().createFilter();
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
  const oldTodayOnly = props.getProperty('contactLogTodayOnly') === 'true';
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

  // 3. Update Visuals & Master Table Filter safely
  updateVisualSettings_(settings);
  
  // 3b. Contact Log filter window changed: rewrite the Column P rule and refresh the filter
  if (oldTodayOnly !== newTodayOnly) {
    try { refreshContactLogFilter_(); } catch (e) { console.warn("Could not refresh Contact Log filter: " + e.message); }
  }

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
 * Re-applies the Contact Log conditional formatting (which reads the filter-window
 * setting) and re-sets the Column P filter so rows re-hide/re-show right away.
 */
function refreshContactLogFilter_() {
  resetContactLogFormatting();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Contact Log");
  const filter = sheet ? sheet.getFilter() : null;
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
  const masterSheet = ss.getSheetByName("Master Table");

  // Dynamically map current Statuses from the Master Table
  const statusMap = new Map();
  if (masterSheet) {
    const mData = masterSheet.getDataRange().getValues();
    let osisIdx = 0, statIdx = 6; 
    for(let i=0; i < Math.min(10, mData.length); i++) {
        for(let c=0; c<mData[i].length; c++) {
            if(String(mData[i][c]).toUpperCase() === "OSIS") osisIdx = c;
            if(String(mData[i][c]).toUpperCase() === "STATUS") statIdx = c;
        }
    }
    for(let i = 4; i < mData.length; i++) {
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
    .setTitle('🔄 System Upgrade Wizard')
    .setWidth(550)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'System Upgrade Wizard');
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
    const shownVersionDialog = checkForUpdates_(false);
    // Only nag about stale ATS data when we didn't just open the version
    // dialog — a spreadsheet can show only one modal at a time.
    if (!shownVersionDialog) {
      checkRawDataFreshness_();
    }
    // Optional module (Z#Guide.gs): open the Getting Started checklist until it's done
    if (typeof guideOnOpen_ === "function") {
      try { guideOnOpen_(); } catch (e) { console.warn("Getting Started checklist not opened: " + e.message); }
    }
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
const DATA_GUIDE_URL = "https://docs.google.com/document/d/1iKzHa5Mh-K90qNLLcNtPllKPdqDzMH_4_Zjtrw-Nj6w/edit?usp=sharing";

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
  const hasInputMode = typeof openInputDataMode === "function"; // optional module (Z#InputDataMode.gs)
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
        ${hasInputMode
          ? `<button type="button" class="btn pri" onclick="go()">Go to Input Data Mode</button>`
          : `<a class="btn pri" href="${DATA_GUIDE_URL}" target="_blank" rel="noopener">How to update your data</a>`}
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
  if (typeof openInputDataMode === "function") openInputDataMode();
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
// One source of truth for the shareable template links, so you
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
//   the result into the Backend Version section (cols D:E), which travels with
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
 * Mirrors the resolved links into the Backend Version section, cells D2:E4.
 * Uses columns D:E so it never collides with the build date (A2), the
 * Auto-Migrate toggle (B3), or the saved-settings block (rows 5+, cols A:B).
 */
function saveTemplateLinksToVersionSheet_(links) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateVersionSheet_(ss);
  if (!sheet) return;
  sheet.getRange(BACKEND_VERSION.LINKS_ROW, 4, 3, 2).setValues([
    ["TEMPLATE LINKS (auto-managed — do not edit)", ""],
    ["BLANK copy",   links.blankCopyUrl   || ""],
    ["MIGRATE copy", links.migrateCopyUrl || ""]
  ]);
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

/** Mirrors the guide link into the Backend Version section (D5 label / E5 value). */
function saveGuideLinkToVersionSheet_(url) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateVersionSheet_(ss);
  if (!sheet) return;
  sheet.getRange(BACKEND_VERSION.GUIDE_ROW, 4, 1, 2).setValues([["GUIDE link", url || ""]]);
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
