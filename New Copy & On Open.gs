// When runInitialSetupWrapper() calls onOpen() to refresh the menu, we don't
// want onOpen's own "auto-open Migration Wizard" block to fire too — the
// wrapper opens it authoritatively right after. This global (shared across all
// .gs files in the one Apps Script scope) suppresses the duplicate for that
// single synchronous call. onOpen's block still fires on genuine reopens.
var IS_RUNNING_SETUP = false;

/**
 * Auto-Migrate mode is controlled by the hidden "Version" tab, cell B2:
 * "Yes" turns it on (upgrade-template copies), anything else is off
 * (new-user copies). The cell travels with every copy of the sheet, and
 * reading the bound spreadsheet's own cell is allowed even in the simple
 * onOpen trigger. Set it with the toggleAutoMigratePopup() dev tool or by
 * typing directly into B2.
 */
function isAutoMigrateOn_() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Version");
    if (!sheet) return false;
    return String(sheet.getRange("B2").getValue()).trim().toLowerCase() === "yes";
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
// 1. MENU OVERRIDE
// ==========================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();

  try {
    if (typeof applyQBlackTextFilter_ === "function") applyQBlackTextFilter_();
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
    // ℹ️ The Auto-Migrate toggle now lives in the "Version" tab cell B2, not
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
          "To get started, please click '🚀 App Menu' at the top of the screen and select '🚨 Initial Setup (Run Once)'. This will authorize the system and open your clickable User Guide!\n\nIf you prefer to view the guide right now, you can copy and paste this link into a new tab:\nhttps://docs.google.com/document/d/1iKzHa5Mh-K90qNLLcNtPllKPdqDzMH_4_Zjtrw-Nj6w/edit?usp=sharing",
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
    getStartedMenu.addItem("🔔 Check for Updates", "manualCheckForUpdates"); // 🆕
    getStartedMenu.addSeparator();
    getStartedMenu.addItem("🏗️ Build Sheets Only", "runBuildOnly"); 
    getStartedMenu.addToUi();

    // =======================================================
    // 🆕 AUTO-OPEN MIGRATION WIZARD (Upgrade-template copies)
    // Fires on every open until the sheets are built. Controlled by the
    // hidden "Version" tab, cell B2 = "Yes" (see isAutoMigrateOn_).
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
    menu.addSeparator(); 
    menu.addItem("⚙️ Settings", "openSettingsDialog"); 
    menu.addItem("📖 Open User Guide", "showWelcomeDialog");
    menu.addItem("🔔 Check for Updates", "manualCheckForUpdates"); // 🆕
    
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
  }

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
  const htmlContent = `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 10px; text-align: center;">
      <h2 style="color: #0F172A; margin-top: 0;">Welcome to the PC Tracker!</h2>
      <p style="color: #334155; font-size: 14px; margin-bottom: 25px; line-height: 1.5; text-align: left;">
        This system is designed to automate your directory, track your parent interactions, and easily compile your PCAR data. 
        <br><br><b>Next Steps:</b> Click <b>🚀 App Menu</b> at the top of your screen to safely <b>Import Old Data</b> from a previous version, or close this window and paste your new student data directly into the <b>RAW Data</b> tab!
      </p>
      
      <a href="https://docs.google.com/document/d/1iKzHa5Mh-K90qNLLcNtPllKPdqDzMH_4_Zjtrw-Nj6w/edit?usp=sharing" target="_blank" style="display: inline-block; padding: 12px 24px; background-color: #2196F3; color: white; text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
        📖 Read the User Guide
      </a>
      
      <p style="margin-top: 30px; font-size: 12px; color: #94a3b8;">
        You can always access this guide later by clicking<br><b>🚀 App Menu ➔ 📖 Open User Guide</b> at the top of the screen.
      </p>
    </div>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(420)
    .setHeight(290);
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "📖 PC Tracker User Guide");
}

// ==========================================
// 1.2. SIDEBAR LAUNCHERS & RUN MODES
// ==========================================
function runBuildOnly() {
  CacheService.getScriptCache().put('RUN_MODE', 'BUILD_ONLY', 600);
  const html = HtmlService.createHtmlOutputFromFile('ProgressSidebar')
      .setTitle('Build Sheets Manager')
      .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

function openProgressBarSidebar() {
  CacheService.getScriptCache().put('RUN_MODE', 'FULL_SYNC', 600);
  const html = HtmlService.createHtmlOutputFromFile('ProgressSidebar')
      .setTitle('Build & Sync Manager')
      .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ==========================================
// 1.5. SETUP WRAPPER 
// ==========================================
function runInitialSetupWrapper() {
  if (typeof setupAllInstallableTriggers === "function") setupAllInstallableTriggers();
  const props = PropertiesService.getDocumentProperties();
  
  props.setProperty('SETUP_COMPLETE', 'true');
  
  // Save this exact sheet's ID so the Clone Detector knows this is the original!
  props.setProperty('ORIGINAL_SHEET_ID', SpreadsheetApp.getActiveSpreadsheet().getId());
  
// =======================================================
  // NEW: FORCE DEFAULT SETTINGS DIRECTLY (Bypasses backend functions!)
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
    emailNotifications: 'false'
  };
  
  // Write variables straight into memory to eliminate any TypeError possibilities.
  for (const [key, value] of Object.entries(defaultSettings)) {
    props.setProperty(key, value);
  }

  // =======================================================
  // NEW: HIDE UNNECESSARY SHEETS
  // =======================================================
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Included both "RAW Data" and "Raw Data" to be safe against case variations
  const visibleSheets = ["RAW Data", "Raw Data", "Contact Log", "Combined Contact Tracking", "PCAR", "Directory"];
  
  ss.getSheets().forEach(sheet => {
    if (visibleSheets.includes(sheet.getName())) {
      sheet.showSheet();
    } else {
      sheet.hideSheet();
    }
  });
  
  // Instant UI swap without refresh! Suppress onOpen's own auto-migrate open
  // for this synchronous call so the wizard isn't opened twice — the
  // authoritative open happens in the if/else right below.
  IS_RUNNING_SETUP = true;
  onOpen();
  IS_RUNNING_SETUP = false;

  // =======================================================
  // 🆕 UPGRADE TEMPLATE vs NEW-USER TEMPLATE
  // Only one modal can be open at a time, so we pick the right one:
  //   Version!B2 = "Yes" -> jump straight into the Migration Wizard
  //   Version!B2 = "No"  -> show the normal Welcome guide
  // =======================================================
  if (isAutoMigrateOn_()) {
    openMigrationWizard();
  } else {
    // POP UP THE WELCOME DIALOG NOW THAT THE USER HAS CLICKED!
    showWelcomeDialog();
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
// 5. RESILIENT SIDEBAR BACKEND
// ==========================================
function startResilientSync() {
  const cache = CacheService.getScriptCache();
  const currentState = cache.get('SYNC_UI_STATE');
  
  if (currentState) {
    const phase = JSON.parse(currentState).phase;
    if (phase !== 'idle' && phase !== 'done' && phase !== 'killed') {
      return; 
    }
  }

  cache.put('SYNC_UI_STATE', JSON.stringify({ phase: 'building', progress: 0, total: 0 }), 21600);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (typeof buildAllDerivedSheets === "function") buildAllDerivedSheets();
  if (typeof formatDirectoryAuto_ === "function") formatDirectoryAuto_(ss.getSheetByName("Directory"));
  const psSheet = ss.getSheetByName("Parents Divided");
  if (psSheet) psSheet.hideSheet();

  const phoneSheet = ss.getSheetByName("Phone Contacts");
  const total = phoneSheet ? Math.max(0, phoneSheet.getLastRow() - 1) : 0;

  cache.put('SYNC_UI_STATE', JSON.stringify({ phase: 'syncing', progress: 0, total: total }), 21600);
  
  // ==============================================================
  // ✨ FIRST BUILD UNLOCK DETECTOR ✨
  // Triggers ONLY when the builder successfully finishes the very first time!
  // ==============================================================
  const props = PropertiesService.getDocumentProperties();
  if (props.getProperty('SETUP_COMPLETE') === 'true' && props.getProperty('FULL_MENU_UNLOCKED') !== 'true') {
    props.setProperty('FULL_MENU_UNLOCKED', 'true');
    try { onOpen(); } catch(e) { console.error(e); } // Attempt silent UI refresh
    
    // Send a long-lasting toast notification to refresh the page
    ss.toast("All sheets are built! Please refresh your webpage (Press F5) to explore and see your new menu options.", "🎉 Build Complete!", 15);
  }

  syncToGoogleContacts();
}

function getSyncStatus() {
  const cache = CacheService.getScriptCache();
  const stateStr = cache.get('SYNC_UI_STATE');
  return stateStr ? JSON.parse(stateStr) : { phase: 'idle' };
}

function resetSyncStatus() {
  CacheService.getScriptCache().remove('SYNC_UI_STATE');
}
