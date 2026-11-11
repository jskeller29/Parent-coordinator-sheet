This is a Google Apps Script project (a Google Sheets add-on for student/roster management).
- .gs files are Apps Script server-side code (JavaScript that runs on Google's servers)
- .html files are dialog/sidebar UIs served via HtmlService
- Frontend HTML talks to backend .gs functions via google.script.run
- There is no build step and no modules — all .gs files share one global scope

Branches:
- main: the live code BEFORE the cleanup (old file layout, 22 .gs files), incl. the Z# Contact Log rename.
- Test-Version: main + the cleanup (plan stages 1-3 below) + the owner's 🐞 Debug menu. It stays here until the
  owner has tested it on a copy of the sheet and says to merge it into main (it is ahead of main, so the merge is
  a fast-forward).
- Old claude/* session branches are deleted once everything in them is in main.

Files (shared = every school's sheet):
- Core.gs: Contact Log column constants, findEndRowInColumnE_, Student List lookup (getStudentListSheet_), Backend
  tab helpers, owner hook point (getOwnerHooks_, runAction_), small text / phone / OSIS / date helpers.
- Build.gs: buildAllDerivedSheets + everything it builds (Student List, Combined Contact Tracking, Parents Divided,
  Directory, Phone Contacts), new / discharged student detection, the notes-line helpers, nightlyMasterSync, build
  sidebar endpoints. BUILD_STEPS + reportBuildStep_ drive the sidebar's "Building Sheets & Directory" checklist
  (cache key BUILD_STEP; getSyncStatus returns buildSteps / buildStep with the state).
- Edits.gs: installedOnEdit, ghost typists (Contact Log / Notes / Events), Raw Data paste clean-up, quick updates
  (updateSingleOsisDelta_ / updateOsisDeltas_ / updateContactAndPhoneOnly), Events tables.
- Menus.gs: onOpen menus, setup, Settings, Student Overrides + report, Rollover, Migration, Version Check, links.
- Contacts Sync.gs: Google Contacts sync (syncToGoogleContacts, searchAndUpsertSingleContact), Full Sync endpoints,
  EMERGENCY_STOP_SYNC / RESTART_SYNC.
- Sheet Lock.gs: "baby-proofing" (see SHEET LOCK below): Google warning-only protections + toggleSheetLock
  (App Menu "🔓 Unlock Sheet" / "🔒 Lock Sheet").
- Regenerate.gs: reset / fake-data tools (resetToDisplayMode, resetAndLoadFakeData) and Contact Log formatting.
  Owner asked to keep it as is (not moved, not cleaned); change it only when the owner asks (e.g. KEEP_TABS_).
  Student List rename (owner's request): KEEP_TABS_ says "Student List", deleteLegacySheets_ calls
  getStudentListSheet_ first (an old "Master Table" is renamed, not deleted) and resetMasterTable_ uses it.
- Styles.html: the one shared look ("District": navy #183e69, Source Sans 3 text, Source Serif 4 headings,
  checkbox-style switches, 3px corners) for every dialog and sidebar. HTML files pull it in with
  <?!= include_('Styles') ?> and are opened with htmlFromFile_(name) (Core.gs, = createTemplateFromFile().evaluate());
  dialogs built as strings in .gs (Welcome, Report a Bug, ATS reminder, Create Event) start with include_('Styles').
  No other <? ... ?> tags may appear in the HTML files. Z#ContactLogSidebar keeps its own style (left as is).
- Z#ContactLog.gs (+ Z#ContactLogSidebar.html): self-contained optional "📇 Contact Log" tool; left as is.
  "Z#" = a temporary file for trying out a new feature: no other code may rely on Z# files (the only link is the
  guarded `if (typeof buildContactLogMenu_ === "function")` line in onOpen, a no-op without the file).
- Z#InputDataMode.gs (+ Z#InputDataSidebar.html): optional "📥 Input Data Mode" (🚀 App Menu, both the getting-started
  and full menus, via 2 guarded `addInputDataModeMenuItem_` lines in onOpen). Hides all visible tabs (remembered),
  then one temp tab at a time ("Student Data", "Parent Square Data", "NYCSA Data"); paste anywhere or upload the
  CSV/Excel file (SheetJS from cdnjs, read in the sidebar). Sidebar polls every 2s, finds the header row, refuses the
  wrong export and numbers Excel cut off (E+ / #### / ID-phone columns ending in 000; "Use it anyway" allowed),
  then replaces that Raw Data section (rows 7+, same headers/stamps as cleanRawDataPasteZone_ /
  updateRawDataTimestamp_). Skip keeps old data. Step 4 "Site / Class" (manual, Save button): Class | Teacher Name |
  Site (Optional), one row per ATS class code (Raw Data R) + classes already in BH:BI, prefilled from BH:BI; Site
  dropdown = hidden col Z formula of every typed site (other values allowed). Save rewrites BH:BI rows 7+ as ONE ROW
  PER SITE (owner's format): BH = site, BI = that site's classes, one per line, "Y25: Carl" or bare "Y24" (no
  teacher); sites in number order, no-site classes last. Read by parseClassCell_ (Build.gs, used by
  buildClassReferenceMap_ and Z#): one class per line, a line may still hold the old "Y41: Tessa Y25: Sam" (read
  exactly as before; bare-code lines are the only addition). ":" stripped from names, classes with no teacher and
  no site dropped, codes must be letters/digits. Raw Data rows 7+ are forced to 21px (lockRawDataRowHeights_, Edits.gs:
  after every Raw Data edit/paste and Input Data Mode write) so multi-line BI cells do not make rows tall; date in doc property
  INPUT_DATA_MODE_SITECLASS_UPDATED. Progress bar at the top (idm_goToStep) jumps to any step; the one temp tab is
  renamed + reset for it; finishing a step goes to the next index (the last one finishes). Exit always asks: Save & exit
  (idm_saveAndExit: saves Site / Class or moves an unmoved paste first; a problem keeps the mode on) / Exit without
  saving / Cancel. State in doc property
  INPUT_DATA_MODE_STATE; onOpen toasts a reminder if the mode was left on.
  Never deletes the tab the user is on (Apps Script then throws "Sheet <id> not found" on every later call): the one
  temp tab is cleared + renamed between steps; at the end it is deleted by a separate sidebar call
  (idm_deleteLeftoverTab, id in INPUT_DATA_MODE_DELETE_TAB; also retried on open). IDM_DEBUG = temporary debug panel.
  Step 5 "Your Lists" (step key lists, lists: true; needs Z#OwnData.gs, every od_* call guarded): no paste / upload /
  polling / Skip; the lists card shows od_listSummaries_ (idm_view_ "lists", null = files missing -> note, Finish still
  works) with Paste new version / Edit setup / ✏️ rename / 🗑️ delete, "➕ Add another list" (od_createList opens
  Z#OwnDataSidebar IN PLACE of this sidebar; od_backToInputMode reopens it) and Finish (idm_finishLists:
  results.lists = { saved, rows: ready lists, lists: true } or skipped, so the end screen offers Build). Site / Class
  is no longer the last step: its button is "✅ Save Site / Class".

- Z#OwnData.gs (+ Z#OwnDataSidebar.html): optional "Your Lists" (owner's request 2026-11-05): your own lists (bus
  routes, PTA, trip forms) matched to students or guardians; reached only from Input Data Mode step 5 (no menu item).
  One tab per list "📥 <name>" (rows 1-3 = OWN DATA band / status / column chips, header row 4 after a paste, data from
  row 5, tool columns "✔ Matched" / "How" after one blank spacer column). Its setup = one row in the Backend OWN_LISTS
  section (Tab ID, List, Setup JSON, Updated; od_loadLists_ / od_saveLists_): by (student / guardian), roles (osis,
  student / guardian as {col} or {first, last}, phone, email, class), vars (label, show = Additional Information,
  column = Student List / Parents Divided column, blank / notOnList = nothing | text), picks (row hash -> OSIS /
  guardian key / "skip"), headers, title, stats.
  - Paste: od_takePaste_ finds the header row (od_findHeaderRow_; title rows above it are remembered), moves it to
    row 4, data as plain text. A re-paste with the same headers keeps the setup (od_rolesByHeader_) and the picks.
  - Matching (od_matchStudentRow_ / od_matchGuardianRow_): your pick, then OSIS (normalizeOsis_), exact name token
    set (either name order), phone / email via guardians (Raw Data incl. overrides), then best guesses (short-name
    prefix of 3+ letters, OD_NICKNAMES_, last initial "Diana P.", a guardian's name -> their children); 2+ candidates =
    "pick one", none = "not found". Index = Student List A5:E + Raw Data guardians (od_indexFromSheet_).
  - Sidebar: paste view polls od_wizardCheck every 2 s; 1 Columns (dropdowns 'Column C “OSIS”', "Two columns (first +
    last name)…", OSIS hint "5 of 6 filled cells are OSIS in your Student List"), 2 Variables (keep / name / blank /
    not on list / where, live "How it reads" preview, privacy warning), 3 Match (od_matchPreview; "✓ Accept all N best
    guesses" with a warning, pick / search picker / skip / undo). Save = od_saveList (writes the tool columns + stats).
    Done view: during Input Data Mode only "← Back to Input Data Mode" (a build then would run with the tabs hidden);
    outside it "🏗️ Build Sheets Now" (od_build).
  - Build: buildAllDerivedSheets calls od_buildOutputs_(ss, masterTable, guardianRows) right after
    addLastInteractionColumn_ (guarded typeof + try/catch: an error only logs "Your Lists skipped"). It re-matches every
    ready list, rewrites the tool columns, returns studentCols (Student List columns after Last Interaction, siblings'
    values joined " / ", addListColumns_) and guardianCols (Parents Divided, keyed OSIS|guardian name,
    addGuardianListColumns_), and rewrites the Backend ADDL_INFO section (OSIS, Student, Guardian, Additional
    Information; Guardian blank = the whole student). The not-on-list text goes to everyone not on the list.
  - Phone Contacts: readAdditionalInfo_ / additionalInfoText_ (Build.gs) add "\n\nAdditional Information:\n<lines>"
    before "Communications:" (shared) / before "Note:" and in the Old Master Table history block (owner,
    ownerBuildPhoneContactsData_). No lists = no block = everything exactly as before (old-vs-new offline run identical
    apart from the 2 new Backend titles on a NEW Backend tab). Combined Contact Tracking is not touched.
    readAdditionalInfo_ keeps the rows in the doc cache (ADDL_INFO_ROWS_<spreadsheet id>, 6 h) because the owner's
    per-edit Phone Contacts update (ownerUpdateOsisDeltas_) calls it; backendClearBlock_ / backendAppendRows_ /
    backendBlockRange_ on ADDL_INFO clear it (a write clears first). Keyed by spreadsheet ID: the owner's "copy, then
    load fake data on the copy" builds the copy with this script's cache. A build without Z#OwnData.gs clears ADDL_INFO
    (no-op if never written), so removing the file also removes its lines.
  - Resets: list tabs are deleted by deleteLegacySheets_ (not in KEEP_TABS_); resetDerivedSheets_ clears ADDL_INFO +
    OWN_LISTS. Migration and Rollover do NOT carry lists over (not built; owner to decide).
  - google.script.run endpoints: od_createList, od_openList, od_repasteList, od_renameList, od_deleteList, od_getWizard,
    od_wizardCheck, od_readPaste, od_matchPreview, od_saveList, od_backToInputMode, od_build (+ idm_finishLists).
    Tested offline (own_test: shared + owner) and by a Chromium click-through of both sidebars wired to the real
    server code on the mock sheet (not yet tried on a real sheet).

- Z#Guide.gs (+ Z#GuideSidebar.html): optional "✅ Getting Started" (bare-bones version, owner's pick 2026-11-05; a full
  tour / tab pop-ups / Start Here tab / header notes were proposed and NOT built). 3 guarded hooks in Menus.gs:
  addGuideMenuItems_ (both 🚀 App Menus), showGuideWelcome_ (runInitialSetupWrapper, instead of showWelcomeDialog on
  new-user copies; "📖 Open User Guide" still opens showWelcomeDialog), guideOnOpen_ (checkVersionOnOpen).
  Welcome pop-up after setup: 4 steps, "Start with Input Data Mode" (only if openInputDataMode exists), video / guide
  links. Checklist sidebar: 1 load data (ticks when Raw Data A1 has a real date; also shows Parent Square AG1, NYCSA
  AP1, Sites and classes BH7:BI7 / INPUT_DATA_MODE_SITECLASS_UPDATED), 2 build (Student List lastRow > 4, same test as
  the menus), 3 settings (doc property ZG_SETTINGS_SEEN, set when opened from the checklist), 4 first contact (any
  Contact Log C above END). Opens by itself on every open until all 4 are done, not during Input Data Mode, unless
  unticked (doc property ZG_AUTO_OPEN = "off"). "This tab" panel (zg_activeTab polled every 3 s) + "Hidden tabs" list
  (Settings tabs: ⚙️ Settings → Tabs & columns; others: View → Hidden sheets). Link PLACEHOLDERS: GUIDE_LINKS at the
  top of Z#Guide.gs (empty = "link coming soon").

- Z#ContactLogDebug.gs (+ Z#ContactLogDebugSidebar.html): TEMPORARY "🐞 Contact Log Debug" sidebar (🚀 App Menu, both
  menus, guarded addContactLogDebugMenuItem_ lines in buildMenus_). Live list of runs (Contact Log / Notes / Events
  edits, row deletes) with each step's ms since the script started (⏱ ghost typist incl. data read vs write+paint,
  auto-fill, notes sync, Combined row written, search bar); "Check search bar" snapshot (settings, triggers, END,
  Communications cell + formula?, OSIS box, notes shown, Combined vs logs for that OSIS, newest rows); "Copy report";
  "Put formula back" (the "Use script" button went with the script-drawn box). "Anonymize" (on by default, cld_report(anon) / cldAnonymizer_): Student List names ->
  Student-N / Guardian-N / Name-N, OSIS -> OSIS-N, phones / emails masked, note texts -> "[N chars, K lines, #hash]"
  (cldNoteSummary_, same text = same hash), Contact Log D -> (filled)/(blank), unknown typed text -> [typed text].
  Recording = shared logger in Core.gs (clDbgStart_ / clDbg_ / clDbgEnd_ / clDbgRun_, wrapped
  around installedOnEdit / installedOnChange): on only while doc cache key CL_DEBUG_ON exists (the sidebar re-puts it
  every 2 s poll, 5-min TTL); runs in CL_DEBUG_LOG (newest 25). Off = one cache read per edit.

Owner-only file (in the repo, deployed ONLY to the owner's own sheet): z_MINE.gs
- Never copy it into other schools' sheets. It defines ownerHooks_(); shared code calls getOwnerHooks_() (Core.gs),
  which returns those hooks on the owner's sheet and null everywhere else. No shared function is redefined, so file
  load order does not matter. Shared code must work without z_MINE.gs; owner-only behavior (e.g. "Old Master Table" /
  "Old Phone Contacts") goes into a hook in z_MINE.gs, never into the shared files.
- Hooks: Student List status snapshot, Old Master Table history (Student List + Directory), Old Master Table OSIS
  that Combined Contact Tracking also keeps (oldMasterTableOsis), Old Phone Contacts merge,
  last-word name split (splitFullName_), tall Phone Contacts layout, quick update that also rewrites Phone Contacts,
  contacts sync (4 phones / 3 emails per row, group list without pageSize, no 50% delete check, summary email),
  Contact Log site column K (shared: H), Montserrat list, the New Students tab (ownerLogNewStudents_), 🐞 Debug menu +
  toasts.
- Shared Directory rule: draws only from Raw Data (no overrides, no discharged students). On the owner's sheet the
  hook then merges Old Phone Contacts and adds non-Raw-Data students from Old Master Table ONLY when its column F
  (Status) is "MASTER".
- 🐞 Debug menu (owner only): "Reset + Load Fake Data" and "Reset to Display Mode" (type RESET, then Regenerate's own
  confirm), "Copy Sheet, Then Reset the Copy" (makes a Drive copy
  in the same folder and runs resetToDisplayModeSteps_ on it), "Copy Sheet, Then Load Fake Data on the Copy" (runs
  loadFakeDataSteps_ on the copy via ownerRunOnCopy_: SpreadsheetApp.setActiveSpreadsheet(copy) because the build /
  buildEventTable / buildOverrideReport use the active spreadsheet, then switches back and restores this script's
  document properties with setProperties(saved, true), since the build's ATS snapshot lives there; refuses to run if
  the switch did not happen; sets OWNER_HOOKS_CACHE_ = null for the run so the copy is built like a
  shared sheet, because the owner deletes z_MINE.gs from the copy). resetToDisplayMode / resetAndLoadFakeData = confirm + these *Steps_ functions + alert;
  loadFakeDataSteps_ throws unless ss is the active spreadsheet. buildMenus_ also adds the owner menus before
  Initial Setup (a copy's setup flag is not copied) and "Toggle Debug" (doc property OWNER_DEBUG_TOASTS). (Color Tabs was removed at the owner's request.)
- All three resets call deleteLegacySheets_ (Regenerate.gs), which deletes EVERY tab not in KEEP_TABS_ (Raw Data /
  RAW Data, Contact Log, Notes, Events, PCAR, Send Out, Student List, Combined Contact Tracking, Parents Divided,
  Directory, Phone Contacts, Students Override Report, Backend), incl. Old Master Table / Old Phone Contacts, New
  Students (owner's choice: the owner's next build re-creates it) and the 📥 Your Lists tabs (resetDerivedSheets_ also
  clears the Backend ADDL_INFO / OWN_LISTS sections). It runs getBackendSheet_ first so old Version / Type List etc. tabs are migrated before they are deleted.
- All three resets also clear the Notes tab (resetNotes_, A:D rows 2+).
- Contact Log student lookup box. Owner's layout (made by hand 2026-11-11; small rows = 10 px): small row (Communications
  formula's name switch), OSIS | Student | Class, small row, guardian table ("Relationship" header + guardian rows),
  small row, Communications | Student Info (labels are HSTACK formulas; Communications reads the OSIS box with
  ADDRESS(ROW()-6, COLUMN()+1) from D = the OSIS box in column E, Student Info (column L, moved from K) with
  ADDRESS(ROW()-6, 5) = column E, Class reads the last Contact Log OSIS itself).
  Old layout: top row, OSIS | Student | "Student Information" title, Communications + Student Information text, spacer,
  guardian table, ending at END + 8. resetContactLog_ ends with fixContactLogLookupBox_ (Regenerate.gs): deletes up to
  11 rows right under the box (findContactLogLookupBox_: new layout = under the last Communications row; old layout /
  no box found: END+9..END+19 as before, which in the new layout would delete Communications; never a row the box
  uses), leftover "Parent Contact (COPY)" header rows (+ the row under it when blank / #REF!), then
  drawContactLogLookupBoxOutline_: thick black outline, top = row above "OSIS", left = "Communications" column, right =
  furthest of the Student Information / Student Info merge, the Communications-row merges and the guardian header's
  last filled cell, bottom = last Communications row (old layout: END + 8). It does not depend on the student shown
  (the old outline looked for the first empty guardian row right after the reset had emptied them, so it stopped at the
  guardian header). The reset does not move rows or change heights / inner lines (an automatic old -> new conversion
  was tried the same day and removed: the owner built the layout by hand). The box helpers (findContactLogLookupBox_,
  drawContactLogLookupBoxOutline_) are in Edits.gs (search bar section), since the build uses them too.
  Guardian rows (owner's request 2026-11-11): syncLookupBoxGuardianRows_ (Edits.gs) makes the guardian table as tall
  as the student with the most guardians in Parents Divided needs (A = OSIS, B not blank, as the table's FILTER; at
  least 2, at most 10), so a 3rd-5th guardian from New/Edit Student does not spill into the small row (#REF!). Run at
  the end of every build (saving / deleting an override runs one; syncLookupBoxGuardianRowsForBuild_) and by every
  reset (Parents Divided just emptied, so extra rows go). Guardian rows = header + 1 down to the small row(s) (<= 15 px)
  right above Communications (none found = not touched). Added rows go under the last guardian row: its height, no
  line between guardian rows, guardian 3 / 5 white, 4 cream #fff8ee (LOOKUP_BOX_ROW_COLORS_), the box's thick sides +
  thick line under the new last row. Extra rows are deleted when no student needs them (not if one holds a formula),
  thick line under the new last row. Each time, recountLookupBoxRowOffsets_ re-counts every ROW()-n / ROW()+n in the
  box (Communications 6 -> 7 for 3 guardians, Student Info the same). Tested only on a mock sheet so far.
- Send Out tab = a Sheets Table "Table1": A Date, B Send Out Title, C Send Out Type, D On Behalf Of, E Where Did You
  Send (dropdowns C:E), F Population, G # Reached, H Attachments (file chips); I "Behalf of List", J "Sending Location
  List" (kept). Dropdown options come from Backend BH3:BJ3 formulas (fixed options + values already typed:
  C Events / Services, D Parent Coordinator / Principal / PTA / School, E Backpack / DOE Messaging / Parents Square).
  Resets clear A:H rows 2+ (resetSendOut_, clearContent only); fake data writes 9 rows
  A:G column by column (loadFakeSendOut_; a refused column is skipped with console.warn).
- Contact Log Follow-up (N): options No Followup Needed / Followup Needed / Followup Needed - ASAP / Followed Up /
  Resolved / Skipped / Newer Interaction. The Contact Log is a Sheets Table and N a "Dropdown" typed column: Apps
  Script may not change its options ("not allowed on cells in typed columns"), so code never touches N's dropdown
  (set by hand in the Table; owner removed the reset step that tried). Autofill default for N = "No Followup Needed" (Edits.gs). Hide rule for P
  (getContactLogHideFormula_, owner's 2026-11-04 formula): =AND($B2<>"", INT($B2)<>TODAY(), TRIM($N2)<> Followup Needed / Followup Needed - ASAP / Need to respond / Waiting For Response /
  Waiting - Check Up). The date part follows Settings → Notes → "Contact Log shows" (contactLogTodayOnly, 2026-11-11 fix:
  it had stopped changing the rule): "Today only" = INT($B2)<>TODAY() (the owner's formula exactly); "This week and last
  week" (default, never saved) = B outside last week's Sunday .. this week's Saturday (same weeks as the A:B color bands).
  Used by every reset, the nightly run (resetContactLogFormatting) and every Settings save: refreshContactLogFilter_
  (Menus.gs) swaps ONLY the P custom-formula rule (rule.copy(), Type colors and other rules kept; added back if missing),
  writes nothing when it already matches, then re-sets the P filter. The old "Yes - Important" bold rule in resetContactLogFormatting is still there (now unused).
- Fake data (Reset + Load Fake Data, Regenerate.gs; owner's request 2026-11-04): 18 ATS students (incl. siblings
  Belcher Louise/Gene + override-only Tina, Addams Wednesday/Pugsley), Parent Square, NYCSA, Site/Class in the new
  BH:BI format (one row per site, classes one per line; 5 of 6 sites have 2-3 classes; "Y12"/"Y50" have no teacher),
  26 Contact Log rows (30-row log; each row carries its Type / Follow-up / Follow-up note), 3 events (Fall PTC,
  Graduation Workshop, Spring Festival; Backend Event Log rows now in the real column order), 5 Notes rows, 3
  New/Edit Student overrides (New Student Tina Belcher, Peter Parker + guardian, Diana Prince hidden from report),
  then the Students Override Report. PCAR gets no fake data (built from other tabs); Send Out: see below.
  With Debug on, runAction_ toasts ✅ / ⚠️ / ❌ after each edit, open, Build Only, Full Sync, nightly and contacts sync,
  including errors that are normally only logged (console.error / console.warn).

Deploying (manual paste into the Apps Script editor; master template + owner sheet):
- Delete the old files first or the project will not load (duplicate consts/functions): Backend, Build.gs Helpers,
  Build.gs beginning, Constants, Event, LinkResolver, MigrationAssistant, New Copy & On Open, NewStudents, Phone,
  Quick Updates Installed, QuickUpdates, RolloverBackend, Settings Backend, Student Override Report,
  Student Overrides Backend, VersionCheck, ZProduction, z_PhoneOverride, z_PhoneOverride 4+, z_PhoneOverride.gs Add On.
- Then paste Core, Build (replaces the old Build), Edits, Menus, Contacts Sync, Sheet Lock (new); owner sheet also z_MINE.
  Student List rename (2026-11-05): also Regenerate, SettingsDialog, ProgressSidebar, Z#Guide, Z#ContactLogDebug,
  Z#InputDataSidebar (they name the tab).
  UI refresh: also add Styles.html (new) and re-paste SettingsDialog, ProgressSidebar, NewStudentDialog,
  ManageOverridesDialog, RolloverWizard, MigrationUI, VersionDialog. Regenerate (re-paste: KEEP_TABS_ tab clean-up, Color Tabs removed), the Z# files
  and the manifest are unchanged. The Contact Log tool is now named
  Z#ContactLog / Z#ContactLogSidebar (delete the old #ContactLog / ContactLogSidebar files, or they define it twice).
- Input Data Mode: also paste Z#InputDataMode (script) + Z#InputDataSidebar (HTML); optional, menu item only appears
  with them.
- Your Lists (2026-11-05): paste the new Z#OwnData (script) + Z#OwnDataSidebar (HTML) and re-paste Core, Build,
  Regenerate, Z#InputDataMode, Z#InputDataSidebar (owner sheet: also z_MINE). Optional: without the two new files step 5
  says they are missing and every build / phone contact is exactly as before.
- Getting Started: also paste Z#Guide (script) + Z#GuideSidebar (HTML); optional, needs Styles.html; works best with
  Input Data Mode.
- Lookup box layout (2026-11-11): re-paste Regenerate, Edits, Z#ContactLogDebug, Build (the box itself is set up by
  hand in the sheet; the code finds it, cleans under it, outlines it and adds / removes guardian rows).

Student List tab (was "Master Table"; renamed 2026-11-05 at the owner's request):
- Every lookup goes through getStudentListSheet_ (Core.gs), which renames an old copy's "Master Table" tab the first
  time, so its Discharged history carries over. Never getSheetByName("Master Table") directly: writeSheet_ would make
  a second, empty tab next to it, and a reset would delete a tab missing from KEEP_TABS_. "Old Master Table" (owner
  only) is a different tab and keeps its name. Function / variable / key names were NOT renamed (buildMasterTableData_,
  masterTable, findInMasterTable_, Settings key showMasterTable).
- Lists built when the files load spell the name out (file load order must not matter): TAB_VISIBILITY
  ["Student List", "Master Table"], KEEP_TABS_, SHEET_LOCK_BUILT_TABS, Z#Guide's tab list, BUILD_STEPS. The search
  bar formula (SEARCH_BAR_COMMS_FORMULA) reads 'Student List'!A:B; restoreSearchBarCommunicationsFormula renames an old
  tab first. A tab renamed by hand / by setName keeps its protections, data validations and formula references.
- Header on row 4, data from row 5: OSIS, Student, Guardian, Site, Class, Sibling, Status, Last Interaction (H).
  Last Interaction (added 2026-11-05) = date of the newest Contact Log row (not hidden) or event for that student,
  with anyone (sibling rows: newest of the siblings), else "None logged". interactionsByOsis_ (Build.gs) is shared
  with Parents Divided's per-guardian Last Interaction (moved word for word). Added after the owner's Old Master Table
  hook, so every row is padded to the same width. Only a build / nightly updates it (like Parents Divided). The build
  reads the Events tab once at step 1 (buildBackendEventLogData_) so the column sees today's events; step 2 writes
  those same rows to the Backend Event Log.

Backend tab (see Core.gs):
- One hidden "Backend" tab replaces the old Version, New/Edit Student, Backend_Event_Log and Type List tabs.
  Row 1 = section titles; each old layout sits one row lower. A:E Version & Settings, G:AV New/Edit Student,
  AX:BF Event Log, BH:BJ Send Out list (BH3:BJ3 = the owner's formulas for the Send Out dropdown options; resets
  clear BH:BJ rows 3+ and write SEND_OUT_LIST_FORMULAS_ back via writeSendOutListFormulas_), BL Type List,
  BN:BQ Additional Information and BS:BV Your Lists setup (both Z#OwnData.gs). The last two are "grow" sections: an
  existing Backend tab gets their title the first time they are written (backendBlockReady_); before that reads return
  nothing and clears do nothing. A new / migrated Backend tab gets every title at once (writeBackendTitles_).
- Sections share rows: never deleteRow/insertRow, getDataRange/getLastRow, appendRow or writeSheet_ on it.
  Use the backend* helpers in Core.gs, which only touch one section's columns.
- Old copies migrate automatically the first time getBackendSheet_ runs.

New Students (detection: NEW / DISCHARGED STUDENTS section of Build.gs; tab: z_MINE.gs, owner's sheet only):
- Since 2026-11-04 the "New Students" TAB is created/updated only on the owner's sheet (ownerLogNewStudents_ hook,
  after the Directory). Shared sheets no longer create it (an old one stays but stops updating); the build report in
  the sidebar still lists new / discharged students on every sheet.
- Only Raw Data (ATS, A:AE) counts. Every build compares Raw Data's OSIS list with the previous build's
  (snapshot in document properties, tied to the sheet ID) and logs new / re-entered (from Discharged) students
  (Date Entered, Source = always "ATS", OSIS, Student, Guardian, Site, Class, Language).
  New/Edit Student and Old Master Table are ignored. Only appended to; a row is re-dated when the student re-enters.
- detectAtsChanges_ (shared, runs right after the Student List) is the single source for new / discharged students:
  the owner's New Students tab, the build report
  (sidebar) and the owner's nightly email all use it. Discharged = "Current" in the old Student List, in Raw Data at
  the last build, and no longer in Raw Data. The snapshot is saved at the end of the build.

UI REFRESH (2026-11-04, on Test-Version, not yet tried on a sheet): owner's picks, same functions / endpoints / saved values.
- Every dialog/sidebar uses Styles.html; Tailwind CDN removed; emoji removed inside dialogs and from dialog titles
  (menus, alert/confirm texts and menu paths like "🚀 App Menu" keep theirs). Browser confirm()/alert() kept.
- Tab visibility (owner's request 2026-11-05): TAB_VISIBILITY (Menus.gs) lists every tab but Backend (key, tab
  name(s), default). Shown by default: Contact Log, Combined Contact Tracking, PCAR, Directory (DEFAULT_VISIBLE_SHEETS
  is built from it, used at setup / resets). Old keys phoneContacts / parentsDivided / notesTab / sendOut kept; new keys
  showContactLog / showCombined / showPcar / showDirectory / showRawData / showMasterTable / showEvents /
  showOverrideReport / showNewStudents (never saved = default; in SETTING_KEYS). Settings checkboxes use data-tab="key";
  getSettings().tabsPresent hides rows for tabs the sheet doesn't have (New Students on shared sheets).
- Settings: 780x580, section list on the left (Contacts & sync / Tabs & columns / Live updates / Notes / Student data).
  "Rebuild sheets" dropdown = disableRebuilds + disableNightlySync (4 combos); notes timing buttons = autoNotes;
  "Fill in Date, # and Follow-up" = !disableContactAutofill; "Contact Log shows" = contactLogTodayOnly; date-range
  dropdown fills the same dates as the old preset buttons. Checked offline: 96 save / reset payloads identical to the old dialog.
- Build & Sync sidebar: live build checklist (from getSyncStatus's buildSteps / buildStep, see BUILD_STEPS in Build.gs;
  the current step shows its detail line) + phone-sync progress bar.
- Data update reminder (showDataReminderDialog_) restyled; Settings has it under Student data. Z#InputDataSidebar
  (trial file) keeps its own style.
- Create Event: type buttons (PTC, Workshop, Family event, Meeting) only fill the Type box. New/Edit Student: one grouped
  form, entry type as two buttons over a hidden #entryType input. Manage Overrides: cards. Rollover: step markers.

SHEET LOCK (Sheet Lock.gs; on by default, doc property SHEET_LOCK_OFF = "true" turns it off). WARNINGS ONLY (owner's
choice 2026-11-05): the code never puts cells back after an edit; the earlier undo layer (snapshots, onSelectionChange
copies, "clear 2+ rows?" popup, column A auto-restore) was removed because its extra script writes made Ctrl+Z need
several presses. Don't bring it back without asking.
- All are RANGE protections (owner's pick 2026-11-05: a whole-tab SHEET protection puts a 🔒 icon on the tab). Built
  tabs = one range over the whole grid; PCAR = 4 ranges around H2:I2. A range covers the grid as it was when locked;
  rows a build adds are covered when applySheetLocks_ runs at the end of the build. Old SHEET ones are removed as stale.
- Google "Heads up, are you sure?" warning-only protections (description starts "🔒 Sheet Lock"; scripts are not
  blocked) on: row 1 of Contact Log / Notes / Send Out; Contact Log END bar row (except O, the FILTER? checkbox) and
  everything below it; row 2 (headers) of every Events table (EVENT_TABLE_HEADERS); PCAR except H2:I2; the whole built
  tabs (Student List, Combined Contact Tracking, Parents Divided, Directory, Phone Contacts, Students Override Report,
  New Students). NOT on: Raw Data, Backend, owner-only tabs, Contact Log entry rows (incl. column A, so deleting one
  row never warns), Notes / Send Out rows 2+, Events row 1 and rows 3+.
- applySheetLocks_ adds/refreshes them (only our tagged ones; unchanged ones kept, stale ones removed; never takes over
  someone else's sheet protection) at the end of every build, at Initial Setup, after buildEventTable / an event delete
  (Events only) and buildOverrideReport. It also deletes the old version's doc properties (SHEET_LOCK_SNAP_*, SHEET_LOCK_CL_END).

Contact Log search bar (rows below END; its cells are formulas typed in the sheet, not in the repo):
- The "Communications" box is the owner's FORMULA (SEARCH_BAR_COMMS_FORMULA in Edits.gs, HSTACK in the label cell,
  reads Combined Contact Tracking D). Owner's decision 2026-11-05 after trying a script-drawn box the same day: a
  script can never be instant on a lookup and it did not follow row deletes. All script-drawn code (refresh / read
  box / "cell decides" / mode property SEARCH_BAR_COMMS_MODE / useCode... / checkSearchBarCommunications) was removed.
  Kept: findSearchBarComms_ (label in any column within 15 rows below END, also an erroring HSTACK formula),
  combinedNotesReader_ / logsNotesReader_ (debug sidebar), restoreSearchBarCommunicationsFormula (puts the formula in
  the label cell and clears the cell right of it, where it spills; also the debug sidebar's "Put formula back").
  The formula finds the OSIS box by counting rows up (ROW()-1 in the old layout); searchBarOsisRowsAbove_ (rows up to
  the "OSIS" label) + searchBarCommsFormula_ write it with the right count, and the debug sidebar reads the OSIS box there.
  Everything below is about getting Combined written sooner, which the formula then shows.
- Speed (2026-11-05; on the owner's sheet each sheet call cost 0.4-1.5 s with a 1,243-row Contact Log, so fewer
  calls is the main lever). Single-cell Contact Log edit, in order:
  1. contactLogPrefetch_: ONE cache.getAll (remembered END row CL_END_ROW_<sheetId> + ghost typist data). When END is
     <= 60 rows below the edited row, ONE read (edited row .. END, B:N) checks END and gives the auto-fill its row
     values (preRowVals); else a one-cell END check (column E scanned only when END moved).
  2. Ghost typist from the cache (Student List A5:C + Directory C:E, ghostTypistData_, key GHOST_TYPIST_DATA, 6 h,
     skipped > 95 KB; cleared by every build and the owner's reset menus, ownerClearEditCaches_ in z_MINE.gs). It
     writes only the cells it fills (C / E:F / H:I). Then autofillContactLogDefaults_ (B / G / N, fill-if-blank) and
     ONE flush paints both. The typist's OSIS is reused (typistOsis), not re-read.
  3. Row clean-up (borders, checkbox, blank row, END bar), writes only: END's row is known (checked at the start,
     +1 when a blank row was added), and the week formula check (patchMissingFormulasInRow_, reads) runs after the
     notes. Owner's order (2026-11-05): clean-up before the notes write, so the row looks finished first.
  4. Quick notes write (syncNotesForRowOsis_(osis, true) -> updateCombinedNotesQuick_, returns { texts, missing }):
     Contact Log read + Event Log / Notes tab from notesSourcesData_ (doc cache NOTES_SRC_EVENTS / NOTES_SRC_NOTES,
     6 h, Notes not cached if a cell is a Date; cleared by every EVENT_LOG helper in Core.gs, any Events / Notes edit,
     every build, the owner's resets) + Combined A:A, one cell written, then a flush (the formula follows). Runs
     under LockService.getDocumentLock (20 s tryLock; writes flushed before release) so two edits can't write older
     notes over newer ones. In a try/catch: an error is logged and the normal end-of-edit sync retries.
     A cleared / typed-over OSIS cell also updates the OSIS it had (e.oldValue, syncNotesForRowOsis_ extraOsis).
     Multi-cell edits (pastes, cleared rows: the old OSIS is unknown) run resyncNotesAfterRowDelete_(ss, sheet,
     endRow): compare every Combined student with the logs, quick-write the changed ones (> 25: full rebuild).
  5. Owner's sheet only: ownerUpdateOsisDeltas_ (PENDING_OWNER_NOTES_) with opts.combinedAlreadyWritten = true
     unless a student had no Combined row: Phone Contacts only (no in-memory Combined rebuild, buildTrackingNotesMap_
     reads the Combined tab; never rewrites Combined, which could put back older notes). The early write never falls
     back to updateContactAndPhoneOnly on the owner's sheet.
  Pastes / rows whose OSIS appears later use the old end-of-edit sync. Notes typist: writes only the cells it fills
  and flushes before the checkbox / sync. Events typist unchanged. TextFinder lookups (Notes / Events) kept on
  purpose: the in-memory cache would change matches with stray spaces.

Workflow:
- Push every change to Test-Version (not main) unless the owner says otherwise for that change. main only gets
  changes when the owner asks for it (e.g. merging Test-Version once it has been tested on a copy of the sheet).

Notes lines (Combined Contact Tracking, and Phone Contacts notes built from it):
- Each line is "𝘿𝙖𝙩𝙚 - {𝗠𝗲𝘁𝗵𝗼𝗱 - 𝗣𝗲𝗿𝘀𝗼𝗻} note [follow-up]" (event: "... - {𝗘𝘃𝗲𝗻𝘁 - [𝗣𝗲𝗿𝘀𝗼𝗻]} |Event|"), no bullet.
  Date = Sans-Serif Bold Italic (digits = sans bold), braces = Sans-Serif Bold, manual block label "𝙉𝙤𝙩𝙚𝙨:" (bold
  italic sans), all via boldText_ / noteLineHeader_ (Core.gs). Owner's pick 2026-11-05 (was the serif Mathematical
  Bold / Bold Italic set): sans-serif matches the sheet's Montserrat. Unicode bold is used because the search bar's
  Communications formula can't show real (rich-text) bold for part of a cell. Accented letters stay plain. Changing
  the style changes every notes text: run a build once after, or the delete / clear check sees all students as
  changed (> 25 -> full rebuild), and the next contacts sync updates every contact's notes once.
- Built in ONE place: contactLogNoteLine_ / eventNoteLine_ / joinNoteLines_ (Build.gs), used by both the full build
  (buildCombinedContactTrackingData_) and the quick update (updateOsisDeltas_).
- Only students on the Student List get a Combined Contact Tracking row (2026-11-05, owner's request): an OSIS from
  the Contact Log / Events / Notes that isn't on it (typo, another school) is left out. knownOsisSet_ (Core.gs)
  compares normalizeOsis_ keys, so a lost leading zero still matches. The owner's sheet also keeps Old Master Table
  students (hook oldMasterTableOsis, even before a build copies them into the Student List). The quick write
  (updateCombinedNotesQuickLocked_) checks only an OSIS that has no Combined row yet (knownStudentOsis_: the ghost
  typist cache's Student List copy, so usually no sheet read): unknown = no row, not "missing", no rebuild loop.
- Divider lines (owner's trial, 2026-11-05): in Combined Contact Tracking each dated line has a "───" line above and
  under it (one shared line between two; manual Notes block not wrapped). Length = fixed 36 "─" (noteDivider_,
  NOTE_DIVIDER_LENGTH_; owner's pick 2026-11-05, was column D width / 7px). buildTrackingNotesMap_ strips them
  (stripNoteDividers_), so Phone Contacts / Google Contacts notes are exactly as before.
- Order (owner's change, 2026-11-04): manual "𝙉𝙤𝙩𝙚𝙨:" block first, then dated lines NEWEST first (joinNoteLines_ sorts
  oldest-first then reverses, so same-day ties flip too). Was oldest first.

CODE CLEANUP (the old "POSSIBLE FUTURE PLAN"): stages 1-3 done on Test-Version (2026-11-04), stage 4 not started.
Goal was same exact functionality, cleaner files. 22 old .gs files -> 7 shared + z_MINE.gs. Verified with an offline
old-vs-new run (mock SpreadsheetApp, the owner's sample XLSX + Regenerate's fake data, shared + owner): 62 scenarios
(builds, nightly, sync incl. resume / mass-delete guard, quick updates, Contact Log / Notes / Events / Raw Data edits
and pastes, settings, overrides, migration, version check, Contact Log tool) give identical sheets, properties,
triggers, contacts, toasts and alerts; the only additions are the owner's Debug menu / toasts.
  1. Done: word-for-word moves, dead code deleted (isTruthy_, compareText_, findEndRowInColumnG_,
     getDirectoryLabelResourceName_, getGuideLink, runInitialBulkContactSync, showSidebar), always-true typeof
     checks dropped (kept: buildContactLogMenu_ for the optional module, and inside Regenerate.gs / Z#ContactLog.gs).
  2. Done: owner hooks (above); the owner copies of buildAllDerivedSheets and syncToGoogleContacts are gone.
  3. Done: one notes-line formatter, one last-name sort (Student List + Combined), one Directory sort, one Contact Log
     hide filter + week formula, safeStr_ -> toText_, one Notes/Events single-cell Student List lookup
     (findInMasterTable_) and one Events OSIS lookup (findMasterRowByOsis_).
     NOT merged because the versions behave differently (merging would change results): Contact Log single-cell
     typist (in-memory, trims) vs paste lookup (buildGhostLookup_, splits groups) vs TextFinder ones; the 3 Directory
     site/class lookups (sibling order differs; one is in Z#ContactLog.gs); the 4 settings lists (different keys/defaults).
     Not done: constants for the 123 sheet-name literals ("Raw Data" vs "RAW Data" lookups differ by place).
  Speed-ups (same results): Combined Contact Tracking groups rows by OSIS once (was every row x every student); sorts
  reuse one collator (compareNatural_); a sibling / multi-row edit reads the logs once (updateOsisDeltas_), and on the
  owner's sheet rebuilds once per edit instead of once per OSIS; owner Phone Contacts groups guardians once; Notes /
  Events typist reads the 3 identity cells in one call; sync reads header positions once.
  4. Not started — optional behavior fixes, one at a time, only with the owner's go-ahead.
Row deletes (2026-11-05): installable change trigger installedOnChange (Edits.gs) -> on REMOVE_ROW with the Contact Log
active: resyncNotesAfterRowDelete_ redraws the END bar border, recomputes every Combined student's notes from the logs
(notesTextForOsis_, rows grouped by OSIS once, Contact Log to END + notesSourcesData_), quick-writes those that differ
(updateCombinedNotesQuick_), then on the owner's sheet the Phone-Contacts-only owner update (> 25 differ:
updateContactAndPhoneOnly, owner-aware through buildPhoneContactsData_). Also run for Contact Log multi-cell edits
(pastes / cleared rows). Made by
setupAllInstallableTriggers and, once per sheet, by checkVersionOnOpen (ensureChangeTrigger_, doc property
CHANGE_TRIGGER_ADDED, so a 2nd user's open doesn't add a 2nd copy).
Must keep: 15 menu handler names (+5 owner Debug items, + toggleSheetLock), 39 google.script.run endpoints, trigger
handlers (installedOnEdit, installedOnChange, nightlyMasterSync, checkVersionOnOpen, syncToGoogleContacts),
all sheet names (except the owner's Master Table -> Student List rename), Backend cells,
property/cache keys.
Owner decisions still open (current behavior kept): add shared sync safety checks (50% delete stop, label pageSize)
to owner sync?; owner hides Contact Log col K (11) vs shared H (8) when no site (K looks stale); Montserrat on New
Students (owner build skips it); owner splitFullName_ changes the shared Directory name split - intended?; ghost typist
TextFinder vs in-memory scan (differ only on stray spaces); carry Your Lists over at Migration / Rollover (now they are not).
Data update reminder (Menus.gs, replaced the old 30-day ATS-only reminder; same checkRawDataFreshness_ name): centered
pop-up when Raw Data A1 / AG1 / AP1 (ATS / Parent Square / NYCSA, only sections ever loaded) is older than setting
"dataReminder" (week default / 2weeks / month / off; in SETTING_KEYS so it is saved to the Backend settings block and
shown in ⚙️ Settings). Checked by checkVersionOnOpen and by installedOnEdit (checkDataReminderWhileOpen_, user cache:
once an hour); once a day (ATS_REMINDER_LAST_SHOWN); not during Input Data Mode. "Go to Input Data Mode" only if
openInputDataMode exists (guarded), else the guide link. Skip -> Remind After 2 Weeks / 1 Month (setting + wait in doc
property DATA_REMINDER_SNOOZE_UNTIL) / Remove Reminders / Just skip for today. Endpoints dataReminderOpenInputMode,
dataReminderChoose. Saving a different choice in Settings clears the wait.
Settings saved but never read: autoOverride, ignoreHideCheckboxes, emailNotifications.
Nightly trigger (nightlyMasterSync): 10 PM, NIGHTLY_SYNC_HOUR (Menus.gs), both at Initial Setup and when a Settings
save re-makes it (owner's pick 2026-11-05; was 11 PM at setup but 2 AM after a Settings save). Existing triggers keep
their old hour until Settings is saved once or Initial Setup is run again.
Bugs noticed (not fixed): quick update ignores the Settings notes date range; getExistingMasterTable_ has no header but buildPhoneContactsData_ slice(1)s it
(first student loses Student List details); EMERGENCY_STOP_SYNC leaves SYNC_EDITED_COUNT; TOGGLE_EMAIL_NOTIFICATIONS
needs two clicks first time; repo manifest lacks script.send_mail so the owner email may never send;
purgeDischargedStudents and the Hide Discharged filter take ROW 1 as the header, but the Student List header is on
row 4, so Purge finds no Status column (returned 0 in the offline run) and the filter hides nothing, unless the real
tab has a header copy on row 1 (check on a copy); generateFakeLegacyData (owner) writes 47 columns into a new
26-column "Old Phone Contacts" tab, so it fails right after a reset.
