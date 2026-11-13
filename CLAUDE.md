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
  tab helpers, owner hook point (getOwnerHooks_, runAction_), small text / phone / OSIS / date helpers, SHEET_LOOK_
  (the tabs' colors and font, see SHEET LOOK).
- Build.gs: buildAllDerivedSheets + everything it builds (Student List, Combined Contact Tracking, Parents Divided,
  Directory, Phone Contacts), new / discharged student detection, the notes-line helpers, nightlyMasterSync, build
  sidebar endpoints. BUILD_STEPS + reportBuildStep_ drive the sidebar's "Building Sheets & Directory" checklist
  (cache key BUILD_STEP; getSyncStatus returns buildSteps / buildStep with the state).
- Edits.gs: installedOnEdit, ghost typists (Contact Log / Notes / Events), Raw Data paste clean-up, quick updates
  (updateSingleOsisDelta_ / updateOsisDeltas_ / updateContactAndPhoneOnly), Events tables. (The temporary
  🐞 Contact Log Debug sidebar is the owner's test file z_ContactLogDebug.gs, see below.)
- Menus.gs: onOpen menus, setup, 👋 Welcome tab (new copies, see below), Getting Started (welcome + checklist, see
  below), Settings, Student Overrides + report, Rollover, Migration, Version Check, data update reminder, links.
- Input Data.gs: Input Data Mode (part 1) + Your Lists (part 2), see below.
- Contacts Sync.gs: Google Contacts sync (syncToGoogleContacts, searchAndUpsertSingleContact), Full Sync endpoints,
  EMERGENCY_STOP_SYNC / RESTART_SYNC.
- Sheet Lock.gs: "baby-proofing" (see SHEET LOCK below): Google warning-only protections + toggleSheetLock
  (App Menu "🔓 Unlock Sheet" / "🔒 Lock Sheet").
- Regenerate.gs: reset / fake-data tools (resetToDisplayMode, resetAndLoadFakeData) and Contact Log formatting.
  Owner asked to keep it as is (not moved, not cleaned); change it only when the owner asks (e.g. KEEP_TABS_).
  Student List rename (owner's request): KEEP_TABS_ says "Student List", deleteLegacySheets_ calls
  getStudentListSheet_ first (an old "Master Table" is renamed, not deleted) and resetMasterTable_ uses it.
  Sheet look (owner's request 2026-11-12, see SHEET LOOK): forceVibrantColors_ (Raw Data: Montserrat on the whole
  tab, DATE UPDATED titles left-aligned, Parent Square header #f6dfc4 instead of #ff9900), the fake data's 2 Raw Data
  fonts, resetMasterTable_ deletes the Student List's columns after H (STUDENT_LIST_COLUMNS_; clears
  min(9, columns)), and resetContactLogFormatting's Workshop rows (J = "Workshop", C:O) are pale green
  SHEET_LOOK_.workshopRow #e6f1ea (was solid #b6d7a8).
- Styles.html: the one shared look ("District": navy #183e69, Source Sans 3 text, Source Serif 4 headings,
  checkbox-style switches, 3px corners) for every dialog and sidebar. HTML files pull it in with
  <?!= include_('Styles') ?> and are opened with htmlFromFile_(name) (Core.gs, = createTemplateFromFile().evaluate());
  dialogs built as strings in .gs (Welcome, Report a Bug, ATS reminder, Create Event) start with include_('Styles').
  No other <? ... ?> tags may appear in the HTML files. Z#ContactLogSidebar keeps its own style (left as is).
- Z#ContactLog.gs (+ Z#ContactLogSidebar.html): self-contained optional "📇 Contact Log" tool (➕ New Interaction /
  🔔 Manage Notes / Follow-ups); left as is. It is the ONLY Z# file left: the owner keeps it as Z# for later
  (2026-11-11, when every other Z# file was coalesced into the shared files).
  "Z#" = a temporary file for trying out a new feature: no other code may rely on Z# files (the only link is the
  guarded `if (typeof buildContactLogMenu_ === "function")` line in onOpen, a no-op without the file).

COALESCED Z# FILES (owner's request 2026-11-11): Input Data Mode, Your Lists, Getting Started and Contact Log Debug are
no longer optional trial files. Their code moved word for word (function names, google.script.run endpoints, document
property / cache / Backend keys unchanged); the typeof guards in shared code and the "file missing" fallbacks were
removed, so these files are now REQUIRED (a missing one = onOpen error / no menu). Old -> new:
  Z#InputDataMode.gs + Z#OwnData.gs -> Input Data.gs (part 1 / part 2); Z#Guide.gs -> Menus.gs section GETTING
  STARTED; Z#ContactLogDebug.gs -> end of Edits.gs (later the owner test file z_ContactLogDebug.gs); HTML: Z#InputDataSidebar -> InputDataSidebar, Z#OwnDataSidebar ->
  YourListsSidebar, Z#GuideSidebar -> GettingStartedSidebar, Z#ContactLogDebugSidebar -> ContactLogDebugSidebar.
  Only behavior change: none on a sheet that had all the Z# files (the guards were always true there).

- Input Data Mode (Input Data.gs part 1 + InputDataSidebar.html; was Z#InputDataMode.gs): "📥 Input Data Mode"
  (🚀 App Menu, both the getting-started and full menus, via the 2 `addInputDataModeMenuItem_` lines in buildMenus_;
  also opened by the data update reminder and Getting Started). Hides all visible tabs (remembered),
  then one temp tab at a time ("Student Data", "Parent Square Data", "NYCSA Data"); paste anywhere or upload the
  CSV/Excel file (SheetJS from cdnjs, read in the sidebar). Which file to pick (owner 2026-11-13): IDM_STEPS' last
  howTo line + fileHint (shown on the upload card): Student Data starts with general_, Parent Square = parents.csv,
  NYCSA starts with Students_With_NonNycsa_Accts_Rpt_ (also in docs/USER_GUIDE.md "Where to get each export"). Sidebar polls every 2s, finds the header row, refuses the
  wrong export and numbers Excel cut off (E+ / #### / ID-phone columns ending in 000; "Use it anyway" allowed),
  then replaces that Raw Data section (rows 7+, same headers/stamps as cleanRawDataPasteZone_ /
  updateRawDataTimestamp_). Skip keeps old data. Step 4 "Site / Class" (manual, Save button): Class | Teacher Name |
  Site (Optional), one row per ATS class code (Raw Data R) + classes already in BH:BI, prefilled from BH:BI; Site
  dropdown = hidden col Z formula of every typed site (other values allowed). Save rewrites BH:BI rows 7+ as ONE ROW
  PER SITE (owner's format): BH = site, BI = that site's classes, one per line, "Y25: Carl" or bare "Y24" (no
  teacher); sites in number order, no-site classes last. Read by parseClassCell_ (Build.gs, used by
  buildClassReferenceMap_ and Input Data Mode): one class per line, a line may still hold the old "Y41: Tessa Y25: Sam" (read
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
  (idm_deleteLeftoverTab, id in INPUT_DATA_MODE_DELETE_TAB; also retried on open). Debug panel: only with
  z_InputDataDebug.gs (see below; IDM_DEBUG is gone). The temp tab is fully cleared between steps (idm_clearTab_: every
  dropdown, hidden column Z, frozen row). Sidebar on Styles.html since 2026-11-12 (htmlFromFile_, title without emoji).
  Step 5 "Your Lists" (step key lists, lists: true; part 2 of the same file): no paste / upload /
  polling / Skip; the lists card shows od_listSummaries_ (idm_view_ "lists")
  with Paste new version / Edit setup / ✏️ rename / 🗑️ delete, "➕ Add another list" (od_createList opens
  YourListsSidebar IN PLACE of this sidebar; od_backToInputMode reopens it) and Finish (idm_finishLists:
  results.lists = { saved, rows: ready lists, lists: true } or skipped, so the end screen offers Build). Site / Class
  is no longer the last step: its button is "✅ Save Site / Class".
  Wrong paste (owner's request 2026-11-12): any data problem on a PASTE (no header row, wrong export, no rows, numbers
  cut off; not "Raw Data tab missing") clears the temp tab right away (idm_processRows_ fail -> idm_resetTempSheet_,
  notice.cleared); the sidebar keeps the message up while it goes on polling for the next paste (stickyError) and
  its button reads OK. Cut-off numbers: the read cells come back in notice.heldRows (Dates as their shown text) and
  "Use it anyway" sends them to idm_submitRows(rows, true), the upload path. Uploads are never "cleared" (nothing on
  the tab).
  Build before step 5 (owner's request 2026-11-12): Your Lists matches rows to the Student List, so idm_view_ says
  needsBuild on step 5 when data was moved to Raw Data in this run (state.dataChanged, set by every section write and
  the Site / Class save) or the Student List was never built (studentListIsBuilt_) and no build was tried yet. The
  sidebar then calls idm_buildForLists by itself (spinner shows the build's step from getSyncStatus): no Student Data
  in Raw Data (C7 empty) -> no build, a note says to load it first; else buildAllDerivedSheets + formatDirectoryAuto_
  WITHOUT the document lock (the build lock only), then every tab the build made / showed is hidden again (added to
  state.hiddenIds, shown at the end) and the temp tab is active again. Result in state.listsBuild ({ ok, message },
  shown on the lists card); dataChanged = false. A failed build is reported, not retried, until new data comes in.
  End screen (2026-11-12): "Build sheets now" only when idm_finalBuildNeeded_ (something came in and NOT
  idm_builtAlready_ = step 5's build ok and no dataChanged / listsChanged since; od_saveList / od_deleteList set
  state.listsChanged via idm_noteListsChanged_); else "Your sheets are already built with the new data." "Next steps"
  (openGuideChecklist) while the Getting Started checklist isn't all done (idm_guideNext_; primary button when no
  build is needed). Closing the mode (idm_cleanup_) shows the tabs it hid EXCEPT the Your Lists tabs (every list in
  OWN_LISTS), and hides those too (owner's request 2026-11-12; od_openList / od_repasteList show one again).
  idm_deleteLeftoverTab (the separate call after the end screen) then runs refreshMenusAfterFirstBuild_.
  "Finish or exit" (was "Exit Input Data Mode"; owner's request 2026-11-12, build at any point): the panel asks
  "Leave Input Data Mode?" with "Save, exit and build my sheets" (primary) / "Save and exit (build later)" / "Exit
  without saving" / Cancel. Build = idm_saveAndExit (a save problem keeps the mode on and shows it), then the sidebar
  calls idm_build in a SEPARATE call (deletes the temp tab, opens the Build sidebar), unless builtAlready ("already up
  to date") or no Student Data. idm_view_.canBuild (Student Data in Raw Data, idm_canBuild_, or on the Student Data
  step) greys the button with a note; the exit result's canBuild is checked again after the save.

- Your Lists (Input Data.gs part 2 + YourListsSidebar.html; was Z#OwnData.gs; owner's request 2026-11-05): your own lists (bus
  routes, PTA, trip forms) matched to students or guardians; reached only from Input Data Mode step 5 (no menu item).
  One tab per list "📥 <name>" (rows 1-3 = OWN DATA band / status / column chips, header row 4 after a paste, data from
  row 5, tool columns "✔ Matched" / "How" after one blank spacer column). Its setup = one row in the Backend OWN_LISTS
  section (Tab ID, List, Setup JSON, Updated; od_loadLists_ / od_saveLists_): by (student / guardian), roles (osis,
  student / guardian as {col} or {first, last}, phone, email, class), vars (label, show = Additional Information,
  column = Student List / Parents Divided column, blank / notOnList = nothing | text), picks (row hash -> OSIS /
  guardian key / "skip"), headers, title, stats.
  - Paste: od_takePaste_ finds the header row (od_findHeaderRow_; title rows above it are remembered), moves it to
    row 4, data as plain text. A re-paste with the same headers keeps the setup (od_rolesByHeader_) and the picks.
    Upload (owner's request 2026-11-12): "Option 1: Upload the file" in the paste view (SheetJS 0.18.5 from cdnjs,
    first sheet, raw values, Dates -> M/D/YYYY, like InputDataSidebar). od_uploadCsv(text) (Utilities.parseCsv via
    idm_parseCsv_) / od_uploadRows(rows): written as plain text from row 4 of the reset list tab, then od_readTab_
    (shared with od_readPaste) reads it exactly like a paste. A list with a header but no rows is cleared
    (od_resetListTab_, cleared: true): before, the header stayed and the 2 s poll re-read it in a loop.
  - Matching (od_matchStudentRow_ / od_matchGuardianRow_): your pick, then OSIS (normalizeOsis_), exact name token
    set (either name order), phone / email via guardians (Raw Data incl. overrides), then BEST GUESSES WITH A
    CONFIDENCE (owner's request 2026-11-12; od_rankStudents_ / od_rankPersons_ / od_decideRanked_): every candidate
    gets 0-100% from od_nameScore_ = 45% first name + 55% last name (od_spellAlike_ = Jaro-Winkler and edit ratio
    averaged; first names also get the rules: same 100, nickname OD_NICKNAMES_ 95, "Sam"/Samuel 90, initial 72), also
    tried in the other order ("SIMPSON BART") and as a share of name parts (middle / two last names); "high" (90+)
    needs first AND last each >= 0.9, else 89 max. The old rules (od_studentGuesses_: part of the name, short name,
    "Diana P.", first name only, a guardian's name -> their children) keep their reason and give a floor (90 / 88 /
    78 / 60 / 85). Boosts: parent's name fits +5, same phone / email +6, same class +4 (capped 99). Candidates =
    people whose LAST name starts with a letter of the row's name (byLastInitial; one-word names: any part). Top two
    within OD_CONF_TIE_ (4) points = "pick one" (with %), but only from medium (75) up: a LOW tie (siblings sharing a
    last name with someone not on the list) = "not found" + suggestions; top >= OD_CONF_MIN_ (60) = best guess (+ up
    to 2 "also close"); else "not found" with the closest 3 >= 50% as suggestions. Column guess: a list with parent
    first + last columns and a first-name-only child column is a GUARDIAN list (od_guessRoles_, stats.multi). Levels: high 90+ / medium 75+ / low 60+.
    Never used without a person: guesses only become matches as Your picks (the build only uses "ok" rows). Sidebar:
    % chips, "Accept the N high-confidence" + "Accept all N" buttons, the tab's ✔ Matched column shows the %.
    Tested offline on sample names (typos, swapped order, nicknames, initials, siblings); 500 unmatched rows x 1,500
    students = ~0.9 s in Node. Index = Student List OSIS..Class (sidebar: the Backend copy, sibling-group rows
    skipped) + Raw Data guardians (od_indexFromSheet_).
  - VIEW, step 4 of the setup sidebar (owner's request 2026-11-12; Save moved there): how the kept values read.
    setup.view (Setup JSON): layout "lines" ("Name: value" per line, as before) | "oneLine" ("List: v1, v2");
    boldNames / boldValues / italic (character bold via boldText_, like the notes); sep comma / dot / bar / slash;
    namesBefore; start list / none / custom + startText; columns "each" (a column per value, as before) | "one" (one
    column named after the list, plain text); order = column HEADERS (survives moved columns). A list saved before
    has no view: od_fillSetup_ gives od_defaultView_(false) = exactly the old text (checked old-vs-new on the
    test-data lists: Additional Information + both column sets identical); a NEW list starts with bold names
    (od_emptySetup_). od_viewText_ (Input Data.gs) is the ONE place the text is made (od_buildOutputs_ via
    od_viewVars_ + od_varLabel_); the sidebar preview has a copy between "// VIEW TEXT" and "// END VIEW TEXT"
    (YourListsSidebar.html) that test_input_flow.js compares with the server for 192 views: change both together.
    SIMPLE + ADVANCED (owner 2026-11-12): the step shows only One per line / One line, "Bold the names" and the
    preview (Matched rows from od_matchPreview's preview, ‹ › / Blank cell / Not on list); "Advanced options"
    (<details>, open by itself when one of them is not the default) has start, separator, names before values,
    bold values, bold italic, order ↑↓, and the Student List / Parents Divided column choice.
    Not built yet (second round): a "Sentence" layout ("Takes bus {Route} from {Stop}").
  - Sidebar: paste view polls od_wizardCheck every 2 s; 1 Columns (dropdowns 'Column C “OSIS”', "Two columns (first +
    last name)…", OSIS hint "5 of 6 filled cells are OSIS in your Student List"), 2 Variables (keep / name / blank /
    not on list / where, live "How it reads" preview, privacy warning), 3 Match (od_matchPreview; "✓ Accept all N best
    guesses" with a warning, pick / search picker / skip / undo). Save = od_saveList (writes the tool columns + stats).
    Done view: during Input Data Mode only "← Back to Input Data Mode" (a build then would run with the tabs hidden);
    outside it "🏗️ Build Sheets Now" (od_build).
  - Build: buildAllDerivedSheets calls od_buildOutputs_(ss, masterTable, guardianRows) right after
    addLastInteractionColumn_ (try/catch: an error only logs "Your Lists skipped"). It re-matches every
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
    load fake data on the copy" builds the copy with this script's cache. od_buildOutputs_ clears ADDL_INFO when there
    are no lists (no-op if never written).
  - Resets: list tabs are deleted by deleteLegacySheets_ (not in KEEP_TABS_); resetDerivedSheets_ clears ADDL_INFO +
    OWN_LISTS. Migration carries them over (owner 2026-11-12): ma_migrateSimpleSheets (step 4 "Copy events, student
    changes and your lists") calls od_migrateLists_: each old list tab is copied (copyTo, same name) with its setup and
    picks under the new tab ID; a same-name list already here is replaced; the old sheet is only read; step 6's build
    fills the columns. Rollover puts each list tab in the archive PDF (next year's sheet gets them through Migration).
  - google.script.run endpoints: od_createList, od_openList, od_repasteList, od_renameList, od_deleteList, od_getWizard,
    od_wizardCheck, od_readPaste, od_uploadCsv, od_uploadRows, od_matchPreview, od_saveList, od_backToInputMode,
    od_build (+ idm_finishLists, idm_buildForLists).
    Tested offline (own_test: shared + owner) and by a Chromium click-through of both sidebars wired to the real
    server code on the mock sheet (not yet tried on a real sheet).

- Getting Started (Menus.gs section GETTING STARTED + GettingStartedSidebar.html; was Z#Guide.gs): "✅ Getting Started"
  (bare-bones version, owner's pick 2026-11-05; a full tour / tab pop-ups / Start Here tab / header notes were proposed
  and NOT built). Called from Menus.gs: addGuideMenuItems_ (both 🚀 App Menus), showGuideWelcome_
  (runInitialSetupWrapper, instead of showWelcomeDialog on new-user copies; "📖 Open User Guide" still opens
  showWelcomeDialog), guideOnOpen_ (checkVersionOnOpen, in a try/catch).
  Welcome pop-up after setup: 4 steps, "Start with Input Data Mode", video / guide
  links. Checklist sidebar: 1 load data (ticks when Raw Data A1 has a real date; also shows Parent Square AG1, NYCSA
  AP1, Sites and classes BH7:BI7 / INPUT_DATA_MODE_SITECLASS_UPDATED), 2 build (Student List lastRow > 4, same test as
  the menus), 3 settings (doc property ZG_SETTINGS_SEEN, set when opened from the checklist), 4 first contact (any
  Contact Log C above END). (Step 2 = studentListIsBuilt_, Core.gs, since 2026-11-11.) Opens by itself on every open until all 4 are done, not during Input Data Mode, unless
  unticked (doc property ZG_AUTO_OPEN = "off"). "This tab" panel (zg_activeTab polled every 3 s) + "Hidden tabs" list
  (Settings tabs: ⚙️ Settings → Tabs & columns; others: View → Hidden sheets). Link PLACEHOLDERS: GUIDE_LINKS at the
  top of the GETTING STARTED section (empty = "link coming soon").

- 🐞 Contact Log Debug (z_ContactLogDebug.gs + z_ContactLogDebugSidebar.html since 2026-11-12, owner's testing only;
  was Z#ContactLogDebug.gs, then briefly the end of Edits.gs): TEMPORARY sidebar (🚀 App Menu, both menus, the 2
  guarded `typeof addContactLogDebugMenuItem_` lines in buildMenus_). Delete both z_ files for schools: no menu item,
  and the Core.gs recorder never turns on. Live list of runs (Contact Log / Notes / Events
  edits, row deletes) with each step's ms since the script started (⏱ ghost typist incl. data read vs write+paint,
  auto-fill, notes sync, Combined row written, search bar); "Check search bar" snapshot (settings, triggers, END,
  Communications cell + formula?, OSIS box, notes shown, Combined vs logs for that OSIS, newest rows); "Copy report";
  "Put formula back" (the "Use script" button went with the script-drawn box). "Anonymize" (on by default, cld_report(anon) / cldAnonymizer_): Student List names ->
  Student-N / Guardian-N / Name-N, OSIS -> OSIS-N, phones / emails masked, note texts -> "[N chars, K lines, #hash]"
  (cldNoteSummary_, same text = same hash), Contact Log D -> (filled)/(blank), unknown typed text -> [typed text].
  Recording = shared logger in Core.gs (clDbgStart_ / clDbg_ / clDbgEnd_ / clDbgRun_, wrapped
  around installedOnEdit / installedOnChange): on only while doc cache key CL_DEBUG_ON exists (the sidebar re-puts it
  every 2 s poll, 5-min TTL); runs in CL_DEBUG_LOG (newest 25). Off = one cache read per edit.

Owner testing files (2026-11-12; keep them on the owner's sheet, DELETE them from the master template / schools'
copies; shared code reaches them only through typeof checks): z_ContactLogDebug.gs + z_ContactLogDebugSidebar.html
(above), z_NewCopy.gs (see below) and z_InputDataDebug.gs: the Input Data Mode debug log (idm_debugEnabled_ +
idm_debugSheets_). With it every Input Data Mode sidebar reply carries its step log and the sidebar shows the Debug
panel; without it idm_log_ does nothing and the panel stays hidden. Input Data.gs reaches it only through typeof
checks. Keep it on the owner's sheet while testing; DELETE it from the master template / schools' copies.
Before Initial Setup the simple onEdit (Menus.gs) undoes every edit with "🔒 Action Blocked", EXCEPT while any z_ file
is in the project (ownerTestFilesPresent_: typeof ownerHooks_ / addContactLogDebugMenuItem_ / idm_debugEnabled_ /
ownerNewCopyCheck_; owner's request 2026-11-12): the owner's sheet and test copies can be typed in before setup; a
school's copy (no z_ files) is still blocked. tests/lib/load.js: the default load = no z_ files at all (owner: true adds
z_MINE + z_NewCopy, debugFiles: true the two debug files).

z_NewCopy.gs (owner only, 2026-11-12): the copy part of DEPLOY.md as a pop-up on the first open of a NEW COPY of the
owner's sheet, and as 🐞 Debug → 📋 New Copy Checklist (HTML dialog, ticks in doc property OWNER_NEW_COPY_TICKS;
endpoints ownerNewCopySetTick / ownerNewCopyClearTicks). Called only from z_MINE.gs ownerAddMenus_ (2 typeof-guarded
lines, after addToUi). Doc properties don't travel with a copy, so ownerNewCopyCheck_ asks once "Is this your MAIN
sheet?" (Yes -> OWNER_HOME_SHEET_ID; No -> checklist, OWNER_COPY_CHECKLIST_SHOWN = that copy's ID); the constant
OWNER_HOME_SHEET_ID_ skips the question on the main sheet. Plain ui.alert because a new copy only has the simple
onOpen (no HtmlService there). Keep NEW_COPY_STEPS_ in step with DEPLOY.md.

Owner-only file (in the repo, deployed ONLY to the owner's own sheet): z_MINE.gs
- Never copy it into other schools' sheets. It defines ownerHooks_(); shared code calls getOwnerHooks_() (Core.gs),
  which returns those hooks on the owner's sheet and null everywhere else. No shared function is redefined, so file
  load order does not matter. Shared code must work without z_MINE.gs; owner-only behavior (e.g. "Old Master Table" /
  "Old Phone Contacts") goes into a hook in z_MINE.gs, never into the shared files.
- Hooks: Student List status snapshot, Old Master Table history (Student List + Directory), Old Master Table OSIS
  that Combined Contact Tracking also keeps (oldMasterTableOsis), Old Phone Contacts merge,
  last-word name split (splitFullName_), tall Phone Contacts layout, quick update that also rewrites Phone Contacts,
  contacts sync (4 phones / 3 emails per row, match = name AND phone / email, summary email incl. "Not saved" and
  the note; since 2026-11-12 the same label lookup (pageSize 1000) and "more than half" check as every sheet),
  Contact Log site column (H like the shared sheets since 2026-11-12: K is Method 2), Montserrat list, the New Students tab (ownerLogNewStudents_), 🐞 Debug menu +
  toasts.
- 🐞 Debug → "🔧 Fix Dropdown Criteria (Student List → Backend)" (ownerMenuFixDropdownCriteria, one-time tool,
  2026-11-11): every tab's data validations (one getDataValidations per tab) listing Student List columns A:G are
  repointed to the same Backend column from row 3 (rule.copy(), other settings kept), per run of cells in a column.
  Needs a build first (refuses while BX is empty). Sheets Table "Dropdown" columns throw -> listed "change by hand";
  custom-formula rules mentioning the Student List are listed, not changed. Running it again changes nothing.
- Shared Directory rule: draws only from Raw Data (no overrides, no discharged students). On the owner's sheet the
  hook then merges Old Phone Contacts and adds non-Raw-Data students from Old Master Table ONLY when its column F
  (Status) is "MASTER".
- 🐞 Debug menu (owner only): "Reset + Load Fake Data" and "Reset to Display Mode" (type RESET, then Regenerate's own
  confirm), "📦 Make Release Copies" (below; "Copy Sheet and Set to Migrate" was removed 2026-11-12 with the MIGRATE
  template), "Copy Sheet, Then Reset the Copy" (makes a Drive copy
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
- Contact Log Follow-up (N): the 6 options, EXACT text (owner's screenshot 2026-11-12): "No Followup Needed",
  "Followup Needed", "Followup Needed - ASAP", "Followed up / Resolved", "Skipped / Newer Interaction", "Waiting for
  Response". Code writes only these (auto-fill default, Migration map, Z#ContactLog, fake data); tests/test_fixes.js
  fails if any other value is written. The grey-out rule compares case-insensitively (Sheets <>), so its "Waiting For
  Response" matches; its "Need to respond" / "Waiting - Check Up" are old options (harmless). The Contact Log is a Sheets Table and N a "Dropdown" typed column: Apps
  Script may not change its options ("not allowed on cells in typed columns"), so code never touches N's dropdown
  (set by hand in the Table; owner removed the reset step that tried). Autofill default for N = "No Followup Needed" (Edits.gs). Hide rule for P
  (getContactLogHideFormula_, owner's 2026-11-04 formula): =AND($B2<>"", INT($B2)<>TODAY(), TRIM($N2)<> Followup Needed / Followup Needed - ASAP / Need to respond / Waiting For Response /
  Waiting - Check Up). The date part follows Settings → Notes → "Contact Log shows" (contactLogTodayOnly, 2026-11-11 fix:
  it had stopped changing the rule): "Today only" = INT($B2)<>TODAY() (the owner's formula exactly); "This week and last
  week" (default, never saved) = B outside last week's Sunday .. this week's Saturday (same weeks as the A:B color bands).
  Used by every reset, the nightly run (resetContactLogFormatting) and every Settings save: refreshContactLogFilter_
  (Menus.gs) swaps ONLY the P custom-formula rule (rule.copy(), Type colors and other rules kept; added back if missing),
  writes nothing when it already matches, then re-sets the P filter. The old "Yes - Important" bold rule was removed from resetContactLogFormatting (owner 2026-11-12: that value is
  now "Followup Needed - ASAP"; the next reset / nightly drops it from the sheet). Z#ContactLog writes / reads the new
  values too (Followup Needed / No Followup Needed / Resolved; open = the grey-out rule's list, cl_isOpenFollowup_).
  Migration maps old values: Yes -> Followup Needed, Yes - Important -> Followup Needed - ASAP, No -> No Followup
  Needed (migrateFollowupValue_, Menus.gs).
- Fake data (Reset + Load Fake Data, Regenerate.gs; owner's request 2026-11-04): 18 ATS students (incl. siblings
  Belcher Louise/Gene + override-only Tina, Addams Wednesday/Pugsley), Parent Square, NYCSA, Site/Class in the new
  BH:BI format (one row per site, classes one per line; 5 of 6 sites have 2-3 classes; "Y12"/"Y50" have no teacher),
  26 Contact Log rows (30-row log; each row carries its Type / Follow-up / Follow-up note; Site / Class H:I filled
  from the Directory right after the build, fillContactLogSiteClass_, owner 2026-11-13: the typist that fills them
  never runs on script writes, so the Example had them empty), 3 events (Fall PTC,
  Graduation Workshop, Spring Festival; Backend Event Log rows now in the real column order), 5 Notes rows, 3
  New/Edit Student overrides (New Student Tina Belcher, Peter Parker + guardian, Diana Prince hidden from report),
  then the Students Override Report. PCAR gets no fake data (built from other tabs); Send Out: see below.
  With Debug on, runAction_ toasts ✅ / ⚠️ / ❌ after each edit, open, Build Only, Full Sync, nightly and contacts sync,
  including errors that are normally only logged (console.error / console.warn).

Deploying (manual paste into the Apps Script editor; master template + owner sheet):
- MANIFESTS (2026-11-12; was one "App Scripts.Json"): the editor's appsscript.json (Project Settings → tick "Show
  'appsscript.json' manifest file in editor") gets "App Scripts (General).Json" on the master template / every
  school, and "App Scripts (Mine).Json" on the owner's sheet only. Mine = General + script.send_mail +
  userinfo.email (the owner's summary email). Keep the two in step: a scope added for shared code goes in BOTH.
  Changing a project's scopes makes its user re-approve once (run any menu item): until then that user's
  triggers (nightly, edit) stop with "authorization required". So the shared one only changes when shared code
  needs a new permission.
- Delete the old files first or the project will not load (duplicate consts/functions): Backend, Build.gs Helpers,
  Build.gs beginning, Constants, Event, LinkResolver, MigrationAssistant, New Copy & On Open, NewStudents, Phone,
  Quick Updates Installed, QuickUpdates, RolloverBackend, Settings Backend, Student Override Report,
  Student Overrides Backend, VersionCheck, ZProduction, z_PhoneOverride, z_PhoneOverride 4+, z_PhoneOverride.gs Add On.
- Then paste Core, Build (replaces the old Build), Edits, Menus, Contacts Sync, Sheet Lock (new); owner sheet also z_MINE.
  Student List rename (2026-11-05): also Regenerate, SettingsDialog, ProgressSidebar and the (now coalesced) Getting
  Started / Contact Log Debug / Input Data Mode files (they name the tab).
  UI refresh: also add Styles.html (new) and re-paste SettingsDialog, ProgressSidebar, NewStudentDialog,
  ManageOverridesDialog, RolloverWizard, MigrationUI, VersionDialog. Regenerate (re-paste: KEEP_TABS_ tab clean-up, Color Tabs removed), the Z# files
  and the manifest are unchanged. The Contact Log tool is now named
  Z#ContactLog / Z#ContactLogSidebar (delete the old #ContactLog / ContactLogSidebar files, or they define it twice).
- Lookup box layout (2026-11-11): re-paste Regenerate, Edits, Z#ContactLogDebug (now in Edits), Build (the box itself is set up by
  hand in the sheet; the code finds it, cleans under it, outlines it and adds / removes guardian rows).
- Student List without sibling rows + header on row 1 (2026-11-11): re-paste Core, Build, Edits, Menus, Regenerate,
  Z#OwnData, Z#Guide, Z#ContactLogDebug (now Input Data, Menus, Edits; owner sheet: also z_MINE), then Build Sheets once (fills Backend BX:CD,
  writes the Student List from row 1, repoints the Events dropdowns); owner sheet: then 🐞 Debug → Fix Dropdown
  Criteria for dropdowns made by hand. The master template's Student List tab needs the same row-1 layout (the build
  writes from row 1 anyway, but frozen rows / header look are set by hand).
- Coalesced Z# files (2026-11-11), the newest step, on top of the above:
  1. DELETE these 8 files first (their code now lives in the files below; keeping them = duplicate functions / consts,
     the project will not load): Z#InputDataMode, Z#InputDataSidebar, Z#OwnData, Z#OwnDataSidebar, Z#Guide,
     Z#GuideSidebar, Z#ContactLogDebug, Z#ContactLogDebugSidebar.
  2. ADD the script file "Input Data" and the HTML files InputDataSidebar, YourListsSidebar, GettingStartedSidebar
     (HTML file names must match exactly, without ".html"). Owner's sheet only: z_ContactLogDebug (script) +
     z_ContactLogDebugSidebar (HTML), z_InputDataDebug and z_NewCopy (scripts); never on the master template / schools.
  3. RE-PASTE Core, Build, Edits, Menus (owner sheet: also z_MINE); Regenerate too if the two 2026-11-11 steps above
     are not deployed yet. Contacts Sync, Sheet Lock, the other HTML files, Z#ContactLog / Z#ContactLogSidebar and the
     manifest are unchanged by the coalesce.
  All of these are now required (no longer optional; the z_ test files excepted). Needs Styles.html.
- Sync rewrite + owner picks (2026-11-12, later): re-paste Core, Edits, Menus, Contacts Sync, Regenerate (fake data
  "Resolved") and the HTML ProgressSidebar, RolloverWizard; owner sheet also z_MINE. The Contact Log Debug files are
  now z_ContactLogDebug (script) + z_ContactLogDebugSidebar (HTML): on the owner's sheet delete ContactLogDebugSidebar
  and add both z_ files; on the master template / schools delete ContactLogDebugSidebar and add neither.
- Fixes + Your Lists confidence + restyle (2026-11-12): re-paste Core, Build, Edits, Menus, Contacts Sync, Sheet Lock,
  Regenerate, Input Data, Z#ContactLog (scripts) and InputDataSidebar, YourListsSidebar, MigrationUI (HTML); owner sheet
  also z_MINE and ADD z_InputDataDebug (owner testing only; never on the master template / schools). Owner sheet: if
  Contact Log column K (Method 2) is hidden from the old site rule, unhide it once by hand.
- 👋 Welcome tab (2026-11-12): re-paste Menus and Build. Nothing to add to the template (the tab is made on the copy).
- 🐞 Debug → 🔄 Copy Sheet and Set to Migrate (2026-11-12; removed again the same day, see ONE TEMPLATE below).
- One template for everyone (2026-11-12): re-paste Menus (script) and MigrationUI, RolloverWizard (HTML) on every
  sheet; owner sheet also z_MINE and z_NewCopy. Move old "(MIGRATE)" files out of the template folder and put the
  (BLANK) copy link in the Version Tracker's column D.
- Release names + version stamp + New/Edit Student header row (2026-11-12): re-paste Core, Build, Regenerate (every
  sheet) and z_MINE (owner sheet).
- 📌 YOUR LINKS box + 🔗 My Links + 📦 Make All 3 Copies (2026-11-12): re-paste Menus, GettingStartedSidebar (every
  sheet) and z_MINE (owner sheet).
- Welcome tab after a reset + permissions (2026-11-12, later): re-paste Menus, Regenerate (script) and
  GettingStartedSidebar (HTML); then run Reset to Display Mode on the master template; upload the new
  docs/PC_Tracker_User_Guide.pdf to the guide folder.
- Wrong paste cleared, Your Lists upload, build before step 5, new-user flow (2026-11-12): re-paste Menus, Build,
  Input Data (scripts) and InputDataSidebar, YourListsSidebar (HTML).
- Your Lists View step + "Save, exit and build" (2026-11-12): re-paste Input Data (script) and InputDataSidebar,
  YourListsSidebar (HTML).
- Version Tracker links (2026-11-12): re-paste Core, Menus (script) and GettingStartedSidebar (HTML) on every sheet;
  owner sheet also z_MINE. In the Version Tracker: name the headers D "Blank Copy", E "Example", F "Document", clear
  the old values in E (the old "Blank Link" column, or they would be read as the example), fill D/E/F in the newest
  row; then Reset to Display Mode on the master template (the Welcome tab gets the example row).
- No "Action Blocked" while a z_ file is there (2026-11-12): re-paste Menus (every sheet).
- Make Release Copies: the Blank's Welcome tab links the new Example (2026-11-12): re-paste Menus (every sheet) and
  z_MINE (owner sheet).
- New-user flow fixes + bug pass (2026-11-12): re-paste Core, Build, Edits, Menus, Input Data (scripts) and
  GettingStartedSidebar, InputDataSidebar, ProgressSidebar, SettingsDialog, NewStudentDialog, ManageOverridesDialog
  (HTML) on every sheet.
- Input Data Mode file names (2026-11-13): re-paste Input Data (script) and InputDataSidebar (HTML) on every sheet.
- Sheet look + 🎨 Apply Sheet Style (2026-11-12, latest): re-paste Core, Build, Edits, Menus, Regenerate on every
  sheet; owner sheet also z_MINE and the manifest from App Scripts (Mine).Json (adds the Sheets service; same
  scopes, no re-approval). Then on the owner's sheet 🐞 Debug → 🎨 Apply Sheet Style once, then Make Release Copies.

NEW-USER FLOW (owner's bug lists 2026-11-12, "act as a new user who is not tech savvy"; walked step by step in
tests/test_input_flow.js "a new user, step by step", loaded WITHOUT z_ / Z# files like the owner's test copy):
- 1. New copy opens: only the 👋 Welcome tab (setup version: permissions, steps), toast -> Initial Setup.
- 2. 🚨 Initial Setup answers at once: runInitialSetupWrapper only runs requireAllScopes and opens
  showGuideWelcome_({ setup: true }) = "Setting up your tracker" spinner window (no link look-ups before it shows:
  zg_runSetup returns { migrate, links } and the window's fillLinks adds them). The window calls zg_runSetup ->
  runInitialSetup_: under the DOCUMENT lock (60 s wait), skipped when SETUP_FINISHED (doc property) = this sheet's ID,
  so a second window / double click never runs it twice (two runs could re-show tabs after Input Data Mode hid them);
  body = runInitialSetupSteps_. A failure shows "Try again".
- 3. During and after setup ONLY the Welcome tab shows until the first build (owner: Contact Log / PCAR / Combined
  showing during setup looked wrong): newCopy (= Student List not built) -> showWelcomeTabOnly_ +
  updateVisualSettings_(settings, columnsOnly = true); a built sheet -> default tabs + settings as before. Last step:
  welcomeTabAfterSetup_ rewrites the tab in place as "You're set up" (writeWelcomeTab_ into, welcomeDoneRows_; C3 =
  WELCOME_DONE_TITLE_). The window asks "Is this your first PC Tracker?"; zg_runSetup opened the checklist.
- 4. "Start Input Data Mode" (and Getting Started's button, the data reminder, Your Lists' Back): only OPEN the
  sidebar (openInputDataModeSidebar_); the sidebar's first call idm_open does the tab work (idm_startOrResume_).
  Owner saw it twice: started from the pop-up, the sidebar opened but the tabs were not switched / hidden; tab work
  from the sidebar's own calls is known to work. The menu item (openInputDataMode, also guideOnOpen_) does it in the
  menu run AND the sidebar checks again. Resume = same screen (idm_focusTempTab_: on the mode's tab, others hidden
  WITHOUT remembering them: only tabs showing at the START come back at the end; step 5 keeps list tabs). Every
  paste-step poll (idm_check): active tab not the paste tab -> idm_focusTempTab_ (one getActiveSheet call otherwise).
  Paste tab: A1 "PASTE HERE" + A2 IDM_HINT (grey: "Closed the side panel? 🚀 App Menu → 📥 Input Data Mode ...");
  neither is data (idm_tempHasData_, idm_processTemp_ strip both). Started from the Welcome tab: it is hidden, never
  remembered, prevActiveId = Contact Log; ending before a build (idm_cleanup_) lands on the Welcome tab.
- 5. Every REAL open (onOpen(e) -> buildMenus_(realOpen); code redrawing menus never moves the user) while not built
  and Input Data Mode off: welcomeTabAfterSetup_(ss, true) shows + opens the Welcome tab with a toast. Input Data
  Mode left on: guideOnOpen_ re-opens its sidebar (and the open toast says the panel is on the right).
- 6. The FIRST build (end of buildAllDerivedSheetsRun_, welcomeTabAfterBuild_; only with setup done, no
  WELCOME_TAB_KEEP and the Student List really built): the starting tabs (startingTabs_ = TAB_VISIBILITY + settings)
  come out (inside Input Data Mode: added to its hiddenIds, shown when it ends), then deleteWelcomeTab_(ss, true)
  (user on it: showStartingTabs_ + leaveWelcomeTab_, deleted next time). A failed / empty build keeps the guide.
  The 10 PM rebuild skips a sheet with no Student List and no ATS rows (nightlyMasterSyncLocked_; no "didn't finish").
- Getting Started: Welcome tab in ZG_TABS ("This tab"); "Hidden tabs" list only after the first build; the "Open this
  by itself" checkbox is now "Finish setup" (zg_finishSetup: ZG_AUTO_OPEN = off + alert "you can always open it:
  🚀 App Menu → ✅ Getting Started", sidebar closes; confirm when steps are left; hidden once all 4 are done).
  zg_setAutoOpen kept for sidebars still open. "Not now" = zg_notNow (toast naming the next step).
- Clone detector also clears INPUT_DATA_MODE_STATE, INPUT_DATA_MODE_DELETE_TAB, ZG_AUTO_OPEN.
- Clumsy-user pass (owner 2026-11-12, "do stupid things and make mistakes"): 🏗️ Build Sheets Only with no student data
  yet (Student List never built + no ATS rows) = a plain note in the Build sidebar (SYNC_UI_STATE killed + notice; the
  sidebar shows notice in plain text) instead of an empty build reported as done; the Welcome tab after setup has a
  warning-only range protection (WELCOME_PROTECT_DESC_, not Sheet Lock's tag); Skip on Student Data never loaded asks
  first; Site / Class Save with nothing filled in never erases BH:BI (error when something was saved, else on like Skip;
  "Save and exit" just exits).

BUG PASS (owner 2026-11-12, "go through every other feature"; tests/test_edits.js + tests/test_features.js):
- Contact Log: a paste wider than one column that covers E looks each row up by its own OSIS (fillPastedContactLogRows_
  numCols); by the parent's name, siblings' rows all got the first child's OSIS.
- Notes tab: the delete box never deletes without the "are you sure?" (no UI = no delete); an OSIS (A) cleared / typed
  over updates the old student's notes (e.oldValue).
- Settings dialog: Save / Reset / Purge are off until getSettings has answered (a quick click or a failed load saved the
  form's defaults over the real settings); failures of the load and the purge are shown.
- New / Edit Student: an OSIS with fewer than 6 digits is refused (every reader skips it, so the student never showed).
  Manage Overrides / New Event: failures are shown (they used to hang on "Deleting…" / "Creating…").
- writeSheet_ pads ragged rows (one short row stopped the whole build).
- Migration: ma_verify refuses this tracker's own link and a sheet with no Contact Log (before anything is copied);
  Raw Data grows to fit last year's rows / columns; Events (copied as values) get each table's look, delete box and
  dropdowns back (restyleEventTables_, Edits.gs; the tab is cleared of old merges / validations / formats first); a new
  copy's tabs wait for step 6's build (updateVisualSettings_ columnsOnly while not built).
- Rollover: archive named for the school year that ends (rolloverArchiveName_: June 2027 = 2026-2027, Sept-Dec = the
  year that started); sets disableNightlySync so someone else's nightly trigger skips the retired tracker.
- Notes line with no method and no person: "𝘿𝙖𝙩𝙚 - note", never "{}" (noteLineHeader_, Core.gs). Run a build once.

First build -> full menus (2026-11-12): refreshMenusAfterFirstBuild_ (Menus.gs) runs onOpen() when setup is done,
FULL_MENU_UNLOCKED isn't "true", Input Data Mode is off and the Student List is built. Called at the end of
startBuildOnlyProcessRun_ (after phase done_build_only, so buildMenus_ shows no "sync in progress" alert; toast
"🎉 Sheets built") and by idm_deleteLeftoverTab. Before, only Full Sync and Migration redrew them; Build Sheets Only
needed a reload. Getting Started texts: step 2 "Build your sheets (Input Data Mode does this for you)".

👋 Welcome tab (Menus.gs, WELCOME TAB section; owner's request 2026-11-12, not yet tried on a sheet):
- New copies only: while SETUP_COMPLETE is not "true", every open (buildMenus_, simple onOpen, before the welcome
  alert) shows the "👋 Welcome" tab and hides every other tab (showWelcomeTabOnly_). The code makes it when missing
  (buildWelcomeTab_: first tab, 27 rows x 4 columns, no gridlines, navy band, 4 numbered steps: App Menu / Initial
  Setup / Allow access (incl. "hasn't verified this app" -> Advanced) / Load your data, or "Bring over your old data"
  on an upgrade copy (isAutoMigrateOn_); User Guide + bug form links (GUIDE_FOLDER_URL, BUG_REPORT_FORM_URL)). An
  existing tab of that name is reused as is. No merged cells (Sheets doesn't grow a merged row for wrapped text).
- The old blocking pre-setup pop-up ("👋 Welcome to the PC Tracker!" / "🎉 Welcome to the newest version!") is now a
  30 s toast pointing at 🚀 App Menu → 🚨 Initial Setup (the Welcome tab has the steps; owner's "make the new-user flow
  simple", 2026-11-12).
- Initial Setup: leaveWelcomeTab_ shows + activates the Contact Log (else the first default tab) and hides it, then
  the normal visibility. It is DELETED by deleteWelcomeTab_ at the next open (buildMenus_ after setup, not while
  IS_RUNNING_SETUP) or build (start of buildAllDerivedSheets), never in the setup call (deleting the tab the user is
  on breaks the session, see Input Data Mode). If the user is on it then (setup stopped half way), it is hidden, the
  default tabs come back, and the next open / build deletes it.
- applyDefaultSheetVisibility_ now shows the default tabs first, then hides the rest (same end result; Sheets refuses
  to hide the last visible tab, which the Welcome tab is on a new copy).
- Not in KEEP_TABS_: every reset deletes it (deleteLegacySheets_ first shows the default tabs when it is the only
  visible one, since Sheets can't delete the last visible tab, and clears the keep mark below).
- Reset to Display Mode (owner's request 2026-11-12, "the welcome the only tab open"): ends with
  showWelcomeAfterReset_ (fresh tab, User Guide link = getGuideLink_, then only it shows), so a copy of the template
  OPENS on it (no flash of the template's tabs before the copy's onOpen runs). On the reset sheet itself setup may have
  run, and the next open / build would delete the tab: doc property WELCOME_TAB_KEEP (= that sheet's ID; not set on
  the owner's "Copy Sheet, Then Reset the Copy" copy, whose setup hasn't run anyway) makes deleteWelcomeTab_ keep it
  until Initial Setup runs there (clearWelcomeKeep_) or the next reset / fake-data load. Reset alert says so.
- Permissions BEFORE the click (owner's request 2026-11-12): the tab starts with "BEFORE YOU START: WHAT GOOGLE WILL
  ASK YOU TO ALLOW", WELCOME_PERMISSIONS_ = the 6 manifest scopes in Google's words (name bold via rich text, reason
  plain), then SET UP YOUR COPY; step 3 = the real screens (Authorization required → OK → account; hasn't verified →
  Advanced → Go to … (unsafe); Select all → Continue). A menu click can't show anything of ours before Google's
  screens, so the explanation has to be on the tab. "Connect to an external service" = only Rollover's PDF download
  (UrlFetchApp); the update check uses the Sheets permission (openById). Keep this list in step with the manifest.
- Initial Setup starts with ScriptApp.requireAllScopes(FULL) (granular consent: an unticked permission is asked for
  again before anything runs; typeof-guarded).
- Getting-started 🚀 App Menu order: ✅ Getting Started, 📥 Input Data Mode, 🏗️ Build Sheets Only, ---, Import Old Data,
  User Guide, Report a Bug, Updates. showWelcomeDialog's "Next steps" point at Getting Started (was "paste into the
  RAW Data tab"). Getting Started links: zg_links_() fills empty DOCUMENT links with the User Guide (download data:
  DATA_GUIDE_URL); an empty VIDEO link is left out (no more "coming soon").
- docs (2026-11-12): Part 1 rewritten for this path, plus Settings labels, Follow-up values, search bar position,
  Events type buttons, Your Lists upload + View step, "RAW Data" -> "Raw Data", the external-service row.

USER GUIDE (docs/, owner's request 2026-11-12 "style of the style guide, less AI looking"): docs/USER_GUIDE.md is the
ONLY source. `node docs/build_guide.js` makes guide.html, PC_Tracker_User_Guide.pdf (Playwright; the jump links are then
turned from named destinations into plain page links with Python pypdf, so Drive / GitHub previews follow them more
often) and PC_Tracker_User_Guide.docx (docx package; headings are real Heading 1-3 so a Google Doc gets heading links;
contents = internal links to bookmarks; footer page numbers, none on the title page). Before, the text lived in three
hand-kept copies (md, guide.html, build_guide.js). Look = Styles.html "District" (navy, Source Sans 3 / Source Serif 4,
thin lines, 3px corners; notes = > [!NOTE] / [!IMPORTANT] / [!WARNING] boxes). Writing: plain short sentences, no
decorative emoji (only menu / tab labels), no em dashes, the sheet's own words; maintainer notes in <!-- comments -->.
Parts: Before you start (what it does, get your copy = the template folder link), 1 Setting up, 2 Daily use, 3 New
versions and the end of the year, 4 Privacy, 5 Help. docs/README.md = how to rebuild and put it in Drive.
Title page "Start here" box (2026-11-12): a > [!IMPORTANT] block before the first "##" is the title page's box
(doc.startBox in build_guide.js, HTML + docx): Make your copy (Blank /copy link) + See an example tracker; both start
as the template folder link and the owner updates them in the Google Doc at every release. The Doc is edited in place
(its link = tracker F + the #bookmark= links in GUIDE_LINKS), never re-uploaded.

📌 YOUR LINKS (owner's request 2026-11-12): every link the owner pastes is in ONE box at the very top of Menus.gs, with
plain-language comments (how to paste, where each shows, empty = no video button / the User Guide opens):
GUIDE_LINKS (setupVideo, downloadDataVideo, downloadDataDoc = the old DATA_GUIDE_URL doc, guideDoc, buildDoc,
settingsDoc, contactLogDoc), GUIDE_FOLDER_ID, TEMPLATE_FOLDER_ID, BUG_REPORT_FORM_URL, VERSION_TRACKER_ID (moved there
from their sections, same names; DATA_GUIDE_URL = GUIDE_LINKS.downloadDataDoc). setupVideo: Welcome tab "🎬 Watch the
setup video" (when set) + checklist Help. linkList_() = all of them with "where it shows" (🐞 Debug → 🔗 My Links,
ownerMenuMyLinks in z_MINE.gs: set / not set + Open links). Keep new links in that box.
VERSION TRACKER LINKS (owner 2026-11-12, "folder = backup"): the Version Tracker's first tab holds the 3 shared links,
one row per release: D "Blank Copy" (the template's /copy link), E "Example" (a filled-in example tracker), F "Document"
(the User Guide). trackerColumns_ finds them by header name (blank|template|copy or "Link" / example / document|guide;
a header not found falls back to D/E/F only when that header is empty or "Column N", so the old E "Blank Link" is never
the example); trackerLinksFromRows_ = per column the newest dated row with an http link; trackerLinks_ (script cache
TRACKER_LINKS_V1, 6 h; never throws, fromTracker false when unreadable; cols = where each column is, for My Links).
computeUpdatePayload_ reads the tracker once (readVersionTracker_) and refreshes that cache too.
Order: blank copy = tracker D, else scanTemplateFolder_'s (BLANK) file, else Backend E3, else the folder; example =
tracker E, else Backend E6 only when the tracker couldn't be read (read fine + no E = no button; the saved E6 is
cleared); guide = tracker F, else the guide folder's newest file, else Backend E5, else the folder. getTemplateLinks_
returns exampleUrl; the "See an example tracker" button shows only with a link: Welcome tab (buildWelcomeTab_
opts.exampleUrl from showWelcomeAfterReset_), the welcome pop-up's fresh view, 📖 Open User Guide (showWelcomeDialog),
Getting Started Help (zg_links_().example / zg_safeExampleLink_). Looked up only when something needs a link, then
remembered 6 h; fresh at the nightly run, 🔔 Check for Updates, 🔗 My Links and Reset to Display Mode; never in the
simple onOpen.
Backend D2:E6 "LINKS (automatic ...)" = the tracker's OWN saved copy (D3:E4 BLANK / upgrade copy links, the same file,
D5:E5 the guide link, D6:E6 "EXAMPLE copy"): written after every successful look-up, read only when a sheet can read
neither the Version Tracker nor the folders. Never the place to paste links; deleting them is harmless. An old second
"GUIDE link" row at D6:E6 is cleared when the guide link is saved (saveGuideLinkToVersionSheet_), then D6 holds the
example.
Guide folder backup = the newest file DIRECTLY in it (guideFolderFiles_ / scanGuideFolder_: a MM/DD/YY date in the
NAME counts first, else the last change; ties: the one Drive lists last), remembered 6 h (script cache GUIDE_LINK_V1;
the nightly refreshResolvedLinks_ clears it). 🔗 My Links and Reset to Display Mode's Welcome tab look again first
(refreshResolvedLinks_). My Links shows the tracker's D / E / F first (which column, or 'no column named "Example"'),
then the folders as backup, the guide folder's files in rank order with why.
🐞 Debug → 📦 Make Release Copies (ownerMenuMakeAllCopies, z_MINE.gs; was "Make All 3 Copies"): a pop-up makes
"… - (EXAMPLE)" (fake data, as "Copy Sheet, Then Load Fake Data"; then every tab shown except Backend, any Welcome
tab deleted, ownerShowAllTabsButBackend_; the copy gets developer metadata PC_EXAMPLE_COPY (EXAMPLE_COPY_MARK_,
Menus.gs) so showWelcomeTabOnly_ does nothing on it when it opens before setup; a Reset to Display Mode removes it) and "… - (BLANK)" (B3 = No + display reset), ONE
google.script.run call each (ownerMakeCopy(kind); each run has its own 6-minute limit), showing each link; a failed one
doesn't stop the others. This sheet is never changed. The done message shows the Version Tracker row to add: A = the
version date, D = the Blank copy's /copy link, E = the Example copy's link (trackerLink).
Nobody runs Initial Setup on the Blank copy and its Welcome tab's links are written by its reset (then kept in every
school's copy until it runs setup), so the pop-up makes the Example FIRST and passes its link to ownerMakeCopy("blank",
exampleUrl): WELCOME_LINKS_OVERRIDE_ (Menus.gs, { exampleUrl, guideUrl }, read by showWelcomeAfterReset_) is set for
that one reset, and ownerSaveCopyLinks_ writes the copy's own Backend D2:E6 (Blank / upgrade = its own /copy link, that
Example, the guide now; saveTemplateLinksToVersionSheet_ / saveGuideLinkToVersionSheet_ take an optional ss). The
guide link on it is the tracker's F at that moment, so a new guide's link goes in F BEFORE making the copies. After
a school's Initial Setup every link is read live (the Welcome tab is then deleted).
ONE TEMPLATE FOR EVERYONE (owner 2026-11-12; there is no MIGRATE template any more): the first window after Initial
Setup (showGuideWelcome_) asks "Is this your first PC Tracker?": "Yes, start fresh" = the old welcome (4 steps, Input
Data Mode), "No, bring over last year's tracker" = zg_action('importOld') -> openMigrationWizard (title "Bring Over Last
Year's Tracker"). The wizard (MigrationUI.html) lists the trackers in the user's Drive (ma_findOldTrackers, Menus.gs:
DriveApp.searchFiles on spreadsheet names with MA_TRACKER_NAME_WORDS_ incl. Rollover's "Retired on", "'me' in owners"
with a fallback without it, this sheet left out, newest 8; never throws) to click; pasting a link stays. The Welcome
tab's step 4 says the question is coming. Old upgrade copies (B3 = Yes) still work as before (straight to the wizard;
isAutoMigrateOn_ kept). Template links: scanTemplateFolder_ sets migrateCopyUrl = the (BLANK) file (a (MIGRATE) file
is only used when there is no BLANK one); Backend D4 says "UPGRADE copy (same file as BLANK)". Rollover: one button
"Get next year's tracker" (+ "choose No, bring over last year's tracker"); 📖 Open User Guide: one "Make a new copy".
Release copies (2026-11-12, owner): Make Release Copies names the copy
"Parent Coordinator Sheet - (TYPE) (MM/dd/yy)" (OWNER_RELEASE_NAME_ / ownerReleaseName_, never this sheet's name; the
two test-copy items keep "<this sheet's name> - Display Mode / Fake Data (date)"). Every release copy gets its version
stamp (ownerStampVersion_: Backend A2 = today via stampVersionDate_, B3 = No); the Example copy had
kept this sheet's old date (fake data has no reset). The pop-up / done message shows that date: it is the date for the
Version Tracker's column A (a later tracker date makes brand-new copies say "New Version Available").
Version Tracker (VERSION_TRACKER_ID, first tab): A date, B Size "Major..." = major, C Notes, D Blank Copy = "Get the
latest version" (the newest dated row's link), E Example, F Document (see VERSION TRACKER LINKS). Older checkers read
only A:D by position, so D must stay the blank copy's link.

Student List tab (was "Master Table"; renamed 2026-11-05 at the owner's request):
- Every lookup goes through getStudentListSheet_ (Core.gs), which renames an old copy's "Master Table" tab the first
  time, so its Discharged history carries over. Never getSheetByName("Master Table") directly: writeSheet_ would make
  a second, empty tab next to it, and a reset would delete a tab missing from KEEP_TABS_. "Old Master Table" (owner
  only) is a different tab and keeps its name. Function / variable / key names were NOT renamed (buildMasterTableData_,
  masterTable, findInMasterTable_, Settings key showMasterTable).
- Lists built when the files load spell the name out (file load order must not matter): TAB_VISIBILITY
  ["Student List", "Master Table"], KEEP_TABS_, SHEET_LOCK_BUILT_TABS, Getting Started's tab list (ZG_TABS), BUILD_STEPS. The search
  bar formula (SEARCH_BAR_COMMS_FORMULA) reads 'Student List'!A:B; restoreSearchBarCommunicationsFormula renames an old
  tab first. A tab renamed by hand / by setName keeps its protections, data validations and formula references.
- A:H only (owner 2026-11-12): every reset deletes the columns after H (the lists are cleared too); a build adds a
  Your Lists column back through writeSheet_ and gives its header the Status header's look.
- ONE ROW PER STUDENT (owner's request 2026-11-11): the tab has no sibling-group rows ("111, 222" / "Ann X & Ben X");
  the Sibling column stays. The build keeps the full table in memory (Combined Contact Tracking, Phone Contacts,
  New Students, build report unchanged), writes it OSIS..Status WITH the group rows to the Backend STUDENT_LIST
  section (BX:CD), then writes the tab filtered by isSiblingGroupRow_ (OSIS cell has a comma).
- The autofill reads the Backend copy (so a sibling group can still be typed / picked): studentLookupSource_ /
  studentLookupRows_ / studentLookupColumn_ (Core.gs) feed the Contact Log ghost typist (ghostTypistData_, also
  knownStudentOsis_), the paste lookup (buildGhostLookup_), the Notes / Events typists (findInMasterTable_(src, ...),
  findMasterRowByOsis_ on data rows), getExistingMasterTable_ (quick updates) and the Events dropdowns
  (eventDropdownRules_: Backend!BX3:BX / BY3:BY / BZ3:BZ). Until a build has written BX (header "OSIS" in BX2) they
  fall back to the Student List tab (row 5+), whose old build still has the group rows. Every build also runs
  pointEventDropdownsAtBackend_ (Edits.gs): Events columns whose row-3 rule still lists the Student List tab are
  repointed (one row read after the first time). Readers that skip comma rows anyway read the tab: Discharged
  history, status snapshots, Your Lists index, Hide Discharged / Purge, isBuilt checks, Contact Log Debug anonymizer.
  The search bar's Communications formula SPLITs the OSIS, so 'Student List'!A:B still works.
- Header on ROW 1, data from row 2 (owner moved it 2026-11-11; was row 4 / row 5): OSIS, Student, Guardian, Site, Class,
  Sibling, Status, Last Interaction (H). STUDENT_LIST_HEADER_ROW (Core.gs) = where the build writes it. Readers find the
  header themselves (studentListHeaderRow_: "OSIS" in A1:A10, else row 1), so a tab not rebuilt since still works:
  studentListIsBuilt_ (menus + Getting Started), the lookup fallback, resetMasterTable_ (clears under the header),
  getAllStudentsForEdit and ownerReadMasterStatusSnapshot_ (start under the header), buildMasterTableData_ /
  readMasterStatusSnapshot_ (already scanned for "OSIS"). Your Lists' sidebar index (od_indexFromSheet_) and the Contact
  Log Debug anonymizer read the Backend copy (studentLookupRows_). Hide Discharged / Purge start at
  studentListHeaderRow_ (2026-11-12).
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
  BN:BQ Additional Information and BS:BV Your Lists setup (both Your Lists, Input Data.gs), BX:CD Student List (OSIS..Status incl.
  the sibling-group rows; every build; read by the autofill). PLAIN RANGE, no Sheets Table (owner's choice 2026-11-11):
  the Contact Log's dropdowns read Backend!$BX$3:$BX (OSIS) / $BY$3:$BY (Student) / $BZ$3:$BZ (Guardian), open-ended, so
  every student is included. A Table was tried first and dropped: it had its header on row 3 (it took the first student
  "249004359 | Alexa Abbey ... | Column 6 | Column 7" as its column names), so every build wrote students into the Table
  header; Apps Script also cannot resize a Table (no Sheets advanced service). Left from then (harmless): keepHeader
  (backendWriteBlock_ writes the header row only when it differs, rewrites rows 3+; resetDerivedSheets_ clears rows 3+)
  and the build's try/catch around this write (console.warn + toast), so a Backend problem never stops the build. The last three are "grow"
  sections: an existing Backend tab gets their title the first time they are written (backendBlockReady_); before that
  reads return nothing and clears do nothing. A new / migrated Backend tab gets every title at once (writeBackendTitles_).
- New/Edit Student (G:AV) = title row 1, header row 2, students from row 3 (like every section; the old separate tab had
  its header on row 1). The owner's sheet had it one row lower (header row 3): readers skip rows without a real OSIS,
  but a reset clears rows 3+ = that header, so copies lost it. fixOverridesHeaderRow_ (Core.gs, OVERRIDES_HEADERS_ =
  the 42 names) moves such a section up (G:AV only, never over a student row): at the start of every build and in
  resetDerivedSheets_ before the clear; after the clear it writes the headers into an empty row 2.
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
- Data update reminder (showDataReminderDialog_) restyled; Settings has it under Student data. InputDataSidebar
  and YourListsSidebar use Styles.html too since 2026-11-12 (same class names, colors from the shared tokens, no emoji
  inside; menu paths like "🚀 App Menu" keep theirs; server status emoji are turned into ✓ / – markers).
- Create Event: type buttons (PTC, Workshop, Family event, Meeting) only fill the Type box. New/Edit Student: one grouped
  form, entry type as two buttons over a hidden #entryType input. Manage Overrides: cards. Rollover: step markers.

SHEET LOOK (owner's pick 2026-11-12; SHEET_LOOK_ in Core.gs = every color / font the code paints on a tab):
- Montserrat everywhere. Tabs you type in (Contact Log, Events, Notes, Send Out) = the Contact Log's brown header
  (~#6b4e37) + cream rows; tabs the build makes (Student List, Parents Divided, Directory, Combined Contact Tracking,
  Phone Contacts, Students Override Report) = navy #183e69 (the dialogs' navy) + light grey #f3f5f8 rows; PCAR =
  a slightly darker brown (owner 2026-11-13, SHEET_LOOK_.pcar: titles #5b422f, This / Last Week dates #4a3627,
  section rows #efe3d3 with #5b422f text, own-dates cells #fff2cc; the owner may also set these by hand).
  Red / amber / green only where they mean something. The Contact Log keeps its hand-made look (table colors, lookup
  box, END bar, week / today rules); its Workshop rows are pale green now (SHEET_LOOK_.workshopRow, owner 2026-11-12).
  Navy everywhere = give typing the built values.
- Dropdown chip colors (by hand, Table column settings; owner was given this scheme 2026-11-12): Follow-up N = dark
  red ASAP (the only dark chip), light orange Needed, light yellow Waiting, light green Resolved, light gray No Followup
  / Skipped (most rows = quiet); Method J = light chips only, Phone light gray (most rows), Walk In light blue,
  Workshop light green, Workshop Attendee light cyan, Parent-Teacher light purple, Internal light yellow.
- Code that paints (so a by-hand change there would come back): Events tables (styleEventTableAt_, Edits.gs: brown
  title row with white text over name / type / place / date, delete box #f8e9e7, header row #efe3d3, thin #e3d8ca
  lines + brown outline; used by buildEventTable, restyleEventTables_ (Migration) and 🎨 with lookOnly; was green /
  tan / yellow / blue / red + grey + black lines, typed out twice); Contact Log Type colors (applyTypeDropdownColors:
  SHEET_LOOK_.typeColors by Type List place, no write when the L rules already match; they were random and
  resetContactLogFormatting re-rolled them every night); Contact Log rows: formatContactLogRows_ (Core.gs) sets
  Montserrat on every row it formats (edits, the new blank row, resets) and resetContactLog_ on the header (owner
  2026-11-13: Make Release Copies' copies showed Roboto; was only column D, on edits, in Roboto); Directory site
  colors (formatDirectoryAuto_: SHEET_LOOK_.siteColors); Students Override Report (styleOverrideReport_, Menus.gs:
  navy header, Montserrat; its tab is cleared every rebuild); Raw Data (forceVibrantColors_, see Regenerate.gs).
  Your Lists headers after a table: styleLastInteractionHeader_ (Student List: every header from Last Interaction on)
  and Parents Divided's after Last Interaction copy the header before them (styleHeadersLike_; a source with no fill
  of its own gives no fill, never white).
- 🐞 Debug → 🎨 Apply Sheet Style (ownerMenuApplySheetStyle, z_MINE.gs; owner's request "a button in debug to fix
  colors and fonts"): the hand-made parts, once, on the owner's sheet (Make Release Copies carries them on):
  Montserrat on 12 tabs (Contact Log: header + rows above END, not the END bar / lookup box); Sheets Table colors through the Sheets API (ownerSheetsApi_:
  the "Sheets" advanced service, in App Scripts (Mine).Json only, else UrlFetch to sheets.googleapis.com; updateTable
  rowsProperties header / first / second band) = Notes + Send Out get the Contact Log TABLE's own colors (or its
  header cell's fill), the 4 built tables navy; Send Out's table "Table1" -> "Send Out"; header rows painted with the
  table color + white bold (over Send Out's old D1:E1 fills); Phone Contacts: header painted first, then each banding
  recolored on its own with a flush (Google refused it on the owner's sheet 2026-11-13: "Service Spreadsheets failed
  while accessing document"; a refusal = the row colors listed by hand with the values, not a failed step); widths (Student List / Parents
  Divided fit within [OWNER_STYLE_WIDTHS_ min, 320], Send Out B 240); PCAR (boxes = row-1 titles over 2 columns;
  all SHEET_LOOK_.pcar brown, the own-dates box's date cells yellow #fff2cc, a missing section title copied from
  the first box, labels plain + wrapped); the Student List's empty columns after
  the last header ("Column 9" counts as empty; Your Lists columns stay); tab "RAW Data" -> "Raw Data" (every lookup
  takes both); then repaints the code-painted parts now (the Contact Log's rules via resetContactLogFormatting, as the
  10 PM run does). Reports Send Out dropdowns that read Backend BH/BI/BJ from
  row 4 (the lists start on row 3: the first choice shows a red corner) or a typed list without the first choice.
  Without the Sheets API: the table colors are listed by hand with the hex codes. Never changes values except the
  PCAR title and the two names. By hand only: dropdown chip colors (Contact Log J / N), the Contact Log's widths, tab
  colors (the owner removed Color Tabs before). Tested offline (tests/test_style.js, fake Sheets API); not yet run on
  a real sheet, so the Sheets API's table fields are untested against Google.

SHEET LOCK (Sheet Lock.gs; on by default, doc property SHEET_LOCK_OFF = "true" turns it off). WARNINGS ONLY (owner's
choice 2026-11-05): the code never puts cells back after an edit; the earlier undo layer (snapshots, onSelectionChange
copies, "clear 2+ rows?" popup, column A auto-restore) was removed because its extra script writes made Ctrl+Z need
several presses. Don't bring it back without asking.
- All are RANGE protections (owner's pick 2026-11-05: a whole-tab SHEET protection puts a 🔒 icon on the tab). Built
  tabs = one range over the whole grid; PCAR = 4 ranges around H2:I2. A range covers the grid as it was when locked;
  rows a build adds are covered when applySheetLocks_ runs at the end of the build. Old SHEET ones are removed as stale.
- Google "Heads up, are you sure?" warning-only protections (description starts "🔒 Sheet Lock"; scripts are not
  blocked) on: row 1 of Contact Log / Notes / Send Out; Contact Log END bar row (except O, the FILTER? checkbox) and
  everything below it except the lookup box cells typed in (lookupBoxInputCells_: the OSIS box right of the "OSIS"
  label and the name switch one row up / two columns right; 2026-11-12, rangesAround_); row 2 (headers) of every Events table (EVENT_TABLE_HEADERS); PCAR except H2:I2; the whole built
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
  2. Ghost typist from the cache (Student List OSIS / Student / Guardian from the Backend copy BX:BZ + Directory C:E, ghostTypistData_, key GHOST_TYPIST_DATA, 6 h,
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
  and flushes before the checkbox / sync. Events typist unchanged. Notes / Events lookups (findInMasterTable_): TextFinder
  first, then (only when it misses) an in-memory scan that ignores stray spaces; the Contact Log typist compares the
  same way (typistKey_: trimmed, inner spaces collapsed, any case; owner 2026-11-12).

Workflow:
- OFFLINE TESTS (tests/, 2026-11-12): `cd tests && npm install && npm test` runs check.js (whole project compiles as
  one script, no duplicate names, every HTML file / google.script.run call / menu item resolves) and the logic tests
  (test_matching, test_fixes, test_sync with a fake People API, test_style with a fake Sheets API; lib/fakesheet.js
  logs every formatting call in ss._log and keeps fills, merges, widths, rules); render_sidebars.js = optional Playwright screenshots.
  lib/load.js loads the .gs files into one Node context with fake Apps Script services. Run them before every push
  and add a test for logic you change; they don't cover sheet reads / writes / dialogs (try those on a copy).
- TEST DATA (test-data/, 2026-11-12): fake exports for every Input Data Mode step (ATS incl. a "next month" file with
  2 new / 1 left and a broken one with cut-off numbers, Parent Square, NYCSA, Site and Class) and 4 Your Lists files
  (bus routes = student list with title rows / typos / initials / wrong OSIS, PTA = guardian list, field trip =
  first + last columns, after school = "LAST, FIRST" + nicknames), each .csv + .xlsx, plus "PC Tracker test data.zip".
  Same 18 students as Regenerate's fake data (make_test_data.py reads them from Regenerate.gs; needs openpyxl).
  README.md there says what each file should show; tests/test_testdata.js checks those counts.
- DEPLOY.md (repo root): the owner's one-page checklist (files to delete, the files each copy should have, manifests,
  after-deploy steps, a quick test, going live). Keep it in step when files are added / renamed.
- Push every change to Test-Version (not main) unless the owner says otherwise for that change. main only gets
  changes when the owner asks for it (e.g. merging Test-Version once it has been tested on a copy of the sheet).
- OTHER CLAUDE SESSIONS (owner's request 2026-11-12): two sessions pushing to Test-Version at once collide. At the
  start of every task AND right before every push: (1) list the owner's sessions (claude-code-remote list_sessions,
  mine: true) and (2) git fetch origin Test-Version and look at commits not in this checkout (their
  "Claude-Session:" trailer names the session). If another session is running / working, or pushed commits this
  session has not built on, WARN THE OWNER right away (which session, what it pushed) before going on; rebase or
  redo this session's change on top of theirs, never overwrite them.

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
Must keep: 15 menu handler names (+6 owner Debug items, + toggleSheetLock), 39 google.script.run endpoints, trigger
handlers (installedOnEdit, installedOnChange, nightlyMasterSync, checkVersionOnOpen, syncToGoogleContacts),
all sheet names (except the owner's Master Table -> Student List rename), Backend cells,
property/cache keys.
Owner decisions still open (current behavior kept): owner splitFullName_ changes the shared Directory name split
(owner: keep it).
Decided 2026-11-12: site column = H on the owner's sheet too; Your Lists carried over at Migration and archived in
the Rollover PDF (processRolloverBackend copies each list tab into the archive; next year's sheet gets them through
Migration); the owner sync has the label pageSize + "more than half" check; Montserrat on New Students (owner's
montserratSheets); ghost typists ignore stray spaces (typistKey_).
Data update reminder (Menus.gs, replaced the old 30-day ATS-only reminder; same checkRawDataFreshness_ name): centered
pop-up when Raw Data A1 / AG1 / AP1 (ATS / Parent Square / NYCSA, only sections ever loaded) is older than setting
"dataReminder" (week default / 2weeks / month / off; in SETTING_KEYS so it is saved to the Backend settings block and
shown in ⚙️ Settings). Checked by checkVersionOnOpen and by installedOnEdit (checkDataReminderWhileOpen_, user cache:
once an hour); once a day (ATS_REMINDER_LAST_SHOWN); not during Input Data Mode. "Go to Input Data Mode" (always, since
the coalesce; DATA_GUIDE_URL is kept but no longer linked). Skip -> Remind After 2 Weeks / 1 Month (setting + wait in doc
property DATA_REMINDER_SNOOZE_UNTIL) / Remove Reminders / Just skip for today. Endpoints dataReminderOpenInputMode,
dataReminderChoose. Saving a different choice in Settings clears the wait.
Settings autoOverride, ignoreHideCheckboxes, emailNotifications were saved but never read: removed from getSettings'
defaults and SETTING_KEYS (owner 2026-11-12). Old values stay in existing sheets' properties, unused; loading the
Backend settings block skips keys not in SETTING_KEYS.
Nightly trigger (nightlyMasterSync): 10 PM, NIGHTLY_SYNC_HOUR (Menus.gs), both at Initial Setup and when a Settings
save re-makes it (owner's pick 2026-11-05; was 11 PM at setup but 2 AM after a Settings save). Existing triggers keep
their old hour until Settings is saved once or Initial Setup is run again.
ONE nightly job per sheet (2026-11-12, Build.gs "ONE NIGHTLY JOB PER SHEET"): triggers are per user, so a second
person's Settings save used to add a second nightly (two builds, contacts synced into their account). Doc properties
NIGHTLY_TRIGGER_UID (+ _SET) / NIGHTLY_LAST_RUN: a firing trigger that isn't the remembered one deletes itself
(e.triggerUid; it runs as its maker), unless the remembered one hasn't run for 3 days (then it takes over);
manageNightlyTrigger_ makes none for another person; checkVersionOnOpen (runs as the setup person) remembers their
trigger on older sheets (claimNightlyTrigger_); the trigger itself checks disableNightlySync when it fires.
A failed / skipped nightly run (error, or the build lock busy) leaves doc property NIGHTLY_FAILED_MSG ("Last night's
rebuild didn't finish / didn't run: <reason> ..."); buildMenus_ toasts it on the next open (before the sync message,
which then waits for the open after), and any finished build clears it (end of buildAllDerivedSheetsRun_).
BUILD / SYNC LOCK (2026-11-12, withBuildLock_ in Core.gs): ScriptLock around buildAllDerivedSheets (body now
buildAllDerivedSheetsRun_), syncToGoogleContacts, Full Sync (build + sync as one job) and the nightly; re-entrant within
a run; busy = Error.buildBusy "Another \"X\" is running right now" (Build Only / Full Sync show it in the sidebar; a busy
resume re-schedules itself in 2 min). Contact Log edits / Input Data Mode use the Document lock, so builds never hold
them up. Full Sync errors end as 'killed' with the error (was stuck on 'building' for 6 h).
CONTACTS SYNC (rewritten 2026-11-12, owner's "fix 1-3"; see "HOW THE SYNC WORKS" in Contacts Sync.gs):
- One Connections.list pass per run (buildContactsIndex_: byResource / byName "given|family" / label members); rows
  find their contact there (findIndexedContact_: the 6 h row cache first, People.get only for a cached contact not
  listed yet; owner: name AND phone / email filter). A contact identical to its row (contactMatchesPayload_: names,
  phones + types, emails + types, notes, organization, label) is not written; the 0.75 s sleep only after a write.
- searchAndUpsertSingleContact / ownerUpsertContact_ return { failed: true, resourceName } instead of null and THROW a
  quota error (isQuotaError_), so the loop's 15-minute pause works. A failed row's contact is kept, and any failure
  in a sync = no deletions this time (note shown). Deletion = label members not used by a row of this run nor
  (rows done in an earlier run) any contact with a row's names (protectByName).
- "More than half" (label > 10 contacts): worked out at row 1 of a sync from the whole sheet; pop-up if someone is
  there (askMassDeletion_), else blocked + note; SYNC_DELETE_DECISION = confirmed / blocked / none-needed, read by the
  clean-up of the last run (asks again only for none-needed). Now on the owner's sheet too.
- Time: SCRIPT_START_MS_ (Core.gs, set when the files load = run start); rows stop at 4.5 min after it, and with less
  than 45 s left at the start (a long build) the sync starts in a fresh run (SYNC_START_ROW kept, so a busy lock
  re-schedules it). Contacts can't be read: retry in 15 min (3 times), then stop with a message, nothing changed.
- State in script props: SYNC_START_ROW / _CREATED_COUNT / _EDITED_COUNT / _FAILED_COUNT / SYNC_DELETE_DECISION /
  SYNC_LIST_RETRIES (clearSyncRunState_, also used by EMERGENCY_STOP_SYNC). The old STALE_CHUNK_* cache is no longer used.
  SYNC_UI_STATE 'done' carries failed + note (ProgressSidebar shows them; log list "failed").
- Tested offline with a fake People API (test_sync: normal, failed save, quota pause + resume, over-half blocked /
  confirmed, late start, owner rows); not yet run against real Google Contacts.
The owner's summary email (MailApp + Session.getActiveUser().getEmail(), z_MINE.gs) needs the script.send_mail and
userinfo.email scopes: they are in "App Scripts (Mine).Json" only (see MANIFESTS under Deploying).
Fixed 2026-11-12: quick notes ignored the Settings date range (notesDateRange_ / inNotesDateRange_, Build.gs, used by
the build and notesTextForOsis_); getExistingMasterTable_ now has a header row; EMERGENCY_STOP_SYNC clears
SYNC_EDITED_COUNT; TOGGLE_EMAIL_NOTIFICATIONS turns on with the first click; purgeDischargedStudents and the Hide
Discharged filter start at studentListHeaderRow_; generateFakeLegacyData adds the 47 columns first; Full Sync stuck
on "building" + no build / sync lock + per-user nightly triggers (see above); Z#ContactLog follow-up values; Sheet Lock
leaves the lookup box inputs free; Input Data Mode clears the temp tab fully; Migration maps old Follow-up values; a
Contact Log edit waits 1 s only when the OSIS cell says "Loading..."; the owner sync counts a contact as edited only
when its phones / emails / notes / site really changed (ownerContactChanges_).
