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
- Core.gs: Contact Log column constants, findEndRowInColumnE_, Backend tab helpers, owner hook point
  (getOwnerHooks_, runAction_), small text / phone / OSIS / date helpers.
- Build.gs: buildAllDerivedSheets + everything it builds (Master Table, Combined Contact Tracking, Parents Divided,
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

- Z#Guide.gs (+ Z#GuideSidebar.html): optional "✅ Getting Started" (bare-bones version, owner's pick 2026-11-05; a full
  tour / tab pop-ups / Start Here tab / header notes were proposed and NOT built). 3 guarded hooks in Menus.gs:
  addGuideMenuItems_ (both 🚀 App Menus), showGuideWelcome_ (runInitialSetupWrapper, instead of showWelcomeDialog on
  new-user copies; "📖 Open User Guide" still opens showWelcomeDialog), guideOnOpen_ (checkVersionOnOpen).
  Welcome pop-up after setup: 4 steps, "Start with Input Data Mode" (only if openInputDataMode exists), video / guide
  links. Checklist sidebar: 1 load data (ticks when Raw Data A1 has a real date; also shows Parent Square AG1, NYCSA
  AP1, Sites and classes BH7:BI7 / INPUT_DATA_MODE_SITECLASS_UPDATED), 2 build (Master Table lastRow > 4, same test as
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
  "Put formula back". "Anonymize" (on by default, cld_report(anon) / cldAnonymizer_): Master Table names ->
  Student-N / Guardian-N / Name-N, OSIS -> OSIS-N, phones / emails masked, note texts -> "[N chars, K lines, #hash]"
  (cldNoteSummary_, same text = same hash), Contact Log D -> (filled)/(blank), unknown typed text -> [typed text].
  Recording = shared logger in Core.gs (clDbgStart_ / clDbg_ / clDbgEnd_ / clDbgRun_, wrapped
  around installedOnEdit / installedOnChange): on only while doc cache key CL_DEBUG_ON exists (the sidebar re-puts it
  every 2 s poll, 5-min TTL); runs in CL_DEBUG_LOG (newest 25). Off = one cache read per edit.
- Search bar formula self-repair: in formula mode refreshSearchBarCommunications_ calls ensureSearchBarFormula_, which
  puts SEARCH_BAR_COMMS_FORMULA back ONCE if the Communications cell has no formula (e.g. left as text by the code
  mode), then sets SEARCH_BAR_COMMS_MODE = "formula" so it never checks again. Runs on a search bar edit / build.

Owner-only file (in the repo, deployed ONLY to the owner's own sheet): z_MINE.gs
- Never copy it into other schools' sheets. It defines ownerHooks_(); shared code calls getOwnerHooks_() (Core.gs),
  which returns those hooks on the owner's sheet and null everywhere else. No shared function is redefined, so file
  load order does not matter. Shared code must work without z_MINE.gs; owner-only behavior (e.g. "Old Master Table" /
  "Old Phone Contacts") goes into a hook in z_MINE.gs, never into the shared files.
- Hooks: Master Table status snapshot, Old Master Table history (Master Table + Directory), Old Phone Contacts merge,
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
  RAW Data, Contact Log, Notes, Events, PCAR, Send Out, Master Table, Combined Contact Tracking, Parents Divided,
  Directory, Phone Contacts, Students Override Report, Backend), incl. Old Master Table / Old Phone Contacts and New
  Students (owner's choice: the owner's next build re-creates it). It runs getBackendSheet_ first so old Version / Type List etc. tabs are migrated before they are deleted.
- All three resets also clear the Notes tab (resetNotes_, A:D rows 2+).
- Contact Log student lookup box (2 rows below END: OSIS / Student / Communications / Student Information + guardian
  table): resetContactLog_ ends with fixContactLogLookupBox_, which deletes leftover "Parent Contact (COPY)" header rows
  (+ the row under it when blank / #REF!) and redraws a thick black outline (left = "Communications" column, right = end
  of the merged "Student Information" title, top = END + 2, bottom = last guardian row under "Relationship").
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
  Waiting - Check Up), always (contactLogTodayOnly no longer changes it). The old "Yes - Important" bold rule in resetContactLogFormatting is still there (now unused).
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
  UI refresh: also add Styles.html (new) and re-paste SettingsDialog, ProgressSidebar, NewStudentDialog,
  ManageOverridesDialog, RolloverWizard, MigrationUI, VersionDialog. Regenerate (re-paste: KEEP_TABS_ tab clean-up, Color Tabs removed), the Z# files
  and the manifest are unchanged. The Contact Log tool is now named
  Z#ContactLog / Z#ContactLogSidebar (delete the old #ContactLog / ContactLogSidebar files, or they define it twice).
- Input Data Mode: also paste Z#InputDataMode (script) + Z#InputDataSidebar (HTML); optional, menu item only appears
  with them.
- Getting Started: also paste Z#Guide (script) + Z#GuideSidebar (HTML); optional, needs Styles.html; works best with
  Input Data Mode.

Backend tab (see Core.gs):
- One hidden "Backend" tab replaces the old Version, New/Edit Student, Backend_Event_Log and Type List tabs.
  Row 1 = section titles; each old layout sits one row lower. A:E Version & Settings, G:AV New/Edit Student,
  AX:BF Event Log, BH:BJ Send Out list (BH3:BJ3 = the owner's formulas for the Send Out dropdown options; resets
  clear BH:BJ rows 3+ and write SEND_OUT_LIST_FORMULAS_ back via writeSendOutListFormulas_), BL Type List.
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
- detectAtsChanges_ (shared, runs right after the Master Table) is the single source for new / discharged students:
  the owner's New Students tab, the build report
  (sidebar) and the owner's nightly email all use it. Discharged = "Current" in the old Master Table, in Raw Data at
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
  tabs (Master Table, Combined Contact Tracking, Parents Divided, Directory, Phone Contacts, Students Override Report,
  New Students). NOT on: Raw Data, Backend, owner-only tabs, Contact Log entry rows (incl. column A, so deleting one
  row never warns), Notes / Send Out rows 2+, Events row 1 and rows 3+.
- applySheetLocks_ adds/refreshes them (only our tagged ones; unchanged ones kept, stale ones removed; never takes over
  someone else's sheet protection) at the end of every build, at Initial Setup, after buildEventTable / an event delete
  (Events only) and buildOverrideReport. It also deletes the old version's doc properties (SHEET_LOCK_SNAP_*, SHEET_LOCK_CL_END).

Contact Log search bar (rows below END; its other cells are formulas typed in the sheet, not in the repo):
- THE CELL DECIDES (2026-11-05; owner wants "no formula"): if the Communications cell holds a formula, the formula
  draws the box and the script does nothing; if it is plain text, the script writes the notes one column right. Doc
  property SEARCH_BAR_COMMS_MODE is no longer read (it was set to "formula" while the cell had no formula, leaving the
  box blank). restoreSearchBarCommunicationsFormula puts the formula in; useCodeForSearchBarCommunications turns it
  back into the plain label and draws. ensureSearchBarFormula_ / isSearchBarCodeMode_ were removed.
  readSearchBarBox_ = 2 calls (3x3 block display + formulas around the label); the label's place (rows below END,
  column) is cached in doc cache CL_COMMS_AT_<sheetId> and checked against the block. A row edit redraws the box
  right after the quick notes write with that write's own texts (LAST_NOTES_TEXTS_), no second read; "▶ Name" uses
  the ghost typist's cached Master Table. The typist's OSIS is reused (typistOsis) instead of re-reading E.
- Notes sources cache (notesSourcesData_): Backend Event Log (display) + Notes tab (values, not cached if a cell is a
  Date) in doc cache NOTES_SRC_EVENTS / NOTES_SRC_NOTES for 10 min (skipped > 95 KB). Cleared by every EVENT_LOG
  write/clear/append/delete/sort/range helper in Core.gs, any Events / Notes edit (handleEdit_) and every build. On
  the owner's sheet each sheet call cost 0.4-1.5 s (1,243-row Contact Log), so fewer calls is the main speed lever.
- Speed (2026-11-05): a single-cell Contact Log edit whose row has an OSIS runs the quick notes update
  (syncNotesForRowOsis_ -> updateOsisDeltas_) RIGHT AFTER the ghost typist / autofill, before the row repair /
  blank row / END sweep, so Combined (and the formula) update sooner; the old end-of-edit sync still covers pastes
  and rows whose OSIS only appears later. Ghost typist: Master Table A5:C + Directory C:E cached in the document
  cache for 10 min (ghostTypistData_, key GHOST_TYPIST_DATA, skipped if > 95 KB), cleared at the end of every build
  (clearGhostTypistCache_); Regenerate's resets don't clear it (10 min max).
  END row: contactLogPrefetch_ gets the remembered END row (doc cache CL_END_ROW_<sheetId>) + the typist data in ONE
  cache.getAll, trusts it after a one-cell "END" check (else scans column E); rememberEndRow_ after the blank-row
  insert and the END sweep. The typist writes only the cells it fills (C / E:F / H:I) without reading the row first,
  so a warm edit makes 1 sheet read before its flush.
  Early sync = updateCombinedNotesQuick_ (one Combined cell from the logs; the shared body of updateOsisDeltas_). On
  the owner's sheet the heavy ownerUpdateOsisDeltas_ (re-reads Raw Data, rebuilds Combined + Phone Contacts in memory)
  runs at the END of the edit instead (PENDING_OWNER_NOTES_), and the early write never falls back to
  updateContactAndPhoneOnly there (that writes Phone Contacts in the shared layout).
  The early notes sync is in a try/catch (an error is logged and the end-of-edit sync retries), so the row clean-up
  always runs. Notes typist: writes only the cells it fills (no row read) and flushes before the checkbox / sync.
  Events typist unchanged (already minimal for its TextFinder lookup). TextFinder lookups (Notes / Events) kept on
  purpose: switching to the in-memory cache would change matches with stray spaces.
- "Communications" box drawn by code (CONTACT LOG SEARCH BAR section of Edits.gs; owner's 2nd try, 2026-11-05; the
  1st try, which read Combined Contact Tracking, was reverted). findSearchBarComms_ finds the label in ANY column
  within 15 rows below END (display "Communications", or a formula containing "Communications" that shows an error,
  e.g. #REF! from a blocked spill); refreshSearchBarCommunications_ writes the label as text there (replacing the
  HSTACK formula) and the notes one column right (plain text); OSIS / name switch are read relative to the label
  after a flush. checkSearchBarCommunications (run from the editor) toasts what it found or why it skipped. OSIS = E of the row above, "▶ Name" (Master Table B) when F two rows above is not blank, "No Notes",
  leading "Notes:" stripped (old formula's rules). Notes are built straight from Contact Log (to END) / Backend Event
  Log / Notes with notesTextForOsis_ (also used by updateOsisDeltas_), so it never waits for Combined — EXCEPT a
  search bar lookup (edit below END) and the end of a build read Combined D instead (quickest; Combined is current
  then), unless Settings autoNotes is off (Combined may be stale).
  Runs: edit below END; a Contact Log row edit whose OSIS (or old OSIS) is in the box, right after the ghost typist
  and before the quick update; after Notes / Events edits; end of buildAllDerivedSheets. Writes only when changed.
  restoreSearchBarCommunicationsFormula puts the formula back (doc property SEARCH_BAR_COMMS_MODE = "formula");
  useCodeForSearchBarCommunications switches back.

Workflow:
- Push every change to Test-Version (not main) unless the owner says otherwise for that change. main only gets
  changes when the owner asks for it (e.g. merging Test-Version once it has been tested on a copy of the sheet).

Notes lines (Combined Contact Tracking, and Phone Contacts notes built from it):
- Each line is "𝑫𝒂𝒕𝒆 - {𝐌𝐞𝐭𝐡𝐨𝐝 - 𝐏𝐞𝐫𝐬𝐨𝐧} note [follow-up]" (event: "... - {𝐄𝐯𝐞𝐧𝐭 - [𝐏𝐞𝐫𝐬𝐨𝐧]} |Event|"), no bullet.
  Date = Bold Italic serif (digits stay bold), braces = Mathematical Bold (lingojam BoldTextGenerator) via boldText_ /
  noteLineHeader_ (Core.gs).
- Built in ONE place: contactLogNoteLine_ / eventNoteLine_ / joinNoteLines_ (Build.gs), used by both the full build
  (buildCombinedContactTrackingData_) and the quick update (updateOsisDeltas_).
- Divider lines (owner's trial, 2026-11-05): in Combined Contact Tracking each dated line has a "───" line above and
  under it (one shared line between two; manual Notes block not wrapped). Length = fixed 36 "─" (noteDivider_,
  NOTE_DIVIDER_LENGTH_; owner's pick 2026-11-05, was column D width / 7px). buildTrackingNotesMap_ strips them
  (stripNoteDividers_), so Phone Contacts / Google Contacts notes are exactly as before.
- Order (owner's change, 2026-11-04): manual "𝑵𝒐𝒕𝒆𝒔:" block first, then dated lines NEWEST first (joinNoteLines_ sorts
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
  3. Done: one notes-line formatter, one last-name sort (Master Table + Combined), one Directory sort, one Contact Log
     hide filter + week formula, safeStr_ -> toText_, one Notes/Events single-cell Master Table lookup
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
(notesTextForOsis_, rows grouped by OSIS once) and runs updateOsisDeltas_ for those that differ (> 25 differ: updateContactAndPhoneOnly). Made by
setupAllInstallableTriggers and, once per sheet, by checkVersionOnOpen (ensureChangeTrigger_, doc property
CHANGE_TRIGGER_ADDED, so a 2nd user's open doesn't add a 2nd copy).
Must keep: 15 menu handler names (+5 owner Debug items, + toggleSheetLock), 39 google.script.run endpoints, trigger
handlers (installedOnEdit, installedOnChange, nightlyMasterSync, checkVersionOnOpen, syncToGoogleContacts),
all sheet names, Backend cells,
property/cache keys.
Owner decisions still open (current behavior kept): add shared sync safety checks (50% delete stop, label pageSize)
to owner sync?; owner hides Contact Log col K (11) vs shared H (8) when no site (K looks stale); Montserrat on New
Students (owner build skips it); owner splitFullName_ changes the shared Directory name split - intended?; ghost typist
TextFinder vs in-memory scan (differ only on stray spaces).
Data update reminder (Menus.gs, replaced the old 30-day ATS-only reminder; same checkRawDataFreshness_ name): centered
pop-up when Raw Data A1 / AG1 / AP1 (ATS / Parent Square / NYCSA, only sections ever loaded) is older than setting
"dataReminder" (week default / 2weeks / month / off; in SETTING_KEYS so it is saved to the Backend settings block and
shown in ⚙️ Settings). Checked by checkVersionOnOpen and by installedOnEdit (checkDataReminderWhileOpen_, user cache:
once an hour); once a day (ATS_REMINDER_LAST_SHOWN); not during Input Data Mode. "Go to Input Data Mode" only if
openInputDataMode exists (guarded), else the guide link. Skip -> Remind After 2 Weeks / 1 Month (setting + wait in doc
property DATA_REMINDER_SNOOZE_UNTIL) / Remove Reminders / Just skip for today. Endpoints dataReminderOpenInputMode,
dataReminderChoose. Saving a different choice in Settings clears the wait.
Settings saved but never read: autoOverride, ignoreHideCheckboxes, emailNotifications.
Bugs noticed (not fixed): quick update ignores the Settings notes date range; nightly trigger 11 PM at setup
but 2 AM after Settings save; getExistingMasterTable_ has no header but buildPhoneContactsData_ slice(1)s it
(first student loses Master Table details); EMERGENCY_STOP_SYNC leaves SYNC_EDITED_COUNT; TOGGLE_EMAIL_NOTIFICATIONS
needs two clicks first time; repo manifest lacks script.send_mail so the owner email may never send.
