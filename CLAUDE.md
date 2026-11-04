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
  Directory, Phone Contacts, New Students), the notes-line helpers, nightlyMasterSync, build sidebar endpoints.
- Edits.gs: installedOnEdit, ghost typists (Contact Log / Notes / Events), Raw Data paste clean-up, quick updates
  (updateSingleOsisDelta_ / updateOsisDeltas_ / updateContactAndPhoneOnly), Events tables.
- Menus.gs: onOpen menus, setup, Settings, Student Overrides + report, Rollover, Migration, Version Check, links.
- Contacts Sync.gs: Google Contacts sync (syncToGoogleContacts, searchAndUpsertSingleContact), Full Sync endpoints,
  EMERGENCY_STOP_SYNC / RESTART_SYNC.
- Regenerate.gs: reset / fake-data tools (resetToDisplayMode, resetAndLoadFakeData) and Contact Log formatting.
  Owner asked to keep it exactly as is (not moved, not cleaned).
- Z#ContactLog.gs (+ Z#ContactLogSidebar.html): self-contained optional "📇 Contact Log" tool; left as is.
  "Z#" = a temporary file for trying out a new feature: no other code may rely on Z# files (the only link is the
  guarded `if (typeof buildContactLogMenu_ === "function")` line in onOpen, a no-op without the file).

Owner-only file (in the repo, deployed ONLY to the owner's own sheet): z_MINE.gs
- Never copy it into other schools' sheets. It defines ownerHooks_(); shared code calls getOwnerHooks_() (Core.gs),
  which returns those hooks on the owner's sheet and null everywhere else. No shared function is redefined, so file
  load order does not matter. Shared code must work without z_MINE.gs; owner-only behavior (e.g. "Old Master Table" /
  "Old Phone Contacts") goes into a hook in z_MINE.gs, never into the shared files.
- Hooks: Master Table status snapshot, Old Master Table history (Master Table + Directory), Old Phone Contacts merge,
  last-word name split (splitFullName_), tall Phone Contacts layout, quick update that also rewrites Phone Contacts,
  contacts sync (4 phones / 3 emails per row, group list without pageSize, no 50% delete check, summary email),
  Contact Log site column K (shared: H), Montserrat list without New Students, 🐞 Debug menu + toasts.
- Shared Directory rule: draws only from Raw Data (no overrides, no discharged students). On the owner's sheet the
  hook then merges Old Phone Contacts and adds non-Raw-Data students from Old Master Table ONLY when its column F
  (Status) is "MASTER".
- 🐞 Debug menu (owner only): "Reset + Load Fake Data" and "Reset to Display Mode" (type RESET, then Regenerate's own
  confirm; they also delete Old Master Table / Old Phone Contacts) and "Toggle Debug" (doc property OWNER_DEBUG_TOASTS).
  With Debug on, runAction_ toasts ✅ / ⚠️ / ❌ after each edit, open, Build Only, Full Sync, nightly and contacts sync,
  including errors that are normally only logged (console.error / console.warn).

Deploying (manual paste into the Apps Script editor; master template + owner sheet):
- Delete the old files first or the project will not load (duplicate consts/functions): Backend, Build.gs Helpers,
  Build.gs beginning, Constants, Event, LinkResolver, MigrationAssistant, New Copy & On Open, NewStudents, Phone,
  Quick Updates Installed, QuickUpdates, RolloverBackend, Settings Backend, Student Override Report,
  Student Overrides Backend, VersionCheck, ZProduction, z_PhoneOverride, z_PhoneOverride 4+, z_PhoneOverride.gs Add On.
- Then paste Core, Build (replaces the old Build), Edits, Menus, Contacts Sync; owner sheet also z_MINE.
  Regenerate, the other HTML files and the manifest are unchanged. The Contact Log tool is now named
  Z#ContactLog / Z#ContactLogSidebar (delete the old #ContactLog / ContactLogSidebar files, or they define it twice).

Backend tab (see Core.gs):
- One hidden "Backend" tab replaces the old Version, New/Edit Student, Backend_Event_Log and Type List tabs.
  Row 1 = section titles; each old layout sits one row lower. A:E Version & Settings, G:AV New/Edit Student,
  AX:BF Event Log, BH:BJ Send Out list (not used by code; cleared on reset), BL Type List.
- Sections share rows: never deleteRow/insertRow, getDataRange/getLastRow, appendRow or writeSheet_ on it.
  Use the backend* helpers in Core.gs, which only touch one section's columns.
- Old copies migrate automatically the first time getBackendSheet_ runs.

New Students tab (see the NEW STUDENTS TAB section of Build.gs):
- Only Raw Data (ATS, A:AE) counts. Every build compares Raw Data's OSIS list with the previous build's
  (snapshot in document properties, tied to the sheet ID) and logs new / re-entered (from Discharged) students
  (Date Entered, Source = always "ATS", OSIS, Student, Guardian, Site, Class, Language).
  New/Edit Student and Old Master Table are ignored. Only appended to; a row is re-dated when the student re-enters.
- detectAtsChanges_ is the single source for new / discharged students: the New Students tab, the build report
  (sidebar) and the owner's nightly email all use it. Discharged = "Current" in the old Master Table, in Raw Data at
  the last build, and no longer in Raw Data. The snapshot is saved at the end of the build.

Workflow:
- Push every change to Test-Version (not main) unless the owner says otherwise for that change. main only gets
  changes when the owner asks for it (e.g. merging Test-Version once it has been tested on a copy of the sheet).

Notes lines (Combined Contact Tracking, and Phone Contacts notes built from it):
- Each line is "𝑫𝒂𝒕𝒆 - {𝐌𝐞𝐭𝐡𝐨𝐝 - 𝐏𝐞𝐫𝐬𝐨𝐧} note [follow-up]" (event: "... - {𝐄𝐯𝐞𝐧𝐭 - [𝐏𝐞𝐫𝐬𝐨𝐧]} |Event|"), no bullet.
  Date = Bold Italic serif (digits stay bold), braces = Mathematical Bold (lingojam BoldTextGenerator) via boldText_ /
  noteLineHeader_ (Core.gs).
- Built in ONE place: contactLogNoteLine_ / eventNoteLine_ / joinNoteLines_ (Build.gs), used by both the full build
  (buildCombinedContactTrackingData_) and the quick update (updateOsisDeltas_).

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
Must keep: 15 menu handler names (+3 owner Debug items), 39 google.script.run endpoints, trigger handlers
(installedOnEdit, nightlyMasterSync, checkVersionOnOpen, syncToGoogleContacts), all sheet names, Backend cells,
property/cache keys.
Owner decisions still open (current behavior kept): add shared sync safety checks (50% delete stop, label pageSize)
to owner sync?; owner hides Contact Log col K (11) vs shared H (8) when no site (K looks stale); Montserrat on New
Students (owner build skips it); owner splitFullName_ changes the shared Directory name split - intended?; ghost typist
TextFinder vs in-memory scan (differ only on stray spaces).
Settings saved but never read: autoOverride, ignoreHideCheckboxes, emailNotifications.
Bugs noticed (not fixed): quick update ignores the Settings notes date range; nightly trigger 11 PM at setup
but 2 AM after Settings save; getExistingMasterTable_ has no header but buildPhoneContactsData_ slice(1)s it
(first student loses Master Table details); EMERGENCY_STOP_SYNC leaves SYNC_EDITED_COUNT; TOGGLE_EMAIL_NOTIFICATIONS
needs two clicks first time; repo manifest lacks script.send_mail so the owner email may never send.
