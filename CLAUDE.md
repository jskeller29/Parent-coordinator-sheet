This is a Google Apps Script project (a Google Sheets add-on for student/roster management).
- .gs files are Apps Script server-side code (JavaScript that runs on Google's servers)
- .html files are dialog/sidebar UIs served via HtmlService
- Frontend HTML talks to backend .gs functions via google.script.run
- There is no build step and no modules — all .gs files share one global scope

Owner-only files (in the repo, but deployed ONLY to the owner's own sheet):
- z_PhoneOverride, z_PhoneOverride 4+, z_PhoneOverride.gs Add On are never copied into other schools' sheets.
  The "z_" prefix makes them load last so they can redefine shared functions for that one sheet.
- Shared code must not depend on them. Owner-only behavior (e.g. pulling from "Old Master Table" /
  "Old Phone Contacts") belongs in those files, never in the shared .gs files.
- Shared Directory rule: draws only from Raw Data (no overrides, no discharged students).
  The owner's sheet (z_PhoneOverride 4+) overrides the build: Raw Data base + Old Phone Contacts merge, and
  adds non-Raw-Data students from Old Master Table ONLY when its column F (Status) is "MASTER".

Backend tab (see Backend.gs):
- One hidden "Backend" tab replaces the old Version, New/Edit Student, Backend_Event_Log and Type List tabs.
  Row 1 = section titles; each old layout sits one row lower. A:E Version & Settings, G:AV New/Edit Student,
  AX:BF Event Log, BH:BJ Send Out list (not used by code; cleared on reset), BL Type List.
- Sections share rows: never deleteRow/insertRow, getDataRange/getLastRow, appendRow or writeSheet_ on it.
  Use the backend* helpers in Backend.gs, which only touch one section's columns.
- Old copies migrate automatically the first time getBackendSheet_ runs.

New Students tab (see NewStudents.gs):
- Only Raw Data (ATS, A:AE) counts. Every build compares Raw Data's OSIS list with the previous build's
  (snapshot in document properties, tied to the sheet ID) and logs new / re-entered (from Discharged) students
  (Date Entered, Source = always "ATS", OSIS, Student, Guardian, Site, Class, Language).
  New/Edit Student and Old Master Table are ignored. Only appended to; a row is re-dated when the student re-enters.
- detectAtsChanges_ (NewStudents.gs) is the single source for new / discharged students: the New Students tab,
  the build report (sidebar) and the owner's nightly email all use it. Discharged = "Current" in the old
  Master Table, in Raw Data at the last build, and no longer in Raw Data. The snapshot is saved at the end of the build.

Workflow:
- The owner wants every finished change merged into main: after pushing the working branch, also push it to main.

Notes lines (Combined Contact Tracking, and Phone Contacts notes built from it):
- Each line is "𝑫𝒂𝒕𝒆 - {𝐌𝐞𝐭𝐡𝐨𝐝 - 𝐏𝐞𝐫𝐬𝐨𝐧} note [follow-up]" (event: "... - {𝐄𝐯𝐞𝐧𝐭 - [𝐏𝐞𝐫𝐬𝐨𝐧]} |Event|"), no bullet.
  Date = Bold Italic serif (digits stay bold), braces = Mathematical Bold (lingojam BoldTextGenerator) via boldText_ / noteLineHeader_ (Build.gs Helpers.gs).
- Built in TWO places that must match: buildCombinedContactTrackingData_ (Build.gs) and updateSingleOsisDelta_ (QuickUpdates.gs).

POSSIBLE FUTURE PLAN (code-structure cleanup; reviewed 2026-11-02, not started, needs owner go-ahead)
Until the owner says go: keep the CURRENT file layout/schema for all changes (do not restructure).
Goal: same exact functionality, cleaner files. Do it in stages, each checked against the old code:
  1. Move only: functions copied word-for-word into new files, delete dead code, drop always-true
     `typeof fn === "function"` checks (107 of them; they only hide typos).
  2. Owner hook points: shared code calls an owner version only if it exists
     (`if (typeof ownerAdjustDirectory_ === "function") ...`), so no function is defined twice and load
     order stops mattering. Removes the 180-line copy of buildAllDerivedSheets in z_PhoneOverride 4+ and the
     ~200-line copy of syncToGoogleContacts in z_PhoneOverride.
  3. Merge duplicates: one notes-line formatter (build + quick update), one Master Table "ghost typist"
     lookup (6 versions now), one Directory site/class lookup (3), one Contact Log filter / week formula,
     Raw Data header list, last-name sort, safeStr_ -> toText_, one settings table (now 4 lists),
     constants for sheet names (123 literals; "Raw Data" vs "RAW Data") and Raw Data / Contact Log columns.
  4. Optional behavior fixes (list below), one at a time.
Must keep: 15 menu handler names, 39 google.script.run endpoints, trigger handlers (installedOnEdit,
nightlyMasterSync, checkVersionOnOpen, syncToGoogleContacts), all sheet names, Backend cells, property/cache keys.
Verify with an offline old-vs-new build diff on fake data (shared + owner) + a click-through checklist.
Deploy is manual paste into the Apps Script editor (master template + owner sheet); clasp is an option.
Proposed files: Config, Util, Backend, App Menu & Setup, Settings, Build, Build Data, Build Sheets,
New Students, Quick Updates, Edit Triggers (incl. ghost typist), Contact Log Sheet, Raw Data Sheet, Events,
Student Overrides, Contact Log Tool, Contacts Sync, Version & Links, Migration & Rollover, Dev Tools;
owner: z_Owner Build, z_Owner Contacts, z_Owner Dev Tools. HTML files unchanged.
Owner decisions still open: add shared sync safety checks (50% delete stop, label pageSize) to owner sync?;
owner hides Contact Log col K (11) vs shared H (8) when no site (K looks stale); Montserrat on New Students
(owner build skips it); owner splitFullName_ override changes the shared Directory name split - intended?;
ghost typist TextFinder vs in-memory scan (differ only on stray spaces).
Dead code: isTruthy_, compareText_, findEndRowInColumnG_, getDirectoryLabelResourceName_, getGuideLink,
runInitialBulkContactSync, showSidebar (opens missing 'Sidebar' file); settings saved but never read:
autoOverride, ignoreHideCheckboxes, emailNotifications.
Bugs noticed (not fixed): quick update ignores the Settings notes date range; nightly trigger 11 PM at setup
but 2 AM after Settings save; getExistingMasterTable_ has no header but buildPhoneContactsData_ slice(1)s it
(first student loses Master Table details); EMERGENCY_STOP_SYNC leaves SYNC_EDITED_COUNT; TOGGLE_EMAIL_NOTIFICATIONS
needs two clicks first time; repo manifest lacks script.send_mail so the owner email may never send.
