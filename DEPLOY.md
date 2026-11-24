# Deploy checklist (Test-Version)

The Apps Script editor has no "update": you make its file list match this page by hand.
Do it on a **copy** first. Every name below is exactly what the editor shows
(scripts end in `.gs`, pages in `.html`).

## 1. Delete old files (if they are there)
A file from an older version left in the project = the project won't load
(the same function twice).

- Scripts from before the cleanup: Backend, Build.gs Helpers, Build.gs beginning, Constants,
  Event, LinkResolver, MigrationAssistant, New Copy & On Open, NewStudents, Phone,
  Quick Updates Installed, QuickUpdates, RolloverBackend, Settings Backend,
  Student Override Report, Student Overrides Backend, VersionCheck, ZProduction,
  z_PhoneOverride, z_PhoneOverride 4+, z_PhoneOverride.gs Add On, #ContactLog
- Old trial files: Z#InputDataMode, Z#OwnData, Z#Guide, Z#ContactLogDebug
- Old pages: ContactLogSidebar, Z#InputDataSidebar, Z#OwnDataSidebar, Z#GuideSidebar,
  Z#ContactLogDebugSidebar, ContactLogDebugSidebar

## 2. The files each copy should have

| | Your own sheet | Master template + schools |
|---|---|---|
| **Scripts** | Core, Build, Edits, Menus, Input Data, Contacts Sync, Sheet Lock, Regenerate, Z#ContactLog | same |
| **Your scripts** | z_MINE, z_ContactLogDebug, z_InputDataDebug, z_NewCopy | **none** (never copy these) |
| **Pages** | Styles, SettingsDialog, ProgressSidebar, NewStudentDialog, ManageOverridesDialog, RolloverWizard, MigrationUI, VersionDialog, InputDataSidebar, YourListsSidebar, GettingStartedSidebar, FindFamily, Z#ContactLogSidebar | same |
| **Your pages** | z_ContactLogDebugSidebar | **none** |
| **Manifest** (`appsscript.json`) | contents of `App Scripts (Mine).Json` | contents of `App Scripts (General).Json` |

To see `appsscript.json`: Project Settings (gear) → tick
**Show "appsscript.json" manifest file in editor**.

Paste each file's full contents from Test-Version (replace everything in it), and create
the ones that don't exist yet with the same name (script: + → Script; page: + → HTML).

**Making a new copy of your sheet** (for a school, the master template or a test): open the
copy and a pop-up lists what to change in it (z_NewCopy.gs, your sheet only). The same list
with tick boxes is in 🐞 Debug → 📋 New Copy Checklist.

**This update (2026-11-23, your 10 ideas):** re-paste the scripts Core, Build, Edits, Menus, Input Data,
Contacts Sync, Regenerate and the pages InputDataSidebar, ManageOverridesDialog, NewStudentDialog,
SettingsDialog, ProgressSidebar, YourListsSidebar; **add** the page **FindFamily** (new). Your sheet: also
z_MINE. No new permissions (same manifests). Then do step 10 below once.
The testers' fixes from the same day, the Student List's new Home Language column (Build) and the orange
Follow-up color (11/24: Core, Edits, Build, Regenerate) are in these same files, so there is nothing more to paste. If you
already pasted the 10 ideas before them, re-paste Core, Build, Menus, Input Data, Contacts Sync and the pages
FindFamily, InputDataSidebar, ManageOverridesDialog, NewStudentDialog, ProgressSidebar, YourListsSidebar.

## 3. After pasting
1. Save, then reload the spreadsheet.
2. Run any menu item and **approve the permissions** if asked. Until you do, your nightly
   and edit triggers stop with "authorization required".
3. **🚀 App Menu → 🏗️ Build Sheets Only** once. It fills the Backend student copy
   (BX:CD), writes the Student List from row 1, and refreshes the notes.
4. Open **⚙️ Settings** and click Save once (sets the 10 PM nightly job and remembers it).
5. Your sheet only: **🐞 Debug → 🔧 Fix Dropdown Criteria** once (dropdowns made by hand).
6. Your sheet only: if Contact Log column **K** (Method 2) is hidden, unhide it.
7. Master template only: the Student List header on row 1, frozen, as on your sheet.
8. Your sheet only: **🐞 Debug → 🎨 Apply Sheet Style** once (colors + fonts on every tab; needs the
   `App Scripts (Mine).Json` manifest pasted, which turns on Google's Sheets service: no new permission).
   Its pop-up ends with what a script can't do (the Contact Log's dropdown chip colors, its widths).
   📦 Make Release Copies then gives the Example and Blank copies the same look.
9. Master template and your sheet: the **Person Spoke With** dropdown, once, by hand (a script can't
   change a Table's dropdown column). Contact Log → Person Spoke With column → Data validation →
   **Dropdown (from a range)** `='Backend'!$CF2:$CO2` (a 2 on BOTH ends, no `$` before them), Advanced options →
   **Show a warning** (other people can still be typed). Then **🏗️ Build Sheets Only** once more.
   The build in step 3 already wrote the list on the hidden Backend tab (CF2: row n = Contact Log row n).
   The next build sees the dropdown reads the Backend and takes the old formula out of Contact Log **R2**:
   that is what lets **Delete table row** work again, and deleting / inserting Contact Log rows can't
   break the list any more (release check 4.3 / 4.4). The same build takes out any checkboxes after column P
   and deletes the empty columns after it, so the Contact Log ends at **P** again (2026-11-15: older builds
   added Q:AA with copies of P's checkboxes). **Changed 2026-11-14**: it was `='Contact Log'!$R2:$AA2`;
   a sheet still set to that keeps working the old way (the build keeps R2 and R:AA; only checkboxes after P
   go) until you change it. Leave R1 empty. Schools' sheets: the same rule by hand after their next build, or
   they get it with next year's copy of the template; until then their build adds no columns after P.

10. Master template and your sheet: the **Method 2** dropdown, once, by hand (2026-11-23; a script can't change a
   Table's dropdown column). First **🏗️ Build Sheets Only** once: it writes the list on the hidden Backend tab
   (column **CQ**: your fixed choices first, then every Method 2 already typed). Then Contact Log → **Method 2**
   column (K) → Data validation → **Dropdown (from a range)** `='Backend'!$CQ$3:$CQ`, Advanced options →
   **Show a warning** (NOT "Reject the input": a new choice can only join the list after it is typed once).
   To change the fixed choices (now Call, Email, Text, Parents Square, In School), edit the formula in
   **Backend CQ3** on the sheet: the build keeps a formula there that reads Contact Log K.
   Nothing to do by hand for the **light red Method** color: every build adds that rule when it is missing
   (and the 10 PM run draws it with the others).

## 4. Test it on copies: the release checklist
Everything for a new version, step by step, with what you should see, in four parts: before you start (the rules,
and this version on your copies), test with fake data (what a new user sees, what's new, day to day: the main test),
a short check with your real data, then the release steps. Do the 61 **essentials** before any school gets a
version; the 29 checks marked Optional can wait. Three ways to use the same list:
- **The web checklist** (best), **round 3** (2026-11-23, your 10 ideas): https://claude.ai/artifact/68r88BXD58pjvejrLmYLCQ
  (only you can open it). It marks every new check "New" and every changed one "Fixed, re-test", shows what you wrote
  in round 2 on each check, and has a third view, **Try to break it**, where you can also add tests of your own.
  (Round 2's page, https://claude.ai/artifact/VgrJLTu3a2Vaixz8JLQHYc, and round 1's,
  https://claude.ai/artifact/SdhtaNhpTbvt6NMCXmcLa1, keep their answers.) Mark each
  check Works, Problem or Skip and write what you saw. It saves as you go, on any device. When you're done, tell
  Claude "release check done": Claude reads your answers there. It can also copy the report or save it as a PDF.
- **On paper:** `docs/release-check/PC_Tracker_Release_Check.pdf` (also a button on the web checklist).
- **In GitHub:** Issues → New issue → **Release check** (offered once Test-Version is merged into main; until then ask
  Claude to open one). Ticks save by themselves; write failures as comments.

Read its **Before you start** part first: phone sync and "bring over last year's tracker" on a test copy can touch
your real Google Contacts and pause your live sheet. When you're done, move the test copies to the trash: a tracker
in the trash skips its 10 PM rebuild by itself. The checks come from `docs/release-check/checks.js`
(`node docs/release-check/build.js` remakes all three).

## 5. Going live (after the copy works)
0. Pilot first: 1 or 2 friendly schools use it for a week or two (the checklist's last box).
1. Merge Test-Version into main.
2. Update the master template (General manifest, no z_ files).
3. Add a row to the Version Tracker so schools get the update prompt: A = the version date, D "Blank Copy" = the
   template's /copy link, E "Example" / F "Document" = the example tracker / User Guide when they changed (blank =
   the newest earlier row's link is kept). A NEW User Guide: put its F link in BEFORE 🐞 Debug → 📦 Make Release
   Copies (the Blank copy's Welcome tab gets the guide link then); D and E come from the pop-up, after.
