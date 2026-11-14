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
| **Pages** | Styles, SettingsDialog, ProgressSidebar, NewStudentDialog, ManageOverridesDialog, RolloverWizard, MigrationUI, VersionDialog, InputDataSidebar, YourListsSidebar, GettingStartedSidebar, Z#ContactLogSidebar | same |
| **Your pages** | z_ContactLogDebugSidebar | **none** |
| **Manifest** (`appsscript.json`) | contents of `App Scripts (Mine).Json` | contents of `App Scripts (General).Json` |

To see `appsscript.json`: Project Settings (gear) → tick
**Show "appsscript.json" manifest file in editor**.

Paste each file's full contents from Test-Version (replace everything in it), and create
the ones that don't exist yet with the same name (script: + → Script; page: + → HTML).

**Making a new copy of your sheet** (for a school, the master template or a test): open the
copy and a pop-up lists what to change in it (z_NewCopy.gs, your sheet only). The same list
with tick boxes is in 🐞 Debug → 📋 New Copy Checklist.

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
   **Dropdown (from a range)** `='Contact Log'!$R2:$AA2` (a 2 on BOTH ends, no `$` before them), Advanced options →
   **Show a warning** (other people can still be typed). The build in step 3 already wrote the formula
   in Contact Log **R2** and hid columns R:AA; leave R1 empty and Q blank. Schools' sheets: the same rule
   by hand after their next build, or they get it with next year's copy of the template.

## 4. Test it on copies: the release checklist
Everything for a new version, step by step, with what you should see, in four parts: set up the test copies,
test with fake data (the main test), a short check with your real data, then the release steps. Do the 48
**essentials** before any school gets a version; the 27 checks marked Optional can wait. Three ways to use the same
list:
- **The web checklist** (best): https://claude.ai/artifact/SdhtaNhpTbvt6NMCXmcLa1 (only you can open it). Mark each
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
