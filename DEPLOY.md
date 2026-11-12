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

## 4. Quick test on the copy
- [ ] 📥 Input Data Mode: all 5 steps with a real export (paste and upload). Paste a wrong export once: the tab
      empties by itself. Step 5 builds the sheets first; the end screen only offers Build if something changed
- [ ] Your Lists: upload a CSV and an Excel list as well as pasting one; after closing Input Data Mode the 📥 list
      tabs are hidden
- [ ] Your Lists step 4 View: try One line + bold names, Build, and look at a phone contact's Additional Information;
      a list saved before this version reads exactly as it did
- [ ] Input Data Mode: "Finish or exit" → "Save, exit and build my sheets" in the middle of the steps
- [ ] Your Lists: paste a list with a few misspelled names; check the % and the picks
- [ ] Contact Log: type a parent name, paste a few rows, delete a row; the search bar notes follow
- [ ] 🧱 Full Sync while you watch (the "more than half" question, if it comes, is asked at the start)
- [ ] On a fresh template copy, before its first build: 🚀 App Menu → 🔄 Import Old Data from last year's tracker
      (Follow-ups translated, lists copied; the menu item is gone once a sheet is built)
- [ ] The lookup box: type an OSIS below END (no "are you sure?" warning)
- [ ] A fresh copy of the master template (no z_ / Z# files): only the 👋 Welcome tab shows. 🚨 Initial Setup opens
      "Setting up your tracker" at once; during and after it ONLY the Welcome tab shows (now "You're set up"). Yes,
      start fresh → Start Input Data Mode: only the yellow Student Data tab shows, with "PASTE HERE" and the grey "Closed
      the side panel?" line. Close the sidebar, reload: it comes back on its own. Exit before building: back on the
      Welcome tab. After the first build: Contact Log, Combined, PCAR, Directory appear and the Welcome tab goes.
      Getting Started → Finish setup: a pop-up, and it no longer opens by itself
- [ ] Mistakes a new user makes: 🏗️ Build Sheets Only before loading data (a note, nothing built); type on the
      Welcome tab after setup ("Heads up" warning); Skip on Student Data (asks first); empty Site / Class + Save
      (your saved sites stay); ⚙️ Settings: Save is greyed out until "Loading your settings…" goes away
- [ ] Bring over last year's tracker: pasting this tracker's own link is refused; the old events come over with their
      delete box and dropdowns; the archive PDF from 🔄 End of Year Rollover is named for the year that ends
- [ ] Before Initial Setup: typing in a copy WITHOUT z_ files shows "🔒 Action Blocked" and undoes it; in a copy that
      still has a z_ file, typing works
- [ ] Reset to Display Mode on the master template: only the 👋 Welcome tab shows (with the 6 permissions explained
      first), also after reopening it. Then Google's permission page on a copy: Select all → Continue
- [ ] Your sheet: 🐞 Debug → 📦 Make Release Copies: two copies "Parent Coordinator Sheet - (EXAMPLE / BLANK) (date)"
      with a link each and one version date (that date goes in the Version Tracker's column A); 🔗 My Links
      lists every link from the 📌 YOUR LINKS box at the top of Menus
- [ ] Backend New/Edit Student (G:AV): after one build, the headers are on row 2 and students start on row 3, on your
      sheet and in each copy
- [ ] One template: on a copy of the (BLANK) file, Initial Setup asks "Is this your first PC Tracker?". Start fresh =
      the four steps; Bring over last year's tracker = the wizard, which lists your old trackers (pick one, Bring it over)
- [ ] Template folder: only the newest "(BLANK)" file in it (move old "(MIGRATE)" files to a subfolder); the Version
      Tracker's column D = that (BLANK) file's /copy link
- [ ] Version Tracker headers: D "Blank Copy", E "Example", F "Document" (clear the old values in E first); 🔗 My Links
      shows "Version Tracker, column D/E/F" with each link, and "See an example tracker" appears in 📖 Open User Guide
- [ ] Make Release Copies: the Blank copy's 👋 Welcome tab has "See an example tracker" opening the Example copy made
      in the same run (no Initial Setup on the Blank copy, nothing to redo there)

- [ ] 🎨 Apply Sheet Style: Montserrat on every tab; Notes / Send Out brown like the Contact Log; Student List,
      Parents Divided, Directory, Combined navy; PCAR's boxes navy + brown with yellow date cells; no "Column 9" on
      the Student List; "Raw Data" tab name; Raw Data's "DATE UPDATED" titles visible at the left. Then: a new event
      (brown title row), a Type typed in the Contact Log keeps its color after the 10 PM run, a reset leaves the
      Student List at A:H, a Your Lists list adds its column back with a navy header

## 5. Going live (after the copy works)
1. Merge Test-Version into main.
2. Update the master template (General manifest, no z_ files).
3. Add a row to the Version Tracker so schools get the update prompt: A = the version date, D "Blank Copy" = the
   template's /copy link, E "Example" / F "Document" = the example tracker / User Guide when they changed (blank =
   the newest earlier row's link is kept). A NEW User Guide: put its F link in BEFORE 🐞 Debug → 📦 Make Release
   Copies (the Blank copy's Welcome tab gets the guide link then); D and E come from the pop-up, after.
