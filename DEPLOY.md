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
- [ ] A fresh copy of the master template: only the 👋 Welcome tab shows; after 🚨 Initial Setup your tabs are
      back, and the Welcome tab is gone after a reload (or the first build)
- [ ] Reset to Display Mode on the master template: only the 👋 Welcome tab shows (with the 6 permissions explained
      first), also after reopening it. Then Google's permission page on a copy: Select all → Continue

## 5. Going live (after the copy works)
1. Merge Test-Version into main.
2. Update the master template (General manifest, no z_ files).
3. Add a row to the Version Tracker so schools get the update prompt.
