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
