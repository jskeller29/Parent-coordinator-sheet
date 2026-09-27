This is a Google Apps Script project (a Google Sheets add-on for student/roster management).
- .gs files are Apps Script server-side code (JavaScript that runs on Google's servers)
- .html files are dialog/sidebar UIs served via HtmlService
- Frontend HTML talks to backend .gs functions via google.script.run
- There is no build step and no modules — all .gs files share one global scope

Owner-only files (NOT in this repo):
- z_PhoneOverride, z_PhoneOverride 4+, z_PhoneOverride.gs Add On live only in the owner's own sheet.
  The "z_" prefix makes them load last so they can redefine shared functions for that one sheet.
- Shared code must not depend on them. Owner-only behavior (e.g. pulling from "Old Master Table" /
  "Old Phone Contacts") belongs in those files, never in the shared .gs files.
- Shared Directory rule: draws only from Raw Data (no overrides, no discharged students).
  The owner's sheet may add Old Master Table / Old Phone Contacts students, but still excludes discharged.
