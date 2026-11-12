# Offline tests

These run on a computer with Node.js (18 or newer). Nothing here talks to Google, and
nothing in this folder is pasted into Apps Script.

```bash
cd tests
npm install      # once (installs acorn for check.js)
npm test         # runs everything below
```

| File | What it checks |
|---|---|
| `check.js` | Every `.gs` file parses; all of them together compile as one script (Apps Script shares one global scope); no name is declared twice; every HTML file the code opens exists (and every HTML file is used); every `google.script.run` call and menu item points at a real function. `node check.js <older copy of the repo>` also lists which functions changed. |
| `test_matching.js` | Your Lists matching: typos, other name order, short names, initials, siblings, guardians, "not found", plus a speed check (500 rows x 1,500 students). |
| `test_fixes.js` | Notes date range in the quick update, Sheet Lock ranges around the lookup box, Follow-up values (Migration + Contact Log tool), one nightly job per sheet, the "last night's rebuild didn't finish" message, the build lock, the quick-rebuild header row, ghost typists ignoring stray spaces. |
| `test_sync.js` | The phone sync against a fake Google Contacts: normal sync, a failed save (nothing removed), Google's limit (15-minute pause, then resume), "more than half" (blocked when no one is there, removed after Yes), a long build before the sync, the owner's tall rows. |
| `test_testdata.js` | The files in `test-data/` go through Input Data Mode (header found, right export, the broken file refused) and Your Lists (expected matched / best guesses / not found per list). |
| `render_sidebars.js` | Optional: screenshots of the Input Data Mode and Your Lists sidebars with sample data (needs Playwright + Chromium). Output in `tests/out/`. |

`lib/load.js` loads the project's `.gs` files into one Node context with small fake Apps
Script services (properties, cache, lock). Add a test by loading the project and calling
the function you changed:

```js
const { loadProject } = require('./lib/load');
const g = loadProject();              // a school's sheet; { owner: true } also loads z_MINE.gs
console.log(g.typistKey_("  Ann  Smith "));
```

These tests cover plain logic only. Anything that reads or writes a sheet, opens a dialog or
calls Google still has to be tried on a copy of the real sheet.
