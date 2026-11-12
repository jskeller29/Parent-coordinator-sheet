# PC Tracker test data

Fake files for trying everything out on a **copy** of the sheet. All of it is made up:
the same 18 students and parents as 🐞 Debug → 🧪 Reset + Load Fake Data (Simpsons,
Addams, Belchers, …), so these files also work on a sheet that already has the fake data.

**Download:** `PC Tracker test data.zip` has every file. Each file is there as **.csv**
and **.xlsx** (Excel). Use either one.

## Input Data Mode (🚀 App Menu → 📥 Input Data Mode)
Upload the file at its step (📂 Upload file), or open it, press Ctrl+A, copy, and paste
on the step's tab.

| File | Step | What you should see |
|---|---|---|
| `1 - Student Data (ATS)` | 1 Student Data | 18 students moved to Raw Data. The Excel file has 2 title rows above the column names, which are skipped. |
| `1b - Student Data (ATS) - next month (2 new, 1 left)` | 1 Student Data, later | Load it after a first build, then build again: the build report shows 2 new students (Steve Rogers, Lucia Lopez) and 1 discharged (Tony Stark). |
| `1c - Student Data (ATS) - BROKEN numbers cut off` (.csv only) | 1 Student Data | Refused: "numbers look cut off by Excel" for Student ID and both phone columns (they read like 1.23E+08). "Use it anyway" is offered. |
| `2 - Parent Square` | 2 Parent Square | 28 rows. Upload it at step 1 instead and it is refused ("This looks like Parent Square data"). |
| `3 - NYCSA (Students with Non-NYCSA Account)` | 3 NYCSA | 25 rows. |
| `4 - Site and Class` | 4 Site / Class | Paste the three columns (with the header) over A1 of the Site / Class tab, then Save. 6 sites, 12 classes; Y12 and Y50 have no teacher. |

## Your Lists (step 5 of Input Data Mode → Add another list)
Name the list, paste the whole file (headers too) in the yellow box, then go through
Columns → Variables → Match. These counts are against the 18 ATS students above.

| File | Each row is a | What you should see on Match |
|---|---|---|
| `5a - Your Lists - Bus Routes` | student | The Excel file has 3 title rows above the headers (left out). **10 matched** (OSIS, "SIMPSON, LISA", "MORALES MILES" in the other order). **6 best guesses:** Violet Incredable 96% (typo), Lou Belcher 96% (Louise), Pugsly Adams 89%, Diana P. 84%, Bruce B. 84% (Banner, not Wayne), Mirabel Madrigal 100% with a wrong OSIS. **2 not found:** Steve Rogers (found once you load `1b`), Nobody Here. Peter Parker has no route: try "If the cell is blank → Write this: No bus". |
| `5b - Your Lists - PTA Members` | guardian | Guessed as a guardian list (parent first + last columns). **11 matched**, incl. typos the phone or email settles (Marj → Marge, Adams → Addams, Pennywise → Pennyworth). **1 best guess:** Martha K. → Martha Kent 84%. **1 not found:** Natasha Romanoff. |
| `5c - Your Lists - Field Trip Forms` | student | First and last name in two columns. **9 matched.** **1 not found:** Tina Belcher, with Louise / Gene Belcher as the closest (on a sheet with 🧪 Reset + Load Fake Data, Tina is a New/Edit Student and matches). Wednesday and Bruce Banner have no form: try "If the cell is blank → Not returned". |
| `5d - Your Lists - After School Sign-up` | student | "LAST, FIRST" names. **6 matched.** **2 best guesses:** STARK, ANTHONY → Tony Stark 98% (nickname), BELCHER, EUGENE → Gene Belcher 86%. |

Then try: **Accept the high-confidence** (90%+), pick someone with "Pick someone else",
Save, finish Input Data Mode, **Build**, and look at the new Student List / Parents
Divided columns and the "Additional Information" lines in Phone Contacts.

## Remaking the files
`python3 make_test_data.py` (needs `openpyxl`) rebuilds every file and the zip from the fake
data in `Regenerate.gs`. `tests/test_testdata.js` checks the counts above.
