---
name: Release check
about: Tick through every check on a copy before a new version goes to schools
title: "Release check: (version date)"
---

<!--
Made by docs/release-check/build.js from docs/release-check/checks.js: change the checks THERE and
run the build (this file, the web checklist and the printable PDF are made from it). One issue per release;
ticks save by themselves; anything that fails goes in a comment.
-->

The easier way to do this: the web checklist (steps, notes, a report Claude can read). Ask Claude for its link.
How to put the new version on a copy: [DEPLOY.md](https://github.com/jskeller29/Parent-coordinator-sheet/blob/Test-Version/DEPLOY.md), sections 1 to 3.
Test files: [test-data](https://github.com/jskeller29/Parent-coordinator-sheet/tree/Test-Version/test-data) (its README says what each one should show).
Something fails? Write it as a comment on this issue: the step, what you saw, a screenshot, and any red error from Extensions → Apps Script → **Executions**.
Checks marked _(optional)_ can wait. Do the other 48 (the essentials) before any school gets this version.

## Part 1: Set up for testing
Make the test copies. Your live sheet is never touched.

### Before you start
_Four rules for every test copy._
- [ ] **0.1 One Google account in this browser.** Use a browser window signed in to only your school Google account (a private window where you sign in once works too). **You should see:** Only one account shows when you click your picture at the top right of Google Sheets. _Why: Several accounts in one browser can make the 🚀 App Menu slow or missing._
- [ ] **0.2 Phone sync stays off on test copies.** Never turn on ⚙️ Settings → Contacts & sync → "Sync contacts to my phone every night" on a test copy. Stage 6 is the one exception, from a second Google account. A new copy starts with it off, so there is nothing to turn off: just look once after each copy's Initial Setup. **You should see:** The switch is off when you open ⚙️ Settings on a test copy. _Why: A test copy with it on would put its families into the "School Directory" label of YOUR Google Contacts every night at 10 PM. Before this version a copy took the switch from the sheet it was copied from._
- [ ] **0.3 Never bring over from your live sheet.** When a test asks you to bring over last year's tracker, pick a test copy, never your live sheet. **You should see:** Your live sheet never asks "Keep this tracker paused?". _Why: Bringing a tracker over pauses that tracker's 10 PM rebuild. If it happens to your live sheet: open it and answer No._
- [ ] **0.4 The test files.** Download "PC Tracker test data.zip" from the test-data folder on GitHub and unzip it. **You should see:** A folder with 1 - Student Data (ATS), 2 - Parent Square, 3 - NYCSA, 4 - Site and Class and four 5 - Your Lists files.

### 1. Make the test copies
_A TEST master (a copy of your sheet with the new version), then a fake-data copy and a test template made from it._
- [ ] **1.1 The TEST master.** Open your live tracker → File → Make a copy. Name it "TEST master" and the date. Open it. When it asks "Is this your MAIN sheet?", answer No, and close the New Copy Checklist. **You should see:** It has all your data and your 🐞 Debug menu.
- [ ] **1.2 The new version on the TEST master.** In the TEST master: Extensions → Apps Script. Follow DEPLOY.md sections 1 and 2: delete the old files it lists, paste every file from Test-Version, keep your z_ files, paste App Scripts (Mine).Json into appsscript.json. Save. **You should see:** Saving shows no red error. The file list matches the table in DEPLOY.md section 2.
- [ ] **1.3 Setting up the TEST master.** Reload. 🚀 App Menu → 🚨 Initial Setup (a copy asks again) → approve the permissions → close the question window. Open ⚙️ Settings and look: phone sync is off (a new copy starts with it off). Close it. Then 🚀 App Menu → 🏗️ Build Sheets Only. 🐞 Debug → 🔧 Fix Dropdown Criteria, then 🐞 Debug → 🎨 Apply Sheet Style. Set the Person Spoke With dropdown by hand (DEPLOY.md section 3, step 9). **You should see:** Phone sync was already off, without you changing it. The build sidebar ticks every step. Apply Sheet Style ends with its list of things to do by hand.
- [ ] **1.4 The fake-data copy.** On the TEST master: 🐞 Debug → 🧪 Copy Sheet, Then Load Fake Data on the Copy. Open the link it gives. Answer No to "Is this your MAIN sheet?". Extensions → Apps Script: delete every file starting with z_ and paste App Scripts (General).Json into appsscript.json. Save. Now it acts like a school's sheet. Reload → 🚀 App Menu → 🚨 Initial Setup → approve → close the question window. Look in ⚙️ Settings: phone sync is off. **You should see:** The fake students (Simpsons, Addams, Belchers…), 26 Contact Log rows, 3 events and 5 Notes. No 🐞 Debug menu.
- [ ] **1.5 The test template.** On the TEST master: 🐞 Debug → 📄 Copy Sheet, Then Reset the Copy. Open the link it gives. Answer No to the MAIN question, delete every z_ file, paste App Scripts (General).Json. Save. Don't run Initial Setup on it. For every new-school test, make a copy of it (File → Make a copy): that copy is what a school gets. **You should see:** Only the 👋 Welcome tab shows. Extensions → Apps Script shows no file starting with z_.

## Part 2: Test with fake data
The main test. Fake students and fake exports, so nothing real can be hurt and every answer is known.

### 2. A brand-new school
_A copy of the test template, and the test files._
- [ ] **2.1 Opening a fresh copy.** Open the test template → File → Make a copy → open the copy. Count the seconds until 🚀 App Menu shows at the top. **You should see:** Only the 👋 Welcome tab shows, starting with the 6 permissions Google will ask for. A note at the bottom right points to 🚀 App Menu → 🚨 Initial Setup. The menu shows. The very first open can take up to a minute; if it still isn't there, reload. (Seconds until the menu showed: ___)
- [ ] **2.2 Typing is blocked before setup.** _(optional)_ Type anything in a cell of the Welcome tab and press Enter. **You should see:** "🔒 Action Blocked", and your typing is undone.
- [ ] **2.3 Initial Setup and Google's screens.** 🚀 App Menu → 🚨 Initial Setup (Run Once). Authorization required → OK → pick your account. "Google hasn't verified this app" → Advanced → Go to … (unsafe). Select all → Continue. Count the seconds until the question shows. **You should see:** "Setting up your tracker" opens right away, with a spinner. Only the Welcome tab shows during and after setup; it now says "You're set up". The 🚀 App Menu changes to the getting-started menu. (Seconds until the question showed: ___)
- [ ] **2.4 "Is this your first PC Tracker?".** Click "Yes, start fresh". **You should see:** "How to download your data" and "Watch the video" are there at once; "See an example tracker" can come a few seconds later. The ✅ Getting Started checklist opens on the right.
- [ ] **2.5 Starting Input Data Mode.** Click "Start Input Data Mode". **You should see:** Only the yellow "Student Data" tab shows, with PASTE HERE in A1 and a grey "Closed the side panel?" line. The side panel says "Data 1 of 5: Student Data (ATS / Insight)" and which file to pick (its name starts with general_).
- [ ] **2.6 The side panel comes back.** _(optional)_ Close the side panel with its X, then reload the page. **You should see:** Input Data Mode comes back by itself, on the same step.
- [ ] **2.7 Leaving before any build.** _(optional)_ In the side panel: Finish or exit → Exit without saving. Then 🚀 App Menu → 📥 Input Data Mode to start it again. **You should see:** After the exit you are back on the 👋 Welcome tab, and no other tab shows.
- [ ] **2.8 Student Data: upload the file.** Upload file → pick "1 - Student Data (ATS).xlsx" from the test files. **You should see:** The next step opens with a green box at the top: "Student Data (ATS) completed!" and "18 rows moved to Raw Data. Next: Parent Square." The same message shows for a few seconds at the bottom of the sheet.
- [ ] **2.9 Parent Square: a wrong paste, then the right one.** Open "1 - Student Data (ATS).csv", press Ctrl+A, Ctrl+C, and paste it on the Parent Square Data tab with Ctrl+Shift+V. Then do the same with "2 - Parent Square.csv". **You should see:** The wrong paste: a red message says it looks like another export, and the tab empties by itself. The right paste: "Parent Square completed!" and "28 rows moved to Raw Data. Next: NYCSA."
- [ ] **2.10 NYCSA.** Upload "3 - NYCSA (Students with Non-NYCSA Account).xlsx". **You should see:** "NYCSA completed!" and "25 rows moved to Raw Data. Next: Site / Class."
- [ ] **2.11 Site / Class.** Open "4 - Site and Class.csv", copy the three columns with their header and paste them over A1 of the Site / Class tab. Click ✅ Save Site / Class. **You should see:** "Site / Class completed!" with the number of classes saved, and "Next: Your Lists."
- [ ] **2.12 Your Lists builds the sheets first.** Wait on the Your Lists step. **You should see:** A spinner shows the build's steps, then "Sheets built, so your lists can find your students."
- [ ] **2.13 A list with an icon (upload).** Add another list → name it "Bus Routes" → pick 🚌 → Make the tab & add the list. In its panel: Upload the file → "5a - Your Lists - Bus Routes (student list).xlsx". Go through Columns → Variables → Match (click "Accept the … high-confidence") → View, then Save. **You should see:** The new tab is named "🚌 Bus Routes". On Match: 10 matched, 6 best guesses with a % (Violet Incredable 96%), 2 not found.
- [ ] **2.14 A pasted guardian list.** _(optional)_ Back to Input Data Mode → Add another list → "PTA". Open "5b - Your Lists - PTA Members (guardian list).csv", copy everything and paste it in the yellow box on the new tab. Go through the steps and Save. **You should see:** It is guessed as a guardian list. On Match: 11 matched, 1 best guess (Martha K. 84%), 1 not found.
- [ ] **2.15 Changing a list's icon.** _(optional)_ On the Your Lists card, click the 🚌 next to Bus Routes and pick 🎒. Then type any other emoji in "Or type any emoji" and click Use it. **You should see:** The tab is renamed each time ("🎒 Bus Routes", then your emoji). Letters or numbers in that box are refused with a message.
- [ ] **2.16 Finishing Input Data Mode.** Click Finish on the Your Lists step. **You should see:** The end screen offers "Build sheets now" only if something changed after step 5's build. The list tabs are hidden afterwards. Contact Log, Combined Contact Tracking, PCAR and Directory show; the Welcome tab is gone.
- [ ] **2.17 Save and exit from Site / Class.** _(optional)_ 🚀 App Menu → 📥 Input Data Mode → click Site / Class in the bar at the top. Change one teacher's name, then Finish or exit → "Save, exit and build my sheets". **You should see:** The note at the bottom says "Site / Class completed" with NO "Next:" part. Your tabs come back and the build runs.
- [ ] **2.18 Next month's data.** _(optional)_ 📥 Input Data Mode → upload "1b - Student Data (ATS) - next month (2 new, 1 left).xlsx" → Finish or exit → Save, exit and build. **You should see:** The build report shows 2 new students (Steve Rogers, Lucia Lopez) and 1 discharged (Tony Stark).
- [ ] **2.19 Logging contacts.** In the Contact Log, type a parent's name in the Guardian column of the empty row (try Morticia Addams), then Enter. Log 3 contacts. **You should see:** The student, OSIS, date, # and follow-up fill in by themselves. Person Spoke With lists only that student's guardians (Morticia and Gomez for an Addams child). The Communications box under END shows the new notes.
- [ ] **2.20 Finishing the Getting Started checklist.** _(optional)_ ✅ Getting Started → Finish setup. **You should see:** A pop-up says you can always open it from 🚀 App Menu → ✅ Getting Started, and it no longer opens by itself.
- [ ] **2.21 Mistakes a new user makes.** _(optional)_ On a second copy of the test template, right after Initial Setup: 🏗️ Build Sheets Only before loading any data. Type on the Welcome tab. In Input Data Mode: Skip on Student Data. On Site / Class with nothing filled in: Save. Open ⚙️ Settings and try Save straight away. **You should see:** Build: a note, nothing built. Typing: a "Heads up" warning. Skip: it asks first. The empty Site / Class save doesn't erase saved sites. Settings: Save stays grey until "Loading your settings…" goes away.

### 3. A returning school
_A copy of the test template brings over "last year's tracker": a copy of your fake-data copy._
- [ ] **3.1 Last year's tracker, for the test.** Open your fake-data copy → File → Make a copy. Name it "Last year (test)". Open it → 🚀 App Menu → 🚨 Initial Setup → approve → close the question window. Look in ⚙️ Settings: phone sync is off. 📥 Input Data Mode → upload "1b - Student Data (ATS) - next month (2 new, 1 left).xlsx" → Finish or exit → Save, exit and build. **You should see:** It has the fake data, and the build report shows Tony Stark discharged. It plays last year's tracker, so your fake-data copy stays as it is for stage 4.
- [ ] **3.2 Choosing the old tracker.** Make a new copy of the test template → Initial Setup → "No, bring over last year's tracker". First paste this new tracker's own link and try it. Then pick "Last year (test)" from the list (or paste its link). **You should see:** Its own link is refused. The wizard lists your trackers.
- [ ] **3.3 All 6 steps.** Let the wizard run. **You should see:** Each step ticks, no "Success!" pop-up comes half way, and the done screen shows. The done screen says "One last step: stop the old tracker's nightly rebuild".
- [ ] **3.4 Everything came over.** Look through the new tracker. **You should see:** The Contact Log rows, the Notes tab, the Type list, Send Out, and the 3 events with their delete box and dropdowns. The New / Edit Student changes (Tina Belcher, Peter Parker), and Your Lists with their icons if you added any. Tony Stark is still on the Student List, as Discharged.
- [ ] **3.5 Person Spoke With after the move.** Open the Person Spoke With dropdown on a few rows. **You should see:** Each lists that row's student's guardians.
- [ ] **3.6 The Import Old Data menu item.** _(optional)_ Look in the 🚀 App Menu on a new copy before its first build, then on the tracker you just brought over. **You should see:** 🔄 Import Old Data is there before the first build and gone after it.
- [ ] **3.7 The old tracker pauses itself.** Open "Last year (test)" again. **You should see:** It asks "Keep this tracker paused?". No means it rebuilds again at 10 PM; Yes keeps it paused.
- [ ] **3.8 End of Year Rollover.** _(optional)_ On "Last year (test)": 🚀 App Menu → 🔄 End of Year Rollover, and follow it. **You should see:** The archive PDF is named for the school year that ends (in June 2026: 2025-2026).

### 4. Day to day
_Your fake-data copy._
- [ ] **4.1 Typing in the Contact Log.** Type a fake parent's name (Marge Simpson) in the Guardian column of the empty row and press Enter. Count the seconds until the row is filled in. **You should see:** The student, OSIS, date, # and follow-up fill in by themselves. (Seconds until the row filled in: ___)
- [ ] **4.2 Pasting and deleting rows.** _(optional)_ Paste a few rows into the Contact Log, then delete one row. **You should see:** The notes in the Communications box follow (the deleted row's note goes away).
- [ ] **4.3 Deleting row 2.** Delete Contact Log row 2 (right-click its number → Delete row). Wait a few seconds. **You should see:** The Person Spoke With dropdowns list names again, on row 2 and further down.
- [ ] **4.4 Inserting a row above row 2.** Right-click row 2's number → Insert 1 row above. Wait a few seconds. Open the Person Spoke With dropdown on two rows with different students. Then delete the empty row. **You should see:** Each row lists ITS OWN student's guardians (not the row above's).
- [ ] **4.5 A student with 3 guardians.** _(optional)_ 👤 Student Overrides → ✨ New / Edit Student: add a third guardian to Wednesday Addams, Save, and let the build run. Log a contact for Wednesday and one for another student. **You should see:** Wednesday's row lists all 3 guardians; the other row lists only its own.
- [ ] **4.6 The formula comes back after a build.** _(optional)_ Show columns R to AA (the small arrows between Q and AB), clear cell R2, then 🏗️ Build Sheets Only. **You should see:** The formula is back in R2 and the columns are hidden again.
- [ ] **4.7 The lookup box.** _(optional)_ Copy an OSIS from the Student List and type it in the OSIS box under END. **You should see:** No "are you sure?" warning; the student's details and notes fill in.
- [ ] **4.8 Your Lists: how it reads.** _(optional)_ 📥 Input Data Mode → click Your Lists in the bar at the top → Add another list → upload "5c - Your Lists - Field Trip Forms". Step 2 Variables: keep the form column, shown in Additional Information and as a Student List column. Step 4 View: "One line" and "Bold the names" → Save → Back to Input Data Mode → Finish → Build sheets now. **You should see:** Phone Contacts' notes show "Additional Information" reading like the preview. The Student List has the list's column, with a navy header.
- [ ] **4.9 The sheet's look.** _(optional)_ Look over the tabs (the look came from 🎨 Apply Sheet Style on the TEST master). **You should see:** Montserrat everywhere. Notes and Send Out brown like the Contact Log; Student List, Parents Divided, Directory and Combined navy. PCAR's boxes a darker brown with yellow date cells; no "Column 9" on the Student List; the tab is named "Raw Data" with its DATE UPDATED titles at the left.
- [ ] **4.10 Events and colors.** _(optional)_ 📆 Events → 📋 New Event and add a parent. Log a Workshop row in the Contact Log, and type a Type in it. **You should see:** The event has a brown title row and its dropdowns work; the Workshop row is pale green; the Type gets its color.
- [ ] **4.11 Backend New / Edit Student headers.** _(optional)_ Show the hidden Backend tab and look at columns G to AV. Hide it again after. **You should see:** The headers are on row 2 and students start on row 3.
- [ ] **4.12 Settings and Build.** ⚙️ Settings → change nothing → Save. Then 🚀 App Menu → 🏗️ Build Sheets Only. **You should see:** Settings closes without an error; the build sidebar ticks every step.
- [ ] **4.13 The next morning.** Open the fake-data copy the morning after. Then Extensions → Apps Script → Executions. **You should see:** No "Last night's rebuild didn't finish" note; Last Interaction shows yesterday's contacts; the Type from 4.10 kept its color. nightlyMasterSync ran around 10 PM with no red error.

### 5. Someone else (optional)
_Share your fake-data copy with a colleague, or with your own second Google account, as an Editor._
- [ ] **5.1 Another person opens it.** _(optional)_ They open the copy and count the seconds until the 🚀 App Menu shows. **You should see:** The menus show for them too. (Seconds until the menu showed for them: ___)
- [ ] **5.2 They work in it.** _(optional)_ They log a contact and open 📥 Input Data Mode. **You should see:** The Contact Log fills in for them, and Input Data Mode opens.
- [ ] **5.3 Still one nightly rebuild.** _(optional)_ They open ⚙️ Settings → Save. The next morning, you open Extensions → Apps Script → Executions. **You should see:** ONE nightlyMasterSync run around 10 PM, not two.

### 6. Phone sync
_Only from a SECOND Google account (a free Gmail account works), so the fake families land in that account's contacts, not yours._
- [ ] **6.1 A copy of its own.** Share your fake-data copy with the second account as an Editor (Share → its email). In a browser signed in to that account: open it → File → Make a copy. In that copy: 🚀 App Menu → 🚨 Initial Setup → approve. **You should see:** The copy belongs to the second account (its name shows in that account's Drive).
- [ ] **6.2 Full Sync while you watch.** ⚙️ Settings → Contacts & sync → turn on "Sync contacts to my phone every night" → Save. 🚀 App Menu → 🧱 Full Sync (Build + Push to Google Contacts). Watch it to the end. **You should see:** The bar fills and it ends with a summary. The "more than half" question, if it comes at all, comes at the start.
- [ ] **6.3 The contacts.** Open contacts.google.com in that account → the "School Directory" label. Then move that copy to the trash (a tracker in the trash skips its nightly job). **You should see:** The fake families with their phones and emails; the notes newest first; "Additional Information" from a list if the copy had one.

## Part 3: A short check with your real data
Fake data is small. These few checks show speed and this month's real exports.

### 7. Your real data
_The TEST master: your real data with the new version. Leave phone sync off on it._
- [ ] **7.1 Typing at real size.** Type a real parent's name in the Contact Log's empty row and press Enter. Count the seconds until the row is filled in. **You should see:** As fast as on your live sheet, or faster. (Seconds until the row filled in: ___)
- [ ] **7.2 Opening at real size.** Close and reopen the TEST master 3 times. Count the seconds until the 🚀 App Menu shows. **You should see:** The menu shows quickly. A few seconds after opening, the old Contact Log rows hide. (Seconds until the menu showed: ___)
- [ ] **7.3 Row 2 at real size.** _(optional)_ Delete Contact Log row 2, wait a few seconds. Then insert a row above row 2, wait, and delete it again. **You should see:** Each row's Person Spoke With dropdown lists its own student's guardians after each change.
- [ ] **7.4 This month's real exports.** 📥 Input Data Mode: load your real ATS, Parent Square and NYCSA exports (one upload, one paste). Finish → Build. **You should see:** Each step says "completed!" with sensible row counts; the build report's new and discharged students look right.
- [ ] **7.5 Your own lists.** _(optional)_ If your sheet has Your Lists lists: after the build, look at a few phone contacts' Additional Information. **You should see:** Your lists read exactly as they did before this version.
- [ ] **7.6 The next morning.** Open the TEST master the morning after. Then Extensions → Apps Script → Executions. **You should see:** No "Last night's rebuild didn't finish" note; nightlyMasterSync ran around 10 PM with no red error.

## Part 4: Release the new version
Once the tests pass: your sheet, the release copies, a pilot, then every school.

### 8. Release it
_In this order, once the tests pass._
- [ ] **8.1 Problems fixed.** Tell Claude "release check done" (Claude reads your answers here) or give the report. Re-test each check that gets a fix. **You should see:** No check is left at Problem.
- [ ] **8.2 Final check with Claude.** Ask Claude for the final check: the tests, the file list in DEPLOY.md, what to paste where. **You should see:** Claude reports everything in order.
- [ ] **8.3 Your live sheet.** Put the new version on your live tracker: DEPLOY.md sections 1 to 3 (as you did on the TEST master). **You should see:** It opens, builds and logs a contact as the TEST master did.
- [ ] **8.4 Release copies.** On your live sheet: 🐞 Debug → 📦 Make Release Copies (Example + Blank). **You should see:** Two copies, "Parent Coordinator Sheet - (EXAMPLE) (date)" and "- (BLANK) (date)", a link each and one version date (it goes in the Version Tracker's column A).
- [ ] **8.5 The Blank copy becomes the template.** Open the BLANK copy: delete every file starting with z_ and paste App Scripts (General).Json (its New Copy Checklist says so). **You should see:** Only the 👋 Welcome tab shows; "See an example tracker" on it opens the EXAMPLE copy made in the same run.
- [ ] **8.6 The template folder.** Put the BLANK copy in the template folder; move older (BLANK) and (MIGRATE) files to a subfolder. **You should see:** Only the newest (BLANK) file is in the folder itself.
- [ ] **8.7 Pilot.** Give 1 or 2 friendly schools the BLANK copy's /copy link and let them use it for a week or two. **You should see:** Nothing they report is left unfixed.
- [ ] **8.8 The Version Tracker row.** Add a row: A = the version date, D "Blank Copy" = the BLANK's /copy link, E "Example" = the EXAMPLE's link, F "Document" = the User Guide (when it changed). **You should see:** Every school's tracker now offers "New Version Available".
- [ ] **8.9 Links.** _(optional)_ 🐞 Debug → 🔗 My Links. Then 📖 Open User Guide. **You should see:** Every link from the 📌 YOUR LINKS box, and "Version Tracker, column D/E/F" with each link. "See an example tracker" shows in 📖 Open User Guide.
- [ ] **8.10 The User Guide Doc.** _(optional)_ Ask Claude for the list of what changed in docs/USER_GUIDE.md, and make those changes in the Google Doc. **You should see:** The Google Doc matches this version.
- [ ] **8.11 Merge into main.** _(optional)_ Ask Claude to merge Test-Version into main. **You should see:** main is the version schools have.
- [ ] **8.12 Clean up the test copies.** Move each test copy to the trash: the TEST master, the fake-data copy, "Last year (test)", the test template and the copies made from it. **You should see:** Nothing else to do: a tracker in the trash skips its 10 PM rebuild by itself. _Why: Every copy that ran Initial Setup has its own 10 PM rebuild under your account, and Google can keep running a trashed sheet's script until the trash is emptied. This version checks for the trash, so trashing is enough._
