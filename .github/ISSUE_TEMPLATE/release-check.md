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

### Before you start
_Do these once. Test copies can reach your real Google Contacts and your live sheet._
- [ ] **0.1 One Google account in this browser.** Use a browser window signed in to only your school Google account (a private window where you sign in once works too). **You should see:** Only one account shows when you click your picture at the top right of Google Sheets. _Why: Several accounts in one browser can make the 🚀 App Menu slow or missing._
- [ ] **0.2 Phone sync off on every test copy.** On each test copy: ⚙️ Settings → Contacts & sync → turn off "Sync contacts to my phone every night" → Save. Leave it on only on the one copy of your real data in stage 5. **You should see:** The switch is off after you open ⚙️ Settings again. _Why: Every copy syncs into the same "School Directory" label in your Google Contacts, the 10 PM rebuild too._
- [ ] **0.3 Bring over only from a copy.** Before stage 3, make a copy of an old tracker (File → Make a copy). Bring over from that copy, never from your live sheet. **You should see:** Your live sheet never asks "Keep this tracker paused?". _Why: Bringing a tracker over pauses that tracker's 10 PM rebuild. If it happens to your live sheet: open it and answer No._

### 1. Put the new version on copies
_A copy of your own sheet and the master template._
- [ ] **1.1 A test copy of your own sheet.** Open your live tracker → File → Make a copy. Put "TEST" and the date in its name. Open the copy. When it asks "Is this your MAIN sheet?", answer No: its New Copy Checklist opens. **You should see:** The copy has all your data and your 🐞 Debug menu.
- [ ] **1.2 Paste the new files.** In the copy: Extensions → Apps Script. Follow DEPLOY.md sections 1 and 2: delete the old files it lists, paste every file from Test-Version, keep your z_ files, paste App Scripts (Mine).Json into appsscript.json. Save. **You should see:** Saving shows no red error. The file list matches the table in DEPLOY.md section 2.
- [ ] **1.3 The after-pasting steps.** Reload the sheet. Run any menu item and approve the permissions. 🚀 App Menu → 🏗️ Build Sheets Only. Then ⚙️ Settings → Save. 🐞 Debug → 🔧 Fix Dropdown Criteria, then 🐞 Debug → 🎨 Apply Sheet Style. Set the Person Spoke With dropdown by hand (DEPLOY.md section 3, step 9). **You should see:** The build sidebar ticks every step. Apply Sheet Style ends with its list of things to do by hand.
- [ ] **1.4 The master template.** Put the same files on the master template, WITHOUT any z_ file, with App Scripts (General).Json as its manifest. Check its Student List header is on row 1 (frozen) and set its Person Spoke With dropdown by hand. **You should see:** Extensions → Apps Script on the template shows no file starting with z_.

### 2. A brand-new school
_A fresh copy of the master template, like a school makes. Keep the test-data files at hand._
- [ ] **2.1 Opening a fresh copy.** Open the template's /copy link (or File → Make a copy of the template) and click Make a copy. Count the seconds until 🚀 App Menu shows at the top. **You should see:** Only the 👋 Welcome tab shows, starting with the 6 permissions Google will ask for. A note at the bottom right points to 🚀 App Menu → 🚨 Initial Setup. The menu shows. The very first open can take up to a minute; if it still isn't there, reload. (Seconds until the menu showed: ___)
- [ ] **2.2 Typing is blocked before setup.** Type anything in a cell of the Welcome tab and press Enter. **You should see:** "🔒 Action Blocked", and your typing is undone.
- [ ] **2.3 Initial Setup and Google's screens.** 🚀 App Menu → 🚨 Initial Setup (Run Once). Authorization required → OK → pick your account. "Google hasn't verified this app" → Advanced → Go to … (unsafe). Select all → Continue. Count the seconds until the question shows. **You should see:** "Setting up your tracker" opens right away, with a spinner. Only the Welcome tab shows during and after setup; it now says "You're set up". The 🚀 App Menu changes to the getting-started menu. (Seconds until the question showed: ___)
- [ ] **2.4 "Is this your first PC Tracker?".** Click "Yes, start fresh". **You should see:** "How to download your data" and "Watch the video" are there at once; "See an example tracker" can come a few seconds later. The ✅ Getting Started checklist opens on the right.
- [ ] **2.5 Starting Input Data Mode.** Click "Start Input Data Mode". **You should see:** Only the yellow "Student Data" tab shows, with PASTE HERE in A1 and a grey "Closed the side panel?" line. The side panel says "Data 1 of 5: Student Data (ATS / Insight)" and which file to pick (its name starts with general_).
- [ ] **2.6 The side panel comes back.** Close the side panel with its X, then reload the page. **You should see:** Input Data Mode comes back by itself, on the same step.
- [ ] **2.7 Leaving before any build.** In the side panel: Finish or exit → Exit without saving. Then 🚀 App Menu → 📥 Input Data Mode to start it again. **You should see:** After the exit you are back on the 👋 Welcome tab, and no other tab shows.
- [ ] **2.8 Student Data: upload the file.** Upload file → pick "1 - Student Data (ATS).xlsx" from test-data. **You should see:** The next step opens with a green box at the top: "Student Data (ATS) completed!" and "18 rows moved to Raw Data. Next: Parent Square." The same message shows for a few seconds at the bottom of the sheet.
- [ ] **2.9 Parent Square: a wrong paste, then the right one.** Open "1 - Student Data (ATS).csv", press Ctrl+A, Ctrl+C, and paste it on the Parent Square Data tab with Ctrl+Shift+V. Then do the same with "2 - Parent Square.csv". **You should see:** The wrong paste: a red message says it looks like another export, and the tab empties by itself. The right paste: "Parent Square completed!" and "28 rows moved to Raw Data. Next: NYCSA."
- [ ] **2.10 NYCSA.** Upload "3 - NYCSA (Students with Non-NYCSA Account).xlsx". **You should see:** "NYCSA completed!" and "25 rows moved to Raw Data. Next: Site / Class."
- [ ] **2.11 Site / Class.** Open "4 - Site and Class.csv", copy the three columns with their header and paste them over A1 of the Site / Class tab. Click ✅ Save Site / Class. **You should see:** "Site / Class completed!" with the number of classes saved, and "Next: Your Lists."
- [ ] **2.12 Your Lists builds the sheets first.** Wait on the Your Lists step. **You should see:** A spinner shows the build's steps, then "Sheets built, so your lists can find your students."
- [ ] **2.13 A list with an icon (upload).** Add another list → name it "Bus Routes" → pick 🚌 → Make the tab & add the list. In its panel: Upload the file → "5a - Your Lists - Bus Routes (student list).xlsx". Go through Columns → Variables → Match (click "Accept the … high-confidence") → View, then Save. **You should see:** The new tab is named "🚌 Bus Routes". On Match: 10 matched, 6 best guesses with a % (Violet Incredable 96%), 2 not found.
- [ ] **2.14 A pasted guardian list.** Back to Input Data Mode → Add another list → "PTA". Open "5b - Your Lists - PTA Members (guardian list).csv", copy everything and paste it in the yellow box on the new tab. Go through the steps and Save. **You should see:** It is guessed as a guardian list. On Match: 11 matched, 1 best guess (Martha K. 84%), 1 not found.
- [ ] **2.15 Changing a list's icon.** On the Your Lists card, click the 🚌 next to Bus Routes and pick 🎒. Then type any other emoji in "Or type any emoji" and click Use it. **You should see:** The tab is renamed each time ("🎒 Bus Routes", then your emoji). Letters or numbers in that box are refused with a message.
- [ ] **2.16 Finishing Input Data Mode.** Click Finish on the Your Lists step. **You should see:** The end screen offers "Build sheets now" only if something changed after step 5's build. The list tabs are hidden afterwards. Contact Log, Combined Contact Tracking, PCAR and Directory show; the Welcome tab is gone.
- [ ] **2.17 Save and exit from Site / Class.** 🚀 App Menu → 📥 Input Data Mode → click Site / Class in the bar at the top. Change one teacher's name, then Finish or exit → "Save, exit and build my sheets". **You should see:** The note at the bottom says "Site / Class completed" with NO "Next:" part. Your tabs come back and the build runs.
- [ ] **2.18 Next month's data.** 📥 Input Data Mode → upload "1b - Student Data (ATS) - next month (2 new, 1 left).xlsx" → Finish or exit → Save, exit and build. **You should see:** The build report shows 2 new students (Steve Rogers, Lucia Lopez) and 1 discharged (Tony Stark).
- [ ] **2.19 Logging contacts.** In the Contact Log, type a parent's name in the Guardian column of the empty row (try Morticia Addams), then Enter. Log 3 contacts. **You should see:** The student, OSIS, date, # and follow-up fill in by themselves. Person Spoke With lists only that student's guardians (Morticia and Gomez for an Addams child). The Communications box under END shows the new notes.
- [ ] **2.20 Finishing the Getting Started checklist.** ✅ Getting Started → Finish setup. **You should see:** A pop-up says you can always open it from 🚀 App Menu → ✅ Getting Started, and it no longer opens by itself.
- [ ] **2.21 Mistakes a new user makes.** On a second fresh copy, right after Initial Setup: 🏗️ Build Sheets Only before loading any data. Type on the Welcome tab. In Input Data Mode: Skip on Student Data. On Site / Class with nothing filled in: Save. Open ⚙️ Settings and try Save straight away. **You should see:** Build: a note, nothing built. Typing: a "Heads up" warning. Skip: it asks first. The empty Site / Class save doesn't erase saved sites. Settings: Save stays grey until "Loading your settings…" goes away.

### 3. A returning school
_A fresh copy of the (BLANK) template, and a COPY of an old tracker to bring over._
- [ ] **3.1 Choosing the old tracker.** Fresh copy of the template → Initial Setup → "No, bring over last year's tracker". First paste this new tracker's own link and try it. Then pick the COPY of the old tracker from the list. **You should see:** Its own link is refused. The wizard lists your old trackers.
- [ ] **3.2 All 6 steps.** Let the wizard run. **You should see:** Each step ticks, no "Success!" pop-up comes half way, and the done screen shows. The done screen says "One last step: stop the old tracker's nightly rebuild".
- [ ] **3.3 Everything came over.** Look through the new tracker. **You should see:** The Notes tab, the Type list, Send Out, and the events with their delete box and dropdowns. Your Lists, with their icons, and the old "Notes" / "Followup Notes" columns. Follow-ups translated (Yes is now Followup Needed); a Discharged student still has notes.
- [ ] **3.4 Person Spoke With after the move.** Open the Person Spoke With dropdown on a few rows. **You should see:** Each lists that row's student's guardians.
- [ ] **3.5 The Import Old Data menu item.** Look in the 🚀 App Menu on a fresh copy before its first build, then after a build. **You should see:** 🔄 Import Old Data is there before the first build and gone after it.
- [ ] **3.6 The old copy pauses itself.** Open the old COPY you brought over from. **You should see:** It asks "Keep this tracker paused?". No means it rebuilds again at 10 PM; Yes keeps it paused.
- [ ] **3.7 End of Year Rollover.** On a test copy: 🚀 App Menu → 🔄 End of Year Rollover, and follow it. **You should see:** The archive PDF is named for the school year that ends (in June 2026: 2025-2026).

### 4. Your real data, day to day
_The TEST copy of your own sheet from stage 1._
- [ ] **4.1 Typing in the Contact Log.** Type a parent's name in the empty row and press Enter. Count the seconds until the student, date and follow-up are filled in. **You should see:** It feels as fast as before, or faster. (Seconds until the row filled in: ___)
- [ ] **4.2 Pasting and deleting rows.** Paste a few rows into the Contact Log, then delete one row. **You should see:** The notes in the Communications box follow (the deleted row's note goes away).
- [ ] **4.3 Deleting row 2.** Delete Contact Log row 2 (right-click its number → Delete row). Wait a few seconds. **You should see:** The Person Spoke With dropdowns list names again, on row 2 and further down.
- [ ] **4.4 Inserting a row above row 2.** Right-click row 2's number → Insert 1 row above. Wait a few seconds. Open the Person Spoke With dropdown on two rows with different students. Then delete the empty row. **You should see:** Each row lists ITS OWN student's guardians (not the row above's).
- [ ] **4.5 Students with 3 guardians.** 👤 Student Overrides → ✨ New / Edit Student: add a third guardian to a student, Save, and let the build run. Log a contact for that student and one for the next student. **You should see:** That row lists all 3 guardians; the next row lists only its own.
- [ ] **4.6 The formula comes back after a build.** Show columns R to AA (the small arrows between Q and AB), clear cell R2, then 🏗️ Build Sheets Only. **You should see:** The formula is back in R2 and the columns are hidden again.
- [ ] **4.7 The lookup box.** Type an OSIS in the OSIS box under END. **You should see:** No "are you sure?" warning; the student's details and notes fill in.
- [ ] **4.8 Your Lists: how it reads.** 📥 Input Data Mode → Your Lists → Edit setup on a list → step 4 View → "One line" and "Bold the names" → Save. Back to Input Data Mode → Finish → Build. **You should see:** A phone contact's Additional Information reads like the preview. A list saved before this version reads exactly as it did.
- [ ] **4.9 Apply Sheet Style: the look.** Look over the tabs after 🎨 Apply Sheet Style (stage 1). **You should see:** Montserrat everywhere. Notes and Send Out brown like the Contact Log; Student List, Parents Divided, Directory and Combined navy. PCAR's boxes a darker brown with yellow date cells; no "Column 9" on the Student List; the tab is named "Raw Data" with its DATE UPDATED titles at the left.
- [ ] **4.10 Apply Sheet Style: what follows.** 📆 Events → 📋 New Event. Log a Workshop row in the Contact Log. Type a Type in the Contact Log. Add a Your Lists column and build. Next morning, look again. **You should see:** The event has a brown title row; Workshop rows are pale green; the Type keeps its color after the 10 PM run; the list's column has a navy header.
- [ ] **4.11 Backend New / Edit Student headers.** After a build, show the hidden Backend tab and look at columns G to AV. Hide it again after. **You should see:** The headers are on row 2 and students start on row 3.
- [ ] **4.12 Settings and Build.** ⚙️ Settings → change nothing → Save. Then 🚀 App Menu → 🏗️ Build Sheets Only. **You should see:** Settings closes without an error; the build sidebar ticks every step.
- [ ] **4.13 Opening the sheet.** Close and reopen the sheet 3 times. Count the seconds until the 🚀 App Menu shows. **You should see:** The menu shows quickly. A few seconds after opening, the old Contact Log rows hide. (Seconds until the menu showed: ___)
- [ ] **4.14 Your real exports.** Load this month's real ATS, Parent Square and NYCSA exports with 📥 Input Data Mode (one upload, one paste). **You should see:** Each step says "completed!" with sensible row counts; the build report's new and discharged students look right.
- [ ] **4.15 The next morning.** Open the sheet the morning after. Then Extensions → Apps Script → Executions. **You should see:** No "Last night's rebuild didn't finish" note; Last Interaction shows yesterday's contacts. nightlyMasterSync ran around 10 PM with no red error.

### 5. Phone sync
_Only the test copy of your real data. Turn phone sync on there for this stage._
- [ ] **5.1 Full Sync while you watch.** ⚙️ Settings → Contacts & sync → turn on "Sync contacts to my phone every night" → Save. 🚀 App Menu → 🧱 Full Sync (Build + Push to Google Contacts). Watch it to the end. **You should see:** The bar fills and it ends with a summary. The "more than half" question, if it comes at all, comes at the start.
- [ ] **5.2 Contacts on your phone.** Open 3 synced contacts on your phone (one with a Your Lists value). **You should see:** Names, phones and emails are right; the notes are newest first; Additional Information shows the list's value.

### 6. Someone else
_Share a test copy with a colleague, or with your own second Google account, as an Editor._
- [ ] **6.1 Another person opens it.** They open the copy and count the seconds until the 🚀 App Menu shows. **You should see:** The menus show for them too. (Seconds until the menu showed for them: ___)
- [ ] **6.2 They work in it.** They log a contact and open 📥 Input Data Mode. **You should see:** The Contact Log fills in for them, and Input Data Mode opens.
- [ ] **6.3 Still one nightly rebuild.** They open ⚙️ Settings → Save. The next morning, you open Extensions → Apps Script → Executions. **You should see:** ONE nightlyMasterSync run around 10 PM, not two.

### 7. Release copies and links
_Your own sheet and the template folder._
- [ ] **7.1 Reset to Display Mode.** Run 🧹 Reset to Display Mode on the master template (type RESET, confirm). Close it and open it again. **You should see:** Only the 👋 Welcome tab shows, also after reopening.
- [ ] **7.2 Make Release Copies.** 🐞 Debug → 📦 Make Release Copies. **You should see:** Two copies, "Parent Coordinator Sheet - (EXAMPLE) (date)" and "- (BLANK) (date)", a link each and one version date (it goes in the Version Tracker's column A).
- [ ] **7.3 The Blank copy's Welcome tab.** Open the BLANK copy. **You should see:** "See an example tracker" on its 👋 Welcome tab opens the EXAMPLE copy made in the same run.
- [ ] **7.4 Template folder and Version Tracker.** Look in the template folder and the Version Tracker. **You should see:** Only the newest (BLANK) file in the folder (old (MIGRATE) files in a subfolder). Headers D "Blank Copy", E "Example", F "Document"; D in the newest row = the BLANK file's /copy link.
- [ ] **7.5 My Links.** 🐞 Debug → 🔗 My Links. Then 📖 Open User Guide. **You should see:** Every link from the 📌 YOUR LINKS box, and "Version Tracker, column D/E/F" with each link. "See an example tracker" shows in 📖 Open User Guide.
- [ ] **7.6 The User Guide Doc.** Ask Claude for the list of what changed in docs/USER_GUIDE.md, and make those changes in the Google Doc. **You should see:** The Google Doc matches this version.

### 8. Before it goes to schools
_Last steps, in this order._
- [ ] **8.1 Problems fixed.** Give Claude your report (or say "release check done": Claude can read your answers here). Re-test what gets fixed. **You should see:** No check is left at Problem.
- [ ] **8.2 Final check with Claude.** Ask Claude for the final check: tests, the file list in DEPLOY.md, no z_ files in the template, the Version Tracker row. **You should see:** Claude reports everything in order.
- [ ] **8.3 Pilot.** Give it to 1 or 2 friendly schools for a week or two. **You should see:** Nothing they report is left unfixed.
- [ ] **8.4 Clean up the test copies.** On each test copy: ⚙️ Settings → Live updates → Rebuild sheets → "Only when I click Build", phone sync off, Save. Then move it to the trash. **You should see:** No test copy rebuilds or syncs at night any more.
- [ ] **8.5 Going live.** DEPLOY.md section 5: merge Test-Version into main, update the master template, add the Version Tracker row. **You should see:** Schools get "New Version Available".
