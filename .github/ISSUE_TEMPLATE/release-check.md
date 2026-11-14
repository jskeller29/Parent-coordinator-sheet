---
name: Release check
about: Tick through every check on a copy before a new version goes to schools
title: "Release check: (version date)"
---

<!--
One issue per release. Tick each box as you go (it saves by itself).
Anything that fails: write it as a comment below (the step, what you saw, a screenshot,
any red error from Extensions → Apps Script → Executions). Claude can read the comments.
How to put a version on a copy: DEPLOY.md, sections 1 to 3.
Keep this list in step with the code: new checks go HERE (DEPLOY.md section 4 points to this file).
-->

How to put the new version on a copy: [DEPLOY.md](https://github.com/jskeller29/parent-coordinator-sheet/blob/Test-Version/DEPLOY.md), sections 1 to 3.
Test files for Input Data Mode and Your Lists: [test-data](https://github.com/jskeller29/parent-coordinator-sheet/tree/Test-Version/test-data) (its README says what each one should show).
Something fails? Write it as a comment on this issue: the step, what you saw, a screenshot, and any red error from Extensions → Apps Script → **Executions**.

### Before you start (test copies can touch your real things)
- [ ] Your browser is signed in to **one** Google account (several accounts can make the 🚀 App Menu slow or missing)
- [ ] **Phone sync is off** on every test copy (⚙️ Settings → Contacts & sync), except the one copy of your real data in step 5. Every copy syncs into the same "School Directory" label in your Google Contacts, the 10 PM rebuild too
- [ ] "Bring over last year's tracker" is only tested **from a copy** of an old tracker, never from your live sheet: it pauses the old tracker's 10 PM rebuild. If it happened anyway: open your live sheet and answer **No** to "Keep this tracker paused?"
- [ ] Done with a test copy: turn its 10 PM rebuild off in ⚙️ Settings, then move it to the trash

### 1. The new version on copies
- [ ] A copy of your own sheet: old files deleted, files pasted from Test-Version, the Mine manifest (DEPLOY.md 1 and 2)
- [ ] Its after-pasting steps (DEPLOY.md 3): reload, approve permissions, 🏗️ Build Sheets Only, ⚙️ Settings → Save, 🔧 Fix Dropdown Criteria, 🎨 Apply Sheet Style, the Person Spoke With dropdown by hand
- [ ] A copy of the master template: the same files without any z_ file, the General manifest, the Student List header on row 1, the Person Spoke With dropdown by hand

### 2. A brand-new school (a fresh copy of the master template, no z_ files)
- [ ] It opens on the 👋 Welcome tab only, with the 6 permissions explained first, and the 🚀 App Menu shows. Seconds until the menu showed: ___
- [ ] Typing before setup: "🔒 Action Blocked", and the typing is undone (in a copy that still has a z_ file, typing works)
- [ ] 🚨 Initial Setup: Google's screens (Authorization required → OK → your account; "hasn't verified" → Advanced → Go to … → Select all → Continue). "Setting up your tracker" opens at once, and ONLY the Welcome tab shows during and after (now "You're set up"). Seconds until the question: ___
- [ ] "Is this your first PC Tracker?": "How to download your data" and "Watch the video" show at once, "See an example tracker" a few seconds later
- [ ] Yes, start fresh → Start Input Data Mode: only the yellow Student Data tab shows, with "PASTE HERE" and the grey "Closed the side panel?" line
- [ ] Close the side panel and reload: Input Data Mode comes back by itself
- [ ] Each export from test-data (upload one, paste another): the next step opens with "✅ … completed!" at the top, and a toast on the sheet
- [ ] Paste a wrong export once: the tab empties by itself and the message stays
- [ ] Site / Class: type teachers and a site, ✅ Save Site / Class
- [ ] Your Lists: Add another list and pick an icon (🚌): its tab is "🚌 (name)". Upload a CSV, upload an Excel file, paste one; a few misspelled names get a %; accept or pick
- [ ] Your Lists: click a list's icon on its card and pick another: the tab is renamed
- [ ] Step 5 builds the sheets first; the end screen only offers Build if something changed
- [ ] After Input Data Mode closes, the list tabs are hidden
- [ ] "Finish or exit" → "Save, exit and build my sheets" in the middle of the steps, with a file pasted: its toast says "completed" and no "Next:"
- [ ] Exit before any build: back on the Welcome tab. After the first build: Contact Log, Combined, PCAR and Directory appear and the Welcome tab goes
- [ ] Log 3 contacts: the parent's name fills in the student; Person Spoke With lists only that student's guardians; the Communications box under END shows the notes
- [ ] ✅ Getting Started → Finish setup: a pop-up, and it no longer opens by itself
- [ ] Mistakes a new user makes: 🏗️ Build Sheets Only before loading data (a note, nothing built); typing on the Welcome tab after setup ("Heads up" warning); Skip on Student Data (asks first); an empty Site / Class + Save (saved sites stay); ⚙️ Settings: Save is greyed out until "Loading your settings…" goes away

### 3. A returning school (a copy of the (BLANK) template + a COPY of an old tracker)
- [ ] Initial Setup → "No, bring over last year's tracker": the wizard lists your old trackers; pick the copy. Pasting this tracker's own link is refused
- [ ] All 6 steps tick and the done screen shows (no "Success!" pop-up half way)
- [ ] It all came over: the Notes tab, Type list, Send Out, events (with their delete box and dropdowns), Your Lists (with their icons), the old "Notes" / "Followup Notes" columns; Follow-ups translated; a Discharged student still has notes
- [ ] Person Spoke With lists guardians after the move
- [ ] Before the first build 🔄 Import Old Data is in the 🚀 App Menu; after a build it's gone
- [ ] Open the old copy: it asks "Keep this tracker paused?" (No = it rebuilds again)
- [ ] 🔄 End of Year Rollover: the archive PDF is named for the school year that ends

### 4. Your real data, day to day (a copy of your own sheet)
- [ ] Contact Log: type a parent name (the student, date and follow-up fill in as fast as before), paste a few rows, delete a row: the notes follow
- [ ] Delete row 2: within a few seconds the Person Spoke With dropdowns list names again
- [ ] Insert a row above row 2: within a few seconds each row's dropdown lists its OWN student's guardians
- [ ] A student with 2 or 3 guardians (one added with ✨ New / Edit Student) lists only them; the next student's row lists theirs
- [ ] Clear cell R2 (show columns R:AA first, hide them after) and 🏗️ Build Sheets Only: the formula comes back
- [ ] The lookup box: type an OSIS below END (no "are you sure?" warning)
- [ ] Your Lists step 4 View: One line + bold names, Build, then a phone contact's Additional Information; a list saved before this version reads exactly as it did
- [ ] 🎨 Apply Sheet Style: Montserrat on every tab; Notes / Send Out brown like the Contact Log; Student List, Parents Divided, Directory, Combined navy; PCAR's boxes a darker brown with yellow date cells; no "Column 9" on the Student List; the "Raw Data" tab name; Raw Data's "DATE UPDATED" titles at the left
- [ ] Then: a new event (brown title row); Workshop rows in the Contact Log pale green; a Type typed in the Contact Log keeps its color after the 10 PM run; a reset leaves the Student List at A:H; a Your Lists list adds its column back with a navy header
- [ ] Backend New/Edit Student (G:AV): after a build the headers are on row 2 and students start on row 3
- [ ] ⚙️ Settings → Save; 🏗️ Build Sheets Only finishes in its sidebar
- [ ] Close and reopen a few times: the 🚀 App Menu shows quickly; old Contact Log rows hide a few seconds after opening
- [ ] Next morning: no "Last night's rebuild didn't finish" note, and Last Interaction is updated

### 5. Phone sync (only on the copy of your real data)
- [ ] 🧱 Full Sync while you watch (the "more than half" question, if it comes, is asked at the start)
- [ ] 3 contacts on your phone: their notes, and Additional Information from a list

### 6. Someone else
- [ ] Share a test copy with a colleague (or your second Google account): they open it (the menus show), log a contact, open Input Data Mode
- [ ] After they save ⚙️ Settings, the next morning Extensions → Apps Script → Executions shows ONE nightlyMasterSync run at 10 PM

### 7. Release copies and links (your sheet)
- [ ] 🧹 Reset to Display Mode on the master template: only the 👋 Welcome tab shows, also after reopening it
- [ ] 🐞 Debug → 📦 Make Release Copies: "Parent Coordinator Sheet - (EXAMPLE / BLANK) (date)", a link each and one version date (it goes in the Version Tracker's column A)
- [ ] The Blank copy's 👋 Welcome tab: "See an example tracker" opens the Example copy made in the same run
- [ ] Template folder: only the newest "(BLANK)" file (old "(MIGRATE)" files in a subfolder); the Version Tracker's column D = that file's /copy link
- [ ] Version Tracker headers: D "Blank Copy", E "Example", F "Document" (old values in E cleared); 🔗 My Links shows "Version Tracker, column D/E/F" with each link; "See an example tracker" shows in 📖 Open User Guide
- [ ] 🔗 My Links lists every link from the 📌 YOUR LINKS box
- [ ] The User Guide Google Doc matches docs/USER_GUIDE.md for this version (ask Claude for the list of what changed)

### 8. Before it goes to schools
- [ ] Everything that failed is a comment on this issue and has been fixed with Claude
- [ ] Final check with Claude: the tests pass, the file list matches DEPLOY.md, no z_ files in the master template, the Version Tracker row
- [ ] Pilot: 1 or 2 friendly schools use it for a week or two
- [ ] DEPLOY.md section 5: merge Test-Version into main, update the master template, add the Version Tracker row
