<!--
  SOURCE OF THE USER GUIDE. Edit this file only, then rebuild the Word, HTML and PDF versions:
      node docs/build_guide.js
  (see docs/README.md). Comments like this one are notes for whoever maintains the guide;
  they never appear in the guide itself.
  Notes boxes:  > [!NOTE]  (grey, a tip)   > [!IMPORTANT]  (navy, read this)   > [!WARNING]  (red, can lose data)
-->

# PC Tracker User Guide

Setting up and using the Parent Coordinator Tracker

Made by a parent coordinator for parent coordinators. This is not an official DOE document.
Questions: use 🚀 App Menu → 🐞 Report a Bug, or email jacob.keller@p94m.org.

<!--
  START HERE box. Update BOTH links at every release, at the same time as the Version Tracker row:
    Make your copy        = the Blank copy's /copy link (Version Tracker column D)
    See an example tracker = the Example copy's link   (Version Tracker column E)
  In the Google Doc: right-click each link → Edit link → paste. Until then the copy link opens the
  template folder (it always works); so does the example link, until you paste the Example copy's link.
-->

> [!IMPORTANT]
> **Start here.** [Make your copy](https://drive.google.com/drive/folders/1goKFrxDJOSAPxp1HUXETX2dns5c6HXk2):
> your own private tracker, in your Google Drive. Then follow Part 1 below.
>
> Want to look around first? [See an example tracker](https://drive.google.com/drive/folders/1goKFrxDJOSAPxp1HUXETX2dns5c6HXk2)
> with made-up students (view only).

## Before you start

### What the tracker does

The PC Tracker is a Google Sheet with a script attached. You load your student exports
once, and the tracker builds the rest: a family directory, a contact log that fills in
student details as you type, a history of every contact per student, and your PCAR
counts. If you want, it can also put the school directory on your phone through Google
Contacts.

| Kind of tab | Tabs | What you do there |
|---|---|---|
| Tabs you type in | Contact Log, Events, Notes, Send Out | Log calls, meetings and events. |
| Tabs that fill themselves | Student List, Directory, Combined Contact Tracking, Parents Divided, Phone Contacts, PCAR | Read them. A build rewrites them, so don't type in them. |
| Data and settings | Raw Data, Backend | Raw Data holds your exports. Backend is hidden and holds the tracker's own settings. |

### Get your copy

Every coordinator works in their own copy. Click **Make your copy** in the box at the top of
this guide, then **Make a copy**. (Or open the
[PC Tracker template folder](https://drive.google.com/drive/folders/1goKFrxDJOSAPxp1HUXETX2dns5c6HXk2),
open the newest file, its name ends in "(BLANK)" and a date, and choose File → Make a copy.)
The copy is private and lives in your own Google Drive.

The same template is for everyone. If you used the tracker last year, you bring your old
data over during setup.

Want to see one filled in first? "See an example tracker" (the box at the top of this guide,
the 👋 Welcome tab, and 🚀 App Menu → 📖 Open User Guide) opens a tracker with made-up students. Look around in it,
but keep your real work in your own copy.

## Part 1: Setting up

### Quick start

Setup takes about 10 minutes.

1. Make your copy (see above). It opens on the 👋 Welcome tab, the only tab showing.
2. Read the Welcome tab. It explains each permission Google will ask for.
3. Click 🚀 App Menu → 🚨 Initial Setup (Run Once) and allow it.
4. Follow the Getting Started checklist that opens on the right.

Right after Initial Setup, a window asks whether this is your first PC Tracker. If you
used one last year, choose "No, bring over last year's tracker" (Part 3).

### Step 1: Open your copy

Until setup is done, your copy shows only the 👋 Welcome tab. It explains the
permissions, lists the setup steps and links to this guide and the bug report form. A
message at the bottom of the screen points you to Initial Setup. After setup, the
Welcome tab stays and shows what to do next. It comes back each time you open the
sheet, so a window you closed by accident is never a dead end. It goes away by itself
once your sheets are built.

### Step 2: What Google will ask for

Initial Setup asks Google for permission once. Google shows its own screens before the
tracker can show anything, so here is what each permission is for, in Google's words:

| Google's wording | What the tracker uses it for |
|---|---|
| See, edit, create and delete all your Google Sheets spreadsheets | Building and updating the tabs in this tracker. It also reads the shared Version Tracker sheet to tell you about a new version, and opens last year's tracker when you use Import Old Data. |
| See, edit, download and permanently delete your contacts | Only the phone sync, which is off until you turn it on in Settings. It only adds, changes and removes contacts in its own "School Directory" label. |
| See, edit, create and delete all of your Google Drive files | Saving your end-of-year PDF archive next to the tracker, finding the newest User Guide and template links, and finding your old tracker when you bring it over. |
| Connect to an external service | Only the end-of-year PDF, which it downloads from Google's own export page. |
| Allow this application to run when you are not present | The updates when you type in the Contact Log, the nightly rebuild at 10 PM and the update check when you open the sheet. |
| Display and run third-party web content in prompts and sidebars | The tracker's own pop-ups and side panels, such as Settings and Input Data Mode. "Third-party" means not made by Google. |

All of this happens inside your own Google account. There is no outside server, and
nothing about your families is sent to the person who made the tracker. Part 4 has more.

### Step 3: Run Initial Setup

1. Wait for 🚀 App Menu to appear at the top of the screen, after Help. It can take a
   few seconds. If it isn't there after a minute, reload the page.
2. Click 🚀 App Menu → 🚨 Initial Setup (Run Once).
3. "Authorization required": click OK, then choose your Google account.
4. "Google hasn't verified this app": click Advanced, then "Go to … (unsafe)". Google
   shows this for tools made by a person rather than a company. It does not mean
   something is wrong.
5. On the list of permissions, tick Select all, then click Continue. Setup needs all of
   them. If one is left unticked, Google asks again.
6. If nothing happens after that, click 🚨 Initial Setup (Run Once) again.

A "Setting up your tracker" window opens right away. Keep it open: setup takes about
20 seconds. Setup then:

- turns on the automatic parts (Contact Log updates, the 10 PM rebuild, the update check),
- keeps you on the Welcome tab, which now shows what to do next. Your everyday tabs
  (Contact Log, Combined Contact Tracking, PCAR and Directory) stay hidden until your
  sheets are built, because they are empty until then. You can turn on others later in
  Settings → Tabs & columns,
- asks "Is this your first PC Tracker?": "Yes, start fresh" shows the first steps, and
  "No, bring over last year's tracker" opens the Migration Wizard (Part 3),
- opens the Getting Started checklist.

> [!NOTE]
> Before setup, the sheet won't let you type. You'll see "Action Blocked:
> Initialization process needs to occur first." Run Initial Setup and it goes away.

### Step 4: The Getting Started checklist

The checklist opens on the right after setup and ticks each step off by itself:

1. Load your student data (Step 5).
2. Build your sheets. Input Data Mode does this for you.
3. Check your settings (Part 2).
4. Log your first parent contact (Part 2).

It opens each time you open the sheet until all four are done, or until you click
Finish setup at the bottom of it. You can always reopen it from 🚀 App Menu → ✅ Getting
Started. If you close Input Data Mode half way, it opens again by itself the next time
you open the sheet, or from 🚀 App Menu → 📥 Input Data Mode.
Below the steps, "This tab" explains the tab you are on, and after your first build
"Hidden tabs" lists the tabs you can turn on.

### Step 5: Load your data

Open 🚀 App Menu → 📥 Input Data Mode, or click Open Input Data Mode in the checklist.
It takes you through each export in turn, then two optional steps: Site / Class (the
teacher and site for each class) and Your Lists. Your other tabs are hidden while you
work. If you close its panel by accident, the yellow tab still says how to bring it back:
🚀 App Menu → 📥 Input Data Mode. It picks up where you left off.

For each export, either click Upload file and pick the file you downloaded, or open the
file, select everything (Ctrl+A, Mac ⌘+A), copy it, and paste it anywhere on the
temporary tab with Ctrl+Shift+V (Mac ⌘+Shift+V). The data moves into Raw Data on its
own. If you pasted the wrong export, or Excel shortened the numbers (they show as
7.19E+09), the tracker tells you and changes nothing. Skip keeps the old data for that
export.

Before the Your Lists step, Input Data Mode builds your sheets. Finish or exit ends it
at any step. Save, exit and build my sheets saves what is on the tab, brings your tabs
back and builds.

#### Where to get each export

Student Data (required), from ATS / Insight:

1. Sign in to TeachHub and open Insight.
2. Go to Tools → Downloader.
3. Under "Data Available for Download", select only Guardian Contact Information, then
   download.

Parent Square (optional):

1. Sign in to ParentSquare and open Admin.
2. Under Data Assistant, select Parents.
3. Click Export CSV.

NYCSA (optional):

1. Sign in to TeachHub and open Family Access Management.
2. Export the "Students with Non-NYCSA Account" report.

#### Pasting into Raw Data yourself

You can also paste an export straight into the Raw Data tab (turn it on in Settings →
Tabs & columns): Student Data into cell A5, Parent Square into AG6, NYCSA into AP2.

> [!IMPORTANT]
> Paste values only, with Ctrl+Shift+V or Edit → Paste special → Values only. A normal
> paste can bring in formatting or formulas that break the build. Input Data Mode
> does this for you.

#### Your own lists

On the Your Lists step you can add any other list, such as bus routes, after-school
sign-ups, PTA members or trip forms. Click Add another list, give it a name, and paste
the whole sheet (headers too) into the new 📥 tab, or upload the file. The side panel
then has four steps:

1. Columns. It finds the header row and guesses what each column is. Fix any guess, and
   say whether each row is a student or a guardian.
2. Variables. Pick the other columns to keep (a route, a fee, a form status), name
   them, and choose where they show: on the phone contact under Additional Information,
   and as a column in the Student List or Parents Divided.
3. Match. It works out who each row is: by OSIS first, then name, then phone or email,
   then best guesses for nicknames and short names. You confirm the rows it isn't sure
   of. Check a few best guesses before you accept them all.
4. View. Choose how the values read (one per line or all on one line) and click Save.

Your choices are remembered. When a new version of the list comes in, click Paste new
version and only the new rows need you. A list with an OSIS column matches best.

> [!WARNING]
> Anything shown under Additional Information goes into your Google Contacts, and onto
> your phone after the next sync. Keep private lists (IEP, housing, health) off the
> phone contacts.

#### Fixing a student's details

ATS can lag behind real enrollment. To add a student who isn't in ATS yet, or to fix a
name or a parent's phone, use 👤 Student Overrides → ✨ New / Edit Student. Your change
shows in the Student List, Combined Contact Tracking, Parents Divided and Phone
Contacts at the next build. The Directory always shows ATS data only. An override
removes itself once ATS has the same information.

### Step 6: Build the sheets

Input Data Mode builds your sheets for you. After new data, overrides or lists, build
again with 🚀 App Menu → 🏗️ Build Sheets Only. A side panel shows each step as it
runs, then a short report: new students, discharged students, and families who signed
up for NYCSA or ParentSquare. Click a number to see the names.

After the first build, the full menus appear: ⚙️ Settings, 👤 Student Overrides and
📆 Events. If you don't see them, reload the page.

> [!NOTE]
> Raw Data is the only tab you paste exports into. Everything the tracker builds is
> rebuilt from it, so building again is always safe.

## Part 2: Daily use

### The menus

The tracker adds its menus after Help: 🚀 App Menu, 👤 Student Overrides and 📆 Events.

| 🚀 App Menu item | What it does |
|---|---|
| 🧱 Full Sync | Builds the sheets and updates your Google Contacts. Shows only when the phone sync is on. |
| 🏗️ Build Sheets Only | Rebuilds every tab from your data without touching Google Contacts. |
| 📥 Input Data Mode | Loads new exports, Site / Class and your lists (Part 1, Step 5). |
| ✅ Getting Started | The setup checklist. |
| ⚙️ Settings | Which tabs show, when sheets rebuild, the phone sync and more (below). |
| 📖 Open User Guide | A short welcome window with a link to this guide. |
| 🐞 Report a Bug | A short form for anything broken or confusing. |
| 🔔 Check for Updates | Checks for a newer version of the tracker. |
| 🔄 End of Year Rollover | Archives the year and freezes the sheet (Part 3). |

| 👤 Student Overrides item | What it does |
|---|---|
| ✨ New / Edit Student | Add a student who isn't in ATS, or correct a name, phone, email, language, grade or up to five guardians. You can also give a student a status label or hide them from the report. |
| 🗂️ Manage Overrides | Lists your changes. Edit or delete any of them. |
| 📄 Override Report | A tab comparing ATS with your changes. |

### The tabs

| Tab | What it is | Shown at first |
|---|---|---|
| Contact Log | Where you log every parent contact. | Yes |
| Combined Contact Tracking | Every student's history in one place: Contact Log notes, events and Notes tab notes, newest first. Useful before a meeting. | Yes |
| PCAR | Your PCAR counts, filled in from the other tabs. | Yes |
| Directory | Each student's site, class and parents' phone and email, from ATS. | Yes |
| Student List | One row per student: guardians, site, class, siblings, status, last interaction, and a column for each student list you keep. | No |
| Parents Divided | One row per guardian, with Parent Square and NYCSA sign-up. Useful for mail merges. | No |
| Phone Contacts | What the phone sync sends to Google Contacts. | No |
| Notes | Notes you always want at the top of a student's history. | No |
| Events | Attendance for events and workshops. | No |
| Send Out | Messages and flyers you sent, for the PCAR. | No |
| Raw Data | Your exports: ATS, Parent Square, NYCSA, sites and classes. | No |
| 📥 your lists | One tab per list from Your Lists. | No |
| Backend | The tracker's own settings and data. Leave it alone, except the Type List in column BL. | No, always hidden |

Turn tabs on or off in ⚙️ Settings → Tabs & columns. Hidden tabs keep updating.

### The Contact Log

Log every call, email, meeting and walk-in here, one row each.

| Column | What to enter |
|---|---|
| Date | Fills in today if left blank. Type a date to log an earlier contact. |
| Guardian, OSIS or Student | Type one of them: the parent's name (C), the OSIS (E) or the student's name (F). The other two fill in. |
| Person Spoke With | Who you actually spoke to. |
| # | How many people this counts for in the PCAR. A call is 1; an email to 30 parents is 30. Fills in 1 if blank. |
| Site and Class | Fill in from the student. Turn them on in Settings. |
| Method 1 and Method 2 | How you reached the family, for the PCAR. |
| Type | Your own topic list (Attendance, Behavior …). Turn it on in Settings; edit the choices in the Backend tab, column BL. |
| Event Description | What you talked about. |
| Follow-up | No Followup Needed, Followup Needed, Followup Needed - ASAP, Followed up / Resolved, Skipped / Newer Interaction, or Waiting for Response. Fills in No Followup Needed if blank. |
| Additional Notes | What happens next. |
| Hide | Tick it to leave the row out of Combined Contact Tracking and the phone contacts. It still counts for the PCAR. |

A few things happen on their own:

- When you type a name or OSIS, the rest of the row's student details fill in, for
  siblings too. If the tracker can't find the name, a message says so. Pasting a whole
  column of names or OSIS numbers works too.
- Date, # and Follow-up fill in only where you left them blank. Anything you type stays.
- A blank row is always kept above the dark END bar. The checkbox on the END bar
  switches a filter that shows only this week and last week (or only today, see
  Settings) plus open follow-ups.

Below the END bar is a search box. Type a student's OSIS or name to see their
communications before you call.

If you get "Could not find … in the Student List", that student isn't in the sheets yet.
Run 🏗️ Build Sheets Only and try again.

### Events

Use the Events tab for workshops, conferences and other group events, so they don't
crowd the Contact Log. Click 📆 Events → 📋 New Event, fill in the name, type, location
and date, and the tracker adds an attendance table. Enter each family by OSIS, student
name or guardian name; you can paste a whole column, or several OSIS numbers in one cell
separated by commas. Attendance shows in each student's Combined Contact Tracking and in
the PCAR.

### Settings

Open 🚀 App Menu → ⚙️ Settings. Choose a section on the left, then click Save & apply.

| Section | Setting | What it does |
|---|---|---|
| Contacts & sync | Sync contacts to my phone every night | Puts the school directory in your Google Contacts during the nightly run. |
| Tabs & columns | Show these tabs | Which tabs show. Contact Log, Combined Contact Tracking, PCAR and Directory show at first. |
| Tabs & columns | Contact Log columns | Shows or hides Site & Class and Type in the Contact Log. |
| Live updates | Rebuild sheets | After every edit and nightly (the default), nightly and when you click Build, after every edit but not nightly, or only when you click Build. No nightly run also means no nightly phone sync. |
| Live updates | Contact Log notes reach the other tabs | As you type (the default), or only when you build. |
| Live updates | Fill in Date, # and Follow-up for me | The Contact Log auto-fill. The name and OSIS lookup works either way. |
| Notes | Contact Log shows | This week and last week (the default), or today only. Open follow-ups always show. |
| Notes | Compiled notes cover | All time, the last month, 6 months or year, or your own dates. |
| Notes | Include notes I've hidden | Puts rows you ticked Hide on back into the compiled notes. |
| Student data | Hide discharged students in the Student List | Hides those rows. Turn it off to see them again. |
| Student data | Remind me to update my data | A reminder when an export is older than a week (the default), 2 weeks or a month. Or never. |
| Student data | Purge discharged students | Deletes discharged students from the Student List for good. |

Reset to defaults puts every setting back. Your settings are saved in the sheet, and
Import Old Data brings them into next year's copy.

### Build and Full Sync

🏗️ Build Sheets Only works inside the spreadsheet: it rebuilds the Student List,
Directory, Combined Contact Tracking and the other tabs from your data.

🧱 Full Sync does a build and then updates your Google Contacts. It keeps a "School
Directory" label in your contacts matching the Phone Contacts tab: it adds new parents,
updates changed ones and removes those who left. A side panel shows the progress. You
can close it; the sync keeps going. Large schools may take more than one run, and it
picks up where it stopped. With the phone sync on, the nightly run does the same.

## Part 3: New versions and the end of the year

### Bring over last year's data

Choose "No, bring over last year's tracker" in the window after Initial Setup, or open
🚀 App Menu → 🔄 Import Old Data at any time. The wizard lists the trackers it finds in your
Google Drive: click yours, or paste its link if it isn't listed, then click Bring it over.
It goes through six steps:

1. Connect to the old sheet.
2. Bring over your settings.
3. Copy the ATS Raw Data.
4. Copy events, student changes and your lists.
5. Match up the Contact Log, column by column.
6. Rebuild the directory and the other tabs.

Your old sheet is only read, never changed. Keep it until you have checked the new one.

### Updates

When you open the sheet, the tracker checks at most every few hours whether a newer
version is out, and shows what changed. You can also check with 🚀 App Menu → 🔔 Check
for Updates. In that window you can dismiss the update for now, dismiss it until a
newer one comes out, ask to hear about major updates only, or stop the automatic check.

An update doesn't change your current sheet. To move to a new version, click Get the
latest version, make a copy, run Initial Setup, and choose "No, bring over last year's
tracker". Your data, lists and settings come with it.

### End of Year Rollover

When the school year is over, run 🚀 App Menu → 🔄 End of Year Rollover. It:

- saves a PDF archive of your Contact Log, Combined Contact Tracking and your lists in
  the same Drive folder as the tracker,
- turns off the nightly sync so the old sheet can't change next year's contacts,
- gives you the link to next year's tracker. Make a copy, run Initial Setup, and choose
  "No, bring over last year's tracker".

Keep the old sheet and the PDF as your record for the year.

## Part 4: Privacy

The person who made the tracker can't see your sheet, your students, your contacts or
your Drive. The tracker has no server or database of its own. The script runs inside
your Google account, as you.

- When you allow the script, you give the permission to the script in your own account.
  Nobody else receives anything.
- The tracker reads and writes your own spreadsheet, and your own Google Contacts only
  if you turn on the phone sync.
- You can read the code: Extensions → Apps Script in the sheet.
- You can remove access at any time: myaccount.google.com → Security → Your connections
  to third-party apps & services.

The "Google hasn't verified this app" screen appears for tools like this one that are
not published in Google's Marketplace. Part 1, Step 2 lists each permission.

These are the only times the tracker looks outside your sheet:

- The update check reads a shared, view-only Version Tracker sheet. It sends nothing
  about your students.
- Import Old Data lists the trackers in your Drive by name, and opens only the one you
  pick.
- A bug report sends only what you type into the form.

Who else can see your copy is up to you. It is in your Drive, and you decide whether to
share it.

## Part 5: Help

### A typical day

1. Before a call, look the student up in the search box below the Contact Log's END bar.
2. After a contact, type the name or OSIS in a new row and fill in the notes and the
   follow-up.
3. To fix a name or phone, or add a missing student, use 👤 Student Overrides → ✨ New /
   Edit Student.
4. When you download a new export, load it with 📥 Input Data Mode.
5. For a big batch of entries, set Settings → Live updates → Rebuild sheets to "Only when
   I click Build", then build once at the end.

### Tips

- Siblings are linked by OSIS. To log one contact for several siblings, pick the entry
  with all their names (for example "Louise & Gene Belcher").
- Type your own date, # or follow-up and the auto-fill leaves it alone.
- Hide discharged students only hides them. Purge deletes them for good.
- If a menu looks out of date, reload the page.
- The data reminder appears once a day when an export is older than your setting. Change
  it in Settings → Student data.

### Something isn't working

| Problem | What to do |
|---|---|
| The tracker's menus don't appear | Wait a few seconds and reload the page. On a new copy, run 🚨 Initial Setup first. |
| You only see the Welcome tab | Setup hasn't run on this copy. Click 🚀 App Menu → 🚨 Initial Setup (Run Once). |
| Google keeps asking for permission | Setup needs every permission. Tick Select all, then Continue. |
| "Action Blocked" when you type | Setup hasn't run yet. Run 🚨 Initial Setup (Run Once). |
| The directory looks wrong after a paste | Load the export again with 📥 Input Data Mode, or paste values only, then build. |
| "Could not find … in the Student List" | Run 🏗️ Build Sheets Only, then try again. |
| "Background Sync in Progress" | A build or sync is running. Wait for it to finish before entering data. |
| The update check misses a new version | Use 🔔 Check for Updates. If it still says you're up to date, report it. |
| The phone sync stopped or made duplicates | Open 🧱 Full Sync and click Restart run. The next full run tidies the label. |
| A list row says "not found" | That person isn't in the Student List. Add them with ✨ New / Edit Student, build, then open the list again from Input Data Mode, or pick the student by hand. |
| Anything else | 🚀 App Menu → 🐞 Report a Bug, or email jacob.keller@p94m.org. |

### Things to avoid

- Typing in the tabs the tracker builds (Student List, Directory and so on). The next
  build overwrites it.
- Renaming or deleting the tracker's tabs, including the hidden Backend tab.
- Purging discharged students unless you are sure.
- Deleting your old tracker before you have checked the new one.

### Words used in this guide

| Word | Meaning |
|---|---|
| ATS / Insight | The DOE student system your Student Data export comes from, through TeachHub. |
| Build | Rebuilding the tabs the tracker makes from your data. |
| OSIS | A student's ID number. The tracker uses it to match students, siblings and parents. |
| Override | A change you make to a student's details without changing ATS. |
| PCAR | The Parent Coordinator Activity Report. The tracker fills in its counts. |
| Phone sync | Putting the school directory into your Google Contacts, under the "School Directory" label. |
| Your Lists | Your own lists (bus routes, PTA and so on), matched to students or guardians. |
| Additional Information | The lines from Your Lists on a phone contact. |
