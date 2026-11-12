# 🚀 Parent Coordinator Tracker — Complete Setup & User Guide

**Set-up · Buttons & Features · Migration & Updates · Privacy**

> A Google Sheets tool that turns your ATS export into a live directory, tracks every
> family contact, and compiles your PCAR — automatically.
>
> Made by a fellow Parent Coordinator — this is **NOT** an official DOE sheet.
> Questions or errors? jacob.keller@p94m.org

> **About these files:** This Markdown file is the editable source of the guide.
> `docs/PC_Tracker_User_Guide.docx` is the ready-to-use Word version — drag it into
> Google Drive to open it as a Google Doc. `docs/PC_Tracker_User_Guide.pdf` is a
> read-only copy with clickable "Jump to" buttons. `docs/guide.html` renders the PDF,
> and `docs/build_guide.js` regenerates the .docx. Keep all four in sync when editing.

---

## Contents

- [Get Your Copy](#get-your-copy)
- [What this tool is](#what-this-tool-is)
- [Part 1 — Setting Up Your Sheet](#part-1--setting-up-your-sheet)
- [Part 2 — Buttons & Features](#part-2--buttons--features)
- [Part 3 — Migration & Updates](#part-3--migration--updates)
- [Part 4 — Security & Privacy](#part-4--security--privacy)
- [Tips, Workflow & Troubleshooting](#tips-workflow--troubleshooting)
- [Getting this into Google Drive](#getting-this-into-google-drive)

---

## Get Your Copy

Start from one of the two templates below. Each opens a **"Make a copy"** prompt, so you
get your own private sheet in your own Google Drive.

- 📄 **New Coordinator — Make a Blank Copy** → *(paste the current (BLANK) copy link)*
- 🔄 **Upgrading — Make a Migrate Copy** → *(paste the current (MIGRATE) copy link)*

| Template | Who it's for |
|---|---|
| 📄 Blank Copy | Brand-new coordinators starting fresh — paste your ATS export and go. |
| 🔄 Migrate Copy | Coordinators upgrading from an older version — it auto-opens the Migration Wizard to bring last year's data and settings forward. |

> **Coordinator note — update these two links each release.** Both templates live in this
> Drive folder:
> <https://drive.google.com/drive/folders/1goKFrxDJOSAPxp1HUXETX2dns5c6HXk2>. Replace the
> two links above with the current **(BLANK)** and **(MIGRATE)** "make a copy" links
> whenever you ship a new version. Inside the sheet, the Welcome dialog and End-of-Year
> Rollover resolve these automatically from the folder — this guide is the one place you
> swap them by hand.

---

## What this tool is

The Parent Coordinator (PC) Tracker is a **Google Sheets add-on** built with Google
Apps Script. You paste your student exports into one tab, and the tool automatically
builds and maintains everything else: a clean class directory, a per-student contact
log, PCAR reporting numbers, and — if you want it — a copy of the directory pushed into
your phone through Google Contacts.

It's organized into three kinds of sheets:

- **The Engine** — Raw Data and the Backend tab's Type List, where you paste your system exports and
  customize your dropdowns.
- **The Workspaces** — Contact Log and Events, where you'll spend most of your day
  logging calls and workshops.
- **The Output** — automated sheets like Directory, Student List, Combined Contact
  Tracking, Parents Divided, PCAR, and Phone Contacts. You never type into
  these — they build themselves.

> **New here?** Do the one-time setup in Part 1, load your data, run Build, and you're
> running in a few minutes.

---

## Part 1 — Setting Up Your Sheet

### Quick Start (4 steps, about 10 minutes)

1. **Make your own copy** of the template (it opens a "Make a copy" prompt, or use
   *File → Make a copy*). Your copy opens on the **👋 Welcome** tab, the only tab showing.
2. **Read what Google will ask.** The Welcome tab (and Step 2 below) explains each permission
   before you click anything.
3. **Click 🚀 App Menu → 🚨 Initial Setup (Run Once)** and allow it: **OK** → your account →
   **Advanced** → **Go to … (unsafe)** → tick **Select all** → **Continue**.
4. **Follow your ✅ Getting Started checklist** (it opens on the right): Load your student data →
   Build your sheets → Check your settings → Log your first parent contact.

> **Moving from last year's tracker?** Run Initial Setup first, then use **🚀 App Menu → 🔄
> Import Old Data** (an upgrade copy opens it by itself). See Part 3.

### Step 1 — Make a copy and open it

You always start from a shared template. Making a copy drops a fresh, private
spreadsheet into your own Google Drive. The template is never touched by your work, and
no one else can see your copy.

Your copy opens on the **👋 Welcome** tab. Until setup is done it is the only tab showing:
it explains each permission Google will ask for, lists the setup steps, and links to this
guide and the bug report form. A message at the bottom of the screen also points you to
**🚀 App Menu → 🚨 Initial Setup (Run Once)**. Initial Setup removes the Welcome tab and
brings your tabs back.

### Step 2 — Know what Google will ask (permissions)

Initial Setup asks Google for permission once. Google shows its own screens first (the
tool can't show anything of its own until you allow it), so the Welcome tab explains them
before you click. Here is the same list, in Google's words:

| Permission Google shows | What the tool uses it for |
|---|---|
| See, edit, create and delete all your Google Sheets spreadsheets | Builds and updates the tabs in this tracker. It also reads the shared Version Tracker sheet to tell you when a newer version is out, and opens last year's tracker when you use 🔄 Import Old Data, which is why it needs more than this one sheet. |
| See, edit, download and permanently delete your contacts | Only the phone sync (⚙️ Settings → Contacts & sync), which stays off until you turn it on. It only adds, changes and removes contacts in its own "School Directory" label. |
| See, edit, create and delete all of your Google Drive files | 🔄 End of Year Rollover saves a PDF archive next to your tracker. It also finds the newest User Guide and template links in the shared folder. |
| Connect to an external service | Only that end-of-year PDF: it downloads the PDF from Google's own export page. Nothing about your families is sent anywhere. |
| Allow this application to run when you are not present | The automatic parts: instant updates when you type in the Contact Log, the nightly rebuild (and phone sync, if on) at 10 PM, and the update check when you open the sheet. |
| Display and run third-party web content in prompts and sidebars inside Google applications | The tool's own pop-ups and side panels: Settings, Input Data Mode, Getting Started, the build progress. "Third-party" just means not made by Google. |

> ✅ **Bottom line on permissions:** Every permission above works only inside your own
> Google account. There is no server or database belonging to the developer — see Part 4.

### Step 3 — Run Initial Setup and allow it

1. Wait until **🚀 App Menu** appears at the top of the screen, next to Help. It takes a
   few seconds; not there after a minute? Reload the page.
2. Click **🚀 App Menu → 🚨 Initial Setup (Run Once)**.
3. **Authorization required:** click **OK**, then choose your Google account.
4. **Google hasn't verified this app:** click **Advanced**, then **Go to … (unsafe)**.
   "Unverified" means Google hasn't done a formal brand review of a tool made by a person —
   *not* that anything is wrong.
5. On the list of permissions, tick **Select all**, then click **Continue**. Setup needs
   all of them; if one is left unticked, Google asks again for the missing ones.
6. Nothing happened after that? Click **🚀 App Menu → 🚨 Initial Setup (Run Once)** again.

What Initial Setup does:

- Installs the automatic triggers: instant Contact Log updates, the nightly rebuild/sync
  (around 10 PM) and the update check when you open the sheet.
- Records this copy as your original (so a later copy of it starts clean).
- Removes the Welcome tab and shows your everyday tabs: Contact Log, Combined Contact
  Tracking, PCAR and Directory (turn on more in ⚙️ Settings → Tabs & columns).
- Opens the welcome pop-up and your **✅ Getting Started** checklist.

> 🔒 **Until setup is done, the sheet is intentionally locked.** If you type before
> running Initial Setup, you'll see "🔒 Action Blocked — Initialization process needs to
> occur first." That's expected.

### Step 4 — Follow the Getting Started checklist

The checklist opens on the right after setup and ticks itself off as you go:

1. **Load your student data** — opens 📥 Input Data Mode (Step 5).
2. **Build your sheets** — Input Data Mode does this for you (Step 6).
3. **Check your settings** — opens ⚙️ Settings (Part 2).
4. **Log your first parent contact** — takes you to the Contact Log (Part 2).

It opens by itself each time you open the sheet until all four are done (untick **Open
this by itself** to stop that), and you can always reopen it from **🚀 App Menu → ✅ Getting
Started**. Under the steps, **This tab** says what the tab you're on is for, and **Hidden
tabs** lists the tabs you can turn on in ⚙️ Settings → Tabs & columns.

### Step 5 — Load your data (📥 Input Data Mode)

Before logging calls, the system needs to know who your students are. Open **🚀 App Menu →
📥 Input Data Mode** (or **Open Input Data Mode** in the checklist) and the sheet walks you
through the three exports one at a time, then two optional steps: **Site / Class** (each
class's teacher and site) and **Your Lists** (below). Your other tabs are hidden while you
work. For each export, either click **📂 Upload file** and pick the file you downloaded, or
open the file, press Ctrl+A (Mac: ⌘+A), copy (Ctrl+C / ⌘+C), and paste anywhere on the
temporary tab with Ctrl+Shift+V (Mac: ⌘+Shift+V). The data moves into Raw Data by itself,
and the sidebar shows when each one was last updated. It warns you, without changing Raw
Data, if you pasted the wrong export or if Excel cut numbers off (like 7.19E+09). Use
**Skip** to keep the old data for an export. Before the Your Lists step, it builds your
sheets for you. **Finish or exit** ends it at any step: **Save, exit and build my sheets**
saves what's on the tab, brings your tabs back and builds. The end screen offers **Build
sheets now** only if something changed since that build, and **Next steps** back to your
Getting Started checklist.

**Student Data (ATS / Insight) — required**

1. Go to TeachHub and sign in.
2. In the search bar, look up and open Insight.
3. Go to Tools → Downloader.
4. Under "Data Available for Download," select only **Guardian Contact Information**,
   then download. Upload or paste it in Input Data Mode.

**Parent Square — optional**

1. Go to ParentSquare, sign in, and open Admin.
2. Under Data Assistant, select Parents.
3. Click Export CSV. Upload or paste it in Input Data Mode.

**NYCSA — optional**

1. Go to TeachHub and sign in.
2. Open Family Access Management.
3. Export the "Students with Non-NYCSA Account" report. Upload or paste it in Input Data
   Mode.

**Pasting straight into Raw Data instead.** You can also paste each export into its
colored box on the **Raw Data** tab (show it in ⚙️ Settings → Tabs & columns): Student
Data into cell **A5**, Parent Square into **AG6**, NYCSA into **AP2**.

> 🚨 **Pasting into Raw Data? Always paste values only.** Use *Edit → Paste Special →
> Values Only* (Ctrl+Shift+V). Pasting normally can drag in formatting or formulas that
> break the build. (Input Data Mode does this for you.)

**Your own lists (Input Data Mode, Your Lists step).** Have a list that isn't an ATS
export, like bus routes, after-school sign-ups, PTA members or trip forms? On the **Your
Lists** step, click **➕ Add another list**, name it, and paste the whole sheet (headers
too) into the yellow box on the new **📥 (name)** tab, or upload the file (CSV or Excel).
The sidebar then takes you through
four short steps:

1. **Columns.** It finds the header row by itself (title rows above it are fine) and
   guesses what each column is, shown like *OSIS: Column C “OSIS”*. Fix a guess with its
   dropdown, and say whether each row is a **student** or a **guardian**. A name can be
   one column (“LAST, FIRST” is fine) or two (first + last).
2. **Variables.** The other columns (a route, a fee, a form status). Tick **keep**, give
   it a name, and choose what a blank cell shows and what everyone who isn't on the list
   gets: nothing, or your own text (like “No bus”). Then choose where it shows: under
   **Additional Information** on the phone contact (above Communications), and/or as a
   column in the **Student List** (student lists) or **Parents Divided** (guardian lists).
3. **Match.** It works out who each row is: OSIS first, then names, then phone or email,
   then best guesses for short names and nicknames (Sam → Samuel). Rows it isn't sure of
   are listed for you: confirm a best guess, pick one of several possible students, search
   the whole Student List, or skip the row. **✓ Accept all best guesses** takes them all
   at once; check a couple first, because best guesses can be wrong.
4. **View.** How the kept values read on the phone contact and in the Student List /
   Parents Divided: one per line or all on one line, names in bold or not, with a live
   preview. Click **Save** here.

Your picks are remembered: when you get a new version of the list, click **Paste new
version** and only the new rows need you. The list's tab shows who each row matched (✔
Matched / How). What you keep reaches the Student List, Parents Divided and the phone
contacts at the next build; Combined Contact Tracking doesn't change. Add as many lists as
you like.

> ⚠️ **Anything shown under Additional Information goes into your Google Contacts.** It
> is on the guardian's contact (and your phone) after the next sync. Keep private lists
> (IEP, housing, health) off, or keep them on their own tab only.

**Missing or wrong data? Use the New/Edit Student override.** District databases like
ATS can lag behind real enrollment. If you need to add a student or fix contact info
before it hits ATS, use **👤 Student Overrides → New / Edit Student**. When you build,
your overrides are applied to the Student List, Combined Contact Tracking, Parents
Divided, Phone Contacts, and the Ghost Typist. The **Directory** tab and the build
report's new / discharged students always show ATS (Raw Data) only. Overrides also tidy themselves up: a "New Student"
override is removed once ATS has that student, an edit that ATS now matches exactly is
removed too, and a student's class always follows Raw Data once they're in ATS.
(More in Part 2.)

**Customizing dropdowns — the Type List.** The optional **Type** column in the
Contact Log is driven by the Type List in the hidden **Backend** tab (column BL). Unhide Backend and edit that column to set your own
conversation-topic categories (for example: Attendance, Behavior, Enrollment, IEP).

### Step 6 — Build the sheets

Input Data Mode builds your sheets for you. To build again later (after new data, overrides
or lists), run **🚀 App Menu → 🏗️ Build Sheets Only**. A progress sidebar generates every derived tab and checks off each step as it
goes (reading your data, Student List, Event Log, Combined Contact Tracking, Parents
Divided, Directory, Phone Contacts, finishing up), with a line under each saying what it
builds. When it finishes, the sidebar shows a short report — new students, discharged
students, and families newly signed up for NYCSA or ParentSquare (click a count to see the
names). The *first* successful build adds the full menus right away: ⚙️ Settings, 👤
Student Overrides and 📆 Events (don't see them? reload the page with F5).

> ✅ **The golden rule of this sheet:** Raw Data is the only tab you paste into.
> Everything else is generated — so re-running Build always safely rebuilds it. Never
> hand-type into Student List, Directory, or other generated tabs.

---

## Part 2 — Buttons & Features

After setup, three custom menus sit at the top of the screen, to the right of "Help":
**🚀 App Menu**, **👤 Student Overrides**, and **📆 Events**.

### 🚀 App Menu

| Menu item | What it does |
|---|---|
| 🧱 **Full Sync (Build + Push to Google Contacts)** | Rebuilds every sheet **and** pushes the directory to your Google Contacts. Only appears when phone syncing is on. Opens a progress sidebar you can pause or stop. |
| 🏗️ **Build Sheets Only (No Push)** | Your everyday "refresh everything" button. Regenerates all derived tabs from Raw Data + Contact Log + overrides, without touching your phone contacts. |
| ⚙️ **Settings** | Show/hide tabs and columns, control live updates, filter notes, manage discharged students (full list below). |
| ✅ **Getting Started** | The four-step checklist for a new sheet (Part 1, Step 4). |
| 📖 **Open User Guide** | A short welcome with the next steps and a button that opens this guide. |
| 🐞 **Report a Bug** | Opens a short Google Form to report anything broken or confusing. |
| 📥 **Input Data Mode** | Walks you through loading each export, then Site / Class and Your Lists, one step at a time, and builds your sheets (Part 1, Step 5). |
| 🔔 **Check for Updates** | Manually checks whether a newer template has been released, and lets you set update preferences. |
| 🔄 **End of Year Rollover** | Archives the year to a PDF, freezes the sheet, and hands you next year's fresh template (Part 3). |

### 👤 Student Overrides

"Overrides" let you correct or add student and guardian info without editing the raw
ATS data.

| Menu item | What it does |
|---|---|
| ✨ **New / Edit Student** | Add a student who isn't in ATS, or correct a name, phone, email, language, grade, or up to 5 guardians (a student's class always follows Raw Data once they're in ATS). Also set a status **Label** and **"Hide from Report (Stealth Mode)."** Saving opens the Build sidebar and rebuilds the sheets (if rebuilds are paused in Settings, it just saves). |
| 🗂️ **Manage Overrides** | See every override you've made; edit or delete any of them. |
| 📄 **Override Report** | A side-by-side "Students Override Report" comparing raw ATS data against your overrides. Stealth-mode students are skipped. |

### 📆 Events

Use the Events tab for large-scale attendances (workshops, graduations, fairs) so they
don't clutter your daily Contact Log. Click **📆 Events → 📋 New Event**. A pop-up with a
calendar picker asks for the event name, type (buttons for PTC, Workshop, Family event and
Meeting fill in the Type box), and location, then builds an attendance table. Log
attendees by OSIS, student name, or guardian name — type one, paste a whole column, or
put several OSIS numbers in one cell separated by commas — and the Ghost Typist fills in
the rest. Attendees map back to each student's Combined Contact Tracking and tally
automatically in PCAR.

### The tabs — what each one is for

| Tab | Purpose | Visible by default? |
|---|---|---|
| **Raw Data** | Your export paste zone — the only tab you paste raw data into (ATS / ParentSquare / NYCSA). | Hidden (show in Settings) |
| **👋 Welcome** | The first page a new copy opens on, with the permissions and the setup steps. Removed by Initial Setup. | Only before setup |
| **Contact Log** | The heart of the tool. Log every guardian contact here (deep dive below). | Yes |
| **Student List** | The generated roster (called “Master Table” in older copies): OSIS, student, guardian, site, class, siblings, status, the date of the last interaction, and a column for each student list you keep (Your Lists). Data starts on row 5. | Hidden (generated) |
| **Directory** | A clean, printable class directory built from your Raw Data (ATS): each student's site, class/teacher, and parents' email and phone. Overrides and discharged students aren't included. | Yes (generated) |
| **Combined Contact Tracking** | One comprehensive history block per student — all notes and event attendances together. Great before an IEP meeting. Only students on the Student List get a block, so notes logged under a mistyped or unknown OSIS don't show here. | Yes (generated) |
| **PCAR** | Your auto-compiled PCAR numbers (Phone, Walk-In, Workshops, PTCs, totals). | Yes (generated) |
| **Phone Contacts** | The exact data pushed to Google Contacts, including the Additional Information lines from Your Lists. | Hidden (optional) |
| **Parents Divided** | One row per guardian (with ParentSquare / NYCSA flags, and a column for each guardian list you keep). Handy for mail merges. | Hidden (optional) |
| **Notes** | An optional manual notes tab (one row per family) with its own auto-fill and duplicate protection. Its notes show at the top of that student's Combined Contact Tracking. | Hidden (optional) |
| **Send Out** | Optional tab tied to extra PCAR rows. | Hidden (optional) |
| **📥 Your lists** | One tab per list you add in Input Data Mode (Your Lists step): your pasted list, plus ✔ Matched / How columns showing who each row is. | Yes (yours) |
| **Backend** | Behind-the-scenes sections side by side: Version & saved settings (A:E), New/Edit Student overrides (G:AV), event log (AX:BF), Send Out list (BH:BJ), the Type List that drives the Contact Log Type dropdown (BL), Additional Information (BN:BQ, rewritten by every build) and the Your Lists setups (BS:BV). Only edit the Type List. | Hidden (system) |

### The Contact Log, in depth

This is your primary workspace for phone calls, emails, and 1-on-1 meetings. It's
designed so you type as little as possible.

**The Pre-Call Search Bar.** Below the dark END bar at the bottom of the Contact Log is a
search bar. Before you call a family, type a student's name or OSIS to pull up their
Communications (every note so far) — so you walk into the call with context.

**Logging a new interaction (left → right):**

| Column | What to put there |
|---|---|
| Date | The date of the interaction. Auto-fills to today if blank; type your own to back-date. |
| Guardian / OSIS / Student | Type ONE of three things — the parent's name (column C), the OSIS (E), or the student's name (F). The Ghost Typist fills in the other two. |
| Person Spoke With | The specific parent/guardian you actually spoke to. |
| # | Number of interactions, for PCAR counting. A normal call = 1; a mass email to 30 parents = 30. Auto-fills to 1 if blank. |
| Site & Class | (Optional) Auto-fills from the student's profile; can be hidden in Settings. |
| 📱🚶 Category | Categorize for PCAR — e.g., Phone, Walk-In, Parent-Teacher. |
| Method | The exact method — Call, Email, In School, etc. |
| Type | (Optional) A custom topic dropdown driven by the Backend tab's Type List; can be hidden in Settings. |
| Event Description | Detailed notes about the conversation. |
| Follow-up | No Followup Needed · Followup Needed · Followup Needed - ASAP · Followed up / Resolved · Skipped / Newer Interaction · Waiting for Response. Auto-fills to "No Followup Needed" if blank. |
| Additional Notes | Next steps or what you did after the initial contact (e.g., "Teacher will reach out"). |
| Hide | Check this box to keep this note out of the compiled notes (Combined Contact Tracking and Phone Contacts). The row stays in the Contact Log and still counts for PCAR. |

Three things the Contact Log does for you automatically:

- **Ghost Typist.** Type — or paste a whole column of — OSIS numbers, student names, or
  guardian names, and the other identity columns (plus Site and Class, including for
  siblings) fill in for every row. If it can't match, it warns "Could not find … in the
  Student List" (after a paste, it lists the ones it couldn't find).
- **Smart auto-fill.** Entering real content stamps today's Date, sets # to 1, and
  Follow-up to "No Followup Needed" — but only where you left them blank. Anything you typed yourself is
  never overwritten.
- **Auto-adding rows + the master hide filter.** A fresh blank row is always kept above
  the dark "END" bar. The checkbox on the END-bar row turns on a filter that shows only
  this and last week's rows (or only today's — see Settings), plus anything still marked
  for follow-up.

> **If the Ghost Typist says "Could not find … in the Student List":** That family isn't
> compiled yet. Run **🏗️ Build Sheets Only** once, then try again.

### ⚙️ Settings — every choice explained

Settings has five sections on the left. Click **Save & apply** to keep your changes.

| Section → setting | What it controls |
|---|---|
| **Contacts & sync → Sync contacts to my phone every night** | Adds the school directory to your Google Contacts during the nightly run (and the 🧱 Full Sync menu item appears). |
| **Tabs & columns → Show these tabs** | Show or hide any tab (every tab except Backend). Contact Log, Combined Contact Tracking, PCAR and Directory start shown; the rest start hidden. Hidden tabs keep updating. |
| **Tabs & columns → Contact Log columns** | Shows / hides the Site & Class and Type columns in the Contact Log. |
| **Live updates → Rebuild sheets** | When the Student List, Directory and other tabs refresh: after every edit and nightly (default), nightly and when you click Build, after every edit but not nightly, or only when you click Build. Turning off the nightly run also stops the nightly phone sync. |
| **Live updates → Contact Log notes reach the other tabs** | As you type (default), or only when you click Build. |
| **Live updates → Fill in Date, # and Follow-up for me** | The Contact Log auto-fill. The parent / student / OSIS lookup works either way. |
| **Notes → Contact Log shows** | This week and last week (default), or today only. Important notes and open follow-ups always show. |
| **Notes → Compiled notes cover** | All time, the last month / 6 months / year, or custom dates. |
| **Notes → Include notes I've hidden** | Adds rows you checked off with the row checkbox back into the compiled notes. |
| **Student data → Hide discharged students in the Student List** | Only hides the rows. Turn it off to see them again. |
| **Student data → Remind me to update my data** | A pop-up when Student Data, Parent Square or NYCSA hasn't been updated in a week (default), 2 weeks, a month, or never. |
| **Student data → Purge discharged students** | **Permanently** deletes non-current students from the Student List. You can't undo this. |
| **Reset to defaults** | Returns every setting to its default. |

> ✅ **Settings travel with your copy.** Your choices are saved into the hidden Backend
> tab, so a future upgrade copy inherits them automatically.

### Build vs. Full Sync — know the difference

| Command | What it is |
|---|---|
| 🏗️ **Build / Rebuild Data** | An **internal** action. Takes your Raw Data and constructs all the automated sheets (Directory, Student List, Parents Divided, Phone Contacts, PCAR). It only organizes data inside the spreadsheet. |
| 🧱 **Full Sync with Google Contacts** | An **external** action. Pushes the formatted Phone Contacts tab into your actual Google Contacts via a "School Directory" label — one number per card — so a parent's name, relationship, and student show on caller ID. |

The phone sync runs in "mirror" mode — creating, editing, and removing contacts so the
label always matches your directory. It runs nightly or on demand via **🧱 Full Sync**,
with a progress sidebar you can pause, stop, or restart; it's resilient and resumes if
it hits Google's time limits.

---

## Part 3 — Migration & Updates

### Importing data from an older sheet (Migration Wizard)

Open **🚀 App Menu → 🔄 Import Old Data**. (On an "upgrade" copy it opens by itself.)
Paste the link to your older PC sheet and click **Start Safe Migration**. The wizard
runs six steps:

1. Connect to the old sheet (confirms the link works and you have access).
2. Recover your custom settings (inferred from which tabs were visible in the old sheet).
3. Migrate your ATS Raw Data.
4. Transfer your Events, Send Out, and Student Overrides.
5. Smart-map the Contact Log — it matches columns by their headers, so it still works
   even if the layout changed, and strips blank/"ghost" rows.
6. Finalize and rebuild all directories.

> ✅ **Migration is one-directional and non-destructive.** It reads your old sheet and
> writes into the new one. Nothing is deleted from the old sheet — keep it as a backup.

### Staying current (the update checker)

Each copy remembers its own build date. Quietly, on open (at most once every few hours),
the sheet checks a shared **Version Tracker** to see whether a newer template has
shipped. If so, it pops a short changelog with a link to the newest copy. You can also
check any time with **🚀 App Menu → 🔔 Check for Updates**.

Inside that dialog you can:

- **Dismiss for now** — close it; it pops up again on the next check.
- **Dismiss this update** — hide this update until something even newer ships.
- **Major updates only** — automatic pop-ups only for big releases (a manual check still
  shows everything).
- **Stop automatic update checks** — silence the on-open checks; the manual menu check
  always still works.

**How to actually adopt an update.** Updates don't rewrite your current sheet in place.
To move up a version: open the new template link, **make a copy**, run Initial Setup,
then use the **Migration Wizard** to bring your data across. Your settings come along
automatically.

### End of Year Rollover

When the year ends, run **🚀 App Menu → 🔄 End of Year Rollover**. It:

- Generates a permanent PDF archive of your Contact Log and Combined Contact Tracking,
  saved next to your sheet in Drive.
- Freezes the sheet — turns off nightly and phone sync so it won't overwrite next year's
  data — and renames it "… – Retired on [date]."
- Hands you the link to next year's fresh, fast template.

> **Rollover is safe to run once you're done for the year.** It archives and freezes the
> OLD sheet; you then set up a brand-new copy for the new year. Keep the retired sheet
> and its PDF as your record.

---

## Part 4 — Security & Privacy

> ✅ **The short version: I cannot see your sheet, your students, your contacts, or your
> Drive.** This tool has no server and no database of mine. The code runs entirely inside
> your own Google account. Nothing about your families is ever sent to me or anyone else.
> Once you make a copy, it's yours.

### Why that's true

- **It runs as you, on Google's servers.** When you click Authorize, you grant permission
  to the script — running inside your account — not to me. I receive nothing.
- **There's no outside destination for your data.** The tool reads and writes your own
  spreadsheet and (optionally) your own Google Contacts. It doesn't phone home.
- **You can read every line of the code.** Open *Extensions → Apps Script* in the sheet.
- **You can revoke access anytime.** Go to *myaccount.google.com → Security → Your
  connections to third-party apps & services*, and remove the tool.

### Why the "unverified app" warning appears

That screen shows for personal Apps Script tools that aren't published to Google's
Marketplace. It means Google hasn't run a formal brand review — *not* that the tool is
unsafe. For a permission-by-permission breakdown, see **"What each authorization actually
does"** in Part 1.

### The only things that ever leave your sheet

- **The update check** reads a shared, view-only "Version Tracker" spreadsheet to compare
  dates. It sends nothing about your students.
- **Migration** opens an older sheet only when you paste its link, and only because you
  already have access to it.
- **Bug reports** go through a Google Form you choose to fill out — only the words you type
  are shared.

> **Who can see your copy?** Entirely up to you. Once you make a copy it lives in your
> Drive; you decide who (if anyone) gets shared access.

---

## Tips, Workflow & Troubleshooting

### A good daily workflow

1. Before a call, use the Contact Log search bar to pull up the family's history.
2. Made a contact? In a new row, type the OSIS or a name and let the Ghost Typist fill
   the rest. Add notes and a follow-up flag.
3. Need to fix a name/phone or add a missing student? Use **👤 Student Overrides → New /
   Edit Student**.
4. Fresh ATS export? Paste it into Raw Data (Paste Special → Values Only), then run
   **🏗️ Build Sheets Only**.
5. Big batch? Turn on "Rebuilds: Manual / Nightly Sync Only," then run Build once at the
   end.

### Handy tips

- **Siblings are automatic.** The tool joins siblings by OSIS and fills Site/Class for each.
- **Back-dating works.** Type your own date, count, or follow-up status and the auto-fill
  leaves it alone.
- **Hide vs. Purge discharged.** "Hide Discharged" is a reversible view filter; "Purge"
  permanently deletes.
- **Refresh if a menu looks stale.** Pressing F5 always reloads the latest menu.
- **Fresh-data reminder.** If your ATS, ParentSquare or NYCSA data in Raw Data is older
  than your reminder setting (a week by default), the sheet shows a reminder once a day
  when you open it. Change it or turn it off in ⚙️ Settings → Student data.
- **Lists match best with OSIS.** If a list you add has an OSIS column, every row with an
  OSIS matches exactly; names alone need more checking.

### Troubleshooting / FAQ

| Symptom | Fix |
|---|---|
| The custom menus didn't appear | Wait a few seconds, then press F5. If it's a brand-new copy, run **🚨 Initial Setup** first. |
| I only see the Welcome tab | Setup hasn't been run on this copy yet. Click **🚀 App Menu → 🚨 Initial Setup (Run Once)**; it removes the Welcome tab and shows your tabs. |
| Google keeps asking for permission | Setup needs every permission on the list. Tick **Select all**, then **Continue** (what each one is for: Part 1, Step 2). |
| "🔒 Action Blocked" when I type | Setup hasn't been run yet. Click **🚀 App Menu → 🚨 Initial Setup (Run Once)**. |
| Pasted data looks broken / directory is wrong | Re-paste using **Edit → Paste Special → Values Only**, then run **🏗️ Build Sheets Only**. |
| Ghost Typist: "Could not find … in the Student List" | That family isn't compiled yet. Run **🏗️ Build Sheets Only**, then try again. |
| "📡 Background Sync in Progress" warning | A rebuild/sync is running. Wait for it to finish before entering new data. |
| Update checker says I'm up to date but there's a new version | Use **🔔 Check for Updates** (it bypasses the throttle). If still wrong, report a bug. |
| Phone sync created duplicates / didn't finish | Open Full Sync, use Stop then Restart. The sync mirrors the sheet and cleans up on the next full run. |
| A row on my list says “not found” | That student (or guardian) isn't on the Student List. A new student? Add them with 👤 New / Edit Student, build, then open the list again (📥 Input Data Mode → Your Lists → Edit setup). Or pick them by hand with 🔍 Pick from the Student List. |
| Something's broken or confusing | Use **🚀 App Menu → 🐞 Report a Bug**, or email jacob.keller@p94m.org. |

### Do & Don't

| ✅ Do | 🚫 Don't |
|---|---|
| Paste raw data only into Raw Data — as Values Only. | Type into Student List, Directory, or other generated tabs — they're overwritten on build. |
| Use Overrides to fix names, phones, or add students. | Rename or delete the generated / hidden tabs (including Backend). |
| Run Build after big changes. | Hand-edit the hidden Backend tab's saved settings. |
| Keep your old / retired sheet as a backup until the new one is verified. | Purge discharged students unless you're certain. |

### Glossary

| Term | Meaning |
|---|---|
| **OSIS** | The unique student ID number from ATS. Used to match students, siblings, and contacts. |
| **ATS / Insight** | The NYC DOE student data systems your Raw Data export comes from (via TeachHub). |
| **PCAR** | Parent Coordinator Activity Report — the reporting numbers this tool compiles. |
| **Ghost Typist** | The auto-fill that completes a Contact Log, Notes, or Events row from a single OSIS or name. |
| **Override** | A correction or addition to student/guardian info without changing raw ATS data. |
| **Stealth Mode** | An override flag that keeps a student out of the Override Report. |
| **NYCSA / ParentSquare** | Parent account/communication systems whose exports show each family's sign-up status in Parents Divided and Phone Contacts. |
| **Build** | Regenerating all derived tabs from your current data (**🏗️ Build Sheets Only**). |
| **Your Lists** | Your own lists (bus routes, PTA…) added in Input Data Mode's last step. Each row is matched to a student or guardian; what you keep shows under Additional Information and as Student List / Parents Divided columns. |
| **Additional Information** | The lines from Your Lists on a phone contact, above Communications. Kept in the Backend tab and rewritten by every build. |

---

## Getting this into Google Drive

You can keep and share this guide as a Google Doc in a couple of ways:

### Option A — Upload and convert (recommended)

1. Go to drive.google.com and open the folder where you want it.
2. Drag the **PC_Tracker_User_Guide.docx** file into the window (or click New → File
   upload).
3. Right-click the uploaded file → Open with → Google Docs. Drive converts it to a
   Google Doc; the Jump-to buttons and table of contents stay clickable.
4. (Optional) Delete the original .docx once the Google Doc is created.

### Option B — From inside Google Docs

1. Open docs.google.com and start a blank document.
2. Choose File → Open → Upload, and drop in the .docx.

### Prefer a clickable PDF?

The companion **PC_Tracker_User_Guide.pdf** has a "Jump to" button bar and a clickable
table of contents that jump to each section — handy for reading on a phone or sharing a
read-only copy.

---

*Questions or something broken? Use **🚀 App Menu → 🐞 Report a Bug**, or email
jacob.keller@p94m.org.*
