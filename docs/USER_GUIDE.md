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
>
> 📷 Screenshot placeholders are marked throughout as `📷 Screenshot: …` — drop in
> real images at those spots.

---

## Contents

- [What this tool is](#what-this-tool-is)
- [Part 1 — Setting Up Your Sheet](#part-1--setting-up-your-sheet)
- [Part 2 — Buttons & Features](#part-2--buttons--features)
- [Part 3 — Migration & Updates](#part-3--migration--updates)
- [Part 4 — Security & Privacy](#part-4--security--privacy)
- [Tips, Workflow & Troubleshooting](#tips-workflow--troubleshooting)
- [Getting this into Google Drive](#getting-this-into-google-drive)

---

## What this tool is

The Parent Coordinator (PC) Tracker is a **Google Sheets add-on** built with Google
Apps Script. You paste your student exports into one tab, and the tool automatically
builds and maintains everything else: a clean class directory, a per-student contact
log, PCAR reporting numbers, and — if you want it — a copy of the directory pushed into
your phone through Google Contacts.

It's organized into three kinds of sheets:

- **The Engine** — RAW Data and Type List, where you paste your system exports and
  customize your dropdowns.
- **The Workspaces** — Contact Log and Events, where you'll spend most of your day
  logging calls and workshops.
- **The Output** — automated sheets like Directory, Master Table, Combined Contact
  Tracking, Parents Divided, PCAR, and Phone Contacts. You never type into these — they
  build themselves.

> **New here?** Do the one-time setup in Part 1, load your data, run Build, and you're
> running in a few minutes.

---

## Part 1 — Setting Up Your Sheet

### Quick Start (5 steps)

1. **Make your own copy** of the template (it opens a "Make a copy" prompt, or use
   *File → Make a copy*).
2. Open your copy, wait for **🚀 App Menu** to appear, then click
   **🚀 App Menu → 🚨 Initial Setup (Run Once)**.
3. Authorize the script when Google asks (one-time — see below).
4. Load your data into the **RAW Data** tab, or use **🔄 Import Old Data** if upgrading.
5. Run **🏗️ Build Sheets Only**, then press **F5** to refresh — the full menu unlocks.

### Step 1 — Make a copy and open it

You always start from a shared template. Making a copy drops a fresh, private
spreadsheet into your own Google Drive. The template is never touched by your work, and
no one else can see your copy.

### Step 2 — Run Initial Setup (once)

Wait until the **🚀 App Menu** appears in the top toolbar — don't click anything else
until it shows. Then click **🚀 App Menu → 🚨 Initial Setup (Run Once)**. This does the
behind-the-scenes wiring:

- Authorizes the script to run as you.
- Installs the automatic triggers — instant Contact Log updates, a nightly refresh/sync
  (~11 PM), and a periodic update check.
- Records this copy as your original (a clone-detector so a later copy starts clean).
- Applies the default tab layout and opens this User Guide.

> 📷 Screenshot: The 🚀 App Menu open, showing "🚨 Initial Setup (Run Once)."

> 🔒 **Until setup is done, the sheet is intentionally locked.** If you type before
> running Initial Setup, you'll see "🔒 Action Blocked — Initialization process needs to
> occur first." That's expected.

### Step 3 — The authorization screen

Because this is a personal tool and not a paid Marketplace app, Google shows a warning
like *"Google hasn't verified this app."* To continue, click **Advanced → Go to …
(unsafe) → choose your account → Allow**.

"Unverified" means Google hasn't done a formal brand review — *not* that anything is
wrong. You are granting permission to a script that runs entirely inside **your own**
account.

> 📷 Screenshot: The Google consent screen — "Google hasn't verified this app" →
> Advanced → Go to … (unsafe) → Allow.

#### What each authorization actually does

On the Allow screen, Google lists the permissions below. The tool asks only for what it
needs — and everything runs inside your own account, so approving it does not give the
developer any access.

| Permission Google shows | What it's used for | Why it's needed |
|---|---|---|
| See, edit, create, and delete all your Google Sheets spreadsheets | Reading your RAW Data and writing the Master Table, Directory, Contact Log, PCAR, and every other tab. | This is the whole engine — without it the tool can't build or update your sheets. |
| See, edit, download, and permanently delete your contacts | The optional "School Directory" phone sync that mirrors parent numbers into Google Contacts. | Only used if you turn on phone syncing. If you never enable it, nothing is written to your contacts. |
| See, edit, create, and delete all of your Google Drive files | Saving the year-end PDF archive to Drive, and opening your old sheet during a migration. | Rollover creates a PDF; migration opens the old spreadsheet you point it at. |
| Connect to an external service | A read-only check of the shared "Version Tracker" sheet so you're told when a newer template ships. | The only outside call the tool makes. It reads release dates/links — never your data. |
| Run when you're not present / manage its own triggers | The automatic triggers: instant Contact Log updates, the nightly rebuild/sync, and the update check on open. | Lets the sheet keep itself current without you clicking Build every time. |
| Display and run third-party web content in dialogs and sidebars | The menus, wizards, and sidebars — Settings, Migration, Rollover, and the progress bars. | These custom screens are how you drive the tool. |

> ✅ **Bottom line on permissions:** Every scope above operates only within your own
> Google account. There is no server or database belonging to the developer — see Part 4.

### Step 4 — Load your data (the RAW Data tab)

Before logging calls, the system needs to know who your students are. Export data from
your standard systems and paste it into the brightly-colored "PASTE" boxes on the
**RAW Data** tab.

> 🚨 **CRITICAL — always paste as values only.** When pasting into RAW Data, use
> *Edit → Paste Special → Values Only* (Ctrl+Shift+V). Pasting normally can drag in
> formatting or formulas that break the build.

**Insight / ATS data — Blue section → paste into cell A5 (required)**

1. Go to TeachHub and sign in.
2. In the search bar, look up and open Insight.
3. Go to Tools → Downloader.
4. Under "Data Available for Download," select only **Guardian Contact Information**,
   then download.
5. Copy the data and paste it into **cell A5** (Paste Special → Values Only).

**ParentSquare data — Orange section → paste into cell AG6 (optional)**

1. Go to ParentSquare, sign in, and open Admin.
2. Under Data Assistant, select Parents.
3. Click Export CSV.
4. Copy the data and paste it into **cell AG6**.

**NYCSA data — Green section → paste into cell AP2 (optional)**

1. Go to TeachHub and sign in.
2. Open Family Access Management.
3. Export the "Students with Non-NYCSA Account" report.
4. Copy the data and paste it into **cell AP2**.

> 📷 Screenshot: The RAW Data tab showing the colored PASTE boxes (Blue A5 / Orange AG6 / Green AP2).

**Missing or wrong data? Use the New/Edit Student override.** District databases like
ATS can lag behind real enrollment. If you need to add a student or fix contact info
before it hits ATS, use **👤 Student Overrides → New / Edit Student**. When you build,
the script checks your overrides first and applies them across the entire directory.
(More in Part 2.)

**Customizing dropdowns — the Type List tab.** The optional **Type** column in the
Contact Log is driven by the **Type List** tab. Edit that tab to set your own
conversation-topic categories (for example: Attendance, Behavior, Enrollment, IEP).

### Step 5 — Build the sheets

Run **🚀 App Menu → 🏗️ Build Sheets Only**. A progress sidebar generates every derived
tab. The *first* successful build unlocks the full menu — a toast asks you to press
**F5** to refresh.

> ✅ **The golden rule of this sheet:** RAW Data is the only tab you paste into.
> Everything else is generated — so re-running Build always safely rebuilds it. Never
> hand-type into Master Table, Directory, or other generated tabs.

---

## Part 2 — Buttons & Features

After setup, three custom menus sit at the top of the screen, to the right of "Help":
**🚀 App Menu**, **👤 Student Overrides**, and **📆 Events**.

### 🚀 App Menu

| Menu item | What it does |
|---|---|
| 🧱 **Full Sync (Build + Push to Google Contacts)** | Rebuilds every sheet **and** pushes the directory to your Google Contacts. Only appears when phone syncing is on. Opens a progress sidebar you can pause or stop. |
| 🏗️ **Build Sheets Only (No Push)** | Your everyday "refresh everything" button. Regenerates all derived tabs from RAW Data + Contact Log + overrides, without touching your phone contacts. |
| ⚙️ **Settings** | Show/hide tabs and columns, control live updates, filter notes, manage discharged students (full list below). |
| 📖 **Open User Guide** | Opens this guide. |
| 🐞 **Report a Bug** | Opens a short Google Form to report anything broken or confusing. |
| 🔔 **Check for Updates** | Manually checks whether a newer template has been released, and lets you set update preferences. |
| 🔄 **End of Year Rollover** | Archives the year to a PDF, freezes the sheet, and hands you next year's fresh template (Part 3). |

### 👤 Student Overrides

"Overrides" let you correct or add student and guardian info without editing the raw
ATS data.

| Menu item | What it does |
|---|---|
| ✨ **New / Edit Student** | Add a student who isn't in ATS, or correct a name, phone, email, language, class, grade, or up to 5 guardians. Also set a status **Label** and **"Hide from Report (Stealth Mode)."** Rebuilds in the background (~1 min). |
| 🗂️ **Manage Overrides** | See every override you've made; edit or delete any of them. |
| 📄 **Override Report** | A side-by-side "Students Override Report" comparing raw ATS data against your overrides. Stealth-mode students are skipped. |

### 📆 Events

Use the Events tab for large-scale attendances (workshops, graduations, fairs) so they
don't clutter your daily Contact Log. Click **📆 Events → 📋 New Event**. A pop-up with a
calendar picker asks for the event name, type/category (with a **PTC Quick Fill** button
for "Parent Teacher Conference"), and location, then builds an attendance table. Log
attendees by OSIS — they map back to each student's Combined Contact Tracking and tally
automatically in PCAR.

### The tabs — what each one is for

| Tab | Purpose | Visible by default? |
|---|---|---|
| **RAW Data** | Your export paste zone — the only tab you paste raw data into (ATS / ParentSquare / NYCSA). | Yes |
| **Type List** | Drives the optional Type dropdown in the Contact Log. Edit it to set your own categories. | Yes |
| **Contact Log** | The heart of the tool. Log every guardian contact here (deep dive below). | Yes |
| **Master Table** | The generated master roster: OSIS, student, guardian, status, and more. Data starts on row 5. | Yes (generated) |
| **Directory** | A clean, printable class directory merging ATS, ParentSquare, and NYCSA data. | Yes (generated) |
| **Combined Contact Tracking** | One comprehensive history block per student — all notes and event attendances together. Great before an IEP meeting. | Yes (generated) |
| **PCAR** | Your auto-compiled PCAR numbers (Phone, Walk-In, Workshops, PTCs, totals). | Yes (generated) |
| **Phone Contacts** | The exact data pushed to Google Contacts. | Hidden (optional) |
| **Parents Divided** | One row per guardian (with ParentSquare / NYCSA flags). Handy for mail merges. | Hidden (optional) |
| **Notes** | An optional manual notes tab with its own auto-fill and duplicate protection. | Hidden (optional) |
| **Send Out** | Optional tab tied to extra PCAR rows. | Hidden (optional) |
| **Version / Backend_Event_Log** | Behind-the-scenes tabs (build date, saved settings, event data). Leave them alone. | Hidden (system) |

### The Contact Log, in depth

This is your primary workspace for phone calls, emails, and 1-on-1 meetings. It's
designed so you type as little as possible.

**The Pre-Call Search Bar.** At the top of the Contact Log is a search bar. Before you
call a family, type a student's name or OSIS to instantly pull up your most recent
interaction with them — so you walk into the call with context.

**Logging a new interaction (left → right):**

| Column | What to put there |
|---|---|
| Date | The date of the interaction. Auto-fills to today if blank; type your own to back-date. |
| Find Student | Type ONE of three things — the parent's name, the OSIS, or the student's name. The Ghost Typist fills in the rest of the profile. |
| Person Spoke With | The specific parent/guardian you actually spoke to. |
| # | Number of interactions, for PCAR counting. A normal call = 1; a mass email to 30 parents = 30. Auto-fills to 1 if blank. |
| Site & Class | (Optional) Auto-fills from the student's profile; can be hidden in Settings. |
| 📱🚶 Category | Categorize for PCAR — e.g., Phone, Walk-In, Parent-Teacher. |
| Method | The exact method — Call, Email, In School, etc. |
| Type | (Optional) A custom topic dropdown driven by the Type List tab; can be hidden in Settings. |
| Notes | Detailed notes about the conversation. |
| ⁉️ Follow-Up Needed | A visual flag: Yes · Yes – Important · No · Waiting on Someone · Completed. Auto-fills to "No" if blank. |
| Followup Notes | Next steps or what you did after the initial contact (e.g., "Teacher will reach out"). |
| Hide | Check this box to hide the row from view while keeping its data active for PCAR counting. |

Three things the Contact Log does for you automatically:

- **Ghost Typist.** Type an OSIS, student name, or guardian name and the other identity
  columns (plus Site and Class, including for siblings) fill in. If it can't match, it
  warns "Could not find … in the Master Table."
- **Smart auto-fill.** Entering real content stamps today's Date, sets # to 1, and
  Follow-up to "No" — but only where you left them blank. Anything you typed yourself is
  never overwritten.
- **Auto-adding rows + the master hide filter.** A fresh blank row is always kept above
  the dark "END" bar, and the checkbox on the END-bar row toggles hiding of your hidden
  rows.

> 📷 Screenshot: The Contact Log — the pre-call search bar up top and a row being auto-filled by the Ghost Typist.

> **If the Ghost Typist says "Could not find … in the Master Table":** That family isn't
> compiled yet. Run **🏗️ Build Sheets Only** once, then try again.

### ⚙️ Settings — every toggle explained

| Setting | What it controls |
|---|---|
| **Sync Contacts to Phones (Nightly)** | Master switch for phone sync. When on, the nightly run pushes the directory into Google Contacts (and the Full Sync menu item appears). |
| **Show "Phone Contacts" / "Parents Divided" / "Notes" / "Send Out"** | Show or hide each optional tab. All start hidden. |
| **Show Site & Class (Contact Log)** | Shows/hides the Site and Class columns in the Contact Log. |
| **Show "Type" (Contact Log)** | Shows/hides the Type column in the Contact Log. |
| **Rebuilds: Manual / Nightly Sync Only** | "Fast entry mode." Skips the live rebuild after every edit; sheets refresh on Build or the nightly sync. Great for bulk entry. |
| **Disable Nightly Sync Entirely** | Turns off the automatic nightly rebuild and phone sync completely. |
| **Update Contact Log only on Build** | Notes stop syncing on every edit and update only when you run Build. |
| **Disable Contact Log Auto-Fill** | Stops auto-stamping Date, #, and Follow-up. The parent/student/OSIS lookup still fills in. |
| **Date Range for Compiled Notes** | Limits compiled notes to a window (Last Month / 6 Months / Year presets). Blank = all time. |
| **Disable Hidden Notes (Show All)** | Includes notes you've hidden with the row checkbox in the compiled output. |
| **Hide "Discharged" in Master Table** | Visual only — filters discharged students out of view. |
| **Purge Discharged Students** | **Permanently** deletes discharged students from the Master Table. Cannot be undone. |
| **Reset to Default Settings** | Returns every toggle to its default (optional tabs/columns hidden, features on). |

> 📷 Screenshot: The ⚙️ Settings dialog with its toggle switches.

> ✅ **Settings travel with your copy.** Your choices are saved into the hidden Version
> tab, so a future upgrade copy inherits them automatically.

### Build vs. Full Sync — know the difference

| Command | What it is |
|---|---|
| 🏗️ **Build / Rebuild Data** | An **internal** action. Takes your RAW Data and constructs all the automated sheets (Directory, Master Table, Parents Divided, Phone Contacts, PCAR). It only organizes data inside the spreadsheet. |
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

> 📷 Screenshot: The System Upgrade Wizard with the old-sheet URL box and the six-step checklist.

> ✅ **Migration is one-directional and non-destructive.** It reads your old sheet and
> writes into the new one. Nothing is deleted from the old sheet — keep it as a backup.

### Staying current (the update checker)

Each copy remembers its own build date. Quietly, on open (at most once every few hours),
the sheet checks a shared **Version Tracker** to see whether a newer template has
shipped. If so, it pops a short changelog with a link to the newest copy. You can also
check any time with **🚀 App Menu → 🔔 Check for Updates**.

Inside that dialog you can:

- **Dismiss** — hide this update until something even newer ships.
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

> 📷 Screenshot: The End of Year Rollover wizard — archive summary and the "Get New Template" button.

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
4. Fresh ATS export? Paste it into RAW Data (Paste Special → Values Only), then run
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

### Troubleshooting / FAQ

| Symptom | Fix |
|---|---|
| The custom menus didn't appear | Wait a few seconds, then press F5. If it's a brand-new copy, run **🚨 Initial Setup** first. |
| "🔒 Action Blocked" when I type | Setup hasn't been run yet. Click **🚀 App Menu → 🚨 Initial Setup (Run Once)**. |
| Pasted data looks broken / directory is wrong | Re-paste using **Edit → Paste Special → Values Only**, then run **🏗️ Build Sheets Only**. |
| Ghost Typist: "Could not find … in the Master Table" | That family isn't compiled yet. Run **🏗️ Build Sheets Only**, then try again. |
| "📡 Background Sync in Progress" warning | A rebuild/sync is running. Wait for it to finish before entering new data. |
| Update checker says I'm up to date but there's a new version | Use **🔔 Check for Updates** (it bypasses the throttle). If still wrong, report a bug. |
| Phone sync created duplicates / didn't finish | Open Full Sync, use Stop then Restart. The sync mirrors the sheet and cleans up on the next full run. |
| Something's broken or confusing | Use **🚀 App Menu → 🐞 Report a Bug**, or email jacob.keller@p94m.org. |

### Do & Don't

| ✅ Do | 🚫 Don't |
|---|---|
| Paste raw data only into RAW Data — as Values Only. | Type into Master Table, Directory, or other generated tabs — they're overwritten on build. |
| Use Overrides to fix names, phones, or add students. | Rename or delete the generated / hidden tabs (including Version). |
| Run Build after big changes. | Hand-edit the hidden Version tab's saved settings. |
| Keep your old / retired sheet as a backup until the new one is verified. | Purge discharged students unless you're certain. |

### Glossary

| Term | Meaning |
|---|---|
| **OSIS** | The unique student ID number from ATS. Used to match students, siblings, and contacts. |
| **ATS / Insight** | The NYC DOE student data systems your Raw Data export comes from (via TeachHub). |
| **PCAR** | Parent Coordinator Activity Report — the reporting numbers this tool compiles. |
| **Ghost Typist** | The auto-fill that completes a Contact Log (or Notes) row from a single OSIS or name. |
| **Override** | A correction or addition to student/guardian info without changing raw ATS data. |
| **Stealth Mode** | An override flag that keeps a student out of the Override Report. |
| **NYCSA / ParentSquare** | Parent account/communication systems whose exports feed the directory. |
| **Build** | Regenerating all derived tabs from your current data (**🏗️ Build Sheets Only**). |

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
