# 🚀 Parent Coordinator Tracker — User Guide

**User Guide for New Coordinators** — Set-up · Buttons & Features · Migration & Updates · Privacy

> A Google Sheets tool that turns your ATS export into a live directory, tracks every
> family contact, and compiles your PCAR — automatically.

> **Note:** This Markdown file is the editable source of the guide. A ready-to-use
> Word version lives next to it at `docs/PC_Tracker_User_Guide.docx` — drag it into
> Google Drive to open it as a Google Doc.

---

## What this tool is

The Parent Coordinator (PC) Tracker is a **Google Sheets add-on** built with Google
Apps Script. You paste your ATS student export into one tab, and the tool automatically
builds and maintains everything else: a clean class directory, a per-student contact
log, PCAR reporting numbers, and — if you want it — a copy of the directory pushed into
your phone through Google Contacts.

In plain terms, it does four jobs for you:

- **Directory automation** — one paste of raw data becomes a formatted, printable
  directory that handles siblings, multiple guardians, and overrides.
- **Contact tracking** — log every parent/guardian interaction; the sheet fills in
  names, OSIS, dates, and totals for you.
- **PCAR compilation** — your reporting figures are tallied as you work, so there's
  nothing to add up by hand.
- **Phone sync (optional)** — mirror the directory into a "School Directory" label in
  your Google Contacts so parent numbers show up on your phone.

> **New here?** Read the Quick Start below, do the one-time setup, and you're running in
> about five minutes.

---

## Part 1 — Setting Up Your Sheet

### Quick Start (5 steps)

1. **Make your own copy** of the template from the link you were given (it opens a
   "Make a copy" prompt, or use *File → Make a copy*). Everyone works in their own copy.
2. Open your copy. A welcome pop-up appears. At the top of the screen, click
   **🚀 App Menu → 🚨 Initial Setup (Run Once)**.
3. Authorize the script when Google asks (one-time — see "The authorization screen" below).
4. Add your data: either paste your ATS export into the **RAW Data** tab (brand-new
   users) or use **🚀 App Menu → 🔄 Import Old Data** (upgrading users).
5. Run **🚀 App Menu → 🏗️ Build Sheets Only**. When it finishes, press **F5** to
   refresh — the full menu now unlocks.

### Step-by-step, in detail

**1. Make a copy and open it.** You always start from a shared template. Making a copy
drops a fresh, private spreadsheet into your own Google Drive. The template itself is
never touched by your work, and no one else can see your copy.

**2. Run Initial Setup (once).** The first time your copy opens, the menu shows a single
button: **🚨 Initial Setup (Run Once)**. Clicking it does the behind-the-scenes wiring:

- Authorizes the script to run as you.
- Installs the automatic triggers — instant updates when you edit the Contact Log, a
  nightly refresh/sync (around 11 PM), and a periodic "check for updates."
- Records this copy as your original (a clone-detector so a later copy starts clean).
- Applies the default tab layout and opens this User Guide.

> 🔒 **Until setup is done, the sheet is intentionally locked.** If you try to type
> before running Initial Setup, you'll see "🔒 Action Blocked — Initialization process
> needs to occur first." That's expected — just run Initial Setup.

**3. The authorization screen (why it looks scary, and why it's safe).** Because this is
a personal tool and not a paid Marketplace app, Google shows a warning like *"Google
hasn't verified this app."* To continue, click **Advanced → Go to … (unsafe) → choose
your account → Allow**.

"Unverified" means Google hasn't done a formal brand review — *not* that anything is
wrong. You are granting permission to a script that runs entirely inside **your own**
account. See **Part 4 — Security & Privacy** for exactly what each permission is for and
why the developer can never see your data.

**4. Add your data — two paths.**

| You are… | Do this | Where |
|---|---|---|
| Brand-new (no old sheet) | Paste your raw ATS student export. The tool cleans the paste, fixes number formatting, and time-stamps it. | **RAW Data** tab |
| Upgrading (have an older PC sheet) | Use the Migration Wizard to pull your data across safely. It never deletes anything from your old sheet. | **🚀 App Menu → 🔄 Import Old Data** |

**5. Build the sheets.** Run **🚀 App Menu → 🏗️ Build Sheets Only**. A progress sidebar
opens and generates every derived tab (Master Table, Directory, Combined Contact
Tracking, Phone Contacts, PCAR, and more). The *first* successful build unlocks the full
menu — a toast will ask you to press **F5** to refresh and reveal the new options.

> ✅ **The golden rule of this sheet:** RAW Data is the only tab you paste into.
> Everything else is generated — so re-running Build always safely rebuilds it. Never
> hand-type into Master Table, Directory, or other generated tabs; your edits would be
> overwritten on the next build.

---

## Part 2 — Buttons & Features

After setup, three custom menus sit at the top of the screen, to the right of "Help":
**🚀 App Menu**, **👤 Student Overrides**, and **📆 Events**.

### 🚀 App Menu

| Menu item | What it does |
|---|---|
| 🧱 **Full Sync (Build + Push to Google Contacts)** | Rebuilds every sheet **and** pushes the directory to your Google Contacts. Only appears when phone syncing is turned on in Settings. Opens a progress sidebar you can pause or stop. |
| 🏗️ **Build Sheets Only (No Push)** | Your everyday "refresh everything" button. Regenerates all derived tabs from RAW Data + Contact Log + overrides, without touching your phone contacts. |
| ⚙️ **Settings** | Show/hide tabs and columns, control live updates, filter notes, manage discharged students (full list below). |
| 📖 **Open User Guide** | Opens this guide. |
| 🐞 **Report a Bug** | Opens a short Google Form so you can report anything broken or confusing. |
| 🔔 **Check for Updates** | Manually checks whether a newer template has been released, and lets you set your update preferences. |
| 🔄 **End of Year Rollover** | Archives the year to a PDF, freezes the sheet, and hands you next year's fresh template (see Part 3). |

### 👤 Student Overrides

"Overrides" let you correct or add student and guardian information without editing the
raw ATS data. Use them when ATS is wrong, out of date, or missing someone.

| Menu item | What it does |
|---|---|
| ✨ **New / Edit Student** | Add a student who isn't in ATS, or correct a name, phone, email, language, class, grade, or up to 5 guardians. You can also set a status **Label** and turn on **"Hide from Report (Stealth Mode)."** Saves and rebuilds in the background (~1 minute). |
| 🗂️ **Manage Overrides** | See every override you've made; edit or delete any of them. |
| 📄 **Override Report** | Builds a side-by-side "Students Override Report" comparing the raw ATS data against your overrides. Stealth-mode students are skipped. |

### 📆 Events

**📋 New Event** opens a one-box form: event name, type/category (with a handy **PTC
Quick Fill** button that types "Parent Teacher Conference"), location, and a date picker.
It builds an event table and feeds a backend event log, so events count toward your
Combined Contact Tracking and PCAR.

### The tabs — what each one is for

| Tab | Purpose | Visible by default? |
|---|---|---|
| **RAW Data** | Your ATS paste zone — the only tab you type/paste raw data into. | Yes |
| **Contact Log** | The heart of the tool. Log every guardian contact here (deep dive below). | Yes |
| **Master Table** | The generated master roster: OSIS, student, guardian, status, and more. Data starts on row 5. | Yes (generated) |
| **Directory** | A clean, printable class directory. | Yes (generated) |
| **Combined Contact Tracking** | Per-student contact totals and last-contact dates, pulled from the Contact Log and Events. | Yes (generated) |
| **PCAR** | Your auto-compiled PCAR reporting numbers. | Yes (generated) |
| **Phone Contacts** | The exact data that gets pushed to Google Contacts. | Hidden (optional) |
| **Parents Divided** | One row per guardian, with ParentSquare / NYCSA flags. | Hidden (optional) |
| **Notes** | An optional manual notes tab with its own auto-fill and duplicate protection. | Hidden (optional) |
| **Send Out** | Optional tab tied to extra PCAR rows. | Hidden (optional) |
| **Version / Backend_Event_Log** | Behind-the-scenes tabs (build date, saved settings, event data). Leave them alone. | Hidden (system) |

### The Contact Log, in depth

This is where you'll spend most of your time. It's designed so you type as little as
possible:

- **Ghost Typist.** Type an OSIS, a student's name, *or* a guardian's name into a new row
  and the sheet auto-fills the other identity columns for you — plus Site and Class,
  including for siblings. If it can't match what you typed, it warns "Could not find … in
  the Master Table."
- **Smart auto-fill.** As soon as you enter real content in a row, it stamps today's
  Date, sets # of interactions to 1, and Follow-up to "No." These are fill-if-blank —
  anything you've already typed (a back-dated date, a count of 3, "Yes – Important") is
  never overwritten.
- **Auto-adding rows.** A fresh blank row is always kept above the dark "END" bar, so you
  never run out of space.
- **The master hide filter.** The checkbox on the END-bar row toggles hiding of rows
  you've marked/hidden, so you can collapse old entries and focus on active ones.

> **If the Ghost Typist says "Could not find … in the Master Table":** That family isn't
> in the Master Table yet. Run **🏗️ Build Sheets Only** once so your latest RAW Data /
> overrides are compiled, then try again.

### ⚙️ Settings — every toggle explained

| Setting | What it controls |
|---|---|
| **Sync Contacts to Phones (Nightly)** | Master switch for phone sync. When on, the nightly run pushes the directory into your Google Contacts (and the Full Sync menu item appears). |
| **Show "Phone Contacts" / "Parents Divided" / "Notes" / "Send Out"** | Show or hide each optional tab. All start hidden. |
| **Show Site & Class (Contact Log)** | Shows/hides the Site and Class columns in the Contact Log. |
| **Show "Type" (Contact Log)** | Shows/hides the Type column in the Contact Log. |
| **Rebuilds: Manual / Nightly Sync Only** | "Fast entry mode." Skips the live rebuild after every edit; sheets refresh when you run Build or on the nightly sync. Great for bulk data entry. |
| **Disable Nightly Sync Entirely** | Turns off the automatic nightly rebuild and phone sync completely. |
| **Update Contact Log only on Build** | Notes stop syncing on every edit and update only when you run Build. |
| **Disable Contact Log Auto-Fill** | Stops auto-stamping Date, #, and Follow-up. The parent/student/OSIS lookup still fills in. |
| **Date Range for Compiled Notes** | Limits compiled notes to a window (with Last Month / 6 Months / Year presets). Blank = all time. |
| **Disable Hidden Notes (Show All)** | Includes notes you've hidden with the row checkbox in the compiled output. |
| **Hide "Discharged" in Master Table** | Visual only — filters discharged students out of view. Turn off to show them again. |
| **Purge Discharged Students** | **Permanently** deletes discharged students from the Master Table. Cannot be undone. |
| **Reset to Default Settings** | Returns every toggle to its default (all optional tabs/columns hidden, all features on). |

> ✅ **Settings travel with your copy.** Your choices are quietly saved into the hidden
> Version tab, so if you upgrade to a newer template later, a fresh copy inherits your
> settings automatically.

### Phone sync (Google Contacts)

When phone sync is on, the tool keeps a **"School Directory"** label in your Google
Contacts in step with the sheet. It runs in "mirror" mode — creating, editing, and
removing contacts so the label always matches your directory. It runs nightly, or on
demand via **🧱 Full Sync**. A progress sidebar shows live status and lets you pause,
stop, or restart; the sync is resilient and picks up where it left off if it hits
Google's time limits.

---

## Part 3 — Migration & Updates

### Importing data from an older sheet (Migration Wizard)

Open **🚀 App Menu → 🔄 Import Old Data**. (On an "upgrade" copy of the template it opens
by itself.) Paste the link to your older PC sheet and click **Start Safe Migration**. The
wizard runs six steps:

1. Connect to the old sheet (confirms the link works and you have access).
2. Recover your custom settings (inferred from which tabs were visible in the old sheet).
3. Migrate your ATS Raw Data.
4. Transfer your Events, Send Out, and Student Overrides.
5. Smart-map the Contact Log — it matches columns by their headers, so it still works even
   if the layout changed between versions, and it strips blank/"ghost" rows automatically.
6. Finalize and rebuild all directories.

> ✅ **Migration is one-directional and non-destructive.** It reads your old sheet (which
> you already own) and writes into the new one. Nothing is deleted or changed in your old
> sheet — keep it as a backup until you're happy with the new copy.

### Staying current (the update checker)

Each copy remembers its own build date. Quietly, on open (at most once every few hours),
the sheet checks a shared **Version Tracker** to see whether a newer template has shipped.
If so, it pops a short changelog with a link to the newest copy. You can also check any
time with **🚀 App Menu → 🔔 Check for Updates**.

Inside that dialog you can:

- **Dismiss** — hide this update until something even newer ships.
- **Major updates only** — automatic pop-ups only for big releases (a manual check still
  shows everything).
- **Stop automatic update checks** — silence the on-open checks; the manual menu check
  always still works.

**How to actually adopt an update.** Updates don't rewrite your current sheet in place. To
move up a version: open the new template link, **make a copy**, run Initial Setup, then use
the **Migration Wizard** to bring your data across from your current sheet. Your settings
come along automatically.

### End of Year Rollover

When the year ends, run **🚀 App Menu → 🔄 End of Year Rollover**. It:

- Generates a permanent **PDF archive** of your Contact Log and Combined Contact Tracking,
  saved next to your sheet in Drive.
- **Freezes the sheet** — turns off nightly and phone sync so it won't overwrite next
  year's data — and renames it "… – Retired on [date]."
- Hands you the link to next year's fresh, fast template.

> **Rollover is safe to run once you're done for the year.** It archives and freezes the
> OLD sheet; you then set up a brand-new copy for the new year (and can migrate anything
> you need). Keep the retired sheet and its PDF as your record.

---

## Part 4 — Security & Privacy

> ✅ **The short version: I cannot see your sheet, your students, your contacts, or your
> Drive.** This tool has no server and no database of mine. The code runs entirely inside
> your own Google account. Nothing about your families is ever sent to me or anyone else.

### Why that's true

- **It runs as you, on Google's servers.** When you click Authorize, you are granting
  permission to the script — running inside your account — not to me. I receive nothing.
- **There's no outside destination for your data.** The tool reads and writes your own
  spreadsheet and (optionally) your own Google Contacts. It doesn't phone home.
- **You can read every line of the code.** Open *Extensions → Apps Script* in the sheet to
  see exactly what it does.
- **You can revoke access anytime.** Go to *myaccount.google.com → Security → Your
  connections to third-party apps & services*, and remove the tool. Your data stays; the
  automation simply stops.

### Why the "unverified app" warning appears

That screen shows for personal Apps Script tools that aren't published to Google's
Marketplace. It means Google hasn't run a formal brand review — *not* that the tool is
unsafe. Approving it grants permissions to code operating only within your account.

### What each permission is for

The tool asks only for the access it needs to do its job:

| Permission Google shows | Used for |
|---|---|
| See, edit, create, and delete your spreadsheets | Reading RAW Data and writing the Master Table, Directory, Contact Log, PCAR, and all derived tabs. |
| See, edit, download, and permanently delete your contacts | The optional "School Directory" phone sync (only if you turn it on). |
| See, create, and delete your Google Drive files | Saving the PDF archive at year-end rollover, and opening your old sheet during migration. |
| Connect to an external service | A read-only check of the shared Version Tracker so you're told when a newer template exists. |
| Run when you're not present / manage triggers | The automatic triggers: instant Contact Log updates, the nightly refresh/sync, and the update check. |
| Display and run third-party web content in dialogs | The menus, wizards, and sidebars (Settings, Migration, Rollover, progress bars). |

### The only thing that ever leaves your sheet

- **The update check** reads a shared, view-only "Version Tracker" spreadsheet to compare
  dates. It sends nothing about your students — it only reads a list of release dates and
  links.
- **Migration** opens an older sheet only when you paste its link, and only because you
  already have access to it.
- **Bug reports** go through a Google Form you choose to fill out — only the words you type
  are shared.

---

## Tips, Workflow & Troubleshooting

### A good daily workflow

1. Made a parent contact? Go to the Contact Log, type the OSIS or a name in a new row, and
   let the Ghost Typist fill the rest. Add your notes.
2. Need to fix a name, phone, or add a missing student? Use **👤 Student Overrides → New /
   Edit Student** rather than editing RAW Data by hand.
3. Got a fresh ATS export? Paste it into RAW Data, then run **🏗️ Build Sheets Only**.
4. Doing a big batch of entries? Turn on "Rebuilds: Manual / Nightly Sync Only" in Settings
   for speed, then run Build once at the end.

### Handy tips

- **Siblings are automatic.** The tool joins siblings by their OSIS numbers and fills
  Site/Class for each — you don't manage them separately.
- **Back-dating works.** Type your own date, count, or follow-up status and the auto-fill
  leaves it alone.
- **Hide vs. Purge discharged.** "Hide Discharged" is a reversible view filter; "Purge
  Discharged Students" permanently deletes them. Use Purge only when you're sure.
- **Refresh if a menu looks stale.** Many actions update the menu live, but pressing F5
  always reloads the latest menu.

### Troubleshooting / FAQ

| Symptom | Fix |
|---|---|
| The custom menus didn't appear | Give the sheet a few seconds after opening, then press F5. If it's a brand-new copy, run **🚨 Initial Setup** first. |
| "🔒 Action Blocked" when I type | Setup hasn't been run yet. Click **🚀 App Menu → 🚨 Initial Setup (Run Once)**. |
| Ghost Typist: "Could not find … in the Master Table" | That family isn't compiled yet. Run **🏗️ Build Sheets Only**, then try again. |
| "📡 Background Sync in Progress" warning | A rebuild/sync is running. Wait for it to finish before entering new data, to avoid save conflicts. |
| Update checker says I'm up to date but I know there's a new version | Use **🔔 Check for Updates** (it bypasses the throttle). If it still looks wrong, the tracker date may be off — report a bug. |
| Phone sync created duplicates / didn't finish | Open Full Sync, use Stop then Restart. The sync mirrors the sheet and cleans up on the next full run. |
| Something's broken or confusing | Use **🚀 App Menu → 🐞 Report a Bug** and describe what happened. |

### Do & Don't

| ✅ Do | 🚫 Don't |
|---|---|
| Paste raw data only into RAW Data. | Type into Master Table, Directory, or other generated tabs — they're overwritten on build. |
| Use Overrides to fix names, phones, or add students. | Rename or delete the generated / hidden tabs (including Version). |
| Run Build after big changes. | Hand-edit the hidden Version tab's saved settings. |
| Keep your old / retired sheet as a backup until you've verified the new one. | Purge discharged students unless you're certain. |

### Glossary

| Term | Meaning |
|---|---|
| **OSIS** | The unique student ID number from ATS. The tool uses it to match students, siblings, and contacts. |
| **ATS** | The NYC DOE student data system your Raw Data export comes from. |
| **PCAR** | Parent Coordinator Activity Report — the reporting numbers this tool compiles for you. |
| **Ghost Typist** | The auto-fill that completes a Contact Log (or Notes) row from a single OSIS or name. |
| **Override** | A correction or addition you make to student/guardian info without changing raw ATS data. |
| **Stealth Mode** | An override flag that keeps a student out of the Override Report. |
| **NYCSA / ParentSquare** | Parent account/communication systems tracked as flags on the Parents Divided tab. |
| **Build** | Regenerating all derived tabs from your current data (**🏗️ Build Sheets Only**). |

---

*Questions or something broken? Use **🚀 App Menu → 🐞 Report a Bug**.*
