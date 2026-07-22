// ======================================================================
// Generator for docs/PC_Tracker_User_Guide.docx (the new-user guide).
// This is a documentation helper only — it is NOT part of the Apps Script
// project and never runs on Google's servers.
//
// To regenerate the .docx after editing this file:
//   npm install docx        # one-time; docx is the only dependency
//   node docs/build_guide.js docs/PC_Tracker_User_Guide.docx
//
// Keep docs/USER_GUIDE.md (the Markdown source of truth) in sync with any
// wording changes here.
// ======================================================================
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  PageBreak, LevelFormat, TableOfContents, ExternalHyperlink, PageOrientation
} = require("docx");

// ---------- palette ----------
const NAVY = "0F172A";
const BLUE = "1D4ED8";
const SLATE = "334155";
const MUTE = "64748B";
const LINE = "CBD5E1";
const HEADSHADE = "0F2A4A"; // table header fill (dark navy)
const ROWSHADE = "F1F5F9";  // zebra
const GREEN = "166534";
const RED = "B91C1C";

// ---------- helpers ----------
function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 140 },
    children: [new TextRun({ text, bold: true, color: NAVY, size: 30, font: "Calibri" })],
    border: { bottom: { color: BLUE, size: 12, style: BorderStyle.SINGLE, space: 6 } },
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 90 },
    children: [new TextRun({ text, bold: true, color: BLUE, size: 24, font: "Calibri" })],
  });
}
function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 160, after: 60 },
    children: [new TextRun({ text, bold: true, color: SLATE, size: 21, font: "Calibri" })],
  });
}
function p(runs, opts = {}) {
  const children = (typeof runs === "string")
    ? [new TextRun({ text: runs, color: SLATE, size: 21, font: "Calibri" })]
    : runs;
  return new Paragraph({ spacing: { after: opts.after != null ? opts.after : 120 }, alignment: opts.align, children });
}
function t(text, o = {}) {
  return new TextRun({ text, bold: o.bold, italics: o.italics, color: o.color || SLATE, size: o.size || 21, font: "Calibri" });
}
function bullet(runs, level = 0) {
  const children = (typeof runs === "string") ? [t(runs)] : runs;
  return new Paragraph({ numbering: { reference: "bullets", level }, spacing: { after: 60 }, children });
}
function numbered(runs, ref = "steps") {
  const children = (typeof runs === "string") ? [t(runs)] : runs;
  return new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 80 }, children });
}
function callout(title, body, color = BLUE, shade = "EEF4FF") {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    columnWidths: [9360],
    borders: {
      top: { style: BorderStyle.SINGLE, size: 2, color: shade },
      bottom: { style: BorderStyle.SINGLE, size: 2, color: shade },
      left: { style: BorderStyle.SINGLE, size: 24, color: color },
      right: { style: BorderStyle.SINGLE, size: 2, color: shade },
      insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE },
    },
    rows: [new TableRow({ children: [new TableCell({
      width: { size: 9360, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: shade, color: "auto" },
      margins: { top: 120, bottom: 120, left: 160, right: 160 },
      children: [
        new Paragraph({ spacing: { after: body ? 40 : 0 }, children: [new TextRun({ text: title, bold: true, color, size: 21, font: "Calibri" })] }),
        ...(body ? [new Paragraph({ children: [new TextRun({ text: body, color: SLATE, size: 20, font: "Calibri" })] })] : []),
      ],
    })] })],
  });
}
function spacer(after = 120) { return new Paragraph({ spacing: { after }, children: [] }); }

// ---------- table builder ----------
function cell(content, { header = false, width, shade, bold = false, color, align } = {}) {
  const paras = (Array.isArray(content) ? content : [content]).map((c) =>
    new Paragraph({
      alignment: align,
      spacing: { after: 0 },
      children: (typeof c === "string")
        ? [new TextRun({ text: c, bold: header || bold, color: header ? "FFFFFF" : (color || SLATE), size: header ? 19 : 19, font: "Calibri" })]
        : c,
    })
  );
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: header ? HEADSHADE : (shade || "FFFFFF"), color: "auto" },
    margins: { top: 70, bottom: 70, left: 110, right: 110 },
    children: paras,
  });
}
function table(headers, rows, widths) {
  const total = widths.reduce((a, b) => a + b, 0);
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((hdr, i) => cell(hdr, { header: true, width: widths[i] })),
  });
  const bodyRows = rows.map((r, ri) =>
    new TableRow({
      children: r.map((c, ci) => cell(c, { width: widths[ci], shade: ri % 2 ? ROWSHADE : "FFFFFF" })),
    })
  );
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: widths,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 2, color: LINE },
      bottom: { style: BorderStyle.SINGLE, size: 2, color: LINE },
      left: { style: BorderStyle.SINGLE, size: 2, color: LINE },
      right: { style: BorderStyle.SINGLE, size: 2, color: LINE },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: LINE },
      insideVertical: { style: BorderStyle.SINGLE, size: 2, color: LINE },
    },
    rows: [headerRow, ...bodyRows],
  });
}

// ============================================================
// CONTENT
// ============================================================
const body = [];
const P = (...x) => body.push(...x);

// ---------- TITLE PAGE ----------
P(
  new Paragraph({ spacing: { before: 1400, after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "🚀", size: 96, font: "Calibri" })] }),
  new Paragraph({ spacing: { before: 200, after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Parent Coordinator Tracker", bold: true, color: NAVY, size: 56, font: "Calibri" })] }),
  new Paragraph({ spacing: { before: 60, after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "User Guide for New Coordinators", color: BLUE, size: 30, font: "Calibri" })] }),
  new Paragraph({ spacing: { before: 40, after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Set-up · Buttons & Features · Migration & Updates · Privacy", color: MUTE, size: 20, italics: true, font: "Calibri" })] }),
  new Paragraph({ spacing: { before: 900, after: 0 }, alignment: AlignmentType.CENTER, border: { top: { color: LINE, size: 8, style: BorderStyle.SINGLE, space: 10 } },
    children: [new TextRun({ text: "A Google Sheets tool that turns your ATS export into a live directory, tracks every", color: SLATE, size: 20, font: "Calibri" })] }),
  new Paragraph({ spacing: { after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "family contact, and compiles your PCAR — automatically.", color: SLATE, size: 20, font: "Calibri" })] }),
  new Paragraph({ children: [new PageBreak()] }),
);

// ---------- TOC ----------
P(h1("Contents"));
P(new Paragraph({ spacing: { after: 200 }, children: [
  new TextRun({ text: "Right-click any entry and choose “Update field” if page numbers look off. In Google Docs, the table of contents refreshes with the circular-arrow button.", italics: true, color: MUTE, size: 19, font: "Calibri" }),
]}));
P(new TableOfContents("Guide Contents", { hyperlink: true, headingStyleRange: "1-2" }));
P(new Paragraph({ children: [new PageBreak()] }));

// ============================================================
// INTRO
// ============================================================
P(h1("What this tool is"));
P(p([
  t("The Parent Coordinator (PC) Tracker is a "), t("Google Sheets add-on", { bold: true }),
  t(" built with Google Apps Script. You paste your ATS student export into one tab, and the tool automatically builds and maintains everything else: a clean class directory, a per-student contact log, PCAR reporting numbers, and — if you want it — a copy of the directory pushed into your phone through Google Contacts."),
]));
P(p("In plain terms, it does four jobs for you:"));
P(
  bullet([t("Directory automation — ", { bold: true }), t("one paste of raw data becomes a formatted, printable directory that handles siblings, multiple guardians, and overrides.")]),
  bullet([t("Contact tracking — ", { bold: true }), t("log every parent/guardian interaction; the sheet fills in names, OSIS, dates, and totals for you.")]),
  bullet([t("PCAR compilation — ", { bold: true }), t("your reporting figures are tallied as you work, so there's nothing to add up by hand.")]),
  bullet([t("Phone sync (optional) — ", { bold: true }), t("mirror the directory into a “School Directory” label in your Google Contacts so parent numbers show up on your phone.")]),
);
P(callout("New here? Read the Quick Start on the next page, do the one-time setup, and you're running in about five minutes.", "", BLUE, "EEF4FF"));

// ============================================================
// PART 1 — SETUP
// ============================================================
P(new Paragraph({ children: [new PageBreak()] }));
P(h1("Part 1 — Setting Up Your Sheet"));

P(h2("Quick Start (5 steps)"));
P(
  numbered([t("Make your own copy", { bold: true }), t(" of the template from the link you were given (it opens a “Make a copy” prompt, or use File → Make a copy). Everyone works in their own copy.")]),
  numbered([t("Open your copy. A welcome pop-up appears. At the top of the screen, click "), t("🚀 App Menu → 🚨 Initial Setup (Run Once)", { bold: true }), t(".")]),
  numbered([t("Authorize the script when Google asks (this is a one-time step — see “The authorization screen” below).")]),
  numbered([t("Add your data: either paste your ATS export into the "), t("RAW Data", { bold: true }), t(" tab (brand-new users) or use "), t("🚀 App Menu → 🔄 Import Old Data", { bold: true }), t(" (upgrading users).")]),
  numbered([t("Run "), t("🚀 App Menu → 🏗️ Build Sheets Only", { bold: true }), t(". When it finishes, press "), t("F5", { bold: true }), t(" to refresh — the full menu now unlocks.")]),
);

P(h2("Step-by-step, in detail"));

P(h3("1. Make a copy and open it"));
P(p("You always start from a shared template. Making a copy drops a fresh, private spreadsheet into your own Google Drive. The template itself is never touched by your work, and no one else can see your copy."));

P(h3("2. Run Initial Setup (once)"));
P(p([
  t("The first time your copy opens, the menu shows a single button: "),
  t("🚨 Initial Setup (Run Once)", { bold: true }),
  t(". Clicking it does the behind-the-scenes wiring:"),
]));
P(
  bullet("Authorizes the script to run as you."),
  bullet("Installs the automatic triggers — instant updates when you edit the Contact Log, a nightly refresh/sync (around 11 PM), and a periodic “check for updates.”"),
  bullet("Records this copy as your original (a clone-detector so a later copy starts clean)."),
  bullet("Applies the default tab layout and opens this User Guide."),
);
P(callout("Until setup is done, the sheet is intentionally locked.", "If you try to type before running Initial Setup, you'll see “🔒 Action Blocked — Initialization process needs to occur first.” That's expected — just run Initial Setup.", RED, "FEF2F2"));

P(h3("3. The authorization screen (why it looks scary, and why it's safe)"));
P(p([
  t("Because this is a personal tool and not a paid Marketplace app, Google shows a warning like "),
  t("“Google hasn't verified this app.”", { italics: true }),
  t(" To continue, click "),
  t("Advanced", { bold: true }),
  t(" → "),
  t("Go to … (unsafe)", { bold: true }),
  t(" → choose your account → "),
  t("Allow", { bold: true }),
  t("."),
]));
P(p([
  t("“Unverified” means Google hasn't done a formal brand review — "),
  t("not", { italics: true }),
  t(" that anything is wrong. You are granting permission to a script that runs entirely inside "),
  t("your own", { bold: true }),
  t(" account. See "),
  t("Part 4 — Security & Privacy", { bold: true }),
  t(" for exactly what each permission is for and why the developer can never see your data."),
]));

P(h3("4. Add your data — two paths"));
P(table(
  ["You are…", "Do this", "Where"],
  [
    ["Brand-new (no old sheet)", "Paste your raw ATS student export into the RAW Data tab. The tool cleans the paste, fixes number formatting, and time-stamps it.", "RAW Data tab"],
    ["Upgrading (have an older PC sheet)", "Use the Migration Wizard to pull your data across safely. It never deletes anything from your old sheet.", "🚀 App Menu → 🔄 Import Old Data"],
  ],
  [2100, 4560, 2700],
));

P(spacer(60));
P(h3("5. Build the sheets"));
P(p([
  t("Run "), t("🚀 App Menu → 🏗️ Build Sheets Only", { bold: true }),
  t(". A progress sidebar opens and generates every derived tab (Master Table, Directory, Combined Contact Tracking, Phone Contacts, PCAR, and more). The "),
  t("first", { italics: true }),
  t(" successful build unlocks the full menu — a toast will ask you to press "),
  t("F5", { bold: true }), t(" to refresh and reveal the new options."),
]));
P(callout("The golden rule of this sheet", "RAW Data is the only tab you paste into. Everything else is generated — so re-running Build always safely rebuilds it. Never hand-type into Master Table, Directory, or other generated tabs; your edits would be overwritten on the next build.", GREEN, "F0FDF4"));

// ============================================================
// PART 2 — BUTTONS & FEATURES
// ============================================================
P(new Paragraph({ children: [new PageBreak()] }));
P(h1("Part 2 — Buttons & Features"));
P(p([
  t("After setup, three custom menus sit at the top of the screen, to the right of “Help”: "),
  t("🚀 App Menu", { bold: true }), t(", "), t("👤 Student Overrides", { bold: true }),
  t(", and "), t("📆 Events", { bold: true }), t("."),
]));

P(h2("🚀 App Menu"));
P(table(
  ["Menu item", "What it does"],
  [
    ["🧱 Full Sync (Build + Push to Google Contacts)", "Rebuilds every sheet AND pushes the directory to your Google Contacts. Only appears when phone syncing is turned on in Settings. Opens a progress sidebar you can pause or stop."],
    ["🏗️ Build Sheets Only (No Push)", "Your everyday “refresh everything” button. Regenerates all derived tabs from RAW Data + Contact Log + overrides, without touching your phone contacts."],
    ["⚙️ Settings", "Opens the Settings dialog — show/hide tabs and columns, control live updates, filter notes, manage discharged students (full list below)."],
    ["📖 Open User Guide", "Opens this guide."],
    ["🐞 Report a Bug", "Opens a short Google Form so you can report anything broken or confusing."],
    ["🔔 Check for Updates", "Manually checks whether a newer template has been released, and lets you set your update preferences."],
    ["🔄 End of Year Rollover", "Archives the year to a PDF, freezes the sheet, and hands you next year's fresh template (see Part 3)."],
  ],
  [3760, 5600],
));

P(spacer(60));
P(h2("👤 Student Overrides"));
P(p("“Overrides” let you correct or add student and guardian information without editing the raw ATS data. Use them when ATS is wrong, out of date, or missing someone."));
P(table(
  ["Menu item", "What it does"],
  [
    ["✨ New / Edit Student", "Add a student who isn't in ATS, or correct a name, phone, email, language, class, grade, or up to 5 guardians. You can also set a status Label and turn on “Hide from Report (Stealth Mode).” Saves and rebuilds in the background (~1 minute)."],
    ["🗂️ Manage Overrides", "See every override you've made; edit or delete any of them."],
    ["📄 Override Report", "Builds a side-by-side “Students Override Report” comparing the raw ATS data against your overrides. Stealth-mode students are skipped."],
  ],
  [3300, 6060],
));

P(spacer(60));
P(h2("📆 Events"));
P(p([
  t("📋 New Event", { bold: true }),
  t(" opens a one-box form: event name, type/category (with a handy "),
  t("PTC Quick Fill", { bold: true }),
  t(" button that types “Parent Teacher Conference”), location, and a date picker. It builds an event table and feeds a backend event log, so events count toward your Combined Contact Tracking and PCAR."),
]));

P(spacer(60));
P(h2("The tabs — what each one is for"));
P(table(
  ["Tab", "Purpose", "Visible by default?"],
  [
    ["RAW Data", "Your ATS paste zone — the only tab you type/paste raw data into.", "Yes"],
    ["Contact Log", "The heart of the tool. Log every guardian contact here (deep dive below).", "Yes"],
    ["Master Table", "The generated master roster: OSIS, student, guardian, status, and more. Data starts on row 5.", "Yes (generated)"],
    ["Directory", "A clean, printable class directory.", "Yes (generated)"],
    ["Combined Contact Tracking", "Per-student contact totals and last-contact dates, pulled from the Contact Log and Events.", "Yes (generated)"],
    ["PCAR", "Your auto-compiled PCAR reporting numbers.", "Yes (generated)"],
    ["Phone Contacts", "The exact data that gets pushed to Google Contacts.", "Hidden (optional)"],
    ["Parents Divided", "One row per guardian, with ParentSquare / NYCSA flags.", "Hidden (optional)"],
    ["Notes", "An optional manual notes tab with its own auto-fill and duplicate protection.", "Hidden (optional)"],
    ["Send Out", "Optional tab tied to extra PCAR rows.", "Hidden (optional)"],
    ["Version / Backend_Event_Log", "Behind-the-scenes tabs (build date, saved settings, event data). Leave them alone.", "Hidden (system)"],
  ],
  [2280, 5080, 2000],
));

P(spacer(80));
P(h2("The Contact Log, in depth"));
P(p("This is where you'll spend most of your time. It's designed so you type as little as possible:"));
P(
  bullet([t("Ghost Typist. ", { bold: true }), t("Type an OSIS, a student's name, OR a guardian's name into a new row and the sheet auto-fills the other identity columns for you — plus Site and Class, including for siblings. If it can't match what you typed, it warns “Could not find … in the Master Table.”")]),
  bullet([t("Smart auto-fill. ", { bold: true }), t("As soon as you enter real content in a row, it stamps today's Date, sets # of interactions to 1, and Follow-up to “No.” These are fill-if-blank — anything you've already typed (a back-dated date, a count of 3, “Yes – Important”) is never overwritten.")]),
  bullet([t("Auto-adding rows. ", { bold: true }), t("A fresh blank row is always kept above the dark “END” bar, so you never run out of space.")]),
  bullet([t("The master hide filter. ", { bold: true }), t("The checkbox on the END-bar row toggles hiding of rows you've marked/hidden, so you can collapse old entries and focus on active ones.")]),
);
P(callout("If the Ghost Typist says “Could not find … in the Master Table”", "That family isn't in the Master Table yet. Run 🏗️ Build Sheets Only once so your latest RAW Data / overrides are compiled, then try again.", BLUE, "EEF4FF"));

P(spacer(80));
P(h2("⚙️ Settings — every toggle explained"));
P(table(
  ["Setting", "What it controls"],
  [
    ["Sync Contacts to Phones (Nightly)", "Master switch for phone sync. When on, the nightly run pushes the directory into your Google Contacts (and the Full Sync menu item appears)."],
    ["Show “Phone Contacts” / “Parents Divided” / “Notes” / “Send Out”", "Show or hide each optional tab. All start hidden."],
    ["Show Site & Class (Contact Log)", "Shows/hides the Site and Class columns in the Contact Log."],
    ["Show “Type” (Contact Log)", "Shows/hides the Type column in the Contact Log."],
    ["Rebuilds: Manual / Nightly Sync Only", "“Fast entry mode.” Skips the live rebuild after every edit; sheets refresh when you run Build or on the nightly sync. Great for bulk data entry."],
    ["Disable Nightly Sync Entirely", "Turns off the automatic nightly rebuild and phone sync completely."],
    ["Update Contact Log only on Build", "Notes stop syncing on every edit and update only when you run Build."],
    ["Disable Contact Log Auto-Fill", "Stops auto-stamping Date, #, and Follow-up. The parent/student/OSIS lookup still fills in."],
    ["Date Range for Compiled Notes", "Limits compiled notes to a window (with Last Month / 6 Months / Year presets). Blank = all time."],
    ["Disable Hidden Notes (Show All)", "Includes notes you've hidden with the row checkbox in the compiled output."],
    ["Hide “Discharged” in Master Table", "Visual only — filters discharged students out of view. Turn off to show them again."],
    ["Purge Discharged Students", "Permanently deletes discharged students from the Master Table. Cannot be undone."],
    ["Reset to Default Settings", "Returns every toggle to its default (all optional tabs/columns hidden, all features on)."],
  ],
  [3560, 5800],
));
P(callout("Settings travel with your copy.", "Your choices are quietly saved into the hidden Version tab, so if you upgrade to a newer template later, a fresh copy inherits your settings automatically.", GREEN, "F0FDF4"));

P(spacer(80));
P(h2("Phone sync (Google Contacts)"));
P(p([
  t("When phone sync is on, the tool keeps a "),
  t("“School Directory”", { bold: true }),
  t(" label in your Google Contacts in step with the sheet. It runs in “mirror” mode — creating, editing, and removing contacts so the label always matches your directory. It runs nightly, or on demand via "),
  t("🧱 Full Sync", { bold: true }),
  t(". A progress sidebar shows live status and lets you pause, stop, or restart; the sync is resilient and picks up where it left off if it hits Google's time limits."),
]));

// ============================================================
// PART 3 — MIGRATION & UPDATES
// ============================================================
P(new Paragraph({ children: [new PageBreak()] }));
P(h1("Part 3 — Migration & Updates"));

P(h2("Importing data from an older sheet (Migration Wizard)"));
P(p([
  t("Open "), t("🚀 App Menu → 🔄 Import Old Data", { bold: true }),
  t(". (On an “upgrade” copy of the template it opens by itself.) Paste the link to your older PC sheet and click "),
  t("Start Safe Migration", { bold: true }), t(". The wizard runs six steps:"),
]));
P(
  numbered("Connect to the old sheet (confirms the link works and you have access)."),
  numbered("Recover your custom settings (inferred from which tabs were visible in the old sheet)."),
  numbered("Migrate your ATS Raw Data."),
  numbered("Transfer your Events, Send Out, and Student Overrides."),
  numbered("Smart-map the Contact Log — it matches columns by their headers, so it still works even if the layout changed between versions, and it strips blank/“ghost” rows automatically."),
  numbered("Finalize and rebuild all directories."),
);
P(callout("Migration is one-directional and non-destructive.", "It reads your old sheet (which you already own) and writes into the new one. Nothing is deleted or changed in your old sheet — keep it as a backup until you're happy with the new copy.", GREEN, "F0FDF4"));

P(h2("Staying current (the update checker)"));
P(p([
  t("Each copy remembers its own build date. Quietly, on open (at most once every few hours), the sheet checks a shared "),
  t("Version Tracker", { bold: true }),
  t(" to see whether a newer template has shipped. If so, it pops a short changelog with a link to the newest copy. You can also check any time with "),
  t("🚀 App Menu → 🔔 Check for Updates", { bold: true }), t("."),
]));
P(p("Inside that dialog you can:"));
P(
  bullet([t("Dismiss", { bold: true }), t(" — hide this update until something even newer ships.")]),
  bullet([t("Major updates only", { bold: true }), t(" — automatic pop-ups only for big releases (a manual check still shows everything).")]),
  bullet([t("Stop automatic update checks", { bold: true }), t(" — silence the on-open checks; the manual menu check always still works.")]),
);
P(h3("How to actually adopt an update"));
P(p([
  t("Updates don't rewrite your current sheet in place. To move up a version: open the new template link, "),
  t("make a copy", { bold: true }),
  t(", run Initial Setup, then use the "),
  t("Migration Wizard", { bold: true }),
  t(" to bring your data across from your current sheet. Your settings come along automatically."),
]));

P(h2("End of Year Rollover"));
P(p([
  t("When the year ends, run "), t("🚀 App Menu → 🔄 End of Year Rollover", { bold: true }),
  t(". It:"),
]));
P(
  bullet("Generates a permanent PDF archive of your Contact Log and Combined Contact Tracking, saved next to your sheet in Drive."),
  bullet("Freezes the sheet — turns off nightly and phone sync so it won't overwrite next year's data — and renames it “… – Retired on [date].”"),
  bullet("Hands you the link to next year's fresh, fast template."),
);
P(callout("Rollover is safe to run once you're done for the year.", "It archives and freezes the OLD sheet; you then set up a brand-new copy for the new year (and can migrate anything you need). Keep the retired sheet and its PDF as your record.", BLUE, "EEF4FF"));

// ============================================================
// PART 4 — SECURITY & PRIVACY
// ============================================================
P(new Paragraph({ children: [new PageBreak()] }));
P(h1("Part 4 — Security & Privacy"));
P(callout("The short version: I cannot see your sheet, your students, your contacts, or your Drive.",
  "This tool has no server and no database of mine. The code runs entirely inside your own Google account. Nothing about your families is ever sent to me or anyone else.", GREEN, "F0FDF4"));

P(h2("Why that's true"));
P(
  bullet([t("It runs as you, on Google's servers. ", { bold: true }), t("When you click Authorize, you are granting permission to the script — running inside your account — not to me. I receive nothing.")]),
  bullet([t("There's no outside destination for your data. ", { bold: true }), t("The tool reads and writes your own spreadsheet and (optionally) your own Google Contacts. It doesn't phone home.")]),
  bullet([t("You can read every line of the code. ", { bold: true }), t("Open Extensions → Apps Script in the sheet to see exactly what it does.")]),
  bullet([t("You can revoke access anytime. ", { bold: true }), t("Go to myaccount.google.com → Security → Your connections to third-party apps & services, and remove the tool. Your data stays; the automation simply stops.")]),
);

P(h2("Why the “unverified app” warning appears"));
P(p([
  t("That screen shows for personal Apps Script tools that aren't published to Google's Marketplace. It means Google hasn't run a formal brand review — "),
  t("not", { italics: true }),
  t(" that the tool is unsafe. Approving it grants permissions to code operating only within your account."),
]));

P(h2("What each permission is for"));
P(p("The tool asks only for the access it needs to do its job. Here's the map:"));
P(table(
  ["Permission Google shows", "Used for"],
  [
    ["See, edit, create, and delete your spreadsheets", "Reading RAW Data and writing the Master Table, Directory, Contact Log, PCAR, and all derived tabs."],
    ["See, edit, download, and permanently delete your contacts", "The optional “School Directory” phone sync (only if you turn it on)."],
    ["See, create, and delete your Google Drive files", "Saving the PDF archive at year-end rollover, and opening your old sheet during migration."],
    ["Connect to an external service", "A read-only check of the shared Version Tracker so you're told when a newer template exists."],
    ["Run when you're not present / manage triggers", "The automatic triggers: instant Contact Log updates, the nightly refresh/sync, and the update check."],
    ["Display and run third-party web content in dialogs", "The menus, wizards, and sidebars (Settings, Migration, Rollover, progress bars)."],
  ],
  [4100, 5260],
));

P(h2("The only thing that ever leaves your sheet"));
P(
  bullet([t("The update check ", { bold: true }), t("reads a shared, view-only “Version Tracker” spreadsheet to compare dates. It sends nothing about your students — it only reads a list of release dates and links.")]),
  bullet([t("Migration ", { bold: true }), t("opens an older sheet only when you paste its link, and only because you already have access to it.")]),
  bullet([t("Bug reports ", { bold: true }), t("go through a Google Form you choose to fill out — only the words you type are shared.")]),
);

// ============================================================
// EXTRAS
// ============================================================
P(new Paragraph({ children: [new PageBreak()] }));
P(h1("Tips, Workflow & Troubleshooting"));

P(h2("A good daily workflow"));
P(
  numbered("Made a parent contact? Go to the Contact Log, type the OSIS or a name in a new row, and let the Ghost Typist fill the rest. Add your notes."),
  numbered("Need to fix a name, phone, or add a missing student? Use 👤 Student Overrides → New / Edit Student rather than editing RAW Data by hand."),
  numbered("Got a fresh ATS export? Paste it into RAW Data, then run 🏗️ Build Sheets Only."),
  numbered("Doing a big batch of entries? Turn on “Rebuilds: Manual / Nightly Sync Only” in Settings for speed, then run Build once at the end."),
);

P(h2("Handy tips"));
P(
  bullet([t("Siblings are automatic. ", { bold: true }), t("The tool joins siblings by their OSIS numbers and fills Site/Class for each — you don't manage them separately.")]),
  bullet([t("Back-dating works. ", { bold: true }), t("Type your own date, count, or follow-up status and the auto-fill leaves it alone.")]),
  bullet([t("Hide vs. Purge discharged. ", { bold: true }), t("“Hide Discharged” is a reversible view filter; “Purge Discharged Students” permanently deletes them. Use Purge only when you're sure.")]),
  bullet([t("Refresh if a menu looks stale. ", { bold: true }), t("Many actions update the menu live, but pressing F5 always reloads the latest menu.")]),
);

P(h2("Troubleshooting / FAQ"));
P(table(
  ["Symptom", "Fix"],
  [
    ["The custom menus didn't appear", "Give the sheet a few seconds after opening, then press F5. If it's a brand-new copy, run 🚨 Initial Setup first."],
    ["“🔒 Action Blocked” when I type", "Setup hasn't been run yet. Click 🚀 App Menu → 🚨 Initial Setup (Run Once)."],
    ["Ghost Typist: “Could not find … in the Master Table”", "That family isn't compiled yet. Run 🏗️ Build Sheets Only, then try again."],
    ["“📡 Background Sync in Progress” warning", "A rebuild/sync is running. Wait for it to finish before entering new data, to avoid save conflicts."],
    ["Update checker says I'm up to date but I know there's a new version", "Use 🔔 Check for Updates (it bypasses the throttle). If it still looks wrong, the tracker date may be off — report a bug."],
    ["Phone sync created duplicates / didn't finish", "Open Full Sync, use Stop then Restart. The sync mirrors the sheet and cleans up on the next full run."],
    ["Something's broken or confusing", "Use 🚀 App Menu → 🐞 Report a Bug and describe what happened."],
  ],
  [3560, 5800],
));

P(spacer(80));
P(h2("Do & Don't"));
const doDont = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [4680, 4680],
  borders: {
    top: { style: BorderStyle.SINGLE, size: 2, color: LINE }, bottom: { style: BorderStyle.SINGLE, size: 2, color: LINE },
    left: { style: BorderStyle.SINGLE, size: 2, color: LINE }, right: { style: BorderStyle.SINGLE, size: 2, color: LINE },
    insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: LINE }, insideVertical: { style: BorderStyle.SINGLE, size: 2, color: LINE },
  },
  rows: [
    new TableRow({ tableHeader: true, children: [
      cell("✅ Do", { header: true, width: 4680 }),
      cell("🚫 Don't", { header: true, width: 4680 }),
    ]}),
    new TableRow({ children: [
      cell("Paste raw data only into RAW Data.", { width: 4680, shade: "F0FDF4" }),
      cell("Type into Master Table, Directory, or other generated tabs — they're overwritten on build.", { width: 4680, shade: "FEF2F2" }),
    ]}),
    new TableRow({ children: [
      cell("Use Overrides to fix names, phones, or add students.", { width: 4680 }),
      cell("Rename or delete the generated / hidden tabs (including Version).", { width: 4680 }),
    ]}),
    new TableRow({ children: [
      cell("Run Build after big changes.", { width: 4680, shade: "F0FDF4" }),
      cell("Hand-edit the hidden Version tab's saved settings.", { width: 4680, shade: "FEF2F2" }),
    ]}),
    new TableRow({ children: [
      cell("Keep your old / retired sheet as a backup until you've verified the new one.", { width: 4680 }),
      cell("Purge discharged students unless you're certain.", { width: 4680 }),
    ]}),
  ],
});
P(doDont);

P(spacer(120));
P(h2("Glossary"));
P(table(
  ["Term", "Meaning"],
  [
    ["OSIS", "The unique student ID number from ATS. The tool uses it to match students, siblings, and contacts."],
    ["ATS", "The NYC DOE student data system your Raw Data export comes from."],
    ["PCAR", "Parent Coordinator Activity Report — the reporting numbers this tool compiles for you."],
    ["Ghost Typist", "The auto-fill that completes a Contact Log (or Notes) row from a single OSIS or name."],
    ["Override", "A correction or addition you make to student/guardian info without changing raw ATS data."],
    ["Stealth Mode", "An override flag that keeps a student out of the Override Report."],
    ["NYCSA / ParentSquare", "Parent account/communication systems tracked as flags on the Parents Divided tab."],
    ["Build", "Regenerating all derived tabs from your current data (🏗️ Build Sheets Only)."],
  ],
  [1900, 7460],
));

P(spacer(160));
P(new Paragraph({ alignment: AlignmentType.CENTER, border: { top: { color: LINE, size: 8, style: BorderStyle.SINGLE, space: 8 } }, spacing: { before: 120 },
  children: [new TextRun({ text: "Questions or something broken? Use 🚀 App Menu → 🐞 Report a Bug.", italics: true, color: MUTE, size: 20, font: "Calibri" })] }));

// ============================================================
// DOCUMENT
// ============================================================
const doc = new Document({
  creator: "Parent Coordinator Tracker",
  title: "Parent Coordinator Tracker — User Guide",
  description: "User guide for new coordinators: setup, features, migration, and privacy.",
  styles: {
    default: { document: { run: { font: "Calibri", size: 21, color: SLATE } } },
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [
        { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 200 } } } },
        { level: 1, format: LevelFormat.BULLET, text: "◦", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 200 } } } },
      ]},
      { reference: "steps", levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 220 } } } },
      ]},
    ],
  },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, bottom: 1080, left: 1200, right: 1200 } } },
    children: body,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(process.argv[2] || "PC_Tracker_User_Guide.docx", buf);
  console.log("Wrote", process.argv[2]);
});
