// ======================================================================
// Generator for docs/PC_Tracker_User_Guide.docx (the new-user guide).
// This is a documentation helper only — it is NOT part of the Apps Script
// project and never runs on Google's servers.
//
// To regenerate the .docx after editing this file:
//   npm install docx        # one-time; docx is the only dependency
//   node docs/build_guide.js docs/PC_Tracker_User_Guide.docx
//
// Keep docs/USER_GUIDE.md (the Markdown source of truth) and
// docs/guide.html (used to render the clickable PDF) in sync with any
// wording changes here.
// ======================================================================
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  PageBreak, LevelFormat, TableOfContents, Bookmark, InternalHyperlink, ExternalHyperlink,
} = require("docx");

// The Drive folder that holds the shareable template files. Both buttons on the
// "Get Your Copy" page default to this folder; replace them with the exact
// (BLANK) and (MIGRATE) "make a copy" links each release. (Inside the sheet,
// LinkResolver.gs resolves these automatically — this doc is static.)
const TEMPLATE_FOLDER_URL = "https://drive.google.com/drive/folders/1goKFrxDJOSAPxp1HUXETX2dns5c6HXk2";
const BLANK_COPY_URL = TEMPLATE_FOLDER_URL;   // ← paste the (BLANK) copy link here
const MIGRATE_COPY_URL = TEMPLATE_FOLDER_URL; // ← paste the (MIGRATE) copy link here

// ---------- palette ----------
const NAVY = "0F172A";
const BLUE = "1D4ED8";
const SLATE = "334155";
const MUTE = "64748B";
const LINE = "CBD5E1";
const HEADSHADE = "0F2A4A";
const ROWSHADE = "F1F5F9";
const GREEN = "166534";
const RED = "B91C1C";

// ---------- inline helpers ----------
function t(text, o = {}) {
  return new TextRun({ text, bold: o.bold, italics: o.italics, color: o.color || SLATE, size: o.size || 21, font: "Calibri" });
}
function h1(text, anchor) {
  const run = new TextRun({ text, bold: true, color: NAVY, size: 30, font: "Calibri" });
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 140 },
    border: { bottom: { color: BLUE, size: 12, style: BorderStyle.SINGLE, space: 6 } },
    children: anchor ? [new Bookmark({ id: anchor, children: [run] })] : [run],
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 90 },
    children: [new TextRun({ text, bold: true, color: BLUE, size: 24, font: "Calibri" })],
  });
}
function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3, spacing: { before: 160, after: 60 },
    children: [new TextRun({ text, bold: true, color: SLATE, size: 21, font: "Calibri" })],
  });
}
function p(runs, opts = {}) {
  const children = (typeof runs === "string") ? [t(runs)] : runs;
  return new Paragraph({ spacing: { after: opts.after != null ? opts.after : 120 }, alignment: opts.align, children });
}
function bullet(runs, level = 0) {
  const children = (typeof runs === "string") ? [t(runs)] : runs;
  return new Paragraph({ numbering: { reference: "bullets", level }, spacing: { after: 60 }, children });
}
function numbered(runs, ref = "steps") {
  const children = (typeof runs === "string") ? [t(runs)] : runs;
  return new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 80 }, children });
}
function spacer(after = 120) { return new Paragraph({ spacing: { after }, children: [] }); }

// ---------- callout box ----------
function callout(title, bodyText, color = BLUE, shade = "EEF4FF") {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE }, columnWidths: [9360],
    borders: {
      top: { style: BorderStyle.SINGLE, size: 2, color: shade }, bottom: { style: BorderStyle.SINGLE, size: 2, color: shade },
      left: { style: BorderStyle.SINGLE, size: 24, color }, right: { style: BorderStyle.SINGLE, size: 2, color: shade },
      insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE },
    },
    rows: [new TableRow({ children: [new TableCell({
      width: { size: 9360, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: shade, color: "auto" },
      margins: { top: 120, bottom: 120, left: 160, right: 160 },
      children: [
        new Paragraph({ spacing: { after: bodyText ? 40 : 0 }, children: [new TextRun({ text: title, bold: true, color, size: 21, font: "Calibri" })] }),
        ...(bodyText ? [new Paragraph({ children: [new TextRun({ text: bodyText, color: SLATE, size: 20, font: "Calibri" })] })] : []),
      ],
    })] })],
  });
}


// ---------- data tables ----------
function cell(content, { header = false, width, shade, bold = false, color, align } = {}) {
  const paras = (Array.isArray(content) ? content : [content]).map((c) =>
    new Paragraph({ alignment: align, spacing: { after: 0 },
      children: (typeof c === "string")
        ? [new TextRun({ text: c, bold: header || bold, color: header ? "FFFFFF" : (color || SLATE), size: 19, font: "Calibri" })]
        : c,
    })
  );
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: header ? HEADSHADE : (shade || "FFFFFF"), color: "auto" },
    margins: { top: 70, bottom: 70, left: 110, right: 110 }, children: paras,
  });
}
function table(headers, rows, widths) {
  const total = widths.reduce((a, b) => a + b, 0);
  const headerRow = new TableRow({ tableHeader: true, children: headers.map((hdr, i) => cell(hdr, { header: true, width: widths[i] })) });
  const bodyRows = rows.map((r, ri) => new TableRow({ children: r.map((c, ci) => cell(c, { width: widths[ci], shade: ri % 2 ? ROWSHADE : "FFFFFF" })) }));
  return new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: widths,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 2, color: LINE }, bottom: { style: BorderStyle.SINGLE, size: 2, color: LINE },
      left: { style: BorderStyle.SINGLE, size: 2, color: LINE }, right: { style: BorderStyle.SINGLE, size: 2, color: LINE },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: LINE }, insideVertical: { style: BorderStyle.SINGLE, size: 2, color: LINE },
    },
    rows: [headerRow, ...bodyRows],
  });
}

// ---------- jump-to button bar (internal links) ----------
function jumpButton(label, anchor) {
  return new TableCell({
    width: { size: 1872, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: "1D4ED8", color: "auto" },
    margins: { top: 90, bottom: 90, left: 60, right: 60 },
    verticalAlign: "center",
    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
      new InternalHyperlink({ anchor, children: [new TextRun({ text: label, bold: true, color: "FFFFFF", size: 18, font: "Calibri" })] }),
    ]})],
  });
}
function jumpBar() {
  const items = [
    ["① Setup", "part1"], ["② Features", "part2"], ["③ Migration", "part3"],
    ["④ Privacy", "part4"], ["⑤ Tips & FAQ", "part5"],
  ];
  return new Table({
    width: { size: 9360, type: WidthType.DXA }, columnWidths: items.map(() => 1872),
    borders: {
      top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 8, color: "FFFFFF" }, insideVertical: { style: BorderStyle.SINGLE, size: 8, color: "FFFFFF" },
    },
    rows: [new TableRow({ children: items.map(([l, a]) => jumpButton(l, a)) })],
  });
}

// ---------- big external-link "buttons" (Get Your Copy page) ----------
function linkButtonRow(label, url, fill) {
  return new TableRow({ children: [new TableCell({
    width: { size: 9360, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill, color: "auto" },
    margins: { top: 150, bottom: 150, left: 60, right: 60 },
    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
      new ExternalHyperlink({ link: url, children: [new TextRun({ text: label, bold: true, color: "FFFFFF", size: 24, font: "Calibri" })] }),
    ]})],
  })] });
}
function templateButtons(blankUrl, migrateUrl) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360],
    borders: {
      top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 16, color: "FFFFFF" }, insideVertical: { style: BorderStyle.NONE },
    },
    rows: [
      linkButtonRow("📄  New Coordinator — Make a Blank Copy", blankUrl, "0F9D58"),
      linkButtonRow("🔄  Upgrading — Make a Migrate Copy", migrateUrl, "7C3AED"),
    ],
  });
}

// ============================================================
// CONTENT
// ============================================================
const body = [];
const P = (...x) => body.push(...x);

// ---------- TITLE PAGE ----------
P(
  new Paragraph({ spacing: { before: 1300, after: 0 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "🚀", size: 96, font: "Calibri" })] }),
  new Paragraph({ spacing: { before: 200, after: 0 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Parent Coordinator Tracker", bold: true, color: NAVY, size: 56, font: "Calibri" })] }),
  new Paragraph({ spacing: { before: 60, after: 0 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Complete Setup & User Guide", color: BLUE, size: 30, font: "Calibri" })] }),
  new Paragraph({ spacing: { before: 40, after: 0 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Set-up · Buttons & Features · Migration & Updates · Privacy", color: MUTE, size: 20, italics: true, font: "Calibri" })] }),
  new Paragraph({ spacing: { before: 820, after: 0 }, alignment: AlignmentType.CENTER, border: { top: { color: LINE, size: 8, style: BorderStyle.SINGLE, space: 10 } },
    children: [new TextRun({ text: "A Google Sheets tool that turns your ATS export into a live directory, tracks every", color: SLATE, size: 20, font: "Calibri" })] }),
  new Paragraph({ spacing: { after: 0 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "family contact, and compiles your PCAR — automatically.", color: SLATE, size: 20, font: "Calibri" })] }),
  new Paragraph({ spacing: { before: 240, after: 0 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Made by a fellow Parent Coordinator — this is NOT an official DOE sheet.", color: MUTE, size: 18, italics: true, font: "Calibri" })] }),
  new Paragraph({ spacing: { before: 20, after: 0 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Questions or errors? jacob.keller@p94m.org", color: MUTE, size: 18, font: "Calibri" })] }),
  new Paragraph({ children: [new PageBreak()] }),
);

// ---------- GET YOUR COPY (page 2) ----------
P(h1("Get Your Copy"));
P(p([t("Start from one of the two templates below. Each opens a "), t("“Make a copy”", { bold: true }), t(" prompt, so you get your own private sheet in your own Google Drive.")]));
P(spacer(60));
P(templateButtons(BLANK_COPY_URL, MIGRATE_COPY_URL));
P(spacer(100));
P(table(
  ["Template", "Who it's for"],
  [
    ["📄 Blank Copy", "Brand-new coordinators starting fresh — paste your ATS export and go."],
    ["🔄 Migrate Copy", "Coordinators upgrading from an older version — it auto-opens the Migration Wizard to bring last year's data and settings forward."],
  ],
  [2700, 6660],
));
P(spacer(60));
P(callout("Coordinator note — update these two links each release",
  "Both templates live in this Drive folder: " + TEMPLATE_FOLDER_URL + ". Replace the two buttons above with the current (BLANK) and (MIGRATE) “make a copy” links whenever you ship a new version. Inside the sheet, the Welcome dialog and End-of-Year Rollover resolve these automatically from the folder (see LinkResolver.gs) — this printable guide is the one place you swap them by hand.",
  BLUE, "EEF4FF"));
P(new Paragraph({ children: [new PageBreak()] }));

// ---------- CONTENTS ----------
P(h1("Contents"));
P(p([t("Use the buttons below to jump straight to a section (clickable in the PDF and in Google Docs). The table of contents underneath is also clickable.", { italics: true, color: MUTE, size: 19 })]));
P(jumpBar());
P(spacer(120));
P(new TableOfContents("Guide Contents", { hyperlink: true, headingStyleRange: "1-2" }));
P(spacer(160));
P(callout("📄  How to open this as a Google Doc",
  "Upload this .docx to Google Drive (drag it into drive.google.com, or in Google Docs choose File → Open → Upload). Drive converts it into a Google Doc automatically — the table of contents and the Jump-to buttons stay clickable. See “Getting this into Google Drive” at the end for step-by-step options.",
  BLUE, "EEF4FF"));
P(new Paragraph({ children: [new PageBreak()] }));

// ============================================================
// INTRO
// ============================================================
P(h1("What this tool is"));
P(p([
  t("The Parent Coordinator (PC) Tracker is a "), t("Google Sheets add-on", { bold: true }),
  t(" built with Google Apps Script. You paste your student exports into one tab, and the tool automatically builds and maintains everything else: a clean class directory, a per-student contact log, PCAR reporting numbers, and — if you want it — a copy of the directory pushed into your phone through Google Contacts."),
]));
P(p("It's organized into three kinds of sheets:"));
P(
  bullet([t("The Engine — ", { bold: true }), t("RAW Data and Type List, where you paste your system exports and customize your dropdowns.")]),
  bullet([t("The Workspaces — ", { bold: true }), t("Contact Log and Events, where you'll spend most of your day logging calls and workshops.")]),
  bullet([t("The Output — ", { bold: true }), t("automated sheets like Directory, Master Table, Combined Contact Tracking, Parents Divided, PCAR, and Phone Contacts. You never type into these — they build themselves.")]),
);
P(callout("New here? Do the one-time setup in Part 1, load your data, run Build, and you're running in a few minutes.", "", BLUE, "EEF4FF"));

// ============================================================
// PART 1 — SETUP
// ============================================================
P(new Paragraph({ children: [new PageBreak()] }));
P(h1("Part 1 — Setting Up Your Sheet", "part1"));

P(h2("Quick Start (5 steps)"));
P(
  numbered([t("Make your own copy", { bold: true }), t(" of the template from the link you were given (it opens a “Make a copy” prompt, or use File → Make a copy). Everyone works in their own copy.")]),
  numbered([t("Open your copy and wait for the "), t("🚀 App Menu", { bold: true }), t(" to appear at the top. Click "), t("🚀 App Menu → 🚨 Initial Setup (Run Once)", { bold: true }), t(".")]),
  numbered([t("Authorize the script when Google asks (one-time — see “The authorization screen” below).")]),
  numbered([t("Load your data into the "), t("RAW Data", { bold: true }), t(" tab (brand-new users) or use "), t("🔄 Import Old Data", { bold: true }), t(" (upgrading users).")]),
  numbered([t("Run "), t("🏗️ Build Sheets Only", { bold: true }), t(", then press "), t("F5", { bold: true }), t(" to refresh — the full menu unlocks.")]),
);

P(h2("Step 1 — Make a copy and open it"));
P(p("You always start from a shared template. Making a copy drops a fresh, private spreadsheet into your own Google Drive. The template itself is never touched by your work, and no one else can see your copy."));

P(h2("Step 2 — Run Initial Setup (once)"));
P(p([t("Wait for the spreadsheet to finish loading until the "), t("🚀 App Menu", { bold: true }), t(" appears in the top toolbar — don't click anything else until it shows. Then click "), t("🚀 App Menu → 🚨 Initial Setup (Run Once)", { bold: true }), t(". This does the behind-the-scenes wiring:")]));
P(
  bullet("Authorizes the script to run as you."),
  bullet("Installs the automatic triggers — instant updates when you edit the Contact Log, a nightly refresh/sync (around 11 PM), and a periodic “check for updates.”"),
  bullet("Records this copy as your original (a clone-detector so a later copy starts clean)."),
  bullet("Applies the default tab layout and opens this User Guide."),
);
P(callout("Until setup is done, the sheet is intentionally locked.", "If you try to type before running Initial Setup, you'll see “🔒 Action Blocked — Initialization process needs to occur first.” That's expected — just run Initial Setup.", RED, "FEF2F2"));

P(h2("Step 3 — The authorization screen"));
P(p([t("Because this is a personal tool and not a paid Marketplace app, Google shows a warning like "), t("“Google hasn't verified this app.”", { italics: true }), t(" To continue, click "), t("Advanced", { bold: true }), t(" → "), t("Go to … (unsafe)", { bold: true }), t(" → choose your account → "), t("Allow", { bold: true }), t(".")]));
P(p([t("“Unverified” means Google hasn't done a formal brand review — "), t("not", { italics: true }), t(" that anything is wrong. You are granting permission to a script that runs entirely inside "), t("your own", { bold: true }), t(" account.")]));

P(h3("What each authorization actually does"));
P(p("On the Allow screen, Google lists the permissions below. The tool asks only for what it needs — and everything runs inside your own account, so approving it does not give the developer any access."));
P(table(
  ["Permission Google shows", "What it's used for", "Why it's needed"],
  [
    ["See, edit, create, and delete all your Google Sheets spreadsheets", "Reading your RAW Data and writing the Master Table, Directory, Contact Log, PCAR, and every other tab.", "This is the whole engine — without it the tool can't build or update your sheets."],
    ["See, edit, download, and permanently delete your contacts", "The optional “School Directory” phone sync that mirrors parent numbers into your Google Contacts.", "Only used if you turn on phone syncing in Settings. If you never enable it, nothing is written to your contacts."],
    ["See, edit, create, and delete all of your Google Drive files", "Saving the year-end PDF archive to your Drive, and opening your old sheet during a migration.", "Rollover creates a PDF file; migration opens the old spreadsheet you point it at."],
    ["Connect to an external service", "A read-only check of the shared “Version Tracker” sheet so you're told when a newer template ships.", "The only outside call the tool makes. It reads release dates/links — never your data."],
    ["Run when you're not present / manage its own triggers", "The automatic triggers: instant Contact Log updates, the nightly rebuild/sync (~11 PM), and the update check on open.", "Lets the sheet keep itself current without you clicking Build every time."],
    ["Display and run third-party web content in dialogs and sidebars", "The menus, wizards, and sidebars — Settings, Migration, Rollover, and the progress bars.", "These custom screens are how you drive the tool."],
  ],
  [3000, 3400, 2960],
));
P(callout("Bottom line on permissions", "Every scope above operates only within your own Google account. There is no server or database belonging to the developer — see Part 4 for the full privacy picture.", GREEN, "F0FDF4"));

P(h2("Step 4 — Load your data (the RAW Data tab)"));
P(p([t("Before logging calls, the system needs to know who your students are. Export data from your standard systems and paste it into the brightly-colored “PASTE” boxes on the "), t("RAW Data", { bold: true }), t(" tab.")]));
P(callout("🚨 CRITICAL — always paste as values only", "When pasting into RAW Data, use Edit → Paste Special → Values Only (Ctrl+Shift+V). Pasting normally can drag in formatting or formulas that break the build.", RED, "FEF2F2"));
P(h3("Insight / ATS data — Blue section → paste into cell A5 (required)"));
P(
  numbered("Go to TeachHub and sign in."),
  numbered("In the search bar, look up and open Insight."),
  numbered("Go to Tools → Downloader."),
  numbered("Under “Data Available for Download,” select only Guardian Contact Information, then download."),
  numbered([t("Copy the downloaded data and paste it into "), t("cell A5", { bold: true }), t(" (Paste Special → Values Only).")]),
);
P(h3("ParentSquare data — Orange section → paste into cell AG6 (optional)"));
P(
  numbered("Go to ParentSquare and sign in, then open Admin."),
  numbered("Under Data Assistant, select Parents."),
  numbered("Click Export CSV."),
  numbered([t("Copy the data and paste it into "), t("cell AG6", { bold: true }), t(".")]),
);
P(h3("NYCSA data — Green section → paste into cell AP2 (optional)"));
P(
  numbered("Go to TeachHub and sign in."),
  numbered("Open Family Access Management."),
  numbered("Export the “Students with Non-NYCSA Account” report."),
  numbered([t("Copy the data and paste it into "), t("cell AP2", { bold: true }), t(".")]),
);
P(h3("Missing or wrong data? Use the New/Edit Student override"));
P(p([t("District databases like ATS can lag behind real enrollment. If you need to add a student or fix contact info before it hits ATS, use "), t("👤 Student Overrides → New / Edit Student", { bold: true }), t(". When you build, the script checks your overrides first and applies them across the entire directory. (More in Part 2.)")]));
P(h3("Customizing dropdowns — the Type List tab"));
P(p([t("The optional "), t("Type", { bold: true }), t(" column in the Contact Log is driven by the "), t("Type List", { bold: true }), t(" tab. Edit that tab to set your own conversation-topic categories (for example: Attendance, Behavior, Enrollment, IEP).")]));

P(h2("Step 5 — Build the sheets"));
P(p([t("Run "), t("🚀 App Menu → 🏗️ Build Sheets Only", { bold: true }), t(". A progress sidebar opens and generates every derived tab. The "), t("first", { italics: true }), t(" successful build unlocks the full menu — a toast will ask you to press "), t("F5", { bold: true }), t(" to refresh and reveal the new options.")]));
P(callout("The golden rule of this sheet", "RAW Data is the only tab you paste into. Everything else is generated — so re-running Build always safely rebuilds it. Never hand-type into Master Table, Directory, or other generated tabs; your edits would be overwritten on the next build.", GREEN, "F0FDF4"));

// ============================================================
// PART 2 — BUTTONS & FEATURES
// ============================================================
P(new Paragraph({ children: [new PageBreak()] }));
P(h1("Part 2 — Buttons & Features", "part2"));
P(p([t("After setup, three custom menus sit at the top of the screen, to the right of “Help”: "), t("🚀 App Menu", { bold: true }), t(", "), t("👤 Student Overrides", { bold: true }), t(", and "), t("📆 Events", { bold: true }), t(".")]));

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
P(p([t("Use the Events tab for large-scale attendances (workshops, graduations, fairs) so they don't clutter your daily Contact Log. Click "), t("📆 Events → 📋 New Event", { bold: true }), t(". A pop-up with a calendar picker asks for the event name, type/category (with a "), t("PTC Quick Fill", { bold: true }), t(" button for “Parent Teacher Conference”), and location, then builds an attendance table. Log attendees by OSIS — the records map back to each student's Combined Contact Tracking and tally automatically in PCAR.")]));

P(spacer(60));
P(h2("The tabs — what each one is for"));
P(table(
  ["Tab", "Purpose", "Visible by default?"],
  [
    ["RAW Data", "Your export paste zone — the only tab you paste raw data into (ATS / ParentSquare / NYCSA).", "Yes"],
    ["Type List", "Drives the optional Type dropdown in the Contact Log. Edit it to set your own categories.", "Yes"],
    ["Contact Log", "The heart of the tool. Log every guardian contact here (deep dive below).", "Yes"],
    ["Master Table", "The generated master roster: OSIS, student, guardian, status, and more. Data starts on row 5.", "Yes (generated)"],
    ["Directory", "A clean, printable class directory merging ATS, ParentSquare, and NYCSA data.", "Yes (generated)"],
    ["Combined Contact Tracking", "One comprehensive history block per student — all contact notes and event attendances together. Great before an IEP meeting.", "Yes (generated)"],
    ["PCAR", "Your auto-compiled PCAR reporting numbers (Phone, Walk-In, Workshops, PTCs, totals).", "Yes (generated)"],
    ["Phone Contacts", "The exact data that gets pushed to Google Contacts.", "Hidden (optional)"],
    ["Parents Divided", "One row per guardian (with ParentSquare / NYCSA flags). Handy for mail merges.", "Hidden (optional)"],
    ["Notes", "An optional manual notes tab with its own auto-fill and duplicate protection.", "Hidden (optional)"],
    ["Send Out", "Optional tab tied to extra PCAR rows.", "Hidden (optional)"],
    ["Version / Backend_Event_Log", "Behind-the-scenes tabs (build date, saved settings, event data). Leave them alone.", "Hidden (system)"],
  ],
  [2280, 5080, 2000],
));

P(spacer(80));
P(h2("The Contact Log, in depth"));
P(p("This is your primary workspace for phone calls, emails, and 1-on-1 meetings. It's designed so you type as little as possible."));
P(h3("The Pre-Call Search Bar"));
P(p("At the top of the Contact Log is a search bar. Before you call a family, type a student's name or OSIS to instantly pull up your most recent interaction with them — so you walk into the call with context."));
P(h3("Logging a new interaction (left → right)"));
P(table(
  ["Column", "What to put there"],
  [
    ["Date", "The date of the interaction. Auto-fills to today if you leave it blank; type your own to back-date."],
    ["Find Student", "Type ONE of three things — the parent's name, the OSIS, or the student's name. The Ghost Typist fills in the rest of the profile."],
    ["Person Spoke With", "The specific parent/guardian you actually spoke to."],
    ["#", "Number of interactions, for PCAR counting. A normal call = 1; a mass email to 30 parents = 30. Auto-fills to 1 if blank."],
    ["Site & Class", "(Optional) Auto-fills from the student's profile; can be hidden in Settings."],
    ["📱🚶 Category", "Categorize for PCAR — e.g., Phone, Walk-In, Parent-Teacher."],
    ["Method", "The exact method — Call, Email, In School, etc."],
    ["Type", "(Optional) A custom topic dropdown driven by the Type List tab; can be hidden in Settings."],
    ["Notes", "Detailed notes about the conversation."],
    ["⁉️ Follow-Up Needed", "A visual flag: Yes · Yes – Important · No · Waiting on Someone · Completed. Auto-fills to “No” if blank."],
    ["Followup Notes", "Next steps or what you did after the initial contact (e.g., “Teacher will reach out,” cc'd the teacher)."],
    ["Hide", "Check this box to hide the row from view while keeping its data active for PCAR counting."],
  ],
  [2600, 6760],
));
P(spacer(60));
P(p("Three things the Contact Log does for you automatically:"));
P(
  bullet([t("Ghost Typist. ", { bold: true }), t("Type an OSIS, student name, or guardian name and the other identity columns (plus Site and Class, including for siblings) fill in. If it can't match what you typed, it warns “Could not find … in the Master Table.”")]),
  bullet([t("Smart auto-fill. ", { bold: true }), t("Entering real content stamps today's Date, sets # to 1, and Follow-up to “No” — but only where you left them blank. Anything you typed yourself is never overwritten.")]),
  bullet([t("Auto-adding rows + the master hide filter. ", { bold: true }), t("A fresh blank row is always kept above the dark “END” bar, and the checkbox on the END-bar row toggles hiding of your hidden rows so you can focus on active ones.")]),
);
P(callout("If the Ghost Typist says “Could not find … in the Master Table”", "That family isn't compiled yet. Run 🏗️ Build Sheets Only once so your latest RAW Data / overrides are included, then try again.", BLUE, "EEF4FF"));

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
P(h2("Build vs. Full Sync — know the difference"));
P(table(
  ["Command", "What it is"],
  [
    ["🏗️ Build / Rebuild Data", "An internal action. Takes your RAW Data and constructs all the automated sheets (Directory, Master Table, Parents Divided, Phone Contacts, PCAR). It only organizes data inside the spreadsheet."],
    ["🧱 Full Sync with Google Contacts", "An external action. Takes the formatted Phone Contacts tab and pushes it into your actual Google Contacts via a “School Directory” label — one phone number per card — so a parent's name, relationship, and student show on your phone's caller ID."],
  ],
  [3200, 6160],
));
P(p([t("The phone sync runs in “mirror” mode — creating, editing, and removing contacts so the label always matches your directory. It runs nightly or on demand via "), t("🧱 Full Sync", { bold: true }), t(", with a progress sidebar you can pause, stop, or restart; it's resilient and resumes if it hits Google's time limits.")]));

// ============================================================
// PART 3 — MIGRATION & UPDATES
// ============================================================
P(new Paragraph({ children: [new PageBreak()] }));
P(h1("Part 3 — Migration & Updates", "part3"));

P(h2("Importing data from an older sheet (Migration Wizard)"));
P(p([t("Open "), t("🚀 App Menu → 🔄 Import Old Data", { bold: true }), t(". (On an “upgrade” copy of the template it opens by itself.) Paste the link to your older PC sheet and click "), t("Start Safe Migration", { bold: true }), t(". The wizard runs six steps:")]));
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
P(p([t("Each copy remembers its own build date. Quietly, on open (at most once every few hours), the sheet checks a shared "), t("Version Tracker", { bold: true }), t(" to see whether a newer template has shipped. If so, it pops a short changelog with a link to the newest copy. You can also check any time with "), t("🚀 App Menu → 🔔 Check for Updates", { bold: true }), t(".")]));
P(p("Inside that dialog you can:"));
P(
  bullet([t("Dismiss", { bold: true }), t(" — hide this update until something even newer ships.")]),
  bullet([t("Major updates only", { bold: true }), t(" — automatic pop-ups only for big releases (a manual check still shows everything).")]),
  bullet([t("Stop automatic update checks", { bold: true }), t(" — silence the on-open checks; the manual menu check always still works.")]),
);
P(h3("How to actually adopt an update"));
P(p([t("Updates don't rewrite your current sheet in place. To move up a version: open the new template link, "), t("make a copy", { bold: true }), t(", run Initial Setup, then use the "), t("Migration Wizard", { bold: true }), t(" to bring your data across from your current sheet. Your settings come along automatically.")]));

P(h2("End of Year Rollover"));
P(p([t("When the year ends, run "), t("🚀 App Menu → 🔄 End of Year Rollover", { bold: true }), t(". It:")]));
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
P(h1("Part 4 — Security & Privacy", "part4"));
P(callout("The short version: I cannot see your sheet, your students, your contacts, or your Drive.",
  "This tool has no server and no database of mine. The code runs entirely inside your own Google account. Nothing about your families is ever sent to me or anyone else. Once you make a copy, it's yours.", GREEN, "F0FDF4"));

P(h2("Why that's true"));
P(
  bullet([t("It runs as you, on Google's servers. ", { bold: true }), t("When you click Authorize, you are granting permission to the script — running inside your account — not to me. I receive nothing.")]),
  bullet([t("There's no outside destination for your data. ", { bold: true }), t("The tool reads and writes your own spreadsheet and (optionally) your own Google Contacts. It doesn't phone home.")]),
  bullet([t("You can read every line of the code. ", { bold: true }), t("Open Extensions → Apps Script in the sheet to see exactly what it does.")]),
  bullet([t("You can revoke access anytime. ", { bold: true }), t("Go to myaccount.google.com → Security → Your connections to third-party apps & services, and remove the tool. Your data stays; the automation simply stops.")]),
);

P(h2("Why the “unverified app” warning appears"));
P(p([t("That screen shows for personal Apps Script tools that aren't published to Google's Marketplace. It means Google hasn't run a formal brand review — "), t("not", { italics: true }), t(" that the tool is unsafe. Approving it grants permissions to code operating only within your account. For a permission-by-permission breakdown, see "), t("“What each authorization actually does”", { bold: true }), t(" in Part 1.")]));

P(h2("The only things that ever leave your sheet"));
P(
  bullet([t("The update check ", { bold: true }), t("reads a shared, view-only “Version Tracker” spreadsheet to compare dates. It sends nothing about your students — it only reads a list of release dates and links.")]),
  bullet([t("Migration ", { bold: true }), t("opens an older sheet only when you paste its link, and only because you already have access to it.")]),
  bullet([t("Bug reports ", { bold: true }), t("go through a Google Form you choose to fill out — only the words you type are shared.")]),
);
P(callout("Who can see your copy?", "Entirely up to you. Once you make a copy it lives in your Drive; you decide who (if anyone) gets shared access, exactly like any other Google Sheet.", BLUE, "EEF4FF"));

// ============================================================
// PART 5 — TIPS, WORKFLOW & TROUBLESHOOTING
// ============================================================
P(new Paragraph({ children: [new PageBreak()] }));
P(h1("Tips, Workflow & Troubleshooting", "part5"));

P(h2("A good daily workflow"));
P(
  numbered("Before a call, use the Contact Log search bar to pull up the family's history."),
  numbered("Made a contact? In a new row, type the OSIS or a name and let the Ghost Typist fill the rest. Add your notes and follow-up flag."),
  numbered("Need to fix a name/phone or add a missing student? Use 👤 Student Overrides → New / Edit Student rather than editing RAW Data by hand."),
  numbered("Got a fresh ATS export? Paste it into RAW Data (Paste Special → Values Only), then run 🏗️ Build Sheets Only."),
  numbered("Doing a big batch? Turn on “Rebuilds: Manual / Nightly Sync Only” in Settings for speed, then run Build once at the end."),
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
    ["Pasted data looks broken / directory is wrong", "Re-paste using Edit → Paste Special → Values Only, then run 🏗️ Build Sheets Only."],
    ["Ghost Typist: “Could not find … in the Master Table”", "That family isn't compiled yet. Run 🏗️ Build Sheets Only, then try again."],
    ["“📡 Background Sync in Progress” warning", "A rebuild/sync is running. Wait for it to finish before entering new data, to avoid save conflicts."],
    ["Update checker says I'm up to date but there's a new version", "Use 🔔 Check for Updates (it bypasses the throttle). If it still looks wrong, report a bug."],
    ["Phone sync created duplicates / didn't finish", "Open Full Sync, use Stop then Restart. The sync mirrors the sheet and cleans up on the next full run."],
    ["Something's broken or confusing", "Use 🚀 App Menu → 🐞 Report a Bug, or email jacob.keller@p94m.org."],
  ],
  [3560, 5800],
));

P(spacer(80));
P(h2("Do & Don't"));
const doDont = new Table({
  width: { size: 9360, type: WidthType.DXA }, columnWidths: [4680, 4680],
  borders: {
    top: { style: BorderStyle.SINGLE, size: 2, color: LINE }, bottom: { style: BorderStyle.SINGLE, size: 2, color: LINE },
    left: { style: BorderStyle.SINGLE, size: 2, color: LINE }, right: { style: BorderStyle.SINGLE, size: 2, color: LINE },
    insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: LINE }, insideVertical: { style: BorderStyle.SINGLE, size: 2, color: LINE },
  },
  rows: [
    new TableRow({ tableHeader: true, children: [cell("✅ Do", { header: true, width: 4680 }), cell("🚫 Don't", { header: true, width: 4680 })] }),
    new TableRow({ children: [cell("Paste raw data only into RAW Data — as Values Only.", { width: 4680, shade: "F0FDF4" }), cell("Type into Master Table, Directory, or other generated tabs — they're overwritten on build.", { width: 4680, shade: "FEF2F2" })] }),
    new TableRow({ children: [cell("Use Overrides to fix names, phones, or add students.", { width: 4680 }), cell("Rename or delete the generated / hidden tabs (including Version).", { width: 4680 })] }),
    new TableRow({ children: [cell("Run Build after big changes.", { width: 4680, shade: "F0FDF4" }), cell("Hand-edit the hidden Version tab's saved settings.", { width: 4680, shade: "FEF2F2" })] }),
    new TableRow({ children: [cell("Keep your old / retired sheet as a backup until you've verified the new one.", { width: 4680 }), cell("Purge discharged students unless you're certain.", { width: 4680 })] }),
  ],
});
P(doDont);

P(spacer(120));
P(h2("Glossary"));
P(table(
  ["Term", "Meaning"],
  [
    ["OSIS", "The unique student ID number from ATS. The tool uses it to match students, siblings, and contacts."],
    ["ATS / Insight", "The NYC DOE student data systems your Raw Data export comes from (via TeachHub)."],
    ["PCAR", "Parent Coordinator Activity Report — the reporting numbers this tool compiles for you."],
    ["Ghost Typist", "The auto-fill that completes a Contact Log (or Notes) row from a single OSIS or name."],
    ["Override", "A correction or addition you make to student/guardian info without changing raw ATS data."],
    ["Stealth Mode", "An override flag that keeps a student out of the Override Report."],
    ["NYCSA / ParentSquare", "Parent account/communication systems whose exports feed the directory and Parents Divided tab."],
    ["Build", "Regenerating all derived tabs from your current data (🏗️ Build Sheets Only)."],
  ],
  [1900, 7460],
));

// ---------- Getting into Google Drive ----------
P(new Paragraph({ children: [new PageBreak()] }));
P(h1("Getting this into Google Drive"));
P(p("You can keep and share this guide as a Google Doc in a couple of ways:"));
P(h3("Option A — Upload and convert (recommended)"));
P(
  numbered("Go to drive.google.com and open the folder where you want it."),
  numbered("Drag the PC_Tracker_User_Guide.docx file into the window (or click New → File upload)."),
  numbered("Right-click the uploaded file → Open with → Google Docs. Drive converts it to a Google Doc; the Jump-to buttons and table of contents stay clickable."),
  numbered("(Optional) Delete the original .docx once the Google Doc is created."),
);
P(h3("Option B — From inside Google Docs"));
P(
  numbered("Open docs.google.com and start a blank document."),
  numbered("Choose File → Open → Upload, and drop in the .docx."),
);
P(h3("Prefer a clickable PDF?"));
P(p("A companion PDF (PC_Tracker_User_Guide.pdf) is provided with a “Jump to” button bar and a clickable table of contents that jump to each section — handy for reading on a phone or sharing a read-only copy."));
P(spacer(160));
P(new Paragraph({ alignment: AlignmentType.CENTER, border: { top: { color: LINE, size: 8, style: BorderStyle.SINGLE, space: 8 } }, spacing: { before: 120 },
  children: [new TextRun({ text: "Questions or something broken? Use 🚀 App Menu → 🐞 Report a Bug, or email jacob.keller@p94m.org.", italics: true, color: MUTE, size: 20, font: "Calibri" })] }));

// ============================================================
// DOCUMENT
// ============================================================
const doc = new Document({
  creator: "Parent Coordinator Tracker",
  title: "Parent Coordinator Tracker — Complete Setup & User Guide",
  description: "User guide for new coordinators: setup, features, migration, and privacy.",
  styles: { default: { document: { run: { font: "Calibri", size: 21, color: SLATE } } } },
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
  console.log("Wrote", process.argv[2] || "PC_Tracker_User_Guide.docx");
});
