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
// the Welcome dialog and Rollover resolve these automatically — this doc is static.)
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
function numbered(runs, ref = "steps", instance) {
  const children = (typeof runs === "string") ? [t(runs)] : runs;
  // instance: a list with its own instance number starts again at 1
  const numbering = instance != null ? { reference: ref, level: 0, instance } : { reference: ref, level: 0 };
  return new Paragraph({ numbering, spacing: { after: 80 }, children });
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
  "Both templates live in this Drive folder: " + TEMPLATE_FOLDER_URL + ". Replace the two buttons above with the current (BLANK) and (MIGRATE) “make a copy” links whenever you ship a new version. Inside the sheet, the Welcome dialog and End-of-Year Rollover resolve these automatically from the folder — this printable guide is the one place you swap them by hand.",
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
  bullet([t("The Engine — ", { bold: true }), t("Raw Data and the Backend tab's Type List, where you paste your system exports and customize your dropdowns.")]),
  bullet([t("The Workspaces — ", { bold: true }), t("Contact Log and Events, where you'll spend most of your day logging calls and workshops.")]),
  bullet([t("The Output — ", { bold: true }), t("automated sheets like Directory, Student List, Combined Contact Tracking, Parents Divided, PCAR, and Phone Contacts. You never type into these — they build themselves.")]),
);
P(callout("New here? Do the one-time setup in Part 1, load your data, run Build, and you're running in a few minutes.", "", BLUE, "EEF4FF"));

// ============================================================
// PART 1 — SETUP
// ============================================================
P(new Paragraph({ children: [new PageBreak()] }));
P(h1("Part 1 — Setting Up Your Sheet", "part1"));
P(h2("Quick Start (4 steps, about 10 minutes)"));
P(
  numbered([t("Make your own copy", { bold: true }), t(" of the template (it opens a \"Make a copy\" prompt, or use "), t("File → Make a copy", { italics: true }), t("). Your copy opens on the "), t("👋 Welcome", { bold: true }), t(" tab, the only tab showing.")], "steps", 11),
  numbered([t("Read what Google will ask.", { bold: true }), t(" The Welcome tab (and Step 2 below) explains each permission before you click anything.")], "steps", 11),
  numbered([t("Click 🚀 App Menu → 🚨 Initial Setup (Run Once)", { bold: true }), t(" and allow it: "), t("OK", { bold: true }), t(" → your account → "), t("Advanced", { bold: true }), t(" → "), t("Go to … (unsafe)", { bold: true }), t(" → tick "), t("Select all", { bold: true }), t(" → "), t("Continue", { bold: true }), t(".")], "steps", 11),
  numbered([t("Follow your ✅ Getting Started checklist", { bold: true }), t(" (it opens on the right): Load your student data → Build your sheets → Check your settings → Log your first parent contact.")], "steps", 11),
);
P(callout("Moving from last year's tracker?", "Run Initial Setup first, then use 🚀 App Menu → 🔄 Import Old Data (an upgrade copy opens it by itself). See Part 3.", BLUE, "EEF4FF"));
P(h2("Step 1 — Make a copy and open it"));
P(p([t("You always start from a shared template. Making a copy drops a fresh, private spreadsheet into your own Google Drive. The template is never touched by your work, and no one else can see your copy.")]));
P(p([t("Your copy opens on the "), t("👋 Welcome", { bold: true }), t(" tab. Until setup is done it is the only tab showing: it explains each permission Google will ask for, lists the setup steps, and links to this guide and the bug report form. A message at the bottom of the screen also points you to "), t("🚀 App Menu → 🚨 Initial Setup (Run Once)", { bold: true }), t(". Initial Setup removes the Welcome tab and brings your tabs back.")]));
P(h2("Step 2 — Know what Google will ask (permissions)"));
P(p([t("Initial Setup asks Google for permission once. Google shows its own screens first (the tool can't show anything of its own until you allow it), so the Welcome tab explains them before you click. Here is the same list, in Google's words:")]));
P(table(
  ["Permission Google shows", "What the tool uses it for"],
  [
    ["See, edit, create and delete all your Google Sheets spreadsheets", "Builds and updates the tabs in this tracker. It also reads the shared Version Tracker sheet to tell you when a newer version is out, and opens last year's tracker when you use 🔄 Import Old Data, which is why it needs more than this one sheet."],
    ["See, edit, download and permanently delete your contacts", "Only the phone sync (⚙️ Settings → Contacts & sync), which stays off until you turn it on. It only adds, changes and removes contacts in its own \"School Directory\" label."],
    ["See, edit, create and delete all of your Google Drive files", "🔄 End of Year Rollover saves a PDF archive next to your tracker. It also finds the newest User Guide and template links in the shared folder."],
    ["Connect to an external service", "Only that end-of-year PDF: it downloads the PDF from Google's own export page. Nothing about your families is sent anywhere."],
    ["Allow this application to run when you are not present", "The automatic parts: instant updates when you type in the Contact Log, the nightly rebuild (and phone sync, if on) at 10 PM, and the update check when you open the sheet."],
    ["Display and run third-party web content in prompts and sidebars inside Google applications", "The tool's own pop-ups and side panels: Settings, Input Data Mode, Getting Started, the build progress. \"Third-party\" just means not made by Google."]
  ],
  [3000, 6360],
));
P(callout("Bottom line on permissions", "Every permission above works only inside your own Google account. There is no server or database belonging to the developer — see Part 4.", GREEN, "F0FDF4"));
P(h2("Step 3 — Run Initial Setup and allow it"));
P(
  numbered([t("Wait until "), t("🚀 App Menu", { bold: true }), t(" appears at the top of the screen, next to Help. It takes a few seconds; not there after a minute? Reload the page.")], "steps", 12),
  numbered([t("Click "), t("🚀 App Menu → 🚨 Initial Setup (Run Once)", { bold: true }), t(".")], "steps", 12),
  numbered([t("Authorization required:", { bold: true }), t(" click "), t("OK", { bold: true }), t(", then choose your Google account.")], "steps", 12),
  numbered([t("Google hasn't verified this app:", { bold: true }), t(" click "), t("Advanced", { bold: true }), t(", then "), t("Go to … (unsafe)", { bold: true }), t(". \"Unverified\" means Google hasn't done a formal brand review of a tool made by a person — "), t("not", { italics: true }), t(" that anything is wrong.")], "steps", 12),
  numbered([t("On the list of permissions, tick "), t("Select all", { bold: true }), t(", then click "), t("Continue", { bold: true }), t(". Setup needs all of them; if one is left unticked, Google asks again for the missing ones.")], "steps", 12),
  numbered([t("Nothing happened after that? Click "), t("🚀 App Menu → 🚨 Initial Setup (Run Once)", { bold: true }), t(" again.")], "steps", 12),
);
P(p([t("What Initial Setup does:")]));
P(
  bullet([t("Installs the automatic triggers: instant Contact Log updates, the nightly rebuild/sync (around 10 PM) and the update check when you open the sheet.")]),
  bullet([t("Records this copy as your original (so a later copy of it starts clean).")]),
  bullet([t("Removes the Welcome tab and shows your everyday tabs: Contact Log, Combined Contact Tracking, PCAR and Directory (turn on more in ⚙️ Settings → Tabs & columns).")]),
  bullet([t("Opens the welcome pop-up and your "), t("✅ Getting Started", { bold: true }), t(" checklist.")]),
);
P(callout("🔒 Until setup is done, the sheet is intentionally locked.", "If you type before running Initial Setup, you'll see \"🔒 Action Blocked — Initialization process needs to occur first.\" That's expected.", RED, "FEF2F2"));
P(h2("Step 4 — Follow the Getting Started checklist"));
P(p([t("The checklist opens on the right after setup and ticks itself off as you go:")]));
P(
  numbered([t("Load your student data", { bold: true }), t(" — opens 📥 Input Data Mode (Step 5).")], "steps", 13),
  numbered([t("Build your sheets", { bold: true }), t(" — Input Data Mode does this for you (Step 6).")], "steps", 13),
  numbered([t("Check your settings", { bold: true }), t(" — opens ⚙️ Settings (Part 2).")], "steps", 13),
  numbered([t("Log your first parent contact", { bold: true }), t(" — takes you to the Contact Log (Part 2).")], "steps", 13),
);
P(p([t("It opens by itself each time you open the sheet until all four are done (untick "), t("Open this by itself", { bold: true }), t(" to stop that), and you can always reopen it from "), t("🚀 App Menu → ✅ Getting Started", { bold: true }), t(". Under the steps, "), t("This tab", { bold: true }), t(" says what the tab you're on is for, and "), t("Hidden tabs", { bold: true }), t(" lists the tabs you can turn on in ⚙️ Settings → Tabs & columns.")]));
P(h2("Step 5 — Load your data (📥 Input Data Mode)"));
P(p([t("Before logging calls, the system needs to know who your students are. Open "), t("🚀 App Menu → 📥 Input Data Mode", { bold: true }), t(" (or "), t("Open Input Data Mode", { bold: true }), t(" in the checklist) and the sheet walks you through the three exports one at a time, then two optional steps: "), t("Site / Class", { bold: true }), t(" (each class's teacher and site) and "), t("Your Lists", { bold: true }), t(" (below). Your other tabs are hidden while you work. For each export, either click "), t("📂 Upload file", { bold: true }), t(" and pick the file you downloaded, or open the file, press Ctrl+A (Mac: ⌘+A), copy (Ctrl+C / ⌘+C), and paste anywhere on the temporary tab with Ctrl+Shift+V (Mac: ⌘+Shift+V). The data moves into Raw Data by itself, and the sidebar shows when each one was last updated. It warns you, without changing Raw Data, if you pasted the wrong export or if Excel cut numbers off (like 7.19E+09). Use "), t("Skip", { bold: true }), t(" to keep the old data for an export. Before the Your Lists step, it builds your sheets for you. "), t("Finish or exit", { bold: true }), t(" ends it at any step: "), t("Save, exit and build my sheets", { bold: true }), t(" saves what's on the tab, brings your tabs back and builds. The end screen offers "), t("Build sheets now", { bold: true }), t(" only if something changed since that build, and "), t("Next steps", { bold: true }), t(" back to your Getting Started checklist.")]));
P(h3("Student Data (ATS / Insight) — required"));
P(
  numbered([t("Go to TeachHub and sign in.")], "steps", 14),
  numbered([t("In the search bar, look up and open Insight.")], "steps", 14),
  numbered([t("Go to Tools → Downloader.")], "steps", 14),
  numbered([t("Under \"Data Available for Download,\" select only "), t("Guardian Contact Information", { bold: true }), t(", then download. Upload or paste it in Input Data Mode.")], "steps", 14),
);
P(h3("Parent Square — optional"));
P(
  numbered([t("Go to ParentSquare, sign in, and open Admin.")], "steps", 15),
  numbered([t("Under Data Assistant, select Parents.")], "steps", 15),
  numbered([t("Click Export CSV. Upload or paste it in Input Data Mode.")], "steps", 15),
);
P(h3("NYCSA — optional"));
P(
  numbered([t("Go to TeachHub and sign in.")], "steps", 16),
  numbered([t("Open Family Access Management.")], "steps", 16),
  numbered([t("Export the \"Students with Non-NYCSA Account\" report. Upload or paste it in Input Data Mode.")], "steps", 16),
);
P(h3("Pasting straight into Raw Data instead"));
P(p([t("You can also paste each export into its colored box on the "), t("Raw Data", { bold: true }), t(" tab (show it in ⚙️ Settings → Tabs & columns): Student Data into cell "), t("A5", { bold: true }), t(", Parent Square into "), t("AG6", { bold: true }), t(", NYCSA into "), t("AP2", { bold: true }), t(".")]));
P(callout("🚨 Pasting into Raw Data? Always paste values only.", "Use Edit → Paste Special → Values Only (Ctrl+Shift+V). Pasting normally can drag in formatting or formulas that break the build. (Input Data Mode does this for you.)", RED, "FEF2F2"));
P(h3("Your own lists (Input Data Mode, Your Lists step)"));
P(p([t("Have a list that isn't an ATS export, like bus routes, after-school sign-ups, PTA members or trip forms? On the "), t("Your Lists", { bold: true }), t(" step, click "), t("➕ Add another list", { bold: true }), t(", name it, and paste the whole sheet (headers too) into the yellow box on the new "), t("📥 (name)", { bold: true }), t(" tab, or upload the file (CSV or Excel). The sidebar then takes you through four short steps:")]));
P(
  numbered([t("Columns.", { bold: true }), t(" It finds the header row by itself (title rows above it are fine) and guesses what each column is, shown like "), t("OSIS: Column C “OSIS”", { italics: true }), t(". Fix a guess with its dropdown, and say whether each row is a "), t("student", { bold: true }), t(" or a "), t("guardian", { bold: true }), t(". A name can be one column (“LAST, FIRST” is fine) or two (first + last).")], "steps", 17),
  numbered([t("Variables.", { bold: true }), t(" The other columns (a route, a fee, a form status). Tick "), t("keep", { bold: true }), t(", give it a name, and choose what a blank cell shows and what everyone who isn't on the list gets: nothing, or your own text (like “No bus”). Then choose where it shows: under "), t("Additional Information", { bold: true }), t(" on the phone contact (above Communications), and/or as a column in the "), t("Student List", { bold: true }), t(" (student lists) or "), t("Parents Divided", { bold: true }), t(" (guardian lists).")], "steps", 17),
  numbered([t("Match.", { bold: true }), t(" It works out who each row is: OSIS first, then names, then phone or email, then best guesses for short names and nicknames (Sam → Samuel). Rows it isn't sure of are listed for you: confirm a best guess, pick one of several possible students, search the whole Student List, or skip the row. "), t("✓ Accept all best guesses", { bold: true }), t(" takes them all at once; check a couple first, because best guesses can be wrong.")], "steps", 17),
  numbered([t("View.", { bold: true }), t(" How the kept values read on the phone contact and in the Student List / Parents Divided: one per line or all on one line, names in bold or not, with a live preview. Click "), t("Save", { bold: true }), t(" here.")], "steps", 17),
);
P(p([t("Your picks are remembered: when you get a new version of the list, click "), t("Paste new version", { bold: true }), t(" and only the new rows need you. The list's tab shows who each row matched (✔ Matched / How). What you keep reaches the Student List, Parents Divided and the phone contacts at the next build; Combined Contact Tracking doesn't change. Add as many lists as you like.")]));
P(callout("⚠️ Anything shown under Additional Information goes into your Google Contacts.", "It is on the guardian's contact (and your phone) after the next sync. Keep private lists (IEP, housing, health) off, or keep them on their own tab only.", RED, "FEF2F2"));
P(h3("Missing or wrong data? Use the New/Edit Student override"));
P(p([t("District databases like ATS can lag behind real enrollment. If you need to add a student or fix contact info before it hits ATS, use "), t("👤 Student Overrides → New / Edit Student", { bold: true }), t(". When you build, your overrides are applied to the Student List, Combined Contact Tracking, Parents Divided, Phone Contacts, and the Ghost Typist. The "), t("Directory", { bold: true }), t(" tab and the build report's new / discharged students always show ATS (Raw Data) only. Overrides also tidy themselves up: a \"New Student\" override is removed once ATS has that student, an edit that ATS now matches exactly is removed too, and a student's class always follows Raw Data once they're in ATS. (More in Part 2.)")]));
P(h3("Customizing dropdowns — the Type List"));
P(p([t("The optional "), t("Type", { bold: true }), t(" column in the Contact Log is driven by the Type List in the hidden "), t("Backend", { bold: true }), t(" tab (column BL). Unhide Backend and edit that column to set your own conversation-topic categories (for example: Attendance, Behavior, Enrollment, IEP).")]));
P(h2("Step 6 — Build the sheets"));
P(p([t("Input Data Mode builds your sheets for you. To build again later (after new data, overrides or lists), run "), t("🚀 App Menu → 🏗️ Build Sheets Only", { bold: true }), t(". A progress sidebar generates every derived tab and checks off each step as it goes (reading your data, Student List, Event Log, Combined Contact Tracking, Parents Divided, Directory, Phone Contacts, finishing up), with a line under each saying what it builds. When it finishes, the sidebar shows a short report — new students, discharged students, and families newly signed up for NYCSA or ParentSquare (click a count to see the names). The "), t("first", { italics: true }), t(" successful build adds the full menus right away: ⚙️ Settings, 👤 Student Overrides and 📆 Events (don't see them? reload the page with F5).")]));
P(callout("The golden rule of this sheet", "Raw Data is the only tab you paste into. Everything else is generated — so re-running Build always safely rebuilds it. Never hand-type into Student List, Directory, or other generated tabs.", GREEN, "F0FDF4"));

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
    ["🏗️ Build Sheets Only (No Push)", "Your everyday “refresh everything” button. Regenerates all derived tabs from Raw Data + Contact Log + overrides, without touching your phone contacts."],
    ["⚙️ Settings", "Opens the Settings dialog — show/hide tabs and columns, control live updates, filter notes, manage discharged students (full list below)."],
    ["✅ Getting Started", "The four-step checklist for a new sheet (Part 1, Step 4)."],
    ["📖 Open User Guide", "A short welcome with the next steps and a button that opens this guide."],
    ["🐞 Report a Bug", "Opens a short Google Form so you can report anything broken or confusing."],
    ["📥 Input Data Mode", "Walks you through loading each export, then Site / Class and Your Lists, one step at a time, and builds your sheets (Part 1, Step 5)."],
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
    ["✨ New / Edit Student", "Add a student who isn't in ATS, or correct a name, phone, email, language, grade, or up to 5 guardians (a student's class always follows Raw Data once they're in ATS). You can also set a status Label and turn on “Hide from Report (Stealth Mode).” Saving opens the Build sidebar and rebuilds the sheets (if rebuilds are paused in Settings, it just saves)."],
    ["🗂️ Manage Overrides", "See every override you've made; edit or delete any of them."],
    ["📄 Override Report", "Builds a side-by-side “Students Override Report” comparing the raw ATS data against your overrides. Stealth-mode students are skipped."],
  ],
  [3300, 6060],
));

P(spacer(60));
P(h2("📆 Events"));
P(p([t("Use the Events tab for large-scale attendances (workshops, graduations, fairs) so they don't clutter your daily Contact Log. Click "), t("📆 Events → 📋 New Event", { bold: true }), t(". A pop-up with a calendar picker asks for the event name, type (buttons for PTC, Workshop, Family event and Meeting fill in the Type box), and location, then builds an attendance table. Log attendees by OSIS, student name, or guardian name — type one, paste a whole column, or put several OSIS numbers in one cell separated by commas — and the Ghost Typist fills in the rest. The records map back to each student's Combined Contact Tracking and tally automatically in PCAR.")]));

P(spacer(60));
P(h2("The tabs — what each one is for"));
P(table(
  ["Tab", "Purpose", "Visible by default?"],
  [
    ["Raw Data", "Your export paste zone — the only tab you paste raw data into (ATS / ParentSquare / NYCSA).", "Hidden (show in Settings)"],
    ["👋 Welcome", "The first page a new copy opens on, with the permissions and the setup steps. Removed by Initial Setup.", "Only before setup"],
    ["Contact Log", "The heart of the tool. Log every guardian contact here (deep dive below).", "Yes"],
    ["Student List", "The generated roster (called “Master Table” in older copies): OSIS, student, guardian, site, class, siblings, status, the date of the last interaction, and a column for each student list you keep (Your Lists). Data starts on row 5.", "Hidden (generated)"],
    ["Directory", "A clean, printable class directory built from your Raw Data (ATS): each student's site, class/teacher, and parents' email and phone. Overrides and discharged students aren't included.", "Yes (generated)"],
    ["Combined Contact Tracking", "One comprehensive history block per student — all contact notes and event attendances together. Great before an IEP meeting. Only students on the Student List get a block, so notes logged under a mistyped or unknown OSIS don't show here.", "Yes (generated)"],
    ["PCAR", "Your auto-compiled PCAR reporting numbers (Phone, Walk-In, Workshops, PTCs, totals).", "Yes (generated)"],
    ["Phone Contacts", "The exact data pushed to Google Contacts, including the Additional Information lines from Your Lists.", "Hidden (optional)"],
    ["Parents Divided", "One row per guardian (with ParentSquare / NYCSA flags, and a column for each guardian list you keep). Handy for mail merges.", "Hidden (optional)"],
    ["Notes", "An optional manual notes tab (one row per family) with its own auto-fill and duplicate protection. Its notes show at the top of that student's Combined Contact Tracking.", "Hidden (optional)"],
    ["Send Out", "Optional tab tied to extra PCAR rows.", "Hidden (optional)"],
    ["📥 Your lists", "One tab per list you add in Input Data Mode (Your Lists step): your pasted list, plus ✔ Matched / How columns showing who each row is.", "Yes (yours)"],
    ["Backend", "Behind-the-scenes sections side by side: Version & saved settings (A:E), New/Edit Student overrides (G:AV), event log (AX:BF), Send Out list (BH:BJ), the Type List that drives the Contact Log Type dropdown (BL), Additional Information (BN:BQ, rewritten by every build) and the Your Lists setups (BS:BV). Only edit the Type List.", "Hidden (system)"],
  ],
  [2280, 5080, 2000],
));

P(spacer(80));
P(h2("The Contact Log, in depth"));
P(p("This is your primary workspace for phone calls, emails, and 1-on-1 meetings. It's designed so you type as little as possible."));
P(h3("The Pre-Call Search Bar"));
P(p("Below the dark END bar at the bottom of the Contact Log is a search bar. Before you call a family, type a student's name or OSIS to pull up their Communications (every note so far) — so you walk into the call with context."));
P(h3("Logging a new interaction (left → right)"));
P(table(
  ["Column", "What to put there"],
  [
    ["Date", "The date of the interaction. Auto-fills to today if you leave it blank; type your own to back-date."],
    ["Guardian / OSIS / Student", "Type ONE of three things — the parent's name (column C), the OSIS (E), or the student's name (F). The Ghost Typist fills in the other two."],
    ["Person Spoke With", "The specific parent/guardian you actually spoke to."],
    ["#", "Number of interactions, for PCAR counting. A normal call = 1; a mass email to 30 parents = 30. Auto-fills to 1 if blank."],
    ["Site & Class", "(Optional) Auto-fills from the student's profile; can be hidden in Settings."],
    ["📱🚶 Category", "Categorize for PCAR — e.g., Phone, Walk-In, Parent-Teacher."],
    ["Method", "The exact method — Call, Email, In School, etc."],
    ["Type", "(Optional) A custom topic dropdown driven by the Backend tab's Type List; can be hidden in Settings."],
    ["Event Description", "Detailed notes about the conversation."],
    ["Follow-up", "No Followup Needed · Followup Needed · Followup Needed - ASAP · Followed up / Resolved · Skipped / Newer Interaction · Waiting for Response. Auto-fills to “No Followup Needed” if blank."],
    ["Additional Notes", "Next steps or what you did after the initial contact (e.g., “Teacher will reach out,” cc'd the teacher)."],
    ["Hide", "Check this box to keep this note out of the compiled notes (Combined Contact Tracking and Phone Contacts). The row stays in the Contact Log and still counts for PCAR."],
  ],
  [2600, 6760],
));
P(spacer(60));
P(p("Three things the Contact Log does for you automatically:"));
P(
  bullet([t("Ghost Typist. ", { bold: true }), t("Type — or paste a whole column of — OSIS numbers, student names, or guardian names, and the other identity columns (plus Site and Class, including for siblings) fill in for every row. If it can't match what you typed, it warns “Could not find … in the Student List” (after a paste, it lists the ones it couldn't find).")]),
  bullet([t("Smart auto-fill. ", { bold: true }), t("Entering real content stamps today's Date, sets # to 1, and Follow-up to “No Followup Needed” — but only where you left them blank. Anything you typed yourself is never overwritten.")]),
  bullet([t("Auto-adding rows + the master hide filter. ", { bold: true }), t("A fresh blank row is always kept above the dark “END” bar. The checkbox on the END-bar row turns on a filter that shows only this and last week's rows (or only today's — see Settings), plus anything still marked for follow-up, so you can focus on active ones.")]),
);
P(callout("If the Ghost Typist says “Could not find … in the Student List”", "That family isn't compiled yet. Run 🏗️ Build Sheets Only once so your latest Raw Data / overrides are included, then try again.", BLUE, "EEF4FF"));

P(spacer(80));
P(h2("⚙️ Settings — every choice explained"));
P(p([t("Settings has five sections on the left. Click "), t("Save & apply", { bold: true }), t(" to keep your changes.")]));
P(table(
  ["Section → setting", "What it controls"],
  [
    ["Contacts & sync → Sync contacts to my phone every night", "Adds the school directory to your Google Contacts during the nightly run (and the 🧱 Full Sync menu item appears)."],
    ["Tabs & columns → Show these tabs", "Show or hide any tab (every tab except Backend). Contact Log, Combined Contact Tracking, PCAR and Directory start shown; the rest start hidden. Hidden tabs keep updating."],
    ["Tabs & columns → Contact Log columns", "Shows / hides the Site & Class and Type columns in the Contact Log."],
    ["Live updates → Rebuild sheets", "When the Student List, Directory and other tabs refresh: after every edit and nightly (default), nightly and when you click Build, after every edit but not nightly, or only when you click Build. Turning off the nightly run also stops the nightly phone sync."],
    ["Live updates → Contact Log notes reach the other tabs", "As you type (default), or only when you click Build."],
    ["Live updates → Fill in Date, # and Follow-up for me", "The Contact Log auto-fill. The parent / student / OSIS lookup works either way."],
    ["Notes → Contact Log shows", "This week and last week (default), or today only. Important notes and open follow-ups always show."],
    ["Notes → Compiled notes cover", "All time, the last month / 6 months / year, or custom dates."],
    ["Notes → Include notes I've hidden", "Adds rows you checked off with the row checkbox back into the compiled notes."],
    ["Student data → Hide discharged students in the Student List", "Only hides the rows. Turn it off to see them again."],
    ["Student data → Remind me to update my data", "A pop-up when Student Data, Parent Square or NYCSA hasn't been updated in a week (default), 2 weeks, a month, or never."],
    ["Student data → Purge discharged students", "Permanently deletes non-current students from the Student List. You can't undo this."],
    ["Reset to defaults", "Returns every setting to its default."],
  ],
  [3560, 5800],
));
P(callout("Settings travel with your copy.", "Your choices are quietly saved into the hidden Backend tab, so if you upgrade to a newer template later, a fresh copy inherits your settings automatically.", GREEN, "F0FDF4"));

P(spacer(80));
P(h2("Build vs. Full Sync — know the difference"));
P(table(
  ["Command", "What it is"],
  [
    ["🏗️ Build / Rebuild Data", "An internal action. Takes your Raw Data and constructs all the automated sheets (Directory, Student List, Parents Divided, Phone Contacts, PCAR). It only organizes data inside the spreadsheet."],
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
  bullet([t("Dismiss for now", { bold: true }), t(" — close it; it pops up again on the next check.")]),
  bullet([t("Dismiss this update", { bold: true }), t(" — hide this update until something even newer ships.")]),
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
  numbered("Need to fix a name/phone or add a missing student? Use 👤 Student Overrides → New / Edit Student rather than editing Raw Data by hand."),
  numbered("Got a fresh ATS export? Paste it into Raw Data (Paste Special → Values Only), then run 🏗️ Build Sheets Only."),
  numbered("Doing a big batch? Turn on “Rebuilds: Manual / Nightly Sync Only” in Settings for speed, then run Build once at the end."),
);

P(h2("Handy tips"));
P(
  bullet([t("Siblings are automatic. ", { bold: true }), t("The tool joins siblings by their OSIS numbers and fills Site/Class for each — you don't manage them separately.")]),
  bullet([t("Back-dating works. ", { bold: true }), t("Type your own date, count, or follow-up status and the auto-fill leaves it alone.")]),
  bullet([t("Hide vs. Purge discharged. ", { bold: true }), t("“Hide Discharged” is a reversible view filter; “Purge Discharged Students” permanently deletes them. Use Purge only when you're sure.")]),
  bullet([t("Refresh if a menu looks stale. ", { bold: true }), t("Many actions update the menu live, but pressing F5 always reloads the latest menu.")]),
  bullet([t("Fresh-data reminder. ", { bold: true }), t("If your ATS, ParentSquare or NYCSA data in Raw Data is older than your reminder setting (a week by default), the sheet shows a reminder once a day when you open it. Change it or turn it off in ⚙️ Settings → Student data.")]),
  bullet([t("Lists match best with OSIS. ", { bold: true }), t("If a list you add has an OSIS column, every row with an OSIS matches exactly; names alone need more checking.")]),
);

P(h2("Troubleshooting / FAQ"));
P(table(
  ["Symptom", "Fix"],
  [
    ["The custom menus didn't appear", "Give the sheet a few seconds after opening, then press F5. If it's a brand-new copy, run 🚨 Initial Setup first."],
    ["I only see the Welcome tab", "Setup hasn't been run on this copy yet. Click 🚀 App Menu → 🚨 Initial Setup (Run Once); it removes the Welcome tab and shows your tabs."],
    ["Google keeps asking for permission", "Setup needs every permission on the list. Tick Select all, then Continue (what each one is for: Part 1, Step 2)."],
    ["“🔒 Action Blocked” when I type", "Setup hasn't been run yet. Click 🚀 App Menu → 🚨 Initial Setup (Run Once)."],
    ["Pasted data looks broken / directory is wrong", "Re-paste using Edit → Paste Special → Values Only, then run 🏗️ Build Sheets Only."],
    ["Ghost Typist: “Could not find … in the Student List”", "That family isn't compiled yet. Run 🏗️ Build Sheets Only, then try again."],
    ["“📡 Background Sync in Progress” warning", "A rebuild/sync is running. Wait for it to finish before entering new data, to avoid save conflicts."],
    ["Update checker says I'm up to date but there's a new version", "Use 🔔 Check for Updates (it bypasses the throttle). If it still looks wrong, report a bug."],
    ["Phone sync created duplicates / didn't finish", "Open Full Sync, use Stop then Restart. The sync mirrors the sheet and cleans up on the next full run."],
    ["A row on my list says “not found”", "That student (or guardian) isn't on the Student List. A new student? Add them with 👤 New / Edit Student, build, then open the list again (📥 Input Data Mode → Your Lists → Edit setup). Or pick them by hand with 🔍 Pick from the Student List."],
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
    new TableRow({ children: [cell("Paste raw data only into Raw Data — as Values Only.", { width: 4680, shade: "F0FDF4" }), cell("Type into Student List, Directory, or other generated tabs — they're overwritten on build.", { width: 4680, shade: "FEF2F2" })] }),
    new TableRow({ children: [cell("Use Overrides to fix names, phones, or add students.", { width: 4680 }), cell("Rename or delete the generated / hidden tabs (including Backend).", { width: 4680 })] }),
    new TableRow({ children: [cell("Run Build after big changes.", { width: 4680, shade: "F0FDF4" }), cell("Hand-edit the hidden Backend tab's saved settings.", { width: 4680, shade: "FEF2F2" })] }),
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
    ["Ghost Typist", "The auto-fill that completes a Contact Log, Notes, or Events row from a single OSIS or name."],
    ["Override", "A correction or addition you make to student/guardian info without changing raw ATS data."],
    ["Stealth Mode", "An override flag that keeps a student out of the Override Report."],
    ["NYCSA / ParentSquare", "Parent account/communication systems whose exports show each family's sign-up status in Parents Divided and Phone Contacts."],
    ["Build", "Regenerating all derived tabs from your current data (🏗️ Build Sheets Only)."],
    ["Your Lists", "Your own lists (bus routes, PTA…) added in Input Data Mode's last step. Each row is matched to a student or guardian; what you keep shows under Additional Information and as Student List / Parents Divided columns."],
    ["Additional Information", "The lines from Your Lists on a phone contact, above Communications. Kept in the Backend tab and rewritten by every build."],
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
