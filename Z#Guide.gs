// ======================================================================
// FILE: Z#Guide.gs  (+ Z#GuideSidebar.html)
// PURPOSE: Optional "✅ Getting Started" guide for new parent coordinators
// (Z# = trial file; nothing relies on it). Without these files every hook
// below is a no-op and the sheet starts up exactly as before.
//
//   - Welcome pop-up right after Initial Setup (new-user copies): "Start with
//     Input Data Mode", plus links to the guide document / video.
//   - Getting Started checklist (sidebar): load data (Input Data Mode) -> build ->
//     settings -> first parent contact. Each step ticks itself from the sheet.
//     Opens by itself on every open until all four are done, unless
//     "Open this by itself" is unticked. Closing it is always fine.
//   - "This tab" panel in the same sidebar: what the tab you're on is for, and
//     which tabs are hidden and where to turn them on.
//
// Hooks in shared code (each guarded with typeof, a no-op without this file):
//   Menus.gs buildMenus_            addGuideMenuItems_(menu)   (both 🚀 App Menus)
//   Menus.gs runInitialSetupWrapper showGuideWelcome_()        (instead of showWelcomeDialog)
//   Menus.gs checkVersionOnOpen     guideOnOpen_()              (installable on open)
//
// State (document properties; they don't travel with copies, so every new copy
// starts fresh): ZG_AUTO_OPEN ("off"), ZG_SETTINGS_SEEN ("true").
// ======================================================================

// ----------------------------------------------------------------------
// LINKS — PLACEHOLDERS. Paste the links here when they're ready.
// An empty link shows as "link coming soon".
// ----------------------------------------------------------------------
const GUIDE_LINKS = {
  guideDoc: "",          // DOCUMENT: the getting-started guide (what each tab does, first steps)
  downloadDataVideo: "", // VIDEO: how to download your data (ATS, Parent Square, NYCSA)
  downloadDataDoc: "",   // DOCUMENT: how to download your data, step by step
  buildDoc: "",          // DOCUMENT: what Build Sheets does
  settingsDoc: "",       // DOCUMENT: what each setting does
  contactLogDoc: ""      // DOCUMENT: how to log a parent contact
};

/**
 * What each tab is for, for the "This tab" panel. type: "data" | "type" (you
 * type here) | "report" | "built" (filled in for you). settings = shown or
 * hidden in ⚙️ Settings → Tabs & columns.
 */
const ZG_TABS = [
  { name: "Raw Data", alt: "RAW Data", type: "data", settings: true,
    what: "Where your student data lives: ATS (columns A to AE), Parent Square (AG to AN), NYCSA (AP to BF), and your sites and classes (BH to BI). Row 1 of each part shows when it was last updated. Input Data Mode puts the data in for you." },
  { name: "Contact Log", type: "type", settings: true,
    what: "Your daily log of every parent contact. Type the parent's name in column C and the student and OSIS fill in, along with the date, the # and the follow-up." },
  { name: "Notes", type: "type", settings: true,
    what: "Notes about a student that you always want to see first. They show at the top of that student's notes in Combined Contact Tracking." },
  { name: "Events", type: "type", settings: true,
    what: "Attendance for your events and workshops. Make a table with 📆 Events → 📋 New Event, then add the parents who came." },
  { name: "Send Out", type: "type", settings: true,
    what: "Messages and flyers you sent to families and how many you reached. They count toward your PCAR." },
  { name: "PCAR", type: "report", settings: true,
    what: "Your PCAR report, filled in for you from the other tabs. Only H2:I2 are yours to change." },
  { name: "Student List", type: "built", settings: true,
    what: "One row per student with guardians, siblings, site, class and status. Rebuilt on every build, so don't type here." },
  { name: "Combined Contact Tracking", type: "built", settings: true,
    what: "Every student's notes in one place: Contact Log entries, events and Notes-tab notes, newest first." },
  { name: "Parents Divided", type: "built", settings: true,
    what: "One row per guardian, with contact details and Parents Square / NYCSA sign-up." },
  { name: "Directory", type: "built", settings: true,
    what: "Your family directory, from ATS only: each student, OSIS, site, class and how to reach their parents." },
  { name: "Phone Contacts", type: "built", settings: true,
    what: "The contact list your phone sync sends to Google Contacts (⚙️ Settings → Contacts & sync)." },
  { name: "Students Override Report", type: "built", settings: true,
    what: "The student changes you made with 👤 Student Overrides." }
];

// ======================================================================
// HOOKS CALLED FROM SHARED CODE
// ======================================================================

/** 🚀 App Menu item (both the getting-started and the full menu). */
function addGuideMenuItems_(menu) {
  menu.addItem("✅ Getting Started", "openGuideChecklist");
}

/** Right after Initial Setup on a new-user copy: the checklist, then the welcome pop-up on top. */
function showGuideWelcome_() {
  openGuideChecklist();
  const hasInputMode = typeof openInputDataMode === "function";
  const link = (url, label) => url
    ? `<a class="btn sec" href="${url}" target="_blank" rel="noopener">${label}</a>`
    : `<span class="btn sec" aria-disabled="true" title="Link coming soon">${label} (coming soon)</span>`;

  const html = include_('Styles') + `
    <style>
      .note-dlg { padding: 8px 12px 12px; display: flex; flex-direction: column; gap: 14px; }
      .note-dlg p { line-height: 1.5; }
      .note-dlg ol { margin: 0; padding-left: 20px; display: flex; flex-direction: column; gap: 4px; }
      .note-dlg .foot { display: flex; justify-content: flex-end; gap: 8px; flex-wrap: wrap; }
      .note-dlg .sub { border-top: 1px solid var(--line); padding-top: 14px; display: flex; flex-direction: column; gap: 10px; }
      .note-dlg .sub .foot { justify-content: flex-start; }
    </style>
    <div class="note-dlg">
      <h1>Welcome to the PC Tracker</h1>
      <p>This tracker keeps your parent contacts, events and family directory in one place, and fills in your PCAR for you. Getting started takes four steps:</p>
      <ol>
        <li><b>Load your student data</b> (Input Data Mode walks you through it)</li>
        <li><b>Build your sheets</b></li>
        <li><b>Check your settings</b></li>
        <li><b>Log your first parent contact</b></li>
      </ol>
      <p class="help" style="margin:0">Your checklist is open on the right. It opens by itself until you've done all four.</p>
      <div class="foot">
        <button type="button" class="btn sec" onclick="google.script.host.close()">Not now</button>
        ${hasInputMode ? `<button type="button" class="btn pri" id="goBtn" onclick="go()">Start with Input Data Mode</button>` : ``}
      </div>
      <div class="sub">
        <p class="help" style="margin:0">Before you start: how to download your data from ATS, Parent Square and NYCSA.</p>
        <div class="foot">
          ${link(GUIDE_LINKS.downloadDataVideo, "Watch the video")}
          ${link(GUIDE_LINKS.guideDoc, "Read the guide")}
        </div>
      </div>
      <p class="hint">Moving from last year's tracker? Use 🚀 App Menu → 🔄 Import Old Data instead.</p>
    </div>
    <script>
      function go() {
        const b = document.getElementById('goBtn');
        b.disabled = true; b.textContent = 'Opening…';
        google.script.run
          .withSuccessHandler(function () { google.script.host.close(); })
          .withFailureHandler(function (e) { alert('Could not open Input Data Mode: ' + e.message); b.disabled = false; b.textContent = 'Start with Input Data Mode'; })
          .zg_action('inputData');
      }
    </script>`;
  const out = HtmlService.createHtmlOutput(html).setWidth(480).setHeight(hasInputMode ? 440 : 410);
  SpreadsheetApp.getUi().showModalDialog(out, "Welcome to the PC Tracker");
}

/** Installable on open (checkVersionOnOpen): opens the checklist until its four steps are done. */
function guideOnOpen_() {
  const props = PropertiesService.getDocumentProperties();
  if (props.getProperty('SETUP_COMPLETE') !== 'true') return;
  if (props.getProperty('INPUT_DATA_MODE_STATE')) return; // Input Data Mode has the sidebar
  if (props.getProperty('ZG_AUTO_OPEN') === 'off') return;
  if (zg_checklist_(SpreadsheetApp.getActiveSpreadsheet()).allDone) return;
  openGuideChecklist();
}

// ======================================================================
// MENU HANDLER + SIDEBAR ENDPOINTS
// ======================================================================

function openGuideChecklist() {
  const html = htmlFromFile_('Z#GuideSidebar').setTitle('Getting Started');
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Everything the sidebar shows. */
function zg_getState() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getDocumentProperties();
  return {
    checklist: zg_checklist_(ss),
    autoOpen: props.getProperty('ZG_AUTO_OPEN') !== 'off',
    tab: zg_activeTab(),
    hiddenTabs: zg_hiddenTabs_(ss),
    links: GUIDE_LINKS,
    userGuideUrl: zg_safeGuideLink_(),
    hasInputMode: typeof openInputDataMode === "function"
  };
}

/** The tab you're on (polled by the sidebar every few seconds; one cheap call). */
function zg_activeTab() {
  const name = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  const info = ZG_TABS.find(t => t.name === name || t.alt === name);
  return info ? { name: name, known: true, type: info.type, what: info.what } : { name: name, known: false };
}

/** Checklist buttons. Some open another sidebar or a pop-up. */
function zg_action(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  switch (name) {
    case "inputData":
      if (typeof openInputDataMode !== "function") throw new Error("Input Data Mode isn't installed on this sheet.");
      openInputDataMode();
      return true;
    case "importOld":
      openMigrationWizard();
      return true;
    case "build":
      runBuildOnly();
      return true;
    case "settings":
      PropertiesService.getDocumentProperties().setProperty('ZG_SETTINGS_SEEN', 'true');
      openSettingsDialog();
      return true;
    case "contactLog": {
      const sheet = ss.getSheetByName("Contact Log");
      if (!sheet) throw new Error("There's no Contact Log tab.");
      if (sheet.isSheetHidden()) sheet.showSheet();
      sheet.activate();
      const end = findEndRowInColumnE_(sheet);
      let row = 2;
      if (end && end > 2) {
        const c = sheet.getRange(2, 3, end - 2, 1).getDisplayValues();
        const blank = c.findIndex(r => String(r[0]).trim() === "");
        row = blank === -1 ? end - 1 : 2 + blank;
      }
      sheet.getRange(row, 3).activate();
      return true;
    }
  }
  throw new Error("Unknown step: " + name);
}

/** "Open this by itself when the sheet opens". */
function zg_setAutoOpen(on) {
  const props = PropertiesService.getDocumentProperties();
  if (on) props.deleteProperty('ZG_AUTO_OPEN'); else props.setProperty('ZG_AUTO_OPEN', 'off');
  return !!on;
}

// ======================================================================
// CHECKLIST
// ======================================================================

function zg_checklist_(ss) {
  const props = PropertiesService.getDocumentProperties();
  const data = zg_dataStatus_(ss);
  const steps = [
    { id: "data", title: "Load your student data", done: data.ats.loaded,
      text: "Input Data Mode walks you through it: paste or upload your ATS, Parent Square and NYCSA exports, then your sites and classes.",
      parts: [
        { label: "Student data (ATS)", done: data.ats.loaded, date: data.ats.date },
        { label: "Parent Square", done: data.ps.loaded, date: data.ps.date },
        { label: "NYCSA", done: data.nycsa.loaded, date: data.nycsa.date },
        { label: "Sites and classes", done: data.siteClass }
      ],
      buttons: [{ action: "inputData", label: "Open Input Data Mode", primary: true, needsInputMode: true },
                { action: "importOld", label: "Import last year's tracker" }],
      links: [{ label: "Video: how to download your data", url: GUIDE_LINKS.downloadDataVideo },
              { label: "Guide: how to download your data", url: GUIDE_LINKS.downloadDataDoc }] },
    { id: "build", title: "Build your sheets", done: zg_isBuilt_(ss),
      text: "Fills the Student List, Directory, Combined Contact Tracking and the other tabs from your data. The rest of the menus appear after the first build.",
      buttons: [{ action: "build", label: "Build Sheets Only", primary: true }],
      links: [{ label: "Guide: what Build Sheets does", url: GUIDE_LINKS.buildDoc }] },
    { id: "settings", title: "Check your settings", done: props.getProperty('ZG_SETTINGS_SEEN') === 'true',
      text: "Choose which tabs to show, when sheets rebuild, phone sync, and how the Contact Log behaves.",
      buttons: [{ action: "settings", label: "Open Settings" }],
      links: [{ label: "Guide: what each setting does", url: GUIDE_LINKS.settingsDoc }] },
    { id: "contact", title: "Log your first parent contact", done: zg_hasContact_(ss),
      text: "Type the parent's name in column C of the Contact Log. The student and OSIS fill in for you.",
      buttons: [{ action: "contactLog", label: "Go to the Contact Log" }],
      links: [{ label: "Guide: how to log a parent contact", url: GUIDE_LINKS.contactLogDoc }] }
  ];
  return { steps: steps, doneCount: steps.filter(s => s.done).length, allDone: steps.every(s => s.done) };
}

/** Raw Data row 1 stamps ("ATS - DATE UPDATED: 11/04/26"; "00/00/00" = never loaded). */
function zg_dataStatus_(ss) {
  const raw = ss.getSheetByName("Raw Data") || ss.getSheetByName("RAW Data");
  const read = cell => {
    if (!raw) return { loaded: false, date: "" };
    const text = String(raw.getRange(cell).getDisplayValue());
    const m = text.match(/(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/);
    const loaded = !!m && !/00\/00\/00/.test(text);
    return { loaded: loaded, date: loaded ? m[1] : "" };
  };
  let siteClass = !!PropertiesService.getDocumentProperties().getProperty('INPUT_DATA_MODE_SITECLASS_UPDATED');
  if (!siteClass && raw && raw.getMaxColumns() >= 61 && raw.getMaxRows() >= 7) {
    siteClass = raw.getRange(7, 60, 1, 2).getDisplayValues()[0].join("").trim() !== "";
  }
  return { ats: read("A1"), ps: read("AG1"), nycsa: read("AP1"), siteClass: siteClass };
}

/** Same test the menus use (a student under the Student List header). */
function zg_isBuilt_(ss) {
  return studentListIsBuilt_(ss);
}

/** Any parent name in Contact Log column C above the END bar. */
function zg_hasContact_(ss) {
  const sheet = ss.getSheetByName("Contact Log");
  if (!sheet) return false;
  const end = findEndRowInColumnE_(sheet);
  const last = end ? end - 1 : sheet.getLastRow();
  if (last < 2) return false;
  return sheet.getRange(2, 3, last - 1, 1).getDisplayValues().some(r => String(r[0]).trim() !== "");
}

/** The guide's tabs that are hidden right now, and where to turn each one on. */
function zg_hiddenTabs_(ss) {
  return ZG_TABS.map(info => ({ info: info, sheet: ss.getSheetByName(info.name) || (info.alt ? ss.getSheetByName(info.alt) : null) }))
    .filter(t => t.sheet && t.sheet.isSheetHidden())
    .map(t => ({
      name: t.sheet.getName(),
      what: t.info.what,
      where: t.info.settings ? "⚙️ Settings → Tabs & columns" : "View → Hidden sheets → " + t.sheet.getName()
    }));
}

function zg_safeGuideLink_() {
  try { return typeof getGuideLink_ === "function" ? getGuideLink_() : ""; } catch (e) { return ""; }
}
