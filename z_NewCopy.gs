// ======================================================================
// FILE: z_NewCopy.gs  (owner only, like z_MINE.gs; owner's request 2026-11-12)
// PURPOSE: When you open a NEW COPY of your own sheet, a pop-up lists what to
// change in it (the copy-related part of DEPLOY.md). The same list, with tick
// boxes, is in 🐞 Debug → 📋 New Copy Checklist.
//
// How it knows it is a copy: Google does not copy a sheet's saved settings, so
// the first open of a copy can't tell by itself. It asks once: "Is this your
// MAIN sheet?" Yes = remembered as your main sheet (never asked again there).
// No = a copy: the checklist shows (once per copy). To skip the question on your
// main sheet, paste its ID below (the part of its address between /d/ and /edit).
//
// Hooked in only from z_MINE.gs (ownerAddMenus_, both checks guarded with typeof):
// no shared file mentions it, and the checklist itself says to delete it on
// copies for schools / the master template (with z_MINE and the other z_ files).
// The automatic pop-up is a plain alert because a new copy has only the simple
// onOpen trigger, which can't open a styled dialog.
// ======================================================================

var OWNER_HOME_SHEET_ID_ = ""; // optional: your main sheet's ID

var NEW_COPY_HOME_KEY_ = "OWNER_HOME_SHEET_ID";        // doc property: the main sheet's ID (set when you answer Yes)
var NEW_COPY_SHOWN_KEY_ = "OWNER_COPY_CHECKLIST_SHOWN"; // doc property: this copy's ID once the pop-up was shown
var NEW_COPY_TICKS_KEY_ = "OWNER_NEW_COPY_TICKS";       // doc property: ticked steps in the checklist dialog

/** The checklist: one group per kind of copy (also the text of the pop-up). */
var NEW_COPY_STEPS_ = [
  { id: "school", title: "A copy for a SCHOOL or the MASTER TEMPLATE", steps: [
    { id: "wipe",     text: "Wipe your data: 🐞 Debug → 🧹 Reset to Display Mode. (Next time, make the copy with 🐞 Debug → 📄 Copy Sheet, Then Reset the Copy: it is wiped already.) Do this first: the 🐞 Debug menu goes away with z_MINE." },
    { id: "files",    text: "Extensions → Apps Script: delete z_MINE, z_ContactLogDebug, z_ContactLogDebugSidebar, z_InputDataDebug and z_NewCopy (this checklist)." },
    { id: "manifest", text: "Project Settings → tick \"Show appsscript.json\", then replace appsscript.json with \"App Scripts (General).Json\" (yours has the email permissions)." },
    { id: "tabs",     text: "Check your own tabs are gone (the reset removes them): Old Master Table, Old Phone Contacts, New Students, 📥 lists." },
    { id: "nosetup",  text: "Don't run Initial Setup on it: each school runs it on its own copy (that makes the automatic jobs under their account)." },
    { id: "b3",       text: "Master template only: Backend cell B3 = No. There is one template for everyone now: its Initial Setup asks whether to start fresh or bring over last year's tracker (B3 = Yes is only for an old upgrade copy)." },
    { id: "header",   text: "Master template only: Student List header on row 1, frozen, like your sheet." }
  ] },
  { id: "test", title: "A TEST copy for you", steps: [
    { id: "keep",     text: "Keep the z_ files and the \"App Scripts (Mine)\" manifest." },
    { id: "setup",    text: "Run 🚀 App Menu → 🚨 Initial Setup (Run Once): a copy has no automatic jobs until then. Phone sync starts off: leave it off on a test copy." },
    { id: "build",    text: "Run 🏗️ Build Sheets Only, then open ⚙️ Settings and click Save once." },
    { id: "cleanup",  text: "When you're done testing, move the copy to the trash: a tracker in the trash skips its nightly job by itself." }
  ] }
];

/** The checklist as plain text, for the pop-up. */
function newCopyChecklistText_() {
  const lines = ["This sheet is a copy of your own sheet, so it still has your owner-only files and your data.", ""];
  NEW_COPY_STEPS_.forEach(group => {
    lines.push(group.title + ":");
    group.steps.forEach((st, i) => lines.push((i + 1) + ". " + st.text));
    lines.push("");
  });
  lines.push("The same list with tick boxes: 🐞 Debug → 📋 New Copy Checklist. Everything else: DEPLOY.md in the repo.");
  return lines.join("\n");
}

/**
 * Called at the end of the 🐞 Debug menu (z_MINE.gs ownerAddMenus_, every open, also
 * before Initial Setup on a fresh copy). Shows the checklist once on a new copy.
 */
function ownerNewCopyCheck_(ui) {
  try {
    const props = PropertiesService.getDocumentProperties();
    const id = SpreadsheetApp.getActiveSpreadsheet().getId();
    const home = OWNER_HOME_SHEET_ID_ || props.getProperty(NEW_COPY_HOME_KEY_);
    if (home === id) return;                                  // your main sheet
    if (props.getProperty(NEW_COPY_SHOWN_KEY_) === id) return; // shown on this copy already
    if (!home) {
      const answer = ui.alert("📋 New copy?",
        "Is this your MAIN sheet (the one you use every day)?\n\n" +
        "Yes = remember it as your main sheet (you won't be asked again here).\n" +
        "No = it's a new copy: show what to change in it.", ui.ButtonSet.YES_NO);
      if (answer === ui.Button.YES) { props.setProperty(NEW_COPY_HOME_KEY_, id); return; }
      if (answer !== ui.Button.NO) return; // closed with the X: ask again next time
    }
    props.setProperty(NEW_COPY_SHOWN_KEY_, id);
    ui.alert("📋 New copy: what to change", newCopyChecklistText_(), ui.ButtonSet.OK);
  } catch (e) {
    console.warn("New copy checklist not shown: " + e.message);
  }
}

// ----------------------------------------------------------------------
// 🐞 Debug → 📋 New Copy Checklist (the same list, with tick boxes kept per sheet)
// ----------------------------------------------------------------------
function ownerMenuNewCopyChecklist() {
  const ticks = newCopyTicks_();
  const esc = v => String(v).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const groups = NEW_COPY_STEPS_.map(group => `
      <div class="grp">
        <h2>${esc(group.title)}</h2>
        ${group.steps.map(st => {
          const key = group.id + "." + st.id;
          return `<label class="cbx"><input type="checkbox" data-k="${esc(key)}"${ticks[key] ? " checked" : ""} onchange="tick(this)"><span>${esc(st.text)}</span></label>`;
        }).join("")}
      </div>`).join("");
  const html = include_('Styles') + `
    <style>
      body { padding: 6px 14px 14px; }
      .grp { border: 1px solid var(--line); border-radius: var(--radius); padding: 10px 12px; margin: 10px 0; display: flex; flex-direction: column; gap: 9px; }
      .grp h2 { font-size: 15px; margin-bottom: 2px; }
      .cbx span { font-size: 13px; line-height: 1.45; }
      .foot { display: flex; justify-content: space-between; align-items: center; gap: 8px; margin-top: 6px; }
    </style>
    <p class="help">What to change in a new copy of your sheet. Your ticks are kept on this sheet. Everything else (files, after-deploy steps, the quick test): DEPLOY.md in the repo.</p>
    ${groups}
    <div class="foot">
      <button type="button" class="btn txt" onclick="clearAll()">Clear ticks</button>
      <button type="button" class="btn pri" onclick="google.script.host.close()">Close</button>
    </div>
    <script>
      function tick(el) { google.script.run.ownerNewCopySetTick(el.getAttribute('data-k'), el.checked); }
      function clearAll() {
        document.querySelectorAll('input[type=checkbox]').forEach(function (el) { el.checked = false; });
        google.script.run.ownerNewCopyClearTicks();
      }
    </script>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(560).setHeight(620), "New copy checklist");
}

function newCopyTicks_() {
  try { return JSON.parse(PropertiesService.getDocumentProperties().getProperty(NEW_COPY_TICKS_KEY_) || "{}"); }
  catch (e) { return {}; }
}

/** Checklist dialog: one tick box changed. */
function ownerNewCopySetTick(key, on) {
  const ticks = newCopyTicks_();
  if (on) ticks[String(key)] = true; else delete ticks[String(key)];
  PropertiesService.getDocumentProperties().setProperty(NEW_COPY_TICKS_KEY_, JSON.stringify(ticks));
  return true;
}

function ownerNewCopyClearTicks() {
  PropertiesService.getDocumentProperties().deleteProperty(NEW_COPY_TICKS_KEY_);
  return true;
}
