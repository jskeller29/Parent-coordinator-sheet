// ======================================================================
// FILE: ContactLog.gs
// PURPOSE: The "📇 Contact Log" menu — a Quick Log tool for entering a new
// interaction and a Follow-ups view for managing notes that need follow-up.
//
// SELF-CONTAINED: the only hook into the rest of the app is a single guarded
// line in onOpen (New Copy & On Open.gs):
//     if (typeof buildContactLogMenu_ === "function") buildContactLogMenu_();
// If this file (and its ContactLogSidebar.html) are not included, that line
// is a no-op and nothing else changes.
//
// Contact Log columns (1-based): A week · B date · C guardian · D person
// spoke with · E OSIS · F student · G # · H site · I class · J method1 ·
// K method2 · L type · M notes · N follow-up status · O follow-up note ·
// P hide-note checkbox.
// ======================================================================

// column numbers (kept local so the module doesn't depend on Constants.gs)
var CLQ = { A:1, DATE:2, GUARDIAN:3, PERSON:4, OSIS:5, STUDENT:6, COUNT:7,
            SITE:8, CLASS:9, METHOD:10, TYPE:12, NOTES:13, FSTATUS:14, FNOTE:15, HIDE:16, W:16 };

// ---------------------------------------------------------------------
// MENU + LAUNCHERS
// ---------------------------------------------------------------------
function buildContactLogMenu_() {
  SpreadsheetApp.getUi().createMenu("📇 Contact Log")
    .addItem("➕ New Interaction", "openNewInteraction")
    .addItem("🔔 Manage Notes / Follow-ups", "openManageNotes")
    .addToUi();
}

function openNewInteraction() { CacheService.getScriptCache().put('CL_INIT_TAB', 'new', 120); cl_show_(); }
function openManageNotes()    { CacheService.getScriptCache().put('CL_INIT_TAB', 'followups', 120); cl_show_(); }

/** Opens the tool as a sidebar or a centered dialog, per the saved preference. */
function cl_show_() {
  const mode = PropertiesService.getDocumentProperties().getProperty('CL_VIEW_MODE') || 'sidebar';
  const html = HtmlService.createHtmlOutputFromFile('ContactLogSidebar').setTitle('Contact Log');
  if (mode === 'modal') {
    SpreadsheetApp.getUi().showModalDialog(html.setWidth(720).setHeight(660), '📇 Contact Log');
  } else {
    SpreadsheetApp.getUi().showSidebar(html.setWidth(360));
  }
}

function cl_getInitialTab() { return CacheService.getScriptCache().get('CL_INIT_TAB') || 'new'; }
function cl_getViewMode()   { return PropertiesService.getDocumentProperties().getProperty('CL_VIEW_MODE') || 'sidebar'; }

/** Flip sidebar <-> full-screen and reopen on the given tab. Client closes itself after. */
function cl_switchView(tab) {
  const props = PropertiesService.getDocumentProperties();
  const next = (props.getProperty('CL_VIEW_MODE') === 'modal') ? 'sidebar' : 'modal';
  props.setProperty('CL_VIEW_MODE', next);
  CacheService.getScriptCache().put('CL_INIT_TAB', tab || 'new', 120);
  cl_show_();
  return true;
}

// ---------------------------------------------------------------------
// DATA FOR THE UI
// ---------------------------------------------------------------------

/** Roster for the search box — reuses the SAME source the override dialog uses. */
function cl_getRoster() {
  if (typeof getAllStudentsForEdit === "function") return getAllStudentsForEdit();
  return [];
}

/**
 * Everything the family panel shows for one OSIS: ParentSquare / NYCSA status
 * (from "Parents Divided", matched by header) and the contact history for the
 * family (from the Contact Log, structured).
 */
function cl_getFamilyContext(osis) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const key = String(osis || "").split(",")[0].trim();
  const out = { ps: "", nycsa: "", history: [] };
  if (!key) return out;

  // --- ParentSquare / NYCSA from Parents Divided (header-matched) ---
  const pd = ss.getSheetByName("Parents Divided");
  if (pd && pd.getLastRow() > 1) {
    const data = pd.getDataRange().getValues();
    const head = data[0].map(h => String(h).toLowerCase());
    const idOsis = head.findIndex(h => h.includes("student_id") || h === "osis");
    const idPs   = head.findIndex(h => h.includes("square"));
    const idNy   = head.findIndex(h => h.includes("nycsa"));
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idOsis]).split(",")[0].trim() === key) {
        if (idPs >= 0 && !out.ps) out.ps = String(data[i][idPs]).trim();
        if (idNy >= 0 && !out.nycsa) out.nycsa = String(data[i][idNy]).trim();
        break;
      }
    }
  }

  // --- History from the Contact Log ---
  const cl = ss.getSheetByName("Contact Log");
  if (cl) {
    const rows = cl.getDataRange().getDisplayValues();
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (String(r[CLQ.OSIS - 1]).trim().toUpperCase() === "END") break;
      if (!cl_osisMatch_(r[CLQ.OSIS - 1], key)) continue;
      if (!String(r[CLQ.NOTES - 1]).trim() && !String(r[CLQ.FNOTE - 1]).trim()) continue;
      out.history.push({
        row: i + 1,
        date: String(r[CLQ.DATE - 1]).trim(),
        method: String(r[CLQ.METHOD - 1] || "").trim(),
        type: String(r[CLQ.TYPE - 1] || "").trim(),
        who: String(r[CLQ.PERSON - 1] || "").trim(),
        note: String(r[CLQ.NOTES - 1] || "").trim(),
        fnote: String(r[CLQ.FNOTE - 1] || "").trim(),
        followup: /^yes/i.test(String(r[CLQ.FSTATUS - 1] || "").trim()),
        hidden: cl_truthy_(r[CLQ.HIDE - 1])
      });
    }
    out.history.reverse(); // newest first
  }
  return out;
}

/** All interactions that still need follow-up (status starts with "Yes"). */
function cl_getFollowups() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cl = ss.getSheetByName("Contact Log");
  const list = [];
  if (!cl) return list;
  const rows = cl.getDataRange().getDisplayValues();
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (String(r[CLQ.OSIS - 1]).trim().toUpperCase() === "END") break;
    if (!/^yes/i.test(String(r[CLQ.FSTATUS - 1] || "").trim())) continue;
    list.push({
      row: i + 1,
      date: String(r[CLQ.DATE - 1]).trim(),
      guardian: String(r[CLQ.GUARDIAN - 1] || "").trim(),
      student: String(r[CLQ.STUDENT - 1] || "").trim(),
      osis: String(r[CLQ.OSIS - 1] || "").trim(),
      method: String(r[CLQ.METHOD - 1] || "").trim(),
      type: String(r[CLQ.TYPE - 1] || "").trim(),
      note: String(r[CLQ.NOTES - 1] || "").trim(),
      fnote: String(r[CLQ.FNOTE - 1] || "").trim(),
      hidden: cl_truthy_(r[CLQ.HIDE - 1])
    });
  }
  return list.reverse(); // newest first
}

// ---------------------------------------------------------------------
// WRITES
// ---------------------------------------------------------------------

/** Save a brand-new interaction into a fresh row just above the "END" bar. */
function cl_saveInteraction(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Contact Log");
  if (!sheet) throw new Error("Contact Log sheet not found.");
  if (!e || !e.osis) throw new Error("Pick a family first.");
  if (!e.notes || !String(e.notes).trim()) throw new Error("Add a note before saving.");

  const endRow = (typeof findEndRowInColumnE_ === "function") ? findEndRowInColumnE_(sheet) : null;
  if (!endRow || endRow < 2) throw new Error("Could not find the END bar in the Contact Log.");

  sheet.insertRowBefore(endRow);      // new blank row lands at endRow; END shifts down
  const r = endRow;
  sheet.getRange(r, 1, 1, CLQ.W).clearContent();

  const sc = cl_siteClassForOsis_(ss, e.osis);
  const dateVal = e.date ? new Date(e.date) : (function(){ const d = new Date(); d.setHours(0,0,0,0); return d; })();

  sheet.getRange(r, CLQ.DATE).setValue(isNaN(dateVal) ? new Date() : dateVal);
  sheet.getRange(r, CLQ.GUARDIAN).setValue(e.guardian || "");
  sheet.getRange(r, CLQ.PERSON).setValue(e.person || "");
  sheet.getRange(r, CLQ.OSIS).setValue(e.osis);
  sheet.getRange(r, CLQ.STUDENT).setValue(e.student || "");
  sheet.getRange(r, CLQ.COUNT).setValue(Number(e.count) > 0 ? Number(e.count) : 1);
  sheet.getRange(r, CLQ.SITE).setValue(sc.site);
  sheet.getRange(r, CLQ.CLASS).setValue(sc.cls);
  if (e.method) sheet.getRange(r, CLQ.METHOD).setValue(e.method);
  if (e.type)   sheet.getRange(r, CLQ.TYPE).setValue(e.type);
  sheet.getRange(r, CLQ.NOTES).setValue(String(e.notes).trim());
  sheet.getRange(r, CLQ.FSTATUS).setValue(e.followup ? "Yes" : "No");
  sheet.getRange(r, CLQ.FNOTE).setValue(e.fnote || "");

  if (typeof patchMissingFormulasInRow_ === "function") patchMissingFormulasInRow_(sheet, r);
  if (typeof formatContactLogRows_ === "function") formatContactLogRows_(sheet, r, 1);
  else sheet.getRange(r, CLQ.HIDE).insertCheckboxes();
  sheet.getRange(r, CLQ.HIDE).setValue(e.hide === true).setFontColor("#000000");

  SpreadsheetApp.flush();
  cl_syncOsis_(e.osis);
  if (typeof applyTypeDropdownColors === "function") { try { applyTypeDropdownColors(); } catch (err) { console.error(err); } }
  return { ok: true };
}

/** Update a follow-up interaction from the Manage Notes view. */
function cl_updateFollowup(row, payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Contact Log");
  if (!sheet || !row) return { ok: false };
  payload = payload || {};

  if (payload.fnote !== undefined) sheet.getRange(row, CLQ.FNOTE).setValue(String(payload.fnote));
  if (payload.resolve === true)   sheet.getRange(row, CLQ.FSTATUS).setValue("No");
  if (payload.hide !== undefined) sheet.getRange(row, CLQ.HIDE).setValue(payload.hide === true);

  SpreadsheetApp.flush();
  cl_syncOsis_(String(sheet.getRange(row, CLQ.OSIS).getDisplayValue()).trim());
  return { ok: true };
}

// ---------------------------------------------------------------------
// HELPERS
// ---------------------------------------------------------------------
function cl_truthy_(v) { return v === true || String(v).trim().toLowerCase() === "true"; }

function cl_osisMatch_(cellValue, target) {
  if (typeof osisCellIncludes_ === "function") return osisCellIncludes_(cellValue, target);
  return String(cellValue).split(/[,\s]+/).map(s => s.trim()).indexOf(String(target).trim()) !== -1;
}

/** Runs the single-OSIS delta rebuild for each OSIS in a (possibly grouped) value. */
function cl_syncOsis_(osisValue) {
  if (typeof updateSingleOsisDelta_ !== "function") return;
  const val = String(osisValue || "").trim();
  if (!val) return;
  if (val.indexOf(",") !== -1) {
    updateSingleOsisDelta_(val);
    val.split(",").map(o => o.trim()).filter(Boolean).forEach(o => updateSingleOsisDelta_(o));
  } else {
    updateSingleOsisDelta_(val);
  }
}

/** Pulls Site/Class from the Directory for an OSIS (or comma-group), like the ghost typist. */
function cl_siteClassForOsis_(ss, osisValue) {
  const sites = new Set(), classes = new Set();
  const dir = ss.getSheetByName("Directory");
  if (dir && osisValue) {
    const search = dir.getRange("C:C");
    String(osisValue).split(",").map(o => o.trim()).filter(Boolean).forEach(o => {
      const found = search.createTextFinder(o).matchEntireCell(true).findNext();
      if (found) {
        const vals = dir.getRange(found.getRow(), found.getColumn() + 1, 1, 2).getValues()[0];
        if (String(vals[0]).trim()) sites.add(String(vals[0]).trim());
        if (String(vals[1]).trim()) classes.add(String(vals[1]).trim());
      }
    });
  }
  return { site: Array.from(sites).join(", "), cls: Array.from(classes).join(", ") };
}
