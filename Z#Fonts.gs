// ======================================================================
// FILE: Z#Fonts.gs  (+ Z#FontsSidebar.html)
// PURPOSE: Optional "🔤 Sheet Fonts" tool (Z# = trial file; nothing relies on it).
// A sidebar to cycle through fonts, see each one on the tab you have open,
// then apply one to the built tabs, the Contact Log and any other tabs.
//
// Hook in onOpen (Menus.gs), a no-op without this file:
//     if (typeof addFontsMenuItem_ === "function") addFontsMenuItem_(menu);
//
// Preview is undoable: before a tab is previewed the first time, its fonts are
// saved (run-length encoded) in the user cache (ZF_SNAP_<sheetId>, list in
// ZF_SNAPS, 6 hours). Undo, Apply (for the tabs not applied) and the next open
// put them back. Apply saves the font for the built tabs / Contact Log in
// document properties SHEET_FONT_BUILT / SHEET_FONT_CONTACTLOG (getSheetFont_ in
// Core.gs), so builds and Contact Log edits keep using it.
// ======================================================================

const ZF_SKIP_TABS = ["Backend"];
const ZF_SNAP_LIST = 'ZF_SNAPS';
const ZF_SNAP_PREFIX = 'ZF_SNAP_';
const ZF_SNAP_MAX = 95000; // characters; the cache holds up to 100 KB per key

function addFontsMenuItem_(menu) {
  menu.addItem("🔤 Sheet Fonts", "openFontsTool");
  // A preview left on when the sidebar was closed is put back on the next open
  try { zf_restoreSnapshots_(null); } catch (e) { console.warn("Sheet Fonts preview not undone: " + e.message); }
}

/** Menu handler: opens the sidebar. */
function openFontsTool() {
  const html = HtmlService.createTemplateFromFile('Z#FontsSidebar').evaluate().setTitle('Sheet Fonts');
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Sidebar start-up: tabs (grouped), the saved fonts and any preview still showing. */
function zf_getState() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const built = builtFontSheets_().filter(n => ss.getSheetByName(n));
  const others = ss.getSheets()
    .map(sh => ({ name: sh.getName(), hidden: sh.isSheetHidden() }))
    .filter(t => ZF_SKIP_TABS.indexOf(t.name) === -1 && t.name !== "Contact Log" && built.indexOf(t.name) === -1);
  return {
    active: ss.getActiveSheet().getName(),
    built: built,
    hasContactLog: !!ss.getSheetByName("Contact Log"),
    others: others,
    fonts: { built: getSheetFont_('BUILT'), contactLog: getSheetFont_('CONTACTLOG') },
    defaults: SHEET_FONT_DEFAULTS_,
    previewing: zf_snapIds_().map(id => (zf_sheetById_(ss, id) || { getName: () => null }).getName()).filter(Boolean)
  };
}

/** Shows `font` on the tab that is open now (its old fonts are saved first so Undo can put them back). */
function zf_preview(font) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const name = sheet.getName();
  if (ZF_SKIP_TABS.indexOf(name) !== -1) return { ok: false, tab: name, msg: "This tab is skipped." };

  const rows = Math.max(sheet.getLastRow(), 1);
  const cols = Math.max(sheet.getLastColumn(), 1);
  const range = sheet.getRange(1, 1, rows, cols);
  const cache = CacheService.getUserCache();
  const id = String(sheet.getSheetId());

  if (!cache.get(ZF_SNAP_PREFIX + id)) {
    const json = JSON.stringify({ rows: rows, cols: cols, runs: zf_encode_(range.getFontFamilies()) });
    if (json.length > ZF_SNAP_MAX) {
      return { ok: false, tab: name, msg: "This tab mixes too many fonts to undo a preview safely, so it isn't previewed. You can still apply a font to it." };
    }
    cache.put(ZF_SNAP_PREFIX + id, json, 21600);
    const ids = zf_snapIds_();
    if (ids.indexOf(id) === -1) ids.push(id);
    cache.put(ZF_SNAP_LIST, JSON.stringify(ids), 21600);
  }

  range.setFontFamily(font);
  return { ok: true, tab: name };
}

/** Undo: puts back the fonts of every previewed tab. Returns the tab names. */
function zf_undoPreview() {
  return zf_restoreSnapshots_(null);
}

/**
 * Applies `font` to whole tabs and saves it for the built tabs / Contact Log.
 * opts = { built: bool, contactLog: bool, tabs: [names], theme: bool }
 */
function zf_apply(font, opts) {
  font = String(font || "").trim();
  if (!font) throw new Error("Pick a font first.");
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let names = [];
  if (opts.built) names = names.concat(builtFontSheets_());
  if (opts.contactLog) names.push("Contact Log");
  names = names.concat(opts.tabs || []).filter((n, i, a) => a.indexOf(n) === i && ZF_SKIP_TABS.indexOf(n) === -1);

  // The tabs being applied keep the new font; any other preview is put back
  const appliedIds = [];
  const applied = [];
  names.forEach(n => {
    const sh = ss.getSheetByName(n);
    if (!sh) return;
    appliedIds.push(String(sh.getSheetId()));
    applied.push(n);
  });
  zf_dropSnapshots_(appliedIds);
  const restored = zf_restoreSnapshots_(null);

  applied.forEach(n => {
    const sh = ss.getSheetByName(n);
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setFontFamily(font);
  });

  const props = PropertiesService.getDocumentProperties();
  const save = (group, on) => {
    if (!on) return;
    if (font === SHEET_FONT_DEFAULTS_[group]) props.deleteProperty('SHEET_FONT_' + group);
    else props.setProperty('SHEET_FONT_' + group, font);
  };
  save('BUILT', opts.built);
  save('CONTACTLOG', opts.contactLog && !!ss.getSheetByName("Contact Log"));

  let themeNote = "";
  if (opts.theme) {
    try {
      const theme = ss.getSpreadsheetTheme();
      if (theme) theme.setFontFamily(font);
      else themeNote = "This spreadsheet has no theme, so its default font wasn't changed.";
    } catch (e) {
      themeNote = "The spreadsheet's default font couldn't be changed (" + e.message + ").";
    }
  }
  return { applied: applied, restored: restored, themeNote: themeNote };
}

/** Back to the fonts the sheet always used: Montserrat on the built tabs, Roboto on the Contact Log. */
function zf_resetDefaults() {
  zf_restoreSnapshots_(null);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  builtFontSheets_().forEach(n => {
    const sh = ss.getSheetByName(n);
    if (sh) sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setFontFamily(SHEET_FONT_DEFAULTS_.BUILT);
  });
  const cl = ss.getSheetByName("Contact Log");
  if (cl) cl.getRange(1, 1, cl.getMaxRows(), cl.getMaxColumns()).setFontFamily(SHEET_FONT_DEFAULTS_.CONTACTLOG);
  const props = PropertiesService.getDocumentProperties();
  props.deleteProperty('SHEET_FONT_BUILT');
  props.deleteProperty('SHEET_FONT_CONTACTLOG');
  return true;
}

// ---------------- snapshots ----------------

function zf_snapIds_() {
  try { return JSON.parse(CacheService.getUserCache().get(ZF_SNAP_LIST) || "[]"); } catch (e) { return []; }
}

function zf_sheetById_(ss, id) {
  return ss.getSheets().find(sh => String(sh.getSheetId()) === String(id)) || null;
}

/** Forgets the saved fonts of these tabs (they keep what they have now). */
function zf_dropSnapshots_(ids) {
  const cache = CacheService.getUserCache();
  const left = zf_snapIds_().filter(id => ids.indexOf(id) === -1);
  ids.forEach(id => cache.remove(ZF_SNAP_PREFIX + id));
  cache.put(ZF_SNAP_LIST, JSON.stringify(left), 21600);
}

/** Puts back every saved tab's fonts (only = list of ids, or null for all) and forgets them. */
function zf_restoreSnapshots_(only) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cache = CacheService.getUserCache();
  const ids = zf_snapIds_();
  const restored = [];
  const left = [];
  ids.forEach(id => {
    if (only && only.indexOf(id) === -1) { left.push(id); return; }
    const json = cache.get(ZF_SNAP_PREFIX + id);
    cache.remove(ZF_SNAP_PREFIX + id);
    const sh = zf_sheetById_(ss, id);
    if (!json || !sh) return;
    const snap = JSON.parse(json);
    const rows = Math.min(snap.rows, sh.getMaxRows());
    const cols = Math.min(snap.cols, sh.getMaxColumns());
    const matrix = zf_decode_(snap.runs, snap.rows, snap.cols).slice(0, rows).map(r => r.slice(0, cols));
    sh.getRange(1, 1, rows, cols).setFontFamilies(matrix);
    restored.push(sh.getName());
  });
  cache.put(ZF_SNAP_LIST, JSON.stringify(left), 21600);
  return restored;
}

/** [["Roboto","Roboto"],["Arial","Roboto"]] -> [["Roboto",2],["Arial",1],["Roboto",1]] (row by row) */
function zf_encode_(matrix) {
  const runs = [];
  matrix.forEach(row => row.forEach(f => {
    const font = f || "";
    const last = runs[runs.length - 1];
    if (last && last[0] === font) last[1]++;
    else runs.push([font, 1]);
  }));
  return runs;
}

function zf_decode_(runs, rows, cols) {
  const flat = [];
  runs.forEach(r => { for (let i = 0; i < r[1]; i++) flat.push(r[0] || null); });
  const out = [];
  for (let r = 0; r < rows; r++) out.push(flat.slice(r * cols, (r + 1) * cols));
  return out;
}
