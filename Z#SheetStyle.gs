// ======================================================================
// FILE: Z#SheetStyle.gs  (+ Z#SheetStyleSidebar.html)
// PURPOSE: Optional, TEMPORARY "🎨 Sheet Style Lab" (Z# = trial file; nothing
// relies on it). Helps the owner decide fonts, colors and cell styling for the
// spreadsheet's own tabs.
//
// It never restyles a real tab while you experiment: it copies up to
// ZS_MAX_ROWS rows of the tab you pick (values + formats + column widths) into
// one "Style Preview" tab and styles that copy. Every change re-copies and
// re-applies the whole style, so "Keep as is" always means the original look.
// The picks are kept per tab (user property ZS_PICKS) and shown as text to send
// back, so the final look can be built into the code.
//
// The only thing it changes on real tabs is the optional font Apply: whole tabs
// get the font, and the built tabs / Contact Log font is saved in document
// properties SHEET_FONT_BUILT / SHEET_FONT_CONTACTLOG (getSheetFont_ in Core.gs).
//
// Hook in onOpen (Menus.gs), a no-op without this file:
//     if (typeof addSheetStyleMenuItem_ === "function") addSheetStyleMenuItem_(menu);
// ======================================================================

const ZS_PREVIEW_TAB = "Style Preview";
const ZS_SKIP_TABS = ["Backend", ZS_PREVIEW_TAB];
const ZS_MAX_ROWS = 150;
const ZS_PICKS_KEY = 'ZS_PICKS';

function addSheetStyleMenuItem_(menu) {
  menu.addItem("🎨 Sheet Style Lab", "openSheetStyleLab");
}

/** Menu handler: opens the sidebar. */
function openSheetStyleLab() {
  const html = HtmlService.createTemplateFromFile('Z#SheetStyleSidebar').evaluate().setTitle('Sheet Style Lab');
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Sidebar start-up: tabs, the tab you're on, saved picks and the saved fonts. */
function zs_getState() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const active = ss.getActiveSheet().getName();
  const built = builtFontSheets_().filter(n => ss.getSheetByName(n));
  const tabs = ss.getSheets().map(sh => sh.getName()).filter(n => ZS_SKIP_TABS.indexOf(n) === -1);
  let picks = {};
  try { picks = JSON.parse(PropertiesService.getUserProperties().getProperty(ZS_PICKS_KEY) || "{}"); } catch (e) { picks = {}; }
  return {
    active: active,
    tabs: tabs,
    built: built,
    others: tabs.filter(n => n !== "Contact Log" && built.indexOf(n) === -1),
    hasContactLog: !!ss.getSheetByName("Contact Log"),
    headerRows: tabs.reduce((m, n) => { m[n] = Math.max(ss.getSheetByName(n).getFrozenRows(), 1); return m; }, {}),
    previewExists: !!ss.getSheetByName(ZS_PREVIEW_TAB),
    picks: picks,
    fonts: { built: getSheetFont_('BUILT'), contactLog: getSheetFont_('CONTACTLOG') }
  };
}

/**
 * Re-copies `source` into the Style Preview tab, applies `style` to the copy and
 * shows it. Also saves the picks for that tab.
 */
function zs_preview(source, style) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(source);
  if (!src || ZS_SKIP_TABS.indexOf(source) !== -1) throw new Error('Tab "' + source + '" not found.');

  const rows = Math.max(Math.min(src.getLastRow(), ZS_MAX_ROWS), 1);
  const cols = Math.max(src.getLastColumn(), 1);

  let prev = ss.getSheetByName(ZS_PREVIEW_TAB);
  if (!prev) prev = ss.insertSheet(ZS_PREVIEW_TAB, ss.getSheets().length);
  prev.getBandings().forEach(b => b.remove());
  prev.clear();
  prev.setConditionalFormatRules([]);
  if (prev.getMaxRows() < rows) prev.insertRowsAfter(prev.getMaxRows(), rows - prev.getMaxRows());
  if (prev.getMaxColumns() < cols) prev.insertColumnsAfter(prev.getMaxColumns(), cols - prev.getMaxColumns());
  prev.setRowHeightsForced(1, prev.getMaxRows(), 21);
  prev.setHiddenGridlines(src.hasHiddenGridlines());

  // Values + formats + column widths only, so formulas / dropdown scripts don't run on the copy
  const from = src.getRange(1, 1, rows, cols);
  const to = prev.getRange(1, 1, rows, cols);
  from.copyTo(to, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  from.copyTo(to, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  from.copyTo(to, SpreadsheetApp.CopyPasteType.PASTE_COLUMN_WIDTHS, false);
  for (let r = 1; r <= Math.min(rows, 12); r++) prev.setRowHeight(r, src.getRowHeight(r)); // header rows keep their height
  prev.setFrozenRows(Math.min(src.getFrozenRows(), rows));
  prev.setFrozenColumns(Math.min(src.getFrozenColumns(), cols));

  zs_applyStyle_(prev, rows, cols, style || {});
  prev.activate();

  zs_savePick_(source, style || {});
  return { rows: rows, cols: cols, cut: src.getLastRow() > ZS_MAX_ROWS };
}

/** The style picks, applied to the copy. "keep" (or empty) = leave as copied. */
function zs_applyStyle_(sheet, rows, cols, s) {
  const keep = v => v === undefined || v === null || v === "" || v === "keep";
  const headerRows = Math.min(Math.max(Number(s.headerRows) || 0, 0), rows);
  const all = sheet.getRange(1, 1, rows, cols);
  const head = headerRows ? sheet.getRange(1, 1, headerRows, cols) : null;
  const body = rows > headerRows ? sheet.getRange(headerRows + 1, 1, rows - headerRows, cols) : null;

  if (!keep(s.font)) all.setFontFamily(s.font);
  if (!keep(s.size)) all.setFontSize(Number(s.size));

  if (head) {
    if (!keep(s.headerBg)) head.setBackground(s.headerBg);
    if (!keep(s.headerText)) head.setFontColor(s.headerText);
    if (!keep(s.headerBold)) head.setFontWeight(s.headerBold === "bold" ? "bold" : "normal");
    if (!keep(s.headerSize)) head.setFontSize(Number(s.headerSize));
  }

  if (body) {
    if (s.bodyBg === "clear") body.setBackground(null);
    if (!keep(s.bodyText)) body.setFontColor(s.bodyText);
    if (!keep(s.banding)) {
      sheet.getBandings().forEach(b => b.remove());
      body.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false)
        .setFirstRowColor("#ffffff")
        .setSecondRowColor(s.banding);
    }
    if (s.lines === "rows") body.setBorder(null, null, true, null, null, true, s.lineColor || "#e1e5ea", SpreadsheetApp.BorderStyle.SOLID);
    if (s.lines === "grid") body.setBorder(true, true, true, true, true, true, s.lineColor || "#e1e5ea", SpreadsheetApp.BorderStyle.SOLID);
    if (s.lines === "none") body.setBorder(false, false, false, false, false, false);
    if (!keep(s.rowHeight)) sheet.setRowHeightsForced(headerRows + 1, rows - headerRows, Number(s.rowHeight));
  }

  if (!keep(s.valign)) all.setVerticalAlignment(s.valign);
  if (!keep(s.wrap)) {
    const strategy = { wrap: SpreadsheetApp.WrapStrategy.WRAP, clip: SpreadsheetApp.WrapStrategy.CLIP, overflow: SpreadsheetApp.WrapStrategy.OVERFLOW }[s.wrap];
    if (strategy) all.setWrapStrategy(strategy);
  }
  if (!keep(s.gridlines)) sheet.setHiddenGridlines(s.gridlines === "hide");
}

function zs_savePick_(tab, style) {
  const props = PropertiesService.getUserProperties();
  let picks = {};
  try { picks = JSON.parse(props.getProperty(ZS_PICKS_KEY) || "{}"); } catch (e) { picks = {}; }
  picks[tab] = style;
  try { props.setProperty(ZS_PICKS_KEY, JSON.stringify(picks)); } catch (e) { console.warn("Style picks not saved: " + e.message); }
}

/** Forgets the saved picks for one tab (or all when tab is empty). */
function zs_clearPicks(tab) {
  const props = PropertiesService.getUserProperties();
  if (!tab) { props.deleteProperty(ZS_PICKS_KEY); return {}; }
  let picks = {};
  try { picks = JSON.parse(props.getProperty(ZS_PICKS_KEY) || "{}"); } catch (e) { picks = {}; }
  delete picks[tab];
  props.setProperty(ZS_PICKS_KEY, JSON.stringify(picks));
  return picks;
}

/**
 * Leaving the lab, step 1: go back to a real tab. The Style Preview tab is
 * deleted by a SEPARATE call (zs_deletePreview): deleting the tab the user is on
 * makes every later call fail with "Sheet ... not found".
 */
function zs_leavePreview(backTo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const target = ss.getSheetByName(backTo) || ss.getSheets().find(sh => sh.getName() !== ZS_PREVIEW_TAB && !sh.isSheetHidden());
  if (target) target.activate();
  return true;
}

/** Leaving the lab, step 2. */
function zs_deletePreview() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prev = ss.getSheetByName(ZS_PREVIEW_TAB);
  if (!prev) return false;
  if (ss.getActiveSheet().getName() === ZS_PREVIEW_TAB) return false; // still on it: step 1 first
  ss.deleteSheet(prev);
  return true;
}

/**
 * Optional: puts `font` on whole real tabs and saves it for the built tabs /
 * Contact Log so builds and Contact Log edits keep it.
 * opts = { built: bool, contactLog: bool, tabs: [names], theme: bool }
 */
function zs_applyFont(font, opts) {
  font = String(font || "").trim();
  if (!font) throw new Error("Pick a font first.");
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let names = [];
  if (opts.built) names = names.concat(builtFontSheets_());
  if (opts.contactLog) names.push("Contact Log");
  names = names.concat(opts.tabs || []).filter((n, i, a) => a.indexOf(n) === i && ZS_SKIP_TABS.indexOf(n) === -1);

  const applied = [];
  names.forEach(n => {
    const sh = ss.getSheetByName(n);
    if (!sh) return;
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setFontFamily(font);
    applied.push(n);
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
  return { applied: applied, themeNote: themeNote };
}

/** Back to the fonts the sheet always used: Montserrat on the built tabs, Roboto on the Contact Log. */
function zs_resetFonts() {
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
