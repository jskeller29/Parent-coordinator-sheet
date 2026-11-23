// A whole tracker for end-to-end tests: a new school's copy (the usual tabs) after Initial Setup, with the
// REAL project code on the fake sheet. Same as the testers' harness (release check round 3), without a browser.
//   const t = newTracker();                    // opts: { setup: false } = before Initial Setup
//   t.idmUpload("1 - Student Data (ATS).csv")   // Input Data Mode's upload path (idm_submitCsv), files in order
//   t.build()                                  // 🏗️ Build Sheets (buildAllDerivedSheets)
//   t.tab("Student List")                      // a tab's values (2D array)
//   t.csvText(name)                            // a test-data file's text
// Formulas are not calculated (a cell holds the formula text). Nothing here talks to Google.
const fs = require('fs'), path = require('path');
const { loadProject, propStore, ROOT } = require('./load');
const { makeSpreadsheet, fakeSpreadsheetApp } = require('./fakesheet');

const TABS = ["Contact Log", "Combined Contact Tracking", "PCAR", "Directory", "Raw Data", "Student List", "Notes",
  "Backend", "Events", "Phone Contacts", "Parents List", "Send Out"];

function csvText(name) { return fs.readFileSync(path.join(ROOT, 'test-data', name), 'utf8'); }

// Like Google's Utilities.parseCsv: quoted fields may hold commas, quotes and line breaks
function parseCsv(text) {
  const rows = []; let row = [], cell = '', q = false;
  text = String(text).replace(/^﻿/, '');
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) { if (ch === '"' && text[i + 1] === '"') { cell += '"'; i++; } else if (ch === '"') q = false; else cell += ch; }
    else if (ch === '"') q = true;
    else if (ch === ',') { row.push(cell); cell = ''; }
    else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(cell); rows.push(row); row = []; cell = ''; }
    else cell += ch;
  }
  if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
  return rows;
}

function newTracker(opts) {
  opts = opts || {};
  const ss = makeSpreadsheet(TABS);
  ss.getSheetByName("Backend").hideSheet();
  const app = fakeSpreadsheetApp(ss);
  const ui = app._ui;
  ui.Button = { YES: 'YES', NO: 'NO', OK: 'OK', CANCEL: 'CANCEL', CLOSE: 'CLOSE' };
  ui.ButtonSet = { YES_NO: 'YES_NO', OK: 'OK', OK_CANCEL: 'OK_CANCEL', YES_NO_CANCEL: 'YES_NO_CANCEL' };
  ui.alert = (...a) => { ui._alerts.push(a.map(String).join(" | ")); return 'OK'; };
  const opened = [];
  const rec = kind => (h, title) => opened.push({ kind, name: (h && h._name) || '(html string)', title: title || (h && h._title) || '' });
  ui.showSidebar = rec('sidebar'); ui.showModalDialog = rec('modal'); ui.showModelessDialog = rec('modeless');
  const g = loadProject({ extra: { SpreadsheetApp: app }, noTrialFiles: true });
  g.PropertiesService.getUserProperties = (() => { const u = propStore(); return () => u; })();
  const pad = n => String(n).padStart(2, '0');
  g.Utilities = Object.assign({}, g.Utilities, {
    parseCsv: parseCsv,
    formatDate: (d, tz, pattern) => {
      d = new Date(d); if (isNaN(d)) return ''; pattern = pattern || 'yyyy-MM-dd';
      return pattern.replace(/yyyy|yy|MM|dd|HH|mm|ss|M|d/g, k => ({ yyyy: d.getFullYear(), yy: pad(d.getFullYear() % 100), MM: pad(d.getMonth() + 1),
        dd: pad(d.getDate()), HH: pad(d.getHours()), mm: pad(d.getMinutes()), ss: pad(d.getSeconds()), M: d.getMonth() + 1, d: d.getDate() })[k]);
    }
  });
  g.isAutoMigrateOn_ = () => false;
  g.zg_safeGuideLink_ = () => ""; g.zg_safeExampleLink_ = () => ""; g.getGuideLink_ = () => "https://example.org/guide";
  g.include_ = () => "";
  const out = o => { o.setTitle = t => { o._title = t; return o; }; o.setWidth = () => o; o.setHeight = () => o; o.setSandboxMode = () => o; o.addMetaTag = () => o; return o; };
  g.htmlFromFile_ = name => out({ _name: name });
  g.HtmlService = { createHtmlOutput: h => out({ _html: String(h || '') }), createTemplateFromFile: name => ({ evaluate: () => out({ _name: name }) }),
                    createHtmlOutputFromFile: name => out({ _name: name }), SandboxMode: {} };
  g.DriveApp = { getRootFolder: () => ({ getOwner: () => ({ getEmail: () => "tester@gmail.com" }) }),
                 getFileById: () => ({ isTrashed: () => false, getName: () => "Fake tracker" }), searchFiles: () => ({ hasNext: () => false }) };
  const t = {
    ss, app, ui, g, opened,
    tab: name => { const s = ss.getSheetByName(name); return s ? s.getDataRange().getValues() : null; },
    csvText,
    idmUpload: (...files) => { let v = g.idm_open(); if (v && v.overview) v = g.idm_goToStep(0); files.forEach(f => { v = g.idm_submitCsv(csvText(f)); }); return v; },
    build: () => g.buildAllDerivedSheets()
  };
  if (opts.setup !== false) { g.onOpen({ source: ss }); g.runInitialSetup_(); }
  return t;
}

module.exports = { newTracker, parseCsv, csvText, TABS };
