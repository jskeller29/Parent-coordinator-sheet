// Loads the project's .gs files into ONE Node context (Apps Script shares one global
// scope too), with small fake Apps Script services. For testing plain logic offline:
// nothing here talks to Google.
//   loadProject({ owner, extra, props })
//     owner: true = also load z_MINE.gs + z_NewCopy.gs (the owner's sheet); default = a school's sheet (no z_ files)
//     debugFiles: true = also load the owner's z_ContactLogDebug.gs / z_InputDataDebug.gs
//     noTrialFiles: true = leave out the optional Z# files (Z#ContactLog.gs), like a copy without them
//     extra: globals to add / replace (e.g. a fake ScriptApp or People)
//     props: starting document properties
// The returned context has every global function and var (ctx.od_index_(...)),
// ctx.__get("CONST_NAME") for top-level const / let, and ctx.__props (the fake document properties).
const fs = require('fs'), path = require('path'), vm = require('vm');
const ROOT = path.resolve(__dirname, '..', '..');

function memStore() {
  const m = new Map();
  return {
    get: k => (m.has(k) ? m.get(k) : null), put: (k, v) => m.set(k, String(v)), remove: k => m.delete(k),
    getAll: ks => Object.fromEntries(ks.filter(k => m.has(k)).map(k => [k, m.get(k)])),
    removeAll: ks => ks.forEach(k => m.delete(k))
  };
}
function propStore(start) {
  const m = Object.assign({}, start || {});
  return {
    getProperty: k => (k in m ? m[k] : null), setProperty: (k, v) => { m[k] = String(v); }, deleteProperty: k => { delete m[k]; },
    setProperties: (obj, deleteAllOthers) => { if (deleteAllOthers) Object.keys(m).forEach(k => delete m[k]); Object.keys(obj).forEach(k => { m[k] = String(obj[k]); }); },
    getProperties: () => Object.assign({}, m), _m: m
  };
}

function loadProject(opts) {
  opts = opts || {};
  const docProps = propStore(opts.props), scriptProps = propStore();
  const ctx = {
    console: opts.quiet === false ? console : { log() {}, warn() {}, error() {}, info() {} },
    Date, Math, JSON, Map, Set, Array, Object, String, Number, RegExp, Error, isNaN, parseInt,
    PropertiesService: { getDocumentProperties: () => docProps, getScriptProperties: () => scriptProps },
    CacheService: { getScriptCache: (() => { const c = memStore(); return () => c; })(), getDocumentCache: (() => { const c = memStore(); return () => c; })(), getUserCache: () => memStore() },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }), getDocumentLock: () => ({ tryLock: () => true, waitLock() {}, releaseLock() {} }) },
    Utilities: { formatDate: d => new Date(d).toISOString().slice(0, 10), sleep() {}, parseCsv: t => t.split(/\r?\n/).map(l => l.split(",")) },
    Session: { getScriptTimeZone: () => "America/New_York", getActiveUser: () => ({ getEmail: () => "" }) },
    SpreadsheetApp: {}, HtmlService: {}, ScriptApp: { getProjectTriggers: () => [] }, MailApp: { sendEmail() {} },
    __props: docProps._m, __scriptProps: scriptProps
  };
  Object.assign(ctx, opts.extra || {});
  vm.createContext(ctx);
  fs.readdirSync(ROOT)
    .filter(f => f.endsWith('.gs'))
    .filter(f => opts.owner || !/^z_(MINE|NewCopy)\.gs$/.test(f))
    .filter(f => opts.debugFiles || !/^z_(ContactLogDebug|InputDataDebug)\.gs$/.test(f))
    .filter(f => !opts.noTrialFiles || !/^Z#/.test(f)) // noTrialFiles: a copy without the optional Z# files
    .forEach(f => vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f }));
  // Top-level const / let (not properties of ctx, unlike var and function): ctx.__get("NAME")
  ctx.__get = name => vm.runInContext(name, ctx);
  return ctx;
}

module.exports = { loadProject, memStore, propStore, ROOT };
