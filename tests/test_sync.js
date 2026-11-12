// Phone sync (Contacts Sync.gs + the owner's version in z_MINE.gs) against a FAKE Google Contacts:
// nothing here talks to Google. Scenarios: normal sync, a failed save, Google's limit (pause +
// resume), "more than half" (blocked when no one is there / removed after Yes), a long build
// before the sync, and the owner's tall rows.
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const DIR = require('./lib/load').ROOT;
function world(opts) {
  const contacts = new Map(); let n = 0; const writes = []; const triggers = []; let failUpdate = opts.failUpdate || null; let quotaAt = opts.quotaAt || 0;
  const mk = (given, family, phone, email, notes, inLabel) => {
    const rn = "people/c" + (++n);
    contacts.set(rn, { resourceName: rn, etag: "e" + n, names: [{ givenName: given, familyName: family, displayName: given + " " + family }],
      phoneNumbers: phone ? [{ value: phone, type: "mobile" }] : [], emailAddresses: email ? [{ value: email, type: "home" }] : [],
      biographies: notes ? [{ value: notes }] : [], organizations: [], memberships: inLabel ? [{ contactGroupMembership: { contactGroupResourceName: "contactGroups/sd" } }] : [] });
    return rn;
  };
  let writeCount = 0;
  const clone = x => JSON.parse(JSON.stringify(x));
  const People = {
    ContactGroups: { list: () => ({ contactGroups: [{ name: "School Directory", resourceName: "contactGroups/sd" }] }), create: () => ({ resourceName: "contactGroups/new" }) },
    People: {
      Connections: { list: (me, o) => { const all = [...contacts.values()]; const start = Number(o.pageToken || 0); return { connections: all.slice(start, start + 2).map(clone), nextPageToken: start + 2 < all.length ? String(start + 2) : undefined }; } },
      get: rn => { if (!contacts.has(rn)) throw new Error("404 not found"); return clone(contacts.get(rn)); },
      createContact: d => { writeCount++; if (quotaAt && writeCount >= quotaAt) { quotaAt = 0; throw new Error("Quota exceeded: 429"); } const rn = "people/c" + (++n); contacts.set(rn, Object.assign(clone(d), { resourceName: rn, etag: "e" + n })); writes.push("create " + d.names[0].givenName); return { resourceName: rn, etag: "e" + n }; },
      updateContact: (d, rn) => { writeCount++; if (failUpdate && d.names[0].givenName === failUpdate) throw new Error("Internal error"); contacts.set(rn, Object.assign(clone(d), { resourceName: rn, etag: "e" + (++n) })); writes.push("update " + d.names[0].givenName); return { resourceName: rn, etag: "e" + n }; },
      batchDeleteContacts: o => { o.resourceNames.forEach(rn => contacts.delete(rn)); writes.push("delete " + o.resourceNames.length); },
      deleteContact: rn => contacts.delete(rn)
    }
  };
  const store = () => { const m = new Map(); return { get: k => m.has(k) ? m.get(k) : null, put: (k, v) => m.set(k, String(v)), remove: k => m.delete(k), getAll: ks => Object.fromEntries(ks.filter(k => m.has(k)).map(k => [k, m.get(k)])), removeAll: ks => ks.forEach(k => m.delete(k)) }; };
  const pstore = () => { const m = {}; return { getProperty: k => (k in m ? m[k] : null), setProperty: (k, v) => { m[k] = String(v); }, deleteProperty: k => { delete m[k]; }, getProperties: () => Object.assign({}, m), _m: m }; };
  const scriptCache = store(), docCache = store(), scriptProps = pstore(), docProps = pstore();
  let data = opts.rows;
  const toasts = [];
  const ss = { getSheetByName: name => name === "Phone Contacts" ? { getDataRange: () => ({ getValues: () => data }), getLastRow: () => data.length } : null, toast: (m, t) => toasts.push(t + ": " + m), getId: () => "SS1" };
  const ctx = { console: { log() {}, warn() {}, error() {} }, Date, Math, JSON, Map, Set, Array, Object, String, Number, RegExp, Error, isNaN, parseInt,
    People, PropertiesService: { getScriptProperties: () => scriptProps, getDocumentProperties: () => docProps },
    CacheService: { getScriptCache: () => scriptCache, getDocumentCache: () => docCache, getUserCache: () => store() },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }), getDocumentLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    ScriptApp: { getProjectTriggers: () => triggers.slice(), deleteTrigger: t => triggers.splice(triggers.indexOf(t), 1),
      newTrigger: h => ({ timeBased: () => ({ after: ms => ({ create: () => { const t = { h, ms, getHandlerFunction: () => h, getUniqueId: () => "t" + triggers.length }; triggers.push(t); return t; } }) }) }) },
    Utilities: { sleep() {}, formatDate: d => String(d) }, Session: { getScriptTimeZone: () => "America/New_York", getActiveUser: () => ({ getEmail: () => "" }) },
    MailApp: { sendEmail() {} },
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, getActive: () => ss, getUi: () => { if (opts.ui) return opts.ui; throw new Error("no ui"); } },
    HtmlService: {} };
  vm.createContext(ctx);
  fs.readdirSync(DIR).filter(f => f.endsWith('.gs') && (opts.owner || f !== 'z_MINE.gs') && f !== 'z_InputDataDebug.gs')
    .forEach(f => vm.runInContext(fs.readFileSync(path.join(DIR, f), 'utf8'), ctx, { filename: f }));
  if (opts.lateStart) ctx.SCRIPT_START_MS_ = Date.now() - opts.lateStart;
  return { ctx, contacts, mk, writes, triggers, scriptProps, scriptCache, toasts, setRows: r => { data = r; } };
}
const H = ["First Name", "Last Name", "E-mail 1 - Value", "Phone 1 - Type", "Phone 1 - Value", "Notes", "Organization Name", "Organization Department"];
const row = (p, s, phone, notes) => [p, s, "", "mobile", phone, notes || "", "", ""];
const state = w => JSON.parse(w.scriptCache.get('SYNC_UI_STATE'));
const names = w => [...w.contacts.values()].map(p => p.names[0].givenName).sort().join(",");

// S1: unchanged skipped, changed updated, new created, stale deleted
let w = world({ rows: [H, row("Ann", "Kid A", "2125550101", "n1"), row("Bob", "Kid B", "2125550199", "n2"), row("Eve", "Kid E", "2125550105", "n5")] });
w.mk("Ann", "Kid A", "2125550101", "", "n1", true); w.mk("Bob", "Kid B", "2125550102", "", "n2", true); w.mk("Cat", "Kid C", "2125550103", "", "n3", true); w.mk("Dan", "Kid D", "", "", "", false);
w.ctx.syncToGoogleContacts();
let st = state(w);
assert.deepStrictEqual([st.phase, st.created, st.edited, st.deleted, st.failed], ['done', 1, 1, 1, 0], JSON.stringify(st));
assert(w.writes.join("|") === "update Bob|create Eve|delete 1", w.writes.join("|"));
assert(names(w) === "Ann,Bob,Dan,Eve");
console.log("S1 normal sync ok:", w.writes.join(" | "));

// S2: a failed save keeps everything (no deletions)
w = world({ failUpdate: "Bob", rows: [H, row("Ann", "Kid A", "2125550101", "n1"), row("Bob", "Kid B", "2125550199", "n2")] });
w.mk("Ann", "Kid A", "2125550101", "", "n1", true); w.mk("Bob", "Kid B", "2125550102", "", "n2", true); w.mk("Cat", "Kid C", "2125550103", "", "n3", true);
w.ctx.syncToGoogleContacts();
st = state(w);
assert(st.failed === 1 && st.deleted === 0 && /couldn't be saved/.test(st.note) && names(w) === "Ann,Bob,Cat", JSON.stringify(st));
console.log("S2 failed save ok:", st.note);

// S3: limit reached -> pause 15 min, nothing deleted; the resumed run finishes and cleans up
w = world({ quotaAt: 2, rows: [H, row("Ann", "Kid A", "2125550101", "n1"), row("Bob", "Kid B", "2125550199", "n2"), row("Eve", "Kid E", "2125550105", "n5")] });
w.mk("Ann", "Kid A", "2125550101", "", "n1", true); w.mk("Bob", "Kid B", "2125550102", "", "n2", true); w.mk("Cat", "Kid C", "2125550103", "", "n3", true);
w.ctx.syncToGoogleContacts();
st = state(w);
assert(st.phase === 'quota' && w.scriptProps.getProperty('SYNC_START_ROW') === "3" && w.triggers.length === 1 && w.triggers[0].ms === 15 * 60000 && names(w) === "Ann,Bob,Cat", JSON.stringify(st));
w.triggers.length = 0;
w.ctx.syncToGoogleContacts(); // the 15-minute trigger
st = state(w);
assert(st.phase === 'done' && st.created === 1 && st.edited === 1 && st.deleted === 1 && names(w) === "Ann,Bob,Eve" && !w.scriptProps.getProperty('SYNC_START_ROW'), JSON.stringify(st) + names(w));
console.log("S3 limit pause + resume ok");

// S4: over half, nobody there -> nothing removed, note left
w = world({ rows: [H, row("Ann", "Kid A", "2125550101", "n1")] });
w.mk("Ann", "Kid A", "2125550101", "", "n1", true);
for (let i = 0; i < 11; i++) w.mk("Old" + i, "Kid", "", "", "", true);
w.ctx.syncToGoogleContacts();
st = state(w);
assert(st.deleted === 0 && /over half/.test(st.note) && w.contacts.size === 12, JSON.stringify(st));
// ... and with someone there who says Yes
w = world({ rows: [H, row("Ann", "Kid A", "2125550101", "n1")], ui: { alert: () => "YES", ButtonSet: { YES_NO: 1 }, Button: { YES: "YES" } } });
w.mk("Ann", "Kid A", "2125550101", "", "n1", true);
for (let i = 0; i < 11; i++) w.mk("Old" + i, "Kid", "", "", "", true);
w.ctx.syncToGoogleContacts();
st = state(w);
assert(st.deleted === 11 && w.contacts.size === 1, JSON.stringify(st));
console.log("S4 more-than-half check ok (blocked unattended, removed after Yes)");

// S5: the build used most of the run -> starts in a fresh run
w = world({ lateStart: 4.3 * 60000, rows: [H, row("Ann", "Kid A", "2125550101", "n1")] });
w.ctx.syncToGoogleContacts();
assert(state(w).phase === 'paused' && w.triggers.length === 1 && w.triggers[0].ms === 60000 && w.writes.length === 0);
console.log("S5 late start -> fresh run ok");

// S6: owner sheet (z_MINE): tall rows, phone must agree, unchanged skipped
const OH = ["First Name", "Last Name", "Notes", "Organization Name", "Organization Department", "E-mail 1 - Value", "E-mail 1 - Label", "Phone 1 - Value", "Phone 1 - Type", "Phone 2 - Value", "Phone 2 - Type"];
w = world({ owner: true, rows: [OH, ["Ann", "Kid A", "n1", "", "", "", "", "2125550101", "mobile", "", ""], ["Ann", "Kid Z", "n9", "", "", "", "", "2125550909", "mobile", "", ""]] });
w.mk("Ann", "Kid A", "2125550101", "", "n1", true); w.mk("Gone", "Kid", "", "", "", true);
w.ctx.OWNER_HOOKS_CACHE_ = undefined;
w.ctx.syncToGoogleContacts();
st = state(w);
assert(st.phase === 'done' && st.created === 1 && st.deleted === 1 && w.writes.join("|") === "create Ann|delete 1", JSON.stringify(st) + " " + w.writes.join("|"));
console.log("S6 owner sync ok:", w.writes.join(" | "));
