// ======================================================================
// FILE: Contacts Sync.gs
// PURPOSE: Pushes the Phone Contacts tab into Google Contacts ("School
// Directory" label): the resumable bulk sync, the single-contact upsert, the
// Full Sync sidebar endpoints, and the stop / restart controls.
// ======================================================================

// ==========================================
// 5. RESILIENT SIDEBAR BACKEND
// ==========================================
function startResilientSync() {
  return runAction_("Full Sync", () => startResilientSyncRun_());
}

function startResilientSyncRun_() {
  const cache = CacheService.getScriptCache();
  const currentState = cache.get('SYNC_UI_STATE');
  
  if (currentState) {
    const phase = JSON.parse(currentState).phase;
    if (phase !== 'idle' && phase !== 'done' && phase !== 'killed') {
      return; 
    }
  }

  reportBuildStep_(0); // sidebar's build checklist starts at step 1
  cache.put('SYNC_UI_STATE', JSON.stringify({ phase: 'building', progress: 0, total: 0 }), 21600);

  // Build + sync as one job (one build / sync at a time, Core.gs). Any error ends
  // the run as 'killed' with the error, so the sidebar says what went wrong and the
  // next Full Sync starts normally (before, a crash left 'building' for 6 hours).
  try {
    withBuildLock_("Full Sync", () => startResilientSyncLocked_(cache));
  } catch (e) {
    console.error("Full Sync failed: " + (e.stack || e.message));
    cache.put('SYNC_UI_STATE', JSON.stringify({ phase: 'killed', error: e.message || String(e) }), 21600);
  }
}

function startResilientSyncLocked_(cache) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  buildAllDerivedSheets();
  formatDirectoryAuto_(ss.getSheetByName("Directory"));
  const psSheet = ss.getSheetByName("Parents Divided");
  if (psSheet) psSheet.hideSheet();

  const phoneSheet = ss.getSheetByName("Phone Contacts");
  const total = phoneSheet ? Math.max(0, phoneSheet.getLastRow() - 1) : 0;

  cache.put('SYNC_UI_STATE', JSON.stringify({ phase: 'syncing', progress: 0, total: total }), 21600);
  
  // ==============================================================
  // ✨ FIRST BUILD UNLOCK DETECTOR ✨
  // Triggers ONLY when the builder successfully finishes the very first time!
  // ==============================================================
  const props = PropertiesService.getDocumentProperties();
  if (props.getProperty('SETUP_COMPLETE') === 'true' && props.getProperty('FULL_MENU_UNLOCKED') !== 'true') {
    props.setProperty('FULL_MENU_UNLOCKED', 'true');
    try { onOpen(); } catch(e) { console.error(e); } // Attempt silent UI refresh
    
    // Send a long-lasting toast notification to refresh the page
    ss.toast("All sheets are built! Please refresh your webpage (Press F5) to explore and see your new menu options.", "🎉 Build Complete!", 15);
  }

  syncToGoogleContacts();
}

function getSyncStatus() {
  const cache = CacheService.getScriptCache();
  const stateStr = cache.get('SYNC_UI_STATE');
  const state = stateStr ? JSON.parse(stateStr) : { phase: 'idle' };
  // The build's step list + current step, for the sidebar's "Building Sheets & Directory" checklist
  return Object.assign(state, getBuildProgress_());
}

function resetSyncStatus() {
  CacheService.getScriptCache().remove('SYNC_UI_STATE');
}

// ==========================================
// 1. BULK NIGHTLY SYNC (Enterprise Mirror Mode)
// Resumable: stops before the 6-minute limit and schedules itself to carry on.
// The owner's sheet (z_MINE.gs) reads up to 4 phones / 3 emails per row, matches
// a contact by name AND phone / email, and emails a summary at the end (see
// ownerHooks_().contactsSync); since 2026-11-12 it has the same label lookup and
// "more than half" check as every sheet.
// ==========================================
function syncToGoogleContacts() {
  return runAction_("Google Contacts sync", () => {
    try {
      return withBuildLock_("Phone sync", () => syncToGoogleContactsRun_());
    } catch (e) {
      if (!e.buildBusy) throw e;
      // Something else is building / syncing. A paused sync must not be lost: try again in 2 minutes.
      if (PropertiesService.getScriptProperties().getProperty('SYNC_START_ROW')) {
        ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === "syncToGoogleContacts") ScriptApp.deleteTrigger(t); });
        ScriptApp.newTrigger("syncToGoogleContacts").timeBased().after(2 * 60 * 1000).create();
      }
      console.warn("Phone sync waited: " + e.message);
    }
  });
}

// ----------------------------------------------------------------------
// HOW THE SYNC WORKS (rewritten 2026-11-12, owner's "fix 1-3"):
//   1. ONE look at all your Google Contacts per run (buildContactsIndex_), not one
//      per row: every row then finds its contact in that list (plus the 6-hour
//      row -> contact cache, for contacts made moments ago that Google doesn't
//      list yet). Before, each row paged through every contact.
//   2. A contact that already looks exactly like its row is skipped (no write,
//      no 0.75 s wait).
//   3. Never delete when unsure: a row whose save fails keeps its contact, and if
//      any row failed in this sync, nothing is removed this time. A Google
//      "limit reached" error pauses 15 minutes and carries on (the old code
//      swallowed it, so the remaining rows looked gone and were deleted).
//   4. "More than half" check (shared AND owner sheet): worked out at the start
//      of a sync from the whole sheet, so it also covers syncs that need more
//      than one run; asked with a pop-up when someone is there, else skipped and
//      a message is left. The answer is kept for the runs that follow.
//   5. Time: Google stops a run at 6 minutes counted from the START of the run
//      (SCRIPT_START_MS_, Core.gs), and Full Sync / the nightly build first. The
//      sync now counts from there and carries on in a fresh run when the build
//      left too little time (before, build + 4.5 min could pass 6 min, the run
//      was killed and the next one started at row 1 again).
// Script properties across runs: SYNC_START_ROW, SYNC_CREATED_COUNT,
// SYNC_EDITED_COUNT, SYNC_FAILED_COUNT, SYNC_DELETE_DECISION (ok / blocked),
// SYNC_LIST_RETRIES.
// ----------------------------------------------------------------------
var SYNC_MAX_RUN_MS_ = 4.5 * 60 * 1000;  // stop syncing rows this long after the run started
var SYNC_MIN_START_MS_ = 45 * 1000;      // less than this left: start in a fresh run instead
var CONTACT_FIELDS_ = 'names,phoneNumbers,emailAddresses,memberships,organizations,biographies';

function isQuotaError_(e) {
  return /quota|rate ?limit|429|resource_exhausted|too many requests/i.test(String((e && e.message) || e || ""));
}

/** "given|family" in lowercase: how a row and a contact are matched by name. */
function contactNameKey_(given, family) {
  return String(given || "").trim().toLowerCase() + "|" + String(family || "").trim().toLowerCase();
}

function personNameKey_(p) {
  const n = p && p.names && p.names[0];
  return n ? contactNameKey_(n.givenName, n.familyName) : "";
}

/**
 * Every Google contact, read once: byResource, byName ("given|family" -> people),
 * the "School Directory" label's members. add(person) keeps it current after a
 * create / update. Throws when Google can't be read (the run then changes nothing).
 */
function buildContactsIndex_(labelResource) {
  const index = {
    label: labelResource, byResource: new Map(), byName: new Map(), members: new Set(),
    inLabel: p => !!(p && p.memberships && p.memberships.some(m => m.contactGroupMembership && m.contactGroupMembership.contactGroupResourceName === labelResource)),
    add: p => {
      if (!p || !p.resourceName) return;
      const old = index.byResource.get(p.resourceName);
      if (old) {
        const list = index.byName.get(personNameKey_(old)) || [];
        const at = list.indexOf(old);
        if (at !== -1) list.splice(at, 1);
      }
      index.byResource.set(p.resourceName, p);
      const key = personNameKey_(p);
      if (key !== "|") {
        if (!index.byName.has(key)) index.byName.set(key, []);
        index.byName.get(key).push(p);
      }
      if (index.inLabel(p)) index.members.add(p.resourceName);
    }
  };
  let pageToken;
  do {
    const response = People.People.Connections.list('people/me', { pageSize: 1000, personFields: CONTACT_FIELDS_, pageToken: pageToken });
    (response.connections || []).forEach(p => index.add(p));
    pageToken = response.nextPageToken;
  } while (pageToken);
  index.labelCount = index.members.size;
  return index;
}

/**
 * A row's existing contact: the one this row used last time (6-hour cache), else
 * the people with the same names (filter = extra test, e.g. the owner's phone /
 * email rule), label members first. No Google call unless the cached contact is
 * too new to be in the list.
 */
function findIndexedContact_(index, cacheKey, nameKey, filter) {
  const cache = CacheService.getScriptCache();
  const cached = cacheKey ? cache.get(cacheKey) : null;
  if (cached) {
    let p = index.byResource.get(cached) || null;
    if (!p) {
      try { p = People.People.get(cached, { personFields: CONTACT_FIELDS_ }); index.add(p); }
      catch (e) { cache.remove(cacheKey); p = null; }
    }
    if (p && personNameKey_(p) === nameKey) return p;
  }
  const list = (index.byName.get(nameKey) || []).filter(p => !filter || filter(p));
  list.sort((a, b) => (index.inLabel(b) ? 1 : 0) - (index.inLabel(a) ? 1 : 0));
  return list[0] || null;
}

/** True when the contact already matches what would be written (then no write is needed). */
function contactMatchesPayload_(p, data, labelResource) {
  if (!p) return false;
  const n = (p.names && p.names[0]) || {}, dn = data.names[0];
  if (String(n.givenName || "") !== String(dn.givenName || "") || String(n.familyName || "") !== String(dn.familyName || "")) return false;
  const phones = arr => JSON.stringify((arr || []).map(x => [String(x.value || "").replace(/\D/g, "").replace(/^1(\d{10})$/, "$1") || String(x.value || "").toLowerCase().trim(), String(x.type || "").toLowerCase()]));
  const emails = arr => JSON.stringify((arr || []).map(x => [String(x.value || "").toLowerCase().trim(), String(x.type || "").toLowerCase()]));
  const bio = arr => (arr && arr.length ? String(arr[0].value || "").trim() : "");
  const org = arr => (arr && arr.length ? String(arr[0].name || "").trim() + "|" + String(arr[0].title || "").trim() : "");
  if (phones(p.phoneNumbers) !== phones(data.phoneNumbers)) return false;
  if (emails(p.emailAddresses) !== emails(data.emailAddresses)) return false;
  if (bio(p.biographies) !== bio(data.biographies)) return false;
  if (org(p.organizations) !== org(data.organizations)) return false;
  if (labelResource && !(p.memberships || []).some(m => m.contactGroupMembership && m.contactGroupMembership.contactGroupResourceName === labelResource)) return false;
  return true;
}

/** What to keep in the list after a create / update (Google's reply may leave fields out). */
function storedContact_(payload, reply) {
  return Object.assign({}, payload, reply || {});
}

function syncToGoogleContactsRun_() {
  const elapsed = () => Date.now() - SCRIPT_START_MS_;

  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) return;   
  
  const sheet = ss.getSheetByName("Phone Contacts");
  if (!sheet) return;

  const owner = getOwnerHooks_();
  const ownerSync = owner ? owner.contactsSync : null;
  
  const data = sheet.getDataRange().getValues();
  const props = PropertiesService.getScriptProperties();
  const cache = CacheService.getScriptCache();
  
  let startRow = parseInt(props.getProperty('SYNC_START_ROW')) || 1;
  let createdCount = parseInt(props.getProperty('SYNC_CREATED_COUNT')) || 0; 
  let editedCount = parseInt(props.getProperty('SYNC_EDITED_COUNT')) || 0; 
  let failedCount = parseInt(props.getProperty('SYNC_FAILED_COUNT')) || 0;
  
  let syncLog = JSON.parse(cache.get('SYNC_LOG') || '{"created":[], "edited":[], "deleted":[], "failed":[]}');
  if (!syncLog.failed) syncLog.failed = [];
  const total = data.length - 1;

  // Carry on in a fresh run in `minutes` (time is up, Google's limit, or busy)
  const scheduleResume = (row, minutes, phase) => {
    props.setProperty('SYNC_START_ROW', String(row));
    props.setProperty('SYNC_CREATED_COUNT', String(createdCount));
    props.setProperty('SYNC_EDITED_COUNT', String(editedCount));
    props.setProperty('SYNC_FAILED_COUNT', String(failedCount));
    cache.put('SYNC_LOG', JSON.stringify(syncLog), 21600);
    cache.put('SYNC_UI_STATE', JSON.stringify({ phase: phase, progress: row, total: total }), 21600);
    ScriptApp.getProjectTriggers().forEach(t => {
      if (t.getHandlerFunction() === "syncToGoogleContacts") ScriptApp.deleteTrigger(t);
    });
    ScriptApp.newTrigger("syncToGoogleContacts").timeBased().after(minutes * 60 * 1000).create();
  };

  // The build (Full Sync / nightly) used most of this run: start in a fresh one
  if (SYNC_MAX_RUN_MS_ - elapsed() < SYNC_MIN_START_MS_) {
    scheduleResume(startRow, 1, 'paused');
    ss.toast("The phone sync goes on in a minute (the build used most of this run's time).", "Syncing", 10);
    return;
  }

  // 1. Get or Create Label
  // pageSize matters: the default page is small, so without it a user with many
  // contact groups may not find "School Directory" and we'd create a duplicate.
  let resourceName = "";
  const groups = (ownerSync ? ownerSync.listContactGroups() : People.ContactGroups.list({ pageSize: 1000 })).contactGroups || [];
  const existingGroup = groups.find(g => g.name === "School Directory");
  
  if (existingGroup) {
    resourceName = existingGroup.resourceName;
  } else {
    const newGroup = People.ContactGroups.create({ contactGroup: { name: "School Directory" } });
    resourceName = newGroup.resourceName;
  }

  // 2. Every contact, once (see "HOW THE SYNC WORKS")
  ss.toast("Phase 1/3: Checking existing Google contacts...", "Syncing", 10);
  let index;
  try {
    index = buildContactsIndex_(resourceName);
    props.deleteProperty('SYNC_LIST_RETRIES');
  } catch (e) {
    const tries = (parseInt(props.getProperty('SYNC_LIST_RETRIES')) || 0) + 1;
    console.error("Could not read Google Contacts (try " + tries + "): " + (e && e.message));
    if (tries <= 3 || isQuotaError_(e)) {
      props.setProperty('SYNC_LIST_RETRIES', String(tries));
      scheduleResume(startRow, 15, isQuotaError_(e) ? 'quota' : 'paused');
      ss.toast("Couldn't read your Google Contacts right now. Trying again in 15 minutes; nothing was changed.", "Sync paused", 15);
      return;
    }
    clearSyncRunState_(props, cache);
    const msg = "⚠️ Phone sync stopped: Google Contacts could not be read (" + (e && e.message) + "). Nothing was changed.";
    PropertiesService.getDocumentProperties().setProperty('MISSED_SYNC_MSG', msg);
    throw new Error(msg);
  }

  // Column positions are read once from the header row
  const readRow = ownerSync ? ownerSync.rowReader(data[0]) : phoneContactsRowReader_(data[0]);
  // Contacts the rows of this sheet could be (all with the row's names), for the
  // "more than half" check and for rows done in an earlier run of this sync
  const protectByName = (fromRow, toRow, into) => {
    for (let r = fromRow; r < toRow; r++) {
      const row = readRow(data[r]);
      if (!row.parentName || String(row.parentName).trim() === "") continue;
      (index.byName.get(row.nameKey) || []).forEach(p => into.add(p.resourceName));
    }
    return into;
  };
  const tooManyGone = n => index.labelCount > 10 && n > index.labelCount * 0.5;

  if (startRow === 1) {
    // A new sync: fresh log and counts, and the "more than half" question now
    syncLog = { "created": [], "edited": [], "deleted": [], "failed": [] };
    createdCount = 0; editedCount = 0; failedCount = 0;
    cache.put('SYNC_LOG', JSON.stringify(syncLog), 21600);
    const sure = protectByName(1, data.length, new Set());
    const goneSoon = [...index.members].filter(id => !sure.has(id)).length;
    // confirmed / blocked (asked, or no one here to ask) / none-needed (not over half)
    const decision = tooManyGone(goneSoon) ? (askMassDeletion_(ss, goneSoon, index.labelCount) ? "confirmed" : "blocked") : "none-needed";
    props.setProperty('SYNC_DELETE_DECISION', decision);
  }
  
  // 3. Sync Rows
  ss.toast("Phase 2/3: Syncing spreadsheet rows... Safe to close sidebar.", "Syncing", 10);
  const usedThisRun = new Set();
  for (let i = startRow; i < data.length; i++) {
    
    if (elapsed() > SYNC_MAX_RUN_MS_) {
      scheduleResume(i, 1, 'paused');
      ss.toast("Timer reset. Resuming at row " + i, "Auto-Resume Active", 30);
      return; 
    }
    
    if (i % 5 === 0) {
      const currentStateStr = cache.get('SYNC_UI_STATE');
      if (currentStateStr && JSON.parse(currentStateStr).phase === 'killed') return; 
      cache.put('SYNC_UI_STATE', JSON.stringify({ phase: 'syncing', progress: i, total: total }), 21600);
    }

    const contactRow = readRow(data[i]);
    if (!contactRow.parentName || String(contactRow.parentName).trim() === "") continue;
    
    let result = null;
    try {
      result = contactRow.upsert(resourceName, index);
    } catch (e) {
      if (isQuotaError_(e)) {
        scheduleResume(i, 15, 'quota');
        ss.toast("API Limit reached. Pausing 15 mins.", "Quota Pause", 60);
        return;
      }
      result = { failed: true, name: String(contactRow.parentName), error: String(e && e.message || e) };
    }
    if (!result) continue;
    if (result.resourceName) usedThisRun.add(result.resourceName); // kept, even when this save failed
    if (result.failed) {
      failedCount++;
      if (syncLog.failed.length < 150) syncLog.failed.push(`<b>${result.name}</b><br><span style="color:#555; font-size:11px;">${result.error || "not saved"}</span>`);
      continue;
    }
    if (result.isNew) {
      createdCount++; 
      if (syncLog.created.length < 150) syncLog.created.push(result.name);
      else if (syncLog.created.length === 150) syncLog.created.push("...and more.");
    } else if (result.isEdited) {
      editedCount++; 
      if (syncLog.edited.length < 150) syncLog.edited.push(`<b>${result.name}</b><br><span style="color:#555; font-size:11px;">${result.logStr}</span>`);
      else if (syncLog.edited.length === 150) syncLog.edited.push("...and more.");
    }
    if (result.wrote) Utilities.sleep(750); // only after a write (People API write limit)
  }
  
  // 4. Cleanup Phase: label members no row of the sheet uses any more
  ss.toast("Phase 3/3: Removing old contacts...", "Final Cleanup", 10);
  const keep = protectByName(1, startRow, new Set(usedThisRun)); // rows done in earlier runs: by name
  const staleEntries = [...index.members].filter(id => !keep.has(id)).map(id => {
    const p = index.byResource.get(id);
    return [id, (p && p.names && p.names.length) ? p.names[0].displayName || p.names[0].givenName : "Unknown Contact"];
  });
  const decision = props.getProperty('SYNC_DELETE_DECISION') || "";
  let note = "";
  if (staleEntries.length && failedCount > 0) {
    note = failedCount + " contact(s) couldn't be saved, so no old contacts were removed this time (they will be next time).";
  } else if (staleEntries.length && tooManyGone(staleEntries.length) && decision !== "confirmed") {
    // Over half: removed only when confirmed. "No" at the start (or no one to ask) stays no; a check
    // that wasn't needed at the start (the numbers changed since) is asked now, if someone is here.
    const ok = decision !== "blocked" && askMassDeletion_(ss, staleEntries.length, index.labelCount);
    if (!ok) note = "Skipped removing " + staleEntries.length + " contacts (over half of the list). Run Full Sync and answer Yes to remove them.";
  }
  const deleteNow = staleEntries.length > 0 && !note;
  const deletedCount = deleteNow ? staleEntries.length : 0;

  clearSyncRunState_(props, cache);
  
  if (deleteNow) {
    let deletedLogCount = 0;
    staleEntries.forEach(([id, name]) => {
       if (deletedLogCount < 150) syncLog.deleted.push(name);
       else if (deletedLogCount === 150) syncLog.deleted.push("...and more.");
       deletedLogCount++;
    });

    for (let i = 0; i < deletedCount; i += 500) {
      const chunk = staleEntries.slice(i, i + 500).map(x => x[0]);
      try {
        People.People.batchDeleteContacts({ resourceNames: chunk });
        Utilities.sleep(500); 
      } catch (err) {
        chunk.forEach(res => { try { People.People.deleteContact(res); } catch(e) {} });
      }
    }
  }
  
  // Save final log and update state
  cache.put('SYNC_LOG', JSON.stringify(syncLog), 21600);
  cache.put('SYNC_UI_STATE', JSON.stringify({ phase: 'done', progress: total, total: total, deleted: deletedCount, created: createdCount, edited: editedCount, failed: failedCount, note: note }), 21600);

  const toastMessage = `🗑️ Del: ${deletedCount} | ☎️ New: ${createdCount} | ✏️ Edited: ${editedCount}` + (failedCount ? ` | ⚠️ Not saved: ${failedCount}` : "") + (note ? "\n" + note : "");
  if (note) console.warn("Phone sync: " + note);
  ss.toast(toastMessage, "🎉 Sync Complete!", 15);
  PropertiesService.getDocumentProperties().setProperty('MISSED_SYNC_MSG', toastMessage);

  if (ownerSync) ownerSync.afterSync(deletedCount, createdCount, editedCount, failedCount, note);
}

/** "Remove N contacts (over half)?" pop-up. No one there (a timed run): false, and nothing is asked. */
function askMassDeletion_(ss, count, labelCount) {
  let ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (uiErr) { ui = null; }
  if (!ui) {
    console.warn("Phone sync: would remove " + count + " of " + labelCount + " contacts; no one here to confirm, so they are kept.");
    return false;
  }
  try {
    const response = ui.alert(
      "⚠️ High Deletion Warning",
      `This sync would remove ${count} of your ${labelCount} School Directory contacts (more than half).\n\nRemove them? (No = sync everything else and keep them.)`,
      ui.ButtonSet.YES_NO
    );
    return response === ui.Button.YES;
  } catch (e) { return false; }
}

/** Ends a sync's run state (next sync starts fresh at row 1). */
function clearSyncRunState_(props, cache) {
  ['SYNC_START_ROW', 'SYNC_CREATED_COUNT', 'SYNC_EDITED_COUNT', 'SYNC_FAILED_COUNT', 'SYNC_DELETE_DECISION', 'SYNC_LIST_RETRIES']
    .forEach(k => props.deleteProperty(k));
  cache.remove('STALE_CHUNKS_COUNT'); // (the old version's saved list)
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "syncToGoogleContacts") ScriptApp.deleteTrigger(t);
  });
}

/**
 * Reads one Phone Contacts row (one phone + one email per row) for the sync.
 * Columns are found by header once; the numbers are the old fixed positions.
 * nameKey: how the contact will be named (for the clean-up's "is it still used").
 */
function phoneContactsRowReader_(headers) {
  const idx = (name, fallback) => headers.indexOf(name) !== -1 ? headers.indexOf(name) : fallback;
  const idxFirst = idx("First Name", 0);
  const idxLast = idx("Last Name", 1);
  const idxEmail = idx("E-mail 1 - Value", 2);
  const idxPhoneLabel = idx("Phone 1 - Type", 12);
  const idxPhone = idx("Phone 1 - Value", 13);
  const idxNotes = idx("Notes", 33);
  const idxOrgName = idx("Organization Name", 22);
  const idxOrgTitle = idx("Organization Department", 23);

  return row => {
    const parentName = row[idxFirst];
    return {
      parentName: parentName,
      nameKey: contactNameKey_(parentName, row[idxLast]),
      upsert: (directoryResourceName, index) => searchAndUpsertSingleContact(parentName, row[idxLast], row[idxEmail], row[idxPhone], row[idxPhoneLabel], row[idxNotes], row[idxOrgName], row[idxOrgTitle], directoryResourceName, index)
    };
  };
}

// ==========================================
// 2. SINGLE PUSH (Strict Mirroring)
// Returns { resourceName, isNew, isEdited, wrote, name, logStr }, or
// { failed: true, name, error, resourceName (its contact, kept) } when it could
// not be saved, or null for a row without a name. A Google "limit reached" error
// is thrown so the sync can pause (it used to be swallowed).
// ==========================================
function searchAndUpsertSingleContact(parentName, studentNames, email, phone, phoneLabel, notes, orgName, orgTitle, directoryResourceName = null, index = null) {
  const fullName = String(parentName || "").trim();
  const fNameStr = String(studentNames || "").trim(); // STRICT: We now track the student name
  if (!fullName) return null;
  let existingContact = null;
  try {
    const cache = CacheService.getScriptCache();
    const safeEmail = String(email || "").toLowerCase().trim();
    
    let safePhone = String(phone || "").replace(/\D/g, '');
    if (safePhone.length === 11 && safePhone.startsWith('1')) safePhone = safePhone.substring(1);

    // 1. STRICT CACHE KEY: Includes Student Name so different kids = different cache slots
    const rowIdentity = fullName.toLowerCase() + "|" + fNameStr.toLowerCase() + "|" + safeEmail + "|" + safePhone;
    
    let hash = 0;
    for (let i = 0; i < rowIdentity.length; i++) {
      hash = ((hash << 5) - hash) + rowIdentity.charCodeAt(i);
      hash |= 0; 
    }
    const cacheKey = "CONTACT_V5_" + Math.abs(hash); // Bumped version to clear the old cache

    // 2. STRICT MATCH: the Parent Name AND the Student Name, from the one list of contacts
    if (!index) index = buildContactsIndex_(directoryResourceName); // a single push outside the sync
    existingContact = findIndexedContact_(index, cacheKey, contactNameKey_(fullName, fNameStr));

    let isEdited = false;
    let editDetails = [];
    
    if (existingContact) {
      let currentEmail = existingContact.emailAddresses && existingContact.emailAddresses.length > 0 ? String(existingContact.emailAddresses[0].value).toLowerCase().trim() : "";
      let currentPhone = "";
      if (existingContact.phoneNumbers && existingContact.phoneNumbers.length > 0) {
        currentPhone = String(existingContact.phoneNumbers[0].value).replace(/\D/g, '');
        if (currentPhone.length === 11 && currentPhone.startsWith('1')) currentPhone = currentPhone.substring(1);
      }
      let currentNotes = existingContact.biographies && existingContact.biographies.length > 0 ? String(existingContact.biographies[0].value).trim() : "";
      
      if (currentEmail !== safeEmail) editDetails.push(`Email: ${currentEmail || 'Empty'} ➔ ${safeEmail || 'Empty'}`);
      if (currentPhone !== safePhone) editDetails.push(`Phone: ${currentPhone || 'Empty'} ➔ ${safePhone || 'Empty'}`);
      if (currentNotes !== String(notes || "").trim()) editDetails.push(`Notes updated`);
      
      if (editDetails.length > 0) isEdited = true; 
    }

    // Build the payload exactly as the row dictates. Zero assumptions.
    const contactData = {
      names: [{ givenName: fullName, familyName: fNameStr }],
      organizations: [],
      phoneNumbers: [],
      emailAddresses: (email && String(email).trim() !== "") ? [{ value: String(email), type: 'home' }] : [],
      biographies: notes ? [{ value: String(notes) }] : []
    };

    const orgNameValue = String(orgName || "").trim();
    const orgTitleValue = String(orgTitle || "").trim();
    if (orgNameValue !== "" || orgTitleValue !== "") {
      contactData.organizations.push({ name: orgNameValue, title: orgTitleValue });
    }
    
    if (phone && String(phone).trim() !== "") {
      const pLabel = String(phoneLabel || "").trim();
      contactData.phoneNumbers.push({ value: String(phone), type: pLabel !== "" ? pLabel : 'mobile' });
    }

    // Already exactly like the row (and in the label): nothing to write
    if (existingContact && contactMatchesPayload_(existingContact, contactData, directoryResourceName)) {
      cache.put(cacheKey, existingContact.resourceName, 21600);
      return { resourceName: existingContact.resourceName, isNew: false, isEdited: false, wrote: false, name: fullName };
    }
    
    let attempts = 0;
    
    while (attempts < 3) {
      attempts++;
      try {
        if (directoryResourceName) contactData.memberships = [{ contactGroupMembership: { contactGroupResourceName: directoryResourceName } }];

        if (existingContact) {
          contactData.etag = existingContact.etag;
          if (existingContact.memberships && directoryResourceName) {
            const others = existingContact.memberships.filter(m => m.contactGroupMembership?.contactGroupResourceName !== directoryResourceName);
            contactData.memberships = contactData.memberships.concat(others);
          }
          
          const updated = People.People.updateContact(contactData, existingContact.resourceName, { updatePersonFields: 'names,phoneNumbers,emailAddresses,biographies,memberships,organizations' });
          cache.put(cacheKey, updated.resourceName, 21600);
          index.add(storedContact_(contactData, updated));
          return { resourceName: updated.resourceName, isNew: false, isEdited: isEdited, wrote: true, name: fullName, logStr: editDetails.join('<br>') }; 
          
        } else {
          const created = People.People.createContact(contactData);
          cache.put(cacheKey, created.resourceName, 21600);
          index.add(storedContact_(contactData, created));
          return { resourceName: created.resourceName, isNew: true, isEdited: false, wrote: true, name: fullName }; 
        }
      } catch (updateErr) {
        const errorString = String(updateErr.message || updateErr || "");
        if (errorString.includes("not found") || errorString.includes("404")) {
           cache.remove('DIR_LABEL_LIVE_V2');
           throw updateErr;
        }
        if (isQuotaError_(updateErr)) throw updateErr;
        if (attempts >= 3) throw updateErr; 
        Utilities.sleep(2000 * attempts); 
      }
    }
  } catch (err) {
    if (isQuotaError_(err)) throw err; // the sync pauses 15 minutes and carries on
    console.error(`Sync error for ${parentName}: ${String(err.message || err)}`);
    return { failed: true, name: fullName, error: String(err.message || err), resourceName: existingContact ? existingContact.resourceName : null };
  }
}

function getSyncLogData() {
  const cache = CacheService.getScriptCache();
  const logStr = cache.get('SYNC_LOG');
  if (logStr) return JSON.parse(logStr);
  return { created: [], edited: [], deleted: [], failed: [] };
}

// ==========================================
// 🛑 BACKGROUND CONTROLS (Stop & Restart)
// ==========================================
function EMERGENCY_STOP_SYNC(isSilent) {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "syncToGoogleContacts") {
      ScriptApp.deleteTrigger(t);
    }
  });

  const cache = CacheService.getScriptCache();
  cache.put('SYNC_UI_STATE', JSON.stringify({ phase: 'killed' }), 21600); 
  cache.remove('DIR_LABEL'); 

  // Start row, counts (incl. the edited count, which used to stay behind), failures, the delete decision
  clearSyncRunState_(PropertiesService.getScriptProperties(), cache);

  if (isSilent !== true) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Sync has been safely cancelled.", "🛑 Stopped", 5);
  }
}

function RESTART_SYNC() {
  EMERGENCY_STOP_SYNC(true); 
  startResilientSync();
}
