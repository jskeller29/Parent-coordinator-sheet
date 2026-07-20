// File: Phone.gs
// ==========================================
// 1. BULK NIGHTLY SYNC (Enterprise Mirror Mode)
// ==========================================
function syncToGoogleContacts() {
  const startTime = Date.now();
  const MAX_RUN_TIME = 4.5 * 60 * 1000; 
  
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
  if (!ss) return;   
  
  const sheet = ss.getSheetByName("Phone Contacts");
  if (!sheet) return;
  
  const data = sheet.getDataRange().getValues();
  const props = PropertiesService.getScriptProperties();
  const cache = CacheService.getScriptCache();
  
  let startRow = parseInt(props.getProperty('SYNC_START_ROW')) || 1;
  let createdCount = parseInt(props.getProperty('SYNC_CREATED_COUNT')) || 0; 
  let editedCount = parseInt(props.getProperty('SYNC_EDITED_COUNT')) || 0; 
  
  let syncLog = JSON.parse(cache.get('SYNC_LOG') || '{"created":[], "edited":[], "deleted":[]}');
  
  function saveStaleMap(map) {
    const entries = Object.entries(map);
    const chunkCount = Math.ceil(entries.length / 1000);
    cache.put('STALE_CHUNKS_COUNT', chunkCount.toString(), 21600);
    for (let c = 0; c < chunkCount; c++) {
      cache.put('STALE_CHUNK_' + c, JSON.stringify(entries.slice(c * 1000, (c + 1) * 1000)), 21600);
    }
  }

  // 1. Get or Create Label
  // pageSize matters: the default page is small, so without it a user with many
  // contact groups may not find "School Directory" and we'd create a duplicate.
  let resourceName = "";
  const groups = People.ContactGroups.list({ pageSize: 1000 }).contactGroups || [];
  const existingGroup = groups.find(g => g.name === "School Directory");
  
  if (existingGroup) {
    resourceName = existingGroup.resourceName;
  } else {
    const newGroup = People.ContactGroups.create({ contactGroup: { name: "School Directory" } });
    resourceName = newGroup.resourceName;
  }
  
  // 2. Prepare Stale List 
  let staleMap = {}; 
  if (startRow === 1) {
    ss.toast("Phase 1/3: Checking existing Google contacts...", "Syncing", 10);
    
    // --- THE FIX: Aggressively wipe local memory so old names don't bleed into new runs! ---
    syncLog = { "created": [], "edited": [], "deleted": [] };
    cache.put('SYNC_LOG', JSON.stringify(syncLog), 21600);
    // ---------------------------------------------------------------------------------------

    if (existingGroup) {
      try {
        let pageToken;
        do {
          const response = People.People.Connections.list('people/me', { 
            pageSize: 1000, 
            personFields: 'names,memberships', 
            pageToken: pageToken 
          });
          if (response.connections) {
            response.connections.forEach(p => {
              if (p.memberships && p.memberships.some(m => m.contactGroupMembership?.contactGroupResourceName === resourceName)) {
                const displayName = (p.names && p.names.length > 0) ? p.names[0].displayName : "Unknown Contact";
                staleMap[p.resourceName] = displayName; 
              }
            });
          }
          pageToken = response.nextPageToken;
        } while (pageToken);
        try { saveStaleMap(staleMap); } catch(e) { console.error("saveStaleMap failed: " + (e && e.message)); }
      } catch (e) {
        // Listing existing contacts failed. Log it and continue with an empty stale
        // map (the cleanup phase will simply delete nothing this run rather than
        // deleting from a partial list).
        console.error("Building stale-contact map failed: " + (e && e.message));
      }
    }
  } else {
    const chunkCount = parseInt(cache.get('STALE_CHUNKS_COUNT')) || 0;
    for (let c = 0; c < chunkCount; c++) {
      const chunkData = cache.get('STALE_CHUNK_' + c);
      if (chunkData) JSON.parse(chunkData).forEach(([id, name]) => staleMap[id] = name);
    }
  }
  
  // 3. Sync Rows
  ss.toast("Phase 2/3: Syncing spreadsheet rows... Safe to close sidebar.", "Syncing", 10);
  for (let i = startRow; i < data.length; i++) {
    
    if (Date.now() - startTime > MAX_RUN_TIME) {
      props.setProperty('SYNC_START_ROW', i.toString());
      props.setProperty('SYNC_CREATED_COUNT', createdCount.toString()); 
      props.setProperty('SYNC_EDITED_COUNT', editedCount.toString()); 
      cache.put('SYNC_LOG', JSON.stringify(syncLog), 21600); 
      try { saveStaleMap(staleMap); } catch(e){ console.error(e); }
      
      cache.put('SYNC_UI_STATE', JSON.stringify({ phase: 'paused', progress: i, total: data.length - 1 }), 21600);
      
      // --- THE FIX: Delete existing sync triggers before creating a new one ---
      ScriptApp.getProjectTriggers().forEach(t => {
        if (t.getHandlerFunction() === "syncToGoogleContacts") ScriptApp.deleteTrigger(t);
      });
      
      ScriptApp.newTrigger("syncToGoogleContacts").timeBased().after(60 * 1000).create();
      ss.toast("Timer reset. Resuming at row " + i, "Auto-Resume Active", 30);
      return; 
    }
    
    if (i % 5 === 0) {
      const currentStateStr = cache.get('SYNC_UI_STATE');
      if (currentStateStr && JSON.parse(currentStateStr).phase === 'killed') return; 
      cache.put('SYNC_UI_STATE', JSON.stringify({ phase: 'syncing', progress: i, total: data.length - 1 }), 21600);
    }
    
    const headers = data[0];
    const idxFirst = headers.indexOf("First Name") !== -1 ? headers.indexOf("First Name") : 0;
    const idxLast = headers.indexOf("Last Name") !== -1 ? headers.indexOf("Last Name") : 1;
    const idxEmail = headers.indexOf("E-mail 1 - Value") !== -1 ? headers.indexOf("E-mail 1 - Value") : 2;
    const idxPhoneLabel = headers.indexOf("Phone 1 - Type") !== -1 ? headers.indexOf("Phone 1 - Type") : 12;
    const idxPhone = headers.indexOf("Phone 1 - Value") !== -1 ? headers.indexOf("Phone 1 - Value") : 13;
    const idxNotes = headers.indexOf("Notes") !== -1 ? headers.indexOf("Notes") : 33;
    const idxOrgName = headers.indexOf("Organization Name") !== -1 ? headers.indexOf("Organization Name") : 22;
    const idxOrgTitle = headers.indexOf("Organization Department") !== -1 ? headers.indexOf("Organization Department") : 23;

    const parentName = data[i][idxFirst];
    const studentNames = data[i][idxLast];
    const email = data[i][idxEmail];
    const phoneLabel = data[i][idxPhoneLabel];
    const phone = data[i][idxPhone];
    const notes = data[i][idxNotes];
    const orgName = data[i][idxOrgName];
    const orgTitle = data[i][idxOrgTitle];

    if (!parentName || String(parentName).trim() === "") continue;
    
    try {
      const result = searchAndUpsertSingleContact(parentName, studentNames, email, phone, phoneLabel, notes, orgName, orgTitle, resourceName);
      
      if (result) {
        delete staleMap[result.resourceName]; 
        
        if (result.isNew) {
          createdCount++; 
          if (syncLog.created.length < 150) syncLog.created.push(result.name);
          else if (syncLog.created.length === 150) syncLog.created.push("...and more.");
        } else if (result.isEdited) {
          editedCount++; 
          if (syncLog.edited.length < 150) syncLog.edited.push(`<b>${result.name}</b><br><span style="color:#555; font-size:11px;">${result.logStr}</span>`);
          else if (syncLog.edited.length === 150) syncLog.edited.push("...and more.");
        }
      }
    } catch (e) {
      if (e.message.toLowerCase().includes("quota") || e.message.includes("429")) {
        props.setProperty('SYNC_START_ROW', i.toString());
        props.setProperty('SYNC_CREATED_COUNT', createdCount.toString());
        props.setProperty('SYNC_EDITED_COUNT', editedCount.toString());
        cache.put('SYNC_LOG', JSON.stringify(syncLog), 21600); 
        try { saveStaleMap(staleMap); } catch(e){ console.error(e); }
        
        cache.put('SYNC_UI_STATE', JSON.stringify({ phase: 'quota', progress: i, total: data.length - 1 }), 21600);

        // --- THE FIX: Delete existing sync triggers before creating a new one ---
        ScriptApp.getProjectTriggers().forEach(t => {
          if (t.getHandlerFunction() === "syncToGoogleContacts") ScriptApp.deleteTrigger(t);
        });

        ScriptApp.newTrigger("syncToGoogleContacts").timeBased().after(15 * 60 * 1000).create();
        ss.toast("API Limit reached. Pausing 15 mins.", "Quota Pause", 60);
        return; 
      }
    }
    Utilities.sleep(750); 
  }
  
  // 4. Cleanup Phase
  ss.toast("Phase 3/3: Removing old contacts...", "Final Cleanup", 10);
  const staleEntries = Object.entries(staleMap);
  const deletedCount = staleEntries.length; 
  
  if (startRow === 1 && deletedCount > (data.length * 0.5) && data.length > 10) {
     // getUi() throws when this runs from a time-based / auto-resume trigger
     // (nightly sync, quota resume). Probe for a UI; if there isn't one we must
     // NOT guess our way into deleting most of the list unattended.
     let ui = null;
     try { ui = SpreadsheetApp.getUi(); } catch (uiErr) { ui = null; }

     if (ui) {
       const response = ui.alert(
         "⚠️ High Deletion Warning",
         `You are about to delete ${deletedCount} contacts (more than 50% of your list).\n\nAre you sure you want to proceed and delete these contacts?`,
         ui.ButtonSet.YES_NO
       );
       if (response !== ui.Button.YES) {
         ss.toast("Sync aborted by user to prevent mass deletion.", "Aborted", 5);
         return;
       }
     } else {
       // Unattended run: skip the mass deletion, keep the contacts, and surface a
       // message on next open. Clean up state/triggers so we don't re-warn nightly.
       const skipMsg = `⚠️ Skipped removing ${deletedCount} contacts (over 50% of the list) during an automatic sync. Open the sheet and run "Full Sync" to confirm these removals.`;
       console.warn(skipMsg);
       try { ss.toast(skipMsg, "Mass Deletion Skipped", 15); } catch (e) {}
       PropertiesService.getDocumentProperties().setProperty('MISSED_SYNC_MSG', skipMsg);

       props.deleteProperty('SYNC_START_ROW');
       props.deleteProperty('SYNC_CREATED_COUNT');
       props.deleteProperty('SYNC_EDITED_COUNT');
       cache.remove('STALE_CHUNKS_COUNT');
       ScriptApp.getProjectTriggers().forEach(t => {
         if (t.getHandlerFunction() === "syncToGoogleContacts") ScriptApp.deleteTrigger(t);
       });

       cache.put('SYNC_LOG', JSON.stringify(syncLog), 21600);
       cache.put('SYNC_UI_STATE', JSON.stringify({ phase: 'done', progress: data.length - 1, total: data.length - 1, deleted: 0, created: createdCount, edited: editedCount }), 21600);
       return;
     }
  }

  props.deleteProperty('SYNC_START_ROW');
  props.deleteProperty('SYNC_CREATED_COUNT'); 
  props.deleteProperty('SYNC_EDITED_COUNT'); 
  cache.remove('STALE_CHUNKS_COUNT'); 
  ScriptApp.getProjectTriggers().forEach(t => { 
    if(t.getHandlerFunction() === "syncToGoogleContacts") ScriptApp.deleteTrigger(t); 
  });
  
  if (deletedCount > 0) {
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
  cache.put('SYNC_UI_STATE', JSON.stringify({ phase: 'done', progress: data.length - 1, total: data.length - 1, deleted: deletedCount, created: createdCount, edited: editedCount }), 21600);

  const toastMessage = `🗑️ Del: ${deletedCount} | ☎️ New: ${createdCount} | ✏️ Edited: ${editedCount}`;
  ss.toast(toastMessage, "🎉 Sync Complete!", 15);
  PropertiesService.getDocumentProperties().setProperty('MISSED_SYNC_MSG', toastMessage);
}

// ==========================================
// 2. SINGLE PUSH (Strict Mirroring & Anti-Lag Logic)
// ==========================================
function searchAndUpsertSingleContact(parentName, studentNames, email, phone, phoneLabel, notes, orgName, orgTitle, directoryResourceName = null) {
  try {
    const fullName = String(parentName || "").trim();
    const fNameStr = String(studentNames || "").trim(); // STRICT: We now track the student name
    if (!fullName) return null;

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

    let existingContact = null;
    let cachedResourceName = cache.get(cacheKey);

    if (cachedResourceName) {
      try {
        existingContact = People.People.get(cachedResourceName, { personFields: 'names,phoneNumbers,emailAddresses,memberships,organizations,biographies' });
      } catch (e) { 
        cachedResourceName = null; 
        cache.remove(cacheKey);
      }
    }

    if (!existingContact) {
      let pageToken;
      do {
        let apiOptions = { 
          pageSize: 1000, 
          personFields: 'names,phoneNumbers,emailAddresses,memberships,organizations,biographies' 
        };
        if (pageToken) apiOptions.pageToken = pageToken;

        const response = People.People.Connections.list('people/me', apiOptions);

        if (response.connections) {
          for (let i = 0; i < response.connections.length; i++) {
            const person = response.connections[i];
            
            if (person.names && person.names.length > 0) {
              const apiGiven = String(person.names[0].givenName || "").toLowerCase().trim();
              const apiFamily = String(person.names[0].familyName || "").toLowerCase().trim();
              
              // 2. STRICT API MATCH: Must perfectly match BOTH the Parent Name AND Student Name
              if (apiGiven === fullName.toLowerCase() && apiFamily === fNameStr.toLowerCase()) {
                existingContact = person;
                break; 
              }
            }
          }
        }
        if (existingContact) break;
        pageToken = response.nextPageToken;
      } while (pageToken);
    }

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
    
    let attempts = 0;
    let success = false;
    
    while (!success && attempts < 3) {
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
          return { resourceName: updated.resourceName, isNew: false, isEdited: isEdited, name: fullName, logStr: editDetails.join('<br>') }; 
          
        } else {
          const created = People.People.createContact(contactData);
          cache.put(cacheKey, created.resourceName, 21600);
          return { resourceName: created.resourceName, isNew: true, isEdited: false, name: fullName }; 
        }
      } catch (updateErr) {
        const errorString = String(updateErr.message || updateErr || "");
        if (errorString.includes("not found") || errorString.includes("404")) {
           cache.remove('DIR_LABEL_LIVE_V2');
           throw updateErr;
        }
        if (errorString.includes("quota") || errorString.includes("429")) throw updateErr;
        if (attempts >= 3) throw updateErr; 
        Utilities.sleep(2000 * attempts); 
      }
    }
  } catch (err) {
    console.error(`Sync error for ${parentName}: ${String(err.message || err)}`);
    return null; 
  }
}

// ==========================================
// 3. UI TRIGGER & SIDEBAR HELPERS
// ==========================================
function runInitialBulkContactSync() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("Starting Mirror Sync. This may take several minutes for large lists.", "Syncing", 8);
  syncToGoogleContacts();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('Persistent Info');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getSyncLogData() {
  const cache = CacheService.getScriptCache();
  const logStr = cache.get('SYNC_LOG');
  if (logStr) return JSON.parse(logStr);
  return { created: [], edited: [], deleted: [] };
}
