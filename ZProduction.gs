// ==========================================
// 2. HELPER: GET DIRECTORY LABEL
// ==========================================
function getDirectoryLabelResourceName_() {
  const cache = CacheService.getScriptCache();
  let label = cache.get('DIR_LABEL_LIVE_V2'); 

  if (label) {
    try {
      People.ContactGroups.get(label);
      return label; 
    } catch(e) {
      cache.remove('DIR_LABEL_LIVE_V2'); 
    }
  }

  const groups = People.ContactGroups.list({ pageSize: 1000 }).contactGroups || [];
  const existingGroup = groups.find(g => g.name === "School Directory");
  
  if (existingGroup) {
    cache.put('DIR_LABEL_LIVE_V2', existingGroup.resourceName, 21600);
    return existingGroup.resourceName;
  } else {
    const newGroup = People.ContactGroups.create({ contactGroup: { name: "School Directory" } });
    cache.put('DIR_LABEL_LIVE_V2', newGroup.resourceName, 21600);
    return newGroup.resourceName;
  }
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
  cache.remove('STALE_CHUNKS_COUNT');
  cache.remove('DIR_LABEL'); 

  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('SYNC_START_ROW');
  props.deleteProperty('SYNC_CREATED_COUNT');

  if (isSilent !== true) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Sync has been safely cancelled.", "🛑 Stopped", 5);
  }
}

function RESTART_SYNC() {
  EMERGENCY_STOP_SYNC(true); 
  startResilientSync();
}
