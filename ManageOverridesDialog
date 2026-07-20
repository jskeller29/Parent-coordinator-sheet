<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.2.19/dist/tailwind.min.css" rel="stylesheet">
  <style>
    ::-webkit-scrollbar { width: 8px; }
    ::-webkit-scrollbar-track { background: #f1f5f9; }
    ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }
  </style>
</head>
<body class="bg-gray-100 text-gray-800 font-sans p-4">
  
  <div class="mb-4 pb-2 border-b border-gray-300">
    <h2 class="text-xl font-bold text-gray-700">Manage Overrides</h2>
    <p class="text-xs text-gray-500">View, edit, or delete any manual student modifications. Deleting an entry will restore their default ATS data.</p>
  </div>

  <div id="loading" class="text-center text-blue-600 py-8 font-semibold animate-pulse">
    Loading your custom entries...
  </div>
  
  <div id="list" class="space-y-3 pb-12"></div>
  
  <!-- Close Button -->
  <div class="fixed bottom-0 left-0 right-0 bg-white p-3 border-t text-right shadow-lg">
     <button onclick="closeDialog()" class="px-4 py-2 text-sm font-medium text-gray-700 bg-gray-100 border border-gray-300 rounded hover:bg-gray-200 shadow-sm focus:outline-none focus:ring-2 focus:ring-gray-400">Close</button>
  </div>

  <script>
    window.onload = function() {
       if (typeof google !== 'undefined' && google.script && google.script.run) {
           google.script.run.withSuccessHandler(renderList).getOverridesList();
       } else {
           // Local test fallback
           renderList([
               {lastName: "Simpson", firstName: "Bart", osis: "123456789", classCode: "Y41", type: "Edit Existing Student", date: "05/12/2026"}
           ]);
       }
    }

    function renderList(overrides) {
        document.getElementById('loading').style.display = 'none';
        const list = document.getElementById('list');
        
        if(overrides.length === 0) {
            list.innerHTML = `
            <div class="text-center py-10 bg-white rounded-lg shadow-sm border border-gray-200">
                <span class="text-3xl block mb-2">📋</span>
                <p class="text-gray-500 font-medium">No manual overrides found.</p>
                <p class="text-xs text-gray-400 mt-1">Students added or edited via the menu will appear here.</p>
            </div>`;
            return;
        }

        let html = '';
        overrides.forEach(o => {
            const badgeColor = o.type.includes("New") ? "bg-green-100 text-green-700" : "bg-purple-100 text-purple-700";
            
            html += `
            <div class="bg-white p-4 border border-gray-200 rounded-lg shadow-sm hover:shadow-md transition-shadow">
               <div class="flex justify-between items-start mb-2">
                 <div>
                   <p class="font-bold text-gray-800 text-lg">${o.lastName}, ${o.firstName}</p>
                   <p class="text-xs text-gray-500 mt-0.5"><b>OSIS:</b> ${o.osis} &nbsp;|&nbsp; <b>Class:</b> ${o.classCode}</p>
                 </div>
                 <span class="text-[10px] font-bold px-2 py-1 rounded ${badgeColor} whitespace-nowrap">${o.type}</span>
               </div>
               
               <div class="flex justify-between items-end mt-3 pt-3 border-t border-gray-100">
                 <p class="text-xs text-gray-400">Modified on: ${o.date}</p>
                 <div class="flex space-x-2">
                   <button onclick="editOsis('${o.osis}')" class="px-3 py-1.5 bg-blue-50 text-blue-700 border border-blue-200 rounded text-xs font-semibold hover:bg-blue-100 transition">✏️ Edit</button>
                   <button onclick="deleteOsis('${o.osis}', '${o.firstName}')" class="px-3 py-1.5 bg-red-50 text-red-700 border border-red-200 rounded text-xs font-semibold hover:bg-red-100 transition">🗑️ Delete</button>
                 </div>
               </div>
            </div>`;
        });
        list.innerHTML = html;
    }

    function editOsis(osis) {
        document.getElementById('list').innerHTML = '<div class="text-center py-8 text-blue-500 font-semibold animate-pulse">Launching Editor...</div>';
        if (typeof google !== 'undefined' && google.script && google.script.run) {
            // Tells the backend to save the OSIS, close this screen, and open the edit screen!
            google.script.run.withSuccessHandler(() => google.script.host.close()).setPrefillAndOpenEdit(osis);
        }
    }

    function deleteOsis(osis, firstName) {
        if(!confirm(`Are you sure you want to permanently delete the override for ${firstName}?\n\nIf they are a standard ATS student, their data will revert to the default ATS information.`)) return;
        
        document.getElementById('list').innerHTML = '<div class="text-center py-8 text-red-500 font-semibold animate-pulse">Deleting entry & rebuilding database...</div>';
        
        if (typeof google !== 'undefined' && google.script && google.script.run) {
            google.script.run.withSuccessHandler(() => {
                // Refresh the list seamlessly after deletion
                document.getElementById('list').innerHTML = '<div class="text-center py-8 text-blue-500 font-semibold animate-pulse">Refreshing...</div>';
                google.script.run.withSuccessHandler(renderList).getOverridesList();
            }).deleteOverride(osis);
        }
    }

    function closeDialog() {
       if (typeof google !== 'undefined' && google.script && google.script.host) {
          google.script.host.close();
       }
    }
  </script>
</body>
</html>
