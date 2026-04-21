/**
 * WORKFLOW: DOCUMENTATION (FILE PICKER)
 * Handles the sidebar UI for searching Drive and attaching links to the sheet.
 * DESIGN: "Modern Blue" Theme.
 * STORAGE: Single-cell source of truth (URLs in Cell Value, Metadata in Cell Notes).
 * PERMISSIONS: Auto-grants View access to Testers & SLT Approver.
 */

/**
 * Opens the custom Drive Search sidebar.
 * Triggered from the "Trip Admin" menu.
 */
function openDocPicker() {
  const html = HtmlService.createHtmlOutput(PICKER_HTML_CONTENT)
      .setTitle('Link Drive Document')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Opens the Link Manager sidebar to view/remove existing links.
 * Triggered from the "Trip Admin" menu.
 */
function openLinkManager() {
  const html = HtmlService.createHtmlOutput(MANAGER_HTML_CONTENT)
      .setTitle('Manage Linked Docs')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * SERVER-SIDE: Searches Google Drive for files matching the query.
 * @param {string} query - The search text.
 * @return {Array} List of file objects {id, name, url, icon}.
 */
function searchDriveFiles(query) {
  if (!query || query.length < 3) return [];
  
  // Search for non-trashed files containing the title
  const files = DriveApp.searchFiles(`title contains '${query.replace(/'/g, "\\'")}' and trashed = false`);
  
  const results = [];
  let count = 0;
  
  // Limit to 15 results for performance
  while (files.hasNext() && count < 15) {
    const file = files.next();
    results.push({
      id: file.getId(),
      name: file.getName(),
      url: file.getUrl(),
      icon: getFileIcon(file.getMimeType())
    });
    count++;
  }
  return results;
}

/**
 * SERVER-SIDE: Attaches the selected file link to the Summary Table.
 * Uses the delimiter defined in Config to append to existing data.
 * ATTEMPTS to grant view permissions to System Testers & SLT.
 * @param {Object} fileData - The file details from the frontend.
 * @return {Object} Response payload containing success message and optional warning.
 */
function attachDriveFileToSheet(fileData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  // 1. Find the Target Cell
  const data = getSummaryData(sheet);
  if (!data) throw new Error("Could not read summary data.");

  const cellA1Key = CONFIG.SUMMARY_KEY_DOCS + '_cellA1';
  const cellA1 = data[cellA1Key];

  if (!cellA1) {
    throw new Error(`Label "${CONFIG.SUMMARY_KEY_DOCS}" not found in Summary Table.`);
  }

  const range = sheet.getRange(cellA1);
  
  // 2. Append URL to Cell Value
  const currentVal = range.getValue();
  const delimiter = (typeof CONFIG !== 'undefined' && CONFIG.SUMMARY_DOCS_DELIMITER) ? CONFIG.SUMMARY_DOCS_DELIMITER : "\n";
  
  let newVal = fileData.url;
  
  if (currentVal && String(currentVal).trim() !== "") {
    newVal = currentVal + delimiter + fileData.url;
  }
  
  range.setValue(newVal);
  
  // 3. Append Metadata to Cell Note
  const currentNote = range.getNote();
  const newNote = `Added: ${fileData.name} (${Session.getActiveUser().getEmail()})`;
  range.setNote(currentNote ? currentNote + "\n" + newNote : newNote);

  // 4. AUTOMATIC PERMISSIONS
  let warningMsg = null;
  
  try {
    const file = DriveApp.getFileById(fileData.id);
    
    // Start with System Testers
    const viewersToTx = [...(CONFIG.SYSTEM_TESTERS || [])];

    // Get SLT Approver from the current sheet
    const sltRange = sheet.getRange(CONFIG.SLT_SELECTED_RANGE_NAME);
    if (sltRange) {
      const sltEmail = sltRange.getValue();
      if (sltEmail && String(sltEmail).includes("@")) {
        viewersToTx.push(String(sltEmail));
      }
    }

    // Deduplicate emails using a Set
    const uniqueViewers = [...new Set(viewersToTx)];

    if (uniqueViewers.length > 0) {
      file.addViewers(uniqueViewers);
    }
    
  } catch (e) {
    // The Why: If the user lacks sharing rights, we catch the error to prevent a crash, 
    // alert them on the frontend, and email them a persistent reminder to ask the owner.
    console.warn(`Could not update permissions for ${fileData.name}: ${e.message}`);
    
    // Re-declare uniqueViewers logic purely for the warning message generation
    const viewersToTx = [...(CONFIG.SYSTEM_TESTERS || [])];
    const sltRange = sheet.getRange(CONFIG.SLT_SELECTED_RANGE_NAME);
    if (sltRange) {
      const sltEmail = sltRange.getValue();
      if (sltEmail && String(sltEmail).includes("@")) viewersToTx.push(String(sltEmail));
    }
    const uniqueViewers = [...new Set(viewersToTx)];
    
    // The Why: Attempt to determine who the user should ask for access.
    // Files in Shared Drives might not return an owner, so we fall back to editors.
    let whoToAsk = "the document owner or administrator";
    try {
      const owner = file.getOwner();
      if (owner) {
        whoToAsk = owner.getEmail();
      } else {
        const editors = file.getEditors();
        if (editors && editors.length > 0) {
          whoToAsk = editors.map(ed => ed.getEmail()).join(', ');
        }
      }
    } catch (ownerErr) {
      console.warn("Could not retrieve file owner/editors: " + ownerErr.message);
    }

    // The Why: Send a reminder email to the user so they have a persistent to-do ticket.
    const currentUserEmail = Session.getActiveUser().getEmail();
    const subject = `Action Required: Document Sharing Reminder (${fileData.name})`;
    const body = `Hello,\n\nYou recently attached the document "${fileData.name}" to the Trip Summary.\n\n` +
                 `However, you do not have the required permissions to automatically share this file. ` + 
                 `Please ask ${whoToAsk} to grant View access to the following required accounts:\n\n` +
                 `${uniqueViewers.join('\n')}\n\n` +
                 `Document Link: ${fileData.url}\n\n` +
                 `Thank you.`;
    
    try {
      MailApp.sendEmail(currentUserEmail, subject, body);
    } catch (mailErr) {
      console.warn("Failed to send reminder email: " + mailErr.message);
    }
    
    warningMsg = `You do not have permission to auto-share this file. A reminder email has been sent to you. Please ask ${whoToAsk} to grant View access to: ${uniqueViewers.join(', ')}`;
  }
  
  // The Why: Returning an object allows the frontend to distinguish between a perfect success and a success with caveats.
  return {
    message: `Attached! Cell now contains ${newVal.split(delimiter).length} link(s).`,
    warning: warningMsg
  };
}

/**
 * SERVER-SIDE: Fetches current links and their metadata (from notes).
 * Used by the Link Manager UI to display the list.
 */
function getLinkedDocs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = getSummaryData(sheet);
  if (!data) return [];

  const cellA1Key = CONFIG.SUMMARY_KEY_DOCS + '_cellA1';
  const cellA1 = data[cellA1Key];
  if (!cellA1) return [];

  const range = sheet.getRange(cellA1);
  const val = String(range.getValue());
  const note = range.getNote();

  if (!val || val.trim() === "") return [];

  const delimiter = (typeof CONFIG !== 'undefined' && CONFIG.SUMMARY_DOCS_DELIMITER) ? CONFIG.SUMMARY_DOCS_DELIMITER : "\n";
  
  const urls = val.split(delimiter);
  const notes = note ? note.split('\n') : [];

  return urls.map((url, index) => {
    let displayName = "Unknown File";
    let meta = "";

    if (notes[index]) {
      const parts = notes[index].replace("Added: ", "").split(" (");
      displayName = parts[0];
      if (parts[1]) meta = parts[1].replace(")", "");
    }

    return {
      index: index,
      url: url,
      name: displayName,
      meta: meta
    };
  });
}

/**
 * SERVER-SIDE: Removes a specific link by index.
 * Synchronizes removal from both the Cell Value and Cell Note.
 */
function removeLink(indexToRemove) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = getSummaryData(sheet);
  
  const cellA1Key = CONFIG.SUMMARY_KEY_DOCS + '_cellA1';
  const range = sheet.getRange(data[cellA1Key]);
  
  const val = String(range.getValue());
  const note = range.getNote();
  const delimiter = (typeof CONFIG !== 'undefined' && CONFIG.SUMMARY_DOCS_DELIMITER) ? CONFIG.SUMMARY_DOCS_DELIMITER : "\n";

  let urls = val.split(delimiter);
  let notes = note ? note.split('\n') : [];

  if (indexToRemove >= 0 && indexToRemove < urls.length) {
    // Remove from Value Array
    urls.splice(indexToRemove, 1);
    
    // Remove from Note Array (if exists)
    if (indexToRemove < notes.length) {
      notes.splice(indexToRemove, 1);
    }

    // Write back to sheet
    if (urls.length === 0) {
      range.clearContent();
      range.clearNote();
    } else {
      range.setValue(urls.join(delimiter));
      range.setNote(notes.join('\n'));
    }
    
    return "Removed link.";
  } else {
    throw new Error("Link not found at index " + indexToRemove);
  }
}

/**
 * Helper: Simple icon mapping for UI polish.
 */
function getFileIcon(mimeType) {
  if (mimeType.includes('spreadsheet')) return '📊';
  if (mimeType.includes('document')) return '📝';
  if (mimeType.includes('pdf')) return '📕';
  if (mimeType.includes('image')) return '🖼️';
  return '📄';
}

// =============================================================================
// CLIENT-SIDE UI (HTML/CSS/JS)
// =============================================================================

const PICKER_HTML_CONTENT = `
<!DOCTYPE html>
<html>
  <head>
    <base target="_blank">
    <link href="https://fonts.googleapis.com/css2?family=Google+Sans:wght@400;500;700&family=Roboto:wght@400;500&display=swap" rel="stylesheet">
    <style>
      body { font-family: 'Google Sans', Roboto, Arial, sans-serif; padding: 20px; font-size: 14px; color: #333; background-color: #f8f9fa; }
      
      h3 { font-size: 18px; margin: 0 0 5px 0; color: #1f2937; }
      p.subtitle { font-size: 12px; color: #6b7280; margin: 0 0 20px 0; line-height: 1.4; }
      
      .search-box {
        background: white;
        padding: 6px;
        border-radius: 12px;
        border: 1px solid #e5e7eb;
        box-shadow: 0 1px 2px rgba(0,0,0,0.05);
        display: flex;
        align-items: center;
        margin-bottom: 16px;
        transition: all 0.2s;
      }
      .search-box:focus-within { border-color: #4285f4; ring: 2px solid rgba(66, 133, 244, 0.3); }
      
      input { 
        width: 100%; border: none; outline: none; padding: 8px; font-size: 14px; color: #374151;
      }
      input::placeholder { color: #9ca3af; }

      button#searchBtn {
        width: 100%;
        background-color: #4285f4;
        color: white;
        border: none;
        padding: 10px;
        border-radius: 12px;
        font-weight: 500;
        cursor: pointer;
        box-shadow: 0 1px 3px rgba(66, 133, 244, 0.3);
        transition: transform 0.1s, background-color 0.2s;
      }
      button#searchBtn:hover { background-color: #3367d6; }
      button#searchBtn:active { transform: scale(0.98); }
      button#searchBtn:disabled { background-color: #cbd5e1; cursor: default; box-shadow: none; transform: none; }

      #results { margin-top: 24px; list-style: none; padding: 0; }
      .results-header { font-size: 11px; font-weight: 700; color: #9ca3af; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 8px; }
      
      .file-item {
        background: white;
        padding: 12px;
        border-radius: 8px;
        border: 1px solid #f3f4f6;
        box-shadow: 0 1px 2px rgba(0,0,0,0.05);
        cursor: pointer;
        display: flex;
        align-items: center;
        margin-bottom: 8px;
        transition: all 0.2s;
      }
      .file-item:hover { border-color: #4285f4; box-shadow: 0 4px 6px -1px rgba(66, 133, 244, 0.1); }
      
      .file-icon { 
        margin-right: 12px; font-size: 18px; 
        background: #f8fafc; padding: 6px; border-radius: 6px;
        transition: background 0.2s;
      }
      .file-item:hover .file-icon { background: rgba(66, 133, 244, 0.1); }
      
      .file-name { 
        white-space: nowrap; overflow: hidden; text-overflow: ellipsis; 
        font-weight: 500; color: #374151; font-size: 13px;
      }
      .file-item:hover .file-name { color: #111827; }

      #spinner { display: none; text-align: center; margin-top: 30px; }
      .spinner-icon {
        width: 24px; height: 24px; border: 3px solid #fbbc04; 
        border-top-color: transparent; border-radius: 50%; 
        animation: spin 1s linear infinite; margin: 0 auto 10px auto;
      }
      @keyframes spin { to { transform: rotate(360deg); } }

      #message { margin-top: 15px; border-radius: 6px; padding: 8px; font-size: 12px; display: none; }
      .error { color: #b91c1c; background: #fef2f2; border: 1px solid #fca5a5; display: block !important; }
      .success { color: #15803d; background: #f0fdf4; border: 1px solid #86efac; font-weight: 600; display: block !important; }
      
      /* The Why: Added a specific warning class to distinguish from standard success/error states */
      .warning { color: #9a3412; background: #fff7ed; border: 1px solid #fed7aa; display: block !important; white-space: pre-wrap; line-height: 1.4; }
    </style>
  </head>
  <body>
    <h3>Link Document</h3>
    <p class="subtitle">Search Drive to attach invoices, letters, or risk assessments.</p>
    
    <div class="search-box">
      <input type="text" id="searchInput" placeholder="Search filename (min 3 chars)..." onkeyup="checkEnter(event)">
    </div>
    
    <button id="searchBtn" onclick="runSearch()">Search Drive</button>
    
    <div id="spinner">
      <div class="spinner-icon"></div>
      <span style="color: #6b7280; font-size: 12px;">Searching...</span>
    </div>
    
    <div id="message"></div>
    
    <div id="results-container" style="display:none;">
      <div class="results-header">Search Results</div>
      <ul id="results"></ul>
    </div>

    <script>
      function checkEnter(e) {
        if (e.key === 'Enter') runSearch();
      }

      function runSearch() {
        const query = document.getElementById('searchInput').value;
        if (query.length < 3) {
          showMessage("Please enter at least 3 characters.", "error");
          return;
        }

        document.getElementById('spinner').style.display = 'block';
        document.getElementById('results-container').style.display = 'none';
        document.getElementById('results').innerHTML = '';
        document.getElementById('message').style.display = 'none';
        document.getElementById('searchBtn').disabled = true;

        google.script.run
          .withSuccessHandler(showResults)
          .withFailureHandler(showError)
          .searchDriveFiles(query);
      }

      function showResults(files) {
        document.getElementById('spinner').style.display = 'none';
        document.getElementById('searchBtn').disabled = false;
        
        if (files.length === 0) {
          showMessage("No files found.", "error");
          return;
        }

        document.getElementById('results-container').style.display = 'block';
        const list = document.getElementById('results');
        
        files.forEach(file => {
          const li = document.createElement('li');
          li.className = 'file-item';
          li.innerHTML = '<span class="file-icon">' + file.icon + '</span><span class="file-name">' + file.name + '</span>';
          li.onclick = function() { attachFile(file); };
          list.appendChild(li);
        });
      }

      function attachFile(file) {
        document.getElementById('spinner').style.display = 'block';
        document.querySelector('#spinner span').innerText = 'Attaching URL...';
        document.getElementById('results-container').style.display = 'none';
        
        google.script.run
          .withSuccessHandler(function(res) {
             document.getElementById('spinner').style.display = 'none';
             
             // The Why: We check if the server returned a warning property
             if (res.warning) {
                 showMessage("✅ " + res.message + "\\n\\n⚠️ Action Required: " + res.warning, "warning");
                 // The Why: Extend the timeout so the user actually has time to read the longer message
                 setTimeout(function() { google.script.host.close(); }, 8000);
             } else {
                 showMessage("✅ " + res.message, "success");
                 setTimeout(function() { google.script.host.close(); }, 2000);
             }
          })
          .withFailureHandler(showError)
          .attachDriveFileToSheet(file);
      }

      function showError(err) {
        document.getElementById('spinner').style.display = 'none';
        document.getElementById('searchBtn').disabled = false;
        showMessage("Error: " + err.message, "error");
      }

      function showMessage(msg, type) {
        const el = document.getElementById('message');
        el.innerText = msg;
        el.className = type;
        el.style.display = 'block';
      }
    </script>
  </body>
</html>
`;

const MANAGER_HTML_CONTENT = `
<!DOCTYPE html>
<html>
  <head>
    <base target="_blank">
    <link href="https://fonts.googleapis.com/css2?family=Google+Sans:wght@400;500;700&family=Roboto:wght@400;500&display=swap" rel="stylesheet">
    <style>
      body { font-family: 'Google Sans', Roboto, Arial, sans-serif; padding: 20px; font-size: 14px; color: #333; background-color: #f8f9fa; }
      
      h3 { font-size: 18px; margin: 0 0 5px 0; color: #1f2937; }
      p.subtitle { font-size: 12px; color: #6b7280; margin: 0 0 20px 0; }
      
      .link-item { 
        background: white;
        padding: 16px; 
        border-radius: 12px;
        border: 1px solid #e5e7eb;
        box-shadow: 0 1px 2px rgba(0,0,0,0.05);
        display: flex; 
        flex-direction: column;
        margin-bottom: 12px;
        transition: all 0.2s;
        position: relative;
        overflow: hidden;
      }
      .link-item:hover { border-color: #fbbc04; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); }
      
      .link-top { display: flex; align-items: flex-start; margin-bottom: 12px; }
      .check-icon { color: #4285f4; margin-right: 8px; font-size: 16px; margin-top: 1px; }
      
      .link-name { font-weight: 600; color: #1f2937; font-size: 14px; word-break: break-word; }
      .link-meta { font-size: 11px; color: #6b7280; display: flex; align-items: center; margin-top: 4px; }
      .badge { background: #f3f4f6; padding: 2px 6px; border-radius: 4px; margin-right: 6px; font-weight: 500; }

      .actions { display: flex; border-top: 1px solid #f3f4f6; padding-top: 8px; margin-top: 4px; }
      
      button { 
        flex: 1; border: none; background: transparent; 
        font-size: 11px; font-weight: 600; cursor: pointer; 
        padding: 6px; border-radius: 6px; display: flex; 
        align-items: center; justify-content: center;
        transition: background 0.2s;
      }
      
      .btn-open { color: #6b7280; margin-right: 4px; }
      .btn-open:hover { color: #4285f4; background: #eff6ff; }
      
      .btn-delete { color: #ef4444; margin-left: 4px; }
      .btn-delete:hover { color: #b91c1c; background: #fef2f2; }
      
      .empty-state { text-align: center; color: #9ca3af; margin-top: 40px; font-style: italic; }
      #spinner { text-align: center; margin-top: 20px; color: #6b7280; }
    </style>
  </head>
  <body>
    <h3>Manage Links</h3>
    <p class="subtitle">Current documents attached to this trip.</p>
    
    <div id="spinner">Loading links...</div>
    <div id="list"></div>

    <script>
      window.onload = loadLinks;

      function loadLinks() {
        google.script.run
          .withSuccessHandler(renderList)
          .withFailureHandler(showError)
          .getLinkedDocs();
      }

      function renderList(links) {
        const container = document.getElementById('list');
        document.getElementById('spinner').style.display = 'none';
        container.innerHTML = '';

        if (links.length === 0) {
          container.innerHTML = '<div class="empty-state">No documents linked yet.</div>';
          return;
        }

        links.forEach(link => {
          const div = document.createElement('div');
          div.className = 'link-item';
          
          // Clean up meta (email)
          const metaShort = link.meta ? link.meta.split('@')[0] : 'Unknown';

          div.innerHTML = \`
            <div class="link-top">
              <span class="check-icon">✓</span>
              <div>
                <div class="link-name">\${link.name}</div>
                <div class="link-meta"><span class="badge">User</span> \${metaShort}</div>
              </div>
            </div>
            
            <div class="actions">
              <button class="btn-open" onclick="window.open('\${link.url}', '_blank')">OPEN FILE</button>
              <div style="width:1px; background:#e5e7eb; margin: 4px 0;"></div>
              <button class="btn-delete" onclick="deleteLink(\${link.index})">REMOVE</button>
            </div>
          \`;
          container.appendChild(div);
        });
      }

      function deleteLink(index) {
        if (!confirm('Are you sure you want to remove this link?')) return;
        
        document.getElementById('spinner').style.display = 'block';
        document.getElementById('spinner').innerText = 'Removing...';
        document.getElementById('list').style.opacity = '0.5';

        google.script.run
          .withSuccessHandler(function() {
             loadLinks(); // Reload list
             document.getElementById('list').style.opacity = '1';
          })
          .withFailureHandler(showError)
          .removeLink(index);
      }

      function showError(err) {
        alert('Error: ' + err.message);
        document.getElementById('spinner').style.display = 'none';
        document.getElementById('list').style.opacity = '1';
      }
    </script>
  </body>
</html>
`;