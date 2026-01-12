/**
 * WORKFLOW MONITORING
 * Handles the logic for logging changes to Authorised trips into the Archive.
 * Triggered by the Installable Trigger in Triggers.gs.
 */

/**
 * TRIGGERED FUNCTION (via Installable Trigger)
 * Handles user edits made to a trip sheet AFTER it has been authorised.
 * Includes "Noise Filter" to ignore UI checkboxes.
 * @param {GoogleAppsScript.Events.AppsScriptEvent} e - The event object.
 */
function handlePostAuthEdit(e) {
  const sheet = e.range.getSheet();
  const col = e.range.getColumn();
  const row = e.range.getRow();
  
  // 1. FILTER OUT NOISE (Checkboxes & Buttons)
  // Column 2 (B) = Remove Students ticks
  // Column 17 (Q) = Add Students ticks
  // Row 33 = Master trigger buttons/checkboxes
  if (col === 2 || col === 17) return; 
  if (row === 33 && (col === 2 || col === 17 || col === 19)) return;

  // 2. Check Status & Log
  // We pass the 'e' object to extract specific user edit details
  processLogRequest(sheet, e);
}

/**
 * PUBLIC HELPER: Called by other scripts (Student Tools)
 * Logs an action taken by the system (script) rather than a direct cell edit.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The active sheet.
 * @param {string} message The description of the change (e.g. "Added 3 students").
 */
function logSystemAction(sheet, message) {
  // We pass 'null' for the event object, and provide a custom message
  processLogRequest(sheet, null, message);
}

/**
 * CORE LOGIC: Checks status, finds archive, and writes the log.
 * Handles both Event objects (e) and manual System messages.
 */
function processLogRequest(sheet, e = null, systemMessage = "") {
  // 1. Check Status using CONFIG constant
  let status;
  try {
    status = sheet.getRange(CONFIG.RANGE_STATUS).getValue();
  } catch (err) { return; }

  // Only proceed if Authorised or already Edited
  if (status !== CONFIG.STATUS_AUTHORISED && status !== CONFIG.STATUS_EDITED) {
    return;
  }

  // 2. Avoid infinite loops (ignore edits to status cell itself)
  if (e) {
    const statusRange = sheet.getRange(CONFIG.RANGE_STATUS);
    if (e.range.getRow() === statusRange.getRow() && 
        e.range.getColumn() === statusRange.getColumn()) {
      return;
    }
  }

  // 3. Find Archive
  const tripRef = sheet.getName(); 
  // NOTE: findArchiveSpreadsheet must be available in Helper.gs
  const archiveSS = findArchiveSpreadsheet(tripRef);
  if (!archiveSS) {
    Logger.log(`Could not find archive for ${tripRef} to log changes.`);
    return;
  }

  // 4. Prepare Log Data
  const timestamp = new Date();
  let user = Session.getActiveUser().getEmail();
  if (e && e.user) user = e.user.getEmail(); // Handle installable trigger user

  let cellRef = "System Action";
  let oldVal = "-";
  let newVal = systemMessage;

  // If this came from a real User Edit (e), grab the specific details
  if (e) {
    cellRef = e.range.getA1Notation();
    oldVal = e.oldValue || "";
    newVal = e.value || "";
  }

  // 5. Write to Archive
  // Uses CONFIG.SHEET_LOG ("log") to match your Template
  const logSheet = archiveSS.getSheetByName(CONFIG.SHEET_LOG);
  
  if (logSheet) {
    logSheet.appendRow([timestamp, user, cellRef, oldVal, newVal]);
  } else {
    // Fallback if template is broken/missing tab
    Logger.log(`Log sheet '${CONFIG.SHEET_LOG}' not found in Archive.`);
  }

  // 6. Update Status & Color (if not already marked Edited)
  if (status === CONFIG.STATUS_AUTHORISED) {
    // Change status to "Authorised (edited)"
    sheet.getRange(CONFIG.RANGE_STATUS).setValue(CONFIG.STATUS_EDITED);
    
    // Update Tab Color
    if (typeof updateSheetTabColor === 'function') {
      updateSheetTabColor(sheet, CONFIG.STATUS_EDITED);
    }
    
    // Notify User (Using getParent() to fix the toast bug)
    sheet.getParent().toast("Change recorded in Archive Log.", "Audit Trail");
  }
}
