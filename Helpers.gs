/**
 * HELPERS
 * Shared utility functions used across Triggers, Workflow, and Monitoring scripts.
 */

// =============================================================================
// DATA EXTRACTION HELPERS
// =============================================================================

/**
 * Reads the 2-column summary table and converts it into a data object.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The active sheet.
 * @returns {object} An object with all summary data, or null if it fails.
 */
function getSummaryData(sheet) {
  try {
    const range = sheet.getRange(CONFIG.SUMMARY_TABLE_RANGE_NAME);
    if (!range) return null; 

    const values = range.getValues();
    const data = values.reduce((obj, row, index) => {
      const key = row[0]; 
      const value = row[1];
      if (key) {
        obj[key] = value;
        // Store the A1 notation of the value cell for future updates
        obj[key + '_cellA1'] = range.getCell(index + 1, 2).getA1Notation();
      }
      return obj;
    }, {});
    
    return data;
  } catch (e) {
    console.error("Error getting summary data: " + e.message);
    return null; 
  }
}

/**
 * Gets data from a named range, skipping row 2 (headers) and filtering blank rows.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The active sheet.
 * @param {string} rangeName The named range to read.
 * @returns {{headers: string[], data: any[][]}}
 */
function getRangeData(sheet, rangeName) {
  const range = sheet.getRange(rangeName);
  if (!range) return { headers: [], data: [] };
  
  const values = range.getValues();
  // Assume Row 1 is headers, Row 2 is blank/spacer, Data starts Row 3?
  // Your snippet sliced(2), implying data starts at index 2 (row 3).
  const headers = values[0].filter(String); 
  const data = values.slice(2).filter(row => row[0] !== ""); 
  
  return { headers, data };
}

/**
 * Reads the 'authorisers_emails' named range and returns a flat list of emails.
 * @returns {string[]} An array of SLT email addresses.
 */
function getSltEmailList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const range = ss.getRangeByName(CONFIG.SLT_LIST_RANGE_NAME);
  if (!range) return [];
  return range.getValues().flat().filter(String);
}

// =============================================================================
// UI & FORMATTING HELPERS
// =============================================================================

/**
 * Updates a single value in the summary table and AUTOMATICALLY updates tab color.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The active sheet.
 * @param {string} key The key to find (e.g., "trip_status").
 * @param {*} value The new value to write.
 */
function updateSummaryValue(sheet, key, value) {
  const range = sheet.getRange(CONFIG.SUMMARY_TABLE_RANGE_NAME);
  const values = range.getValues();
  const rowIndex = values.findIndex(row => row[0] === key);
  
  if (rowIndex === -1) {
    throw new Error(`Could not find key "${key}" in summary table.`);
  }
  
  // Update the cell (Column 2)
  range.getCell(rowIndex + 1, 2).setValue(value);

  // Update Tab Color if the Status Changed
  // Note: Check against your specific key for status in the summary table
  if (key === 'trip_status' || key === 'Status') { 
    updateSheetTabColor(sheet, value);
  }
}

/**
 * Updates the sheet tab color based on the trip status.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The trip sheet.
 * @param {String} statusKey - The status string (must match a key in CONFIG.STATUS_COLORS).
 */
function updateSheetTabColor(sheet, statusKey) {
  if (!CONFIG.STATUS_COLORS) return;

  const color = CONFIG.STATUS_COLORS[statusKey];
  if (color) {
    sheet.setTabColor(color);
  } else {
    // If status is not found (e.g. unknown), reset or default
    sheet.setTabColor(null); 
  }
}

// =============================================================================
// CHECKBOX LOGIC (Optimized)
// =============================================================================

/**
 * Sets a specific range to FALSE (un-ticked) efficiently.
 * @param {String} rangeA1 - The A1 notation of the range (e.g., "Q35:Q").
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The active sheet.
 */
function setRangeToFalse(rangeA1, sheet) {
  try { 
    const range = sheet.getRange(rangeA1);
    const numRows = range.getNumRows();
    // Optimized: Create array in memory and set all at once
    range.setValues(Array(numRows).fill([false])); 
  } catch (e) {
    console.warn("setRangeToFalse failed for " + rangeA1);
  }
}

/**
 * Handles logic for Master Checkboxes that control lists of Slave Checkboxes.
 * Hardcoded to start processing from Row 35 as per recent logic.
 * @param {Sheet} sheet - The active sheet.
 * @param {String} master - A1 notation of the master checkbox (unused in logic but good for ref).
 * @param {String} target - A1 notation of the target column (slave checkboxes).
 * @param {String} dataRange - A1 notation of the data column to check for presence.
 * @param {String|Boolean} val - The new value of the master checkbox.
 */
function toggleTickBoxes(sheet, master, target, dataRange, val) {
  const isTrue = (val === 'TRUE' || val === true);
  const lastRow = sheet.getLastRow();
  
  // Logic specifically handles sheets where data starts at Row 35
  const rows = Math.max(0, lastRow - 35 + 1);
  if (rows <= 0) return;
  
  // Get ranges starting from row 35
  const tRange = sheet.getRange(35, sheet.getRange(target).getColumn(), rows, 1);
  const dRange = sheet.getRange(35, sheet.getRange(dataRange).getColumn(), rows, 1).getDisplayValues();
  
  // Map values: Only tick if data exists in dRange
  const newVals = dRange.map(r => [r[0] ? isTrue : false]);
  
  tRange.setValues(newVals);
}

// =============================================================================
// DRIVE & SYSTEM LOGGING
// =============================================================================

/**
 * Search the Archive folder for a spreadsheet containing the specific Trip Ref.
 * Uses Search Query (Fast) instead of Iterator (Slow).
 * @param {string} tripRef The trip reference (e.g., "2511T01").
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet | null} The Spreadsheet object, or null if not found.
 */
function findArchiveSpreadsheet(tripRef) {
  try {
    // FIX: Using CONFIG.FOLDER_ID_ARCHIVE from your Config.gs
    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID_ARCHIVE);
    
    // Search query for exact mimeType and name containment
    // Updated search to look for the 'A' suffix
    const files = folder.searchFiles(`title contains '${tripRef}A' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false`);
    
    if (files.hasNext()) {
      const file = files.next();
      return SpreadsheetApp.openById(file.getId());
    }
    return null;
  } catch (e) {
    console.error("Error finding archive: " + e.message);
    return null;
  }
}

/**
 * Appends a high-level action to the global 'log' tab (System Log).
 * Useful for debugging script executions.
 * @param {string} action The action type (e.g., "Trip Created", "Request Sent").
 * @param {string} tripRef The reference number of the trip.
 * @param {string} details Optional details.
 */
function logToSystem(action, tripRef, details = "") {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Default to "System Log" if not defined in Config
  const sheetName = CONFIG.SYSTEM_LOG_SHEET_NAME || "System Log";
  let logSheet = ss.getSheetByName(sheetName);

  // Create the log sheet if it doesn't exist
  if (!logSheet) {
    logSheet = ss.insertSheet(sheetName);
    logSheet.moveActiveSheet(ss.getNumSheets()); // Move to end
    logSheet.appendRow(["Timestamp", "User", "Trip Ref", "Action", "Details"]);
    logSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#efefef");
    logSheet.setFrozenRows(1);
  }

  const timestamp = new Date();
  const user = Session.getActiveUser().getEmail();

  logSheet.appendRow([timestamp, user, tripRef, action, details]);
}

/**
 * Creates the Archive using the Drive Template.
 * Copies data from the active trip sheet into the specific template tabs.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sourceSheet - The active trip sheet.
 * @param {String} tripRef - Trip Reference (e.g. 2511T01).
 * @param {String} tripName - Trip Name.
 * @return {String} The URL of the created archive.
 */
function createArchiveFromTemplate(sourceSheet, tripRef, tripName) {
  // 1. Setup Files & Folders
  const archiveFolder = DriveApp.getFolderById(CONFIG.FOLDER_ID_ARCHIVE);
  const templateFile = DriveApp.getFileById(CONFIG.ARCHIVE_TEMPLATE_ID);
  
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  
  // Updated naming convention: {Ref}A - {Name}
  const archiveFilename = `${tripRef}A - ${tripName} [${timestamp}]`;
  
  // 2. Make the Copy
  const newFile = templateFile.makeCopy(archiveFilename, archiveFolder);
  const archiveSS = SpreadsheetApp.open(newFile);
  
  // 3. Transfer Data to "summary" tab
  const summaryData = getSummaryData(sourceSheet); 
  const summaryTarget = archiveSS.getSheetByName(CONFIG.SHEET_SUMMARY);
  if (summaryTarget && summaryData) {
    const summaryArr = Object.entries(summaryData)
      .filter(([key]) => !key.endsWith('_cellA1')) // Remove internal keys
      .map(([key, val]) => [key, val]);
    
    if (summaryArr.length > 0) {
      // Write to Col A:B starting at Row 2
      summaryTarget.getRange(2, 1, summaryArr.length, 2).setValues(summaryArr);
    }
  }

  // 4. Transfer Student Data to "studentDetails"
  const studentTarget = archiveSS.getSheetByName(CONFIG.SHEET_STUDENTS);
  if (studentTarget) {
    const { headers, data } = getRangeData(sourceSheet, CONFIG.STUDENT_DATA_RANGE_NAME);
    if (data.length > 0) {
      studentTarget.getRange(2, 1, data.length, data[0].length).setValues(data);
    }
  }

  // 5. Transfer Medical Data to "allMedicalNotes"
  const medicalTarget = archiveSS.getSheetByName(CONFIG.SHEET_MEDICAL);
  if (medicalTarget) {
    const { headers, data } = getRangeData(sourceSheet, CONFIG.MEDICAL_DATA_RANGE_NAME);
    if (data.length > 0) {
      medicalTarget.getRange(2, 1, data.length, data[0].length).setValues(data);
    }
  }

  // 6. Initialize Log
  const logTarget = archiveSS.getSheetByName(CONFIG.SHEET_LOG);
  if (logTarget) {
    logTarget.appendRow([new Date(), Session.getActiveUser().getEmail(), "System", "-", "Archive Created"]);
  }

  return archiveSS.getUrl();
}
