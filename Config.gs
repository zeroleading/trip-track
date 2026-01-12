/**
 * CONFIGURATION FILE
 * Centralises all hardcoded IDs, named ranges, and status logic.
 * Organized for ease of maintenance.
 */
const CONFIG = {
  // ===========================================================================
  // 1. SYSTEM STRUCTURE (Folders, Templates, Permanent Sheets)
  // ===========================================================================
  // The folder where new trip sheets are built/stored
  FOLDER_ID_BUILD: "1_3jh60_8SsWHKu69EP50johgnvzEq-sG",
  
  // The folder where finalised/approved archives are stored
  FOLDER_ID_ARCHIVE: "1xz5FSwWxYYQsUtsnhOkPDvhXHxS3LeFb",
  
  // The Master Template ID used to create Archives
  ARCHIVE_TEMPLATE_ID: "1va96BaQAgws7Px_EXyxyVcAbVSRv5QcP_gj5oInw4ks",

  // Permanent System Sheets (Do not delete these from the spreadsheet)
  SYSTEM_LOG_SHEET_NAME: "log",           // Global logging tab
  TEMPLATE_SHEET_NAME: "0000T00",         // Master Trip Template tab

  // ===========================================================================
  // 2. NAMED RANGES (Internal to the Trip Sheets)
  // ===========================================================================
  // Core Trip Data
  SUMMARY_TABLE_RANGE_NAME: "thisTrip_summaryTable",     // Main trip details (Col A-B)
  RANGE_STATUS: "thisTrip_status",                       // Trip processing status cell
  STUDENT_DATA_RANGE_NAME: "thisTrip_studentData",       // Student list data (Col A-S)
  MEDICAL_DATA_RANGE_NAME: "thisTrip_medicalData",       // Medical notes summary block

  // Lists & Dropdowns
  REFERENCE_LIST_RANGE_NAME: "referenceList",            // List of generated Trip IDs
  SLT_SELECTED_RANGE_NAME: "authorisers_singleSelected", // The chosen SLT approver
  SLT_LIST_RANGE_NAME: "authorisers_emails",             // SLT email dropdown source
  
  // Officer Notifications
  RANGE_OFFICERS_ATTENDANCE: "officers_attendance",      // Attendance Officer emails
  RANGE_OFFICERS_COVER: "officers_cover",                // Cover Supervisor emails

  // ===========================================================================
  // 3. ARCHIVE MAPPING (Target tabs in the Archive File)
  // ===========================================================================
  // These must match the tab names in the ARCHIVE_TEMPLATE_ID file exactly
  SHEET_SUMMARY: "summary",
  SHEET_STUDENTS: "studentDetails",
  SHEET_MEDICAL: "allMedicalNotes",
  SHEET_LOG: "log",

  // ===========================================================================
  // 4. STATUS & VISUALS
  // ===========================================================================
  STATUS_AUTHORISED: "Authorised",
  STATUS_EDITED: "Authorised (edited)",

  // Maps status strings to hex colors for tab visual indicators
  STATUS_COLORS: {
    'Incomplete': '#bdbdbd',            // Grey
    'Pending authorisation': '#4285f4', // Blue
    'Authorised': '#34a853',            // Green
    'Authorised (edited)': '#b7e1cd',   // Light Green
    'Denied': '#ea4335'                 // Red
  },

  // ===========================================================================
  // 5. PERMISSIONS & SYSTEM ADMIN
  // ===========================================================================
  SYSTEM_ADMIN_BCC: "jappleton@csg.school",
  SYSTEM_TESTERS: ["jappleton@csg.school", "tnayagam@csg.school"]
};
