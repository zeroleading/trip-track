/**
 * NAVIGATION HELPERS
 * Functions to quickly switch views/zoom levels for specific sections of the Trip Sheet.
 * Typically assigned to Drawing buttons on the sheet.
 */

/**
 * Sets up the sheet view for "Trip Details".
 * Focuses on the top section (Summary & details).
 */
function setupTripDetailsView() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.setFrozenRows(0);
  sheet.setFrozenRows(5);
  sheet.showRows(8, 21); // Ensure specific rows are visible
  sheet.getRange('A10').activate(); // Anchor
  SpreadsheetApp.flush();
  sheet.getRange('D10').activate(); // Focus
}

/**
 * Sets up the sheet view for "Risk Assessment".
 * Focuses on the RA section (Cols V-Y usually).
 */
function setupRiskAssessmentView() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.setFrozenRows(0);
  sheet.setFrozenRows(5);
  sheet.showRows(8, 21);
  sheet.getRange('Y10').activate();
  SpreadsheetApp.flush();
  sheet.getRange('V10').activate();
}

/**
 * Sets up the sheet view for "Students".
 * Freezes the top header but hides the middle section to maximize student list space.
 */
function setupStudentView() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.setFrozenRows(0);
  sheet.setFrozenRows(33); // Freeze below the headers
  sheet.hideRows(8, 21);   // Hide the trip details/RA section
  sheet.getRange('A35').activate();
  SpreadsheetApp.flush();
  sheet.getRange('D35').activate(); // Focus on Student Name column
}

/**
 * Sets up the sheet view for "Medical Notes".
 * Focuses on the Medical columns.
 */
function setupMedNotesView() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.setFrozenRows(0);
  sheet.setFrozenRows(33);
  sheet.hideRows(8, 21);
  sheet.getRange('AC35').activate();
  SpreadsheetApp.flush();
  sheet.getRange('V35').activate();
}
