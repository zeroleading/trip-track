/**
 * TRIGGERS
 * Handles global events for the spreadsheet.
 */

/**
 * Runs when the spreadsheet is opened. 
 * Creates a custom "Trip Admin" menu.
 */
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Trip Admin')
    .addItem('Add new trip', 'addTrip')
    .addItem('View previous trips', 'viewPrevious')
    .addSeparator()
    .addItem('1. Request Authorisation', 'requestAuthorisation')
    .addItem('2. Approve This Trip (SLT)', 'approveTripWorkflow')
    .addItem('3. Deny This Trip (SLT)', 'denyTripWorkflow')
    .addSeparator()
    .addItem('Add documents to this trip', 'openDocPicker') 
    .addItem('Manage linked documents for this trip', 'openLinkManager') // ADDED ITEM
    .addSeparator()
    .addItem('Force notifications (manual)', 'forceTripNotifications')
    .addSeparator()
    .addItem('View trip evaluation', 'viewTripEvaluation')
    .addSeparator()
    .addSubMenu(ui.createMenu('⚠️ Testing Area')
        .addItem('Test: Request Authorisation', 'test_RequestAuth')
        .addItem('Test: Approval Email', 'test_ApproveTrip')
        .addItem('Test: Denial Email', 'test_DenyTrip')
        .addSeparator()
        .addItem('Test: T-7 Lock-in', 'test_T7')
        .addItem('Test: T-4 Operations (Att/Cov)', 'test_T4')
        .addItem('Test: T-1 Leader Pack', 'test_T1')
        .addItem('Test: T-0 Reminder', 'test_T0')
        .addItem('Test: T+1 Evaluations', 'test_EvaluationEmail')
        .addSeparator()
        .addItem('🔴 Test: Run ALL Email Tests', 'test_RunAllEmails')
        .addSeparator()
        .addItem('Add dummy evaluation (for testing)', 'test_SeedDummyResponse'))
    .addToUi();
}

/**
 * SIMPLE TRIGGER: onEdit
 * Handles FAST UI updates (Tick boxes, Tab Colors).
 */
function onEdit(e) {
  if (!e) return;

  const editedRange = e.range;
  const activeSheet = editedRange.getSheet();
  const editedCellA1Notation = editedRange.getA1Notation();
  const currentSheetName = activeSheet.getName();
  const newEditedCellValue = e.value;
  
  const validSheetNamePattern = /^\d{4}T\d{2}$/; 

  // Configuration for Tick Boxes
  const resetAddStudentsTicksTriggerCell = 'S33';
  const addStudentsTicksRange = 'Q35:Q';
  const masterAddStudentsTickBox = 'Q33';
  const addStudentsTickBoxesRange = 'Q35:Q';
  const addStudentsDataPresenceRange = 'S35:S';
  const masterRemoveStudentsTickBox = 'B33';
  const removeStudentsTickBoxesRange = 'B35:B';
  const removeStudentsDataPresenceRange = 'D35:D';

  if (!validSheetNamePattern.test(currentSheetName)) return;

  // 1. UPDATE TAB COLOR
  if (CONFIG.STATUS_COLORS && CONFIG.STATUS_COLORS[newEditedCellValue]) {
    updateSheetTabColor(activeSheet, newEditedCellValue);
  }

  // 2. STUDENT TICK BOX LOGIC
  if (editedCellA1Notation === resetAddStudentsTicksTriggerCell) {
    setRangeToFalse(addStudentsTicksRange, activeSheet);
    return; 
  }
  if (editedCellA1Notation === masterRemoveStudentsTickBox) {
    toggleTickBoxes(activeSheet, masterRemoveStudentsTickBox, removeStudentsTickBoxesRange, removeStudentsDataPresenceRange, newEditedCellValue);
    return;
  }
  if (editedCellA1Notation === masterAddStudentsTickBox) {
    toggleTickBoxes(activeSheet, masterAddStudentsTickBox, addStudentsTickBoxesRange, addStudentsDataPresenceRange, newEditedCellValue);
    return;
  }
}

/**
 * INSTALLABLE TRIGGER FUNCTION
 * Handles "Heavy" logic like logging changes to the Archive.
 */
function trigger_MonitorChanges(e) {
  if (!e) return;
  
  const sheetName = e.range.getSheet().getName();
  const validSheetNamePattern = /^\d{4}T\d{2}$/;
  
  if (!validSheetNamePattern.test(sheetName)) return;

  handlePostAuthEdit(e);
}
