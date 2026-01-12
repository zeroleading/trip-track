/**
 * WORKFLOW STUDENT TOOLS
 * Handles student list management (Add/Remove) and Sorting.
 */

/**
 * Adds students from the main list (Column S) to the 'added students' list (Column D).
 */
function addTickedStudents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();

  const dataStartRow = 35;
  const masterAddStudentsTickCellA1 = 'Q33'; 

  const sheetLastRow = activeSheet.getLastRow();
  const processedRowCount = Math.max(0, sheetLastRow - dataStartRow + 1);

  if (processedRowCount === 0) {
    ss.toast('No data found.', 'No Action', 5);
    return;
  }

  const mainStudentListRange = activeSheet.getRange(dataStartRow, 19, processedRowCount, 1); 
  const addStudentsTickBoxColumnRange = activeSheet.getRange(dataStartRow, 17, processedRowCount, 1); 
  const currentAddedStudentsColumnRange = activeSheet.getRange('D35:D');

  const mainStudentNames = mainStudentListRange.getValues();
  const addTickBoxStates = addStudentsTickBoxColumnRange.getValues();
  const existingAddedStudentNames = currentAddedStudentsColumnRange.getDisplayValues(); 

  const selectedStudentsToAdd = [];
  const updatedAddTickBoxStates = [];

  for (let i = 0; i < processedRowCount; i++) {
    const currentStudentName = mainStudentNames[i][0];
    const isAddTickBoxChecked = addTickBoxStates[i][0];

    if (isAddTickBoxChecked === true && currentStudentName !== '' && currentStudentName !== null) {
      selectedStudentsToAdd.push([currentStudentName]);
      updatedAddTickBoxStates.push([false]); 
    } else {
      updatedAddTickBoxStates.push([addTickBoxStates[i][0]]); 
    }
  }

  if (selectedStudentsToAdd.length > 0) {
    const lastRowOfAddedStudents = existingAddedStudentNames.filter(String).length + 34;
    const rangeForAddingStudents = activeSheet.getRange(lastRowOfAddedStudents + 1, 4, selectedStudentsToAdd.length, 1); 
    
    rangeForAddingStudents.setValues(selectedStudentsToAdd);
    ss.toast('Selected students have been added to the list.', 'Students Added!', 5);

    addStudentsTickBoxColumnRange.setValues(updatedAddTickBoxStates);
    activeSheet.getRange(masterAddStudentsTickCellA1).setValue('FALSE'); 

    // Log this system action
    const namesAdded = selectedStudentsToAdd.map(row => row[0]).join(", ");
    if (typeof logSystemAction === 'function') {
      logSystemAction(activeSheet, `Added students: ${namesAdded}`);
    }

  } else {
    ss.toast('Please tick the boxes next to the students you wish to add.', 'No Students Selected', 5);
  }
}

/**
 * Removes students from the 'added students' list (Column D).
 */
function removeTickedStudents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();

  const dataStartRow = 35; 
  const masterRemoveStudentsTickCellA1 = 'B33'; 

  const sheetLastRow = activeSheet.getLastRow();
  const processedRowCount = Math.max(0, sheetLastRow - dataStartRow + 1);

  if (processedRowCount === 0) {
    ss.toast('No data found.', 'No Action', 5);
    return;
  }

  const removeStudentsTickBoxColumnRange = activeSheet.getRange(dataStartRow, 2, processedRowCount, 1); 
  const studentsInAddedListRange = activeSheet.getRange(dataStartRow, 4, processedRowCount, 1); 

  const removeTickBoxStates = removeStudentsTickBoxColumnRange.getValues();
  const currentAddedStudentNames = studentsInAddedListRange.getValues();

  const filteredStudentsToKeep = [];
  const studentsRemoved = []; // Track who we are removing
  const updatedRemoveTickBoxStates = []; 

  for (let i = 0; i < processedRowCount; i++) {
    const currentStudentName = currentAddedStudentNames[i][0];
    const isRemoveTickBoxChecked = removeTickBoxStates[i][0];

    if (isRemoveTickBoxChecked !== true && currentStudentName !== '' && currentStudentName !== null) {
      filteredStudentsToKeep.push([currentStudentName]);
      updatedRemoveTickBoxStates.push([false]);
    } else {
      if (currentStudentName) studentsRemoved.push(currentStudentName);
      updatedRemoveTickBoxStates.push([false]);
    }
  }

  const entireAddedStudentsListColumn = activeSheet.getRange(dataStartRow, 4, activeSheet.getMaxRows() - dataStartRow + 1, 1);
  entireAddedStudentsListColumn.clearContent();
  activeSheet.getRange(masterRemoveStudentsTickCellA1).setValue('FALSE');

  // Clear tick boxes first
  removeStudentsTickBoxColumnRange.setValues(updatedRemoveTickBoxStates);

  if (studentsRemoved.length > 0) {
    // Repopulate list
    if (filteredStudentsToKeep.length > 0) {
      const rangeForRepopulatingStudents = activeSheet.getRange(dataStartRow, 4, filteredStudentsToKeep.length, 1);
      rangeForRepopulatingStudents.setValues(filteredStudentsToKeep);
    }
    ss.toast('Selected students have been removed and the list repopulated.', 'Students Removed!', 5);

    // Log this system action
    const namesRemoved = studentsRemoved.join(", ");
    if (typeof logSystemAction === 'function') {
      logSystemAction(activeSheet, `Removed students: ${namesRemoved}`);
    }

  } else {
    ss.toast('No students were selected for removal.', 'No Action', 5);
  }
}

/**
 * Sorts the student name column (D) on the active sheet.
 * Updated Logic:
 * 1. Checks if student count > 50.
 * 2. Checks if majority of Reg Groups are 2-3 chars (e.g. "7C", "10A").
 * 3. If both true: Sorts by Reg Group (Numeric) -> Name.
 * 4. Else: Standard A-Z Name sort.
 */
function sortNames() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const startRow = 35;
  const colIndex = 4; // Column D

  // Get dynamic range of all data in Column D starting at 35
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;

  const range = sheet.getRange(startRow, colIndex, lastRow - startRow + 1, 1);
  const values = range.getValues();
  
  // Filter out empty rows to get actual student list
  const studentList = values.flat().filter(s => typeof s === 'string' && s.trim() !== "");
  const count = studentList.length;

  // Decision Flag
  let useRegGroupSort = false;

  // CRITERIA 1: More than 50 students
  if (count > 50) {
    let validRegCount = 0;
    let targetLengthCount = 0; // Length 2 or 3

    studentList.forEach(name => {
      // Expected Format: "Surname, Forename - Reg"
      const parts = name.split(" - ");
      if (parts.length > 1) {
        validRegCount++;
        // Get the last part as the Reg Group (trims whitespace)
        const reg = parts[parts.length - 1].trim();
        // UPDATED: Check for 2 or 3 characters
        if (reg.length === 2 || reg.length === 3) {
          targetLengthCount++;
        }
      }
    });

    // CRITERIA 2: Majority (>50%) are 2 or 3 chars
    if (validRegCount > 0 && (targetLengthCount / validRegCount > 0.5)) {
      useRegGroupSort = true;
    }
  }

  // EXECUTE SORT
  if (useRegGroupSort) {
    // Custom Sort: Reg Group first, then Name
    studentList.sort((a, b) => {
      const regA = a.split(" - ").pop().trim();
      const regB = b.split(" - ").pop().trim();

      // Compare Reg Groups with numeric sensitivity (so "7C" comes before "10C")
      const result = regA.localeCompare(regB, undefined, {numeric: true, sensitivity: 'base'});
      
      // If Reg Groups are same, fallback to full string (Name)
      if (result === 0) {
        return a.localeCompare(b);
      }
      return result;
    });

    // Apply back to sheet:
    // 1. Clear original range to prevent ghosting
    range.clearContent();
    // 2. Write sorted list
    if (studentList.length > 0) {
      const output = studentList.map(s => [s]);
      sheet.getRange(startRow, colIndex, output.length, 1).setValues(output);
    }
    
    if (typeof logSystemAction === 'function') {
      logSystemAction(sheet, "Student list sorted (Reg Group > Name).");
    }

  } else {
    // Standard Sort (A-Z)
    range.sort(colIndex);
    
    if (typeof logSystemAction === 'function') {
      logSystemAction(sheet, "Student list sorted (A-Z).");
    }
  }
}
