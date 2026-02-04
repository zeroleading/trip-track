/**
 * NEW FUNCTION: View Evaluation
 * Reads the 'Form Responses' sheet, finds the entry for the current trip,
 * and displays the results in a popup.
 */
function viewTripEvaluation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  const tripRef = sheet.getName();

  // 1. Basic Validation
  if (!tripRef.match(/^\d{4}T\d{2}$/)) {
    ui.alert("This does not appear to be a valid Trip Sheet.");
    return;
  }

  // 2. Find the Responses Sheet
  const responseSheet = ss.getSheetByName(CONFIG.FORM_RESPONSES_SHEET_NAME);
  if (!responseSheet) {
    ui.alert("Configuration Error", `Could not find sheet named "${CONFIG.FORM_RESPONSES_SHEET_NAME}".`, ui.ButtonSet.OK);
    return;
  }

  // 3. Search for Data
  const data = responseSheet.getDataRange().getValues();
  if (data.length < 2) {
    ui.alert("No evaluations submitted yet.");
    return;
  }

  const headers = data[0];
  const refColIndex = CONFIG.FORM_COL_INDEX_REF;
  
  // Find matching rows (Iterate backwards to find the latest submission first)
  let foundRow = null;
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][refColIndex]) === tripRef) {
      foundRow = data[i];
      break;
    }
  }

  // 4. Display Result
  if (foundRow) {
    let output = `Evaluation for ${tripRef}:\n\n`;
    
    headers.forEach((header, index) => {
      // Skip the Reference column (refColIndex) to avoid redundancy
      if (index !== refColIndex) {
        let answer = foundRow[index];
        if (answer instanceof Date) {
          answer = Utilities.formatDate(answer, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        }
        
        const answerText = String(answer);

        // Formatting: Long answers get a new line for better readability
        if (answerText.length > 40) {
          output += `• ${header}:\n${answerText}\n\n`;
        } else {
          output += `• ${header}: ${answerText}\n`;
        }
      }
    });

    ui.alert("Trip Evaluation", output, ui.ButtonSet.OK);
  } else {
    ui.alert("Not Found", `No evaluation found for trip ${tripRef}.`, ui.ButtonSet.OK);
  }
}
