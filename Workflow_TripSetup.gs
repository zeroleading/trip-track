/**
 * WORKFLOW SETUP
 * Handles the creation of new trips and navigation.
 * updated: Integrates Master Template "0000T00", Protection copying, and Reference List logic.
 */

/**
 * Initiates the process of adding a new school trip.
 * This is the function called by the "Add new trip" menu item.
 */
function addTrip() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Get the template sheet and its primary protection.
  const templateSheet = ss.getSheetByName('0000T00');
  if (!templateSheet) {
    ui.alert('Error', 'Template sheet "0000T00" not found. Please ensure it exists.', ui.ButtonSet.OK);
    return;
  }
  
  const templateSheetProtection = templateSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
  if (!templateSheetProtection) {
    ui.alert('Error', 'No sheet protection found on the template sheet "0000T00". Please add protection to the template.', ui.ButtonSet.OK);
    return;
  }

  // 2. Confirm trip has been authorised
  const tripBeenAgreed = ui.alert(
    'Has the trip been approved and entered into the school calendar?',
    ui.ButtonSet.YES_NO
  );

  if(tripBeenAgreed !== ui.Button.YES){
    ui.alert('Please seek approval for your trip from SLT before continuing');
    return;
  } 
  
  // 3. Prompt user for trip name.
  const tripNameUserPrompt = ui.prompt(
    'Trip Name',
    'Enter the name of your new trip below',
    ui.ButtonSet.OK_CANCEL
  );

  if (tripNameUserPrompt.getSelectedButton() !== ui.Button.OK) {
    ui.alert('Operation Cancelled', 'Trip creation was cancelled by the user.', ui.ButtonSet.OK);
    return;
  }

  let newTripName = tripNameUserPrompt.getResponseText().trim();
  if (!newTripName) {
    ui.alert('Input Error', 'Trip name cannot be empty. Please try again.', ui.ButtonSet.OK);
    return; 
  }

  // 4. Generate and add new reference number.
  const generatedTripReference = generateAndAddReferenceNumber();
  if (!generatedTripReference) {
    return; // Error was already handled in the helper
  }

  // 5. Create new sheet
  const newTripSheet = templateSheet.copyTo(ss).setName(generatedTripReference);
  let newTripSheetProtection = newTripSheet.protect();

  // 6. Populate data
  // Note: These Named Ranges must exist inside the Template sheet
  newTripSheet.getRange('thisTrip_tripName').setValue(newTripName);
  newTripSheet.getRange('thisTrip_tripRef').setValue(generatedTripReference);

  newTripSheet.showSheet();
  newTripSheet.activate();

  // 7. Copy protection properties
  newTripSheetProtection.setDescription(templateSheetProtection.getDescription());
  newTripSheetProtection.setWarningOnly(templateSheetProtection.isWarningOnly());

  // Get unprotected ranges from template and map them to the new sheet
  const unprotectedTemplateRangesA1 = templateSheetProtection.getUnprotectedRanges().map(range => range.getA1Notation());
  newTripSheetProtection.setUnprotectedRanges(unprotectedTemplateRangesA1.map(a1 => newTripSheet.getRange(a1)));

  // 8. Final UX
  newTripSheet.getRange(10, 4).activateAsCurrentCell(); // Focus on cell D10

  ui.alert('Trip Created', `New trip "${newTripName}" with reference "${generatedTripReference}" has been successfully created!`, ui.ButtonSet.OK);
  ui.alert('Trip Created','Enter data in all fields marked "*" and add students below before sending for authorisation.', ui.ButtonSet.OK);

  // 9. Logging
  if (typeof logToSystem === 'function') {
    logToSystem("Trip Created", generatedTripReference, `Name: ${newTripName}`);
  }
}


/**
 * Generates a new reference number (e.g., 2511T01) and adds it to the 'referenceList'.
 * @returns {string|null} The newly generated reference number, or null if an error occurred.
 */
function generateAndAddReferenceNumber() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const referenceListNamedRangeName = 'referenceList';

  // Get the current date in YYMM format
  const currentDate = new Date();
  const datePrefixYYMM = Utilities.formatDate(currentDate, 'Europe/London', 'yyMM'); // e.g., "2511"

  let referenceListNamedRange;
  try {
    referenceListNamedRange = ss.getRangeByName(referenceListNamedRangeName);
    if (!referenceListNamedRange) {
      throw new Error(`Named range '${referenceListNamedRangeName}' not found.`);
    }
  } catch (e) {
    Logger.log(e.message + " Please ensure the named range 'referenceList' exists.");
    ui.alert('Error', e.message + "\nPlease ensure the named range 'referenceList' exists.", ui.ButtonSet.OK);
    return null; // Return null to indicate failure.
  }

  // Get all existing references
  const existingReferencesList = referenceListNamedRange.getValues().flat().filter(String);

  // Determine the next increment
  let currentMaxIncrement = 0;
  // Regex: Starts with YYMMT, ends with 2 digits
  const referenceNumberRegex = new RegExp(`^${datePrefixYYMM}T(\\d{2})$`);

  existingReferencesList.forEach(ref => {
    const match = ref.match(referenceNumberRegex);
    if (match) {
      const parsedIncrement = parseInt(match[1], 10);
      if (parsedIncrement > currentMaxIncrement) {
        currentMaxIncrement = parsedIncrement;
      }
    }
  });

  const nextReferenceIncrement = (currentMaxIncrement + 1).toString().padStart(2, '0');
  const newGeneratedReference = `${datePrefixYYMM}T${nextReferenceIncrement}`;
  Logger.log(`Generated new reference: ${newGeneratedReference}`);

  // Add the new reference to the next empty row
  const targetSheetForReference = referenceListNamedRange.getSheet();
  const referenceListColumn = referenceListNamedRange.getColumn();
  
  // Note: This appends to the bottom of the sheet based on data presence
  const nextEmptyRowInReferenceList = targetSheetForReference.getLastRow() + 1; 

  targetSheetForReference.getRange(nextEmptyRowInReferenceList, referenceListColumn).setValue(newGeneratedReference);

  return newGeneratedReference;
}

/**
 * PLACEHOLDER: For the 'View previous trips' menu item.
 */
function viewPrevious() {
  SpreadsheetApp.getUi().alert("This function ('viewPrevious') is not yet implemented.");
}
