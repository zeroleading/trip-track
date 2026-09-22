/**
 * WORKFLOW SETUP
 * Handles the creation of new trips and navigation.
 * Updated: Incorporates master switch guard clause, master template protection cloning, and index ordering.
 */

/**
 * Initiates the process of adding a new school trip.
 * Validates system availability and prior approval before cloning the master template sheet.
 * Called directly via the "Add new trip" custom menu trigger.
 * 
 * @returns {void}
 */
function addTrip() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 0. Master Switch Guard Clause
  // Evaluate the master switch before executing any sheet operations or prompting the user.
  // This cleanly halts execution with a polite explanation if submissions are suspended.
  if (!CONFIG.ALLOW_NEW_TRIPS) {
    ui.alert(
      'Trip Creation Suspended',
      CONFIG.TRIP_CREATION_SUSPENDED_MESSAGE || 'Trip creation is currently suspended.',
      ui.ButtonSet.OK
    );
    return;
  }

  // 1. Get the template sheet and its primary protection.
  // Leveraging CONFIG.TEMPLATE_SHEET_NAME ensures alignment with central settings.
  const templateSheetName = CONFIG.TEMPLATE_SHEET_NAME || '0000T00';
  const templateSheet = ss.getSheetByName(templateSheetName);
  if (!templateSheet) {
    ui.alert('Error', `Template sheet "${templateSheetName}" not found. Please ensure it exists.`, ui.ButtonSet.OK);
    return;
  }
  
  const templateSheetProtection = templateSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
  if (!templateSheetProtection) {
    ui.alert('Error', `No sheet protection found on the template sheet "${templateSheetName}". Please add protection to the template.`, ui.ButtonSet.OK);
    return;
  }

  // 2. Confirm trip has been authorised
  const tripBeenAgreed = ui.alert(
    'Has the trip been approved and entered into the school calendar?',
    ui.ButtonSet.YES_NO
  );

  if (tripBeenAgreed !== ui.Button.YES) {
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
    return; // Error handling already managed inside generateAndAddReferenceNumber
  }

  // 5. Create new sheet
  const newTripSheet = templateSheet.copyTo(ss).setName(generatedTripReference);

  // Position new sheet directly adjacent to the primary Menu dashboard
  const menuSheet = ss.getSheetByName(CONFIG.MENU_SHEET_NAME);
  if (menuSheet) {
    newTripSheet.activate();
    ss.moveActiveSheet(menuSheet.getIndex() + 1);
  }

  let newTripSheetProtection = newTripSheet.protect();

  // 6. Populate data into template named ranges
  newTripSheet.getRange('thisTrip_tripName').setValue(newTripName);
  newTripSheet.getRange('thisTrip_tripRef').setValue(generatedTripReference);

  newTripSheet.showSheet();
  newTripSheet.activate();

  // 7. Replicate protection properties
  newTripSheetProtection.setDescription(templateSheetProtection.getDescription());
  newTripSheetProtection.setWarningOnly(templateSheetProtection.isWarningOnly());

  // Map unprotected ranges from template onto the newly created sheet
  const unprotectedTemplateRangesA1 = templateSheetProtection.getUnprotectedRanges().map(range => range.getA1Notation());
  newTripSheetProtection.setUnprotectedRanges(unprotectedTemplateRangesA1.map(a1 => newTripSheet.getRange(a1)));

  // 8. Final User Experience
  newTripSheet.getRange(10, 4).activateAsCurrentCell(); // Focus on cell D10

  ui.alert('Trip Created', `New trip "${newTripName}" with reference "${generatedTripReference}" has been successfully created!`, ui.ButtonSet.OK);
  ui.alert('Trip Created', 'Enter data in all fields marked "*" and add students below before sending for authorisation.', ui.ButtonSet.OK);

  // 9. Audit Logging
  if (typeof logToSystem === 'function') {
    logToSystem("Trip Created", generatedTripReference, `Name: ${newTripName}`);
  }
}

/**
 * Generates a new reference number (e.g., 2511T01) and appends it to the 'referenceList'.
 * Inspects existing entries matching the YYMMT pattern to allocate the next sequential increment.
 * 
 * @returns {string|null} The newly generated reference number, or null if an error occurred.
 */
function generateAndAddReferenceNumber() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const referenceListNamedRangeName = CONFIG.REFERENCE_LIST_RANGE_NAME || 'referenceList';

  // Format current date in YYMM format within the London timezone
  const currentDate = new Date();
  const datePrefixYYMM = Utilities.formatDate(currentDate, 'Europe/London', 'yyMM'); // e.g., "2511"

  let referenceListNamedRange;
  try {
    referenceListNamedRange = ss.getRangeByName(referenceListNamedRangeName);
    if (!referenceListNamedRange) {
      throw new Error(`Named range '${referenceListNamedRangeName}' not found.`);
    }
  } catch (e) {
    Logger.log(e.message + " Please ensure the named range exists.");
    ui.alert('Error', e.message + "\nPlease ensure the named range exists.", ui.ButtonSet.OK);
    return null;
  }

  // Retrieve existing references
  const existingReferencesList = referenceListNamedRange.getValues().flat().filter(String);

  // Determine highest existing increment matching current month prefix
  let currentMaxIncrement = 0;
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

  // Append new reference to next vacant row in designated range sheet
  const targetSheetForReference = referenceListNamedRange.getSheet();
  const referenceListColumn = referenceListNamedRange.getColumn();
  const nextEmptyRowInReferenceList = targetSheetForReference.getLastRow() + 1; 

  targetSheetForReference.getRange(nextEmptyRowInReferenceList, referenceListColumn).setValue(newGeneratedReference);

  return newGeneratedReference;
}