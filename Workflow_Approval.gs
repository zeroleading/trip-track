/**
 * WORKFLOW APPROVAL
 * Handles the Authorisation Request, Approval (Archive creation), and Denial processes.
 * Integrated with Email Notifications and Permission Checks.
 */

/**
 * 1. TRIP LEADER: Request Authorisation
 * Called from the 'Trip Admin' menu.
 */
function requestAuthorisation() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  // 1. Get Trip Data
  const tripData = getSummaryData(sheet);
  if (!tripData) {
    Logger.log("Workflow stopped because tripData was null.");
    return; 
  }
  
  // 2. CHECK MINIMUM REQUIREMENTS
  // Checks the cell mapped to 'trip_minimumRequirements' in the summary table
  const minReq = tripData.trip_minimumRequirements;
  const isMinReqMet = (minReq === true || String(minReq).toUpperCase() === 'TRUE');

  if (!isMinReqMet) {
    ui.alert('Cannot Request Authorisation', 'Enter data in all fields marked "*" and/or at least one student before requesting authorisation.', ui.ButtonSet.OK);
    return; 
  }

  // 3. Status & Logic Checks
  const status = tripData.trip_status;
  const sltEmail = sheet.getRange(CONFIG.SLT_SELECTED_RANGE_NAME).getValue();
  const tripName = tripData.trip_name;
  const tripRef = tripData.trip_ref;

  if (status !== 'Incomplete' && status !== 'Denied') {
    ui.alert('Action Stopped', `This trip's status is "${status}". A request can only be sent if the status is "Incomplete" or "Denied".`, ui.ButtonSet.OK);
    return;
  }
  
  if (!sltEmail || !sltEmail.includes('@')) {
    ui.alert('Action Stopped', 'Please select an SLT Authoriser from the dropdown (range "authorisers_singleSelected") before requesting.', ui.ButtonSet.OK);
    return;
  }

  // 4. Confirmation
  const response = ui.alert('Confirm Request', `This will send an authorisation request to ${sltEmail}, change the status to "Pending authorisation", and sort the sheet. Proceed?`, ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;
  
  try {
    // 5. Update Status
    updateSummaryValue(sheet, 'trip_status', 'Pending authorisation');
    
    // Sort students if the function exists (in Workflow_StudentTools.gs)
    if (typeof sortNames === 'function') sortNames();
    
    // 6. Prepare & Send Email
    const tripLeaderEmail = tripData.trip_leadEmail;
    
    const tripDateStr = String(tripData.trip_date || "Date Not Set");
    const daysToGo = String(tripData.trip_countdown || "N/A");
    
    let timeStr = String(tripData.trip_timeLeaving);
    if (tripData.trip_timeLeaving instanceof Date) {
      timeStr = Utilities.formatDate(tripData.trip_timeLeaving, Session.getScriptTimeZone(), "HH:mm");
    }

    const subject = `ACTION REQUIRED: Trip Authorisation Request - ${tripName} (${tripRef})`;
    const body = `
      <p>Hello,</p>
      <p>A request for authorisation has been submitted for the following trip:</p>
      
      <table style="border-collapse: collapse; width: 100%; max-width: 600px;">
        <tr><td style="padding: 5px; font-weight: bold;">Trip Name:</td><td style="padding: 5px;">${tripName}</td></tr>
        <tr><td style="padding: 5px; font-weight: bold;">Reference:</td><td style="padding: 5px;">${tripRef}</td></tr>
        <tr><td style="padding: 5px; font-weight: bold;">Leader:</td><td style="padding: 5px;">${tripData.trip_leadName}</td></tr>
        <tr><td style="padding: 5px; font-weight: bold;">Date:</td><td style="padding: 5px;">${tripDateStr} (Leaves: ${timeStr})</td></tr>
        <tr><td style="padding: 5px; font-weight: bold;">Countdown:</td><td style="padding: 5px;">${daysToGo}</td></tr>
      </table>

      <p>Please review the details in the spreadsheet and "Approve" or "Deny" the trip using the "Trip Admin" menu.</p>
      
      <p><b><a href="${ss.getUrl()}#gid=${sheet.getSheetId()}" style="font-size: 16px;">Open Trip Spreadsheet</a></b></p>
    `;
    
    MailApp.sendEmail({ to: sltEmail, cc: tripLeaderEmail, subject: subject, htmlBody: body });
    ui.alert('Request Sent', `The request has been sent to ${sltEmail}, the status is "Pending authorisation", and the sheet has been sorted.`, ui.ButtonSet.OK);

    // LOGGING
    logToSystem("Request Sent", tripRef, `Sent to: ${sltEmail}`);

  } catch (e) {
    Logger.log(e);
    ui.alert('An Error Occurred', `The process failed. Error: ${e.message}`, ui.ButtonSet.OK);
    // Revert status on error
    updateSummaryValue(sheet, 'trip_status', status);
  }
}

/**
 * 2. SLT: Approve Trip
 * Called from the 'Trip Admin' menu.
 */
function approveTripWorkflow() {
  const ui = SpreadsheetApp.getUi();
  const activeUserEmail = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  // 1. Permission Check
  const sltEmailList = getSltEmailList();
  if (!sltEmailList.includes(activeUserEmail)) {
    ui.alert('Permission Denied', 'You are not authorized to approve trips.', ui.ButtonSet.OK);
    return;
  }
  
  const tripData = getSummaryData(sheet);
  if (!tripData) {
    Logger.log("Workflow stopped because tripData was null.");
    return; 
  }
  
  // 2. Status Check
  const status = tripData.trip_status;
  if (status !== 'Pending authorisation') {
    ui.alert('Action Stopped', `This trip's status is "${status}". It must be "Pending authorisation" to be approved.`, ui.ButtonSet.OK);
    return;
  }
  
  // 3. Confirmation
  const response = ui.alert('Confirm Approval', `Are you sure you want to approve trip "${tripData.trip_ref}"? This will create the values-only archive sheet and lock the status.`, ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  try {
    SpreadsheetApp.flush(); 
    
    const authDate = new Date();
    const authDateStr = Utilities.formatDate(authDate, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    
    // Update local data object for email usage
    tripData['trip_status'] = 'Authorised';
    tripData['trip_authorisedBy'] = activeUserEmail;
    tripData['trip_authorisationDate'] = authDateStr;

    // 4. Create Archive (Using Template Logic)
    // REPLACED: createArchiveSheet(sheet, tripData) -> createArchiveFromTemplate
    const archiveUrl = createArchiveFromTemplate(sheet, tripData.trip_ref, tripData.trip_name);
    
    // 5. Update Sheet
    updateSummaryValue(sheet, 'trip_status', 'Authorised');
    updateSummaryValue(sheet, 'trip_authorisedBy', activeUserEmail);
    updateSummaryValue(sheet, 'trip_authorisationDate', authDateStr);

    // 6. Send Notifications
    const tripDateStr = String(tripData.trip_date || "Date Not Set");
    const daysToGo = String(tripData.trip_countdown || "N/A");
    
    let timeStr = String(tripData.trip_timeLeaving);
    if (tripData.trip_timeLeaving instanceof Date) {
      timeStr = Utilities.formatDate(tripData.trip_timeLeaving, Session.getScriptTimeZone(), "HH:mm");
    }

    const subject = `Trip Approved: ${tripData.trip_name} (${tripData.trip_ref})`;
    const body = `
      <p>The following trip has been approved by ${activeUserEmail}.</p>
      
      <table style="border-collapse: collapse; width: 100%; max-width: 600px; background-color: #f9f9f9; padding: 10px; border: 1px solid #ddd;">
        <tr><td style="padding: 5px; font-weight: bold;">Trip Name:</td><td style="padding: 5px;">${tripData.trip_name}</td></tr>
        <tr><td style="padding: 5px; font-weight: bold;">Reference:</td><td style="padding: 5px;">${tripData.trip_ref}</td></tr>
        <tr><td style="padding: 5px; font-weight: bold;">Date:</td><td style="padding: 5px;">${tripDateStr} (Leaves: ${timeStr})</td></tr>
        <tr><td style="padding: 5px; font-weight: bold;">Countdown:</td><td style="padding: 5px;">${daysToGo}</td></tr>
      </table>

      <hr>

      <h3>1. Live Trip Management</h3>
      <p>Use this link to edit student details, risk assessments, or trip information. Any changes made here will be logged.</p>
      <p><b><a href="${ss.getUrl()}#gid=${sheet.getSheetId()}" style="font-size: 16px; color: #1a73e8;">Open Live Spreadsheet</a></b></p>

      <br>

      <h3>2. Archive Record (Snapshot)</h3>
      <p>This is a permanent record of the trip details at the moment of authorisation. This file cannot be edited.</p>
      <p><b><a href="${archiveUrl}" style="font-size: 16px; color: #5f6368;">View Archive Snapshot</a></b></p>
    `;
    
    MailApp.sendEmail({
      to: tripData.trip_leadEmail,
      cc: activeUserEmail,
      bcc: 'jappleton@csg.school',
      subject: subject,
      htmlBody: body
    });

    ui.alert('Approval Complete', 'The trip has been approved. The archive snapshot has been created, and the trip leader has been notified.', ui.ButtonSet.OK);

    // LOGGING
    logToSystem("Trip Approved", tripData.trip_ref, `Authorised by: ${activeUserEmail}`);

  } catch (e) {
    Logger.log(e);
    ui.alert('An Error Occurred', `The process failed. Please contact support. Error: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * 3. SLT: Deny Trip
 * Called from the 'Trip Admin' menu.
 */
function denyTripWorkflow() {
  const ui = SpreadsheetApp.getUi();
  const activeUserEmail = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  // 1. Permission Check
  const sltEmailList = getSltEmailList();
  if (!sltEmailList.includes(activeUserEmail)) {
    ui.alert('Permission Denied', 'You are not authorized to deny trips.', ui.ButtonSet.OK);
    return;
  }

  const tripData = getSummaryData(sheet);
  if (!tripData) {
    Logger.log("Workflow stopped because tripData was null.");
    return;
  }
  
  // 2. Status Check
  const status = tripData.trip_status;
  if (status !== 'Pending authorisation') {
    ui.alert('Action Stopped', `This trip's status is "${status}". It must be "Pending authorisation" to be denied.`, ui.ButtonSet.OK);
    return;
  }
  
  // 3. Reason Prompt
  const reasonResponse = ui.prompt('Reason for Denial', 'Please provide a brief reason for denying this trip (this will be emailed to the trip leader).', ui.ButtonSet.OK_CANCEL);
  if (reasonResponse.getSelectedButton() !== ui.Button.OK) return;
  const reasonText = reasonResponse.getResponseText();
  if (!reasonText) {
    ui.alert('Reason Required', 'You must provide a reason for the denial.', ui.ButtonSet.OK);
    return;
  }

  try {
    // 4. Update Status
    updateSummaryValue(sheet, 'trip_status', 'Denied');

    // 5. Send Notification
    const subject = `Trip Denied: ${tripData.trip_name} (${tripData.trip_ref})`;
    const body = `
      <p>Hello,</p>
      <p>Your request for authorisation for the trip <b>${tripData.trip_name} (${tripData.trip_ref})</b> has been denied by ${activeUserEmail}.</p>
      <p><b>Reason:</b><br><i>${reasonText}</i></p>
      <p>Please make the required changes and re-submit for authorisation using the "Trip Admin" menu.</p>
      <p><b>Spreadsheet Link:</b> <a href="${ss.getUrl()}#gid=${sheet.getSheetId()}">${ss.getName()}</a></p>
    `;
    
    MailApp.sendEmail({
      to: tripData.trip_leadEmail,
      cc: activeUserEmail,
      subject: subject,
      htmlBody: body
    });

    ui.alert('Trip Denied', 'The trip status has been set to "Denied" and the trip leader has been notified.', ui.ButtonSet.OK);

    // LOGGING
    logToSystem("Trip Denied", tripData.trip_ref, `Reason: ${reasonText}`);

  } catch (e) {
    Logger.log(e);
    ui.alert('An Error Occurred', `The process failed. Error: ${e.message}`, ui.ButtonSet.OK);
  }
}
