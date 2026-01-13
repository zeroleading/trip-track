/**
 * MESSAGE TESTING SUITE
 * Allows the current user to trigger email notifications to THEMSELVES only.
 * RESTRICTED: Only runs for System Admins or SLT members.
 */

// =============================================================================
// MASTER TEST
// =============================================================================

function test_RunAllEmails() {
  const ui = SpreadsheetApp.getUi();
  
  // Security Check
  if (!isUserAllowedToTest()) {
    ui.alert("Access Denied", "You do not have permission to run system tests.\n(User not in SYSTEM_TESTERS or SLT list)", ui.ButtonSet.OK);
    return;
  }

  const { sheet } = getTestData();
  if (!sheet) return; 

  const confirm = ui.alert(
    "Run All Tests?", 
    "This will send 8 separate emails to your inbox representing the full trip lifecycle.\n\nProceed?", 
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  test_RequestAuth(); Utilities.sleep(500);
  test_ApproveTrip(); Utilities.sleep(500);
  test_DenyTrip();    Utilities.sleep(500);
  test_T7();          Utilities.sleep(500);
  test_T4();          Utilities.sleep(500);
  test_T1();          Utilities.sleep(500);
  test_T0();          Utilities.sleep(500);
  test_EvaluationEmail();

  ui.alert("All 8 test emails have been queued/sent.");
}

// =============================================================================
// APPROVAL WORKFLOW TESTS
// =============================================================================

function test_RequestAuth() {
  const { sheet, tripData, userEmail, ssUrl } = getTestData();
  if (!sheet) return;

  const subject = `[TEST] ACTION REQUIRED: Trip Authorisation Request - ${tripData.trip_name} (${tripData.trip_ref})`;
  
  const tripDateStr = String(tripData.trip_date || "Date Not Set");
  const daysToGo = String(tripData.trip_countdown || "N/A");
  let timeStr = String(tripData.trip_timeLeaving);
  if (tripData.trip_timeLeaving instanceof Date) {
    timeStr = Utilities.formatDate(tripData.trip_timeLeaving, Session.getScriptTimeZone(), "HH:mm");
  }

  const emailBody = `
    <p>Hello,</p>
    <p>A request for authorisation has been submitted for the following trip:</p>
    
    <table style="border-collapse: collapse; width: 100%; max-width: 600px;">
      <tr><td style="padding: 5px; font-weight: bold;">Trip Name:</td><td style="padding: 5px;">${tripData.trip_name}</td></tr>
      <tr><td style="padding: 5px; font-weight: bold;">Reference:</td><td style="padding: 5px;">${tripData.trip_ref}</td></tr>
      <tr><td style="padding: 5px; font-weight: bold;">Leader:</td><td style="padding: 5px;">${tripData.trip_leadName}</td></tr>
      <tr><td style="padding: 5px; font-weight: bold;">Date:</td><td style="padding: 5px;">${tripDateStr} (Leaves: ${timeStr})</td></tr>
      <tr><td style="padding: 5px; font-weight: bold;">Countdown:</td><td style="padding: 5px;">${daysToGo}</td></tr>
    </table>

    <p>Please review the details in the spreadsheet and "Approve" or "Deny" the trip using the "Trip Admin" menu.</p>
    
    <p><b><a href="${ssUrl}" style="font-size: 16px;">Open Trip Spreadsheet</a></b></p>
  `;

  sendTestEmail(userEmail, "SLT Authoriser, Trip Leader", subject, emailBody);
}

function test_ApproveTrip() {
  const { sheet, tripData, userEmail, ssUrl } = getTestData();
  if (!sheet) return;

  const subject = `[TEST] Trip Approved: ${tripData.trip_name} (${tripData.trip_ref})`;
  const archiveUrl = "https://docs.google.com/spreadsheets/d/placeholder"; 
  
  const tripDateStr = String(tripData.trip_date || "Date Not Set");
  const daysToGo = String(tripData.trip_countdown || "N/A");
  let timeStr = String(tripData.trip_timeLeaving);
  if (tripData.trip_timeLeaving instanceof Date) {
    timeStr = Utilities.formatDate(tripData.trip_timeLeaving, Session.getScriptTimeZone(), "HH:mm");
  }

  const emailBody = `
    <p>The following trip has been approved by ${userEmail} (Test User).</p>
    
    <table style="border-collapse: collapse; width: 100%; max-width: 600px; background-color: #f9f9f9; padding: 10px; border: 1px solid #ddd;">
      <tr><td style="padding: 5px; font-weight: bold;">Trip Name:</td><td style="padding: 5px;">${tripData.trip_name}</td></tr>
      <tr><td style="padding: 5px; font-weight: bold;">Reference:</td><td style="padding: 5px;">${tripData.trip_ref}</td></tr>
      <tr><td style="padding: 5px; font-weight: bold;">Date:</td><td style="padding: 5px;">${tripDateStr} (Leaves: ${timeStr})</td></tr>
      <tr><td style="padding: 5px; font-weight: bold;">Countdown:</td><td style="padding: 5px;">${daysToGo}</td></tr>
    </table>

    <hr>

    <h3>1. Live Trip Management</h3>
    <p>Use this link to edit student details, risk assessments, or trip information. Any changes made here will be logged.</p>
    <p><b><a href="${ssUrl}" style="font-size: 16px; color: #1a73e8;">Open Live Spreadsheet</a></b></p>

    <br>

    <h3>2. Archive Record (Snapshot)</h3>
    <p>This is a permanent record of the trip details at the moment of authorisation. This file cannot be edited.</p>
    <p><b><a href="${archiveUrl}" style="font-size: 16px; color: #5f6368;">View Archive Snapshot</a></b></p>
  `;

  sendTestEmail(userEmail, "Trip Leader, SLT, Cover", subject, emailBody);
}

function test_DenyTrip() {
  const { sheet, tripData, userEmail, ssUrl } = getTestData();
  if (!sheet) return;

  const subject = `[TEST] Trip Denied: ${tripData.trip_name} (${tripData.trip_ref})`;
  const reasonText = "Insufficient staffing ratio (Test Reason)";

  const emailBody = `
    <p>Hello,</p>
    <p>Your request for authorisation for the trip <b>${tripData.trip_name} (${tripData.trip_ref})</b> has been denied by ${userEmail}.</p>
    <p><b>Reason:</b><br><i>${reasonText}</i></p>
    <p>Please make the required changes and re-submit for authorisation using the "Trip Admin" menu.</p>
    <p><b>Spreadsheet Link:</b> <a href="${ssUrl}">${sheet.getParent().getName()}</a></p>
  `;

  sendTestEmail(userEmail, "Trip Leader, SLT", subject, emailBody);
}

// =============================================================================
// TIME-BASED WORKFLOW TESTS
// =============================================================================

function test_T7() {
  const { sheet, tripData, userEmail, ssUrl } = getTestData();
  if (!sheet) return;

  const subject = `[TEST] ACTION REQUIRED: 7 Days to Go - ${tripData.trip_name} (${tripData.trip_ref})`;
  
  const emailBody = `
    <h3>Trip Lock-In Required</h3>
    <p>Your trip <b>${tripData.trip_name}</b> departs in one week.</p>
    <p>Please log in immediately and:</p>
    <ul>
      <li>Finalise the student list (tick/untick students).</li>
      <li>Ensure all medical notes are read and risk assessment checked.</li>
      <li>Confirm staffing details.</li>
    </ul>

    <hr>
    <h4>Leader Checklist</h4>
    <ul>
      <li>Parents/carers informed/reminded?</li>
      <li>School mobile acquired?</li>
      <li>Transport and venue confirmed?</li>
    </ul>
    <hr>

    <p><b>Note:</b> Operational lists will be sent to Attendance and Cover in 3 days (T-4). Changes after that point require manual emails to officers.</p>
    <p><a href="${ssUrl}">Open Trip Sheet</a></p>
  `;

  sendTestEmail(userEmail, "Trip Leader, SLT Contact", subject, emailBody);
}

function test_T4() {
  const { sheet, tripData, userEmail, ssUrl } = getTestData();
  if (!sheet) return;

  const tripDateStr = tripData.trip_date instanceof Date ? 
    Utilities.formatDate(tripData.trip_date, Session.getScriptTimeZone(), "dd/MM/yyyy") : String(tripData.trip_date);

  let leaveTime = String(tripData.trip_timeLeaving || "TBC");
  if (tripData.trip_timeLeaving instanceof Date) {
    leaveTime = Utilities.formatDate(tripData.trip_timeLeaving, Session.getScriptTimeZone(), "HH:mm");
  }
  
  let returnTime = String(tripData.trip_timeReturning || "TBC");
  if (tripData.trip_timeReturning instanceof Date) {
    returnTime = Utilities.formatDate(tripData.trip_timeReturning, Session.getScriptTimeZone(), "HH:mm");
  }

  const leadName = tripData.trip_leadName || "Unknown Lead";
  const staffRaw = tripData.trip_staffNames || "";
  
  let staffHtml = `<ul><li><b>Lead:</b> ${leadName}</li>`;
  if (staffRaw) {
    const staffArray = staffRaw.split(',').map(s => s.trim()).filter(s => s);
    staffArray.forEach(s => staffHtml += `<li>${s}</li>`);
  } else {
    staffHtml += `<li>(No other staff listed)</li>`;
  }
  staffHtml += `</ul>`;

  // Attendance
  const studentListHtml = getTestStudentListHtml(sheet);
  const attSubject = `[TEST] TRIP ATTENDANCE: ${tripData.trip_name} (${tripData.trip_ref})`;
  const attBody = `
    <h3>Trip Departing in 4 Days</h3>
    <p><b>Ref:</b> ${tripData.trip_ref}<br>
    <b>Date:</b> ${tripDateStr}<br>
    <b>Times:</b> ${leaveTime} - ${returnTime}<br>
    <b>Leader:</b> ${tripData.trip_leadName}</p>
    
    <h4>Student List for MIS Coding:</h4>
    ${studentListHtml}
    
    <p><i>Please code these students as 'V' (Educational Visit).</i></p>
  `;
  sendTestEmail(userEmail, "Attendance Officer", attSubject, attBody);

  // Cover
  const covSubject = `[TEST] TRIP COVER: ${tripData.trip_name} (${tripData.trip_ref})`;
  const covBody = `
    <h3>Trip Departing in 4 Days</h3>
    <p><b>Date:</b> ${tripDateStr}</p>
    <p><b>Times:</b> ${leaveTime} - ${returnTime}</p>
    
    <h4>Staff Requiring Cover:</h4>
    ${staffHtml}
    
    <p>Please check the spreadsheet below for details.</p>
    <p><a href="${ssUrl}">Open Trip Sheet</a></p>
  `;
  sendTestEmail(userEmail, "Cover Supervisor", covSubject, covBody);
}

function test_T1() {
  const { sheet, tripData, userEmail } = getTestData();
  if (!sheet) return;

  const subject = `[TEST] PRINT NOW: Leader Pack for ${tripData.trip_name}`;
  const body = `
    <h3>T-1 Leader Pack</h3>
    <p>Your trip departs tomorrow.</p>
    <p><b>Please find attached the Leader Pack (PDF).</b></p>
    <p>This document contains:</p>
    <ul>
      <li>Registration List (with Context)</li>
      <li>Emergency Contacts</li>
      <li>Medical & Dietary Red Flags</li>
      <li>Medical Notes Summary</li>
      <li>Risk Assessment Narrative</li>
    </ul>
  `;

  let pdfBlob = null;
  try {
    if (typeof generateLeaderPackPDF === 'function') {
      pdfBlob = generateLeaderPackPDF(sheet, tripData);
    }
  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast("PDF Generation Failed: " + e.message);
  }

  const emailOptions = {
    to: userEmail,
    subject: subject,
    htmlBody: wrapTestBody(userEmail, "Trip Leader", body)
  };

  if (pdfBlob) emailOptions.attachments = [pdfBlob];

  MailApp.sendEmail(emailOptions);
  SpreadsheetApp.getActiveSpreadsheet().toast("T-1 Test Email (with PDF) sent.");
}

function test_T0() {
  const { sheet, tripData, userEmail, ssUrl } = getTestData();
  if (!sheet) return;

  let leaveTime = String(tripData.trip_timeLeaving || "TBC");
  if (tripData.trip_timeLeaving instanceof Date) {
    leaveTime = Utilities.formatDate(tripData.trip_timeLeaving, Session.getScriptTimeZone(), "HH:mm");
  }
  
  let returnTime = String(tripData.trip_timeReturning || "TBC");
  if (tripData.trip_timeReturning instanceof Date) {
    returnTime = Utilities.formatDate(tripData.trip_timeReturning, Session.getScriptTimeZone(), "HH:mm");
  }

  const subject = `[TEST] REMINDER: Trip Departing TODAY - ${tripData.trip_name}`;
  const emailBody = `
    <p>This is a reminder that <b>${tripData.trip_ref}</b> departs today.</p>
    <p><b>Times:</b> ${leaveTime} - ${returnTime}</p>
    <p>Please ensure the register is marked.</p>
    <p style="color: red; font-weight: bold;">⚠️ NOTE: The student list may have changed since the last notification (T-4). Please check the sheet below for the final list.</p>
    <p><a href="${ssUrl}">Open Trip Sheet</a></p>
  `;

  sendTestEmail(userEmail, "Attendance Officer", subject, emailBody);
}

/**
 * TEST: T+1 Evaluation Email
 * Simulates the email sent to the Trip Leader the day after the trip.
 * Uses the configuration from Config.gs to generate the pre-filled link.
 */
function test_EvaluationEmail() {
  // Dependencies from Testing.gs
  if (typeof getTestData !== 'function') {
    SpreadsheetApp.getUi().alert("Error: 'Testing.gs' functions (getTestData) not found.");
    return;
  }

  const { sheet, tripData, userEmail } = getTestData();
  if (!sheet) return;

  // 1. Construct Link
  const formUrl = CONFIG.FORM_EVALUATION_URL;
  const entryId = CONFIG.FORM_ENTRY_ID_TRIP_REF;
  const preFilledUrl = `${formUrl}?usp=pp_url&${entryId}=${tripData.trip_ref}`;

  // 2. Prepare Email Content
  const subject = `[TEST] Trip Evaluation Required: ${tripData.trip_name} (${tripData.trip_ref})`;
  
  const body = `
    <h3>T+1 Evaluation Request</h3>
    <p>This is a test of the email sent to the Trip Leader 1 day after return.</p>
    
    <div style="background-color: #f9f9f9; padding: 15px; border-left: 4px solid #4285f4;">
      <p>Hi there,</p>
      <p>We hope the trip "${tripData.trip_name}" went well yesterday.</p>
      <p>Please take 2 minutes to complete the mandatory evaluation form.</p>
      <p><a href="${preFilledUrl}"><b>Click Here to Complete Evaluation</b></a></p>
      <p style="font-size: 11px; color: #666;">(Link pre-filled with Ref: ${tripData.trip_ref})</p>
    </div>
  `;

  // 3. Send
  if (typeof sendTestEmail === 'function') {
    sendTestEmail(userEmail, "Trip Leader", subject, body);
  } else {
    MailApp.sendEmail({ to: userEmail, subject: subject, htmlBody: body });
    SpreadsheetApp.getActiveSpreadsheet().toast("Test Email Sent");
  }
}

// =============================================================================
// ADDS DUMMY EVALUATION
// =============================================================================

function test_SeedDummyResponse() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getActiveSheet();
  
  // 1. Get Current Trip Ref
  const tripRef = sheet.getName();
  if (!tripRef.match(/^\d{4}T\d{2}$/)) {
    ui.alert("Run this from a valid Trip Sheet.");
    return;
  }

  // 2. Get Response Sheet
  const responseSheet = ss.getSheetByName(CONFIG.FORM_RESPONSES_SHEET_NAME);
  if (!responseSheet) {
    ui.alert(`Response sheet '${CONFIG.FORM_RESPONSES_SHEET_NAME}' not found.`);
    return;
  }

  // 3. Construct Dummy Row
  // We need to ensure the Trip Ref lands in the correct column (CONFIG.FORM_COL_INDEX_REF)
  // Arrays are 0-indexed, so we create an array large enough.
  const colIndex = CONFIG.FORM_COL_INDEX_REF; // e.g. 2 for Column C
  
  const dummyRow = [];
  // Fill up to the target column with empty strings
  for (let i = 0; i < colIndex; i++) {
    dummyRow.push(""); 
  }
  
  // Set specific values
  dummyRow[0] = new Date(); // Timestamp (usually Col A / Index 0)
  dummyRow[1] = "test@example.com";
  dummyRow[colIndex] = tripRef; // The Trip Ref
  
  // Add some dummy answers after the ref
  dummyRow.push("Yes"); // Successfully concluded?
  dummyRow.push("Yes"); // Incidents of note?
  dummyRow.push("Delays due to the wrong leaves on the students."); // Description of incident(s)
  dummyRow.push("Conduct the trip in Spring."); // Changes next time?
  dummyRow.push("The leaves were a vibrant, descending kaleidoscope of amber, russet and gold, dancing in a crisp breeze before forming a whispering, crunching carpet of fiery color. But they got in the way!"); // Comments?

  // 4. Append
  responseSheet.appendRow(dummyRow);
  
  ui.alert("Success", `Dummy evaluation added for ${tripRef}.\n\nYou can now test 'Trip Admin > View Evaluation'.`, ui.ButtonSet.OK);
}

// =============================================================================
// TEST HELPERS
// =============================================================================

function isUserAllowedToTest() {
  const user = Session.getActiveUser().getEmail();
  
  // SAFE GUARD: If config is missing (file not updated), fallback to empty array to prevent crash
  const systemTesters = (typeof CONFIG !== 'undefined' && CONFIG.SYSTEM_TESTERS) ? CONFIG.SYSTEM_TESTERS : [];
  
  // 1. Check System Admins
  if (systemTesters.includes(user)) return true;
  
  // 2. Check SLT List (from Named Range)
  // This logic is safe because getRangeByName returns null if range doesn't exist
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rangeName = (typeof CONFIG !== 'undefined' && CONFIG.SLT_LIST_RANGE_NAME) ? CONFIG.SLT_LIST_RANGE_NAME : "authorisers_emails";
  const range = ss.getRangeByName(rangeName);
  
  if (range) {
    const sltEmails = range.getValues().flat().map(e => String(e).trim());
    if (sltEmails.includes(user)) return true;
  }
  
  return false;
}

function wrapTestBody(recipient, intendedAudience, bodyContent) {
  return `
    <div style="background-color: #fff3cd; color: #856404; padding: 15px; border: 1px solid #ffeeba; margin-bottom: 20px; font-family: sans-serif;">
      <strong>⚠️ TEST MODE</strong><br>
      <strong>Recipient:</strong> ${recipient} (You)<br>
      <strong>Intended Audience:</strong> ${intendedAudience}
    </div>
    <div style="font-family: sans-serif; color: #333;">
      ${bodyContent}
    </div>
  `;
}

function sendTestEmail(recipient, intendedAudience, subject, bodyContent) {
  const html = wrapTestBody(recipient, intendedAudience, bodyContent);
  MailApp.sendEmail({ to: recipient, subject: subject, htmlBody: html });
  SpreadsheetApp.getActiveSpreadsheet().toast(`Sent: ${subject}`, "Test Suite");
}

function getTestData() {
  const ui = SpreadsheetApp.getUi();

  // Security Check
  if (!isUserAllowedToTest()) {
    ui.alert("Access Denied", "You do not have permission to run system tests.", ui.ButtonSet.OK);
    return { sheet: null };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const validSheetPattern = /^\d{4}T\d{2}$/;
  if (!validSheetPattern.test(sheet.getName())) {
    ui.alert("Test Error", "Please run this from a valid Trip Sheet (e.g., 2511T01).", ui.ButtonSet.OK);
    return { sheet: null };
  }

  const tripData = getSummaryData(sheet);
  if (!tripData) {
    ui.alert("Test Error", "Could not read trip summary data.", ui.ButtonSet.OK);
    return { sheet: null };
  }

  return { 
    sheet: sheet, 
    tripData: tripData, 
    userEmail: Session.getActiveUser().getEmail(),
    ssUrl: `${ss.getUrl()}#gid=${sheet.getSheetId()}`
  };
}

function getTestStudentListHtml(sheet) {
  const data = getRangeData(sheet, CONFIG.STUDENT_DATA_RANGE_NAME).data;
  if (!data || data.length === 0) return "<ul><li>No students</li></ul>";
  const listItems = data.map(row => `<li>${row[0]}</li>`).join("");
  return `<ul>${listItems}</ul>`;
}
