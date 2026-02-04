/**
 * WORKFLOW TIME-BASED (v4.2 - Date Fix Applied)
 * Handles T-Minus Automated checks (T-7, T-4, T-1, T-0).
 * Includes Duplicate Prevention, PDF Generation, Manual Force.
 * UPDATED: Fixed "dd/MM/yyyy" parsing issue for non-US locales.
 */

function runDailyChecks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const validSheetPattern = /^\d{4}T\d{2}$/;
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  console.log(`Daily Check Started: ${today}`);

  sheets.forEach(sheet => {
    if (!validSheetPattern.test(sheet.getName())) return;

    const tripData = getSummaryData(sheet);
    if (!tripData || !tripData.trip_date) return;
    
    const status = tripData.trip_status;
    if (status !== CONFIG.STATUS_AUTHORISED && status !== CONFIG.STATUS_EDITED) return;

    // --- FIX STARTS HERE ---
    // We use the helper to ensure dd/mm/yyyy is parsed correctly
    const tripDate = parseBritishDate(tripData.trip_date);
    
    if (!tripDate || isNaN(tripDate.getTime())) {
      console.warn(`Skipped ${sheet.getName()}: Invalid date format (${tripData.trip_date})`);
      return;
    }
    // --- FIX ENDS HERE ---

    tripDate.setHours(0, 0, 0, 0);
    
    const diffTime = tripDate - today;
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

    // Debug log to confirm calculation (remove later if too noisy)
    console.log(`Sheet: ${sheet.getName()} | Trip Date: ${tripDate.toDateString()} | Diff: ${diffDays} days`);

    switch (diffDays) {
      case 7:
        if (!tripData.trip_sentLockIn) sendT7_LockIn(sheet, tripData);
        break;
      case 4:
        sendT4_Operations(sheet, tripData);
        break;
      case 1:
        if (!tripData.trip_sentLeaderPack) sendT1_LeaderPack(sheet, tripData);
        break;
      case 0:
        if (!tripData.trip_sentDeparture) sendT0_AttendanceReminder(sheet, tripData);
        break;
      case -1:
        if (!tripData.trip_sentEvaluation) sendTPlus1_Evaluation(sheet, tripData);
        break;
    }
  });
}

function forceTripNotifications() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  const validSheetPattern = /^\d{4}T\d{2}$/;
  if (!validSheetPattern.test(sheet.getName())) {
    ui.alert("Invalid Sheet", "Please run this command from a valid Trip Sheet.", ui.ButtonSet.OK);
    return;
  }

  const tripData = getSummaryData(sheet);
  if (!tripData) return;

  const prompt = ui.prompt(
    "Force Notification", 
    `Force a notification for trip ${tripData.trip_ref}.\n\n` +
    "Enter the T-Minus number:\n" +
    "• 7  = Lock-in Reminder (Lead)\n" +
    "• 4  = Operations Lists (Att/Cover)\n" +
    "• 1  = Leader Pack (Lead)\n" +
    "• 0  = Departure Reminder (Att)\n",
    ui.ButtonSet.OK_CANCEL
  );

  if (prompt.getSelectedButton() !== ui.Button.OK) return;
  const input = prompt.getResponseText().trim();

  try {
    switch (input) {
      case "7":
        sendT7_LockIn(sheet, tripData);
        ui.alert("T-7 (Lock-in) forced successfully.");
        logToSystem("Manual Force", tripData.trip_ref, "User forced T-7 Notification");
        break;
      case "4":
        sendT4_Operations(sheet, tripData);
        ui.alert("T-4 (Operations) forced successfully.");
        logToSystem("Manual Force", tripData.trip_ref, "User forced T-4 Notification");
        break;
      case "1":
        sendT1_LeaderPack(sheet, tripData);
        ui.alert("T-1 (Leader Pack) forced successfully.");
        logToSystem("Manual Force", tripData.trip_ref, "User forced T-1 Notification");
        break;
      case "0":
        sendT0_AttendanceReminder(sheet, tripData);
        ui.alert("T-0 (Departure) forced successfully.");
        logToSystem("Manual Force", tripData.trip_ref, "User forced T-0 Notification");
        break;
      default:
        ui.alert("Invalid input. Please enter 7, 4, 1, or 0.");
    }
  } catch (e) {
    ui.alert("Error forcing notification: " + e.message);
    console.error(e);
  }
}

// =============================================================================
// WORKFLOW ACTIONS
// =============================================================================

function sendT7_LockIn(sheet, tripData) {
  // Recipient: Trip Lead
  // CC: SLT Authoriser
  // BCC: System Admin
  
  const recipients = tripData.trip_leadEmail;
  const sltEmail = tripData.trip_authorisedBy;
  const subject = `ACTION REQUIRED: 7 Days to Go - ${tripData.trip_name} (${tripData.trip_ref})`;
  
  const body = `
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
    <p><a href="${SpreadsheetApp.getActiveSpreadsheet().getUrl()}#gid=${sheet.getSheetId()}">Open Trip Sheet</a></p>
  `;

  MailApp.sendEmail({ 
    to: recipients, 
    cc: sltEmail,
    bcc: CONFIG.SYSTEM_ADMIN_BCC,
    subject: subject, 
    htmlBody: body 
  });
  
  updateSummaryValue(sheet, 'trip_sentLockIn', new Date());
  logToSystem("T-7 Check", tripData.trip_ref, "Lock-in reminder sent");
}

function sendT4_Operations(sheet, tripData) {
  const tripDateStr = tripData.trip_date instanceof Date ? 
    Utilities.formatDate(tripData.trip_date, Session.getScriptTimeZone(), "dd/MM/yyyy") : String(tripData.trip_date);

  let leaveTime = formatTime(tripData.trip_timeLeaving);
  let returnTime = formatTime(tripData.trip_timeReturning);
  const sltEmail = tripData.trip_authorisedBy;

  // 1. ATTENDANCE
  if (!tripData.trip_sentOpsAttendance) {
    const attendanceEmails = getOfficerEmails(CONFIG.RANGE_OFFICERS_ATTENDANCE);
    if (attendanceEmails.length > 0) {
      const studentListHtml = getStudentListHtml(sheet);
      const attSubject = `TRIP ATTENDANCE: ${tripData.trip_name} (${tripData.trip_ref})`;
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
      MailApp.sendEmail({ 
        to: attendanceEmails.join(","), 
        cc: sltEmail,
        bcc: CONFIG.SYSTEM_ADMIN_BCC,
        subject: attSubject, 
        htmlBody: attBody 
      });
      
      updateSummaryValue(sheet, 'trip_sentOpsAttendance', new Date());
      logToSystem("T-4 Check", tripData.trip_ref, "Attendance list sent");
    }
  }

  // 2. COVER
  if (!tripData.trip_sentOpsCover) {
    const coverEmails = getOfficerEmails(CONFIG.RANGE_OFFICERS_COVER);
    if (coverEmails.length > 0) {
      const covSubject = `TRIP COVER: ${tripData.trip_name} (${tripData.trip_ref})`;
      
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

      const covBody = `
        <h3>Trip Departing in 4 Days</h3>
        <p><b>Date:</b> ${tripDateStr}</p>
        <p><b>Times:</b> ${leaveTime} - ${returnTime}</p>
        
        <h4>Staff Requiring Cover:</h4>
        ${staffHtml}
        
        <p>Please check the spreadsheet below for details.</p>
        <p><a href="${SpreadsheetApp.getActiveSpreadsheet().getUrl()}#gid=${sheet.getSheetId()}">Open Trip Sheet</a></p>
      `;
      MailApp.sendEmail({ 
        to: coverEmails.join(","), 
        cc: sltEmail,
        bcc: CONFIG.SYSTEM_ADMIN_BCC,
        subject: covSubject, 
        htmlBody: covBody 
      });
      
      updateSummaryValue(sheet, 'trip_sentOpsCover', new Date());
      logToSystem("T-4 Check", tripData.trip_ref, "Cover list sent");
    }
  }
}

function sendT1_LeaderPack(sheet, tripData) {
  const leaderEmail = tripData.trip_leadEmail;
  const sltEmail = tripData.trip_authorisedBy;
  if (!leaderEmail) return;

  const subject = `PRINT NOW: Leader Pack for ${tripData.trip_name}`;
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
    <p>Please print this or save it to your device for offline access.</p>
  `;

  let pdfBlob = null;
  try {
    if (typeof generateLeaderPackPDF === 'function') {
      pdfBlob = generateLeaderPackPDF(sheet, tripData);
    } else {
      throw new Error("PDF Generator function missing in Workflow_PDF.gs");
    }
  } catch (e) {
    console.error("Failed to generate PDF: " + e.message);
    logToSystem("PDF Error", tripData.trip_ref, e.message);
  }

  const emailOptions = {
    to: leaderEmail,
    cc: sltEmail,
    bcc: CONFIG.SYSTEM_ADMIN_BCC,
    subject: subject,
    htmlBody: body
  };

  if (pdfBlob) {
    emailOptions.attachments = [pdfBlob];
  } else {
    emailOptions.htmlBody += "<p style='color:red'><b>Error:</b> PDF could not be generated. Please check the system log.</p>";
  }

  MailApp.sendEmail(emailOptions);
  
  updateSummaryValue(sheet, 'trip_sentLeaderPack', new Date());
  logToSystem("T-1 Check", tripData.trip_ref, "Leader Pack (PDF) sent");
}

function sendT0_AttendanceReminder(sheet, tripData) {
  const attendanceEmails = getOfficerEmails(CONFIG.RANGE_OFFICERS_ATTENDANCE);
  const sltEmail = tripData.trip_authorisedBy;
  if (attendanceEmails.length === 0) return;

  let leaveTime = formatTime(tripData.trip_timeLeaving);
  let returnTime = formatTime(tripData.trip_timeReturning);

  const subject = `REMINDER: Trip Departing TODAY - ${tripData.trip_name}`;
  const body = `
    <p>This is a reminder that <b>${tripData.trip_ref}</b> departs today.</p>
    <p><b>Times:</b> ${leaveTime} - ${returnTime}</p>
    <p>Please ensure the register is marked.</p>
    <p style="color: red; font-weight: bold;">⚠️ NOTE: The student list may have changed since the last notification (T-4). Please check the sheet below for the final list.</p>
    <p><a href="${SpreadsheetApp.getActiveSpreadsheet().getUrl()}#gid=${sheet.getSheetId()}">Open Trip Sheet</a></p>
  `;

  MailApp.sendEmail({ 
    to: attendanceEmails.join(","), 
    cc: sltEmail,
    bcc: CONFIG.SYSTEM_ADMIN_BCC,
    subject: subject, 
    htmlBody: body 
  });
  
  updateSummaryValue(sheet, 'trip_sentDeparture', new Date());
  logToSystem("T-0 Check", tripData.trip_ref, "Departure reminder sent");
}

function sendTPlus1_Evaluation(sheet, tripRef) {
  const leaderEmail = sheet.getRange(CONFIG.RANGE_LEADER_EMAIL).getValue();
  const tripName = sheet.getRange(CONFIG.RANGE_TRIP_NAME).getValue();

  if (!leaderEmail) {
    Logger.log(`Skipping email for ${tripRef}: No leader email found.`);
    return;
  }

  // Build Pre-filled URL
  const formUrl = `${CONFIG.FORM_EVALUATION_URL}?usp=pp_url&${CONFIG.FORM_ENTRY_ID_TRIP_REF}=${tripRef}`;

  const subject = `Trip Evaluation Required: ${tripName} (${tripRef})`;
  const body = `
    Hi there,
    
    We hope the trip "${tripName}" went well yesterday.
    
    Please take 2 minutes to complete the mandatory evaluation form.
    Click the link below (Reference number is already filled in for you):
    
    ${formUrl}
  `;

  MailApp.sendEmail({
    to: leaderEmail,
    subject: subject,
    body: body
  });
  
  updateSummaryValue(sheet, 'trip_sentEvaluation', new Date());
  logToSystem("T+1 Evaluation", tripData.trip_ref, "Evaluation form sent");
}

// =============================================================================
// HELPERS
// =============================================================================

function getOfficerEmails(rangeName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const range = ss.getRangeByName(rangeName);
  if (!range) return [];
  const raw = range.getValues().flat();
  return [...new Set(raw.filter(email => email && email.toString().includes("@")))];
}

function getStudentListHtml(sheet) {
  const data = getRangeData(sheet, CONFIG.STUDENT_DATA_RANGE_NAME).data;
  if (!data || data.length === 0) return "<ul><li>No students</li></ul>";
  const listItems = data.map(row => `<li>${row[0]}</li>`).join("");
  return `<ul>${listItems}</ul>`;
}

function formatTime(val) {
  if (!val) return "TBC";
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "HH:mm");
  return String(val);
}

/**
 * Helper to safely parse "dd/MM/yyyy" strings into Date objects.
 * Handles cases where the input is already a Date object.
 */
function parseBritishDate(dateInput) {
  // Case 1: It's already a Date object (Sheets might return this automatically)
  if (dateInput instanceof Date) {
    return new Date(dateInput); // Return a clone
  }

  // Case 2: It's a string like "25/12/2025"
  if (typeof dateInput === 'string') {
    const parts = dateInput.split('/');
    if (parts.length === 3) {
      // new Date(year, monthIndex, day) -> Month is 0-indexed (Jan=0, Dec=11)
      return new Date(parts[2], parts[1] - 1, parts[0]);
    }
  }

  // Fallback: Return null or Invalid Date if format is unrecognizable
  console.warn("Could not parse date:", dateInput);
  return null; 
}
