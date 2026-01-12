/**
 * WORKFLOW PDF GENERATOR
 * Generates the "Leader Pack" PDF attachment for T-1 notifications.
 */

/**
 * Generates a multi-page PDF Leader Pack.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The trip sheet.
 * @param {Object} tripData - The summary data object.
 * @return {GoogleAppsScript.Base.Blob} The PDF blob.
 */
function generateLeaderPackPDF(sheet, tripData) {
  // 1. Get Data Sources
  const students = getRangeData(sheet, CONFIG.STUDENT_DATA_RANGE_NAME).data;
  const medSummary = getRangeData(sheet, CONFIG.MEDICAL_DATA_RANGE_NAME).data; 
  
  // 2. Build HTML Sections
  const css = getPdfCss();
  const header = getPdfHeader(tripData);
  
  const section1 = buildSection1_Register(students);
  const section2 = buildSection2_Contacts(students);
  const section3 = buildSection3_Medical(students);
  const section4 = buildSection4_MedSummary(medSummary);
  const section5 = buildSection5_RiskAndStaff(tripData);

  // 3. Assemble Final HTML
  const htmlContent = `
    <html>
      <head><style>${css}</style></head>
      <body>
        ${header}
        ${section1}
        <div class="page-break"></div>
        
        ${header}
        ${section2}
        <div class="page-break"></div>
        
        ${header}
        ${section3}
        <div class="page-break"></div>
        
        ${header}
        ${section4}
        <div class="page-break"></div>
        
        ${header}
        ${section5}
      </body>
    </html>
  `;

  // 4. Convert to PDF
  const blob = Utilities.newBlob(htmlContent, MimeType.HTML, "LeaderPack.html");
  const pdf = blob.getAs(MimeType.PDF).setName(`LeaderPack - ${tripData.trip_name}.pdf`);
  
  return pdf;
}

// =============================================================================
// SECTION BUILDERS
// =============================================================================

function buildSection1_Register(data) {
  // Cols: Name(0), PP(4), SEN(5), FSM(7)
  if (!data || data.length === 0) return "<p>No student data found.</p>";

  // Headcount Summary
  const count = data.length;
  const headcountHtml = `
    <div style="border: 1px solid #333; padding: 5px; margin-bottom: 15px; width: 300px;">
      Total Students: <b>${count}</b> &nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp; Revised Headcount: __________
    </div>
  `;

  let rows = "";
  data.forEach(row => {
    // Empty string instead of "-"
    const pp = row[4] || "";
    const sen = row[5] || "";
    const fsm = row[7] || "";

    rows += `
      <tr>
        <td><b>${row[0]}</b></td>
        <td>${pp}</td>
        <td>${sen}</td>
        <td>${fsm}</td>
        <td class="box"></td>
        <td class="box"></td>
        <td class="box"></td>
      </tr>`;
  });

  return `
    <h3>Section 1: Registration</h3>
    ${headcountHtml}
    <table>
      <thead>
        <tr>
          <th width="30%">Student</th>
          <th width="10%">PP</th>
          <th width="10%">SEN</th>
          <th width="10%">FSM</th>
          <th width="10%">Out</th>
          <th width="10%">Lunch</th>
          <th width="10%">Back</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

function buildSection2_Contacts(data) {
  // Cols: Name(0), Contact1(1), Contact2(2)
  if (!data || data.length === 0) return "<p>No student data found.</p>";

  let rows = "";
  data.forEach(row => {
    rows += `
      <tr>
        <td><b>${row[0]}</b></td>
        <td style="font-size:10px;">${row[1] || ""}</td>
        <td style="font-size:10px;">${row[2] || ""}</td>
      </tr>`;
  });

  return `
    <h3>Section 2: Emergency Contacts</h3>
    <table>
      <thead>
        <tr>
          <th width="30%">Student</th>
          <th width="35%">Contact 1</th>
          <th width="35%">Contact 2</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

function buildSection3_Medical(data) {
  // Cols: Name(0), Med(6), Diet(8)
  let rows = "";
  let count = 0;

  data.forEach(row => {
    let med = String(row[6] || "").trim();
    let diet = String(row[8] || "").trim();
    const medClean = med.toLowerCase();

    // Skip if Med is "none recorded"/"none" AND diet is empty
    if ((!med || medClean.includes("none recorded") || medClean === "none") && !diet) {
      return; 
    }

    // Grey out "None recorded" text for visual clarity
    if (medClean.includes("none recorded") || medClean === "none") {
      med = `<span style="color:#777;">${med}</span>`;
    }

    rows += `
      <tr>
        <td><b>${row[0]}</b></td>
        <td>${med}</td>
        <td>${diet}</td>
      </tr>`;
    count++;
  });

  if (count === 0) return "<h3>Section 3: Medical & Dietary</h3><p>No flagged conditions found.</p>";

  return `
    <h3>Section 3: Medical & Dietary</h3>
    <p><i>Only showing students with recorded conditions or dietary needs.</i></p>
    <table>
      <thead>
        <tr>
          <th width="30%">Student</th>
          <th width="35%">Medical Conditions</th>
          <th width="35%">Dietary Requirements</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

function buildSection4_MedSummary(data) {
  if (!data || data.length === 0) return "<h3>Section 4: Medical Notes Summary</h3><p>No summary notes recorded.</p>";

  // Logic: Iterate through cells. If a cell starts with a Letter (A-Z), add a divider line above it.
  // Exception: Don't add a line above the very first item.
  let htmlBlock = "";
  
  data.forEach((row, index) => {
    const text = String(row[0]);
    // Regex: Starts with a letter?
    const startsWithLetter = /^[a-zA-Z]/.test(text);
    
    if (index > 0 && startsWithLetter) {
      htmlBlock += `<hr style="border: 0; border-top: 1px solid #ccc; margin: 10px 0;">`;
    } else if (index > 0) {
      // Just a line break for dates/continuations
      htmlBlock += `<br>`;
    }
    
    htmlBlock += text;
  });

  return `
    <h3>Section 4: Medical Notes Summary</h3>
    <p><i>Note: Students in this summary are listed alphabetically.</i></p>
    <div style="border: 1px solid #333; padding: 10px; font-family: sans-serif; white-space: pre-wrap;">
      ${htmlBlock}
    </div>
  `;
}

function buildSection5_RiskAndStaff(tripData) {
  // 1. Build Staff Table
  const leadName = tripData.trip_leadName || "Unknown";
  const leadMobile = tripData.trip_leadMobile || "-";
  
  // Safe Staff Parsing: Split, Zip, then Filter
  // This prevents index mismatch if one string has fewer commas than the other
  const rawNames = (tripData.trip_staffNames || "").split(',');
  const rawMobiles = (tripData.trip_staffMobiles || "").split(',');
  const maxLength = Math.max(rawNames.length, rawMobiles.length);
  
  let staffRows = `<tr><td><b>${leadName} (Lead)</b></td><td>${leadMobile}</td></tr>`;
  
  for (let i = 0; i < maxLength; i++) {
    const name = (rawNames[i] || "").trim();
    const mobile = (rawMobiles[i] || "").trim();
    
    // Only add row if there is a name
    if (name) {
      staffRows += `<tr><td>${name}</td><td>${mobile || "-"}</td></tr>`;
    }
  }

  const staffTable = `
    <table class="ra-table">
      <tr><th width="60%">Staff Name</th><th width="40%">Mobile Number</th></tr>
      ${staffRows}
    </table>
  `;

  // 2. Risk Assessment Logic
  const prior = tripData.trip_raTravelPrior || "None";
  const travelRisks = tripData.trip_raTravelRisks || "None";
  const travelControls = tripData.trip_raTravelControls || "None";
  const behRisks = tripData.trip_raBehaviourRisks || "None";
  const behControls = tripData.trip_raBehaviourControls || "None";

  return `
    <h3>Section 5: Staffing & Risk Assessment</h3>
    
    <h4>Staff Emergency Contacts</h4>
    ${staffTable}

    <h4>1. Travel (Prior)</h4>
    <div class="ra-box">${prior}</div>

    <h4>2. Travel Risks & Controls</h4>
    <table class="ra-table">
      <tr><th>Risks</th><th>Controls</th></tr>
      <tr>
        <td valign="top">${travelRisks}</td>
        <td valign="top">${travelControls}</td>
      </tr>
    </table>

    <h4>3. Behaviour Risks & Controls</h4>
    <table class="ra-table">
      <tr><th>Risks</th><th>Controls</th></tr>
      <tr>
        <td valign="top">${behRisks}</td>
        <td valign="top">${behControls}</td>
      </tr>
    </table>
  `;
}

// =============================================================================
// STYLES & UTILS
// =============================================================================

function getPdfHeader(tripData) {
  return `
    <div class="header">
      <div style="font-weight:bold; color:red; font-size:10px; margin-bottom:5px;">HIGHLY CONFIDENTIAL</div>
      <strong>${tripData.trip_name} (${tripData.trip_ref})</strong><br>
      Lead: ${tripData.trip_leadName} | School Emergency No: 020 7485 3414
    </div>
    <hr>
  `;
}

function getPdfCss() {
  return `
    body { font-family: sans-serif; font-size: 10px; color: #333; }
    h3 { font-size: 14px; margin-bottom: 10px; color: #2c3e50; }
    h4 { font-size: 12px; margin-top: 10px; margin-bottom: 5px; background: #eee; padding: 3px; }
    
    table { width: 100%; border-collapse: collapse; margin-bottom: 15px; }
    th { background-color: #f2f2f2; border: 1px solid #999; padding: 5px; text-align: left; font-size: 9px; }
    td { border: 1px solid #999; padding: 4px; font-size: 9px; vertical-align: top; }
    
    /* FIXED: Matches table cell border exactly to prevent thick black lines */
    .box { border: 1px solid #999; }
    
    .page-break { page-break-after: always; }
    .header { font-size: 12px; text-align: center; margin-bottom: 5px; }
    
    .ra-box { border: 1px solid #ccc; padding: 5px; margin-bottom: 10px; min-height: 30px; }
    .ra-table td { width: 50%; }
  `;
}
