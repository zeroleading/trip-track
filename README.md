Trip-Track 🌍

Trip-Track is a professional-grade Google Apps Script (GAS) system integrated with Google Sheets, designed specifically for educational institutions to manage school trips. It automates the lifecycle of a trip from initial request and SLT authorization to medical data management and post-trip evaluations.

🚀 Key Features

⚖️ Multi-Stage Authorization Workflow

Request & Validate: Trip leaders can submit authorization requests once minimum requirements (marked fields/student count) are met.

SLT Approval/Denial: Senior Leadership Team (SLT) members can approve or deny trips directly via a custom "Trip Admin" menu.

Automated Archiving: Upon approval, the system generates a values-only "Snapshot" archive in a designated Drive folder, ensuring a permanent record of the state at the time of authorization.

📋 Student & Medical Intelligence

Dynamic Student Management: Optimized logic to "tick" and "untick" students from a master list into specific trip itineraries.

Smart Sorting: Automatically switches between A-Z sorting and Registration Group sorting (for groups >50) to facilitate easier roll-calls.

Leader Pack (PDF): One-click generation of a "Leader Pack" PDF including registration lists, emergency contacts, medical/dietary red flags, and risk assessment summaries.

🕵️ Post-Auth Audit Trail

Change Monitoring: Uses installable triggers to monitor any edits made to a trip after it has been authorized.

Audit Logging: Edits are automatically logged in the trip's archive file, recording the timestamp, user, cell reference, old value, and new value.

🔗 Document & Link Management

Drive File Picker: Custom sidebar UI to search Google Drive and attach invoices, letters, or risk assessments directly to the summary table.

Auto-Permissions: The system attempts to automatically grant "View" permissions for attached documents to relevant SLT members and system testers.

📧 Automated Notification Suite

Lifecycle Alerts: Automated emails at key intervals:

T-7: "Lock-in" reminder for leaders.

T-4: Operational alerts for Attendance Officers (MIS coding) and Cover Supervisors.

T-1: Leader Pack PDF delivery.

T-0: Departure day reminder.

T+1: Mandatory evaluation form request.

📁 Repository Structure

The project is modularized for ease of maintenance:

File

Description

Config.gs

Centralizes Folder IDs, Template IDs, status strings, and color hex codes.

Helpers.gs

Core utility functions for data extraction, Drive search, and logging.

Navigation.gs

UI helpers to toggle between "Student View", "RA View", and "Med Notes".

Testing.gs

A robust testing suite allowing admins to simulate the full trip lifecycle emails.

Triggers.gs

Handles onOpen (Menu creation) and onEdit (Fast UI updates/Checkboxes).

Workflow_Approval.gs

Logic for Request, Approve (with Archive creation), and Denial workflows.

Workflow_Docs.gs

HTML/JS sidebar code for the Drive File Picker and Link Manager.

Workflow_Monitoring.gs

Logic for logging post-authorization changes to the Audit Trail.

Workflow_PDF.gs

HTML-to-PDF engine for building the multi-page Leader Pack.

Workflow_StudentTools.gs

Add/Remove student logic and the "Smart Sort" algorithm.

🛠️ Setup & Configuration

1. Spreadsheet Setup

The script expects a "Master Trip Template" tab (default name 0000T00) and a "Menu" tab. Ensure the following named ranges are present:

thisTrip_summaryTable (2-column data block)

thisTrip_studentData (Student details block)

authorisers_emails (List of SLT members)

2. Configuration (Config.gs)

Update the CONFIG object with your specific institutional IDs:

const CONFIG = {
  FOLDER_ID_BUILD: "...",   // Active trip sheets
  FOLDER_ID_ARCHIVE: "...", // Finalized snapshots
  ARCHIVE_TEMPLATE_ID: "...",
  SYSTEM_ADMIN_BCC: "admin@school.com",
  FORM_EVALUATION_URL: "[https://docs.google.com/forms/](https://docs.google.com/forms/)..."
};


3. Triggers

onOpen and onEdit are simple triggers.

Manual Action Required: You must set up an Installable Trigger for the function trigger_MonitorChanges to run on "On edit" to enable the Audit Trail logging (which requires Drive/Email permissions).

📄 License

This project is proprietary to Camden School for Girls.

Developed by zeroleading
