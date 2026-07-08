function updateHourlyTimestamp() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('menu'); // Make sure your tab is named 'menu'
  
  if (sheet) {
    // Writes the exact current date & time into cell G3
    sheet.getRange('G3').setValue(new Date());
  }
}