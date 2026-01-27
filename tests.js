function debugSheetId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0]; // first visible sheet
  Logger.log("Sheet name: " + sheet.getName());
  Logger.log("Sheet ID:   " + sheet.getSheetId());
}