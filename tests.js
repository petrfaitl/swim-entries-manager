function debugSheetId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0]; // first visible sheet
  Logger.log("Sheet name: " + sheet.getName());
  Logger.log("Sheet ID:   " + sheet.getSheetId());
}

function testFlexibleTimeParsing() {
  const tests = [
    "1:02.3", "01:02.30", "1:2.3", "01:02.300", "1:2", "01:2.3",
    "0:59.99", "59.99", "123.45", "2:3", "2:03", "2:3.0",
    "10:5.678", "0:0.05", "", "NT already"
  ];

  tests.forEach(t => {
    const result = formatAsMmSsSs(t);
    Logger.log(`"${t}" → "${result}"`);
  });
}

function test_getNamedTableData() {
  // Assuming a named table exists in your sheet with the name "MyTable"
  const tableName = "Teams";
  const spreadsheetId = "1KFNv7LIAyrDPl4Na8XNsCMXJCg_BDqyuUo_q3lwEklU";
  const data = getNamedTableData(spreadsheetId,tableName);
  // const data = getTableByName(tableName);

  // Process or manipulate table data as needed
  Logger.log(`First cell of the table: ${data[0][0]}`);
}
//test_getNamedTableData();



function test_getAllTables(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = "1KFNv7LIAyrDPl4Na8XNsCMXJCg_BDqyuUo_q3lwEklU";
  const tableApp = TableApp.openById(spreadsheetId);
  const allTables = tableApp.getTables();
  for (table in allTables){
  Logger.log(`All tables in the workbook: ${table}`);
  }

}

