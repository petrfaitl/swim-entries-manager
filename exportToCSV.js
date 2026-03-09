// Note: Shared functions (capitalizeNameSafely, normalizeSchoolYear, etc.)
// are now in EntryDataProcessor.js

/**
 * Shows a dialog to select a visible sheet and exports its data to CSV
 * Integrates with Custom Tools menu (only for owner)
 */
function exportEntriesToCSV() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  // Get all sheets that are NOT hidden
  const visibleSheets = ss.getSheets().filter(sheet => !sheet.isSheetHidden());

  if (visibleSheets.length === 0) {
    ui.alert("No visible sheets found in this spreadsheet.");
    return;
  }

  // Get available tables for team code lookup
  let tableNames = [];
  let hasSchoolsTable = false;
  try {
    const tableApp = TableApp.openById(ss.getId());
    const tablesObj = tableApp.getTables();
    Logger.log('[exportEntriesToCSV] getTables() returned: %s', JSON.stringify(tablesObj));

    if (tablesObj && typeof tablesObj === 'object') {
      Object.keys(tablesObj).forEach(function(sheetName) {
        const sheetData = tablesObj[sheetName];
        if (sheetData && sheetData.tables && Array.isArray(sheetData.tables)) {
          sheetData.tables.forEach(function(tableWrapper) {
            if (tableWrapper && tableWrapper.table && tableWrapper.table.name) {
              tableNames.push(tableWrapper.table.name);
            }
          });
        }
      });
      Logger.log('[exportEntriesToCSV] Extracted table names: %s', tableNames.join(', '));
    } else {
      Logger.log('[exportEntriesToCSV] Unexpected tables structure');
      tableNames = [];
    }
    hasSchoolsTable = tableNames.includes('Schools');
  } catch (err) {
    Logger.log('[exportEntriesToCSV] Could not load tables: %s', err.message);
    tableNames = [];
  }

  // Build HTML for sheet dropdown
  const sheetOptionsHtml = visibleSheets.map(sheet =>
    `<option value="${sheet.getSheetId()}">${sheet.getName()}</option>`
  ).join('');

  // Build HTML for table dropdown
  let tableOptionsHtml = '';
  if (tableNames.length === 0) {
    tableOptionsHtml = '<option value="">No tables found</option>';
  } else {
    tableOptionsHtml = tableNames.map(name => {
      const selected = (name === 'Schools' && hasSchoolsTable) ? ' selected' : '';
      return `<option value="${name}"${selected}>${name}</option>`;
    }).join('');
  }

  const template = HtmlService.createTemplateFromFile('export-csv-ui');
  template.sheetOptionsHtml = sheetOptionsHtml;
  template.tableOptionsHtml = tableOptionsHtml;

  const html = template.evaluate()
                       .setWidth(480)
                       .setHeight(450);

  ui.showModalDialog(html, 'Export Entries to CSV');
}

/**
 * Main processing function: reads the selected sheet and creates CSV
 * AUTO-DETECTS Event/Time columns from header row (J:AA or wherever they are)
 * @param {string} sheetId - ID of the sheet to process
 * @param {string} teamCodeTableName - Name of the table to use for team code lookup (default: "Schools")
 * @return {string} URL to the created CSV file
 */
function processSheetToCSV(sheetId, teamCodeTableName) {
  teamCodeTableName = teamCodeTableName || "Schools"; // Default if not provided
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetById(sheetId);
  const spreadsheetId = ss.getId();

  if (!sheet) {
    throw new Error(`Sheet with ID ${sheetId} not found`);
  }

  const sheetName = sheet.getName();
  Logger.log(`[exportToCSV] Processing sheet: ${sheetName}`);

  // Use shared processing function
  const processedRows = processEntriesSheet(sheet, teamCodeTableName, spreadsheetId);

  const headers = [
    "Team Code", "First Name", "Last Name", "Date of Birth", "Gender",
    "Event Entries", "Entry Times", "School Year"
  ];

  const outputRows = [headers];

  processedRows.forEach(swimmer => {
    outputRows.push([
      swimmer.teamCode,
      swimmer.firstName,
      swimmer.lastName,
      swimmer.dobFormatted,
      swimmer.gender,
      swimmer.eventsStr,
      swimmer.timesStr,
      swimmer.schoolYear
    ]);
  });

  // ── CREATE & RETURN CSV ──────────────────────────────────────────────
  const csvContent = outputRows.map(row =>
    row.map(cell => `"${(cell || "").toString().replace(/"/g, '""')}"`).join(",")
  ).join("\n");

  const fileName = `${sheetName}.csv`;
  const blob = Utilities.newBlob(csvContent, MimeType.CSV, fileName);
  const file = DriveApp.getRootFolder().createFile(blob);

  Logger.log(`CSV created: ${fileName} with ${outputRows.length - 1} data rows`);
  return file.getDownloadUrl();
}
