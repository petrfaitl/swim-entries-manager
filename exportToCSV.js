// Note: Shared functions (capitalizeNameSafely, normalizeSchoolYear, etc.)
// are now in EntryDataProcessor.js

/**
 * Shows a dialog to select a visible sheet and exports its data to CSV
 * Integrates with Custom Tools menu (only for owner)
 */
function exportEntriesToCSV() {
  const html = HtmlService.createTemplateFromFile('export-csv-ui')
      .evaluate()
      .setWidth(500)
      .setHeight(520);

  SpreadsheetApp.getUi().showModalDialog(html, 'Export Entries to CSV');
}

/**
 * Main processing function: reads the selected sheet and creates CSV
 * AUTO-DETECTS Event/Time columns from header row (J:AA or wherever they are)
 * @param {string|number} sheetNameOrId - Name or ID of the sheet to process
 * @param {string} [teamCodeTableName] - Name of the table to use for team code lookup (default: "Schools")
 * @param {string} [fileName] - Optional custom output file name
 * @return {string} URL to the created CSV file
 */
function processSheetToCSV(sheetNameOrId, teamCodeTableName, fileName) {
  teamCodeTableName = teamCodeTableName || "Schools"; // Default if not provided
  const ss = SpreadsheetApp.getActive();
  const sheet = resolveSheet_(ss, sheetNameOrId);
  const spreadsheetId = ss.getId();

  if (!sheet) {
    throw new Error(`Sheet "${sheetNameOrId}" not found`);
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

  let outputFileName = (fileName || "").trim();
  if (outputFileName) {
    if (!outputFileName.toLowerCase().endsWith(".csv")) {
      outputFileName += ".csv";
    }
  } else {
    outputFileName = `${sheetName}.csv`;
  }

  const blob = Utilities.newBlob(csvContent, MimeType.CSV, outputFileName);
  
  let fileId;
  try {
    // Use Advanced Drive Service for restricted scope compatibility
    const file = Drive.Files.create({
      name: outputFileName,
      mimeType: MimeType.CSV
    }, blob);
    fileId = file.id;
    Logger.log(`CSV created via Advanced Drive Service: ${outputFileName} (ID: ${fileId})`);
  } catch (e) {
    Logger.log(`Advanced Drive Service failed: ${e.message}. Falling back to DriveApp.`);
    const file = DriveApp.createFile(blob);
    fileId = file.getId();
    Logger.log(`CSV created via DriveApp: ${outputFileName} (ID: ${fileId})`);
  }

  return "https://drive.google.com/uc?export=download&id=" + fileId;
}

/**
 * Resolves a sheet by name, ID, or case-insensitive match.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string|number} sheetNameOrId
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function resolveSheet_(ss, sheetNameOrId) {
  if (sheetNameOrId === null || sheetNameOrId === undefined || sheetNameOrId === '') {
    return null;
  }

  if (typeof sheetNameOrId === 'string') {
    const sheetByName = ss.getSheetByName(sheetNameOrId);
    if (sheetByName) return sheetByName;
  }

  const idNum = Number(sheetNameOrId);
  if (!isNaN(idNum)) {
    const sheetById = ss.getSheetById(idNum);
    if (sheetById) return sheetById;
  }

  if (typeof sheetNameOrId === 'string') {
    const trimmed = sheetNameOrId.trim().toLowerCase();
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].getName().trim().toLowerCase() === trimmed) {
        return sheets[i];
      }
    }
  }

  return null;
}
