/**
 * Duplicates the template sheet for each name in the given range
 * and renames the copies to match the names.
 *
 * @param {string} templateSheetName - Name of the sheet to duplicate (e.g. "Template")
 * @param {Range} namesRange - Range containing the list of names (single column)
 * @customfunction (you can also call it manually or from menu)
 */
function duplicateTemplateForNames(templateSheetName, sheetName,namesRange) {
  const ss = SpreadsheetApp.getActive();
  const templateSheet = ss.getSheetByName(templateSheetName);

  if (!templateSheet) {
    throw new Error(`Template sheet "${templateSheetName}" not found!`);
  }

  // Get the list of names (single column assumed)
  const rangeValues = "";
  const names = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(namesRange).getValues()
    .flat()
    .map(name => String(name).trim()) // convert to string & trim
    .filter(name => name !== "");     // remove blanks

  if (names.length === 0) {
    SpreadsheetApp.getUi().alert("No valid names found in the range.");
    return;
  }

  // Optional: Prevent duplicates (skip if sheet with this name already exists)
  const existingSheets = new Set(ss.getSheets().map(s => s.getName()));

  let createdCount = 0;

  names.forEach(name => {
    // Skip if sheet already exists
    if (existingSheets.has(name)) {
      Logger.log(`Sheet "${name}" already exists → skipping`);
      return;
    }

    // Duplicate the template
    const newSheet = templateSheet.copyTo(ss);
    newSheet.setName(name);

    // Optional: Move the new sheet to the end (or to a specific position)
    ss.setActiveSheet(newSheet);
    ss.moveActiveSheet(ss.getSheets().length);

    createdCount++;
  });

  SpreadsheetApp.getUi().alert(
    `Success!\nCreated ${createdCount} new sheets from template "${templateSheetName}".`
  );
}

function getTemplateSheetNames() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheets()
    .map(sheet => sheet.getName())
    .filter(name => name.toUpperCase().includes("TEMPLATE"));
}

function getNonTemplateSheetNames() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheets()
    .map(sheet => sheet.getName())
    .filter(name => !name.toUpperCase().includes("TEMPLATE"));
}

function getRelayDefaults() {
  const cfg = getTableConfig();
  return cfg.relayDefaults;
}


/**
 * Shows a dialog to let user select template name and range
 */
function showCreateSheetsDialog() {
  const ui = SpreadsheetApp.getUi();

  const html = HtmlService.createTemplateFromFile('sheets-from-template-ui')
                          .evaluate()
                          .setWidth(450)
                          .setHeight(480);

  ui.showModalDialog(html, 'Create Sheets from Template');
}
