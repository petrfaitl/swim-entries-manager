/**
 * Duplicates the template sheet for each name in the given range
 * and renames the copies to match the names.
 *
 * @param {string} templateSheetName - Name of the sheet to duplicate (e.g. "Template")
 * @param {string} sheetName - Name of the source sheet or TableApp table containing names
 * @param {Range|string} namesRange - Range containing the list of names (single column)
 * @customfunction (you can also call it manually or from menu)
 */
function duplicateTemplateForNames(templateSheetName, sheetName, namesRange) {
  const ss = SpreadsheetApp.getActive();
  const templateSheet = ss.getSheetByName(templateSheetName);

  if (!templateSheet) {
    throw new Error(`Template sheet "${templateSheetName}" not found!`);
  }

  // Resolve source sheet (exact match, case/whitespace-insensitive fallback, or TableApp lookup)
  const sourceSheet = resolveSheetByName_(ss, sheetName);

  if (!sourceSheet) {
    throw new Error(`Sheet "${sheetName}" not found!`);
  }

  // Get the list of names (single column assumed)
  const names = sourceSheet.getRange(namesRange).getValues()
    .flat()
    .map(name => String(name).trim()) // convert to string & trim
    .filter(name => name !== "");     // remove blanks

  if (names.length === 0) {
    throw new Error(`No valid names found in range "${namesRange}" on sheet "${sheetName}".`);
  }

  // Optional: Prevent duplicates (skip if sheet with this name already exists)
  const existingSheets = new Set(ss.getSheets().map(s => s.getName()));

  let createdCount = 0;
  let skippedCount = 0;

  names.forEach(name => {
    // Skip if sheet already exists
    if (existingSheets.has(name)) {
      Logger.log(`Sheet "${name}" already exists → skipping`);
      skippedCount++;
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

  return {
    success: true,
    createdCount: createdCount,
    skippedCount: skippedCount,
    message: `Created ${createdCount} sheet(s) from template "${templateSheetName}"` +
      (skippedCount > 0 ? ` (${skippedCount} skipped as already existing)` : '') + '.'
  };
}

function getTemplateSheetNames() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheets()
    .map(sheet => sheet.getName())
    .filter(name => name.toUpperCase().includes("TEMPLATE"));
}

/**
 * Returns available School/Team table options and default selection
 * for the Create Sheets from Template dialog.
 * @return {{ tables: string[], defaultTable: string }}
 */
function getSchoolTeamTableOptions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (typeof getSchoolTeamTableNames_ === 'function') {
    return getSchoolTeamTableNames_(ss);
  }
  return { tables: ['Schools', 'Teams'], defaultTable: 'Schools' };
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
 * Resolves a sheet from the spreadsheet by exact name, case-insensitive/whitespace-normalized match,
 * or TableApp structured table name mapping.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - Active spreadsheet instance
 * @param {string} sheetName - Target sheet or table name
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null}
 * @private
 */
function resolveSheetByName_(ss, sheetName) {
  if (!ss || !sheetName) return null;

  // 1. Exact match
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) return sheet;

  // 2. Case-insensitive and whitespace-normalized match
  const trimmed = String(sheetName).trim().toLowerCase();
  const normalized = trimmed.replace(/\s+/g, '');
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const sName = sheets[i].getName().trim().toLowerCase();
    if (sName === trimmed || sName.replace(/\s+/g, '') === normalized) {
      return sheets[i];
    }
  }

  // 3. TableApp structured table lookup
  if (typeof TableApp !== 'undefined' && TableApp.openById) {
    try {
      const spreadsheetId = typeof ss.getId === 'function' ? ss.getId() : null;
      if (spreadsheetId) {
        const tableApp = TableApp.openById(spreadsheetId);
        if (tableApp && typeof tableApp.getTableByName === 'function') {
          const table = tableApp.getTableByName(sheetName) || tableApp.getTableByName(normalized);
          if (table && table.sheetName) {
            sheet = ss.getSheetByName(table.sheetName);
            if (sheet) return sheet;
          }
        }
      }
    } catch (e) {
      Logger.log('[resolveSheetByName_] TableApp lookup error: %s', e.message);
    }
  }

  return null;
}

/**
 * Activates the specified sheet in the spreadsheet while a dialog is open.
 * Resolves sheet name with exact match, case-insensitive/whitespace-normalized fallback,
 * or TableApp structured table lookup.
 * @param {string} sheetName - Name of the sheet to activate.
 * @return {boolean} True if successfully activated.
 */
function activateSheetByName(sheetName) {
  if (!sheetName) {
    throw new Error("Sheet name cannot be empty.");
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = resolveSheetByName_(ss, sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found.`);
  }
  if (typeof sheet.isSheetHidden === 'function' && sheet.isSheetHidden()) {
    sheet.showSheet();
  }
  ss.setActiveSheet(sheet);
  if (typeof sheet.activate === 'function') {
    sheet.activate();
  }
  return true;
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
