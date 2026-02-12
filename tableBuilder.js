/**
 * Generic, config‑driven table creation utilities for Google Sheets.
 * Depends on getTableConfig() from tableConfig.js
 */

function getTableAppClient_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();

  if (typeof TableApp !== 'undefined' && TableApp && typeof TableApp.openById === 'function') {
    return TableApp.openById(spreadsheetId);
  }

  throw new Error('TableApp library is not available. Ensure the TableApp library is imported in appsscript.json with userSymbol "TableApp".');
}

function toQuotedSheetA1_(sheetName, a1Range) {
  const escapedSheetName = String(sheetName).replace(/'/g, "''");
  return "'" + escapedSheetName + "'!" + a1Range;
}

/**
 * Creates a named range for a specific column in a table.
 * @param {string} namedRangeName - Name for the named range
 * @param {Sheet} sheet - The sheet containing the table
 * @param {number} startRow - Starting row (data row, not header)
 * @param {number} columnIndex - Column index (1-based)
 * @param {number} numRows - Number of data rows
 */
function createNamedRangeForColumn_(namedRangeName, sheet, startRow, columnIndex, numRows) {
  const ss = SpreadsheetApp.getActive();

  // Remove existing named range if it exists
  const existingRanges = ss.getNamedRanges();
  for (let i = 0; i < existingRanges.length; i++) {
    if (existingRanges[i].getName() === namedRangeName) {
      existingRanges[i].remove();
      Logger.log('[createNamedRangeForColumn_] Removed existing named range: %s', namedRangeName);
      break;
    }
  }

  // Create the range for the column (excluding header)
  const range = sheet.getRange(startRow, columnIndex, numRows, 1);
  ss.setNamedRange(namedRangeName, range);
  Logger.log('[createNamedRangeForColumn_] Created named range "%s" for %s', namedRangeName, range.getA1Notation());
}

function getStructuredTableName_(tableName, tableCfg) {
  if (tableCfg && tableCfg.options && tableCfg.options.structuredTableName) {
    return tableCfg.options.structuredTableName;
  }
  return tableName;
}

function ensureStructuredTable_(sheet, tableName, tableCfg, headerRange, dataRange) {
  const tableApp = getTableAppClient_();
  const structuredTableName = getStructuredTableName_(tableName, tableCfg);
  const startA1 = headerRange.getA1Notation().split(':')[0];
  const endA1 = dataRange.getA1Notation().split(':')[1];
  const fullTableRangeA1 = toQuotedSheetA1_(sheet.getName(), startA1 + ':' + endA1);

  const existing = tableApp.getTableByName(structuredTableName);
  if (existing) {
    existing.setRange(fullTableRangeA1);
    configureTableColumns_(existing, tableCfg);
    return {
      tableName: structuredTableName,
      tableId: existing.getId(),
      action: 'updated',
      rangeA1: fullTableRangeA1
    };
  }

  const created = tableApp.getRange(fullTableRangeA1).create(structuredTableName);
  configureTableColumns_(created, tableCfg);
  return {
    tableName: structuredTableName,
    tableId: created.getId(),
    action: 'created',
    rangeA1: fullTableRangeA1
  };
}

function configureTableColumns_(table, tableCfg) {
  if (!tableCfg || !tableCfg.headers || !tableCfg.columns) return;

  const columnProperties = [];

  for (let i = 0; i < tableCfg.headers.length; i++) {
    const headerName = tableCfg.headers[i];
    const meta = tableCfg.columns[headerName];
    if (!meta) continue;

    const colProp = {
      columnIndex: i,
      columnName: headerName
    };

    // Set column type and validation based on metadata
    if (meta.validation && meta.validation.type === 'list') {
      // Dropdown column with static list
      const values = meta.validation.args && meta.validation.args.values;
      if (values && values.length > 0) {
        colProp.columnType = 'DROPDOWN';
        colProp.dataValidationRule = {
          condition: {
            type: 'ONE_OF_LIST',
            values: values.map(function(v) {
              return { userEnteredValue: String(v) };
            })
          }
        };
      }
    } else if (meta.validation && meta.validation.type === 'range') {
      // Dropdown column with range reference
      const rangeA1 = meta.validation.args && meta.validation.args.rangeA1;
      if (rangeA1) {
        colProp.columnType = 'DROPDOWN';
        colProp.dataValidationRule = {
          condition: {
            type: 'ONE_OF_RANGE',
            values: [{ userEnteredValue: rangeA1 }]
          }
        };
      }
    } else if (meta.validation && meta.validation.type === 'checkbox') {
      // Checkbox column
      colProp.columnType = 'CHECKBOX';
    } else if (meta.type === 'number') {
      // Number column
      colProp.columnType = 'DOUBLE';
    } else if (meta.type === 'date') {
      // Date column
      colProp.columnType = 'DATE';
    } else {
      // If no specific type is set, it defaults to TEXT
      colProp.columnType = 'TEXT';
    }
    columnProperties.push(colProp);
  }

  // Apply column properties to the table
  try {
    table.setColumnProperties(columnProperties, 'columnProperties');
    Logger.log('[configureTableColumns_] Successfully configured %s columns for table "%s"', columnProperties.length, table.getName());
  } catch (e) {
    Logger.log('[configureTableColumns_] Failed to configure columns for table "%s": %s', table.getName(), e.message);
    // Logger.log('[column properties] %s', e);
  }
}

/**
 * Create or place a configured table.
 * @param {string} tableName - Key from configuration (e.g., 'TeamOfficials')
 * @param {Object} [options]
 *  - sheetName: override target sheet name
 *  - startCell: override start cell (A1)
 *  - clearMode: 'rebuild' | 'append'
 *  - schoolYears, genders: optional arrays to override validation lists
 *  - rows: approximate number of rows to pre-prepare for validation (default 50)
 * @return {{tableName:string, sheetName:string, headerA1:string, dataA1:string, rows:number, cols:number}}
 */
function createConfiguredTable(tableName, options) {
  options = options || {};
  const cfg = getTableConfig({ schoolYears: options.schoolYears, genders: options.genders });
  const tableCfg = cfg[tableName];
  if (!tableCfg) {
    throw new Error('Unknown table name: ' + tableName);
  }

  // Resolve placement
  const placement = Object.assign({}, tableCfg.options && tableCfg.options.placement || {});
  if (options.sheetName) placement.targetSheet = options.sheetName;
  if (options.startCell) placement.startCell = options.startCell;
  const sheetName = placement.targetSheet || tableName;
  const startCell = placement.startCell || 'A1';
  const clearMode = options.clearMode || (tableCfg.options && tableCfg.options.clearMode) || 'rebuild';
  const numRows = options.rows || 50;
  const title = options.title ? options.title : tableName;

  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  const startRange = sheet.getRange(startCell);
  const startRow = startRange.getRow();
  const startCol = startRange.getColumn();

  // Prepare headers matrix
  const headers = tableCfg.headers.slice();
  const headerRange = sheet.getRange(startRow, startCol, 1, headers.length);

  if (clearMode === 'rebuild') {
    // Clear the target area (headers + data columns for a reasonable number of rows)
    const clearRange = sheet.getRange(startRow, startCol, numRows + 50, headers.length);
    clearRange.clear({ contentsOnly: true });
    // Note: Do NOT clear validations or set number formats here - TableApp manages these
  }

  // Write headers
  headerRange.setValues([headers]);

  // Apply header formatting
  if (tableCfg.options) {
    if (tableCfg.options.headerBg) headerRange.setBackground(tableCfg.options.headerBg);
    if (typeof tableCfg.options.freezeHeader === 'number') {
      // Freeze the header row relative to sheet
      sheet.setFrozenRows(Math.max(sheet.getFrozenRows(), startRow));
    }
  }

  // Data range (below header)
  const dataStartRow = startRow + 1;
  const dataRange = sheet.getRange(dataStartRow, startCol, numRows, headers.length);

  // Note: Column types and validations are managed by TableApp via configureTableColumns_
  // We only set default values or formulas here if specified
  for (let c = 0; c < headers.length; c++) {
    const header = headers[c];
    const meta = (tableCfg.columns && tableCfg.columns[header]) || {};

    // Apply formula if defined (will be copied down to all rows)
    if (meta.formula) {
      sheet.getRange(dataStartRow, startCol + c, numRows, 1).setFormula(meta.formula);
    }
    // Default values on first row (optional): if defined, pre-fill first row only
    else if (meta.default != null) {
      sheet.getRange(dataStartRow, startCol + c).setValue(meta.default);
    }
  }

  const structuredTable = ensureStructuredTable_(sheet, tableName, tableCfg, headerRange, dataRange);

  // Create named ranges if configured in options
  if (tableCfg.options && tableCfg.options.namedRange) {
    const namedRangeConfig = tableCfg.options.namedRange;
    const columnIndex = headers.indexOf(namedRangeConfig.columnName);

    if (columnIndex !== -1) {
      createNamedRangeForColumn_(namedRangeConfig.name, sheet, dataStartRow, startCol + columnIndex, numRows);
    } else {
      Logger.log('[createConfiguredTable] Warning: Column "%s" not found in table "%s" for named range "%s"',
        namedRangeConfig.columnName, tableName, namedRangeConfig.name);
    }
  }

  return {
    tableName: tableName,
    sheetName: sheetName,
    headerA1: headerRange.getA1Notation(),
    dataA1: dataRange.getA1Notation(),
    rows: numRows,
    cols: headers.length,
    structuredTable: structuredTable
  };
}

/**
 * Convenience: create multiple tables and return a report.
 * @param {string[]} tableNames
 * @param {Object} [options] - passed to each createConfiguredTable
 * @return {Array<Object>} results
 */
function createConfiguredTables(tableNames, options) {
  const results = [];
  tableNames.forEach(function(name) {
    try {
      const result = createConfiguredTable(name, options);
      Logger.log('[createConfiguredTables] Created table "%s" on sheet "%s".', name, result.sheetName);
      results.push(result);
    } catch (err) {
      const message = err && err.message ? err.message : String(err);
      Logger.log('[createConfiguredTables] Failed for "%s": %s', name, message);
      results.push({
        tableName: name,
        error: message
      });
    }
  });
  return results;
}

function getTableNamesForDialog() {
  return listAvailableTables();
}

function createTablesFromDialog(tableNames, options) {
  Logger.log('[createTablesFromDialog] Requested tables: %s, options: %s', JSON.stringify(tableNames), JSON.stringify(options || {}));
  return createConfiguredTables(tableNames, options || {});
}
