/**
 * Generic, config‑driven table creation utilities for Google Sheets.
 * Depends on getTableConfig() from tableConfig.js
 */

/**
 * Internal helper to get a TableApp instance for the current spreadsheet.
 * 
 * @private
 * @return {TableApp} The TableApp instance.
 */
function getTableAppClient_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();

  // Prefer the class if available (for local usage)
  if (typeof TableApp !== 'undefined') {
    return new TableApp(spreadsheetId);
  }

  throw new Error('TableApp library is not available. Ensure TableApp_Local.js is included in your project or the library is imported.');
}

/**
 * Formats a sheet name and A1 range into a quoted A1 notation.
 * 
 * @param {string} sheetName - The name of the sheet.
 * @param {string} a1Range - The A1 range (e.g., "A1:B5").
 * @return {string} Quoted A1 notation (e.g., "'Sheet Name'!A1:B5").
 */
function toQuotedSheetA1_(sheetName, a1Range) {
  const escapedSheetName = String(sheetName).replace(/'/g, "''");
  return "'" + escapedSheetName + "'!" + a1Range;
}

/**
 * Creates a named range for a specific column in a table.
 * 
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

/**
 * Resolves the final table name based on configuration options.
 * 
 * @param {string} tableName - The base table name.
 * @param {Object} tableCfg - The table configuration object.
 * @return {string} The resolved table name.
 */
function getStructuredTableName_(tableName, tableCfg) {
  if (tableCfg && tableCfg.options && tableCfg.options.structuredTableName) {
    return tableCfg.options.structuredTableName;
  }
  return tableName;
}

/**
 * Generates a unique table name if the desired name is already taken.
 * 
 * @param {TableApp} tableApp - The TableApp instance.
 * @param {string} baseName - The desired table name.
 * @return {string} A unique table name.
 */
function getUniqueTableName_(tableApp, baseName) {
  let uniqueName = baseName;
  let counter = 1;
  while (tableApp.getTableByName(uniqueName)) {
    uniqueName = baseName + '_' + counter;
    counter++;
  }
  return uniqueName;
}

/**
 * Ensures a structured table exists in the specified range. Updates if exists, creates if not.
 * 
 * @param {Sheet} sheet - The sheet object.
 * @param {string} tableName - The name for the table.
 * @param {Object} tableCfg - The table configuration object.
 * @param {Range} headerRange - The range containing headers.
 * @param {Range} dataRange - The range containing data.
 * @return {Object} Result metadata about the table created or updated.
 */
function ensureStructuredTable_(sheet, tableName, tableCfg, headerRange, dataRange) {
  const tableApp = getTableAppClient_();
  const structuredTableName = getStructuredTableName_(tableName, tableCfg);
  const startA1 = headerRange.getA1Notation().split(':')[0];
  const endA1 = dataRange.getA1Notation().split(':')[1];
  const fullTableRangeA1 = toQuotedSheetA1_(sheet.getName(), startA1 + ':' + endA1);

  // Ensure all previous SpreadsheetApp changes are visible to the Sheets API
  SpreadsheetApp.flush();

  // Check if we should reuse an existing table
  const existing = tableApp.getTableByName(structuredTableName);
  
  if (existing && existing.sheetName === sheet.getName()) {
    Logger.log('[ensureStructuredTable_] Updating existing table "%s" range to %s', structuredTableName, fullTableRangeA1);
    existing.setRange(fullTableRangeA1);
    configureTableColumns_(existing, tableCfg);
    return {
      tableName: structuredTableName,
      tableId: existing.getId(),
      action: 'updated',
      rangeA1: fullTableRangeA1
    };
  }

  // If it doesn't exist, or exists on another sheet, create a new unique name
  const finalTableName = existing ? getUniqueTableName_(tableApp, structuredTableName) : structuredTableName;

  Logger.log('[ensureStructuredTable_] Creating new table "%s" at range %s', finalTableName, fullTableRangeA1);
  const created = tableApp.getRange(fullTableRangeA1).create(finalTableName);
  configureTableColumns_(created, tableCfg);
  return {
    tableName: finalTableName,
    tableId: created.getId(),
    action: 'created',
    rangeA1: fullTableRangeA1
  };
}

/**
 * Configures table columns (types and validations) based on metadata.
 * 
 * @param {Table} table - The Table instance.
 * @param {Object} tableCfg - The configuration object containing column definitions.
 */
function configureTableColumns_(table, tableCfg) {
  if (!tableCfg || !tableCfg.headers || !tableCfg.columns) {
    Logger.log('[configureTableColumns_] Missing tableCfg data: tableCfg=%s, headers=%s, columns=%s',
      !!tableCfg, !!(tableCfg && tableCfg.headers), !!(tableCfg && tableCfg.columns));
    return;
  }

  Logger.log('[configureTableColumns_] Configuring columns for table. Headers: %s', JSON.stringify(tableCfg.headers));

  const columnProperties = [];

  for (let i = 0; i < tableCfg.headers.length; i++) {
    const headerName = tableCfg.headers[i];
    const meta = tableCfg.columns[headerName];
    if (!meta) {
      Logger.log('[configureTableColumns_] No metadata for column "%s" at index %s', headerName, i);
      continue;
    }

    const colProp = {
      columnIndex: i,
      columnName: headerName
    };

    Logger.log('[configureTableColumns_] Processing column "%s": type=%s, validation=%s',
      headerName, meta.type, JSON.stringify(meta.validation));

    if (meta.validation && meta.validation.type === 'list') {
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
        Logger.log('[configureTableColumns_] Column "%s": DROPDOWN with %s values', headerName, values.length);
      }
    } else if (meta.validation && meta.validation.type === 'range') {
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
      colProp.columnType = 'CHECKBOX';
    } else if (meta.type === 'number') {
      colProp.columnType = 'DOUBLE';
    } else if (meta.type === 'date') {
      colProp.columnType = 'DATE';
    } else {
      colProp.columnType = 'TEXT';
    }
    columnProperties.push(colProp);
  }

  Logger.log('[configureTableColumns_] Applying %s column properties to table "%s"', columnProperties.length, table.getName());

  try {
    table.setColumnProperties(columnProperties, 'columnProperties');
    Logger.log('[configureTableColumns_] ✓ Successfully configured %s columns for table "%s"', columnProperties.length, table.getName());
  } catch (e) {
    Logger.log('[configureTableColumns_] ✗ Failed to configure columns for table "%s": %s', table.getName(), e.message);
    Logger.log('[configureTableColumns_] Error stack: %s', e.stack);
  }
}

/**
 * Create or place a configured table.
 * 
 * @param {string} tableName - Key from configuration (e.g., 'TeamOfficials')
 * @param {Object} [options]
 *  - sheetName: override target sheet name
 *  - startCell: override start cell (A1)
 *  - clearMode: 'rebuild' | 'append'
 *  - schoolYears, genders: optional arrays to override validation lists
 *  - rows: approximate number of rows to pre-prepare for validation (default 50)
 * @return {Object} Placement and structure details.
 */
function createConfiguredTable(tableName, options) {
  options = options || {};
  const cfg = getTableConfig({ schoolYears: options.schoolYears, genders: options.genders });
  const tableCfg = cfg.tables[tableName];
  if (!tableCfg) {
    throw new Error('Unknown table name: ' + tableName);
  }

  const placement = Object.assign({}, tableCfg.options && tableCfg.options.placement || {});
  if (options.sheetName) placement.targetSheet = options.sheetName;
  if (options.startCell) placement.startCell = options.startCell;
  const sheetName = placement.targetSheet || tableName;
  const startCell = placement.startCell || 'A1';
  const clearMode = options.clearMode || (tableCfg.options && tableCfg.options.clearMode) || 'rebuild';
  const numRows = options.rows || (tableCfg.options && tableCfg.options.rows) || 50;

  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    Logger.log('[createConfiguredTable] Created new sheet "%s"', sheetName);
    // CRITICAL: Flush after insertSheet so the Sheets API (v4) can see the new sheetId
    SpreadsheetApp.flush();
  }

  const startRange = sheet.getRange(startCell);
  const startRow = startRange.getRow();
  const startCol = startRange.getColumn();

  const headers = tableCfg.headers.slice();
  const headerRange = sheet.getRange(startRow, startCol, 1, headers.length);

  if (clearMode === 'rebuild') {
    Logger.log('[createConfiguredTable] Clearing target area for "%s"', tableName);
    const clearRange = sheet.getRange(startRow, startCol, numRows + 50, headers.length);
    clearRange.clear({ contentsOnly: true });
  }

  headerRange.setValues([headers]);

  if (tableCfg.options) {
    if (tableCfg.options.headerBg) headerRange.setBackground(tableCfg.options.headerBg);
    if (typeof tableCfg.options.freezeHeader === 'number') {
      sheet.setFrozenRows(Math.max(sheet.getFrozenRows(), startRow));
    }
  }

  const dataStartRow = startRow + 1;
  const dataRange = sheet.getRange(dataStartRow, startCol, numRows, headers.length);

  for (let c = 0; c < headers.length; c++) {
    const header = headers[c];
    const meta = (tableCfg.columns && tableCfg.columns[header]) || {};
    if (meta.formula) {
      sheet.getRange(dataStartRow, startCol + c, numRows, 1).setFormula(meta.formula);
    } else if (meta.default != null) {
      sheet.getRange(dataStartRow, startCol + c).setValue(meta.default);
    }
  }

  // Ensure all SpreadsheetApp changes are flushed before handing off to TableApp (Sheets API)
  SpreadsheetApp.flush();

  const structuredTable = ensureStructuredTable_(sheet, tableName, tableCfg, headerRange, dataRange);

  if (tableCfg.options && tableCfg.options.namedRange) {
    const namedRangeConfig = tableCfg.options.namedRange;
    const columnIndex = headers.indexOf(namedRangeConfig.columnName);
    if (columnIndex !== -1) {
      createNamedRangeForColumn_(namedRangeConfig.name, sheet, dataStartRow, startCol + columnIndex, numRows);
    } else {
      Logger.log('[createConfiguredTable] Warning: Column "%s" not found for named range "%s"',
        namedRangeConfig.columnName, namedRangeConfig.name);
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
 * 
 * @param {string[]} tableNames - Array of table names to create.
 * @param {Object} [options] - Options passed to createConfiguredTable.
 * @return {Array<Object>} Results for each table.
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

/**
 * Lists available tables for dialog selection.
 * 
 * @return {string[]} List of table names.
 */
function getTableNamesForDialog() {
  return listAvailableTables();
}

/**
 * Creates multiple tables from dialog input.
 * 
 * @param {string[]} tableNames - Array of table names.
 * @param {Object} [options] - Creation options.
 * @return {Array<Object>} Results for each table.
 */
function createTablesFromDialog(tableNames, options) {
  Logger.log('[createTablesFromDialog] Requested tables: %s, options: %s', JSON.stringify(tableNames), JSON.stringify(options || {}));
  return createConfiguredTables(tableNames, options || {});
}

/**
 * Reads team names from a specified range in a given sheet.
 * 
 * @param {string} sheetName - Name of the sheet to read from
 * @param {string} rangeA1 - A1 notation range (e.g. "A2:A8")
 * @return {string[]} Array of team names (filters blanks)
 */
function getTeamNamesFromSheet(sheetName, rangeA1) {
  Logger.log('[getTeamNamesFromSheet] Reading from sheet "%s", range %s', sheetName, rangeA1);
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet "' + sheetName + '" does not exist');

  const values = sheet.getRange(rangeA1).getValues();
  const teamNames = values.flat()
    .map(function(name) { return String(name).trim(); })
    .filter(function(name) { return name !== ''; });

  Logger.log('[getTeamNamesFromSheet] Found %s team names', teamNames.length);
  return teamNames;
}

/**
 * Sanitizes a string for use as a structured table name.
 * 
 * @param {string} label - Raw event label
 * @return {string} Sanitized table name
 */
function sanitizeTableName_(label) {
  let sanitized = String(label)
    .replace(/[^A-Za-z0-9_ -]/g, '')
    .replace(/\s+/g, '_')
    .substring(0, 50);

  if (sanitized && /^\d/.test(sanitized)) {
    sanitized = 'R_' + sanitized;
  }
  return sanitized;
}

/**
 * Creates a single relay table with label row, header, and data rows.
 * 
 * @param {Object} eventConfig - { label: string }
 * @param {Sheet} sheetTarget - Target sheet object
 * @param {number} startRow - Row to start the label (1-based)
 * @param {number} startCol - Column to start (1-based)
 * @param {string[]} teamNames - Array of team names to use as columns
 * @param {Object} relayTableCfg - The relay table config from getTableConfig
 * @param {Object} relayDefaults - The relayDefaults config
 * @return {Object} { labelRow, headerRow, dataEndRow, cols, structuredTable }
 */
function createRelayTable_(eventConfig, sheetTarget, startRow, startCol, teamNames, relayTableCfg, relayDefaults) {
  Logger.log('[createRelayTable_] Creating relay table "%s" at row %s', eventConfig.label, startRow);
  const labelRow = startRow;
  const headerRow = startRow + 1;
  const dataStartRow = headerRow + 1;
  const rowsPerTable = relayDefaults.defaultRowsPerTable;
  const dataEndRow = dataStartRow + rowsPerTable - 1;

  const fullHeaders = relayTableCfg.headers.slice().concat(teamNames);
  const totalCols = fullHeaders.length;

  const labelRange = sheetTarget.getRange(labelRow, startCol, 1, totalCols);
  labelRange.merge();
  labelRange.setValue(eventConfig.label);
  labelRange.setBackground('#356853');
  labelRange.setFontColor('#ffffff');
  labelRange.setFontWeight('bold');
  labelRange.setHorizontalAlignment('center');

  const headerRange = sheetTarget.getRange(headerRow, startCol, 1, totalCols);
  headerRange.setValues([fullHeaders]);
  if (relayTableCfg.options && relayTableCfg.options.headerBg) {
    headerRange.setBackground(relayTableCfg.options.headerBg);
  }
  headerRange.setFontWeight('bold');
  headerRange.setFontColor('#ffffff');

  const dataRange = sheetTarget.getRange(dataStartRow, startCol, rowsPerTable, totalCols);

  for (let c = 0; c < relayTableCfg.headers.length; c++) {
    const header = relayTableCfg.headers[c];
    const meta = relayTableCfg.columns[header];
    if (meta && meta.default != null) {
      sheetTarget.getRange(dataStartRow, startCol + c).setValue(meta.default);
    }
  }

  // Ensure all SpreadsheetApp changes are flushed before handing off to TableApp (Sheets API)
  SpreadsheetApp.flush();

  const sanitizedName = sanitizeTableName_(eventConfig.label);
  const structuredTable = ensureStructuredTable_(sheetTarget, sanitizedName, relayTableCfg, headerRange, dataRange);

  return {
    labelRow: labelRow,
    headerRow: headerRow,
    dataEndRow: dataEndRow,
    cols: totalCols,
    structuredTable: structuredTable
  };
}

/**
 * Creates multiple relay tables based on configuration.
 * 
 * @param {Object} config - {
 *   sourceSheet: string,
 *   sourceRange: string (A1 notation, e.g. "A2:A8"),
 *   events: [{ label: string }],
 *   years: [string] (optional),
 *   placement: { sameSheet: boolean, sheetName: string, startCell: string }
 * }
 * @return {Array<Object>} Results for each relay table created.
 */
function createRelayTables(config) {
  Logger.log('[createRelayTables] Starting relay table creation');
  try {
    const teamNames = getTeamNamesFromSheet(config.sourceSheet, config.sourceRange);
    if (teamNames.length === 0) {
      const cfg = getTableConfig();
      Logger.log('[createRelayTables] No team names found, using defaults');
      teamNames.push.apply(teamNames, cfg.defaultClusters);
    }

    const overrides = {};
    if (config.years && config.years.length > 0) {
      overrides.schoolYears = config.years;
    }
    const cfg = getTableConfig(overrides);

    let relayTableCfg = null;
    for (let tableName in cfg.tables) {
      if (cfg.tables[tableName].tableType === 'relay') {
        relayTableCfg = cfg.tables[tableName];
        break;
      }
    }
    if (!relayTableCfg) throw new Error('No relay table configuration found');

    const relayDefaults = cfg.relayDefaults;
    const ss = SpreadsheetApp.getActive();
    const results = [];

    if (config.placement.sameSheet) {
      const sheetName = config.placement.sheetName || 'Relays';
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        SpreadsheetApp.flush();
      }

      const startRange = sheet.getRange(config.placement.startCell || 'A1');
      let currentRow = startRange.getRow();
      const startCol = startRange.getColumn();

      for (let i = 0; i < config.events.length; i++) {
        const event = config.events[i];
        try {
          const tableResult = createRelayTable_(event, sheet, currentRow, startCol, teamNames, relayTableCfg, relayDefaults);
          results.push({
            eventLabel: event.label,
            sanitisedTableName: sanitizeTableName_(event.label),
            sheetName: sheetName,
            labelRow: tableResult.labelRow,
            headerRow: tableResult.headerRow,
            dataRange: sheet.getRange(tableResult.headerRow + 1, startCol, relayDefaults.defaultRowsPerTable, tableResult.cols).getA1Notation(),
            structuredTable: tableResult.structuredTable
          });
          currentRow = tableResult.dataEndRow + 1 + relayDefaults.gapBetweenTables;
        } catch (err) {
          Logger.log('[createRelayTables] Failed for event "%s": %s', event.label, err.message);
          results.push({ eventLabel: event.label, error: err.message });
        }
      }
    } else {
      for (let i = 0; i < config.events.length; i++) {
        const event = config.events[i];
        const sheetName = sanitizeTableName_(event.label);
        try {
          let sheet = ss.getSheetByName(sheetName);
          if (!sheet) {
            sheet = ss.insertSheet(sheetName);
            SpreadsheetApp.flush();
          }
          const tableResult = createRelayTable_(event, sheet, 1, 1, teamNames, relayTableCfg, relayDefaults);
          results.push({
            eventLabel: event.label,
            sanitisedTableName: sanitizeTableName_(event.label),
            sheetName: sheetName,
            labelRow: tableResult.labelRow,
            headerRow: tableResult.headerRow,
            dataRange: sheet.getRange(tableResult.headerRow + 1, 1, relayDefaults.defaultRowsPerTable, tableResult.cols).getA1Notation(),
            structuredTable: tableResult.structuredTable
          });
        } catch (err) {
          Logger.log('[createRelayTables] Failed for event "%s": %s', event.label, err.message);
          results.push({ eventLabel: event.label, error: err.message });
        }
      }
    }
    return results;
  } catch (err) {
    Logger.log('[createRelayTables] Fatal error: %s', err.message);
    throw err;
  }
}

/**
 * Wrapper function called from the dialog to create relay tables.
 * 
 * @param {Object} config - Same as createRelayTables
 * @return {Array<Object>} Results
 */
function createRelayTablesFromDialog(config) {
  Logger.log('[createRelayTablesFromDialog] Called');
  return createRelayTables(config);
}
