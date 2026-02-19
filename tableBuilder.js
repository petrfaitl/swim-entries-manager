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
  const tableCfg = cfg.tables[tableName];
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
  const numRows = options.rows || (tableCfg.options && tableCfg.options.rows) || 50;
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

/**
 * Reads team names from a specified range in a given sheet.
 * @param {string} sheetName - Name of the sheet to read from
 * @param {string} rangeA1 - A1 notation range (e.g. "A2:A8")
 * @return {string[]} Array of team names (filters blanks)
 */
function getTeamNamesFromSheet(sheetName, rangeA1) {
  Logger.log('[getTeamNamesFromSheet] Reading from sheet "%s", range %s', sheetName, rangeA1);

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('Sheet "' + sheetName + '" does not exist');
  }

  // Get values from the specified range
  const values = sheet.getRange(rangeA1).getValues();
  const teamNames = values
    .flat()
    .map(function(name) { return String(name).trim(); })
    .filter(function(name) { return name !== ''; });

  Logger.log('[getTeamNamesFromSheet] Found %s team names: %s', teamNames.length, JSON.stringify(teamNames));
  return teamNames;
}

/**
 * Sanitizes a string for use as a structured table name.
 * Strips characters not allowed: keeps only A-Z a-z 0-9 _ -
 * Replaces spaces with underscores, truncates to 50 chars.
 * @param {string} label - Raw event label
 * @return {string} Sanitized table name
 */
function sanitizeTableName_(label) {
  return String(label)
    .replace(/[^A-Za-z0-9_ -]/g, '')  // Remove invalid chars
    .replace(/\s+/g, '_')              // Replace spaces with underscores
    .substring(0, 50);                 // Truncate to 50 chars
}

/**
 * Creates a single relay table with label row, header, and data rows.
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
  Logger.log('[createRelayTable_] Creating relay table "%s" at row %s, col %s', eventConfig.label, startRow, startCol);

  const labelRow = startRow;
  const headerRow = startRow + 1;
  const dataStartRow = headerRow + 1;
  const rowsPerTable = relayDefaults.defaultRowsPerTable;
  const dataEndRow = dataStartRow + rowsPerTable - 1;

  // Build full header: fixed columns + team names
  const fullHeaders = relayTableCfg.headers.slice().concat(teamNames);
  const totalCols = fullHeaders.length;

  // 1. Write label row (merged cell with event name)
  const labelRange = sheetTarget.getRange(labelRow, startCol, 1, totalCols);
  labelRange.merge();
  labelRange.setValue(eventConfig.label);
  labelRange.setBackground('#1c4587');
  labelRange.setFontColor('#ffffff');
  labelRange.setFontWeight('bold');
  labelRange.setHorizontalAlignment('center');

  // 2. Write header row
  const headerRange = sheetTarget.getRange(headerRow, startCol, 1, totalCols);
  headerRange.setValues([fullHeaders]);
  if (relayTableCfg.options && relayTableCfg.options.headerBg) {
    headerRange.setBackground(relayTableCfg.options.headerBg);
  }
  headerRange.setFontWeight('bold');
  headerRange.setFontColor('#ffffff');

  // 3. Write data rows for fixed columns only (Order, School Year, Gender)
  const dataRange = sheetTarget.getRange(dataStartRow, startCol, rowsPerTable, totalCols);

  // Pre-fill fixed column data with validation (Order, School Year, Gender)
  for (let c = 0; c < relayTableCfg.headers.length; c++) {
    const header = relayTableCfg.headers[c];
    const meta = relayTableCfg.columns[header];

    if (meta && meta.default != null) {
      // Set default value in first row if defined
      sheetTarget.getRange(dataStartRow, startCol + c).setValue(meta.default);
    }
  }

  // 4. Create structured table
  const sanitizedName = sanitizeTableName_(eventConfig.label);
  const structuredTable = ensureStructuredTable_(sheetTarget, sanitizedName, relayTableCfg, headerRange, dataRange);

  Logger.log('[createRelayTable_] Created relay table "%s" (structured name: "%s")', eventConfig.label, sanitizedName);

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
 * @param {Object} config - {
 *   sourceSheet: string,
 *   sourceRange: string (A1 notation, e.g. "A2:A8"),
 *   events: [{ label: string }],
 *   years: [string] (optional, e.g. ['Y5', 'Y6', ...]),
 *   placement: {
 *     sameSheet: boolean,
 *     sheetName: string,
 *     startCell: string (A1 notation)
 *   }
 * }
 * @return {Array} Results for each relay table created
 */
function createRelayTables(config) {
  Logger.log('[createRelayTables] Starting relay table creation with config: %s', JSON.stringify(config));

  try {
    // 1. Get team names from source sheet
    const teamNames = getTeamNamesFromSheet(config.sourceSheet, config.sourceRange);

    if (teamNames.length === 0) {
      // Fallback to DEFAULT_CLUSTERS
      const cfg = getTableConfig();
      Logger.log('[createRelayTables] No team names found, using default clusters');
      teamNames.push.apply(teamNames, cfg.defaultClusters);
    }

    // 2. Get relay table config with custom years if provided
    const overrides = {};
    if (config.years && config.years.length > 0) {
      overrides.schoolYears = config.years;
    }
    const cfg = getTableConfig(overrides);
    let relayTableCfg = null;

    // Find the table with tableType === 'relay'
    for (let tableName in cfg.tables) {
      if (cfg.tables[tableName].tableType === 'relay') {
        relayTableCfg = cfg.tables[tableName];
        break;
      }
    }

    if (!relayTableCfg) {
      throw new Error('No relay table configuration found (tableType: relay)');
    }

    const relayDefaults = cfg.relayDefaults;
    const ss = SpreadsheetApp.getActive();
    const results = [];

    // 3. Process each event
    if (config.placement.sameSheet) {
      // All tables on one sheet
      const sheetName = config.placement.sheetName || 'Relays';
      let sheet = ss.getSheetByName(sheetName);

      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        Logger.log('[createRelayTables] Created new sheet "%s"', sheetName);
      }

      // Parse start cell
      const startRange = sheet.getRange(config.placement.startCell || 'A1');
      let currentRow = startRange.getRow();
      const startCol = startRange.getColumn();

      for (let i = 0; i < config.events.length; i++) {
        const event = config.events[i];

        try {
          const tableResult = createRelayTable_(
            event,
            sheet,
            currentRow,
            startCol,
            teamNames,
            relayTableCfg,
            relayDefaults
          );

          results.push({
            eventLabel: event.label,
            sanitisedTableName: sanitizeTableName_(event.label),
            sheetName: sheetName,
            labelRow: tableResult.labelRow,
            headerRow: tableResult.headerRow,
            dataRange: sheet.getRange(tableResult.headerRow + 1, startCol, relayDefaults.defaultRowsPerTable, tableResult.cols).getA1Notation(),
            structuredTable: tableResult.structuredTable
          });

          // Advance currentRow: label + header + data rows + gap
          currentRow = tableResult.dataEndRow + 1 + relayDefaults.gapBetweenTables;

        } catch (err) {
          Logger.log('[createRelayTables] Failed to create table for event "%s": %s', event.label, err.message);
          results.push({
            eventLabel: event.label,
            error: err.message
          });
        }
      }

    } else {
      // Separate sheet per event
      for (let i = 0; i < config.events.length; i++) {
        const event = config.events[i];
        const sheetName = sanitizeTableName_(event.label);

        try {
          let sheet = ss.getSheetByName(sheetName);

          if (!sheet) {
            sheet = ss.insertSheet(sheetName);
            Logger.log('[createRelayTables] Created new sheet "%s"', sheetName);
          }

          const tableResult = createRelayTable_(
            event,
            sheet,
            1,  // Start at A1
            1,
            teamNames,
            relayTableCfg,
            relayDefaults
          );

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
          Logger.log('[createRelayTables] Failed to create table for event "%s": %s', event.label, err.message);
          results.push({
            eventLabel: event.label,
            error: err.message
          });
        }
      }
    }

    Logger.log('[createRelayTables] Completed: %s tables created', results.length);
    return results;

  } catch (err) {
    Logger.log('[createRelayTables] Fatal error: %s', err.message);
    throw err;
  }
}

/**
 * Wrapper function called from the dialog.
 * @param {Object} config - Same as createRelayTables
 * @return {Array} Results
 */
function createRelayTablesFromDialog(config) {
  Logger.log('[createRelayTablesFromDialog] Called with config: %s', JSON.stringify(config));
  return createRelayTables(config);
}
