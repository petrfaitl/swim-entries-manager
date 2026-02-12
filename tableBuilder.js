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
    return {
      tableName: structuredTableName,
      tableId: existing.getId(),
      action: 'updated',
      rangeA1: fullTableRangeA1
    };
  }

  const created = tableApp.getRange(fullTableRangeA1).create(structuredTableName);
  return {
    tableName: structuredTableName,
    tableId: created.getId(),
    action: 'created',
    rangeA1: fullTableRangeA1
  };
}

/**
 * Create or place a configured table.
 * @param {string} tableName - Key from configuration (e.g., 'TeamOfficials')
 * @param {Object} [options]
 *  - sheetName: override target sheet name
 *  - startCell: override start cell (A1)
 *  - clearMode: 'rebuild' | 'append'
 *  - schoolYears, genders: optional arrays to override validation lists
 *  - rows: approximate number of rows to pre-prepare for validation (default 500)
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
  const numRows = options.rows || 500;

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
    // Also reset validations and formats for this block
    clearRange.clearDataValidations();
    clearRange.setNumberFormat('@'); // default to plain text; specific columns will override
  }

  // Write headers
  headerRange.setValues([headers]);

  // Apply header formatting
  if (tableCfg.options) {
    if (tableCfg.options.boldHeader) headerRange.setFontWeight('bold');
    if (tableCfg.options.headerBg) headerRange.setBackground(tableCfg.options.headerBg);
    if (typeof tableCfg.options.freezeHeader === 'number') {
      // Freeze the header row relative to sheet
      sheet.setFrozenRows(Math.max(sheet.getFrozenRows(), startRow));
    }
  }

  // Add alternating colors (basic: apply to a banded range covering a reasonable block)
  if (tableCfg.options && tableCfg.options.alternatingColors) {
    const bandRange = sheet.getRange(startRow, startCol, Math.max(2, Math.min(numRows + 1, sheet.getMaxRows() - startRow + 1)), headers.length);
    const banding = bandRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
    // Ensure header styled distinctively
    banding.setFirstRowColor(headerRange.getBackground());
  }

  // Data range (below header)
  const dataStartRow = startRow + 1;
  const dataRange = sheet.getRange(dataStartRow, startCol, numRows, headers.length);

  // Apply per-column types and validations
  for (let c = 0; c < headers.length; c++) {
    const header = headers[c];
    const meta = (tableCfg.columns && tableCfg.columns[header]) || {};
    const colRange = sheet.getRange(dataStartRow, startCol + c, numRows, 1);

    // Clear existing validation for safety if not rebuild
    colRange.clearDataValidations();

    // Set number formats for some known types
    if (meta.type === 'number') {
      colRange.setNumberFormat('0');
    } else if (meta.type === 'date') {
      colRange.setNumberFormat('dd/mm/yyyy');
    } else if (meta.type === 'checkbox' && meta.validation && meta.validation.type === 'checkbox') {
      const args = meta.validation.args || {};
      colRange.insertCheckboxes(args.checkedValue || 'TRUE', args.uncheckedValue || 'FALSE');
    } else {
      // default plain text
      colRange.setNumberFormat('@');
    }

    // Validation rules
    if (meta.validation) {
      const v = meta.validation;
      if (v.type === 'list') {
        const args = v.args || {};
        if (args.values && args.values.length) {
          const rule = SpreadsheetApp.newDataValidation()
            .requireValueInList(args.values, true)
            .setAllowInvalid(false)
            .build();
          colRange.setDataValidation(rule);
        } else if (args.rangeA1) {
          const rule = SpreadsheetApp.newDataValidation()
            .requireValueInRange(sheet.getRange(args.rangeA1), true)
            .setAllowInvalid(false)
            .build();
          colRange.setDataValidation(rule);
        }
      } else if (v.type === 'numberRange') {
        const args = v.args || {};
        const builder = SpreadsheetApp.newDataValidation().setAllowInvalid(!!args.allowInvalid);
        if (typeof args.min === 'number' && typeof args.max === 'number') {
          builder.requireNumberBetween(args.min, args.max);
        } else if (typeof args.min === 'number') {
          builder.requireNumberGreaterThanOrEqualTo(args.min);
        } else if (typeof args.max === 'number') {
          builder.requireNumberLessThanOrEqualTo(args.max);
        } else {
          builder.requireNumber();
        }
        colRange.setDataValidation(builder.build());
      } else if (v.type === 'date') {
        const args = v.args || {};
        const builder = SpreadsheetApp.newDataValidation();
        if (args.condition === 'DATE_BEFORE' && args.value) {
          builder.requireDateBefore(args.value);
        } else if (args.condition === 'DATE_AFTER' && args.value) {
          builder.requireDateAfter(args.value);
        } else {
          builder.requireDate();
        }
        colRange.setDataValidation(builder.setAllowInvalid(false).build());
      } else if (v.type === 'checkbox') {
        // handled above by insertCheckboxes
      }
    }

    // Default values on first row (optional): if defined, pre-fill first row only
    if (meta.default != null) {
      sheet.getRange(dataStartRow, startCol + c).setValue(meta.default);
    }
  }

  // Named range (optional)
  if (tableCfg.options && tableCfg.options.namedRange) {
    const existing = ss.getRangeByName(tableCfg.options.namedRange);
    const headerA1 = headerRange.getA1Notation();
    const dataA1 = dataRange.getA1Notation();
    const a1 = sheet.getName() + '!' + headerA1.replace(/:.+$/, '') + ':' + dataA1.split('!').pop().split(':').pop();
    if (existing) ss.removeNamedRange(tableCfg.options.namedRange);
    ss.setNamedRange(tableCfg.options.namedRange, sheet.getRange(headerA1 + ':' + dataRange.getA1Notation().split('!').pop().split(':').pop()));
  }

  const structuredTable = ensureStructuredTable_(sheet, tableName, tableCfg, headerRange, dataRange);

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
  tableNames.forEach(name => {
    results.push(createConfiguredTable(name, options));
  });
  return results;
}

function getTableNamesForDialog() {
  return listAvailableTables();
}

function createTablesFromDialog(tableNames, options) {
  return createConfiguredTables(tableNames, options || {});
}
