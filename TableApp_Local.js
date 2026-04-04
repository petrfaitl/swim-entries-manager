/**
 * TableApp: TableApp is a Google Apps Script library for managing Tables on Google Sheets.
 * (required) Sheets API
 * Author: Kanshi Tanaike
 * https://github.com/tanaikech/TableApp
 *
 * Updated on 20251120 11:45
 * version 1.0.1 - Improved compatibility with 'currentonly' scope and structured tables.
 */

/**
 * Opens the TableApp for a specific Spreadsheet.
 *
 * @param {string} spreadsheetId The Spreadsheet ID.
 * @return {TableApp} The TableApp instance.
 */
function openById(spreadsheetId) {
  return new TableApp(spreadsheetId);
}

/**
 * Class for managing Tables within a Google Spreadsheet.
 * Acts as the main entry point for creating and retrieving tables.
 */
class TableApp {
  /**
   * Opens the TableApp for a specific Spreadsheet.
   *
   * @param {string} spreadsheetId The Spreadsheet ID.
   * @return {TableApp} The TableApp instance.
   */
  static openById(spreadsheetId) {
    return new TableApp(spreadsheetId);
  }

  /**
   * @param {string} spreadsheetId The Spreadsheet ID.
   */
  constructor(spreadsheetId) {
    /** @private @type {string} */
    this.spreadsheetId = spreadsheetId || null;
    /** @private @type {string|null} */
    this.sheetName = null;
    /** @private @type {string|null} */
    this.a1Notation = null;
    /** @private @type {Object|null} */
    this.cachedTables = null;
  }

  /**
   * Internal helper to get a valid Spreadsheet ID, defaulting to active if needed.
   * @private
   */
  _getSid() {
    if (!this.spreadsheetId) {
      try {
        this.spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
      } catch (e) {
        throw new Error("Spreadsheet ID is missing and no active spreadsheet found.");
      }
    }
    return this.spreadsheetId;
  }

  /**
   * Sets the target sheet by name for subsequent operations (like create).
   *
   * @param {string} sheetName The name of the sheet.
   * @return {TableApp} This instance for chaining.
   */
  getSheetByName(sheetName) {
    this.sheetName = sheetName;
    return this;
  }

  /**
   * Sets the target range using A1 notation for creating a table.
   *
   * @param {string} a1Notation The range in A1 notation (e.g., "Sheet1!A1:B5" or "A1:B5").
   * @return {TableApp} This instance for chaining.
   */
  getRange(a1Notation) {
    const parsed = parseA1Notation_(a1Notation);
    if (parsed && parsed.sheetName) {
      this.sheetName = parsed.sheetName;
    }
    this.a1Notation = a1Notation;
    return this;
  }

  /**
   * Creates a new table in the defined range/sheet.
   *
   * @param {string} tableName The name of the new table.
   * @return {Table} The created Table instance.
   */
  create(tableName) {
    const { gridRange, sheetName, sheetId } = this._resolveGridRange();

    const requests = [
      {
        addTable: {
          table: {
            name: tableName,
            range: gridRange,
          },
        },
      },
    ];

    const sid = this._getSid();
    // Flush changes before batchUpdate to ensure new sheets/ranges are recognized
    SpreadsheetApp.flush();
    
    const response = batchUpdate_(sid, requests);
    const newTableObj = response.replies[0].addTable.table;

    this.cachedTables = null;

    return new Table({
      spreadsheetId: sid,
      sheetName: sheetName,
      sheetId: sheetId,
      table: newTableObj,
    });
  }

  /**
   * Retrieves all tables.
   */
  getTables() {
    const data = this._getOrFetchTables();
    return this.sheetName
      ? data.tablesBySheetNames[this.sheetName]?.tables || []
      : data.tablesBySheetNames;
  }

  /**
   * Gets a specific table by its name.
   */
  getTableByName(tableName) {
    const data = this._getOrFetchTables();
    return data.tablesByTableNames[tableName] || null;
  }

  /**
   * Gets a specific table by its ID.
   */
  getTableById(tableId) {
    const data = this._getOrFetchTables();
    return data.tablesByTableIds[tableId] || null;
  }

  /**
   * Internal method to fetch tables with caching.
   * @private
   */
  _getOrFetchTables() {
    if (!this.cachedTables) {
      this.cachedTables = fetchAllTables_(this._getSid());
    }
    return this.cachedTables;
  }

  /**
   * Internal method to resolve A1 notation to GridRange and Sheet properties.
   * Optimized for 'currentonly' scope.
   * @private
   */
  _resolveGridRange() {
    const notation = this.a1Notation || "A1";
    const parsed = parseA1Notation_(notation);

    let targetSheetName = (parsed && parsed.sheetName) ? parsed.sheetName : this.sheetName;

    // Use SpreadsheetApp for sheet metadata to comply with 'currentonly' scope
    // We prefer the active spreadsheet if possible, or open the specific ID
    let ss;
    const sid = this._getSid();
    try {
      const active = SpreadsheetApp.getActiveSpreadsheet();
      if (active && active.getId() === sid) {
        ss = active;
      } else {
        ss = SpreadsheetApp.openById(sid);
      }
    } catch (e) {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    }
    
    const sheets = ss.getSheets();
    let targetSheetId = null;

    if (!targetSheetName) {
      const firstSheet = sheets[0];
      targetSheetName = firstSheet.getName();
      targetSheetId = firstSheet.getSheetId();
    } else {
      const found = sheets.find((s) => s.getName() === targetSheetName);
      if (!found) throw new Error(`Sheet with name "${targetSheetName}" not found.`);
      targetSheetId = found.getSheetId();
    }

    const gridRange = convA1NotationToGridRange_(notation, targetSheetId);
    return { gridRange, sheetName: targetSheetName, sheetId: targetSheetId };
  }
}

/**
 * Class representing a single Table in Google Sheets.
 */
class Table {
  constructor(obj) {
    this.spreadsheetId = obj.spreadsheetId;
    this.sheetName = obj.sheetName;
    this.sheetId = obj.sheetId;
    this.table = obj.table;
    this.rangeAsA1Notation = convGridRangeToA1Notation_(this.table.range, this.sheetName);
  }

  getName() { return this.table.name; }
  getId() { return this.table.tableId; }
  getMetadata() { return this.table; }
  getRange() { return convGridRangeToA1Notation_(this.table.range, this.sheetName); }

  getValues() {
    return valuesGet_(this.spreadsheetId, this.rangeAsA1Notation);
  }

  setName(tableName) {
    if (!tableName || typeof tableName !== "string") throw new Error("Invalid table name.");
    const requests = [{
      updateTable: {
        fields: "name",
        table: { name: tableName, tableId: this.table.tableId },
      },
    }];
    this._updateTable(requests);
    this.table.name = tableName;
    return this;
  }

  setValues(values) {
    if (!Array.isArray(values) || values.length === 0 || !Array.isArray(values[0])) {
      throw new Error("Invalid values. Must be a 2D array.");
    }
    return valuesUpdate_(this.spreadsheetId, values, this.rangeAsA1Notation);
  }

  setRange(a1Notation) {
    if (!a1Notation || typeof a1Notation !== "string") throw new Error("Invalid A1 Notation.");
    const gridRange = convA1NotationToGridRange_(a1Notation, this.sheetId);
    const requests = [{
      updateTable: {
        fields: "range",
        table: { range: gridRange, tableId: this.table.tableId },
      },
    }];
    this._updateTable(requests);
    this.table.range = gridRange;
    this.rangeAsA1Notation = convGridRangeToA1Notation_(gridRange, this.sheetName);
    return this;
  }

  setRowsProperties(rowsProperties, fields = "rowsProperties") {
    const requests = [{
      updateTable: {
        fields,
        table: { rowsProperties, tableId: this.table.tableId },
      },
    }];
    return this._updateTable(requests);
  }

  setColumnProperties(columnProperties, fields = "columnProperties") {
    const requests = [{
      updateTable: {
        fields,
        table: { columnProperties, tableId: this.table.tableId },
      },
    }];
    return this._updateTable(requests);
  }

  remove() {
    const requests = [{ deleteTable: { tableId: this.table.tableId } }];
    try {
      batchUpdate_(this.spreadsheetId, requests);
    } catch (e) {
      throw new Error(`Failed to remove table: ${e.message}`);
    }
    return `${this.table.name} was successfully deleted.`;
  }

  reverse() {
    const fieldsToFetch = "sheets(properties(sheetId),data(rowData(values(userEnteredValue,textFormatRuns,chipRuns,userEnteredFormat,effectiveValue,hyperlink,note,dataValidation))))";
    const obj = sget_(this.spreadsheetId, fieldsToFetch, [this.rangeAsA1Notation]);
    const f = obj.sheets.find(({ properties: { sheetId } }) => sheetId == this.sheetId);
    if (f) {
      this.remove();
      const rows = f.data[0].rowData;
      const requests = [{ updateCells: { rows, range: this.table.range, fields: "*" } }];
      batchUpdate_(this.spreadsheetId, requests);
    } else {
      throw new Error("Sheet not found during reverse operation.");
    }
    return `${this.table.name} was successfully reversed.`;
  }

  copyTo(a1Notation) {
    const parsed = parseA1Notation_(a1Notation);
    const sheetsData = sget_(this.spreadsheetId, "sheets(properties(sheetId,title))");
    let destSheetId, destSheetName;

    if (parsed.sheetName) {
      const s = sheetsData.sheets.find(({ properties: { title } }) => title === parsed.sheetName);
      if (s) {
        destSheetId = s.properties.sheetId;
        destSheetName = s.properties.title;
      } else {
        throw new Error(`Destination sheet "${parsed.sheetName}" not found.`);
      }
    } else {
      destSheetId = this.sheetId;
      destSheetName = this.sheetName;
    }

    const destination = convA1NotationToGridRange_(parsed.range, destSheetId);
    const requests = [{ copyPaste: { source: this.table.range, destination } }];
    batchUpdate_(this.spreadsheetId, requests);

    const tablesData = fetchAllTables_(this.spreadsheetId);
    const copiedTableObj = tablesData.tablesBySheetNames[destSheetName]?.tables?.find((t) => {
      const tr = t.getMetadata().range;
      return (tr.sheetId || 0) === (destSheetId || 0) &&
             tr.startRowIndex === destination.startRowIndex &&
             tr.startColumnIndex === destination.startColumnIndex;
    });

    if (!copiedTableObj) throw new Error("Table copied, but could not retrieve instance.");
    return copiedTableObj;
  }

  _updateTable(requests) {
    try {
      batchUpdate_(this.spreadsheetId, requests);
    } catch (e) {
      throw new Error(`Table update failed: ${e.message}`);
    }
    return this;
  }
}

/* -------------------------------------------------------------------------- */
/*                               PRIVATE HELPERS                              */
/* -------------------------------------------------------------------------- */

/**
 * Wrapper for Sheets.Spreadsheets.get
 * @private
 */
function sget_(spreadsheetId, fields = "*", ranges = []) {
  const sid = spreadsheetId || SpreadsheetApp.getActiveSpreadsheet().getId();
  const options = { fields };
  if (ranges && ranges.length > 0) {
    options.ranges = ranges;
  }
  return Sheets.Spreadsheets.get(sid, options);
}

/**
 * Wrapper for Sheets.Spreadsheets.batchUpdate
 * @private
 */
function batchUpdate_(spreadsheetId, requests) {
  const sid = spreadsheetId || SpreadsheetApp.getActiveSpreadsheet().getId();
  return Sheets.Spreadsheets.batchUpdate({ requests }, sid);
}

/**
 * Wrapper for Sheets.Spreadsheets.Values.get
 * @private
 */
function valuesGet_(spreadsheetId, range) {
  const sid = spreadsheetId || SpreadsheetApp.getActiveSpreadsheet().getId();
  const res = Sheets.Spreadsheets.Values.get(sid, range, {
    valueRenderOption: "FORMATTED_VALUE",
  });
  return res.values;
}

/**
 * Wrapper for Sheets.Spreadsheets.Values.update
 * @private
 */
function valuesUpdate_(spreadsheetId, values, range) {
  const sid = spreadsheetId || SpreadsheetApp.getActiveSpreadsheet().getId();
  const res = Sheets.Spreadsheets.Values.update({ values }, sid, range, {
    valueInputOption: "USER_ENTERED"
  });
  return res.updatedRange;
}

/**
 * Fetch and categorize all tables in the spreadsheet.
 * @private
 */
function fetchAllTables_(spreadsheetId) {
  const result = {
    tablesBySheetNames: {},
    tablesByTableNames: {},
    tablesByTableIds: {},
  };

  const sid = spreadsheetId || SpreadsheetApp.getActiveSpreadsheet().getId();

  try {
    // Omitting ranges to avoid "Entity Not Found" errors during resolution
    const response = Sheets.Spreadsheets.get(sid, {
      fields: "sheets(properties(sheetId,title),tables)"
    });

    if (!response.sheets) return result;

    response.sheets.forEach((apiSheet) => {
      const title = apiSheet.properties.title;
      const sheetId = apiSheet.properties.sheetId;
      const apiTables = apiSheet.tables || [];

      if (apiTables.length > 0) {
        const tableInstances = apiTables.map((t) => new Table({
          spreadsheetId: sid,
          sheetName: title,
          sheetId: sheetId,
          table: t,
        }));

        result.tablesBySheetNames[title] = { sheetId, tables: tableInstances };
        tableInstances.forEach((inst) => {
          const name = inst.getName();
          if (name) result.tablesByTableNames[name] = inst;
          result.tablesByTableIds[inst.getId()] = inst;
        });
      }
    });
  } catch (e) {
    Logger.log(`[TableApp] Error fetching tables structure for ID ${sid}: ${e.message}`);
    // If it fails with 404 and we used a passed ID, try one last time with Active ID
    if (e.message.indexOf("not found") !== -1 && spreadsheetId && spreadsheetId !== SpreadsheetApp.getActiveSpreadsheet().getId()) {
       return fetchAllTables_(null);
    }
  }

  return result;
}

/**
 * Parse A1Notation to separate sheet name and range.
 * @private
 */
function parseA1Notation_(a1Notation) {
  if (!a1Notation) return null;
  // Improved regex for sheet names with single quotes
  const regex = /^(?:'?(.*?)'?!|([^!]+)!)?(.+)$/;
  const match = a1Notation.match(regex);
  if (match) {
    let sheetName = (match[1] || match[2] || "").replace(/''/g, "'");
    const range = match[3];
    return { sheetName: sheetName || null, range };
  }
  return null;
}

/**
 * Converts a column letter to an index (0-based).
 * @private
 */
function columnLetterToIndex_(letter) {
  if (!letter || typeof letter !== "string") return 0;
  letter = letter.toUpperCase();
  return [...letter].reduce((c, e, i, a) =>
    (c += (e.charCodeAt(0) - 64) * Math.pow(26, a.length - i - 1)), -1);
}

/**
 * Converts a column index (0-based) to a letter.
 * @private
 */
function columnIndexToLetter_(index) {
  let a;
  return (a = Math.floor(index / 26)) >= 0
    ? columnIndexToLetter_(a - 1) + String.fromCharCode(65 + (index % 26))
    : "";
}

/**
 * Converts A1Notation to GridRange.
 * @private
 */
function convA1NotationToGridRange_(a1Notation, sheetId = 0) {
  const parsed = parseA1Notation_(a1Notation);
  const rangePart = parsed ? parsed.range : a1Notation;

  const { col, row } = rangePart.toUpperCase().split(":").reduce((o, part) => {
    const r1 = part.match(/[A-Z]+/);
    const r2 = part.match(/[0-9]+/);
    o.col.push(r1 ? columnLetterToIndex_(r1[0]) : null);
    o.row.push(r2 ? Number(r2[0]) : null);
    return o;
  }, { col: [], row: [] });

  col.sort((a, b) => a - b);
  row.sort((a, b) => a - b);

  const startCol = col[0];
  const endCol = col[1] !== undefined ? col[1] : startCol;
  const startRow = row[0];
  const endRow = row[1] !== undefined ? row[1] : startRow;

  const obj = { sheetId };
  if (startRow !== null) obj.startRowIndex = startRow - 1;
  if (endRow !== null) obj.endRowIndex = endRow;
  if (startCol !== null) obj.startColumnIndex = startCol;
  if (endCol !== null) obj.endColumnIndex = (endCol !== null) ? endCol + 1 : null;

  if (obj.endRowIndex === null) delete obj.endRowIndex;
  if (obj.endColumnIndex === null) delete obj.endColumnIndex;

  return obj;
}

/**
 * Converts GridRange to A1Notation.
 * @private
 */
function convGridRangeToA1Notation_(gridrange, sheetName = "") {
  if (!gridrange) return "";
  const startCol = gridrange.hasOwnProperty("startColumnIndex") ? columnIndexToLetter_(gridrange.startColumnIndex) : "A";
  const startRow = gridrange.hasOwnProperty("startRowIndex") ? gridrange.startRowIndex + 1 : "";
  const endCol = gridrange.hasOwnProperty("endColumnIndex") ? columnIndexToLetter_(gridrange.endColumnIndex - 1) : "";
  const endRow = gridrange.hasOwnProperty("endRowIndex") ? gridrange.endRowIndex : "";
  const startStr = `${startCol}${startRow}`;
  const endStr = `${endCol}${endRow}`;
  const rangeStr = (startStr === endStr || !endStr) ? startStr : `${startStr}:${endStr}`;
  return sheetName ? `'${sheetName.replace(/'/g, "''")}'!${rangeStr}` : rangeStr;
}
