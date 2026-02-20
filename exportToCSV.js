/**
 * Shows a dialog to select a visible sheet and exports its data to CSV
 * Integrates with Custom Tools menu (only for owner)
 */
function exportEntriesToCSV() {
  const ss = SpreadsheetApp.getActive(); //SpreadsheetApp.getActiveSpreadsheet()
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
    const tables = tableApp.getAllTables();
    tableNames = tables.map(t => t.getName());
    hasSchoolsTable = tableNames.includes('Schools');
  } catch (err) {
    Logger.log('[exportEntriesToCSV] Could not load tables: %s', err.message);
    tableNames = [];
  }

  // Build HTML for sheet dropdown
  let sheetOptionsHtml = visibleSheets.map(sheet =>
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

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      label { display: block; margin-top: 15px; font-weight: bold; }
      .small { font-size: 12px; color: #666; font-weight: normal; }
      select, button { padding: 8px; margin: 10px 0; width: 100%; font-size: 16px; }
      button { background: #4285f4; color: white; border: none; cursor: pointer; }
      button:hover { background: #3267d6; }
      #status { margin-top: 15px; color: #555; }
    </style>

    <p><strong>Auto-detects Event/Time columns</strong> (handles sheets with/without times)</p>

    <label>Select the sheet to export:</label>
    <select id="sheetId">
      <option value="">-- Select a sheet --</option>
      ${sheetOptionsHtml}
    </select>

    <label>Team Code Table <span class="small">(Team code is required for export)</span></label>
    <select id="teamCodeTable">
      ${tableOptionsHtml}
    </select>

    <button onclick="exportSheet()">Export to CSV</button>

    <div id="status"></div>

    <script>
      function exportSheet() {
        const sheetId = document.getElementById('sheetId').value;
        const teamCodeTable = document.getElementById('teamCodeTable').value;

        if (!sheetId) {
          alert('Please select a sheet.');
          return;
        }

        if (!teamCodeTable) {
          alert('No Team Code table available. Please create a table with team codes (e.g., "Schools") before exporting.');
          return;
        }

        document.getElementById('status').innerText = 'Processing...';
        document.querySelector('button').disabled = true;

        google.script.run
          .withSuccessHandler(function(fileUrl) {
            document.getElementById('status').innerHTML =
              'Success! <a href="' + fileUrl + '" target="_blank">Download CSV</a>';
            document.querySelector('button').disabled = false;
          })
          .withFailureHandler(function(err) {
            document.getElementById('status').innerText = 'Error: ' + err.message;
            document.querySelector('button').disabled = false;
          })
          .processSheetToCSV(sheetId, teamCodeTable);
      }
    </script>
  `)
  .setWidth(420)
  .setHeight(320);

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
  const ss = SpreadsheetApp.getActive(); //SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetById(sheetId);
  const spreadsheetId = ss.getId();

  if (!sheet) {
    throw new Error(`Sheet with ID ${sheetId} not found`);
  }

  const sheetName = sheet.getName();
  const data = sheet.getDataRange().getValues();

  Logger.log(`Processing sheet: ${sheetName} with ${data.length} rows`);

  if (data.length <= 1) {
    throw new Error("No data found in sheet");
  }

  // ── AUTO-DETECT Event/Time columns from HEADER ROW ───────────────────
  const headerRow = data[0];
  const eventTimePairs = detectEventTimeColumns(headerRow);

  Logger.log(`Detected ${eventTimePairs.length} event/time pairs at columns: ${eventTimePairs.map(p => `(${p.eventCol}, ${p.timeCol || 'N/A'})`).join(', ')}`);

  const convertColumnIndex = findConvertColumnIndex(headerRow);

  const headers = [
    "Team Code", "First Name", "Last Name", "Date of Birth", "Gender",
    "Event Entries", "Entry Times", "School Year"
  ];

  const outputRows = [headers];

  // --- FETCH TEAM CODE TABLE DATA ONCE (outside the loop to avoid rate limiting)
  let teamCodeData = null;
  try {
    const tableApp = TableApp.openById(spreadsheetId);
    const table = tableApp.getTableByName(teamCodeTableName);
    if (table) {
      teamCodeData = table.getValues();
      Logger.log('[exportToCSV] Successfully loaded team code data from table "%s" (%s rows)', teamCodeTableName, teamCodeData.length);
    } else {
      Logger.log('[exportToCSV] Warning: Table "%s" not found. All team codes will default to "UNS"', teamCodeTableName);
    }
  } catch (err) {
    Logger.log('[exportToCSV] Error loading team code table: %s. All team codes will default to "UNS"', err.message);
  }

  // Process each data row (skip header)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // Skip empty rows (no first name)
    const firstName = getSafeValue(row, 1);  // Zero indexed
    if (firstName === "" || firstName == null) {
      continue;
    }

    // Fixed columns (updated to match your indices)
     let teamCode = "UNS";
    const lastName   = getSafeValue(row, 2);
    const dobRaw     = getSafeValue(row, 3);
    const gender     = getSafeValue(row, 4);
    const schoolYear = getSafeValue(row, 5);
    const schoolName = getSafeValue(row, 6);

    const dobFormatted = formatDob(dobRaw);
    let convertTimes = false;

    if (convertColumnIndex !== -1) {
      convertTimes = getSafeValue(row, convertColumnIndex); // "Yes" or empty
    }


    // --- LOOKUP TEAM CODE from cached team data
    if (teamCodeData) {
      teamCode = getSchoolCode(teamCodeData, schoolName);
    } else {
      teamCode = 'UNS'; // Default to UNS if no team data available
    }


    // ── EXTRACT EVENTS & TIMES using detected columns ───────────────────
    const eventEntries = [];
    const entryTimes   = [];

    eventTimePairs.forEach(pair => {
      const event = getSafeValue(row, pair.eventCol);
      const eventStr = String(event || "").trim();

      if (eventStr !== "") {
        eventEntries.push(eventStr);

        // Time handling based on column format
        if (pair.timeCol !== null) {
          // Standard format: Event | Time pair
          const timeRaw = getSafeValue(row, pair.timeCol);
          let timeStr = (timeRaw === "" || timeRaw == null)
            ? "NT"
            : formatAsMmSsSs(timeRaw);

          // Convert times from 33.3m pool to standard 25m pool timing when requested.
          // Applies only to 25m and 50m events, assuming swims occurred in a 33.3m pool.
          //  - For 25m events: treat recorded time as 33.3m → convert using RESIZESWIM(33.3, 25, time)
          //  - For 50m events: treat recorded time as 66.6m → convert using RESIZESWIM(66.6, 50, time)
          if (timeStr !== "NT") {
            const convertFlag = String(convertTimes || "").trim().toLowerCase();
            if (convertFlag === "yes") {
              const dist = parseEventDistance(eventStr);
              try {
                if (dist === 25) {
                  timeStr = RESIZESWIM(33.3, 25, timeStr);
                } else if (dist === 50) {
                  timeStr = RESIZESWIM(66.6, 50, timeStr);
                }
              } catch (e) {
                // If anything goes wrong during conversion, keep original timeStr
                Logger.log(`Conversion skipped due to error for event "${eventStr}": ${e && e.message ? e.message : e}`);
              }
            }
          }

          entryTimes.push(timeStr);
        } else {
          // Events-only format: default "NT"
          entryTimes.push("NT");
        }
      }
    });

    const eventsStr = eventEntries.join(", ");
    const timesStr  = entryTimes.join(", ");

    outputRows.push([
      teamCode, firstName, lastName, dobFormatted, gender,
      eventsStr, timesStr, schoolYear
    ]);
  }

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

/**
 * AUTO-DETECTS Event/Time column pairs from header row
 * Only matches clear "Event X" or real event names, then checks NEXT column for Time
 * @param {array} headerRow - Header row values (e.g. data[0])
 * @return {array} Array of {eventCol: number, timeCol: number|null} pairs
 */
function detectEventTimeColumns(headerRow) {
  const pairs = [];

  // Scan from column I (index 8) onwards - earlier columns are unlikely to be events
  for (let col = 5; col < headerRow.length; col++) {
    const header = String(headerRow[col] || "").trim();

    // Skip if header is empty or obviously not an event
    if (!header) continue;

    // Strict Event detection:
    // 1. Contains "Event" (case-insensitive)
    // 2. Or starts with number + "m" (e.g. "25m Butterfly", "100m IM")
    // 3. Or matches known event patterns (very strict)
    const isLikelyEvent =
      /event/i.test(header) ||
      /^\d+m/i.test(header) ||
      /(backstroke|breaststroke|butterfly|freestyle|im|medley|relay|back|free|breast|fly)/i.test(header) &&
      !/time|m:s|\(m:s\.s\)/i.test(header);  // explicitly exclude Time headers

    if (!isLikelyEvent) continue;

    // Found a probable Event column
    const eventCol = col;
    let timeCol = null;

    // Check if NEXT column looks like a Time header
    if (col + 1 < headerRow.length) {
      const nextHeader = String(headerRow[col + 1] || "").trim().toLowerCase();
      if (
        nextHeader.includes('time') ||
        nextHeader.includes('m:s') ||
        nextHeader.includes('(m:s.s)') ||
        nextHeader.includes('m:s.s') ||
        /^\d{1,2}:\d{2}\.\d{2}$/.test(nextHeader)  // rare: time already in header
      ) {
        timeCol = col + 1;
        // Skip the next column in the loop (we've consumed it as Time)
        col++;
      }
    }

    pairs.push({ eventCol, timeCol });
  }

  // Fallback: if nothing detected, assume standard J:AA (9 pairs)
  if (pairs.length === 0) {
    Logger.log("No Event columns detected → using fallback J:AA (indices 9-26)");
    for (let col = 9; col <= 26; col += 2) {
      pairs.push({ eventCol: col, timeCol: col + 1 });
    }
  }

  return pairs;
}

/**
 * Helper: Detects if header looks like an actual event name
 * e.g. "25m Backstroke", "50m Free", "100 IM", etc.
 */
function isEventName(header) {
  const eventPatterns = [
    /\d+m\s*(back|breast|free|fly|im|medley)/i,
    /\d+m\s*(freestyle|backstroke|breaststroke|butterfly)/i,
    /(freestyle|backstroke|breaststroke|butterfly|im|medley)/i,
    /kick|pull|choice|relay/i
  ];
  return eventPatterns.some(pattern => pattern.test(header));
}

/**
 * Helper: Detects if header is a convert utility column
 * Matches labels like: "Convert times from 33m pool", "Convert 33m to 25m", etc.
 * Loosely requires the presence of words "convert" and "33m" (or variants) in any order.
 * @param {string} header
 * @return {boolean}
 */
function isConvertColumn(header) {
  const h = String(header || "").trim().toLowerCase();
  if (!h) return false;

  const hasConvertWord = /(convert|conversion|converted)/.test(h);
  const has33mWord = /(\b33\s*m\b|\b33m\b|33\s*met(er|re)s?)/.test(h);

  // Optional helpers
  // const mentionsPool = /pool/.test(h);

  return hasConvertWord && has33mWord;
}

/**
 * Finds the index (zero-based) of the column with a header like
 * "Convert times from 33m pool" or similar. Returns -1 if not found.
 * Relies on {@link isConvertColumn} for matching variations.
 * @param {Array<string>} headerRow
 * @return {number}
 */
function findConvertColumnIndex(headerRow) {
  if (!Array.isArray(headerRow)) return -1;
  for (let i = 0; i < headerRow.length; i++) {
    if (isConvertColumn(headerRow[i])) return i;
  }
  return -1;
}

/**
 * Safe way to get cell value (handles undefined/null)
 */
function getSafeValue(row, index) {
  return (index < row.length && row[index] != null) ? row[index] : "";
}


/**
 * Parses the event distance from an event string.
 * Supports formats like "25m Freestyle", "50 m Backstroke" (case-insensitive).
 * Only 25m and 50m are recognized for conversion; others return null.
 * @param {string} eventStr
 * @return {number|null} 25, 50, or null if not recognized
 */
function parseEventDistance(eventStr) {
  const s = String(eventStr || "").trim();
  // Match distance at the start: 25m or 50m (allow space before 'm')
  const m = /^\s*(25|50)\s*m\b/i.exec(s);
  if (!m) return null;
  const dist = parseInt(m[1], 10);
  return (dist === 25 || dist === 50) ? dist : null;
}
