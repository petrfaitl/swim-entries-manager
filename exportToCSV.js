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

  // Build HTML for dropdown
  let optionsHtml = visibleSheets.map(sheet => 
    `<option value="${sheet.getSheetId()}">${sheet.getName()}</option>`
  ).join('');

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      select, button { padding: 8px; margin: 10px 0; width: 100%; font-size: 16px; }
      button { background: #4285f4; color: white; border: none; cursor: pointer; }
      button:hover { background: #3267d6; }
      #status { margin-top: 15px; color: #555; }
    </style>

    <p><strong>Auto-detects Event/Time columns</strong> (handles sheets with/without times)</p>
    <p>Select the sheet to export:</p>
    
    <select id="sheetId">
      <option value="">-- Select a sheet --</option>
      ${optionsHtml}
    </select>
    
    <button onclick="exportSheet()">Export to CSV</button>
    
    <div id="status"></div>

    <script>
      function exportSheet() {
        const sheetId = document.getElementById('sheetId').value;
        if (!sheetId) {
          alert('Please select a sheet.');
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
          .processSheetToCSV(sheetId);
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
 * @return {string} URL to the created CSV file
 */
function processSheetToCSV(sheetId) {
  const ss = SpreadsheetApp.getActive(); //SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetById(sheetId);
  
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

  const headers = [
    "Team Code", "First Name", "Last Name", "Date of Birth", "Gender",
    "Event Entries", "Entry Times", "School Year"
  ];

  const outputRows = [headers];

  // Process each data row (skip header)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // Skip empty rows (no first name)
    const firstName = getSafeValue(row, 3);  // Column D (index 3)
    if (firstName === "" || firstName == null) {
      continue;
    }

    // Fixed columns (updated to match your indices)
    const teamCode   = getSafeValue(row, 1);   // B (1)
    const lastName   = getSafeValue(row, 4);   // E (4)
    const dobRaw     = getSafeValue(row, 5);   // F (5)
    const gender     = getSafeValue(row, 6);   // G (6)
    const schoolYear = getSafeValue(row, 7);   // H (7)

    const dobFormatted = formatDob(dobRaw);

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
          const timeStr = (timeRaw === "" || timeRaw == null) 
            ? "NT" 
            : formatAsMmSsSs(timeRaw);
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
  for (let col = 8; col < headerRow.length; col++) {
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
 * Safe way to get cell value (handles undefined/null)
 */
function getSafeValue(row, index) {
  return (index < row.length && row[index] != null) ? row[index] : "";
}

