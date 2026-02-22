/**
 * EntryDataProcessor — Shared data processing logic for swim entries
 *
 * Provides common functions for:
 * - Name capitalization
 * - School year normalization
 * - Team code lookup
 * - Event/Time column detection
 * - Row processing from raw sheet data to clean structured data
 *
 * Used by both exportToCSV.js and SDIFCreator.js
 */

/**
 * Capitalizes names properly, handling all-caps and all-lowercase cases.
 * Preserves existing mixed-case names (like McDonald, MacLeod, O'Brien, etc.)
 * Only "fixes" names that are entirely uppercase or entirely lowercase.
 *
 * @param {string} name - The name to format
 * @return {string} Properly capitalized name
 */
function capitalizeNameSafely(name) {
  if (!name || typeof name !== 'string') return name;

  const trimmed = name.trim();
  if (!trimmed) return trimmed;

  // Check if name is all uppercase or all lowercase
  const isAllUpper = trimmed === trimmed.toUpperCase() && trimmed !== trimmed.toLowerCase();
  const isAllLower = trimmed === trimmed.toLowerCase() && trimmed !== trimmed.toUpperCase();

  // Only fix if entirely one case - preserve mixed case names
  if (!isAllUpper && !isAllLower) {
    return name; // Already has mixed case, preserve it (handles Mac/Mc, O', etc.)
  }

  // Convert to title case: capitalize first letter of each word
  return trimmed.split(/(\s|-|')/).map(function(part) {
    if (!part || part === ' ' || part === '-' || part === "'") {
      return part; // Preserve separators
    }
    return part.charAt(0).toUpperCase() + part.slice(1).toLowerCase();
  }).join('');
}

/**
 * Normalizes school year to SDIF-compatible 2-character format.
 * SDIF school year field is only 2 characters long.
 *
 * Rules:
 * - Years 1-9: "Y1" to "Y9" (or "1" to "9")
 * - Years 10-13: "10" to "13"
 * - Junior/Jn: "Jn"
 * - Intermediate/In: "In"
 * - Senior/Sr: "Sr"
 *
 * @param {string|number} year - School year value (e.g., "Year 5", "Y5", "5", "Junior", "Jn", "Senior")
 * @return {string} Normalized 2-character school year
 */
function normalizeSchoolYear(year) {
  if (!year) return '';

  const yearStr = String(year).trim().toUpperCase();
  if (!yearStr) return '';

  // Check for Junior/Intermediate/Senior
  if (yearStr.includes('JUNIOR') || yearStr === 'JN') return 'Jn';
  if (yearStr.includes('INTERMEDIATE') || yearStr === 'IN') return 'In';
  if (yearStr.includes('SENIOR') || yearStr === 'SR') return 'Sr';

  // Extract numeric part (handles "Year 5", "Y5", "5", etc.)
  const numMatch = yearStr.match(/(\d+)/);
  if (numMatch) {
    const num = parseInt(numMatch[1], 10);
    if (num >= 1 && num <= 9) {
      return 'Y' + num;
    } else if (num >= 10 && num <= 13) {
      if (num === 10){
        return 'Jn';
      }else if (num === 11){
        return 'In';
      }
      return 'Sr';
    }
  }

  // If we can't parse it, return first 2 characters (fallback)
  return yearStr.substring(0, 2);
}

/**
 * AUTO-DETECTS Event/Time column pairs from header row
 * Only matches clear "Event X" or real event names, then checks NEXT column for Time
 * @param {array} headerRow - Header row values (e.g. data[0])
 * @return {array} Array of {eventCol: number, timeCol: number|null} pairs
 */
function detectEventTimeColumns(headerRow) {
  const pairs = [];

  // Scan from column F (index 5) onwards - earlier columns are unlikely to be events
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
    Logger.log("[EntryDataProcessor] No Event columns detected → using fallback J:AA (indices 9-26)");
    for (let col = 9; col <= 26; col += 2) {
      pairs.push({ eventCol: col, timeCol: col + 1 });
    }
  }

  return pairs;
}

/**
 * Finds the index (zero-based) of the column with a header like
 * "Convert times from 33m pool" or similar. Returns -1 if not found.
 * @param {Array<string>} headerRow
 * @return {number}
 */
function findConvertColumnIndex(headerRow) {
  if (!Array.isArray(headerRow)) return -1;
  for (let i = 0; i < headerRow.length; i++) {
    const h = String(headerRow[i] || "").trim().toLowerCase();
    if (!h) continue;

    const hasConvertWord = /(convert|conversion|converted)/.test(h);
    const has33mWord = /(\b33\s*m\b|\b33m\b|33\s*met(er|re)s?)/.test(h);

    if (hasConvertWord && has33mWord) return i;
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
 * Formats date of birth to DD/MM/YYYY format
 * @param {Date|number|string} dobRaw
 * @return {string} Formatted date as DD/MM/YYYY or empty string
 */
function formatDob(dobRaw) {
  if (!dobRaw) return "";

  let date;
  if (dobRaw instanceof Date) {
    date = dobRaw;
  } else if (typeof dobRaw === 'number') {
    // Excel/Sheets serial date
    date = new Date((dobRaw - 25569) * 86400000);
  } else if (typeof dobRaw === 'string') {
    const trimmed = dobRaw.trim();
    if (!trimmed) return "";

    // Try DD/MM/YYYY format first
    const parts = trimmed.split('/');
    if (parts.length === 3) {
      const day = parseInt(parts[0], 10);
      const month = parseInt(parts[1], 10);
      const year = parseInt(parts[2], 10);
      date = new Date(year, month - 1, day);
    } else {
      date = new Date(trimmed);
    }
  }

  if (!date || isNaN(date.getTime())) {
    return "";
  }

  const dd = String(date.getDate()).padStart(2, '0');
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const yyyy = String(date.getFullYear());

  return dd + '/' + mm + '/' + yyyy;
}

/**
 * Formats time as MM:SS.SS
 * @param {Date|number|string} timeRaw
 * @return {string}
 */
function formatAsMmSsSs(timeRaw) {
  if (!timeRaw && timeRaw !== 0) return "NT";

  const s = String(timeRaw).trim();
  if (s === "" || s.toUpperCase() === "NT") return "NT";

  // If it's already in MM:SS.SS format, return as is
  if (/^\d{1,2}:\d{2}\.\d{2}$/.test(s)) return s;

  // Handle serial time (decimal < 1)
  if (typeof timeRaw === 'number' && timeRaw < 1) {
    const totalSeconds = timeRaw * 86400;
    const minutes = Math.floor(totalSeconds / 60);
    const seconds = totalSeconds % 60;
    const mm = String(minutes).padStart(2, '0');
    const ss = seconds.toFixed(2).padStart(5, '0');
    return mm + ':' + ss;
  }

  // Parse string formats like "1:23.45" or "83.45"
  if (s.includes(':')) {
    const parts = s.split(':');
    if (parts.length === 2) {
      const minutes = parseInt(parts[0], 10);
      const secondsPart = parseFloat(parts[1]);
      const mm = String(minutes).padStart(2, '0');
      const ss = secondsPart.toFixed(2).padStart(5, '0');
      return mm + ':' + ss;
    }
  } else {
    // Assume it's seconds only
    const totalSec = parseFloat(s);
    if (!isNaN(totalSec)) {
      const minutes = Math.floor(totalSec / 60);
      const seconds = totalSec % 60;
      const mm = String(minutes).padStart(2, '0');
      const ss = seconds.toFixed(2).padStart(5, '0');
      return mm + ':' + ss;
    }
  }

  return "NT";
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

/**
 * Looks up school code from team code data
 * @param {Array} teamCodeData - Array of rows from Schools table
 * @param {string} schoolName - School name to lookup
 * @return {string} Team code or "UNS" if not found
 */
function getSchoolCode(teamCodeData, schoolName) {
  if (!teamCodeData || !schoolName) return "UNS";

  const normalizedSchoolName = String(schoolName).trim().toLowerCase();
  if (!normalizedSchoolName) return "UNS";

  // Find matching row (case-insensitive)
  for (let i = 1; i < teamCodeData.length; i++) {
    const row = teamCodeData[i];
    if (!row || row.length < 2) continue;

    const tableName = String(row[1] || "").trim().toLowerCase();
    if (tableName === normalizedSchoolName) {
      const code = String(row[0] || "").trim();
      return code || "UNS";
    }
  }

  return "UNS";
}

/**
 * Process a single row from the entries sheet into structured data
 *
 * @param {Array} row - Raw row data from sheet
 * @param {Array} eventTimePairs - Detected event/time column pairs
 * @param {number} convertColumnIndex - Index of convert column, or -1
 * @param {Array} teamCodeData - Team code lookup data
 * @return {Object} Processed swimmer data with: teamCode, firstName, lastName, dobFormatted, gender, eventsStr, timesStr, schoolYear
 */
function processEntryRow(row, eventTimePairs, convertColumnIndex, teamCodeData) {
  // Skip empty rows (no first name)
  const firstNameRaw = getSafeValue(row, 1);  // Zero indexed
  if (firstNameRaw === "" || firstNameRaw == null) {
    return null;
  }

  // Fixed columns
  let teamCode = "UNS";
  const lastNameRaw = getSafeValue(row, 2);

  // Apply safe capitalization to names (only fixes all-caps or all-lowercase)
  const firstName = capitalizeNameSafely(firstNameRaw);
  const lastName = capitalizeNameSafely(lastNameRaw);
  const dobRaw = getSafeValue(row, 3);
  const gender = getSafeValue(row, 4);
  const schoolYearRaw = getSafeValue(row, 5);
  const schoolName = getSafeValue(row, 6);

  // Normalize school year to SDIF 2-character format
  const schoolYear = normalizeSchoolYear(schoolYearRaw);
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
  const entryTimes = [];

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
              Logger.log(`[EntryDataProcessor] Conversion skipped due to error for event "${eventStr}": ${e && e.message ? e.message : e}`);
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
  const timesStr = entryTimes.join(", ");

  return {
    teamCode: teamCode,
    firstName: firstName,
    lastName: lastName,
    dobFormatted: dobFormatted,
    gender: gender,
    eventsStr: eventsStr,
    timesStr: timesStr,
    schoolYear: schoolYear
  };
}

/**
 * Process all rows from an entries sheet
 *
 * @param {Sheet} sheet - Google Sheets sheet object
 * @param {string} teamCodeTableName - Name of the table for team code lookup
 * @param {string} spreadsheetId - Spreadsheet ID
 * @return {Array} Array of processed swimmer objects
 */
function processEntriesSheet(sheet, teamCodeTableName, spreadsheetId) {
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    throw new Error("[EntryDataProcessor] No data found in sheet");
  }

  // ── AUTO-DETECT Event/Time columns from HEADER ROW ───────────────────
  const headerRow = data[0];
  const eventTimePairs = detectEventTimeColumns(headerRow);

  Logger.log(`[EntryDataProcessor] Detected ${eventTimePairs.length} event/time pairs at columns: ${eventTimePairs.map(p => `(${p.eventCol}, ${p.timeCol || 'N/A'})`).join(', ')}`);

  const convertColumnIndex = findConvertColumnIndex(headerRow);

  // --- FETCH TEAM CODE TABLE DATA ONCE (outside the loop to avoid rate limiting)
  let teamCodeData = null;
  try {
    const tableApp = TableApp.openById(spreadsheetId);
    const table = tableApp.getTableByName(teamCodeTableName);
    if (table) {
      teamCodeData = table.getValues();
      Logger.log('[EntryDataProcessor] Successfully loaded team code data from table "%s" (%s rows)', teamCodeTableName, teamCodeData.length);
    } else {
      Logger.log('[EntryDataProcessor] Warning: Table "%s" not found. All team codes will default to "UNS"', teamCodeTableName);
    }
  } catch (err) {
    Logger.log('[EntryDataProcessor] Error loading team code table: %s. All team codes will default to "UNS"', err.message);
  }

  // Process each data row (skip header)
  const processedRows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const processed = processEntryRow(row, eventTimePairs, convertColumnIndex, teamCodeData);
    if (processed) {
      processedRows.push(processed);
    }
  }

  return processedRows;
}
