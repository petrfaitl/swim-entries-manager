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

  let trimmed = name.trim();
  if (!trimmed) return trimmed;

  // Replace problematic characters with proper equivalents
  // Back-tick (`) -> apostrophe (')
  // Smart quotes (",") -> straight quotes (")
  // Various apostrophe-like characters -> standard apostrophe
  trimmed = trimmed
    .replace(/`/g, "'")           // Back-tick to apostrophe
    .replace(/[\u2018\u2019]/g, "'")  // Smart single quotes to apostrophe
    .replace(/[\u201C\u201D]/g, '"')  // Smart double quotes to straight quotes
    .replace(/[\u02BC\u2032]/g, "'"); // Other apostrophe-like chars

  // Check if name is all uppercase or all lowercase
  const isAllUpper = trimmed === trimmed.toUpperCase() && trimmed !== trimmed.toLowerCase();
  const isAllLower = trimmed === trimmed.toLowerCase() && trimmed !== trimmed.toUpperCase();

  // Only fix if entirely one case - preserve mixed case names
  if (!isAllUpper && !isAllLower) {
    return trimmed; // Already has mixed case, preserve it (handles Mac/Mc, O', etc.)
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
 * - Years 10-13: "Jr", "In" or "Sr"
 * - Junior/Jn: "Jr"
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
  if (yearStr.includes('JUNIOR') || yearStr === 'JR') return 'Jr';
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
        return 'Jr';
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
 * Detects the column mapping based on header row
 * Handles tables with or without the '#' numbering column
 *
 * @param {Array} headerRow - Header row from sheet
 * @return {Object} Column mapping with indices for: firstName, lastName, dob, gender, schoolYear, school, hasNumberColumn, eventStartCol
 */
function detectColumnMapping(headerRow) {
  const mapping = {
    firstName: -1,
    lastName: -1,
    dob: -1,
    gender: -1,
    schoolYear: -1,
    school: -1,
    hasNumberColumn: false,
    eventStartCol: -1
  };

  // Check if first column is a number/hash column
  const firstHeader = String(headerRow[0] || "").trim().toLowerCase();
  mapping.hasNumberColumn = firstHeader === '#' || firstHeader === 'no' || firstHeader === 'no.' ||
                            firstHeader === 'number' || firstHeader === '' && String(headerRow[1] || "").trim().toLowerCase().includes('first');

  // Set base offset
  const offset = mapping.hasNumberColumn ? 1 : 0;

  // Detect standard columns by pattern matching
  for (let col = 0; col < Math.min(headerRow.length, 15); col++) {
    const header = String(headerRow[col] || "").trim().toLowerCase();

    if (!header) continue;

    // First Name detection
    if ((header.includes('first') && header.includes('name')) || header === 'first name' || header === 'firstname') {
      mapping.firstName = col;
    }
    // Last Name detection
    else if ((header.includes('last') && header.includes('name')) || header === 'last name' || header === 'lastname' || header === 'surname') {
      mapping.lastName = col;
    }
    // Date of Birth detection
    else if (header.includes('date') && header.includes('birth') || header === 'dob' || header === 'birth date' || header === 'birthdate') {
      mapping.dob = col;
    }
    // Gender detection
    else if (header === 'gender' || header === 'sex' || header === 'm/f') {
      mapping.gender = col;
    }
    // School Year detection
    else if ((header.includes('school') && header.includes('year')) ||
             (header.includes('year') && header.includes('level')) ||
             header === 'year' || header === 'year level' || header === 'grade') {
      mapping.schoolYear = col;
    }
    // School detection
    else if (header === 'school' || header === 'team' || header === 'club' || header === 'house') {
      mapping.school = col;
    }
  }

  // Fallback to default positions if not found
  if (mapping.firstName === -1) mapping.firstName = 0 + offset;
  if (mapping.lastName === -1) mapping.lastName = 1 + offset;
  if (mapping.dob === -1) mapping.dob = 2 + offset;
  if (mapping.gender === -1) mapping.gender = 3 + offset;
  if (mapping.schoolYear === -1) mapping.schoolYear = 4 + offset;
  if (mapping.school === -1) mapping.school = 5 + offset;

  // Event columns typically start after school column
  mapping.eventStartCol = Math.max(mapping.school + 1, 6 + offset);

  Logger.log(`[EntryDataProcessor] Column mapping detected: hasNumberColumn=${mapping.hasNumberColumn}, firstName=${mapping.firstName}, lastName=${mapping.lastName}, dob=${mapping.dob}, gender=${mapping.gender}, schoolYear=${mapping.schoolYear}, school=${mapping.school}, eventStartCol=${mapping.eventStartCol}`);

  return mapping;
}

/**
 * Determines if an event name is valid (real event format)
 * Valid: "25m Freestyle", "100m Backstroke", "50m IM"
 * Invalid: "Event 1", "Event", generic text
 * @param {string} header - Header text to check
 * @return {boolean}
 */
function isRealEventName(header) {
  const h = String(header || "").trim();
  if (!h) return false;

  // Must start with distance (e.g., "25m", "100m")
  const hasDistance = /^\d+m/i.test(h);

  // Must contain stroke/discipline name
  const hasStroke = /(backstroke|breaststroke|butterfly|freestyle|im|medley|relay|back|free|breast|fly|kick)/i.test(h);

  return hasDistance && hasStroke;
}

/**
 * COLUMN FORMAT DETECTION
 * Detects which format the sheet uses:
 * - METHOD_1: "Event 1" | "Time 1" pairs (values in cells)
 * - METHOD_2: "100m Freestyle" | "100m Freestyle Best Time" (Yes/NT in cells)
 *
 * @param {array} headerRow - Header row values
 * @return {object} {method: 'METHOD_1'|'METHOD_2', confidence: number}
 */
function detectColumnFormat(headerRow) {
  let method1Score = 0;
  let method2Score = 0;

  // Scan headers starting from column F (index 5)
  for (let col = 5; col < headerRow.length; col++) {
    const header = String(headerRow[col] || "").trim();
    if (!header) continue;

    // METHOD_1 indicators: "Event X" pattern
    if (/event\s*\d+/i.test(header)) {
      method1Score += 10;
    }

    // METHOD_1 indicators: "Time X" pattern
    if (/time\s*\d+/i.test(header)) {
      method1Score += 5;
    }

    // METHOD_2 indicators: Real event names in headers
    if (isRealEventName(header)) {
      method2Score += 10;
    }

    // METHOD_2 indicators: "Best Time" pattern following event name
    if (/(best\s*time|entry\s*time|pb\s*time)/i.test(header)) {
      method2Score += 5;

      // Check if previous column was a real event name
      if (col > 0 && isRealEventName(headerRow[col - 1])) {
        method2Score += 10; // Strong indicator
      }
    }
  }

  Logger.log(`[EntryDataProcessor] Format detection scores: METHOD_1=${method1Score}, METHOD_2=${method2Score}`);

  // Determine which method based on scores
  if (method1Score > method2Score) {
    return { method: 'METHOD_1', confidence: method1Score };
  } else if (method2Score > method1Score) {
    return { method: 'METHOD_2', confidence: method2Score };
  } else {
    // Default to METHOD_1 if scores are equal
    Logger.log('[EntryDataProcessor] Warning: Unable to determine format confidently, defaulting to METHOD_1');
    return { method: 'METHOD_1', confidence: 0 };
  }
}

/**
 * AUTO-DETECTS Event/Time column pairs from header row (METHOD 1)
 * Only matches clear "Event X" or real event names, then checks NEXT column for Time
 * @param {array} headerRow - Header row values (e.g. data[0])
 * @return {array} Array of {eventCol: number, timeCol: number|null, method: 'METHOD_1'} pairs
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

    pairs.push({ eventCol, timeCol, method: 'METHOD_1' });
  }

  // Fallback: if nothing detected, assume standard J:AA (9 pairs)
  if (pairs.length === 0) {
    Logger.log("[EntryDataProcessor] No Event columns detected → using fallback J:AA (indices 9-26)");
    for (let col = 9; col <= 26; col += 2) {
      pairs.push({ eventCol: col, timeCol: col + 1, method: 'METHOD_1' });
    }
  }

  return pairs;
}

/**
 * AUTO-DETECTS Event/BestTime column pairs from header row (METHOD 2)
 * Looks for real event names (e.g., "100m Freestyle") followed by optional "Best Time" column
 * @param {array} headerRow - Header row values
 * @return {array} Array of {eventCol: number, timeCol: number|null, eventName: string, method: 'METHOD_2'} pairs
 */
function detectEventColumnsWithBestTime(headerRow) {
  const pairs = [];

  // Scan from column F (index 5) onwards
  for (let col = 5; col < headerRow.length; col++) {
    const header = String(headerRow[col] || "").trim();

    // Skip empty headers
    if (!header) continue;

    // Check if this is a real event name
    if (isRealEventName(header)) {
      const eventCol = col;
      const eventName = header;
      let timeCol = null;

      // Check if NEXT column is "Best Time" or similar
      if (col + 1 < headerRow.length) {
        const nextHeader = String(headerRow[col + 1] || "").trim();

        // Look for patterns like "Best Time", "Entry Time", or event name + "Best Time"
        const isBestTimeColumn =
          /(best\s*time|entry\s*time)/i.test(nextHeader) ||
          (nextHeader.toLowerCase().includes(header.toLowerCase()) && /time/i.test(nextHeader));

        if (isBestTimeColumn) {
          timeCol = col + 1;
          col++; // Skip the time column in next iteration
        }
      }

      pairs.push({
        eventCol,
        timeCol,
        eventName,
        method: 'METHOD_2'
      });
    }
  }

  Logger.log(`[EntryDataProcessor] METHOD_2: Detected ${pairs.length} event columns with real event names`);

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


    if (hasConvertWord) return i;
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
 * Reformats time strings into MM:SS.SS format
 * Does NOT recalculate or convert - only parses and reformats existing time values
 *
 * Accepts formats:
 * - MM:SS.SS (e.g. "01:30.95") - minutes:seconds.decimals
 * - SS.SS (e.g. "45.67") - seconds only, no minutes
 * - MM:SS:SS (e.g. "01:30:95") - last : treated as .
 * - MM:SS.S (e.g. "01:05.6") - one decimal, pads to two
 * - SS:DD (e.g. "34:56") - seconds:decimals (interpreted as 34.56 seconds, not 34 minutes)
 * - MM:SS (e.g. "1:23") - minutes:seconds (small first value = minutes)
 * - SS.S (e.g. "45.6") - seconds only, one decimal
 *
 * Disambiguation logic for single colon (X:Y):
 * - If Y has 2 digits AND X >= 10 AND X < 60: treat as SS:DD (seconds.decimals)
 *   Example: "34:56" → 00:34.56 (34.56 seconds)
 * - Otherwise: treat as MM:SS (minutes:seconds)
 *   Example: "1:23" → 01:23.00 (1 minute 23 seconds)
 * - If colon AND dot present: always MM:SS.DD format
 *   Example: "1:23.45" → 01:23.45 (1 minute 23.45 seconds)
 *
 * @param {Date|number|string} timeRaw - Raw time value from cell
 * @return {string} Formatted time as MM:SS.SS or "NT"
 */
function formatAsMmSsSs(timeRaw) {
  if (!timeRaw && timeRaw !== 0) return "NT";

  const s = String(timeRaw).trim();
  if (s === "" || s.toUpperCase() === "NT" || s.toUpperCase() === "DNS" || s.toUpperCase() === "DNF") return "NT";

  // Check for free text (contains letters) - return NT
  if (/[a-zA-Z]/.test(s)) return "NT";

  // Handle serial time from Google Sheets (decimal < 1 represents fraction of a day)
  // Special case: "decimal seconds" format where e.g. 0.17 means 17 seconds
  // This can be passed as a number OR a string like "0.17"
  const numericVal = Number(timeRaw);
  if (!isNaN(numericVal) && numericVal > 0 && numericVal < 1) {
    let totalSeconds;

    if (numericVal >= 0.01) {
      // Rule: values like 0.17 treated as 17 seconds (as per user request)
      totalSeconds = numericVal * 100;
    } else {
      // Standard serial time (fraction of a day)
      totalSeconds = numericVal * 86400;
    }

    const minutes = Math.floor(totalSeconds / 60);
    const seconds = totalSeconds % 60;

    // Validate ranges
    if (minutes >= 60) return "NT";
    if (seconds >= 60) return "NT";

    const mm = String(minutes).padStart(2, '0');
    const secondsInt = Math.floor(seconds);
    const secondsDec = Math.round((seconds - secondsInt) * 100);
    const ss = String(secondsInt).padStart(2, '0');
    const dec = String(secondsDec).padStart(2, '0');

    return mm + ':' + ss + '.' + dec;
  }

  // ─── STRING PARSING ────────────────────────────────────────────────

  let workingStr = s;

  // Step 1: Determine format and handle special cases
  const colonCount = (s.match(/:/g) || []).length;
  const dotCount = (s.match(/\./g) || []).length;

  if (colonCount === 2) {
    // Three parts separated by colons: MM:SS:DD format
    // Last colon should be treated as decimal point
    const lastColonIndex = s.lastIndexOf(':');
    workingStr = s.substring(0, lastColonIndex) + '.' + s.substring(lastColonIndex + 1);
  } else if (dotCount === 2) {
    // Three parts separated by dots: MM.SS.DD format
    // First dot should be treated as a colon
    const firstDotIndex = s.indexOf('.');
    workingStr = s.substring(0, firstDotIndex) + ':' + s.substring(firstDotIndex + 1);
  } else if (colonCount === 1 && dotCount === 0) {
    // Single colon, no decimal: Could be MM:SS or SS:DD
    // Disambiguate based on the value after the colon
    const parts = s.split(':');
    const firstPart = parseInt(parts[0], 10);
    const secondPart = parseInt(parts[1], 10);

    // If second part has exactly 2 digits AND first part < 60, treat as SS:DD (seconds with decimals)
    // Pattern: 34:56 means 34.56 seconds (not 34 minutes 56 seconds)
    // Pattern: 1:23 means 1 minute 23 seconds (first part is small, typical for minutes)
    if (parts[1].length === 2 && firstPart >= 10 && firstPart < 60 && secondPart < 100) {
      // This looks like SS:DD format (e.g., "34:56" = 34.56 seconds)
      workingStr = firstPart + '.' + parts[1];
    } else {
      // This is MM:SS format (e.g., "1:23" = 1 minute 23 seconds)
      // Keep as is - will be processed below
      workingStr = s;
    }
  } else if (colonCount === 1 && dotCount === 1) {
    // Has both colon and dot: This is definitely MM:SS.DD format
    // Keep as is
    workingStr = s;
  }

  // Step 2: Split into components
  let minutes = 0;
  let secondsStr = workingStr;

  if (workingStr.includes(':')) {
    const parts = workingStr.split(':');
    if (parts.length === 2) {
      minutes = parseInt(parts[0], 10);
      secondsStr = parts[1];
    } else {
      return "NT"; // Invalid format
    }
  }

  // Step 3: Parse seconds (integer and decimal parts)
  let secondsInt;
  let decimalPart = "00";

  if (secondsStr.includes('.')) {
    const secParts = secondsStr.split('.');
    secondsInt = parseInt(secParts[0], 10);

    // Handle decimal padding: .6 becomes .60 (not .06)
    // Single digit represents TENTHS, must be RIGHT-padded
    if (secParts[1]) {
      const dec = secParts[1];
      if (dec.length === 1) {
        decimalPart = dec + '0'; // Right-pad: .6 -> .60
      } else if (dec.length === 2) {
        decimalPart = dec;
      } else {
        // Truncate to 2 digits if longer
        decimalPart = dec.substring(0, 2);
      }
    }
  } else {
    // No decimal point
    secondsInt = parseInt(secondsStr, 10);
  }

  // Step 4: Validate
  if (isNaN(minutes) || isNaN(secondsInt)) return "NT";
  if (minutes < 0 || minutes >= 60) return "NT";
  if (secondsInt < 0 || secondsInt >= 60) return "NT";

  // Step 5: Format output
  const mm = String(minutes).padStart(2, '0');
  const ss = String(secondsInt).padStart(2, '0');

  return mm + ':' + ss + '.' + decimalPart;
}

/**
 * Parses the event distance from an event string.
 * Supports formats like "25m Freestyle", "50 m Backstroke" (case-insensitive).
 * Recognizes: 25m, 50m, 100m, 200m, 400m
 * @param {string} eventStr
 * @return {number|null} Distance in meters, or null if not recognized
 */
function parseEventDistance(eventStr) {
  const s = String(eventStr || "").trim();
  // Match distance at the start: 25m, 50m, 100m, 200m, 400m (allow space before 'm')
  const m = /^\s*(25|50|100|200|400)\s*m\b/i.exec(s);
  if (!m) return null;
  const dist = parseInt(m[1], 10);
  // Validate it's one of the supported distances
  return [25, 50, 100, 200, 400].includes(dist) ? dist : null;
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
 * Handles both METHOD_1 (Event/Time pairs) and METHOD_2 (Yes/NT indicators)
 *
 * @param {Array} row - Raw row data from sheet
 * @param {Array} eventTimePairs - Detected event/time column pairs (includes method info)
 * @param {number} convertColumnIndex - Index of convert column, or -1
 * @param {Array} teamCodeData - Team code lookup data
 * @param {Object} columnMapping - Column indices mapping (from detectColumnMapping)
 * @return {Object} Processed swimmer data with: teamCode, firstName, lastName, dobFormatted, gender, eventsStr, timesStr, schoolYear
 */
function processEntryRow(row, eventTimePairs, convertColumnIndex, teamCodeData, columnMapping) {
  // Skip empty rows (no first name) - use dynamic column index
  const firstNameRaw = getSafeValue(row, columnMapping.firstName);
  if (firstNameRaw === "" || firstNameRaw == null) {
    return null;
  }

  // Use dynamic column indices from mapping
  let teamCode = "UNS";
  const lastNameRaw = getSafeValue(row, columnMapping.lastName);

  // Apply safe capitalization to names (only fixes all-caps or all-lowercase)
  const firstName = capitalizeNameSafely(firstNameRaw);
  const lastName = capitalizeNameSafely(lastNameRaw);
  const dobRaw = getSafeValue(row, columnMapping.dob);
  const gender = getSafeValue(row, columnMapping.gender);
  const schoolYearRaw = getSafeValue(row, columnMapping.schoolYear);
  const schoolName = getSafeValue(row, columnMapping.school);

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
    if (pair.method === 'METHOD_2') {
      // METHOD 2: Event name in header, "Yes" or time in event column
      const cellValue = getSafeValue(row, pair.eventCol);
      const cellStr = String(cellValue || "").trim().toUpperCase();

      // Check if swimmer is entered in this event
      const isEntered = cellStr === "YES" || cellStr === "Y" || cellStr !== "" && cellStr !== "NO" && cellStr !== "N";

      if (isEntered) {
        // Use the event name from the header
        eventEntries.push(pair.eventName);

        // Check if there's a time in the Best Time column
        if (pair.timeCol !== null) {
          const timeRaw = getSafeValue(row, pair.timeCol);
          const timeStr = String(timeRaw || "").trim().toUpperCase();

          // If the cell value in event column looks like a time, use that instead
          let finalTime;
          if (/^\d{1,2}:\d{2}/.test(cellStr) || /^\d+\.\d+$/.test(cellStr)) {
            // Event column contains the time
            finalTime = formatAsMmSsSs(cellValue);
          } else if (timeStr === "" || timeStr === "NT" || timeStr === "NO TIME") {
            finalTime = "NT";
          } else {
            finalTime = formatAsMmSsSs(timeRaw);
          }

          // Apply pool conversion if needed
          if (finalTime !== "NT") {
            const convertFlag = String(convertTimes || "").trim();
            if (convertFlag !== "" && convertFlag.toLowerCase() !== "no") {
              const dist = parseEventDistance(pair.eventName);
              if (dist) {
                try {
                  const strokeType = getStrokeType(pair.eventName);
                  const sourceDistance = getSourceDistance(convertFlag, dist, strokeType);
                  if (sourceDistance) {
                    finalTime = RESIZESWIM(sourceDistance, dist, finalTime);
                    Logger.log(`[EntryDataProcessor] Converted ${strokeType} time from ${sourceDistance}m to ${dist}m: ${finalTime}`);
                  } else {
                    Logger.log(`[EntryDataProcessor] No conversion mapping found for ${convertFlag} at ${dist}m distance (${strokeType})`);
                  }
                } catch (e) {
                  Logger.log(`[EntryDataProcessor] Conversion skipped due to error for event "${pair.eventName}": ${e && e.message ? e.message : e}`);
                }
              }
            }
          }

          entryTimes.push(finalTime);
        } else {
          // No time column, default to NT
          entryTimes.push("NT");
        }
      }
    } else {
      // METHOD 1: Event value in cell, time in adjacent column
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

          // Convert times from non-standard pools to standard 25m pool timing when requested
          if (timeStr !== "NT") {
            const convertFlag = String(convertTimes || "").trim();
            if (convertFlag !== "" && convertFlag.toLowerCase() !== "no") {
              const dist = parseEventDistance(eventStr);
              if (dist) {
                try {
                  const strokeType = getStrokeType(eventStr);
                  const sourceDistance = getSourceDistance(convertFlag, dist, strokeType);
                  if (sourceDistance) {
                    timeStr = RESIZESWIM(sourceDistance, dist, timeStr);
                    Logger.log(`[EntryDataProcessor] Converted ${strokeType} time from ${sourceDistance}m to ${dist}m: ${timeStr}`);
                  } else {
                    Logger.log(`[EntryDataProcessor] No conversion mapping found for ${convertFlag} at ${dist}m distance (${strokeType})`);
                  }
                } catch (e) {
                  // If anything goes wrong during conversion, keep original timeStr
                  Logger.log(`[EntryDataProcessor] Conversion skipped due to error for event "${eventStr}": ${e && e.message ? e.message : e}`);
                }
              }
            }
          }

          entryTimes.push(timeStr);
        } else {
          // Events-only format: default "NT"
          entryTimes.push("NT");
        }
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
 * Automatically detects column format (METHOD_1 or METHOD_2) and column mapping
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

  // ── AUTO-DETECT COLUMN MAPPING ──────────────────────────────────────
  const headerRow = data[0];
  const columnMapping = detectColumnMapping(headerRow);

  // ── AUTO-DETECT COLUMN FORMAT ───────────────────────────────────────
  const formatDetection = detectColumnFormat(headerRow);

  Logger.log(`[EntryDataProcessor] Detected column format: ${formatDetection.method} (confidence: ${formatDetection.confidence})`);

  // ── AUTO-DETECT Event/Time columns based on format ──────────────────
  let eventTimePairs;
  if (formatDetection.method === 'METHOD_2') {
    // METHOD 2: Event names in headers with "Yes" indicators
    eventTimePairs = detectEventColumnsWithBestTime(headerRow);
    Logger.log(`[EntryDataProcessor] METHOD_2: Detected ${eventTimePairs.length} event columns with best time pairs`);
  } else {
    // METHOD 1: Traditional "Event X" | "Time X" pairs
    eventTimePairs = detectEventTimeColumns(headerRow);
    Logger.log(`[EntryDataProcessor] METHOD_1: Detected ${eventTimePairs.length} event/time pairs`);
  }

  if (eventTimePairs.length === 0) {
    Logger.log('[EntryDataProcessor] WARNING: No event columns detected. Results may be empty.');
  } else {
    Logger.log(`[EntryDataProcessor] Event columns at indices: ${eventTimePairs.map(p => `(${p.eventCol}, ${p.timeCol || 'N/A'})`).join(', ')}`);
  }

  const convertColumnIndex = findConvertColumnIndex(headerRow);

  // --- FETCH TEAM CODE TABLE DATA ONCE (outside the loop to avoid rate limiting)
  let teamCodeData = null;
  try {
    const tableApp = TableApp.openById(spreadsheetId);
    let table = null;
    try {
      table = tableApp.getTableByName(teamCodeTableName);
    } catch (e) {
      Logger.log('[EntryDataProcessor] TableApp failed to find table by name: %s. Falling back to sheet search.', e.message);
    }

    if (table) {
      teamCodeData = table.getValues();
      Logger.log('[EntryDataProcessor] Successfully loaded team code data from table "%s" (%s rows)', teamCodeTableName, teamCodeData.length);
    } else {
      // Fallback: search sheets if TableApp fails or doesn't find the table
      Logger.log('[EntryDataProcessor] Searching sheets for table data: %s', teamCodeTableName);
      const targetSheet = sheet.getParent().getSheetByName(teamCodeTableName);
      if (targetSheet) {
        teamCodeData = targetSheet.getDataRange().getValues();
        Logger.log('[EntryDataProcessor] Successfully loaded team code data from sheet "%s" (%s rows)', teamCodeTableName, teamCodeData.length);
      } else {
        Logger.log('[EntryDataProcessor] Warning: Table or Sheet "%s" not found. All team codes will default to "UNS"', teamCodeTableName);
      }
    }
  } catch (err) {
    Logger.log('[EntryDataProcessor] Error loading team code table: %s. All team codes will default to "UNS"', err.message);
  }

  // Process each data row (skip header)
  const processedRows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const processed = processEntryRow(row, eventTimePairs, convertColumnIndex, teamCodeData, columnMapping);
    if (processed) {
      processedRows.push(processed);
    }
  }

  Logger.log(`[EntryDataProcessor] Successfully processed ${processedRows.length} swimmers using ${formatDetection.method}`);

  return processedRows;
}
