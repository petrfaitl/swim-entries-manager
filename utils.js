/**
 * Formats DOB to DD/MM/YYYY
 * Handles: Date objects, numbers (serial date), strings in various formats
 * @param {any} dob - Raw DOB value from sheet
 * @return {string} Formatted date or empty string if invalid
 */
function formatDob(dob) {
  if (!dob) return "";

  // Case 1: Already a Date object (common when cell is formatted as Date)
  if (dob instanceof Date) {
    const day   = String(dob.getDate()).padStart(2, '0');
    const month = String(dob.getMonth() + 1).padStart(2, '0'); // months are 0-indexed
    const year  = dob.getFullYear();
    return `${day}/${month}/${year}`;
  }

  // Case 2: Number (Google Sheets serial date number)
  if (typeof dob === 'number' && !isNaN(dob)) {
    // Convert serial date to JS Date
    const date = new Date((dob - 25569) * 86400 * 1000); // 25569 = days from 1900-01-01 to 1970-01-01
    const day   = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year  = date.getFullYear();
    return `${day}/${month}/${year}`;
  }

  // Case 3: String – try to parse common formats
  if (typeof dob === 'string') {
    const cleaned = dob.trim();

    // Try parsing as DD/MM/YYYY or D/M/YY etc.
    const parts = cleaned.split(/[-\/.]/); // split on -, /, .
    if (parts.length === 3) {
      let [d, m, y] = parts.map(p => p.trim());
      
      // Handle short year (e.g. 25 → 2025)
      if (y.length === 2) {
        y = (parseInt(y, 10) < 50 ? "20" : "19") + y;
      }
      
      // Try to make valid numbers
      const day   = parseInt(d, 10);
      const month = parseInt(m, 10);
      const year  = parseInt(y, 10);

      if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
        return `${String(day).padStart(2, '0')}/${String(month).padStart(2, '0')}/${year}`;
      }
    }

    // Fallback: return original if we can't parse
    return cleaned;
  }

  // Fallback for anything else
  return dob.toString().trim();
}


/**
 * Formats time coming from Google Sheets cell → "mm:ss.SS"
 * Handles: numbers, strings in many formats, and Date objects from Sheets
 */
function formatAsMmSsSs(input) {
  // ── 0. Handle Empty/Null ──────────────────────────────────────────────
  if (input === "" || input == null) {
    return "NT";
  }

  let totalSeconds = 0;

  // ── 1. Handle Date object from Sheets time/duration cells ───────────
  if (input instanceof Date) {
    const hours   = input.getHours();
    const minutes = input.getMinutes();
    const seconds = input.getSeconds();
    const ms      = input.getMilliseconds();
    // Note: This resets every 24h. For strict durations > 24h, 
    // you would need the raw cell value, but this works for time-of-day.
    totalSeconds = (hours * 3600) + (minutes * 60) + seconds + (ms / 1000);
  }

  // ── 2. Handle STRICT Numeric Input (Primitive Numbers Only) ─────────
  // We use `typeof` to ensure we don't accidentally grab strings like "1:02"
  else if (typeof input === 'number') {
    let num = input;
    // Heuristic: If < 1, Sheets treats it as a fraction of a day (e.g. 0.5 = 12:00pm)
    // If > 1, we assume it is raw seconds.
    totalSeconds = (num < 1 && num > 0) ? num * 86400 : num;
  }

  // ── 3. Handle String Inputs ─────────────────────────────────────────
  else if (typeof input === 'string') {
    const clean = input.trim().replace(/\s+/g, ''); // Remove whitespace

    // Case A: Contains colon → parse as mm:ss[.sss]
    if (clean.includes(':')) {
      const parts = clean.split(':');
      
      // Safety: If structure isn't x:y, return original
      if (parts.length < 2) return clean; 

      const minStr = parts[0];
      const secStr = parts[1];

      const minutes = parseInt(minStr, 10) || 0;
      
      // Simplify logic: Let parseFloat handle "02.3", "2.300", etc.
      const secondsRaw = parseFloat(secStr) || 0;

      totalSeconds = (minutes * 60) + secondsRaw;
    }

    // Case B: No colon, but looks like a number (e.g. "59.99", "123.45")
    // We check !isNaN to ensure it is actually a number
    else if (!isNaN(parseFloat(clean))) {
      totalSeconds = parseFloat(clean);
    }

    // Fallback: Unknown format (e.g. "NT", "DNF")
    else {
      return clean;
    }
  }
  
  // Fallback for objects/arrays that aren't Dates
  else {
    return String(input);
  }

  // ── 4. Final Formatting to "mm:ss.SS" ───────────────────────────────
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = totalSeconds % 60;

  // .toFixed(2) handles the rounding/padding of the decimal part automatically
  const secFormatted = seconds.toFixed(2).padStart(5, '0');

  // PadStart ensures 01:05.00 format
  return `${minutes.toString().padStart(2, '0')}:${secFormatted}`;
}

function testFlexibleTimeParsing() {
  const tests = [
    "1:02.3", "01:02.30", "1:2.3", "01:02.300", "1:2", "01:2.3",
    "0:59.99", "59.99", "123.45", "2:3", "2:03", "2:3.0",
    "10:5.678", "0:0.05", "", "NT already"
  ];

  tests.forEach(t => {
    const result = formatAsMmSsSs(t);
    Logger.log(`"${t}" → "${result}"`);
  });
}
