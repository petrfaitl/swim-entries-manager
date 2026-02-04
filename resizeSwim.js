/**
 * Function: RESIZESWIM
 *
 * Description:
 * This function adjusts the time taken to swim a given distance proportionally
 * when the target distance changes. It calculates the new time based on the
 * ratio of the target distance to the initial distance. The `initialTime`
 * input can be provided in the format `MM:SS.ss` (minutes and seconds) or
 * `SS.ss` (only seconds). If a number is passed, it is automatically converted
 * to string format before processing.
 *
 * Parameters:
 * @param {number} initialDistance - The initial distance swum (e.g., in meters or yards).
 * @param {number} targetDistance - The target distance to swim.
 * @param {string|number} initialTime - The time taken to swim the initial distance.
 * Format should be either `MM:SS.ss` (e.g., "1:30.50") or `SS.ss` (e.g., "35.33").
 * Numbers (e.g., 35.33) will be automatically converted to strings.
 *
 * Returns:
 * @return {number|null} - The adjusted time (in seconds) to swim the target distance,
 * proportionally scaled. Returns `null` if the input format is invalid or an error
 * occurs.
 *
 * Usage Examples:
 * console.log(RESIZESWIM(33.3, 25, "35.33"));   // Output: 26.76
 * console.log(RESIZESWIM(33.3, 25, "1:35.33")); // Output: 71.62
 * console.log(RESIZESWIM(33.3, 25, 35.33));     // Output: 26.76
 *
 * Error Handling:
 * - If `initialTime` is improperly formatted or invalid (e.g., non-numeric values),
 * the function will log the error and return `null`.
 * - Numeric inputs will be auto-converted to strings.
 */

function RESIZESWIM(initialDistance, targetDistance, initialTime) {
  // Helper function to parse time from MM:SS.ss or SS.ss format into total seconds
  function parseTimeToSeconds(time) {
    if(!time){
      return 0;
    }

    if (typeof time === "number") {
      time = time.toString();
    }

    if (typeof time !== "string") {
      throw new Error("The time should be a string in MM:SS.ss or SS.ss format");
    }

    const parts = time.split(":");
    if (parts.length === 2) {
      // MM:SS.ss format
      const minutes = parseFloat(parts[0]);
      const seconds = parseFloat(parts[1]);
      if (isNaN(minutes) || isNaN(seconds)) {
        return seconds;
        // throw new Error("Invalid time format - unable to parse minutes or seconds.");
      }
      return minutes * 60 + seconds;
    } else if (parts.length === 1) {
      // SS.ss format
      const seconds = parseFloat(time);
      if (isNaN(seconds)) {
        throw new Error("Invalid time format - unable to parse seconds.");
      }
      return seconds;
    } else {
      // Invalid time format
      throw new Error("The time format is invalid. Use MM:SS.ss or SS.ss");
    }

  }

  // Helper function to convert seconds to MM:SS.ss format
  function formatSecondsToTime(seconds) {
    const minutes = Math.floor(seconds / 60);
    const remainingSeconds = (seconds % 60).toFixed(2);
    if (minutes > 0) {
      return `${minutes}:${remainingSeconds.padStart(5, "0")}`;
    } else {
      return `${remainingSeconds}`;
    }
  }

  // Convert initialTime to seconds
  const initialTimeInSeconds = parseTimeToSeconds(initialTime);

  // Adjust the time proportionally to the distance
  const adjustedTimeInSeconds =
    (initialTimeInSeconds / initialDistance) * targetDistance;

  // Convert back to formatted time
  return formatSecondsToTime(parseFloat(adjustedTimeInSeconds.toFixed(2)));
}

// Example usage:
//console.log(RESIZESWIM(66, 50, "35.33")); // Output: 26.76
// console.log(RESIZESWIM(33, 25, "1:35.33")); // Example for MM:SS.ss format
