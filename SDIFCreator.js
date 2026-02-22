/**
 * SDIFCreator — SDIF v3 Meet Entry File Generator
 *
 * Standalone module for Google Apps Script.
 * Can be used as a GAS Library (deploy → Manage Deployments → Library).
 * No external dependencies. Embed in project or import as library.
 *
 * Public API:
 *   generate(config)
 *   generateFromData(swimmerRows, meetInfo, teamLookup, outputFileName)
 *   buildSdifString(swimmerRows, meetInfo, teamLookup)
 *   showDialog()
 *   getMeetInfo(tableName)
 *   getTeamLookup(tableName)
 */

var SDIFCreator = (function() {

  // --- Constants ---
  var SDIF_ORG_CODE    = '8';    // Code 001 — FINA
  var SDIF_FILE_CODE   = '01';   // Code 003 — meet entries
  var SDIF_COUNTRY     = 'NZL';
  var SDIF_MEET_TYPE   = '7';    // Code 005 — Junior
  var SDIF_COURSE      = 'S';    // Code 013 — short course metres
  var SDIF_CITIZEN     = 'NZL';
  var SDIF_VERSION     = '3.0';
  var SDIF_EVENT_AGE   = '20';   // Code 025 — open age
  var SDIF_EVENT_SEX   = 'X';    // Mixed
  var SDIF_CREATOR = "GAS SDIF Creator";
  var SDIF_CREATOR_VERSION = "v1.0";

  var STROKE_CODES_ = {
    'free': '1', 'freestyle': '1', 'free style': '1',
    'back': '2', 'backstroke': '2', 'back stroke': '2',
    'breast': '3', 'breaststroke': '3', 'breast stroke': '3',
    'fly': '4', 'butterfly': '4',
    'im': '5', 'individual medley': '5', 'individual': '5', 'medley': '5'
  };

  var CONFIG_ = {
    rightJustifySeedTime: false
  };

  // --- Public API ---

  /**
   * @param {Object} config
   * @param {string}   config.entriesSheetName   Sheet name containing swimmer entries
   * @param {string}   [config.meetTableName]     Named table for meet info. Default: "Meet"
   * @param {string}   [config.schoolsTableName]  Named table for school lookup. Default: "Schools"
   * @param {string}   [config.outputFileName]    File name for .sd3 output.
   * @param {string}   [config.outputFolderId]    Drive folder ID. Default: root folder
   * @returns {string} Download URL of the .sd3 file
   */
  function generate(config) {
    config = config || {};
    var meetTableName = config.meetTableName || "Meet";
    var schoolsTableName = config.schoolsTableName || "Schools";
    var entriesSheetName = config.entriesSheetName;

    if (!entriesSheetName) {
      throw new Error("SDIFCreator: entriesSheetName is required");
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var entriesSheet = ss.getSheetByName(entriesSheetName);
    if (!entriesSheet) {
      throw new Error("SDIFCreator: Sheet '" + entriesSheetName + "' not found");
    }

    // Use shared processing logic from EntryDataProcessor.js
    Logger.log('[SDIFCreator] Processing entries sheet using shared EntryDataProcessor');
    var processedSwimmers = processEntriesSheet(entriesSheet, schoolsTableName, ss.getId());

    if (!processedSwimmers || processedSwimmers.length === 0) {
       throw new Error("SDIFCreator: No valid swimmer data found in entries sheet");
    }

    Logger.log('[SDIFCreator] Processed %s swimmers', processedSwimmers.length);

    var meetInfo = getMeetInfo(meetTableName);
    var teamLookup = getTeamLookup(schoolsTableName);

    var outputFileName = config.outputFileName;
    if (!outputFileName) {
      var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
      var meetNameClean = (meetInfo.meetName || "Meet").replace(/[^a-z0-9]/gi, '_');
      outputFileName = meetNameClean + "-" + dateStr + ".sd3";
    }

    return generateFromData(processedSwimmers, meetInfo, teamLookup, outputFileName, config.outputFolderId);
  }

  /**
   * @param {Array}    swimmerRows    Array of positional arrays OR plain objects.
   * @param {Object}   meetInfo       Plain object with meet details
   * @param {Object[]} teamLookup     Array of { code, teamName, regionCode }
   * @param {string}   outputFileName File name for Drive output including extension
   * @param {string}   [folderId]     Drive folder ID
   * @returns {string} Download URL
   */
  function generateFromData(swimmerRows, meetInfo, teamLookup, outputFileName, folderId) {
    var sdifString = buildSdifString(swimmerRows, meetInfo, teamLookup);
    return saveToFile_(sdifString, outputFileName, folderId);
  }

  /**
   * @param {Array}    swimmerRows
   * @param {Object}   meetInfo
   * @param {Object[]} teamLookup
   * @returns {string} Complete SDIF file contents as a string
   */
  function buildSdifString(swimmerRows, meetInfo, teamLookup) {
    return buildSdifString_(swimmerRows, meetInfo, teamLookup);
  }

  function showDialog() {
    var html = HtmlService.createHtmlOutputFromFile('SDIFCreator_Dialog')
        .setWidth(450)
        .setHeight(400);
    SpreadsheetApp.getUi().showModalDialog(html, 'Generate SDIF (.sd3) File');
  }

  /**
   * @param {string} tableName  Default: "Meet"
   * @returns {Object} meetInfo plain object
   */
  function getMeetInfo(tableName) {
    tableName = tableName || "Meet";
    return readMeetInfo_(tableName);
  }

  /**
   * @param {string} tableName  Default: "Schools"
   * @returns {Object[]} Array of { code, teamName, regionCode }
   */
  function getTeamLookup(tableName) {
    tableName = tableName || "Schools";
    return buildTeamLookup_(tableName);
  }

  // --- Internal Functions ---

  function buildSdifString_(swimmerRows, meetInfo, teamLookup) {
    var normalisedSwimmers = swimmerRows.map(function(row) {
      return normaliseSwimmerRow_(row);
    }).filter(function(s) { return s !== null; });

    if (normalisedSwimmers.length === 0) {
      throw new Error("SDIFCreator: No valid swimmer records to process");
    }

    // Group by team code
    var teamsMap = {};
    var teamOrder = [];
    normalisedSwimmers.forEach(function(swimmer) {
      if (!teamsMap[swimmer.teamCode]) {
        teamsMap[swimmer.teamCode] = [];
        teamOrder.push(swimmer.teamCode);
      }
      teamsMap[swimmer.teamCode].push(swimmer);
    });

    // Resolve team info
    var teamLookupMap = {};
    teamLookup.forEach(function(t) {
      teamLookupMap[t.code] = t;
    });

    var counts = {
      bRecords: 0,
      cRecords: 0,
      d0Records: 0,
      d1Records: 0,
      swimmers: 0
    };

    var lines = [];

    // A0
    lines.push(emitA0_(meetInfo));

    // B1
    lines.push(emitB1_(meetInfo));
    counts.bRecords++;

    teamOrder.forEach(function(teamCode) {
      var teamInfo = teamLookupMap[teamCode];
      var teamName = teamInfo ? teamInfo.teamName : teamCode;
      var regionCode = teamInfo ? teamInfo.regionCode : "";

      if (!teamInfo) {
        Logger.log('[SDIFCreator] WARNING: Team code %s not found in Schools table. Using defaults.', teamCode);
      }

      // C1
      lines.push(emitC1_(teamCode, teamName, regionCode));
      counts.cRecords++;

      var swimmersInTeam = teamsMap[teamCode];
      swimmersInTeam.forEach(function(swimmer) {
        var events = parseEvents_(swimmer.events, swimmer.times);
        var validEventsCount = 0;

        events.forEach(function(event) {
          if (event.isValid) {
            lines.push(emitD0_(swimmer, event, event.seedTime, meetInfo));
            counts.d0Records++;
            validEventsCount++;
          } else {
            Logger.log('[SDIFCreator] WARNING: Skipping invalid event for %s %s: %s (%s)',
                       swimmer.firstName, swimmer.lastName, event.eventStr, event.errorMessage);
          }
        });

        // D1
        lines.push(emitD1_(swimmer, regionCode));
        counts.d1Records++;
        counts.swimmers++;
      });
    });

    // Z0
    var totalD = counts.d0Records + counts.d1Records;
    lines.push(emitZ0_({
      bRecords: counts.bRecords,
      cRecords: counts.cRecords,
      dRecords: totalD,
      swimmers: counts.swimmers
    }));

    return lines.join('');
  }

  // --- Emitters ---

  function emitA0_(meetInfo) {
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMddyyyy");
    return buildRecord_([
      pad_("A0", 2),
      pad_(SDIF_ORG_CODE, 1),
      pad_(SDIF_VERSION, 8),
      pad_(SDIF_FILE_CODE, 2),
      pad_("", 30),
      pad_(meetInfo.softwareName || SDIF_CREATOR, 20),
      pad_(meetInfo.softwareVersion || SDIF_CREATOR_VERSION, 10),
      pad_(meetInfo.contactName || "", 20),
      pad_(meetInfo.contactPhone || "", 12),
      pad_(today, 8),
      pad_("", 42),
      pad_("", 2), // LSC
      pad_("", 3)
    ]);
  }

  function emitB1_(meetInfo) {
    var startDate = dateToMMDDYYYY_(meetInfo.startDate);
    var endDate = dateToMMDDYYYY_(meetInfo.endDate);

    return buildRecord_([
      pad_("B1", 2),
      pad_(SDIF_ORG_CODE, 1),
      pad_("", 8),
      pad_(meetInfo.meetName || "", 30),
      pad_(meetInfo.venue || "", 22), // Venue
      pad_(meetInfo.address1 || meetInfo.address || "", 22),
      pad_(meetInfo.city || "", 20),
      pad_(meetInfo.region || "", 2),
      pad_(meetInfo.postcode || "", 10),
      pad_(meetInfo.country || SDIF_COUNTRY, 3),
      pad_(meetInfo.meetType || SDIF_MEET_TYPE, 1),
      pad_(startDate, 8),
      pad_(endDate, 8),
      pad_("0", 4, true), // Altitude
      pad_("", 8),
      pad_(meetInfo.course || SDIF_COURSE, 1),
      pad_("", 10)
    ]);
  }

  function emitC1_(teamCode, teamName, regionCode) {
    var fullTeamCode = (regionCode || "").substring(0, 2) + teamCode.substring(0, 4);
    fullTeamCode = pad_(fullTeamCode, 6);

    return buildRecord_([
      pad_("C1", 2),
      pad_(SDIF_ORG_CODE, 1),
      pad_("", 8),
      fullTeamCode,
      pad_(teamName || "", 30),
      pad_("", 16), // Abbr name
      pad_("", 22), // Address 1
      pad_("", 22), // Address 2
      pad_("", 20), // City
      pad_("", 2),  // State
      pad_("", 10), // Postcode
      pad_(SDIF_COUNTRY, 3),
      pad_("", 1),  // Region
      pad_("", 6),
      pad_("", 1),  // 5th team char
      pad_("", 10)
    ]);
  }

  function emitD0_(swimmer, event, seedTime, meetInfo) {
    var mmNumber = buildMMNumber_(swimmer.teamCode, swimmer.lastName, swimmer.firstName, swimmer.dob);
    var fullName = (swimmer.lastName + ", " + swimmer.firstName).substring(0, 28);
    var dob = dateToMMDDYYYY_(swimmer.dob);
    var meetStartDate = dateToMMDDYYYY_(meetInfo.startDate);
    var eventAgeCode = meetInfo.eventAgeCode || "13UN";

    return buildRecord_([
      pad_("D0", 2),
      pad_(SDIF_ORG_CODE, 1),
      pad_("", 8),
      pad_(fullName, 28),
      pad_(mmNumber, 12),
      pad_("", 1), // Attach
      pad_("", 3),
      pad_(dob, 8),
      pad_(swimmer.schoolYear || "", 2),
      pad_(swimmer.gender, 1),
      pad_(SDIF_EVENT_SEX, 1),
      pad_(event.distance, 4),
      pad_(event.strokeCode, 1),
      pad_("", 4), // Event num
      pad_(eventAgeCode, 4),
      pad_(meetStartDate, 8),
      pad_(seedTime, 8, CONFIG_.rightJustifySeedTime),
      pad_(meetInfo.course || SDIF_COURSE, 1),
      pad_("", 63)
    ]);
  }

  function emitD1_(swimmer, regionCode) {
    var mmNumber = buildMMNumber_(swimmer.teamCode, swimmer.lastName, swimmer.firstName, swimmer.dob);
    var fullName = (swimmer.lastName + ", " + swimmer.firstName).substring(0, 28);
    var dob = dateToMMDDYYYY_(swimmer.dob);
    var fullTeamCode = (regionCode || "").substring(0, 2) + swimmer.teamCode.substring(0, 4);
    fullTeamCode = pad_(fullTeamCode, 6);

    return buildRecord_([
      pad_("D1", 2),
      pad_(SDIF_ORG_CODE, 1),
      pad_("", 8),
      fullTeamCode,
      pad_("", 1), // 5th char
      pad_(fullName, 28),
      pad_("", 1), // Future
      pad_(mmNumber, 12),
      pad_("", 1), // Attach
      pad_(SDIF_CITIZEN, 3),
      pad_(dob, 8),
      pad_(swimmer.schoolYear || "", 2),
      pad_(swimmer.gender, 1),
      pad_("", 86)
    ]);
  }

  function emitZ0_(counts) {
    var line = [
      pad_("Z0", 2),
      pad_(SDIF_ORG_CODE, 1),
      pad_("", 8),
      pad_(SDIF_FILE_CODE, 2),
      pad_("", 30),
      pad_(counts.bRecords, 3, true),
      pad_(counts.bRecords, 3, true), // Meets
      pad_(counts.cRecords, 4, true),
      pad_(counts.cRecords, 4, true), // Teams
      pad_(counts.dRecords, 6, true),
      pad_(counts.swimmers, 6, true),
      pad_("", 91)
    ].join('');

    if (line.length !== 160) {
      Logger.log('[SDIFCreator] WARNING: Z0 record length %s != 160', line.length);
    }
    return line; // No CRLF for Z0 as per spec/reference
  }

  function buildRecord_(parts) {
    var line = parts.join('');
    if (line.length !== 160) {
      Logger.log('[SDIFCreator] WARNING: record length %s != 160: "%s"', line.length, line.substring(0, 20));
      // Force it to 160 if we must? Better to fix emitters.
    }
    return line + '\r\n';
  }

  // --- Formatting & Parsing Helpers ---

  function pad_(value, length, rightJustify) {
    var str = String(value === null || value === undefined ? "" : value);
    if (str.length > length) {
      return str.substring(0, length);
    }
    var padding = new Array(length - str.length + 1).join(' ');
    return rightJustify ? padding + str : str + padding;
  }

  function dateToMMDDYYYY_(dateVal) {
    if (!dateVal) return "        ";

    var date;
    if (dateVal instanceof Date) {
      date = dateVal;
    } else if (typeof dateVal === 'number') {
      // Excel/Sheets serial date
      date = new Date((dateVal - 25569) * 86400000);
    } else if (typeof dateVal === 'string') {
      var trimmed = dateVal.trim();
      if (!trimmed) return "        ";

      // Try DD/MM/YYYY
      var parts = trimmed.split('/');
      if (parts.length === 3) {
        var day = parseInt(parts[0], 10);
        var month = parseInt(parts[1], 10);
        var year = parseInt(parts[2], 10);
        date = new Date(year, month - 1, day);
      } else {
        date = new Date(trimmed);
      }
    }

    if (!date || isNaN(date.getTime())) {
      return "        ";
    }

    var mm = pad_(date.getMonth() + 1, 2, true).replace(/ /g, '0');
    var dd = pad_(date.getDate(), 2, true).replace(/ /g, '0');
    var yyyy = pad_(date.getFullYear(), 4, true).replace(/ /g, '0');

    return mm + dd + yyyy;
  }

  function dobToDDMMYY_(dobDDMMYYYY) {
    // Input is usually "DD/MM/YYYY" from normaliseSwimmerRow_
    var parts = dobDDMMYYYY.split('/');
    if (parts.length !== 3) return "000000";

    var dd = pad_(parts[0], 2, true).replace(/ /g, '0');
    var mm = pad_(parts[1], 2, true).replace(/ /g, '0');
    var yy = parts[2].substring(2, 4);

    return dd + mm + yy;
  }

  function buildMMNumber_(teamCode, lastName, firstName, dob) {
    var lastNameChar = (lastName && lastName[0] ? lastName[0] : 'X').toUpperCase();
    var firstNameChar = (firstName && firstName[0] ? firstName[0] : 'X').toUpperCase();
    var dobPart = dobToDDMMYY_(dob);

    return teamCode.substring(0, 3).toUpperCase()
      + lastNameChar
      + firstNameChar
      + 'Z'
      + dobPart;
  }

  function parseSeedTime_(timeStr) {
    if (!timeStr) return pad_("NT", 8);
    var s = timeStr.trim().toUpperCase();
    if (s === "NT" || s === "") return pad_("NT", 8);

    // Strip whitespace
    s = s.replace(/\s/g, '');

    var minutes = 0;
    var seconds = 0;
    var decimals = 0;

    if (s.indexOf(':') !== -1) {
      var parts = s.split(':');
      if (parts.length === 3) { // hh:mm:ss.dd or mm:ss:dd
         // Assuming mm:ss:dd if it's swimming
         minutes = parseInt(parts[0], 10);
         seconds = parseInt(parts[1], 10);
         decimals = parseInt(parts[2], 10);
      } else if (parts.length === 2) { // mm:ss.dd or ss:dd
         if (parts[1].indexOf('.') !== -1) {
           minutes = parseInt(parts[0], 10);
           var ssdd = parts[1].split('.');
           seconds = parseInt(ssdd[0], 10);
           decimals = parseInt(ssdd[1], 10);
         } else {
           // ss:dd or mm:ss?
           // If parts[0] > 59, it's probably ss:dd
           var p0 = parseInt(parts[0], 10);
           if (p0 > 59) {
             seconds = p0;
             decimals = parseInt(parts[1], 10);
           } else {
             minutes = p0;
             seconds = parseInt(parts[1], 10);
           }
         }
      }
    } else if (s.indexOf('.') !== -1) {
      var parts = s.split('.');
      seconds = parseInt(parts[0], 10);
      decimals = parseInt(parts[1], 10);
    } else {
      seconds = parseInt(s, 10);
    }

    if (isNaN(minutes)) minutes = 0;
    if (isNaN(seconds)) seconds = 0;
    if (isNaN(decimals)) decimals = 0;

    var mm = pad_(minutes, 2, true).replace(/ /g, '0');
    var ss = pad_(seconds, 2, true).replace(/ /g, '0');
    var dd = pad_(decimals, 2).replace(/ /g, '0');

    return mm + ":" + ss + "." + dd;
  }

  function parseEvents_(eventsStr, timesStr) {
    if (!eventsStr) return [];
    var events = eventsStr.split(',').map(function(s) { return s.trim(); });
    var times = (timesStr || "").split(',').map(function(s) { return s.trim(); });

    return events.map(function(eventStr, i) {
      var time = times[i] || "NT";
      var parsed = parseDistanceAndStroke_(eventStr);
      parsed.eventStr = eventStr;
      parsed.seedTime = parseSeedTime_(time);
      return parsed;
    });
  }

  function parseDistanceAndStroke_(eventStr) {
    var s = eventStr.trim().replace(/,$/, '');
    var match = s.match(/^(\d+)(m)?\s+(.+)$/i);
    if (!match) {
      return { distance: "", strokeCode: "", isValid: false, errorMessage: "Unrecognised event format" };
    }

    var distance = match[1];
    var strokeName = match[3].toLowerCase();
    var strokeCode = getStrokeCode_(strokeName);

    if (!strokeCode) {
       return { distance: distance, strokeCode: "", isValid: false, errorMessage: "Unrecognised stroke: " + strokeName };
    }

    return { distance: distance, strokeCode: strokeCode, isValid: true };
  }

  function getStrokeCode_(strokeName) {
    return STROKE_CODES_[strokeName] || null;
  }

  function normaliseGender_(raw) {
    if (!raw) return "X";
    var s = String(raw).trim().toUpperCase();
    if (s === "M" || s === "MALE" || s === "B" || s === "BOY") return "M";
    if (s === "F" || s === "FEMALE" || s === "G" || s === "GIRL") return "F";
    Logger.log('[SDIFCreator] WARNING: Unrecognised gender "%s", using "X"', raw);
    return "X";
  }

  function normaliseSwimmerRow_(row) {
    var s = {};

    // Handle both old format (arrays) and new format (objects from EntryDataProcessor)
    if (Array.isArray(row)) {
      if (row.length < 8) return null;
      s.teamCode   = String(row[0]);
      s.firstName  = String(row[1]);
      s.lastName   = String(row[2]);
      s.dob        = String(row[3]);
      s.gender     = String(row[4]);
      s.events     = String(row[5]);
      s.times      = String(row[6]);
      s.schoolYear = String(row[7]);
    } else {
      // New format from EntryDataProcessor
      s.teamCode   = row.teamCode;
      s.firstName  = row.firstName;
      s.lastName   = row.lastName;
      s.dob        = row.dobFormatted;  // Already formatted as DD/MM/YYYY
      s.gender     = row.gender;
      s.events     = row.eventsStr;     // Already formatted as comma-separated string
      s.times      = row.timesStr;      // Already formatted as comma-separated string
      s.schoolYear = row.schoolYear;    // Already normalized
    }

    if (!s.firstName || !s.lastName) {
      Logger.log('[SDIFCreator] ERROR: Missing swimmer name: %s', JSON.stringify(s));
      return null;
    }

    s.gender = normaliseGender_(s.gender);

    // DOB formatting - only needed for old array format
    // New format already has dobFormatted as DD/MM/YYYY
    if (s.dob instanceof Date) {
       s.dob = Utilities.formatDate(s.dob, Session.getScriptTimeZone(), "dd/MM/yyyy");
    } else if (typeof s.dob === 'number') {
       var d = new Date((s.dob - 25569) * 86400000);
       s.dob = Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
    }

    return s;
  }

  function readTableByName_(tableName, spreadsheetId) {
    var ss = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();

    // Try TableApp if available
    if (typeof TableApp !== 'undefined' && TableApp.openById) {
      try {
        var tableApp = TableApp.openById(ss.getId());
        var table = tableApp.getTableByName(tableName);
        if (table) {
          var values = table.getValues();
          if (values && values.length > 0) {
            var headers = values[0];
            var rows = [];
            for (var i = 1; i < values.length; i++) {
              var rowObj = {};
              for (var j = 0; j < headers.length; j++) {
                rowObj[headers[j]] = values[i][j];
              }
              rows.push(rowObj);
            }
            return rows;
          }
        }
      } catch (e) {
        Logger.log('[SDIFCreator] TableApp failed, falling back to sheet reading: %s', e.message);
      }
    }

    // Fallback: search sheets
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName() === tableName) {
        var values = sheets[i].getDataRange().getValues();
        if (values.length > 0) {
          var headers = values[0];
          var rows = [];
          for (var r = 1; r < values.length; r++) {
            var rowObj = {};
            for (var c = 0; c < headers.length; c++) {
              rowObj[headers[c]] = values[r][c];
            }
            rows.push(rowObj);
          }
          return rows;
        }
      }
    }

    return [];
  }

  function readMeetInfo_(tableName) {
    var rows = readTableByName_(tableName);
    if (rows.length === 0) {
      throw new Error("SDIFCreator: Meet table not found or empty");
    }

    var row = rows[0];
    var info = {};

    // Map with case-insensitive keys
    Object.keys(row).forEach(function(key) {
      var k = key.toLowerCase().replace(/\s/g, '');
      if (k === 'meetname') info.meetName = row[key];
      else if (k === 'startdate') info.startDate = row[key];
      else if (k === 'enddate') info.endDate = row[key];
      else if (k === 'venue') info.venue = row[key];
      else if (k === 'venueaddress1') info.address1 = row[key];
      else if (k === 'city') info.city = row[key];
      else if (k === 'region') info.region = row[key];
      else if (k === 'postcode') info.postcode = row[key];
      else if (k === 'country') info.country = row[key];
      else if (k === 'course') info.course = row[key];
      else if (k === 'contactname') info.contactName = row[key];
      else if (k === 'contactphone') info.contactPhone = row[key];
      else if (k === 'meettype') info.meetType = row[key];
      else if (k === 'eventagecode') info.eventAgeCode = row[key];
    });

    return info;
  }

  function buildTeamLookup_(tableName) {
    var rows = readTableByName_(tableName);
    return rows.map(function(row) {
      var info = { code: "", teamName: "", regionCode: "" };
      Object.keys(row).forEach(function(key) {
        var k = key.toLowerCase().replace(/\s/g, '');
        if (k === 'code') info.code = String(row[key]);
        else if (k === 'teamname') info.teamName = String(row[key]);
        else if (k === 'regioncode') info.regionCode = String(row[key]);
        else if (k === 'school' && !info.teamName) info.teamName = String(row[key]);
      });
      return info;
    });
  }

  function saveToFile_(content, fileName, folderId) {
    var blob = Utilities.newBlob(content, 'text/plain', fileName);
    var folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
    var file = folder.createFile(blob);
    return file.getDownloadUrl();
  }

  // --- Exposed Namespace ---
  return {
    generate: generate,
    generateFromData: generateFromData,
    buildSdifString: buildSdifString,
    showDialog: showDialog,
    getMeetInfo: getMeetInfo,
    getTeamLookup: getTeamLookup
  };

})();

/**
 * Server-side function called from dialog.
 */
function processSheetToSd3(sheetName, meetTableName, schoolsTableName, fileName) {
  return SDIFCreator.generate({
    entriesSheetName: sheetName,
    meetTableName: meetTableName,
    schoolsTableName: schoolsTableName,
    outputFileName: fileName
  });
}

/**
 * Smoke test for SDIFCreator
 */
function SDIFCreator_runSmokeTest() {
  var meetInfo = {
    meetName: "Test Meet",
    startDate: "22/02/2026",
    endDate: "22/02/2026",
    region: "BP",
    course: "S"
  };

  var swimmerRows = [
    ["TPS", "Niko", "Smith", "10/01/2013", "M", "25m Freestyle", "NT", "Y5"],
    ["TPS", "Test", "Swimmer", "01/01/2014", "F", "50m Freestyle, 100m Backstroke", "29.50, 1:02.34", "Y6"]
  ];

  var teamLookup = [
    { code: "TPS", teamName: "Test Primary School", regionCode: "BP" }
  ];

  var sdif = SDIFCreator.buildSdifString(swimmerRows, meetInfo, teamLookup);
  var lines = sdif.split('\r\n').filter(function(l) { return l.length > 0; });

  Logger.log("--- Smoke Test Output ---");
  lines.forEach(function(line, i) {
    Logger.log("Line %s (%s chars): %s", i+1, line.length, line);
    if (line.length !== 160 && line.substring(0, 2) !== "Z0") {
       Logger.log("ERROR: Line %s length is not 160!", i+1);
    }
  });

  // Basic assertions
  if (lines[0].substring(0, 2) !== "A0") throw new Error("Line 1 should be A0");
  if (lines[1].substring(0, 2) !== "B1") throw new Error("Line 2 should be B1");
  if (lines[2].substring(0, 2) !== "C1") throw new Error("Line 3 should be C1");

  // Niko Davis MM number: TPSDNZ100415
  var d0_niko = lines[3];
  var mmNumber = d0_niko.substring(39, 51);
  Logger.log("Niko MM Number: %s", mmNumber);
  if (mmNumber !== "TPSDNZ100415") throw new Error("Incorrect MM number for Niko: " + mmNumber);

  // Distance 25m -> "25  "
  var distance = d0_niko.substring(67, 71);
  if (distance !== "25  ") throw new Error("Incorrect distance: '" + distance + "'");

  // Seed time NT -> "NT      "
  var seedTime = d0_niko.substring(88, 96);
  if (seedTime !== "NT      ") throw new Error("Incorrect seed time: '" + seedTime + "'");

  // Z0 check
  var z0 = lines[lines.length - 1];
  if (z0.substring(0, 2) !== "Z0") throw new Error("Last line should be Z0");

  Logger.log("Smoke test passed!");
  return "Smoke test passed!";
}
