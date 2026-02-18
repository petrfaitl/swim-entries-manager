/**
 * Central configuration for all tables the add-on can create.
 * Each table has:
 *  - headers: ordered array of column display names
 *  - columns: object keyed by header with metadata:
 *      { required, type, default, validation: { type, args } }
 *  - options: table-level options { freezeHeader, headerBg, placement, clearMode }
 *  - notes: optional description
 *
 * Validation types supported by the builder:
 *  - list: args.values (array of strings) OR args.rangeA1 (string) for range-driven lists
 *  - numberRange: args.min, args.max, args.allowInvalid (default false)
 *  - date: args.condition ('DATE_IS_VALID' | 'DATE_BEFORE' | 'DATE_AFTER'), args.value (Date or string)
 *  - checkbox: args.checkedValue, args.uncheckedValue
 *  - text: no validation
 *
 * Placement options:
 *  - { targetSheet: string | null, startCell: string (A1) }
 *  If targetSheet omitted, builder uses table name as the sheet name.
 *
 * clearMode:
 *  - 'rebuild' (default): clears whole sheet/area and rebuilds header and formats
 *  - 'append': keeps existing header/format; positions cursor at first empty row (not yet used by UI)
 */

function getTableConfig(overrides) {
  const DEFAULT_CLUSTERS = [
    'Team One',
    'Team Two',
    'Team Three',
    'Team Four',
    'Team Five',
    'Team Six',
    'Team Seven'
  ];


  const relayDefaults = {
    suggestedEventNames: ['4x25m Girls Freestyle Relay', '4x25m Boys Freestyle Relay', '4x25m Freestyle Grand Relay'],
    defaultRowsPerTable: 4,   // fixed swimmer rows per relay table
    gapBetweenTables: 3       // blank rows between tables when on same sheet
  };

  const YEARS = (overrides && overrides.schoolYears) || ['Y5','Y6','Y7','Y8','Y9','Y10','Y11','Y12','Y13'];
  const GENDERS = (overrides && overrides.genders) || ['Female','Male'];

  /** Utility to build a map from headers for concise config */
  const toCols = (spec) => spec; // pass-through for readability

  const config = {
    TeamOfficials: {
      tableType: 'core',
      title: 'Team Officials*',
      headers: [
        'Cluster or School','Sport Coordinator','Email','Contact Phone','School/Organisation',
        'Team Manager Name','Team Manager Contact','Team Manager Email',
        'Qualified Officials Name','Qualified Officials Email',
        'Timekeeper 1 Name','Timekeeper 2 Name','Timekeeper 3 Name','Other Helpers'
      ],
      columns: toCols({
        'Cluster or School': { type: 'text' },
        'Sport Coordinator': { type: 'text' },
        'Email': { type: 'text' },
        'Contact Phone': { type: 'text' },
        'School/Organisation': { type: 'text' },
        'Team Manager Name': { type: 'text' },
        'Team Manager Contact': { type: 'text' },
        'Team Manager Email': { type: 'text' },
        'Qualified Officials Name': { type: 'text' },
        'Qualified Officials Email': { type: 'text' },
        'Timekeeper 1 Name': { type: 'text' },
        'Timekeeper 2 Name': { type: 'text' },
        'Timekeeper 3 Name': { type: 'text' },
        'Other Helpers': { type: 'text' }
      }),
      options: {
        freezeHeader: 1,
        rows:15,
        headerBg: '#356853',
        title: 'Team Officials',
        required: true,
        placement: { targetSheet: "Team Officials", startCell: 'A1' },
        clearMode: 'rebuild'
      },
      notes: 'All events with detailed information.'
    },

    DetailedEvents: {
      tableType: 'core',
      title: 'Detailed Events',
      headers: ['Events','Gender','Min Age','Max Age','Discipline','Distance','Event Type','Event No.'],
      columns: toCols({
        'Events': { type: 'text' },
        'Gender': { type: 'text',  validation: { type: 'list', args: { values: GENDERS } } },
        'Min Age': { type: 'number', validation: { type: 'numberRange', args: { min: 5, max: 18 } } },
        'Max Age': { type: 'number', validation: { type: 'numberRange', args: { min: 5, max: 18 } } },
        'Discipline': { type: 'text', validation: { type: 'list', args: { values: ['Freestyle','Backstroke','Breaststroke','Butterfly','Individual Medley', 'Relay Medley'] } } },
        'Distance': { type: 'number', validation: { type: 'numberRange', args: { min: 25, max: 400 } } },
        'Event Type': { type: 'text', validation: { type: 'list', args: { values: ['Individual','Relay'] } } },
        'Event No.': { type: 'number' }
      }),
      options: {
        freezeHeader: 1, headerBg: '#356853',title: 'Events for Meet Manager',required:false,
        placement: { targetSheet: "Events", startCell: 'A1' }, clearMode: 'rebuild'
      },
      notes: 'Full events list for Meet Manager alignment.'
    },

    EventsForTemplate: {
      tableType: 'core',
      title: 'Events for Individual Entries Template*',
      headers: ['Distance','Discipline','Events'],
      columns: toCols({
        'Distance': { type: 'text',validation: { type: 'list', args: { values: ['25m','50m','100m','200m','400m'] } } },
        'Discipline': { type: 'text', validation: { type: 'list', args: { values: ['Freestyle','Backstroke','Breaststroke','Butterfly','Individual Medley'] } } },
        'Events': { type: 'text', formula: '=IF(AND(B2<>"",A2<>""),A2&" "&B2,"")' }
      }),
      options: {
        freezeHeader: 1, headerBg: '#356853',
        rows:10,required:true,
        placement: { targetSheet: "EventsForTemplate", startCell: 'A1' }, clearMode: 'rebuild',
        namedRange: { name: 'EventsList', columnName: 'Events' }
      },
      notes: 'Simplified event names used for dropdown validation in individual entries.'
    },

    Schools: {
      tableType: 'core',
      title: 'Schools for Individual Entries Template',
      headers: ['Team Name','School','Cluster','Code'],
      columns: toCols({
        'Team Name': { type: 'text' },
        'School': { type: 'text' },
        'Cluster': { type: 'text', validation: { type: 'list', args: { values: DEFAULT_CLUSTERS } } },
        'Code': { type: 'text' }
      }),
      options: {
        freezeHeader: 1, headerBg: '#356853', rows:30, required:true,
        placement: { targetSheet: "SchoolsForTemplate", startCell: 'A1' }, clearMode: 'rebuild',
        namedRange: { name: 'SchoolsList', columnName: 'School' }
      }
    },

    IndividualEventsTemplate: {
      tableType: 'core',
      title: 'Individual Events Template*',
      headers: ['#','First Name','Last Name','Date of Birth','Gender','School Year','School',
        'Event 1','Time 1 (m:s.S)','Event 2','Time 2 (m:s.S)','Event 3','Time 3 (m:s.S)','Event 4','Time 4 (m:s.S)','Event 5','Time 5 (m:s.S)','Event 6','Time 6 (m:s.S)','Event 7','Time 7 (m:s.S)','Event 8','Time 8 (m:s.S)','Event 9','Time 9 (m:s.S)','Convert times from 33m pool'
      ],
      columns: toCols({
        '#': { type: 'number' },
        'First Name': { type: 'text' },
        'Last Name': { type: 'text' },
        'Date of Birth': { type: 'date', validation: { type: 'date', args: { condition: 'DATE_IS_VALID' } } },
        'Gender': { type: 'text', validation: { type: 'list', args: { values: GENDERS } } },
        'School Year': { type: 'text', validation: { type: 'list', args: { values: YEARS } } },
        'School': { type: 'text', validation: { type: 'range', args: { rangeA1: '=SchoolsList' } }  },
        'Event 1': { type: 'text', validation: { type: 'range', args: { rangeA1: '=EventsList' } } }, 'Time 1 (m:s.S)': { type: 'text' },
        'Event 2': { type: 'text', validation: { type: 'range', args: { rangeA1: '=EventsList' } } }, 'Time 2 (m:s.S)': { type: 'text' },
        'Event 3': { type: 'text', validation: { type: 'range', args: { rangeA1: '=EventsList' } } }, 'Time 3 (m:s.S)': { type: 'text' },
        'Event 4': { type: 'text', validation: { type: 'range', args: { rangeA1: '=EventsList' } } }, 'Time 4 (m:s.S)': { type: 'text' },
        'Event 5': { type: 'text', validation: { type: 'range', args: { rangeA1: '=EventsList' } } }, 'Time 5 (m:s.S)': { type: 'text' },
        'Event 6': { type: 'text', validation: { type: 'range', args: { rangeA1: '=EventsList' } } }, 'Time 6 (m:s.S)': { type: 'text' },
        'Event 7': { type: 'text', validation: { type: 'range', args: { rangeA1: '=EventsList' } } }, 'Time 7 (m:s.S)': { type: 'text' },
        'Event 8': { type: 'text', validation: { type: 'range', args: { rangeA1: '=EventsList' } } }, 'Time 8 (m:s.S)': { type: 'text' },
        'Event 9': { type: 'text', validation: { type: 'range', args: { rangeA1: '=EventsList' } } }, 'Time 9 (m:s.S)': { type: 'text' },
        'Convert times from 33m pool': { type: 'text', validation: { type: 'list', args: { values: ['Yes', ''] } }, default: '' }
      }),
      options: {
        freezeHeader: 1, headerBg: '#356853', title: "Individual Entries Template", rows:30,required: true,
        placement: { targetSheet: 'INDIVIDUAL_EVENTS_TEMPLATE', startCell: 'A1' }, clearMode: 'rebuild'
      },
      notes: 'Template for duplication into per-cluster sheets.'
    },

    // Relay sheet consists of multiple small tables on one sheet.
    // The builder can place individual sections; here we define a single 4x swimmers block header set.
    // NOTE: Team name columns are NOT included in headers — they are dynamically appended at runtime
    // from actual team names provided via the relay table creation dialog.
    RelayEntry: {
      tableType: 'relay',
      title: 'Relays',
      headers: ['Order','School Year','Gender'],  // Team columns appended dynamically at build time
      columns: toCols({
        'Order': { type: 'text',  validation: { type: 'list', args: { values: ['1st Swimmer','2nd Swimmer','3rd Swimmer','4th Swimmer'] } } },
        'School Year': { type: 'text', validation: { type: 'list', args: { values: YEARS } } },
        'Gender': { type: 'text', validation: { type: 'list', args: { values: GENDERS } } }
        // Team name columns are free-text for swimmer names, added dynamically
      }),
      options: {
        freezeHeader: 0, headerBg: '#356853',rows:4, required: false,
        placement: { targetSheet: 'Relays', startCell: 'A1' },
        clearMode: 'rebuild'
      }
    }
  };

  return {
    tables: config,
    relayDefaults: relayDefaults,
    defaultClusters: DEFAULT_CLUSTERS  // Fallback if no team names available
  };
}

/**
 * Utility to list available table names for the UI.
 */
function listAvailableTables(overrides) {
  const cfg = getTableConfig(overrides);
  const tables = Object.keys(cfg.tables).map(function(name) {
    return {
      name: name,
      title: cfg.tables[name].title || name
    };
  });
  Logger.log('[listAvailableTables] Found %s tables: %s', tables.length, JSON.stringify(tables));
  return tables;
}
