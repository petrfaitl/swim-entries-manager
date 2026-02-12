/**
 * Central configuration for all tables the add-on can create.
 * Each table has:
 *  - headers: ordered array of column display names
 *  - columns: object keyed by header with metadata:
 *      { required, type, default, validation: { type, args } }
 *  - options: table-level options { freezeHeader, boldHeader, headerBg, alternatingColors, namedRange, placement, clearMode }
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
    'School One',
    'School Two',
    'School Three',
    'School Four',
    'School Five',
    'School Six',
    'School Seven'
  ];

  const WBOP_CLUSTERS = [
    'WBOP - East Cluster',
    'WBOP - Mount Cluster',
    'WBOP - North Cluster',
    'WBOP - Papamoa Cluster',
    'WBOP - South Cluster',
    'WBOP - Te Puke Cluster',
    'WBOP - West Cluster'
  ];

  const YEARS = (overrides && overrides.schoolYears) || ['Y5','Y6','Y7','Y8','Y9','Y10','Y11','Y12','Y13'];
  const GENDERS = (overrides && overrides.genders) || ['Female','Male'];

  /** Utility to build a map from headers for concise config */
  const toCols = (spec) => spec; // pass-through for readability

  const config = {
    TeamOfficials: {
      headers: [
        'Cluster/School','Sport Coordinator','Email','Contact Phone','School/Organisation',
        'Team Manager Name','Team Manager Contact','Team Manager Email',
        'Qualified Officials Name','Qualified Officials Email',
        'Timekeeper 1 Name','Timekeeper 2 Name','Timekeeper 3 Name','Timekeeper 4 Name'
      ],
      columns: toCols({
        'Cluster/School': { required: true, type: 'text', default: DEFAULT_CLUSTERS[0], validation: { type: 'list', args: { values: DEFAULT_CLUSTERS } } },
        'Sport Coordinator': { type: 'text', required: false },
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
        'Timekeeper 4 Name': { type: 'text' }
      }),
      options: {
        freezeHeader: 1,
        boldHeader: true,
        headerBg: '#e3f2fd',
        alternatingColors: true,
        namedRange: 'AllEventsWithDetailsTable',
        placement: { targetSheet: null, startCell: 'A1' },
        clearMode: 'rebuild'
      },
      notes: 'All events with detailed information.'
    },

    DetailedEvents: {
      headers: ['Events','Gender','Min Age','Max Age','Discipline','Distance','Event Type','Event No.'],
      columns: toCols({
        'Events': { type: 'text', required: true },
        'Gender': { type: 'text', required: true, validation: { type: 'list', args: { values: GENDERS } } },
        'Min Age': { type: 'number', validation: { type: 'numberRange', args: { min: 5, max: 18 } } },
        'Max Age': { type: 'number', validation: { type: 'numberRange', args: { min: 5, max: 18 } } },
        'Discipline': { type: 'text', validation: { type: 'list', args: { values: ['Freestyle','Backstroke','Breaststroke','Butterfly','Individual Medley', 'Relay Medley'] } } },
        'Distance': { type: 'number', validation: { type: 'numberRange', args: { min: 25, max: 400 } } },
        'Event Type': { type: 'text', validation: { type: 'list', args: { values: ['Individual','Relay'] } } },
        'Event No.': { type: 'number' }
      }),
      options: {
        freezeHeader: 1, boldHeader: true, headerBg: '#fff3e0', alternatingColors: true,
        namedRange: 'EventsTable', placement: { targetSheet: null, startCell: 'A1' }, clearMode: 'rebuild'
      },
      notes: 'Full events list for Meet Manager alignment.'
    },

    EventsForTemplate: {
      headers: ['Events','Discipline','Distance'],
      columns: toCols({
        'Events': { type: 'text', required: true },
        'Discipline': { type: 'text' },
        'Distance': { type: 'number' }
      }),
      options: {
        freezeHeader: 1, boldHeader: true, headerBg: '#ede7f6', alternatingColors: true,
        namedRange: 'EventsForTemplateTable', placement: { targetSheet: null, startCell: 'A1' }, clearMode: 'rebuild'
      },
      notes: 'Simplified event names used for dropdown validation in individual entries.'
    },

    Schools: {
      headers: ['Team Name','School','Cluster','Code'],
      columns: toCols({
        'Team Name': { type: 'text', required: true },
        'School': { type: 'text', required: true },
        'Cluster': { type: 'text', validation: { type: 'list', args: { values: DEFAULT_CLUSTERS } } },
        'Code': { type: 'text' }
      }),
      options: {
        freezeHeader: 1, boldHeader: true, headerBg: '#e8f5e9', alternatingColors: true,
        namedRange: 'Teams', placement: { targetSheet: null, startCell: 'A1' }, clearMode: 'rebuild'
      }
    },

    IndividualEventsTemplate: {
      headers: ['#','Team Code','First Name','Last Name','Date of Birth','Gender','School Year','School',
        'Event 1','Time 1 (m:s.S)','Event 2','Time 2 (m:s.S)','Event 3','Time 3 (m:s.S)','Event 4','Time 4 (m:s.S)','Event 5','Time 5 (m:s.S)','Event 6','Time 6 (m:s.S)','Event 7','Time 7 (m:s.S)','Event 8','Time 8 (m:s.S)','Event 9','Time 9 (m:s.S)','Convert times from 33m pool'
      ],
      columns: toCols({
        '#': { type: 'number' },
        'Team Code': { type: 'text' },
        'First Name': { type: 'text', required: true },
        'Last Name': { type: 'text', required: true },
        'Date of Birth': { type: 'date', validation: { type: 'date', args: { condition: 'DATE_IS_VALID' } } },
        'Gender': { type: 'text', validation: { type: 'list', args: { values: GENDERS } } },
        'School Year': { type: 'text', validation: { type: 'list', args: { values: YEARS } } },
        'School': { type: 'text' },
        'Event 1': { type: 'text' }, 'Time 1 (m:s.S)': { type: 'text' },
        'Event 2': { type: 'text' }, 'Time 2 (m:s.S)': { type: 'text' },
        'Event 3': { type: 'text' }, 'Time 3 (m:s.S)': { type: 'text' },
        'Event 4': { type: 'text' }, 'Time 4 (m:s.S)': { type: 'text' },
        'Event 5': { type: 'text' }, 'Time 5 (m:s.S)': { type: 'text' },
        'Event 6': { type: 'text' }, 'Time 6 (m:s.S)': { type: 'text' },
        'Event 7': { type: 'text' }, 'Time 7 (m:s.S)': { type: 'text' },
        'Event 8': { type: 'text' }, 'Time 8 (m:s.S)': { type: 'text' },
        'Event 9': { type: 'text' }, 'Time 9 (m:s.S)': { type: 'text' },
        'Convert times from 33m pool': { type: 'checkbox', validation: { type: 'checkbox', args: { checkedValue: 'TRUE', uncheckedValue: 'FALSE' } }, default: 'FALSE' }
      }),
      options: {
        freezeHeader: 1, boldHeader: true, headerBg: '#f3e5f5', alternatingColors: true,
        namedRange: 'IndividualEventsTemplate', placement: { targetSheet: 'INDIVIDUAL_EVENTS_TEMPLATE', startCell: 'A1' }, clearMode: 'rebuild'
      },
      notes: 'Template for duplication into per-cluster sheets.'
    },

    // Relay sheet consists of multiple small tables on one sheet.
    // The builder can place individual sections; here we define a single 4x swimmers block header set.
    RelayEntryBlock: {
      headers: ['Order','School Year','Gender'].concat(DEFAULT_CLUSTERS),
      columns: toCols({
        'Order': { type: 'text', required: true, validation: { type: 'list', args: { values: ['1st Swimmer','2nd Swimmer','3rd Swimmer','4th Swimmer'] } } },
        'School Year': { type: 'text', validation: { type: 'list', args: { values: YEARS } } },
        'Gender': { type: 'text', validation: { type: 'list', args: { values: ['Boys','Girls'] } } }
        // Cluster columns are free-text to type swimmer names or can later be validated to Schools list
      }),
      options: {
        freezeHeader: 1, boldHeader: true, headerBg: '#e1f5fe', alternatingColors: true,
        namedRange: null,
        placement: { targetSheet: 'Relays', startCell: 'A1' },
        clearMode: 'rebuild'
      }
    }
  };

  return config;
}

/**
 * Utility to list available table names for the UI.
 */
function listAvailableTables(overrides) {
  const cfg = getTableConfig(overrides);
  const names = Object.keys(cfg);
  Logger.log('[listAvailableTables] Found %s tables: %s', names.length, JSON.stringify(names));
  return names;
}
