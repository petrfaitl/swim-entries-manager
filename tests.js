function debugSheetId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0]; // first visible sheet
  Logger.log("Sheet name: " + sheet.getName());
  Logger.log("Sheet ID:   " + sheet.getSheetId());
}

function testFlexibleTimeParsing() {
  const tests = [
    "1:02.3", "01:02.30", "1:2.3", "01:02.300", "1:2", "01:2.3",
    "0:59.99", "59.99", "123.45", "2:3", "2:03", "2:3.0",
    "10:5.678", "0:0.05", "", "NT already",
    "1.02.34", "0.59.99", "1.2.3",
    "0.17", 0.17, "0.59", 0.59, "0.005"
  ];

  tests.forEach(t => {
    const result = formatAsMmSsSs(t);
    Logger.log(`"${t}" → "${result}"`);
  });
}

function test_getNamedTableData() {
  // Assuming a named table exists in your sheet with the name "MyTable"
  const tableName = "Teams";
  const spreadsheetId = "1KFNv7LIAyrDPl4Na8XNsCMXJCg_BDqyuUo_q3lwEklU";
  const data = getNamedTableData(spreadsheetId,tableName);
  // const data = getTableByName(tableName);

  // Process or manipulate table data as needed
  Logger.log(`First cell of the table: ${data[0][0]}`);
}
//test_getNamedTableData();



function testCompetitionTypesRegistry() {
  const types = getCompetitionTypes();
  const expectedKeys = ['SCHOOL', 'COLLEGE', 'AGE_GROUP', 'OPEN', 'MASTERS'];
  expectedKeys.forEach(key => {
    if (!types[key]) throw new Error(`Missing competition type: ${key}`);
    if (!types[key].hasAgeUpDate) throw new Error(`Competition type ${key} should have hasAgeUpDate set to true`);
  });

  if (types.SCHOOL.sdifCode !== '7' || !types.SCHOOL.hasSchoolYear) {
    throw new Error('Invalid SCHOOL configuration');
  }
  if (types.COLLEGE.sdifCode !== '7' || !types.COLLEGE.hasSchoolYear) {
    throw new Error('Invalid COLLEGE configuration');
  }
  if (types.AGE_GROUP.sdifCode !== '1' || types.AGE_GROUP.hasSchoolYear) {
    throw new Error('Invalid AGE_GROUP configuration');
  }
  if (types.OPEN.sdifCode !== '0' || types.OPEN.hasSchoolYear) {
    throw new Error('Invalid OPEN configuration');
  }
  if (types.MASTERS.sdifCode !== '8' || types.MASTERS.hasSchoolYear) {
    throw new Error('Invalid MASTERS configuration');
  }

  Logger.log('✓ testCompetitionTypesRegistry passed');
  return true;
}

function testAgeBracketsTableSchema() {
  const brackets = getDefaultAgeBrackets();
  if (!brackets || !brackets.school || !brackets.college) {
    throw new Error('DEFAULT_AGE_BRACKETS is missing school or college definitions');
  }
  if (brackets.school.length !== 4 || brackets.school[0].schoolYear !== 'Y5' || brackets.school[0].minAge !== 9 || brackets.school[0].maxAge !== 9) {
    throw new Error('DEFAULT_AGE_BRACKETS school mapping invalid');
  }
  if (brackets.college.length !== 3 || brackets.college[1].schoolYear !== 'Intermediate' || brackets.college[1].minAge !== 14 || brackets.college[1].maxAge !== 15) {
    throw new Error('DEFAULT_AGE_BRACKETS college mapping invalid');
  }

  const schoolCfg = getTableConfig({ competitionType: 'school' });
  const ageBracketsTable = schoolCfg.tables.AgeBrackets;
  if (!ageBracketsTable) {
    throw new Error('AgeBrackets table definition missing from tableConfig');
  }

  const expectedHeaders = ['School Year', 'Min Age', 'Max Age'];
  if (JSON.stringify(ageBracketsTable.headers) !== JSON.stringify(expectedHeaders)) {
    throw new Error('AgeBrackets table headers mismatch: ' + JSON.stringify(ageBracketsTable.headers));
  }

  if (ageBracketsTable.options.placement.targetSheet !== 'Meet' || (ageBracketsTable.options.placement.startCell !== 'A5' && ageBracketsTable.options.placement.startCell !== 'A4')) {
    throw new Error('AgeBrackets placement mismatch: ' + JSON.stringify(ageBracketsTable.options.placement));
  }

  const minAgeCol = ageBracketsTable.columns['Min Age'];
  const maxAgeCol = ageBracketsTable.columns['Max Age'];
  if (!minAgeCol || minAgeCol.type !== 'number' || !minAgeCol.validation || minAgeCol.validation.type !== 'numberRange') {
    throw new Error('AgeBrackets Min Age column configuration invalid');
  }
  if (!maxAgeCol || maxAgeCol.type !== 'number' || !maxAgeCol.validation || maxAgeCol.validation.type !== 'numberRange') {
    throw new Error('AgeBrackets Max Age column configuration invalid');
  }

  Logger.log('✓ testAgeBracketsTableSchema passed');
  return true;
}

function testTableConfigGeneration() {
  // Test School
  const schoolCfg = getTableConfig({ competitionType: 'school' });
  const schoolHeaders = schoolCfg.tables.IndividualEventsTemplate.headers;
  if (!schoolHeaders.includes('School Year') || !schoolHeaders.includes('School')) {
    throw new Error('School template headers missing School Year or School');
  }
  if (schoolHeaders.includes('Team')) {
    throw new Error('School template should not include Team header');
  }

  // Test Age Group
  const ageGroupCfg = getTableConfig({ competitionType: 'age_group' });
  const ageGroupHeaders = ageGroupCfg.tables.IndividualEventsTemplate.headers;
  if (ageGroupHeaders.includes('School Year')) {
    throw new Error('Age Group template should not include School Year');
  }
  if (!ageGroupHeaders.includes('Team')) {
    throw new Error('Age Group template should include Team header');
  }

  // Test Meet table schema
  const meetHeaders = schoolCfg.tables.Meet.headers;
  if (!meetHeaders.includes('Competition Type') || !meetHeaders.includes('Age-up Date')) {
    throw new Error('Meet table headers missing Competition Type or Age-up Date');
  }
  const ageUpCol = schoolCfg.tables.Meet.columns['Age-up Date'];
  if (!ageUpCol || ageUpCol.type !== 'date' || !ageUpCol.validation || ageUpCol.validation.args.condition !== 'DATE_IS_VALID') {
    throw new Error('Meet table Age-up Date column configuration invalid or missing date validation');
  }

  Logger.log('✓ testTableConfigGeneration passed');
  return true;
}

function testSchoolAndTeamCodeLookup() {
  const schoolsData = [
    ['Team Name', 'School', 'Cluster', 'Code', 'Region Code'],
    ['ACG Tauranga', 'ACG Tauranga', 'Cluster 1', 'ACGT', 'BP'],
    ['Bethlehem College', 'Bethlehem College', 'Cluster 1', 'BCOL', 'BP']
  ];
  if (getSchoolCode(schoolsData, 'ACG Tauranga') !== 'ACGT') {
    throw new Error('Failed to lookup school code in Schools table');
  }

  const teamsData = [
    ['Team Name', 'Team', 'Code', 'Region Code'],
    ['Mount Masters', 'Mount Masters', 'MMST', 'BP'],
    ['Otumoetai Swimming Club', 'Otumoetai', 'OTUM', 'BP']
  ];
  if (getSchoolCode(teamsData, 'Otumoetai') !== 'OTUM') {
    throw new Error('Failed to lookup team code in Teams table');
  }

  Logger.log('✓ testSchoolAndTeamCodeLookup passed');
  return true;
}

function testSDIFMeetTypeResolution() {
  const compTypes = [
    { type: 'School', expectedCode: '7' },
    { type: 'College', expectedCode: '7' },
    { type: 'Age Group', expectedCode: '1' },
    { type: 'Open', expectedCode: '0' },
    { type: 'Masters', expectedCode: '8' }
  ];

  compTypes.forEach(({ type, expectedCode }) => {
    const meetInfo = {
      meetName: `${type} Championship`,
      competitionType: type,
      startDate: '08/11/2026',
      endDate: '08/11/2026',
      course: 'S',
      venue: 'Aquatic Centre',
      city: 'Tauranga',
      region: 'BP',
      country: 'NZL'
    };

    const teamLookup = [{ code: 'TEST', teamName: 'Test Team', regionCode: 'BP' }];
    const swimmerRows = [{
      teamCode: 'TEST',
      firstName: 'Jane',
      lastName: 'Swimmer',
      dobFormatted: '10/02/2010',
      gender: 'Female',
      eventsStr: '50m Freestyle',
      timesStr: '00:30.00'
    }];

    const sdifStr = SDIFCreator.buildSdifString(swimmerRows, meetInfo, teamLookup);
    const b1Line = (sdifStr.split('\r\n').find(l => l.startsWith('B1')) || sdifStr.split('\n').find(l => l.startsWith('B1')));
    if (!b1Line || !b1Line.includes(`NZL${expectedCode}`)) {
      throw new Error(`SDIF meet type mismatch for ${type}: expected code ${expectedCode}`);
    }
  });

  Logger.log('✓ testSDIFMeetTypeResolution passed for all 5 competition types');
  return true;
}

function testAgeCalculationRoutine() {
  const age1 = calculateExactAge('15/06/2012', '01/10/2026');
  if (age1 !== 14) throw new Error(`Expected age 14 for DOB 15/06/2012 on 01/10/2026, got ${age1}`);

  const age2 = calculateExactAge('15/10/2012', '01/10/2026');
  if (age2 !== 13) throw new Error(`Expected age 13 for DOB 15/10/2012 on 01/10/2026, got ${age2}`);

  const ageExactBirthday = calculateExactAge('01/10/2012', '01/10/2026');
  if (ageExactBirthday !== 14) throw new Error(`Expected age 14 on exact birthday, got ${ageExactBirthday}`);

  const ageDayAfterBirthday = calculateExactAge('02/10/2012', '01/10/2026');
  if (ageDayAfterBirthday !== 13) throw new Error(`Expected age 13 for birthday day after cutoff, got ${ageDayAfterBirthday}`);

  const ageLeap1 = calculateExactAge('29/02/2012', '28/02/2026');
  if (ageLeap1 !== 13) throw new Error(`Expected age 13 for leap year birthday before Mar 1, got ${ageLeap1}`);

  const ageLeap2 = calculateExactAge('29/02/2012', '01/03/2026');
  if (ageLeap2 !== 14) throw new Error(`Expected age 14 for leap year birthday on Mar 1, got ${ageLeap2}`);

  Logger.log('✓ testAgeCalculationRoutine passed');
  return true;
}

function testMockDobGeneration() {
  const ageUpDate = '01/01/2026';

  // 1. Single-year bracket (Y6: 10 - 10)
  const daysSeenY6 = new Set();
  const monthsSeenY6 = new Set();
  for (let i = 0; i < 200; i++) {
    const mockDob = generateMockDob(10, 10, ageUpDate);
    const age = calculateExactAge(mockDob, ageUpDate);
    if (age !== 10) {
      throw new Error(`Y6 Mock DOB ${formatDob(mockDob)} produced invalid age ${age} (expected 10)`);
    }
    daysSeenY6.add(mockDob.getDate());
    monthsSeenY6.add(mockDob.getMonth());
  }
  if (daysSeenY6.size < 15 || monthsSeenY6.size < 6) {
    throw new Error('Mock DOB generation lacks expected random day/month variation');
  }

  // 2. Multi-year bracket (Intermediate: 14 - 15)
  for (let i = 0; i < 200; i++) {
    const mockDob = generateMockDob(14, 15, ageUpDate);
    const age = calculateExactAge(mockDob, ageUpDate);
    if (age < 14 || age > 15) {
      throw new Error(`Intermediate Mock DOB ${formatDob(mockDob)} produced invalid age ${age} (expected 14-15)`);
    }
  }

  // 3. Multi-year bracket (Senior: 16 - 18) with non-Jan 1 Age-up Date
  const midSeasonDate = '15/03/2026';
  for (let i = 0; i < 200; i++) {
    const mockDob = generateMockDob(16, 18, midSeasonDate);
    const age = calculateExactAge(mockDob, midSeasonDate);
    if (age < 16 || age > 18) {
      throw new Error(`Senior Mock DOB ${formatDob(mockDob)} produced invalid age ${age} on ${midSeasonDate}`);
    }
  }

  Logger.log('✓ testMockDobGeneration passed with randomized variation across 600 iterations');
  return true;
}

function testProcessEntryRowMockDobAndAge() {
  const colMapping = {
    firstName: 0,
    lastName: 1,
    dob: 2,
    gender: 3,
    schoolYear: 4,
    school: 5
  };
  const eventPairs = [];
  const context = {
    ageUpDate: '01/10/2026',
    ageBracketsMap: {
      'Y6': { schoolYear: 'Y6', minAge: 10, maxAge: 10 },
      'INTERMEDIATE': { schoolYear: 'Intermediate', minAge: 14, maxAge: 15 }
    }
  };

  // Swimmer with empty DOB and School Year "Y6"
  const row1 = ['John', 'Doe', '', 'Male', 'Y6', 'School A'];
  const processed1 = processEntryRow(row1, eventPairs, -1, null, colMapping, context);
  if (!processed1 || !processed1.dobFormatted || processed1.age !== 10 || !processed1.generatedMockDob) {
    throw new Error('processEntryRow failed to generate mock DOB and calculate age for Y6 swimmer');
  }

  // Swimmer with real DOB "15/06/2012"
  const row2 = ['Jane', 'Smith', '15/06/2012', 'Female', '', 'Team B'];
  const processed2 = processEntryRow(row2, eventPairs, -1, null, colMapping, context);
  if (!processed2 || processed2.dobFormatted !== '15/06/2012' || processed2.age !== 14 || processed2.generatedMockDob) {
    throw new Error('processEntryRow failed to calculate exact age for Age Group swimmer with DOB');
  }

  Logger.log('✓ testProcessEntryRowMockDobAndAge passed');
  return true;
}

function testSDIFAgeAndMockDobEmission() {
  const teamLookup = [{ code: 'TEST', teamName: 'Test Team', regionCode: 'BP' }];

  // 1. Age Group meet: DOB provided, no school year -> D0 pos 64-65 has integer age
  const ageGroupMeetInfo = {
    meetName: 'Age Group Champs',
    competitionType: 'Age Group',
    startDate: '01/10/2026',
    endDate: '01/10/2026',
    ageUpDate: '01/10/2026',
    course: 'S',
    venue: 'Aquatic Centre',
    city: 'Tauranga',
    region: 'BP',
    country: 'NZL'
  };

  const ageGroupSwimmers = [
    {
      teamCode: 'TEST',
      firstName: 'Alice',
      lastName: 'Brown',
      dobFormatted: '15/06/2012',
      gender: 'Female',
      eventsStr: '50m Freestyle',
      timesStr: '00:30.00',
      schoolYear: ''
    },
    {
      teamCode: 'TEST',
      firstName: 'Bob',
      lastName: 'Green',
      dobFormatted: '15/10/2012',
      gender: 'Male',
      eventsStr: '50m Freestyle',
      timesStr: '00:32.00',
      schoolYear: ''
    }
  ];

  const sdifAgeGroup = SDIFCreator.buildSdifString(ageGroupSwimmers, ageGroupMeetInfo, teamLookup);
  const d0Lines = sdifAgeGroup.split('\r\n').filter(l => l.startsWith('D0'));
  if (d0Lines.length !== 2) throw new Error('Expected 2 D0 lines in Age Group SDIF');

  const aliceAge = d0Lines[0].substring(63, 65);
  const bobAge = d0Lines[1].substring(63, 65);
  if (aliceAge !== '14') throw new Error(`Expected Alice age "14" at D0 pos 64-65, got "${aliceAge}"`);
  if (bobAge !== '13') throw new Error(`Expected Bob age "13" at D0 pos 64-65, got "${bobAge}"`);

  // 2. School meet: mock DOB generated, school year "Y6" -> D0 pos 64-65 has "Y6", pos 56-63 has MMDDYYYY
  const schoolMeetInfo = {
    meetName: 'Primary School Meet',
    competitionType: 'School',
    startDate: '01/01/2026',
    endDate: '01/01/2026',
    ageUpDate: '01/01/2026',
    course: 'S',
    venue: 'School Pool',
    city: 'Tauranga',
    region: 'BP',
    country: 'NZL'
  };

  // Simulate row processed by EntryDataProcessor
  const bracket = resolveAgeBracket('Y6');
  const mockDob = generateMockDob(bracket.minAge, bracket.maxAge, schoolMeetInfo.ageUpDate);
  const schoolSwimmers = [
    {
      teamCode: 'TEST',
      firstName: 'Charlie',
      lastName: 'Young',
      dobFormatted: formatDob(mockDob),
      gender: 'Male',
      eventsStr: '25m Freestyle',
      timesStr: '00:20.00',
      schoolYear: 'Y6'
    }
  ];

  const sdifSchool = SDIFCreator.buildSdifString(schoolSwimmers, schoolMeetInfo, teamLookup);
  const d0SchoolLine = sdifSchool.split('\r\n').find(l => l.startsWith('D0'));
  if (!d0SchoolLine) throw new Error('Missing D0 line in School SDIF');

  const schoolYearCode = d0SchoolLine.substring(63, 65);
  const schoolDobMMDDYYYY = d0SchoolLine.substring(55, 63);
  if (schoolYearCode !== 'Y6') throw new Error(`Expected school year "Y6" at D0 pos 64-65, got "${schoolYearCode}"`);
  if (!/^\d{8}$/.test(schoolDobMMDDYYYY)) throw new Error(`Expected 8-digit MMDDYYYY DOB at D0 pos 56-63, got "${schoolDobMMDDYYYY}"`);

  // 3. College meet: school year "In" (Intermediate)
  const collegeMeetInfo = {
    meetName: 'College Champs',
    competitionType: 'College',
    startDate: '01/01/2026',
    endDate: '01/01/2026',
    ageUpDate: '01/01/2026',
    course: 'S',
    venue: 'College Pool',
    city: 'Tauranga',
    region: 'BP',
    country: 'NZL'
  };

  const interBracket = resolveAgeBracket('Intermediate');
  const interMockDob = generateMockDob(interBracket.minAge, interBracket.maxAge, collegeMeetInfo.ageUpDate);
  const collegeSwimmers = [
    {
      teamCode: 'TEST',
      firstName: 'Diana',
      lastName: 'Prince',
      dobFormatted: formatDob(interMockDob),
      gender: 'Female',
      eventsStr: '100m Freestyle',
      timesStr: '01:05.00',
      schoolYear: 'In'
    }
  ];

  const sdifCollege = SDIFCreator.buildSdifString(collegeSwimmers, collegeMeetInfo, teamLookup);
  const d0CollegeLine = sdifCollege.split('\r\n').find(l => l.startsWith('D0'));
  if (!d0CollegeLine) throw new Error('Missing D0 line in College SDIF');

  const collegeYearCode = d0CollegeLine.substring(63, 65);
  if (collegeYearCode !== 'In') throw new Error(`Expected college division "In" at D0 pos 64-65, got "${collegeYearCode}"`);

  Logger.log('✓ testSDIFAgeAndMockDobEmission passed for Age Group, School, and College meets');
  return true;
}

function testSDIFExportOptionsFiltering() {
  // Mock Sheet helper
  function createMockSheet(name, isHidden, headers) {
    return {
      getName: () => name,
      isSheetHidden: () => !!isHidden,
      getLastRow: () => (headers && headers.length > 0 ? 5 : 0),
      getLastColumn: () => (headers ? headers.length : 0),
      getRange: (row, col, numRows, numCols) => ({
        getValues: () => [headers || []]
      })
    };
  }

  // Mock Spreadsheet with various sheets
  const mockSheets = [
    createMockSheet('Meet', false, ['Meet Name', 'Start Date']),
    createMockSheet('Events', false, ['Event #', 'Distance', 'Discipline']),
    createMockSheet('Schools', false, ['Team Name', 'School', 'Team Code']),
    createMockSheet('Team Officials', false, ['Role', 'Name']),
    createMockSheet('Sheet1', false, ['Some', 'Random', 'Data']),
    createMockSheet('Sheet 1', false, ['Notes']),
    createMockSheet('Individual Entries Template', false, ['#', 'First Name', 'Last Name', 'Date of Birth', 'Gender']),
    createMockSheet('Takapuna Primary', false, ['#', 'First Name', 'Last Name', 'Date of Birth', 'Gender', '50m Free']),
    createMockSheet('Custom Non Entry', false, ['Budget', 'Expense', 'Total']),
    createMockSheet('Hidden Entry Sheet', true, ['#', 'First Name', 'Last Name', 'Date of Birth', 'Gender'])
  ];

  const mockSpreadsheet = {
    getSheets: () => mockSheets
  };

  const entrySheets = SDIFCreator.getEntriesSheetNames(mockSpreadsheet);
  if (entrySheets.length !== 2) {
    throw new Error(`Expected 2 entry sheets, got ${entrySheets.length}: ${JSON.stringify(entrySheets)}`);
  }
  if (!entrySheets.includes('Individual Entries Template') || !entrySheets.includes('Takapuna Primary')) {
    throw new Error(`Expected 'Individual Entries Template' and 'Takapuna Primary', got: ${JSON.stringify(entrySheets)}`);
  }

  // Test School/Team table detection
  // 1. Only Schools exists
  const ssSchoolsOnly = {
    getSheets: () => [
      createMockSheet('Meet', false),
      createMockSheet('Schools', false)
    ]
  };
  const res1 = SDIFCreator.getSchoolTeamTableNames(ssSchoolsOnly);
  if (JSON.stringify(res1.tables) !== JSON.stringify(['Schools']) || res1.defaultTable !== 'Schools') {
    throw new Error(`Expected Schools only result, got: ${JSON.stringify(res1)}`);
  }

  // 2. Only Teams exists
  const ssTeamsOnly = {
    getSheets: () => [
      createMockSheet('Meet', false),
      createMockSheet('Teams', false)
    ]
  };
  const res2 = SDIFCreator.getSchoolTeamTableNames(ssTeamsOnly);
  if (JSON.stringify(res2.tables) !== JSON.stringify(['Teams']) || res2.defaultTable !== 'Teams') {
    throw new Error(`Expected Teams only result, got: ${JSON.stringify(res2)}`);
  }

  // 3. Both Schools and Teams exist
  const ssBoth = {
    getSheets: () => [
      createMockSheet('Meet', false),
      createMockSheet('Schools', false),
      createMockSheet('Teams', false)
    ]
  };
  const res3 = SDIFCreator.getSchoolTeamTableNames(ssBoth);
  if (res3.tables.length !== 2 || !res3.tables.includes('Schools') || !res3.tables.includes('Teams') || res3.defaultTable !== 'Schools') {
    throw new Error(`Expected Both result, got: ${JSON.stringify(res3)}`);
  }

  // 4. Neither exists
  const ssNeither = {
    getSheets: () => [
      createMockSheet('Sheet1', false)
    ]
  };
  const res4 = SDIFCreator.getSchoolTeamTableNames(ssNeither);
  if (res4.tables.length !== 2 || res4.defaultTable !== 'Schools') {
    throw new Error(`Expected fallback defaults, got: ${JSON.stringify(res4)}`);
  }

  Logger.log('✓ testSDIFExportOptionsFiltering passed');
  return true;
}

function testCSVExportFileNameAndResolution() {
  const sheet1 = {
    getName: () => 'Takapuna Primary',
    getSheetId: () => 12345
  };
  const sheet2 = {
    getName: () => 'Aquablazers',
    getSheetId: () => 67890
  };

  const mockSpreadsheet = {
    getSheetByName: name => {
      if (name === 'Takapuna Primary') return sheet1;
      if (name === 'Aquablazers') return sheet2;
      return null;
    },
    getSheetById: id => {
      if (id === 12345) return sheet1;
      if (id === 67890) return sheet2;
      return null;
    },
    getSheets: () => [sheet1, sheet2]
  };

  // Test name resolution
  const resolvedByName = resolveSheet_(mockSpreadsheet, 'Takapuna Primary');
  if (!resolvedByName || resolvedByName.getName() !== 'Takapuna Primary') {
    throw new Error('resolveSheet_ failed by exact name');
  }

  // Test ID resolution (number and string)
  const resolvedByIdNum = resolveSheet_(mockSpreadsheet, 12345);
  if (!resolvedByIdNum || resolvedByIdNum.getName() !== 'Takapuna Primary') {
    throw new Error('resolveSheet_ failed by numerical ID');
  }

  const resolvedByIdStr = resolveSheet_(mockSpreadsheet, '67890');
  if (!resolvedByIdStr || resolvedByIdStr.getName() !== 'Aquablazers') {
    throw new Error('resolveSheet_ failed by string ID');
  }

  // Test case-insensitive resolution
  const resolvedCaseInsensitive = resolveSheet_(mockSpreadsheet, 'takapuna primary');
  if (!resolvedCaseInsensitive || resolvedCaseInsensitive.getName() !== 'Takapuna Primary') {
    throw new Error('resolveSheet_ failed case-insensitively');
  }

  // Test not found
  const notFound = resolveSheet_(mockSpreadsheet, 'NonExistentSheet');
  if (notFound !== null) {
    throw new Error('resolveSheet_ should return null for non-existent sheet');
  }

  // Test empty/null
  if (resolveSheet_(mockSpreadsheet, null) !== null || resolveSheet_(mockSpreadsheet, '') !== null) {
    throw new Error('resolveSheet_ should return null for empty/null input');
  }

  // Test output file name formatting
  function formatCsvFileName(fileName, sheetName) {
    let outputFileName = (fileName || '').trim();
    if (outputFileName) {
      if (!outputFileName.toLowerCase().endsWith('.csv')) {
        outputFileName += '.csv';
      }
    } else {
      outputFileName = `${sheetName}.csv`;
    }
    return outputFileName;
  }

  if (formatCsvFileName('', 'Takapuna Primary') !== 'Takapuna Primary.csv') {
    throw new Error('formatCsvFileName failed for blank filename');
  }
  if (formatCsvFileName('Hamilton_Primary', 'Takapuna Primary') !== 'Hamilton_Primary.csv') {
    throw new Error('formatCsvFileName failed to append .csv');
  }
  if (formatCsvFileName('Hamilton_Primary.csv', 'Takapuna Primary') !== 'Hamilton_Primary.csv') {
    throw new Error('formatCsvFileName produced duplicate extension');
  }
  if (formatCsvFileName('Hamilton_Primary.CSV', 'Takapuna Primary') !== 'Hamilton_Primary.CSV') {
    throw new Error('formatCsvFileName failed case-preserving .CSV');
  }

  Logger.log('✓ testCSVExportFileNameAndResolution passed');
  return true;
}

function testCreateSheetsSourceTableFiltering() {
  function createMockSheet(name, hidden) {
    return {
      getName: () => name,
      isSheetHidden: () => !!hidden
    };
  }

  // 1. School meet setup
  const ssSchoolMeet = {
    getSheets: () => [
      createMockSheet('Meet', false),
      createMockSheet('Events', false),
      createMockSheet('Schools', false),
      createMockSheet('Team Officials', false),
      createMockSheet('Individual Entries Template', false),
      createMockSheet('Takapuna Primary', false),
      createMockSheet('Sheet1', false)
    ]
  };

  const schoolOpts = getSchoolTeamTableNames_(ssSchoolMeet);
  if (schoolOpts.tables.length !== 1 || schoolOpts.tables[0] !== 'Schools' || schoolOpts.defaultTable !== 'Schools') {
    throw new Error(`testCreateSheetsSourceTableFiltering: Expected ['Schools'] with default 'Schools', got: ${JSON.stringify(schoolOpts)}`);
  }

  // 2. Age group / club meet setup
  const ssAgeGroupMeet = {
    getSheets: () => [
      createMockSheet('Meet', false),
      createMockSheet('Events', false),
      createMockSheet('Teams', false),
      createMockSheet('Team Officials', false),
      createMockSheet('Individual Entries Template', false),
      createMockSheet('Aquablazers', false)
    ]
  };

  const ageGroupOpts = getSchoolTeamTableNames_(ssAgeGroupMeet);
  if (ageGroupOpts.tables.length !== 1 || ageGroupOpts.tables[0] !== 'Teams' || ageGroupOpts.defaultTable !== 'Teams') {
    throw new Error(`testCreateSheetsSourceTableFiltering: Expected ['Teams'] with default 'Teams', got: ${JSON.stringify(ageGroupOpts)}`);
  }

  // 3. Both Schools and Teams exist
  const ssBoth = {
    getSheets: () => [
      createMockSheet('Meet', false),
      createMockSheet('Events', false),
      createMockSheet('Schools', false),
      createMockSheet('Teams', false),
      createMockSheet('Sheet1', false)
    ]
  };

  const bothOpts = getSchoolTeamTableNames_(ssBoth);
  if (bothOpts.tables.length !== 2 || !bothOpts.tables.includes('Schools') || !bothOpts.tables.includes('Teams') || bothOpts.defaultTable !== 'Schools') {
    throw new Error(`testCreateSheetsSourceTableFiltering: Expected both with default 'Schools', got: ${JSON.stringify(bothOpts)}`);
  }

  // 4. Neither exists
  const ssNeither = {
    getSheets: () => [
      createMockSheet('Sheet1', false),
      createMockSheet('Individual Entries Template', false)
    ]
  };

  const neitherOpts = getSchoolTeamTableNames_(ssNeither);
  if (neitherOpts.tables.length !== 2 || neitherOpts.defaultTable !== 'Schools') {
    throw new Error(`testCreateSheetsSourceTableFiltering: Expected fallback defaults, got: ${JSON.stringify(neitherOpts)}`);
  }

  // 5. Case-insensitivity check (e.g., 'schools', 'TEAMS')
  const ssCasing = {
    getSheets: () => [
      createMockSheet('schools', false)
    ]
  };
  const casingOpts = getSchoolTeamTableNames_(ssCasing);
  if (casingOpts.tables.length !== 1 || casingOpts.tables[0] !== 'schools' || casingOpts.defaultTable !== 'schools') {
    throw new Error(`testCreateSheetsSourceTableFiltering: Expected 'schools' preserving case, got: ${JSON.stringify(casingOpts)}`);
  }

  // 6. Test source sheet resolution in duplicateTemplate logic
  const mockSchoolsSheet = {
    getName: () => 'Schools',
    getRange: rangeStr => ({
      getValues: () => [['Takapuna Primary'], ['Hamilton East'], ['']]
    })
  };
  const mockTemplateSheet = {
    getName: () => 'Individual Entries Template',
    copyTo: ss => ({
      setName: n => {}
    })
  };

  const mockAppSS = {
    getSheetByName: name => {
      if (name === 'Individual Entries Template') return mockTemplateSheet;
      if (name === 'Schools') return mockSchoolsSheet;
      return null;
    },
    getSheets: () => [mockTemplateSheet, mockSchoolsSheet]
  };

  // Case-insensitive resolution helper check
  function resolveSourceSheetForTemplate(ss, sheetName) {
    let sourceSheet = ss.getSheetByName(sheetName);
    if (!sourceSheet && sheetName) {
      const trimmed = String(sheetName).trim().toLowerCase();
      const sheets = ss.getSheets();
      for (let i = 0; i < sheets.length; i++) {
        if (sheets[i].getName().trim().toLowerCase() === trimmed) {
          sourceSheet = sheets[i];
          break;
        }
      }
    }
    return sourceSheet;
  }

  const resolvedExact = resolveSourceSheetForTemplate(mockAppSS, 'Schools');
  if (!resolvedExact || resolvedExact.getName() !== 'Schools') {
    throw new Error('resolveSourceSheetForTemplate failed exact match');
  }

  const resolvedCase = resolveSourceSheetForTemplate(mockAppSS, 'schools');
  if (!resolvedCase || resolvedCase.getName() !== 'Schools') {
    throw new Error('resolveSourceSheetForTemplate failed case-insensitive match');
  }

  Logger.log('✓ testCreateSheetsSourceTableFiltering passed');
  return true;
}

function testActivateSheetByName() {
  let activeSheet = null;
  let activated = false;
  let unhidden = false;

  function createMockSheet(name, hidden) {
    let isHidden = !!hidden;
    return {
      getName: () => name,
      isSheetHidden: () => isHidden,
      showSheet: () => { isHidden = false; unhidden = true; },
      activate: () => { activated = true; }
    };
  }

  const sheetSchools = createMockSheet('Schools', true);
  const sheetTeams = createMockSheet('Teams', false);
  const sheetMeet = createMockSheet('Meet', false);

  const mockSS = {
    getSheetByName: name => {
      if (name === 'Schools') return sheetSchools;
      if (name === 'Teams') return sheetTeams;
      if (name === 'Meet') return sheetMeet;
      return null;
    },
    getSheets: () => [sheetMeet, sheetSchools, sheetTeams],
    setActiveSheet: sheet => { activeSheet = sheet; }
  };

  // Helper function replicating activateSheetByName logic with injected ss
  function activateSheetHelper(ss, sheetName) {
    if (!sheetName) {
      throw new Error("Sheet name cannot be empty.");
    }
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      const trimmed = String(sheetName).trim().toLowerCase();
      const sheets = ss.getSheets();
      for (let i = 0; i < sheets.length; i++) {
        if (sheets[i].getName().trim().toLowerCase() === trimmed) {
          sheet = sheets[i];
          break;
        }
      }
    }
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found.`);
    }
    if (typeof sheet.isSheetHidden === 'function' && sheet.isSheetHidden()) {
      sheet.showSheet();
    }
    ss.setActiveSheet(sheet);
    if (typeof sheet.activate === 'function') {
      sheet.activate();
    }
    return true;
  }

  // 1. Exact match and unhide
  activeSheet = null;
  activated = false;
  unhidden = false;
  const res1 = activateSheetHelper(mockSS, 'Schools');
  if (!res1 || activeSheet !== sheetSchools || !activated || !unhidden) {
    throw new Error('testActivateSheetByName: Exact match / unhide failed');
  }

  // 2. Case-insensitive fallback
  activeSheet = null;
  activated = false;
  const res2 = activateSheetHelper(mockSS, 'teams');
  if (!res2 || activeSheet !== sheetTeams || !activated) {
    throw new Error('testActivateSheetByName: Case-insensitive match failed');
  }

  // 3. Empty name error
  let caughtEmpty = false;
  try {
    activateSheetHelper(mockSS, '');
  } catch (e) {
    caughtEmpty = true;
  }
  if (!caughtEmpty) {
    throw new Error('testActivateSheetByName: Failed to catch empty sheetName');
  }

  // 4. Missing sheet error
  let caughtMissing = false;
  try {
    activateSheetHelper(mockSS, 'NonExistentSheet');
  } catch (e) {
    caughtMissing = true;
  }
  if (!caughtMissing) {
    throw new Error('testActivateSheetByName: Failed to catch missing sheetName');
  }

  Logger.log('✓ testActivateSheetByName passed');
  return true;
}

function testDuplicateTemplateForNames() {
  const createdSheets = [];
  const sheetsList = [];

  function createMockSheet(name) {
    return {
      getName: () => name,
      copyTo: ss => {
        let copyName = name + ' (Copy)';
        const newSheet = {
          getName: () => copyName,
          setName: n => { copyName = n; }
        };
        createdSheets.push(newSheet);
        sheetsList.push(newSheet);
        return newSheet;
      }
    };
  }

  const templateSheet = createMockSheet('Individual Entries Template');
  sheetsList.push(templateSheet);

  const existingSchoolSheet = createMockSheet('Existing Primary');
  sheetsList.push(existingSchoolSheet);

  const schoolsSheet = {
    getName: () => 'Schools',
    getRange: rangeStr => {
      if (rangeStr === 'B2:B5') {
        return {
          getValues: () => [['Takapuna Primary'], ['Existing Primary'], ['Hamilton East'], ['']]
        };
      }
      if (rangeStr === 'A1:A1') {
        return {
          getValues: () => [['']]
        };
      }
      return { getValues: () => [] };
    }
  };
  sheetsList.push(schoolsSheet);

  const mockSS = {
    getSheetByName: name => {
      if (name === 'Individual Entries Template') return templateSheet;
      if (name === 'Schools') return schoolsSheet;
      if (name === 'Existing Primary') return existingSchoolSheet;
      return null;
    },
    getSheets: () => sheetsList,
    setActiveSheet: s => {},
    moveActiveSheet: pos => {}
  };

  function duplicateTemplateHelper(ss, templateSheetName, sheetName, namesRange) {
    const templateSheet = ss.getSheetByName(templateSheetName);
    if (!templateSheet) {
      throw new Error(`Template sheet "${templateSheetName}" not found!`);
    }

    let sourceSheet = ss.getSheetByName(sheetName);
    if (!sourceSheet && sheetName) {
      const trimmed = String(sheetName).trim().toLowerCase();
      const sheets = ss.getSheets();
      for (let i = 0; i < sheets.length; i++) {
        if (sheets[i].getName().trim().toLowerCase() === trimmed) {
          sourceSheet = sheets[i];
          break;
        }
      }
    }

    if (!sourceSheet) {
      throw new Error(`Sheet "${sheetName}" not found!`);
    }

    const names = sourceSheet.getRange(namesRange).getValues()
      .flat()
      .map(name => String(name).trim())
      .filter(name => name !== "");

    if (names.length === 0) {
      throw new Error(`No valid names found in range "${namesRange}" on sheet "${sheetName}".`);
    }

    const existingSheets = new Set(ss.getSheets().map(s => s.getName()));

    let createdCount = 0;
    let skippedCount = 0;

    names.forEach(name => {
      if (existingSheets.has(name)) {
        skippedCount++;
        return;
      }

      const newSheet = templateSheet.copyTo(ss);
      newSheet.setName(name);
      ss.setActiveSheet(newSheet);
      ss.moveActiveSheet(ss.getSheets().length);
      createdCount++;
    });

    return {
      success: true,
      createdCount: createdCount,
      skippedCount: skippedCount,
      message: `Created ${createdCount} sheet(s) from template "${templateSheetName}"` +
        (skippedCount > 0 ? ` (${skippedCount} skipped as already existing)` : '') + '.'
    };
  }

  // 1. Success case with 2 created and 1 skipped
  const result = duplicateTemplateHelper(mockSS, 'Individual Entries Template', 'schools', 'B2:B5');
  if (!result.success || result.createdCount !== 2 || result.skippedCount !== 1) {
    throw new Error('testDuplicateTemplateForNames: Expected 2 created, 1 skipped. Got: ' + JSON.stringify(result));
  }

  // 2. Empty names error
  let caughtEmpty = false;
  try {
    duplicateTemplateHelper(mockSS, 'Individual Entries Template', 'Schools', 'A1:A1');
  } catch (e) {
    caughtEmpty = true;
  }
  if (!caughtEmpty) {
    throw new Error('testDuplicateTemplateForNames: Failed to catch empty names range');
  }

  // 3. Missing template error
  let caughtMissingTemplate = false;
  try {
    duplicateTemplateHelper(mockSS, 'Missing Template', 'Schools', 'B2:B5');
  } catch (e) {
    caughtMissingTemplate = true;
  }
  if (!caughtMissingTemplate) {
    throw new Error('testDuplicateTemplateForNames: Failed to catch missing template');
  }

  Logger.log('✓ testDuplicateTemplateForNames passed');
  return true;
}

function runAllDiagnosticTests() {
  testCompetitionTypesRegistry();
  testAgeBracketsTableSchema();
  testTableConfigGeneration();
  testAgeCalculationRoutine();
  testMockDobGeneration();
  testProcessEntryRowMockDobAndAge();
  testSchoolAndTeamCodeLookup();
  testSDIFMeetTypeResolution();
  testSDIFAgeAndMockDobEmission();
  testSDIFExportOptionsFiltering();
  testCSVExportFileNameAndResolution();
  testCreateSheetsSourceTableFiltering();
  testActivateSheetByName();
  testDuplicateTemplateForNames();
  testFlexibleTimeParsing();
  Logger.log('=== ALL DIAGNOSTIC TESTS PASSED ===');
}

