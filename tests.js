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
  });

  if (types.SCHOOL.sdifCode !== '7' || !types.SCHOOL.hasSchoolYear || types.SCHOOL.hasAgeUpDate) {
    throw new Error('Invalid SCHOOL configuration');
  }
  if (types.COLLEGE.sdifCode !== '7' || !types.COLLEGE.hasSchoolYear || !types.COLLEGE.hasAgeUpDate) {
    throw new Error('Invalid COLLEGE configuration');
  }
  if (types.AGE_GROUP.sdifCode !== '1' || types.AGE_GROUP.hasSchoolYear || !types.AGE_GROUP.hasAgeUpDate) {
    throw new Error('Invalid AGE_GROUP configuration');
  }
  if (types.OPEN.sdifCode !== '0' || types.OPEN.hasSchoolYear || types.OPEN.hasAgeUpDate) {
    throw new Error('Invalid OPEN configuration');
  }
  if (types.MASTERS.sdifCode !== '8' || types.MASTERS.hasSchoolYear || !types.MASTERS.hasAgeUpDate) {
    throw new Error('Invalid MASTERS configuration');
  }

  Logger.log('✓ testCompetitionTypesRegistry passed');
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

function runAllDiagnosticTests() {
  testCompetitionTypesRegistry();
  testTableConfigGeneration();
  testSchoolAndTeamCodeLookup();
  testSDIFMeetTypeResolution();
  testFlexibleTimeParsing();
  Logger.log('=== ALL DIAGNOSTIC TESTS PASSED ===');
}

