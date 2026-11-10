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
 *  - range: args.rangeA1 (string formula/named range reference e.g. '=SchoolsList')
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

const COMPETITION_TYPES = {
	SCHOOL: {
		id: "school",
		name: "School",
		orgSheetName: "Schools",
		orgLabel: "School",
		sdifCode: "7",
		hasSchoolYear: true,
		hasAgeUpDate: true,
		schoolYears: ["Y5", "Y6", "Y7", "Y8"],
		defaultEvents: [
			{ distance: "25m", discipline: "Freestyle" },
			{ distance: "25m", discipline: "Backstroke" },
			{ distance: "25m", discipline: "Breaststroke" },
			{ distance: "25m", discipline: "Butterfly" },
			{ distance: "50m", discipline: "Freestyle" },
			{ distance: "50m", discipline: "Backstroke" },
			{ distance: "50m", discipline: "Breaststroke" },
			{ distance: "50m", discipline: "Butterfly" },
			{ distance: "100m", discipline: "Freestyle" },
			{ distance: "100m", discipline: "Backstroke" },
			{ distance: "100m", discipline: "Breaststroke" },
			{ distance: "100m", discipline: "Individual Medley" },
		],
	},
	COLLEGE: {
		id: "college",
		name: "College",
		orgSheetName: "Schools",
		orgLabel: "School",
		sdifCode: "7",
		hasSchoolYear: true,
		hasAgeUpDate: true,
		schoolYears: ["Y9", "Junior", "Intermediate", "Senior"],
		defaultEvents: [
			{ distance: "50m", discipline: "Freestyle" },
			{ distance: "50m", discipline: "Backstroke" },
			{ distance: "50m", discipline: "Breaststroke" },
			{ distance: "50m", discipline: "Butterfly" },
			{ distance: "100m", discipline: "Freestyle" },
			{ distance: "100m", discipline: "Backstroke" },
			{ distance: "100m", discipline: "Breaststroke" },
			{ distance: "100m", discipline: "Individual Medley" },
			{ distance: "200m", discipline: "Freestyle" },
			{ distance: "200m", discipline: "Individual Medley" },
		],
	},
	AGE_GROUP: {
		id: "age_group",
		name: "Age Group",
		orgSheetName: "Teams",
		orgLabel: "Team",
		sdifCode: "1",
		hasSchoolYear: false,
		hasAgeUpDate: true,
		defaultEvents: [
			{ distance: "50m", discipline: "Freestyle" },
			{ distance: "50m", discipline: "Backstroke" },
			{ distance: "50m", discipline: "Breaststroke" },
			{ distance: "50m", discipline: "Butterfly" },
			{ distance: "100m", discipline: "Freestyle" },
			{ distance: "100m", discipline: "Backstroke" },
			{ distance: "100m", discipline: "Breaststroke" },
			{ distance: "100m", discipline: "Butterfly" },
			{ distance: "200m", discipline: "Freestyle" },
			{ distance: "200m", discipline: "Individual Medley" },
		],
	},
	OPEN: {
		id: "open",
		name: "Open",
		orgSheetName: "Teams",
		orgLabel: "Team",
		sdifCode: "0",
		hasSchoolYear: false,
		hasAgeUpDate: true,
		defaultEvents: [
			{ distance: "50m", discipline: "Freestyle" },
			{ distance: "50m", discipline: "Backstroke" },
			{ distance: "50m", discipline: "Breaststroke" },
			{ distance: "50m", discipline: "Butterfly" },
			{ distance: "100m", discipline: "Freestyle" },
		],
	},
	MASTERS: {
		id: "masters",
		name: "Masters",
		orgSheetName: "Teams",
		orgLabel: "Club",
		sdifCode: "8",
		hasSchoolYear: false,
		hasAgeUpDate: true,
		defaultEvents: [
			{ distance: "50m", discipline: "Freestyle" },
			{ distance: "100m", discipline: "Backstroke" },
			{ distance: "50m", discipline: "Butterfly" },
			{ distance: "100m", discipline: "Freestyle" },
			{ distance: "50m", discipline: "Breaststroke" },
			{ distance: "200m", discipline: "Freestyle" },
			{ distance: "100m", discipline: "Breaststroke" },
			{ distance: "25m", discipline: "Butterfly" },
			{ distance: "100m", discipline: "Individual Medley" },
			{ distance: "50m", discipline: "Backstroke" },
			{ distance: "25m", discipline: "Freestyle" },
			{ distance: "400m", discipline: "Freestyle" },
			{ distance: "800m", discipline: "Freestyle" },
			{ distance: "1500m", discipline: "Freestyle" },
		],
	},
};

const DEFAULT_AGE_BRACKETS = {
	school: [
		{ schoolYear: "Y5", minAge: 9, maxAge: 9 },
		{ schoolYear: "Y6", minAge: 10, maxAge: 10 },
		{ schoolYear: "Y7", minAge: 11, maxAge: 11 },
		{ schoolYear: "Y8", minAge: 12, maxAge: 12 },
	],
	college: [
		{ schoolYear: "Junior", minAge: 13, maxAge: 13 },
		{ schoolYear: "Intermediate", minAge: 14, maxAge: 15 },
		{ schoolYear: "Senior", minAge: 16, maxAge: 18 },
	],
};

function getDefaultAgeBrackets() {
	return DEFAULT_AGE_BRACKETS;
}

function getCompetitionTypes() {
	return COMPETITION_TYPES;
}

function getCompetitionTypePreset(typeKey) {
	if (!typeKey) return COMPETITION_TYPES.SCHOOL;
	const key = String(typeKey)
		.trim()
		.toUpperCase()
		.replace(/[\s-]+/g, "_");
	return COMPETITION_TYPES[key] || COMPETITION_TYPES.SCHOOL;
}

function getTableConfig(overrides) {
	const DEFAULT_CLUSTERS = [
		"Team One",
		"Team Two",
		"Team Three",
		"Team Four",
		"Team Five",
		"Team Six",
		"Team Seven",
	];

	/**
	 * Pool conversion configuration
	 * Maps source pool types to their distance conversions (stroke-dependent)
	 * All conversions target 25m pool distances
	 *
	 * Structure: { poolType: { targetDistance: { stroke: sourceDistance } } }
	 * Special stroke keys:
	 * - 'default': applies to Freestyle, Backstroke, Breaststroke, Butterfly
	 * - 'IM': Individual Medley (different distances due to stroke changes)
	 *
	 * Usage: When converting times from non-standard pools to 25m pools,
	 * use RESIZESWIM(sourceDistance, targetDistance, time)
	 */
	const POOL_CONVERSIONS = {
		"From 33m pool": {
			25: {
				default: 33.3,
				IM: 33.3,
			},
			50: {
				default: 66.6,
				IM: 66.6,
			},
			100: {
				default: 99.9, // 3 lengths (Freestyle, Back, Breast, Fly)
				IM: 133.2, // 4 lengths (Individual Medley - stroke changes per length)
			},
			200: {
				default: 199.8,
				IM: 266.4,
			},
			400: {
				default: 399.6,
				IM: 532.8,
			},
		},
		"From 18m pool": {
			25: { default: 18, IM: 18 },
			50: { default: 36, IM: 36 },
			100: { default: 72, IM: 72 },
			200: { default: 144, IM: 144 },
			400: { default: 288, IM: 288 },
		},
		"From LC": {
			25: { default: 50, IM: 50 },
			50: { default: 50, IM: 50 },
			100: { default: 100, IM: 100 },
			200: { default: 200, IM: 200 },
			400: { default: 400, IM: 400 },
		},
	};

	const relayDefaults = {
		suggestedEventNames: [
			"4x25m Girls Freestyle Relay",
			"4x25m Boys Freestyle Relay",
			"4x25m Freestyle Grand Relay",
		],
		defaultRowsPerTable: 4, // fixed swimmer rows per relay table
		gapBetweenTables: 3, // blank rows between tables when on same sheet
	};

	const compPreset = getCompetitionTypePreset(
		overrides && overrides.competitionType,
	);
	const YEARS = (overrides && overrides.schoolYears) ||
		compPreset.schoolYears || ["Y5", "Y6", "Y7", "Y8"];
	const GENDERS = ["Female", "Male"];
	const COUNTRIES = (overrides && overrides.country) || [
		"NZL",
		"AUS",
		"CAN",
		"GBR",
		"USA",
	];
	const REGIONS = (overrides && overrides.region) || [
		"AK",
		"BP",
		"CB",
		"CO",
		"ED",
		"HP",
		"MW",
		"NL",
		"NM",
		"OT",
		"SL",
		"TR",
		"WG",
		"WK",
		"WN",
		"WP",
	];

	/** Utility to build a map from headers for concise config */
	const toCols = (spec) => spec; // pass-through for readability

	const hasSchoolYear =
		overrides && overrides.hasSchoolYear !== undefined
			? overrides.hasSchoolYear
			: compPreset.hasSchoolYear;
	const orgColName =
		compPreset.orgSheetName === "Schools" ? "School" : "Team";
	const orgNamedRange =
		compPreset.orgSheetName === "Schools" ? "=SchoolsList" : "=TeamsList";

	// Build Individual Events Template headers and columns dynamically
	let indTemplateHeaders = [
		"#",
		"First Name",
		"Last Name",
		"Date of Birth",
		"Gender",
	];
	const indTemplateCols = {
		"#": { type: "number" },
		"First Name": { type: "text" },
		"Last Name": { type: "text" },
		"Date of Birth": {
			type: "date",
			validation: { type: "date", args: { condition: "DATE_IS_VALID" } },
		},
		Gender: {
			type: "text",
			validation: { type: "list", args: { values: GENDERS } },
		},
	};

	if (hasSchoolYear) {
		indTemplateHeaders.push("School Year");
		indTemplateCols["School Year"] = {
			type: "text",
			validation: { type: "list", args: { values: YEARS } },
		};
	}
	indTemplateHeaders.push(orgColName);
	indTemplateCols[orgColName] = {
		type: "text",
		validation: { type: "range", args: { rangeA1: orgNamedRange } },
	};

	for (let i = 1; i <= 9; i++) {
		indTemplateHeaders.push("Event " + i);
		indTemplateHeaders.push("Time " + i + " (m:s.S)");
		indTemplateCols["Event " + i] = {
			type: "text",
			validation: { type: "range", args: { rangeA1: "=EventsList" } },
		};
		indTemplateCols["Time " + i + " (m:s.S)"] = { type: "text" };
	}
	indTemplateHeaders.push("Convert times");
	indTemplateCols["Convert times"] = {
		type: "text",
		validation: {
			type: "list",
			args: { values: ["From 33m pool", "From 18m pool", "No"] },
		},
		default: "",
	};

	const config = {
		Meet: {
			tableType: "core",
			title: "Meet Details*",
			headers: [
				"Meet Name",
				"Competition Type",
				"Start Date",
				"End Date",
				"Age-up Date",
				"Course",
				"Contact Name",
				"Contact Phone",
				"Venue",
				"Venue Address1",
				"City",
				"Region",
				"Postcode",
				"Country",
			],
			columns: toCols({
				"Meet Name": { type: "text" },
				"Competition Type": {
					type: "text",
					validation: {
						type: "list",
						args: {
							values: [
								"School",
								"College",
								"Age Group",
								"Open",
								"Masters",
							],
						},
					},
					default: "School",
				},
				"Start Date": {
					type: "date",
					validation: {
						type: "date",
						args: { condition: "DATE_IS_VALID" },
					},
				},
				"End Date": {
					type: "date",
					validation: {
						type: "date",
						args: { condition: "DATE_IS_VALID" },
					},
				},
				"Age-up Date": {
					type: "date",
					validation: {
						type: "date",
						args: { condition: "DATE_IS_VALID" },
					},
				},
				Course: {
					type: "text",
					validation: { type: "list", args: { values: ["S", "L"] } },
					default: "S",
				},
				"Contact Name": { type: "text" },
				"Contact Phone": { type: "text" },
				Venue: { type: "text" },
				"Venue Address1": { type: "text" },
				City: { type: "text" },
				Region: {
					type: "text",
					validation: { type: "list", args: { values: REGIONS } },
					default: "BP",
				},
				Postcode: { type: "text" },
				Country: {
					type: "text",
					validation: { type: "list", args: { values: COUNTRIES } },
					default: "NZL",
				},
			}),
			options: {
				freezeHeader: 1,
				headerBg: "#356853",
				title: "Meet Details",
				rows: 1,
				required: true,
				placement: { targetSheet: "Meet", startCell: "A1" },
				clearMode: "rebuild",
			},
		},
		AgeBrackets: {
			tableType: "core",
			title: "School Year Age Brackets",
			headers: ["School Year", "Min Age", "Max Age"],
			columns: toCols({
				"School Year": {
					type: "text",
					validation: { type: "list", args: { values: YEARS } },
				},
				"Min Age": {
					type: "number",
					validation: {
						type: "numberRange",
						args: { min: 5, max: 25 },
					},
				},
				"Max Age": {
					type: "number",
					validation: {
						type: "numberRange",
						args: { min: 5, max: 25 },
					},
				},
			}),
			options: {
				freezeHeader: 0,
				headerBg: "#356853",
				title: "School Year Age Brackets",
				rows: 4,
				required: false,
				placement: { targetSheet: "Meet", startCell: "A5" },
				clearMode: "rebuild",
			},
			notes: "Age brackets mapping school years to eligible min and max ages for mock DOB calculation.",
		},
		TeamOfficials: {
			tableType: "core",
			title: "Team Officials*",
			headers: [
				"Cluster or School",
				"Sport Coordinator",
				"Email",
				"Contact Phone",
				"School/Organisation",
				"Team Manager Name",
				"Team Manager Contact",
				"Team Manager Email",
				"Qualified Officials Name",
				"Qualified Officials Email",
				"Timekeeper 1 Name",
				"Timekeeper 2 Name",
				"Timekeeper 3 Name",
				"Other Helpers",
			],
			columns: toCols({
				"Cluster or School": { type: "text" },
				"Sport Coordinator": { type: "text" },
				Email: { type: "text" },
				"Contact Phone": { type: "text" },
				"School/Organisation": { type: "text" },
				"Team Manager Name": { type: "text" },
				"Team Manager Contact": { type: "text" },
				"Team Manager Email": { type: "text" },
				"Qualified Officials Name": { type: "text" },
				"Qualified Officials Email": { type: "text" },
				"Timekeeper 1 Name": { type: "text" },
				"Timekeeper 2 Name": { type: "text" },
				"Timekeeper 3 Name": { type: "text" },
				"Other Helpers": { type: "text" },
			}),
			options: {
				freezeHeader: 1,
				rows: 15,
				headerBg: "#356853",
				title: "Team Officials",
				required: true,
				placement: { targetSheet: "Team Officials", startCell: "A1" },
				clearMode: "rebuild",
			},
			notes: "All events with detailed information.",
		},

		EventsForTemplate: {
			tableType: "core",
			title: "Events*",
			headers: ["Distance", "Discipline", "Events"],
			columns: toCols({
				Distance: {
					type: "text",
					validation: {
						type: "list",
						args: {
							values: [
								"25m",
								"50m",
								"100m",
								"200m",
								"400m",
								"800m",
								"1500m",
							],
						},
					},
				},
				Discipline: {
					type: "text",
					validation: {
						type: "list",
						args: {
							values: [
								"Freestyle",
								"Backstroke",
								"Breaststroke",
								"Butterfly",
								"Individual Medley",
							],
						},
					},
				},
				Events: {
					type: "text",
					formula: '=IF(AND(B2<>"",A2<>""),A2&" "&B2,"")',
				},
			}),
			options: {
				freezeHeader: 1,
				headerBg: "#356853",
				rows: 15,
				required: true,
				placement: { targetSheet: "Events", startCell: "A1" },
				clearMode: "rebuild",
				namedRange: { name: "EventsList", columnName: "Events" },
			},
			notes: "Simplified event names used for dropdown validation in individual entries.",
		},

		Schools: {
			tableType: "core",
			title: "Schools*",
			headers: ["Team Name", "School", "Cluster", "Code", "Region Code"],
			columns: toCols({
				"Team Name": { type: "text", default: "Team Name" },
				School: { type: "text", default: "School Name" },
				Cluster: {
					type: "text",
					validation: {
						type: "list",
						args: { values: DEFAULT_CLUSTERS },
					},
				},
				Code: { type: "text", default: "UNSP" },
				"Region Code": {
					type: "text",
					validation: { type: "list", args: { values: REGIONS } },
					default: "BP",
				},
			}),
			options: {
				freezeHeader: 1,
				headerBg: "#356853",
				rows: 30,
				required: true,
				placement: { targetSheet: "Schools", startCell: "A1" },
				clearMode: "rebuild",
				namedRange: { name: "SchoolsList", columnName: "School" },
			},
		},

		Teams: {
			tableType: "core",
			title: "Teams*",
			headers: ["Team Name", "Team", "Code", "Region Code"],
			columns: toCols({
				"Team Name": { type: "text", default: "Team Name" },
				Team: { type: "text", default: "Team / Club Name" },
				Code: { type: "text", default: "UNSP" },
				"Region Code": {
					type: "text",
					validation: { type: "list", args: { values: REGIONS } },
					default: "BP",
				},
			}),
			options: {
				freezeHeader: 1,
				headerBg: "#356853",
				rows: 30,
				required: true,
				placement: { targetSheet: "Teams", startCell: "A1" },
				clearMode: "rebuild",
				namedRange: { name: "TeamsList", columnName: "Team" },
			},
		},

		IndividualEventsTemplate: {
			tableType: "core",
			title: "Individual Events Template*",
			headers: indTemplateHeaders,
			columns: toCols(indTemplateCols),
			options: {
				conversions: POOL_CONVERSIONS, // Reference to pool conversion configuration
				freezeHeader: 1,
				headerBg: "#356853",
				title: "Individual Entries Template",
				rows: 30,
				required: true,
				placement: {
					targetSheet: "Individual Events Template",
					startCell: "A1",
				},
				clearMode: "rebuild",
			},
			notes: "Template for duplication into per-cluster sheets.",
		},
		SampleIndividualEntries: {
			tableType: "core",
			title: "Sample Individual Entries",
			headers: [
				"#",
				"First Name",
				"Last Name",
				"Date of Birth",
				"Gender",
				"School Year",
				"School",
				"Event 1",
				"Time 1 (m:s.S)",
				"Event 2",
				"Time 2 (m:s.S)",
				"Event 3",
				"Time 3 (m:s.S)",
				"Event 4",
				"Time 4 (m:s.S)",
				"Event 5",
				"Time 5 (m:s.S)",
				"Event 6",
				"Time 6 (m:s.S)",
				"Event 7",
				"Time 7 (m:s.S)",
				"Event 8",
				"Time 8 (m:s.S)",
				"Event 9",
				"Time 9 (m:s.S)",
				"Convert times",
			],
			columns: toCols({
				"#": { type: "number", default: "1" },
				"First Name": { type: "text", default: "Leah" },
				"Last Name": { type: "text", default: "Smith" },
				"Date of Birth": {
					type: "date",
					validation: {
						type: "date",
						args: { condition: "DATE_IS_VALID" },
					},
					default: "08/11/2012",
				},
				Gender: {
					type: "text",
					validation: { type: "list", args: { values: GENDERS } },
					default: "Female",
				},
				"School Year": {
					type: "text",
					validation: { type: "list", args: { values: YEARS } },
					default: "Y5",
				},
				School: {
					type: "text",
					validation: {
						type: "list",
						args: {
							values: [
								"ACG",
								"Bethlehem College",
								"Arataki",
								"Greenpark",
								"Rangitoto Primary",
								"Takapuna",
								"Omanu",
							],
						},
					},
					default: "ACG",
				},
				"Event 1": {
					type: "text",
					validation: {
						type: "list",
						args: {
							values: [
								"25m Breaststroke",
								"25m Freestyle",
								"50m Backstroke",
								"100m Freestyle",
								"100m Individual Medley",
								"50m Breaststroke",
							],
						},
					},
					default: "25m Breaststroke",
				},
				"Time 1 (m:s.S)": { type: "text", default: "00:45.98" },
				"Event 2": {
					type: "text",
					validation: {
						type: "list",
						args: {
							values: [
								"25m Breaststroke",
								"25m Freestyle",
								"50m Backstroke",
								"100m Freestyle",
								"100m Individual Medley",
								"50m Breaststroke",
							],
						},
					},
					default: "50m Breaststroke",
				},
				"Time 2 (m:s.S)": { type: "text", default: "01:01.20" },
				"Event 3": {
					type: "text",
					validation: {
						type: "list",
						args: {
							values: [
								"25m Breaststroke",
								"25m Freestyle",
								"50m Backstroke",
								"100m Freestyle",
								"100m Individual Medley",
								"50m Breaststroke",
							],
						},
					},
					default: "50m Backstroke",
				},
				"Time 3 (m:s.S)": { type: "text", default: "00:59.21" },
				"Event 4": {
					type: "text",
					validation: {
						type: "list",
						args: {
							values: [
								"25m Breaststroke",
								"25m Freestyle",
								"50m Backstroke",
								"100m Freestyle",
								"100m Individual Medley",
								"50m Breaststroke",
							],
						},
					},
					default: "100m Individual Medley",
				},
				"Time 4 (m:s.S)": { type: "text", default: "NT" },
				"Event 5": {
					type: "text",
					validation: {
						type: "list",
						args: {
							values: [
								"25m Breaststroke",
								"25m Freestyle",
								"50m Backstroke",
								"100m Freestyle",
								"100m Individual Medley",
								"50m Breaststroke",
							],
						},
					},
					default: "100m Freestyle",
				},
				"Time 5 (m:s.S)": { type: "text", default: "1:42.67" },
				"Event 6": {
					type: "text",
					validation: {
						type: "list",
						args: {
							values: [
								"25m Breaststroke",
								"25m Freestyle",
								"50m Backstroke",
								"100m Freestyle",
								"100m Individual Medley",
								"50m Breaststroke",
							],
						},
					},
					default: "25m Freestyle",
				},
				"Time 6 (m:s.S)": { type: "text", default: "00:34.65" },
				"Event 7": {
					type: "text",
					validation: {
						type: "list",
						args: {
							values: [
								"25m Breaststroke",
								"25m Freestyle",
								"50m Backstroke",
								"100m Freestyle",
								"100m Individual Medley",
								"50m Breaststroke",
							],
						},
					},
				},
				"Time 7 (m:s.S)": { type: "text" },
				"Event 8": {
					type: "text",
					validation: {
						type: "list",
						args: {
							values: [
								"25m Breaststroke",
								"25m Freestyle",
								"50m Backstroke",
								"100m Freestyle",
								"100m Individual Medley",
								"50m Breaststroke",
							],
						},
					},
				},
				"Time 8 (m:s.S)": { type: "text" },
				"Event 9": {
					type: "text",
					validation: {
						type: "list",
						args: {
							values: [
								"25m Breaststroke",
								"25m Freestyle",
								"50m Backstroke",
								"100m Freestyle",
								"100m Individual Medley",
								"50m Breaststroke",
							],
						},
					},
				},
				"Time 9 (m:s.S)": { type: "text" },
				"Convert times": {
					type: "text",
					validation: {
						type: "list",
						args: {
							values: ["From 33m pool", "From 18m pool", "No"],
						},
					},
					default: "",
				},
			}),
			options: {
				conversions: POOL_CONVERSIONS, // Reference to pool conversion configuration
				freezeHeader: 1,
				headerBg: "#139997",
				title: "Sample Individual Entries",
				rows: 10,
				required: true,
				placement: {
					targetSheet: "How to fill entries",
					startCell: "A1",
				},
				clearMode: "rebuild",
			},
			notes: "Sample table to show how to fill swimmer data.",
		},
		DetailedEvents: {
			tableType: "optional",
			title: "Detailed Events",
			headers: [
				"Events",
				"Gender",
				"Min Age",
				"Max Age",
				"Discipline",
				"Distance",
				"Event Type",
				"Event No.",
			],
			columns: toCols({
				Events: { type: "text" },
				Gender: {
					type: "text",
					validation: { type: "list", args: { values: GENDERS } },
				},
				"Min Age": {
					type: "number",
					validation: {
						type: "numberRange",
						args: { min: 5, max: 18 },
					},
				},
				"Max Age": {
					type: "number",
					validation: {
						type: "numberRange",
						args: { min: 5, max: 18 },
					},
				},
				Discipline: {
					type: "text",
					validation: {
						type: "list",
						args: {
							values: [
								"Freestyle",
								"Backstroke",
								"Breaststroke",
								"Butterfly",
								"Individual Medley",
								"Relay Medley",
							],
						},
					},
				},
				Distance: {
					type: "number",
					validation: {
						type: "numberRange",
						args: { min: 25, max: 400 },
					},
				},
				"Event Type": {
					type: "text",
					validation: {
						type: "list",
						args: { values: ["Individual", "Relay"] },
					},
				},
				"Event No.": { type: "number" },
			}),
			options: {
				freezeHeader: 1,
				headerBg: "#356853",
				title: "Events for Meet Manager",
				required: false,
				placement: { targetSheet: "MMEvents", startCell: "A1" },
				clearMode: "rebuild",
			},
			notes: "Full events list for Meet Manager alignment.",
		},

		// Relay sheet consists of multiple small tables on one sheet.
		// The builder can place individual sections; here we define a single 4x swimmers block header set.
		// NOTE: Team name columns are NOT included in headers — they are dynamically appended at runtime
		// from actual team names provided via the relay table creation dialog.
		RelayEntry: {
			tableType: "relay",
			title: "Relays",
			headers: ["Order", "School Year", "Gender"], // Team columns appended dynamically at build time
			columns: toCols({
				Order: {
					type: "text",
					validation: {
						type: "list",
						args: {
							values: [
								"1st Swimmer",
								"2nd Swimmer",
								"3rd Swimmer",
								"4th Swimmer",
							],
						},
					},
				},
				"School Year": {
					type: "text",
					validation: { type: "list", args: { values: YEARS } },
				},
				Gender: {
					type: "text",
					validation: { type: "list", args: { values: GENDERS } },
				},
				// Team name columns are free-text for swimmer names, added dynamically
			}),
			options: {
				freezeHeader: 0,
				headerBg: "#356853",
				rows: 4,
				required: false,
				placement: { targetSheet: "Relays", startCell: "A1" },
				clearMode: "rebuild",
			},
		},
	};

	return {
		tables: config,
		relayDefaults: relayDefaults,
		defaultClusters: DEFAULT_CLUSTERS, // Fallback if no team names available
		poolConversions: POOL_CONVERSIONS, // Pool conversion configuration
	};
}

/**
 * Utility to list available table names for the UI.
 */
function listAvailableTables(overrides) {
	const cfg = getTableConfig(overrides);
	const tables = Object.keys(cfg.tables)
		.filter(function (name) {
			// Only include core tables, exclude relay tables
			return cfg.tables[name].tableType === "core";
		})
		.map(function (name) {
			return {
				name: name,
				title: cfg.tables[name].title || name,
			};
		});
	Logger.log(
		"[listAvailableTables] Found %s tables: %s",
		tables.length,
		JSON.stringify(tables),
	);
	return tables;
}

/**
 * Gets pool conversion configuration
 * @return {Object} Pool conversion mappings
 */
function getPoolConversions() {
	const cfg = getTableConfig();
	return cfg.poolConversions;
}

/**
 * Parses stroke type from event string
 * @param {string} eventStr - Event string (e.g., "100m Individual Medley", "50m Freestyle")
 * @return {string} 'IM' for Individual Medley, 'default' for all other strokes
 */
function getStrokeType(eventStr) {
	if (!eventStr) return "default";
	const s = String(eventStr).toLowerCase();

	// Check for Individual Medley variations
	if (
		s.includes("individual medley") ||
		s.includes("medley") ||
		/\bim\b/.test(s)
	) {
		return "IM";
	}

	// All other strokes (Freestyle, Backstroke, Breaststroke, Butterfly) use 'default'
	return "default";
}

/**
 * Gets the source distance for a given target distance, pool type, and stroke
 * @param {string} poolType - Pool type key (e.g., "From 33m pool")
 * @param {number} targetDistance - Target distance in 25m pool (25, 50, 100, 200, 400)
 * @param {string} [strokeType='default'] - Stroke type ('default' or 'IM')
 * @return {number|null} Source distance or null if not found
 */
function getSourceDistance(poolType, targetDistance, strokeType) {
	const conversions = getPoolConversions();
	if (!conversions[poolType]) return null;

	const poolMap = conversions[poolType];
	if (!poolMap[targetDistance]) return null;

	const stroke = strokeType || "default";
	const distanceMap = poolMap[targetDistance];

	// Return stroke-specific distance, fallback to 'default' if stroke not found
	return distanceMap[stroke] || distanceMap["default"] || null;
}
