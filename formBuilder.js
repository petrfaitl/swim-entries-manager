/**
 * Programmatic Google Form Generator & Sync Service for Swim Entries Manager.
 * Generates a competition-type-aware Google Form linked to the active spreadsheet
 * so parents or individual swimmers can submit entries directly into the export pipeline.
 */

const FORM_PROPERTIES_KEYS = {
	FORM_ID: "swimEntriesFormId",
	PUBLISHED_URL: "swimEntriesFormPublishedUrl",
	EDIT_URL: "swimEntriesFormEditUrl",
	DEST_SHEET: "swimEntriesFormDestSheet",
};

const FORM_PLACEHOLDER_VALUES = new Set([
	"school name",
	"team name",
	"team / club name",
	"team/club name",
	"unsp",
	"uns",
]);

const DEFAULT_FORM_PRIVACY_NOTICE =
	"Privacy Notice: Information submitted in this form (swimmer name, date of birth, gender, school/team, and event entries) " +
	"is collected solely for organising and administering this swim meet.";

/**
 * Normalises a cell value to a trimmed string.
 * @param {*} val
 * @return {string}
 */
function cleanFormOptionText_(val) {
	if (val === null || val === undefined) return "";
	return String(val).trim();
}

/**
 * Checks whether a value is a template placeholder or blank.
 * @param {string} val
 * @return {boolean}
 */
function isFormPlaceholder_(val) {
	const s = cleanFormOptionText_(val);
	if (!s) return true;
	return FORM_PLACEHOLDER_VALUES.has(s.toLowerCase());
}

/**
 * Reads Meet metadata from the Meet sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @return {Object}
 */
function readFormMeetInfo_(ss) {
	if (typeof readMeetInfo_ === "function") {
		const info = readMeetInfo_("Meet");
		if (info && Object.keys(info).length > 0) return info;
	}

	const meetSheet = ss.getSheetByName("Meet");
	if (!meetSheet) return {};

	const data = meetSheet.getDataRange().getValues();
	if (!data || data.length < 2) return {};

	const headers = data[0];
	const row = data[1];
	const info = {};

	headers.forEach(function (h, idx) {
		const key = String(h || "")
			.trim()
			.toLowerCase()
			.replace(/[^a-z0-9]/g, "");
		if (!key) return;
		if (key === "meetname") info.meetName = cleanFormOptionText_(row[idx]);
		else if (key === "competitiontype")
			info.competitionType = cleanFormOptionText_(row[idx]);
		else if (key === "startdate") info.startDate = row[idx];
		else if (key === "enddate") info.endDate = row[idx];
		else if (key === "ageupdate") info.ageUpDate = row[idx];
		else if (key === "course") info.course = cleanFormOptionText_(row[idx]);
		else if (key === "venue") info.venue = cleanFormOptionText_(row[idx]);
		else if (key === "city") info.city = cleanFormOptionText_(row[idx]);
	});

	return info;
}

/**
 * Extracts clean, deduplicated event names from the Events sheet.
 * Supports both calculated "Events" column and Distance + Discipline fallback.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @return {string[]}
 */
function extractFormEvents_(ss) {
	const eventsSheet = ss.getSheetByName("Events");
	if (!eventsSheet) {
		throw new Error(
			'Required sheet "Events" not found. Please run Meet Setup first.',
		);
	}

	const data = eventsSheet.getDataRange().getValues();
	if (!data || data.length < 2) {
		throw new Error(
			'The "Events" sheet has no event rows. Please add at least one event before generating an Entry Form.',
		);
	}

	const headers = data[0].map((h) => cleanFormOptionText_(h).toLowerCase());
	const eventsColIdx = headers.findIndex(
		(h) => h === "events" || h === "event" || h === "event name",
	);
	const distColIdx = headers.findIndex((h) => h === "distance");
	const discColIdx = headers.findIndex(
		(h) => h === "discipline" || h === "stroke",
	);

	const seen = new Set();
	const events = [];

	for (let r = 1; r < data.length; r++) {
		const row = data[r];
		let eventName =
			eventsColIdx !== -1 ? cleanFormOptionText_(row[eventsColIdx]) : "";

		if (!eventName && distColIdx !== -1 && discColIdx !== -1) {
			const dist = cleanFormOptionText_(row[distColIdx]);
			const disc = cleanFormOptionText_(row[discColIdx]);
			if (dist && disc) {
				eventName = `${dist} ${disc}`;
			}
		}

		if (!isFormPlaceholder_(eventName)) {
			const key = eventName.toLowerCase();
			if (!seen.has(key)) {
				seen.add(key);
				events.push(eventName);
			}
		}
	}

	if (events.length === 0) {
		throw new Error(
			'No valid events found in the "Events" sheet. Please add at least one event (Distance and Discipline) before generating an Entry Form.',
		);
	}

	return events;
}

/**
 * Extracts clean, deduplicated organisation names from Schools or Teams sheet.
 * Filters out placeholder defaults such as "School Name", "Team Name", and "Team / Club Name".
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Object} compPreset
 * @return {string[]}
 */
function extractFormOrganizations_(ss, compPreset) {
	const orgSheetName = compPreset.orgSheetName || "Schools";
	const orgLabel = compPreset.orgLabel || "School";
	const orgSheet = ss.getSheetByName(orgSheetName);

	if (!orgSheet) {
		throw new Error(
			`Required sheet "${orgSheetName}" not found. Please run Meet Setup first.`,
		);
	}

	const data = orgSheet.getDataRange().getValues();
	if (!data || data.length < 2) {
		throw new Error(
			`The "${orgSheetName}" sheet is empty. Please add at least one ${orgLabel} before generating an Entry Form.`,
		);
	}

	const headers = data[0].map((h) => cleanFormOptionText_(h).toLowerCase());
	let nameColIdx = headers.indexOf(orgLabel.toLowerCase());
	if (nameColIdx === -1) nameColIdx = headers.indexOf("school");
	if (nameColIdx === -1) nameColIdx = headers.indexOf("team");
	if (nameColIdx === -1) nameColIdx = headers.indexOf("team name");
	if (nameColIdx === -1) nameColIdx = Math.min(1, headers.length - 1);

	const seen = new Set();
	const organizations = [];

	for (let r = 1; r < data.length; r++) {
		const val = cleanFormOptionText_(data[r][nameColIdx]);
		if (!isFormPlaceholder_(val)) {
			const key = val.toLowerCase();
			if (!seen.has(key)) {
				seen.add(key);
				organizations.push(val);
			}
		}
	}

	if (organizations.length === 0) {
		throw new Error(
			`No valid entries found in "${orgSheetName}". Please replace placeholder rows with at least one real ${orgLabel} before generating an Entry Form.`,
		);
	}

	return organizations;
}

/**
 * Extracts School Year options from AgeBrackets on the Meet sheet, falling back to preset defaults.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Object} compPreset
 * @return {string[]}
 */
function extractFormSchoolYears_(ss, compPreset) {
	const defaults =
		compPreset && compPreset.schoolYears
			? compPreset.schoolYears.slice()
			: ["Y5", "Y6", "Y7", "Y8"];
	const meetSheet = ss.getSheetByName("Meet");
	if (!meetSheet) return defaults;

	const data = meetSheet.getDataRange().getValues();
	let headerRowIdx = -1;
	for (let r = 0; r < data.length; r++) {
		const rowStr = data[r]
			.map((c) => cleanFormOptionText_(c).toLowerCase())
			.join(" ");
		if (rowStr.includes("school year") && rowStr.includes("min age")) {
			headerRowIdx = r;
			break;
		}
	}
	if (headerRowIdx === -1) return defaults;

	const syColIdx = data[headerRowIdx].findIndex((h) =>
		/school\s*year/i.test(String(h || "")),
	);
	if (syColIdx === -1) return defaults;

	const years = [];
	const seen = new Set();
	for (let r = headerRowIdx + 1; r < data.length; r++) {
		const sy = cleanFormOptionText_(data[r][syColIdx]);
		if (sy && !seen.has(sy.toLowerCase())) {
			seen.add(sy.toLowerCase());
			years.push(sy);
		}
	}

	return years.length > 0 ? years : defaults;
}

/**
 * Validates that Meet, Events, and Schools/Teams exist and contain non-placeholder entries.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 * @return {{ meetInfo: Object, compPreset: Object, events: string[], organizations: string[], schoolYears: string[] }}
 */
function validateFormPrerequisites_(ss) {
	ss = ss || SpreadsheetApp.getActiveSpreadsheet();
	if (!ss) {
		throw new Error("No active spreadsheet found.");
	}

	const meetSheet = ss.getSheetByName("Meet");
	if (!meetSheet) {
		throw new Error(
			'Required sheet "Meet" not found. Please run Meet Setup first.',
		);
	}

	const meetInfo = readFormMeetInfo_(ss);
	if (!meetInfo || (!meetInfo.meetName && !meetInfo.competitionType)) {
		throw new Error(
			'The "Meet" sheet appears to be empty. Please complete Meet Setup first.',
		);
	}

	const compPreset = getCompetitionTypePreset(meetInfo.competitionType);
	const events = extractFormEvents_(ss);
	const organizations = extractFormOrganizations_(ss, compPreset);
	const schoolYears = compPreset.hasSchoolYear
		? extractFormSchoolYears_(ss, compPreset)
		: [];

	return {
		meetInfo: meetInfo,
		compPreset: compPreset,
		events: events,
		organizations: organizations,
		schoolYears: schoolYears,
	};
}

/**
 * Returns the current stored Form status and prerequisite readiness for the Form Setup dialog.
 *
 * @return {Object}
 */
function getStoredFormStatus() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const props = PropertiesService.getDocumentProperties();
	const formId = props.getProperty(FORM_PROPERTIES_KEYS.FORM_ID) || "";
	const publishedUrl =
		props.getProperty(FORM_PROPERTIES_KEYS.PUBLISHED_URL) || "";
	const editUrl = props.getProperty(FORM_PROPERTIES_KEYS.EDIT_URL) || "";
	const destinationSheet =
		props.getProperty(FORM_PROPERTIES_KEYS.DEST_SHEET) || "";

	let hasExistingForm = Boolean(formId);
	let formMissing = false;

	let collectEmailAndAllowEdits = true;
	let eventSlots = 5;
	let requireDobForSchool = false;
	let allowOtherTeam = false;

	if (formId && typeof FormApp !== "undefined") {
		try {
			const existingForm = FormApp.openById(formId);
			if (typeof existingForm.collectsEmail === "function") {
				collectEmailAndAllowEdits = Boolean(existingForm.collectsEmail());
			}
			const existingItems =
				typeof existingForm.getItems === "function"
					? existingForm.getItems()
					: [];
			let maxEventSlot = 0;
			existingItems.forEach(function (item) {
				const t = cleanFormOptionText_(item.getTitle());
				const m = /^Event\s+(\d+)$/i.exec(t);
				if (m) {
					const n = parseInt(m[1], 10);
					if (n > maxEventSlot) maxEventSlot = n;
				}
				if (t.toLowerCase() === "date of birth") {
					requireDobForSchool = true;
				}
				if (
					String(item.getType()) === "MULTIPLE_CHOICE" &&
					(t.toLowerCase() === "club" || t.toLowerCase() === "team")
				) {
					try {
						const mc = item.asMultipleChoiceItem();
						allowOtherTeam =
							typeof mc.hasOtherOption === "function"
								? Boolean(mc.hasOtherOption())
								: true;
					} catch (ignoreErr) {
						allowOtherTeam = true;
					}
				}
			});
			if (maxEventSlot >= 1 && maxEventSlot <= 9) {
				eventSlots = maxEventSlot;
			}
		} catch (e) {
			Logger.log(
				'[formBuilder] Stored formId "%s" could not be opened: %s',
				formId,
				e.message,
			);
			hasExistingForm = false;
			formMissing = true;
		}
	}

	let prereqReady = false;
	let prereqError = "";
	let meetName = "";
	let competitionType = "";
	let hasSchoolYear = false;
	let orgLabel = "School";
	let eventCount = 0;
	let organizationCount = 0;

	try {
		const prereq = validateFormPrerequisites_(ss);
		prereqReady = true;
		meetName = prereq.meetInfo.meetName || "";
		competitionType = prereq.compPreset.name || "";
		hasSchoolYear = Boolean(prereq.compPreset.hasSchoolYear);
		orgLabel = prereq.compPreset.orgLabel || "School";
		eventCount = prereq.events.length;
		organizationCount = prereq.organizations.length;
	} catch (err) {
		prereqError = err.message;
	}

	return {
		hasExistingForm: hasExistingForm,
		formMissing: formMissing,
		formId: hasExistingForm ? formId : "",
		publishedUrl: hasExistingForm ? publishedUrl : "",
		editUrl: hasExistingForm ? editUrl : "",
		destinationSheet: hasExistingForm ? destinationSheet : "",
		prereqReady: prereqReady,
		prereqError: prereqError,
		meetName: meetName,
		competitionType: competitionType,
		hasSchoolYear: hasSchoolYear,
		orgLabel: orgLabel,
		eventCount: eventCount,
		organizationCount: organizationCount,
		collectEmailAndAllowEdits: collectEmailAndAllowEdits,
		eventSlots: eventSlots,
		requireDobForSchool: requireDobForSchool,
		allowOtherTeam: allowOtherTeam,
		defaultPrivacyNotice: DEFAULT_FORM_PRIVACY_NOTICE,
	};
}

/**
 * Deduplicates columns in the linked response sheet when a Form item was deleted
 * and recreated (e.g., changing LIST <-> MULTIPLE_CHOICE for Club/Team, or toggling DOB).
 * Preserves the rightmost (actively linked) column, migrates any non-empty values from
 * earlier orphaned columns where the active column is blank, and deletes the orphaned columns.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} sheetName
 * @return {number} Number of duplicate columns removed
 */
function cleanupDuplicateFormColumns_(ss, sheetName) {
	if (!ss || !sheetName) return 0;
	const sheet = ss.getSheetByName(sheetName);
	if (
		!sheet ||
		typeof sheet.getLastColumn !== "function" ||
		typeof sheet.getLastRow !== "function"
	) {
		return 0;
	}

	const lastCol = sheet.getLastColumn();
	const lastRow = sheet.getLastRow();
	if (lastCol < 2 || lastRow < 1) return 0;

	const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
	const colsByHeader = {};

	headers.forEach(function (h, idx) {
		const key = cleanFormOptionText_(h).toLowerCase();
		if (!key) return;
		if (!colsByHeader[key]) colsByHeader[key] = [];
		colsByHeader[key].push(idx);
	});

	const colsToDelete = [];
	const data =
		lastRow > 1
			? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues()
			: null;

	Object.keys(colsByHeader).forEach(function (key) {
		const indices = colsByHeader[key];
		if (indices.length <= 1) return;

		// Keep the rightmost column (bound to the most recently created Form item)
		const keepIdx = indices[indices.length - 1];
		const orphanedIndices = indices.slice(0, -1);

		if (data) {
			let migrated = false;
			for (let r = 0; r < data.length; r++) {
				if (cleanFormOptionText_(data[r][keepIdx]) !== "") continue;
				for (let j = orphanedIndices.length - 1; j >= 0; j--) {
					const val = data[r][orphanedIndices[j]];
					if (cleanFormOptionText_(val) !== "") {
						data[r][keepIdx] = val;
						migrated = true;
						break;
					}
				}
			}
			if (migrated) {
				const keepColValues = data.map(function (row) {
					return [row[keepIdx]];
				});
				sheet
					.getRange(2, keepIdx + 1, lastRow - 1, 1)
					.setValues(keepColValues);
			}
		}

		orphanedIndices.forEach(function (idx) {
			colsToDelete.push(idx);
		});
	});

	if (colsToDelete.length === 0) return 0;

	// Delete right-to-left so earlier column indices do not shift
	colsToDelete.sort(function (a, b) {
		return b - a;
	});
	colsToDelete.forEach(function (idx) {
		sheet.deleteColumn(idx + 1);
	});

	Logger.log(
		'[formBuilder] Removed %s duplicate orphaned column(s) from "%s".',
		colsToDelete.length,
		sheetName,
	);
	return colsToDelete.length;
}

/**
 * Finds an existing Form item by exact title (case-insensitive) and matching ItemType,
 * or creates a new item of the requested type, returning the typed item instance.
 *
 * @param {GoogleAppsScript.Forms.Form} form
 * @param {GoogleAppsScript.Forms.Item[]} existingItems
 * @param {string} title
 * @param {string} expectedType - 'TEXT' | 'DATE' | 'LIST' | 'MULTIPLE_CHOICE'
 * @return {GoogleAppsScript.Forms.TextItem|GoogleAppsScript.Forms.DateItem|GoogleAppsScript.Forms.ListItem|GoogleAppsScript.Forms.MultipleChoiceItem}
 */
function findOrCreateFormItem_(form, existingItems, title, expectedType) {
	const targetTitle = title.trim().toLowerCase();
	const matches = existingItems.filter(function (item) {
		return (
			cleanFormOptionText_(item.getTitle()).toLowerCase() === targetTitle
		);
	});

	if (matches.length > 0) {
		const exactMatch = matches.find(function (item) {
			return String(item.getType()) === expectedType;
		});

		if (exactMatch) {
			// Delete any extra duplicate items with the same title
			matches.forEach(function (item) {
				if (item !== exactMatch) {
					try {
						form.deleteItem(item);
					} catch (e) {
						Logger.log(
							'[formBuilder] Could not delete duplicate item "%s": %s',
							title,
							e.message,
						);
					}
				}
			});
			if (expectedType === "TEXT") return exactMatch.asTextItem();
			if (expectedType === "DATE") return exactMatch.asDateItem();
			if (expectedType === "LIST") return exactMatch.asListItem();
			if (expectedType === "MULTIPLE_CHOICE")
				return exactMatch.asMultipleChoiceItem();
		}

		// Remove all mismatched-type items so it can be recreated cleanly
		matches.forEach(function (item) {
			try {
				form.deleteItem(item);
			} catch (e) {
				Logger.log(
					'[formBuilder] Could not delete mismatched item "%s": %s',
					title,
					e.message,
				);
			}
		});
	}

	let created;
	if (expectedType === "TEXT") {
		created = form.addTextItem();
	} else if (expectedType === "DATE") {
		created = form.addDateItem();
	} else if (expectedType === "LIST") {
		created = form.addListItem();
	} else if (expectedType === "MULTIPLE_CHOICE") {
		created = form.addMultipleChoiceItem();
	} else {
		throw new Error("Unsupported Form item type: " + expectedType);
	}

	created.setTitle(title);
	return created;
}

/**
 * Creates a new Google Form linked to the active spreadsheet, or syncs/updates
 * the existing Form in place using the stored Form ID.
 *
 * @param {Object} [options]
 * @param {number} [options.eventSlots=5] - Number of Event/Time slot pairs (1..9)
 * @param {boolean} [options.collectEmailAndAllowEdits=true] - Collect email and allow response edits
 * @param {boolean} [options.requireDobForSchool=false] - Whether DOB is required for School/College meets
 * @param {boolean} [options.allowOtherTeam=false] - Allow write-in "Other" team for Open/Masters
 * @param {string} [options.formDescription] - Optional Form description / privacy notice
 * @return {Object} Summary with formId, publishedUrl, editUrl, destinationSheet, eventCount, organizationCount
 */
function createOrUpdateEntryForm(options) {
	options = options || {};
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const prereq = validateFormPrerequisites_(ss);
	const { meetInfo, compPreset, events, organizations, schoolYears } = prereq;

	const rawSlots = parseInt(options.eventSlots, 10);
	const eventSlots =
		!isNaN(rawSlots) && rawSlots >= 1 && rawSlots <= 9 ? rawSlots : 5;
	const collectEmailAndAllowEdits =
		options.collectEmailAndAllowEdits !== false;
	const requireDobForSchool = Boolean(options.requireDobForSchool);
	const allowOtherTeam =
		Boolean(options.allowOtherTeam) && !compPreset.hasSchoolYear;

	const props = PropertiesService.getDocumentProperties();
	const storedFormId = props.getProperty(FORM_PROPERTIES_KEYS.FORM_ID);

	let form = null;
	let isUpdate = false;
	let destSheetName =
		props.getProperty(FORM_PROPERTIES_KEYS.DEST_SHEET) || "";

	if (storedFormId) {
		try {
			form = FormApp.openById(storedFormId);
			isUpdate = true;
		} catch (e) {
			Logger.log(
				'[formBuilder] Stored formId "%s" unavailable (%s); creating a new Form.',
				storedFormId,
				e.message,
			);
			form = null;
			isUpdate = false;
		}
	}

	const formTitle = (meetInfo.meetName || "Swim Meet") + " - Swimmer Entries";
	const existingSheetNamesBefore = new Set(
		ss.getSheets().map((s) => s.getName()),
	);

	if (!form) {
		form = FormApp.create(formTitle);
		form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
		SpreadsheetApp.flush();

		// Locate newly created Form Responses sheet and rename to 'Form Entries' if available
		const sheetsAfter = ss.getSheets();
		const newResponseSheet = sheetsAfter.find(
			(s) => !existingSheetNamesBefore.has(s.getName()),
		);
		if (newResponseSheet) {
			if (!existingSheetNamesBefore.has("Form Entries")) {
				newResponseSheet.setName("Form Entries");
				destSheetName = "Form Entries";
			} else {
				destSheetName = newResponseSheet.getName();
			}
		} else {
			destSheetName = "Form Responses 1";
		}
	} else if (!destSheetName) {
		const formEntriesSheet = ss.getSheetByName("Form Entries");
		destSheetName = formEntriesSheet ? "Form Entries" : "Form Responses 1";
	}

	const descriptionText =
		cleanFormOptionText_(options.formDescription) ||
		`Submit individual swimmer entries for ${meetInfo.meetName || "this swim meet"} (${compPreset.name}).\n\n${DEFAULT_FORM_PRIVACY_NOTICE}`;

	form.setTitle(formTitle);
	form.setDescription(descriptionText);
	form.setCollectEmail(collectEmailAndAllowEdits);
	form.setAllowResponseEdits(collectEmailAndAllowEdits);

	const existingItems = form.getItems();

	// 1. First Name
	const firstNameItem = findOrCreateFormItem_(
		form,
		existingItems,
		"First Name",
		"TEXT",
	);
	firstNameItem.setRequired(true);

	// 2. Last Name
	const lastNameItem = findOrCreateFormItem_(
		form,
		existingItems,
		"Last Name",
		"TEXT",
	);
	lastNameItem.setRequired(true);

	// 3. Date of Birth (always shown & required for non-school meets; shown for School/College only when requireDobForSchool is checked)
	const includeDob = !compPreset.hasSchoolYear || requireDobForSchool;
	if (includeDob) {
		const dobItem = findOrCreateFormItem_(
			form,
			existingItems,
			"Date of Birth",
			"DATE",
		);
		dobItem.setRequired(true);
		dobItem.setHelpText("Enter the swimmer's date of birth.");
	} else {
		const existingDobItem = existingItems.find(function (item) {
			return (
				cleanFormOptionText_(item.getTitle()).toLowerCase() ===
				"date of birth"
			);
		});
		if (existingDobItem) {
			try {
				form.deleteItem(existingDobItem);
			} catch (e) {
				Logger.log(
					"[formBuilder] Could not delete Date of Birth item: %s",
					e.message,
				);
			}
		}
	}

	// 4. Gender
	const genderItem = findOrCreateFormItem_(
		form,
		existingItems,
		"Gender",
		"LIST",
	);
	genderItem.setChoiceValues(["Female", "Male"]);
	genderItem.setRequired(true);

	// 5. School Year (School / College only)
	if (compPreset.hasSchoolYear) {
		const syItem = findOrCreateFormItem_(
			form,
			existingItems,
			"School Year",
			"LIST",
		);
		syItem.setChoiceValues(schoolYears);
		syItem.setRequired(true);
	}

	// 6. School / Team
	const orgTitle = compPreset.orgLabel || "School";
	if (allowOtherTeam) {
		const orgMcItem = findOrCreateFormItem_(
			form,
			existingItems,
			orgTitle,
			"MULTIPLE_CHOICE",
		);
		orgMcItem.setChoiceValues(organizations);
		orgMcItem.showOtherOption(true);
		orgMcItem.setRequired(true);
		orgMcItem.setHelpText(
			'Select your team/club, or choose "Other" to enter an unattached/unlisted team.',
		);
	} else {
		const orgListItem = findOrCreateFormItem_(
			form,
			existingItems,
			orgTitle,
			"LIST",
		);
		orgListItem.setChoiceValues(organizations);
		orgListItem.setRequired(true);
	}

	// 7. Event 1..N and Time 1..N pairs (METHOD_1)
	const timeValidation = FormApp.createTextValidation()
		.setHelpText(
			"Enter a valid time as MM:SS.SS (e.g. 01:15.40), SS.SS (e.g. 34.50), or NT",
		)
		.requireTextMatchesPattern("^(\\d{1,2}:)?\\d{1,2}\\.\\d{2}$|^NT$")
		.build();

	for (let i = 1; i <= eventSlots; i++) {
		const eventItem = findOrCreateFormItem_(
			form,
			existingItems,
			"Event " + i,
			"LIST",
		);
		eventItem.setChoiceValues(events);
		eventItem.setRequired(i === 1);

		const timeItem = findOrCreateFormItem_(
			form,
			existingItems,
			"Time " + i,
			"TEXT",
		);
		timeItem.setRequired(false);
		timeItem.setHelpText(
			"Seed time in MM:SS.SS or SS.SS format (e.g. 1:15.40 or 34.50), or NT. Leave blank for NT.",
		);
		timeItem.setValidation(timeValidation);
	}

	// Flush pending Form/Sheet updates and remove any orphaned duplicate columns in the linked sheet
	if (typeof SpreadsheetApp.flush === "function") {
		SpreadsheetApp.flush();
	}
	cleanupDuplicateFormColumns_(ss, destSheetName);

	const publishedUrl = form.getPublishedUrl();
	const editUrl = form.getEditUrl();
	const formId = form.getId();

	props.setProperties({
		[FORM_PROPERTIES_KEYS.FORM_ID]: formId,
		[FORM_PROPERTIES_KEYS.PUBLISHED_URL]: publishedUrl,
		[FORM_PROPERTIES_KEYS.EDIT_URL]: editUrl,
		[FORM_PROPERTIES_KEYS.DEST_SHEET]: destSheetName,
	});

	Logger.log(
		'[formBuilder] %s Form "%s" (id=%s) linked to sheet "%s"',
		isUpdate ? "Updated" : "Created",
		formTitle,
		formId,
		destSheetName,
	);

	return {
		success: true,
		isUpdate: isUpdate,
		formId: formId,
		publishedUrl: publishedUrl,
		editUrl: editUrl,
		destinationSheet: destSheetName,
		eventCount: events.length,
		organizationCount: organizations.length,
	};
}
