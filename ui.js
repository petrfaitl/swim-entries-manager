/**
 * Add menu item to Custom Tools
 */
function onOpen(e) {
	const menu = SpreadsheetApp.getUi().createAddonMenu();
	if (e && e.authMode === ScriptApp.AuthMode.NONE) {
		Logger.log("Not enough permissions to create menu yet.");
		menu.addItem("Start Add-on", "showSidebar");
	} else {
		menu.addItem("Meet Setup", "showMeetSetupDialog")
			.addItem("Create Sheets from Template", "showCreateSheetsDialog")
			.addItem("Create / Sync Entry Form", "showFormSetupDialog")
			.addItem("Create Relay Tables", "showCreateRelayTablesDialog")
			.addSeparator()
			.addItem("Export for Meet Manager", "showSDIFCreatorDialog")
			.addItem("Export to CSV", "exportEntriesToCSV");
	}
	menu.addToUi();
}

function onHomepage(e) {
	onOpen(e);
	return buildAdminSidebarCard();
}

// Alias for manifest homepageTrigger if needed
function showHomepage(e) {
	return onHomepage(e);
}

function showCreateTablesDialog() {
	return showMeetSetupDialog();
}

function showCreateRelayTablesDialog() {
	const html = HtmlService.createTemplateFromFile("relay-ui-full")
		.evaluate()
		.setWidth(520)
		.setHeight(720);

	SpreadsheetApp.getUi().showModalDialog(html, "Create Relay Tables");
}

/**
 * Opens the Google Form Setup & Sync modal dialog.
 */
function showFormSetupDialog() {
	const html = HtmlService.createTemplateFromFile("form-setup-ui")
		.evaluate()
		.setWidth(520)
		.setHeight(640);

	SpreadsheetApp.getUi().showModalDialog(html, "Create / Sync Entry Form");
}

function buildAdminSidebarCard() {
	const encodedImageURL =
		"https://www.lvwasc.co.nz/wp-content/uploads/2026/03/icon_pool_swimmer.png";

	// ── Card 1: Templates ──────────────────────────────────────────────────
	const templatesCard = CardService.newCardBuilder().setHeader(
		CardService.newCardHeader()
			.setTitle("Swim Meet Templates")
			.setSubtitle(
				"Creates essential template tables for managing a meet.",
			)
			.setImageUrl(encodedImageURL),
	);

	// Section 1: Meet Setup
	const createTablesSection = CardService.newCardSection()
		.setHeader("Meet Setup")
		.setCollapsible(true)
		.addWidget(
			CardService.newDecoratedText()
				.setText("<b>STEP 1</b>")
				.setStartIcon(
					CardService.newIconImage().setMaterialIcon(
						CardService.newMaterialIcon().setName("looks_one"),
					),
				),
		)
		.addWidget(
			CardService.newTextParagraph().setText(
				"Configure your meet parameters, build your event and team sheets and swimmer entry templates.",
			),
		)
		.addWidget(
			CardService.newButtonSet().addButton(
				CardService.newTextButton()
					.setText("Meet Setup")
					.setTextButtonStyle(CardService.TextButtonStyle.FILLED)
					.setBackgroundColor("#0397B3")
					.setOnClickAction(
						CardService.newAction().setFunctionName(
							"showMeetSetupDialog",
						),
					),
			),
		);

	// Section 2: Create Sheets from Templates
	const createSheetsSection = CardService.newCardSection()
		.setHeader("Create Team Entry Sheets")
		.setCollapsible(true)
		.addWidget(
			CardService.newDecoratedText()
				.setText("<b>STEP 2</b>")
				.setStartIcon(
					CardService.newIconImage().setMaterialIcon(
						CardService.newMaterialIcon().setName("looks_two"),
					),
				),
		)
		.addWidget(
			CardService.newTextParagraph().setText(
				"After setting up your meet parameters, duplicate the Individual Entry template sheet for each school/cluster or team to allow coordinators to fill in entries separately.",
			),
		)
		.addWidget(
			CardService.newButtonSet().addButton(
				CardService.newTextButton()
					.setText("Create Team Sheets")
					.setTextButtonStyle(CardService.TextButtonStyle.FILLED)
					.setBackgroundColor("#0397B3")
					.setOnClickAction(
						CardService.newAction().setFunctionName(
							"showCreateSheetsDialog",
						),
					),
			),
		);

	// Section 2b: Create / Sync Entry Form (Optional)
	const createFormSection = CardService.newCardSection()
		.setHeader("Entry Form (Optional)")
		.setCollapsible(true)
		.addWidget(
			CardService.newTextParagraph().setText(
				"Generate or sync a public Google Form for parents or individual swimmers to submit entries directly into your workbook.",
			),
		)
		.addWidget(
			CardService.newButtonSet().addButton(
				CardService.newTextButton()
					.setText("Create / Sync Entry Form")
					.setTextButtonStyle(CardService.TextButtonStyle.FILLED)
					.setBackgroundColor("#0397B3")
					.setOnClickAction(
						CardService.newAction().setFunctionName(
							"showFormSetupDialog",
						),
					),
			),
		);

	// Section 3: Create Relay Tables (Optional)
	const createRelaySection = CardService.newCardSection()
		.setHeader("Create Relay Tables (Optional)")
		.setCollapsible(true)
		.addWidget(
			CardService.newDecoratedText()
				.setText("<b>STEP 3</b>")
				.setStartIcon(
					CardService.newIconImage().setMaterialIcon(
						CardService.newMaterialIcon().setName("looks_3"),
					),
				),
		)
		.addWidget(
			CardService.newTextParagraph().setText(
				"This is optional and only needed if your meet includes relay events.",
			),
		)
		.addWidget(
			CardService.newButtonSet().addButton(
				CardService.newTextButton()
					.setText("Create Relay Tables")
					.setTextButtonStyle(CardService.TextButtonStyle.FILLED)
					.setBackgroundColor("#0397B3")
					.setOnClickAction(
						CardService.newAction().setFunctionName(
							"showCreateRelayTablesDialog",
						),
					),
			),
		);

	// Navigation to Export Card
	const navToExportSection = CardService.newCardSection()

		.addWidget(
			CardService.newTextParagraph().setText(
				"<b>Next:</b> Export entries for Meet Manager",
			),
		)
		.addWidget(
			CardService.newButtonSet().addButton(
				CardService.newTextButton()
					.setText("Go to Export →")
					.setTextButtonStyle(CardService.TextButtonStyle.FILLED)
					.setOnClickAction(
						CardService.newAction().setFunctionName(
							"buildExportCard",
						),
					),
			),
		);

	templatesCard
		.addSection(createTablesSection)
		.addSection(createSheetsSection)
		.addSection(createFormSection)
		.addSection(createRelaySection)
		.addSection(navToExportSection)
		.setFixedFooter(
			CardService.newFixedFooter()
				.setPrimaryButton(
					CardService.newTextButton()
						.setTextButtonStyle(
							CardService.TextButtonStyle.BORDERLESS,
						)
						.setBackgroundColor("#e2f9ff")
						.setText("Help")
						.setIcon(CardService.Icon.DESCRIPTION)
						.setOnClickAction(
							CardService.newAction().setFunctionName(
								"showHelpCard",
							),
						),
				)
				.setSecondaryButton(
					CardService.newTextButton()
						.setTextButtonStyle(
							CardService.TextButtonStyle.BORDERLESS,
						)
						.setText("About")
						.setIcon(CardService.Icon.BOOKMARK)
						.setOnClickAction(
							CardService.newAction().setFunctionName(
								"showAboutCard",
							),
						),
				),
		);

	return templatesCard.build();
}

function buildExportCard() {
	const encodedImageURL =
		"https://www.lvwasc.co.nz/wp-content/uploads/2026/03/icon_file_download.png";

	// ── Card 2: Export ─────────────────────────────────────────────────────
	const exportCard = CardService.newCardBuilder().setHeader(
		CardService.newCardHeader()
			.setTitle("Export Entries")
			.setSubtitle("Generate meet entry files for Hy-Tek Meet Manager.")
			.setImageUrl(encodedImageURL),
	);

	const sdifExportSection = CardService.newCardSection()
		.setHeader("Export for Meet Manager (Recommended)")
		.setCollapsible(false)
		.addWidget(
			CardService.newDecoratedText()
				.setText("<b>STEP 4</b>")
				.setStartIcon(
					CardService.newIconImage().setMaterialIcon(
						CardService.newMaterialIcon().setName("looks_4"),
					),
				),
		)
		.addWidget(
			CardService.newTextParagraph().setText(
				"Generate an SDIF (.sd3) file ready for import into Hy-Tek Meet Manager. This is the recommended export method for most users.",
			),
		)
		.addWidget(
			CardService.newButtonSet().addButton(
				CardService.newTextButton()
					.setText("Export for Meet Manager")
					.setTextButtonStyle(CardService.TextButtonStyle.FILLED)
					.setBackgroundColor("#0397B3")
					.setOnClickAction(
						CardService.newAction().setFunctionName(
							"showSDIFCreatorDialog",
						),
					),
			),
		);

	const csvExportSection = CardService.newCardSection()
		.setHeader("Alternative: Export to CSV")
		.setCollapsible(true)
		.addWidget(
			CardService.newTextParagraph().setText(
				"Export entries as a CSV file for manual processing or use with external tools.",
			),
		)
		.addWidget(
			CardService.newButtonSet().addButton(
				CardService.newTextButton()
					.setText("Export to CSV")
					.setTextButtonStyle(CardService.TextButtonStyle.FILLED)
					.setBackgroundColor("#0397B3")
					.setOnClickAction(
						CardService.newAction().setFunctionName(
							"exportEntriesToCSV",
						),
					),
			),
		);

	// Navigation back to Templates
	const navBackSection = CardService.newCardSection()

		.addWidget(
			CardService.newButtonSet().addButton(
				CardService.newTextButton()
					.setText("← Back to Templates")
					.setTextButtonStyle(CardService.TextButtonStyle.FILLED)
					.setOnClickAction(
						CardService.newAction().setFunctionName(
							"buildAdminSidebarCard",
						),
					),
			),
		);

	exportCard
		.addSection(sdifExportSection)
		.addSection(csvExportSection)
		.addSection(navBackSection)
		.setFixedFooter(
			CardService.newFixedFooter()
				.setPrimaryButton(
					CardService.newTextButton()
						.setTextButtonStyle(CardService.TextButtonStyle.FILLED)
						.setBackgroundColor("#0397B3")
						.setText("Help")
						.setIcon(CardService.Icon.DESCRIPTION)
						.setOnClickAction(
							CardService.newAction().setFunctionName(
								"showHelpCard",
							),
						),
				)
				.setSecondaryButton(
					CardService.newTextButton()
						.setTextButtonStyle(CardService.TextButtonStyle.TEXT)
						.setText("About")
						.setIcon(CardService.Icon.BOOKMARK)
						.setOnClickAction(
							CardService.newAction().setFunctionName(
								"showAboutCard",
							),
						),
				),
		);

	return CardService.newNavigation().pushCard(exportCard.build());
}

function showSDIFCreatorDialog() {
	SDIFCreator.showDialog();
}

/**
 * Builds and returns the About card with project description and coffee link
 */
function showAboutCard() {
	const aboutCard = CardService.newCardBuilder().setHeader(
		CardService.newCardHeader()
			.setTitle("About Swim Entries Manager")
			.setSubtitle("v2.0"),
	);

	const descriptionSection = CardService.newCardSection()
		.addWidget(
			CardService.newTextParagraph().setText(
				"<b>Swim Entries Manager</b> is a Google Workspace Add-on for managing swim meet entries efficiently.",
			),
		)
		.addWidget(
			CardService.newTextParagraph().setText(
				"This tool helps swim meet organisers create structured entry templates, manage school/cluster registrations, and export data directly to SDIF format for Hy-Tek Meet Manager or CSV for other tools.",
			),
		);

	const featuresSection = CardService.newCardSection()
		.setHeader("Key Features")
		.addWidget(
			CardService.newTextParagraph().setText(
				"• <b>Template Management:</b> Generate structured tables with automated validation and named ranges<br><br>• <b>Sheet Duplication:</b> Bulk creation of entry sheets for multiple schools or clusters<br><br>• <b>SDIF Export:</b> Direct export to SDIF v3 format with exception reporting and data validation<br><br>• <b>CSV Export:</b> Traditional CSV format for custom processing",
			),
		);

	const linksSection = CardService.newCardSection()
		.setHeader("Resources")
		.addWidget(
			CardService.newButtonSet().addButton(
				CardService.newTextButton()
					.setText("View on GitHub")
					.setIcon(CardService.Icon.DESCRIPTION)
					.setOpenLink(
						CardService.newOpenLink().setUrl(
							"https://github.com/petrfaitl/swim-entries-manager",
						),
					),
			),
		)
		.addWidget(
			CardService.newButtonSet().addButton(
				CardService.newTextButton()
					.setText("Report Issues")
					.setIcon(CardService.Icon.EMAIL)
					.setOpenLink(
						CardService.newOpenLink().setUrl(
							"https://github.com/petrfaitl/swim-entries-manager/issues",
						),
					),
			),
		);

	const supportSection = CardService.newCardSection()
		.setHeader("Support Development")
		.addWidget(
			CardService.newTextParagraph().setText(
				"If you find this add-on useful, consider supporting its development:",
			),
		)
		.addWidget(
			CardService.newButtonSet().addButton(
				CardService.newTextButton()
					.setText("Buy me a coffee")
					.setTextButtonStyle(CardService.TextButtonStyle.FILLED)
					.setBackgroundColor("#FFDD00")
					.setIcon(CardService.Icon.STORE)
					.setOpenLink(
						CardService.newOpenLink().setUrl(
							"https://buymeacoffee.com/pfaitl",
						),
					),
			),
		);

	const legalSection = CardService.newCardSection().addWidget(
		CardService.newTextParagraph().setText(
			'<font color="#666666"><i>Licensed under MIT License<br>Built with Google Apps Script<br>Contact: recorder@lvwasc.co.nz</i></font>',
		),
	);

	aboutCard
		.addSection(descriptionSection)
		.addSection(featuresSection)
		.addSection(linksSection)
		.addSection(supportSection)
		.addSection(legalSection);

	return CardService.newNavigation().pushCard(aboutCard.build());
}

/**
 * Builds and returns the Help card with essential user guide information
 */
function showHelpCard() {
	const helpCard = CardService.newCardBuilder().setHeader(
		CardService.newCardHeader()
			.setTitle("Help & Quick Start")
			.setSubtitle("Essential guide to using Swim Entries Manager"),
	);

	const quickStartSection = CardService.newCardSection()
		.setHeader("Quick Start Workflow")
		.addWidget(
			CardService.newTextParagraph().setText(
				"<b>1. Meet Setup (STEP 1)</b><br>Select your Competition Type (<i>School</i>, <i>College</i>, <i>Age Group</i>, <i>Open</i>, or <i>Masters</i>) and enter meet details. This auto-generates <i>Meet</i> (with <i>AgeBrackets</i> for School/College), <i>Events</i>, <i>Schools</i> or <i>Teams</i>, and <i>Individual Events Template</i> (plus optional <i>Team Officials</i> and <i>How to fill entries</i>).",
			),
		)
		.addWidget(
			CardService.newTextParagraph().setText(
				"<b>2. Review &amp; Populate Tables</b><br>• <i>Events:</i> Enter Distance and Discipline (Event name is auto-generated)<br>• <i>Schools / Teams:</i> Add participating organisations with Code and Team Name<br>• <i>Team Officials (if enabled):</i> Enter team/school/cluster officials",
			),
		)
		.addWidget(
			CardService.newTextParagraph().setText(
				"<b>3. Collect Entries (STEP 2 / Optional)</b><br>• <i>Create Team Entry Sheets (STEP 2):</i> Duplicate <i>Individual Events Template</i> per school, cluster, or team via <b>Create Team Sheets</b>.<br>• <i>Entry Form (Optional):</i> Generate a public Google Form via <b>Create / Sync Entry Form</b> (responses land in <i>Form Entries</i>).<br>• <i>Create Relay Tables (STEP 3, Optional):</i> Generate relay sheets if your meet includes relay events.",
			),
		)
		.addWidget(
			CardService.newTextParagraph().setText(
				"<b>4. Fill Swimmer Entries</b><br><b>Required:</b> First Name, Last Name, Gender, valid School/Team/Club, plus Date of Birth (<i>Age Group</i>, <i>Open</i>, <i>Masters</i>) or School Year (<i>School</i>, <i>College</i>)<br><b>Recommended:</b> Entry times",
			),
		)
		.addWidget(
			CardService.newTextParagraph().setText(
				"<b>5. Export for Meet Manager (STEP 4)</b><br>Generate an SDIF (.sd3) file and review any Exception Report warnings/errors. Alternative: <b>Export to CSV</b>.",
			),
		);

	const criticalFieldsSection = CardService.newCardSection()
		.setHeader("Critical Data Requirements")
		.addWidget(
			CardService.newTextParagraph().setText(
				"<b>Must be present or swimmer is excluded:</b><br>• First Name &amp; Last Name<br>• Gender<br>• Date of Birth (DD/MM/YYYY) — required for <i>Age Group</i>, <i>Open</i>, and <i>Masters</i>; <i>School</i> and <i>College</i> meets can use School Year (mapped via <i>AgeBrackets</i> on <i>Meet</i>) when DOB is omitted<br>• Team Code &amp; Team Name (must exist in the <i>Schools</i> or <i>Teams</i> table)",
			),
		)
		.addWidget(
			CardService.newTextParagraph().setText(
				'<b>Event Format:</b><br>Events must be formatted as [Distance] [Stroke]<br>✓ Correct: "25m Freestyle", "50m Backstroke"<br>✗ Incorrect: "Y5", "25 Free"',
			),
		);

	const exceptionReportsSection = CardService.newCardSection()
		.setHeader("Understanding Exception Reports")
		.addWidget(
			CardService.newTextParagraph().setText(
				"When exporting for Meet Manager, the system validates all critical data. If issues are found, an <b>Exception Report</b> is generated that details:",
			),
		)
		.addWidget(
			CardService.newTextParagraph().setText(
				"• Missing critical data (swimmer excluded)<br>• Missing teams in Schools / Teams table<br>• Invalid event format (events skipped)<br>• Duplicate event entries (duplicates skipped, warning only)<br>• Swimmers with no events (warning only)",
			),
		)
		.addWidget(
			CardService.newTextParagraph().setText(
				"<b>Always review exception reports before importing to Meet Manager!</b>",
			),
		);

	const tipsSection = CardService.newCardSection()
		.setHeader("Best Practices")
		.addWidget(
			CardService.newTextParagraph().setText(
				"✓ Fill in all required fields before exporting<br>✓ Use dropdown values when available<br>✓ Add all teams/schools to the Schools or Teams table first<br>✓ Re-sync your Entry Form (<b>Create / Sync Entry Form</b>) after updating Events, Schools, or Teams<br>✓ Use correct event naming format<br>✓ Run test export early to catch data issues",
			),
		);

	const resourcesSection = CardService.newCardSection()
		.setHeader("Need More Help?")
		.addWidget(
			CardService.newButtonSet().addButton(
				CardService.newTextButton()
					.setText("Full User Guide")
					.setIcon(CardService.Icon.DESCRIPTION)
					.setOpenLink(
						CardService.newOpenLink().setUrl(
							"https://github.com/petrfaitl/swim-entries-manager/blob/main/docs/USER_GUIDE.md",
						),
					),
			),
		)
		.addWidget(
			CardService.newButtonSet().addButton(
				CardService.newTextButton()
					.setText("Report Issues")
					.setIcon(CardService.Icon.EMAIL)
					.setOpenLink(
						CardService.newOpenLink().setUrl(
							"https://github.com/petrfaitl/swim-entries-manager/issues",
						),
					),
			),
		)
		.addWidget(
			CardService.newTextParagraph().setText(
				'<font color="#666666"><i>Email: recorder@lvwasc.co.nz</i></font>',
			),
		);

	helpCard
		.addSection(quickStartSection)
		.addSection(criticalFieldsSection)
		.addSection(exceptionReportsSection)
		.addSection(tipsSection)
		.addSection(resourcesSection);

	return CardService.newNavigation().pushCard(helpCard.build());
}

/**
 * Gets the names of all visible sheets in the active spreadsheet.
 * @returns {string[]} Array of sheet names.
 */
function getVisibleSheetNames() {
	return SpreadsheetApp.getActiveSpreadsheet()
		.getSheets()
		.filter((sheet) => !sheet.isSheetHidden())
		.map((sheet) => sheet.getName());
}
