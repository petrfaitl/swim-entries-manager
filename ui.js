/**
 * Add menu item to Custom Tools
 */
function onOpen(e) {
  const menu = SpreadsheetApp.getUi()
                             .createAddonMenu();
  if (e && e.authMode === ScriptApp.AuthMode.NONE) {
    Logger.log("Not enough permissions to create menu yet.");
    menu.addItem('Start Add-on', 'showSidebar');
  } else {
    menu.addItem('Create Core Tables', 'showCreateTablesDialog')
        .addItem('Create Sheets from Template', 'showCreateSheetsDialog')
        .addItem('Create Relay Tables', 'showCreateRelayTablesDialog')
        .addSeparator()
        .addItem('Export for Meet Manager', 'showSDIFCreatorDialog')
        .addItem('Export to CSV', 'exportEntriesToCSV');
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
  const tables = listAvailableTables();
  Logger.log('[showCreateTablesDialog] tables payload: %s', JSON.stringify(tables));

  const template = HtmlService.createTemplateFromFile('core-tables-ui');
  template.tableListHtml = tables.length
    ? tables.map(function (table) {
        return '<label style="display:flex; align-items:center; gap:6px; margin:4px 0; font-size:11.5px; cursor:pointer;"><input type="checkbox" value="' + table.name + '" style="width:auto; cursor:pointer;"> <span>' + table.title + '</span></label>';
      }).join('')
    : 'No tables available.';
  template.tableCount = tables.length;

  const html = template.evaluate()
                       .setWidth(520)
                       .setHeight(600);

  SpreadsheetApp.getUi()
                .showModalDialog(html, 'Create Core Tables');
}

function showCreateRelayTablesDialog() {
  const html = HtmlService.createTemplateFromFile('relay-ui-full')
                          .evaluate()
                          .setWidth(520)
                          .setHeight(720);

  SpreadsheetApp.getUi()
                .showModalDialog(html, 'Create Relay Tables');
}

function buildAdminSidebarCard() {
  const imageBytes = DriveApp.getFileById('1PT-hMqaBpVTNpkN3KiocfgNa7QRWfmXw')
                             .getBlob()
                             .getBytes();
  const encodedImageURL = `data:image/jpeg;base64,${Utilities.base64Encode(imageBytes)}`;

  // ── Card 1: Templates ──────────────────────────────────────────────────
  const templatesCard = CardService.newCardBuilder()
                                   .setHeader(
                                     CardService.newCardHeader()
                                                .setTitle('Swim Meet Templates')
                                                .setSubtitle('Creates essential template tables for managing a meet.')
                                                .setImageUrl(encodedImageURL)
                                   );

  // Section 1: Create Core Tables
  const createTablesSection = CardService.newCardSection()
                                         .setHeader('Create Core Tables')
                                         .setCollapsible(true)
                                         .addWidget(
                                           CardService.newDecoratedText()
                                                      .setText('<b>STEP 1</b>')
                                                      .setStartIcon(CardService.newIconImage()
                                                                               .setMaterialIcon(
                                                                                 CardService.newMaterialIcon()
                                                                                            .setName('looks_one')
                                                                               ))
                                         )
                                         .addWidget(
                                           CardService.newTextParagraph()
                                                      .setText('Generate core tables from a list of tables.')
                                         )
                                         .addWidget(
                                           CardService.newButtonSet()
                                                      .addButton(
                                                        CardService.newTextButton()
                                                                   .setText('Create Core Tables')
                                                                   .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
                                                                   .setBackgroundColor('#0397B3')
                                                                   .setOnClickAction(CardService.newAction()
                                                                                                .setFunctionName('showCreateTablesDialog'))
                                                      )
                                         );

  // Section 2: Create Sheets from Templates
  const createSheetsSection = CardService.newCardSection()
                                         .setHeader('Create Sheets from Templates')
                                         .setCollapsible(true)
                                         .addWidget(
                                           CardService.newDecoratedText()
                                                      .setText('<b>STEP 2</b>')
                                                      .setStartIcon(CardService.newIconImage()
                                                                               .setMaterialIcon(
                                                                                 CardService.newMaterialIcon()
                                                                                            .setName('looks_two')
                                                                               ))
                                         )
                                         .addWidget(
                                           CardService.newTextParagraph()
                                                      .setText('After populating core tables, duplicate the Individual Entry template sheet for each school or cluster from a list of schools/clusters in your spreadsheet.')
                                         )
                                         .addWidget(
                                           CardService.newButtonSet()
                                                      .addButton(
                                                        CardService.newTextButton()
                                                                   .setText('Create Sheets from Templates')
                                                                   .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
                                                                   .setBackgroundColor('#0397B3')
                                                                   .setOnClickAction(CardService.newAction()
                                                                                                .setFunctionName('showCreateSheetsDialog'))
                                                      )
                                         );

  // Section 3: Create Relay Tables (Optional)
  const createRelaySection = CardService.newCardSection()
                                        .setHeader('Create Relay Tables (Optional)')
                                        .setCollapsible(true)
                                        .addWidget(
                                          CardService.newDecoratedText()
                                                     .setText('<b>STEP 3</b>')
                                                     .setStartIcon(CardService.newIconImage()
                                                                              .setMaterialIcon(
                                                                                CardService.newMaterialIcon()
                                                                                           .setName('looks_3')
                                                                              ))
                                        )
                                        .addWidget(
                                          CardService.newTextParagraph()
                                                     .setText('Create relay entry tables using team names from your Team Officials sheet. Tables can be placed on one sheet or separate sheets.')
                                        )
                                        .addWidget(
                                          CardService.newButtonSet()
                                                     .addButton(
                                                       CardService.newTextButton()
                                                                  .setText('Create Relay Tables')
                                                                  .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
                                                                  .setBackgroundColor('#0397B3')
                                                                  .setOnClickAction(CardService.newAction()
                                                                                               .setFunctionName('showCreateRelayTablesDialog'))
                                                     )
                                        );

  // Navigation to Export Card
  const navToExportSection = CardService.newCardSection()

                                        .addWidget(
                                          CardService.newTextParagraph()
                                                     .setText('<b>Next:</b> Export entries for Meet Manager')
                                        )
                                        .addWidget(
                                          CardService.newButtonSet()
                                                     .addButton(
                                                       CardService.newTextButton()
                                                                  .setText('Go to Export →')
                                                                  .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
                                                                  .setOnClickAction(CardService.newAction()
                                                                                               .setFunctionName('buildExportCard'))
                                                     )
                                        );

  templatesCard
    .addSection(createTablesSection)
    .addSection(createSheetsSection)
    .addSection(createRelaySection)
    .addSection(navToExportSection)
    .setFixedFooter(
      CardService.newFixedFooter()
                 .setPrimaryButton(
                   CardService.newTextButton()
                              .setTextButtonStyle(CardService.TextButtonStyle.BORDERLESS)
                              .setBackgroundColor('#e2f9ff')
                              .setText('Help')
                              .setIcon(CardService.Icon.DESCRIPTION)
                              .setOnClickAction(CardService.newAction()
                                                           .setFunctionName('showHelpCard'))
                 )
                 .setSecondaryButton(
                   CardService.newTextButton()
                              .setTextButtonStyle(CardService.TextButtonStyle.BORDERLESS)
                              .setText('About')
                              .setIcon(CardService.Icon.BOOKMARK)
                              .setOnClickAction(CardService.newAction()
                                                           .setFunctionName('showAboutCard'))
                 )
    );

  return templatesCard.build();
}

function buildExportCard() {
  const imageBytes = DriveApp.getFileById('1DQZoQh-ya7I-rVsED3HJjVdobkEFwlvZ')
                             .getBlob()
                             .getBytes();
  const encodedImageURL = `data:image/jpeg;base64,${Utilities.base64Encode(imageBytes)}`;

  // ── Card 2: Export ─────────────────────────────────────────────────────
  const exportCard = CardService.newCardBuilder()
                                .setHeader(
                                  CardService.newCardHeader()
                                             .setTitle('Export Entries')
                                             .setSubtitle('Generate meet entry files for Hy-Tek Meet Manager.')
                                             .setImageUrl(encodedImageURL)
                                );

  const sdifExportSection = CardService.newCardSection()
                                   .setHeader('Export for Meet Manager (Recommended)')
                                   .setCollapsible(false)
                                   .addWidget(
                                     CardService.newDecoratedText()
                                                .setText('<b>STEP 4</b>')
                                                .setStartIcon(CardService.newIconImage()
                                                                         .setMaterialIcon(
                                                                           CardService.newMaterialIcon()
                                                                                      .setName('looks_4')
                                                                         ))
                                   )
                                   .addWidget(
                                     CardService.newTextParagraph()
                                                .setText('Generate an SDIF (.sd3) file ready for import into Hy-Tek Meet Manager. This is the recommended export method for most users.')
                                   )
                                   .addWidget(
                                     CardService.newButtonSet()
                                                .addButton(
                                                  CardService.newTextButton()
                                                             .setText('Export for Meet Manager')
                                                             .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
                                                             .setBackgroundColor('#0397B3')
                                                             .setOnClickAction(CardService.newAction()
                                                                                          .setFunctionName('showSDIFCreatorDialog'))
                                                )
                                   );

  const csvExportSection = CardService.newCardSection()
                                         .setHeader('Alternative: Export to CSV')
                                         .setCollapsible(true)
                                         .addWidget(
                                           CardService.newTextParagraph()
                                                      .setText('Export entries as a CSV file for manual processing or use with external tools.')
                                         )
                                         .addWidget(
                                           CardService.newButtonSet()
                                                      .addButton(
                                                        CardService.newTextButton()
                                                                   .setText('Export to CSV')
                                                                   .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
                                                                   .setBackgroundColor('#0397B3')
                                                                   .setOnClickAction(CardService.newAction()
                                                                                                .setFunctionName('exportEntriesToCSV'))
                                                      )
                                         );

  // Navigation back to Templates
  const navBackSection = CardService.newCardSection()

                                    .addWidget(
                                      CardService.newButtonSet()
                                                 .addButton(
                                                   CardService.newTextButton()
                                                              .setText('← Back to Templates')
                                                              .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
                                                              .setOnClickAction(CardService.newAction()
                                                                                           .setFunctionName('buildAdminSidebarCard'))
                                                 )
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
                              .setBackgroundColor('#0397B3')
                              .setText('Help')
                              .setIcon(CardService.Icon.DESCRIPTION)
                              .setOnClickAction(CardService.newAction()
                                                           .setFunctionName('showHelpCard'))
                 )
                 .setSecondaryButton(
                   CardService.newTextButton()
                              .setTextButtonStyle(CardService.TextButtonStyle.TEXT)
                              .setText('About')
                              .setIcon(CardService.Icon.BOOKMARK)
                              .setOnClickAction(CardService.newAction()
                                                           .setFunctionName('showAboutCard'))
                 )
    );

  return CardService.newNavigation()
                    .pushCard(exportCard.build());
}

function showSDIFCreatorDialog() {
  SDIFCreator.showDialog();
}

/**
 * Builds and returns the About card with project description and coffee link
 */
function showAboutCard() {
  const aboutCard = CardService.newCardBuilder()
                               .setHeader(
                                 CardService.newCardHeader()
                                            .setTitle('About Swim Entries Manager')
                                            .setSubtitle('v2.0')
                               );

  const descriptionSection = CardService.newCardSection()
                                        .addWidget(
                                          CardService.newTextParagraph()
                                                     .setText('<b>Swim Entries Manager</b> is a Google Workspace Add-on for managing swim meet entries efficiently.')
                                        )
                                        .addWidget(
                                          CardService.newTextParagraph()
                                                     .setText('This tool helps swim meet organisers create structured entry templates, manage school/cluster registrations, and export data directly to SDIF format for Hy-Tek Meet Manager or CSV for other tools.')
                                        );

  const featuresSection = CardService.newCardSection()
                                     .setHeader('Key Features')
                                     .addWidget(
                                       CardService.newTextParagraph()
                                                  .setText('• <b>Template Management:</b> Generate structured tables with automated validation and named ranges<br><br>• <b>Sheet Duplication:</b> Bulk creation of entry sheets for multiple schools or clusters<br><br>• <b>SDIF Export:</b> Direct export to SDIF v3 format with exception reporting and data validation<br><br>• <b>CSV Export:</b> Traditional CSV format for custom processing')
                                     );

  const linksSection = CardService.newCardSection()
                                  .setHeader('Resources')
                                  .addWidget(
                                    CardService.newButtonSet()
                                               .addButton(
                                                 CardService.newTextButton()
                                                            .setText('View on GitHub')
                                                            .setIcon(CardService.Icon.DESCRIPTION)
                                                            .setOpenLink(CardService.newOpenLink()
                                                                                   .setUrl('https://github.com/petrfaitl/swim-entries-manager'))
                                               )
                                  )
                                  .addWidget(
                                    CardService.newButtonSet()
                                               .addButton(
                                                 CardService.newTextButton()
                                                            .setText('Report Issues')
                                                            .setIcon(CardService.Icon.EMAIL)
                                                            .setOpenLink(CardService.newOpenLink()
                                                                                   .setUrl('https://github.com/petrfaitl/swim-entries-manager/issues'))
                                               )
                                  );

  const supportSection = CardService.newCardSection()
                                    .setHeader('Support Development')
                                    .addWidget(
                                      CardService.newTextParagraph()
                                                 .setText('If you find this add-on useful, consider supporting its development:')
                                    )
                                    .addWidget(
                                      CardService.newButtonSet()
                                                 .addButton(
                                                   CardService.newTextButton()
                                                              .setText('Buy me a coffee')
                                                              .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
                                                              .setBackgroundColor('#FFDD00')
                                                              .setIcon(CardService.Icon.STORE)
                                                              .setOpenLink(CardService.newOpenLink()
                                                                                     .setUrl('https://buymeacoffee.com/pfaitl'))
                                                 )
                                    );

  const legalSection = CardService.newCardSection()
                                  .addWidget(
                                    CardService.newTextParagraph()
                                               .setText('<font color="#666666"><i>Licensed under MIT License<br>Built with Google Apps Script<br>Contact: recorder@lvwasc.co.nz</i></font>')
                                  );

  aboutCard
    .addSection(descriptionSection)
    .addSection(featuresSection)
    .addSection(linksSection)
    .addSection(supportSection)
    .addSection(legalSection);

  return CardService.newNavigation()
                    .pushCard(aboutCard.build());
}

/**
 * Builds and returns the Help card with essential user guide information
 */
function showHelpCard() {
  const helpCard = CardService.newCardBuilder()
                              .setHeader(
                                CardService.newCardHeader()
                                           .setTitle('Help & Quick Start')
                                           .setSubtitle('Essential guide to using Swim Entries Manager')
                              );

  const quickStartSection = CardService.newCardSection()
                                       .setHeader('Quick Start Workflow')
                                       .addWidget(
                                         CardService.newTextParagraph()
                                                    .setText('<b>1. Create Core Tables (STEP 1)</b><br>Create Team Officials, Events for Template, and Schools tables. Define Gender and School Year values if using Individual Entries Template.')
                                       )
                                       .addWidget(
                                         CardService.newTextParagraph()
                                                    .setText('<b>2. Populate Tables</b><br>• <i>Team Officials:</i> Enter teams/schools/clusters<br>• <i>Events:</i> Enter Distance and Discipline<br>• <i>Schools:</i> Fill Team Name, School, Cluster, and Code')
                                       )
                                       .addWidget(
                                         CardService.newTextParagraph()
                                                    .setText('<b>3. Create Entry Sheets (STEP 2)</b><br>Duplicate the template for each team/school so coordinators work in separate sheets.')
                                       )
                                       .addWidget(
                                         CardService.newTextParagraph()
                                                    .setText('<b>4. Fill Swimmer Entries</b><br><b>Required:</b> First Name, Last Name, Gender, Date of Birth, valid Team Code<br><b>Recommended:</b> Entry times and School Year')
                                       )
                                       .addWidget(
                                         CardService.newTextParagraph()
                                                    .setText('<b>5. Export for Meet Manager (STEP 4)</b><br>Generate SDIF file and review exception reports if warnings appear. Alternative: Export to CSV.')
                                       );

  const criticalFieldsSection = CardService.newCardSection()
                                           .setHeader('Critical Data Requirements')
                                           .addWidget(
                                             CardService.newTextParagraph()
                                                        .setText('<b>Must be present or swimmer is excluded:</b><br>• First Name<br>• Last Name<br>• Date of Birth (DD/MM/YYYY)<br>• Gender<br>• Team Code (must exist in Schools table)<br>• Team Name (from Schools table)')
                                           )
                                           .addWidget(
                                             CardService.newTextParagraph()
                                                        .setText('<b>Event Format:</b><br>Events must be formatted as [Distance] [Stroke]<br>✓ Correct: "25m Freestyle", "50m Backstroke"<br>✗ Incorrect: "Y5", "25 Free"')
                                           );

  const exceptionReportsSection = CardService.newCardSection()
                                             .setHeader('Understanding Exception Reports')
                                             .addWidget(
                                               CardService.newTextParagraph()
                                                          .setText('When exporting for Meet Manager, the system validates all critical data. If issues are found, an <b>Exception Report</b> is generated that details:')
                                             )
                                             .addWidget(
                                               CardService.newTextParagraph()
                                                          .setText('• Missing critical data (swimmer excluded)<br>• Missing teams in Schools table<br>• Invalid event format (events skipped)<br>• Swimmers with no events (warning only)')
                                             )
                                             .addWidget(
                                               CardService.newTextParagraph()
                                                          .setText('<b>Always review exception reports before importing to Meet Manager!</b>')
                                             );

  const tipsSection = CardService.newCardSection()
                                 .setHeader('Best Practices')
                                 .addWidget(
                                   CardService.newTextParagraph()
                                              .setText('✓ Fill in all required fields before exporting<br>✓ Use dropdown values when available<br>✓ Add all teams to Schools table first<br>✓ Use correct event naming format<br>✓ Run test export early to catch data issues')
                                 );

  const resourcesSection = CardService.newCardSection()
                                      .setHeader('Need More Help?')
                                      .addWidget(
                                        CardService.newButtonSet()
                                                   .addButton(
                                                     CardService.newTextButton()
                                                                .setText('Full User Guide')
                                                                .setIcon(CardService.Icon.DESCRIPTION)
                                                                .setOpenLink(CardService.newOpenLink()
                                                                                       .setUrl('https://github.com/petrfaitl/swim-entries-manager/blob/main/docs/USER_GUIDE.md'))
                                                   )
                                      )
                                      .addWidget(
                                        CardService.newButtonSet()
                                                   .addButton(
                                                     CardService.newTextButton()
                                                                .setText('Report Issues')
                                                                .setIcon(CardService.Icon.EMAIL)
                                                                .setOpenLink(CardService.newOpenLink()
                                                                                       .setUrl('https://github.com/petrfaitl/swim-entries-manager/issues'))
                                                   )
                                      )
                                      .addWidget(
                                        CardService.newTextParagraph()
                                                   .setText('<font color="#666666"><i>Email: recorder@lvwasc.co.nz</i></font>')
                                      );

  helpCard
    .addSection(quickStartSection)
    .addSection(criticalFieldsSection)
    .addSection(exceptionReportsSection)
    .addSection(tipsSection)
    .addSection(resourcesSection);

  return CardService.newNavigation()
                    .pushCard(helpCard.build());
}

/**
 * Gets the names of all visible sheets in the active spreadsheet.
 * @returns {string[]} Array of sheet names.
 */
function getVisibleSheetNames() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .filter(sheet => !sheet.isSheetHidden())
    .map(sheet => sheet.getName());
}
