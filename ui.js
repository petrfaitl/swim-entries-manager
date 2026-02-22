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
  const tableListHtml = tables.length
    ? tables.map(function (table) {
              return '<label><input type="checkbox" value="' + table.name + '"> ' + table.title + '</label>';
            })
            .join('')
    : 'No tables available.';
  const html = HtmlService.createHtmlOutput(
                            `
    <style>
      body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 20px; line-height: 1.4; color: #333; }
      h2 {margin: 0 0 8px;}
      fieldset {border: 1px solid #ddd; padding: 10px; margin-bottom: 12px;}
      legend {font-weight: bold;}
      label {display:block; margin: 4px 0; font-size:smaller;}
      .row {display:flex; gap:8px;}
      .row > div {flex:1;}
      select, button { padding: 8px; margin: 10px 0; font-size: 16px; }
      input { padding: 8px; margin: 10px 0; }
      button { background: #0397B3; color: white; border: none; cursor: pointer; border-end-end-radius: 14px;
              border-start-end-radius: 14px;
              border-start-start-radius: 14px;
              border-end-start-radius: 14px;
              padding-inline-start: 24px;
              padding-inline-end: 24px;
              min-inline-size: 250px;
              }
      button:hover { box-shadow: 0 2px 4px 0 rgba(0, 0, 0, 0.4); }
      .small {font-size: 12px; color: #666}
      .subtitle {margin-bottom: 14px;}
      .required {color: #ff3c19;}
      #status {margin-top:8px; white-space: pre-wrap; font-size: smaller;}
    </style>

    <div class="small">Select one or more table types to create or rebuild. You can override dropdown values (e.g., school years) below.</div>
    <div class="small subtitle">Required tables marked with an asterisk (<span class="required">*</span>)</div>

    <fieldset>
      <legend>Tables</legend>
      <div id="tableList">` + tableListHtml + `</div>
      <div id="tableListDebug" class="small">Rendered server-side: ` + tables.length + ` table option(s).</div>
    </fieldset>

    <fieldset>
      <legend>Dropdown values (optional)</legend>
      <div class="row">
        <div>
          <label>School Years (comma-separated)</label>
          <input type="text" id="years" value="Y5,Y6,Y7,Y8,Y9,Junior,Intermediate,Senior" />
        </div>
        <div>
          <label>Genders (comma-separated)</label>
          <input type="text" id="genders" value="Female,Male" />
        </div>
      </div>
    </fieldset>

    <fieldset>
      <legend>Placement (optional overrides for all selected)</legend>
      <div class="row">
        <div>
          <label>Override Sheet Name</label>
          <input type="text" id="sheetName" placeholder="Leave blank to use default" />
        </div>
        <div>
          <label>Start Cell (A1)</label>
          <input type="text" id="startCell" placeholder="A1" />
        </div>
      </div>
    </fieldset>

    <button id="createTablesBtn" type="button">Create Tables</button>
    <div id="status"></div>

    <script>
      function createTables() {
        const names = Array.from(document.querySelectorAll('#tableList input[type=checkbox]:checked')).map(function(cb) {
          return cb.value;
        });

        if (!names.length) {
          alert('Please select at least one table.');
          return;
        }

        const years = document.getElementById('years').value.split(',').map(function(s) {
          return s.trim();
        }).filter(Boolean);

        const genders = document.getElementById('genders').value.split(',').map(function(s) {
          return s.trim();
        }).filter(Boolean);

        const sheetName = document.getElementById('sheetName').value.trim();
        const startCell = document.getElementById('startCell').value.trim();

        const opts = {
          schoolYears: years,
          genders: genders
        };

        if (sheetName) opts.sheetName = sheetName;
        if (startCell) opts.startCell = startCell;

        document.getElementById('status').textContent = 'Working…';

        google.script.run
          .withSuccessHandler(function(res) {
            const lines = res.map(function(r) {
              if (r.error) {
                return r.tableName + ': ERROR - ' + r.error;
              }
              const tableMeta = r.structuredTable
                ? (' [' + r.structuredTable.action + ' table name: ' + r.structuredTable.tableName + ']')
                : '';
              return 'Created table at sheet: ' + r.sheetName + ', cell reference: ' + r.headerA1 + ' / ' + r.dataA1 + tableMeta;
            }).join('\\n');
            document.getElementById('status').textContent = 'Done.\\n' + lines;
          })
          .withFailureHandler(function(err) {
            document.getElementById('status').textContent = 'Error: ' + err.message;
          })
          .createTablesFromDialog(names, opts);
      }

      function bindEvents() {
        const btn = document.getElementById('createTablesBtn');
        if (!btn) {
          document.getElementById('status').textContent = 'Error: Create button not found.';
          return;
        }
        btn.addEventListener('click', createTables);
        document.getElementById('status').textContent = 'Ready.';
      }

      if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', bindEvents);
      } else {
        bindEvents();
      }
    </script>
    `
                          )
                          .setWidth(520)
                          .setHeight(600);

  SpreadsheetApp.getUi()
                .showModalDialog(html, 'Create Core Tables');
}

function showCreateRelayTablesDialog() {
  const html = HtmlService.createHtmlOutput(`
    
    <style>
      body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 20px; line-height: 1.4; color: #333; }
      h2 {margin: 0 0 8px;}
      fieldset {border: 1px solid #ddd; padding: 10px; margin-bottom: 12px;}
      legend {font-weight: bold;}
      label {display:block; margin: 4px 0; font-size:smaller;}
      .row {display:flex; gap:8px;}
      .row > div {flex:1;}
      select {padding: 8px; margin: 10px 0; width:100%;}
      input {padding: 8px; margin: 10px 0;width: calc(100% - 20px);}
      input[type="radio"] {margin:10px; width: auto;}
      button { background: #0397B3; color: white; border: none; cursor: pointer;
               border-radius: 14px;
               padding-inline-start: 24px;
               padding-inline-end: 24px;
               min-inline-size: 250px;
               padding: 8px; margin: 10px 0; font-size: 16px;
             }
      button:hover { box-shadow: 0 2px 4px 0 rgba(0, 0, 0, 0.4); }
      .small {font-size: 12px; color: #666}
      .subtitle {margin-bottom: 14px;}
      #status {margin-top:8px; white-space: pre-wrap; font-size: smaller;}
      #eventNamesList {margin-top: 10px;}
      .event-name-input {margin: 5px 0;}
      .hidden {display: none;}
    </style>

    <div class="small">Create relay tables from team names in your spreadsheet.</div>
    <div class="small subtitle">Tables can be placed on one sheet or separate sheets.</div>

    <fieldset>
      <legend>Team Data</legend>
      <div>
        <label>Sheet with team names</label>
        <select id="sourceSheet"></select>
      </div>
      <div>
        <label>Range with team names <span class="small">(e.g. A2:A8)</span></label>
        <input type="text" id="sourceRange" value="A2:A8" placeholder="A2:A8" />
      </div>
    </fieldset>

    <fieldset>
      <legend>Relay Events</legend>
      <div>
        <label>Number of relay events</label>
        <input type="number" id="eventCount" value="3" min="1" />
      </div>
      <div id="eventNamesList"></div>
      <div>
        <label>School Years <span class="small">(comma-separated)</span></label>
        <input type="text" id="years" value="Y5,Y6,Y7,Y8,Y9,Junior,Intermediate,Senior" />
      </div>
    </fieldset>

    <fieldset>
      <legend>Placement</legend>
      <div>
        <label>
          <input type="radio" name="placement" value="sameSheet" checked /> Place all tables on one sheet
        </label>
        <label>
          <input type="radio" name="placement" value="separateSheets" /> Create separate sheet per event
        </label>
      </div>
      <div id="sameSheetOptions" style="margin-top: 10px;">
        <div>
          <label>Target sheet name</label>
          <input type="text" id="targetSheetName" value="Relays" />
        </div>
        <div>
          <label>Start cell <span class="small">(e.g A1)</span></label>
          <input type="text" id="startCell" value="A1" />
        </div>
      </div>
    </fieldset>

    <button id="createRelayTablesBtn" type="button">Create Relay Tables</button>
    <div id="status"></div>

    <script>
      let eventNames = [];
      let relayDefaults = null;

      function loadSourceSheets() {
        google.script.run
          .withSuccessHandler(function(names) {
            const select = document.getElementById('sourceSheet');
            select.innerHTML = '';
            if (!names || names.length === 0) {
              select.innerHTML = '<option value="">No sheets found</option>';
              return;
            }
            names.forEach(function(name) {
              const option = document.createElement('option');
              option.value = name;
              option.textContent = name;
              select.appendChild(option);
            });
          })
          .withFailureHandler(function(err) {
            alert('Error loading sheets: ' + err.message);
          })
          .getNonTemplateSheetNames();
      }

      function loadRelayDefaults() {
        google.script.run
          .withSuccessHandler(function(defaults) {
            relayDefaults = defaults;
            renderEventNameInputs();
          })
          .withFailureHandler(function(err) {
            console.error('Error loading relay defaults:', err.message);
            relayDefaults = { suggestedEventNames: [] };
            renderEventNameInputs();
          })
          .getRelayDefaults();
      }

      function renderEventNameInputs() {
        const count = parseInt(document.getElementById('eventCount').value) || 1;
        const container = document.getElementById('eventNamesList');
        container.innerHTML = '';

        const defaults = (relayDefaults && relayDefaults.suggestedEventNames) || [];

        for (let i = 0; i < count; i++) {
          const input = document.createElement('input');
          input.type = 'text';
          input.className = 'event-name-input';
          input.placeholder = 'Event ' + (i + 1) + ' name';
          input.value = eventNames[i] || (i < defaults.length ? defaults[i] : '');
          input.dataset.index = i;
          input.addEventListener('input', function() {
            eventNames[parseInt(this.dataset.index)] = this.value;
          });
          container.appendChild(input);
        }
      }

      function updatePlacementOptions() {
        const sameSheet = document.querySelector('input[name="placement"]:checked').value === 'sameSheet';
        const sameSheetOptions = document.getElementById('sameSheetOptions');
        sameSheetOptions.style.display = sameSheet ? 'block' : 'none';
      }

      function createRelayTables() {
        const sourceSheet = document.getElementById('sourceSheet').value.trim();
        const sourceRange = document.getElementById('sourceRange').value.trim();
        const eventCount = parseInt(document.getElementById('eventCount').value);

        if (!sourceSheet) {
          alert('Please select a source sheet.');
          return;
        }

        if (!sourceRange) {
          alert('Please enter a range with team names.');
          return;
        }

        // Collect event names
        const events = [];
        const inputs = document.querySelectorAll('.event-name-input');
        inputs.forEach(function(input) {
          const label = input.value.trim();
          if (label) {
            events.push({ label: label });
          }
        });

        if (events.length === 0) {
          alert('Please enter at least one event name.');
          return;
        }

        // Get school years
        const yearsStr = document.getElementById('years').value.trim();
        const years = yearsStr ? yearsStr.split(',').map(function(y) { return y.trim(); }).filter(function(y) { return y !== ''; }) : [];

        const sameSheet = document.querySelector('input[name="placement"]:checked').value === 'sameSheet';
        const config = {
          sourceSheet: sourceSheet,
          sourceRange: sourceRange,
          events: events,
          years: years,
          placement: {
            sameSheet: sameSheet,
            sheetName: sameSheet ? document.getElementById('targetSheetName').value.trim() : '',
            startCell: sameSheet ? document.getElementById('startCell').value.trim() : ''
          }
        };

        document.getElementById('status').textContent = 'Creating relay tables…';

        google.script.run
          .withSuccessHandler(function(results) {
            const lines = results.map(function(r) {
              if (r.error) {
                return r.eventLabel + ': ERROR - ' + r.error;
              }
              return 'Created: ' + r.eventLabel + ' on sheet "' + r.sheetName + '" (table: ' + r.sanitisedTableName + ')';
            }).join('\\n');

            const successCount = results.filter(function(r) { return !r.error; }).length;
            const message = sameSheet
              ? 'Created ' + successCount + ' relay table(s) on sheet "' + config.placement.sheetName + '".'
              : 'Created ' + successCount + ' relay table(s) on ' + successCount + ' separate sheet(s).';

            document.getElementById('status').textContent = message + '\\n\\n' + lines;
          })
          .withFailureHandler(function(err) {
            document.getElementById('status').textContent = 'Error: ' + err.message;
          })
          .createRelayTablesFromDialog(config);
      }

      function bindEvents() {
        const btn = document.getElementById('createRelayTablesBtn');
        if (!btn) {
          document.getElementById('status').textContent = 'Error: Button not found.';
          return;
        }
        btn.addEventListener('click', createRelayTables);

        document.getElementById('eventCount').addEventListener('input', renderEventNameInputs);
        document.querySelectorAll('input[name="placement"]').forEach(function(radio) {
          radio.addEventListener('change', updatePlacementOptions);
        });

        // Initialize
        loadSourceSheets();
        loadRelayDefaults();
        updatePlacementOptions();
        document.getElementById('status').textContent = 'Ready.';
      }

      if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', bindEvents);
      } else {
        bindEvents();
      }
    </script>
  `)
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
                              .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
                              .setBackgroundColor('#FFDD00')
                              .setText('Buy me a coffee.')
                              .setIcon(CardService.Icon.STORE)
                              .setOpenLink(CardService.newOpenLink()
                                                      .setUrl('https://buymeacoffee.com/pfaitl'))
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
                              .setBackgroundColor('#FFDD00')
                              .setText('Buy me a coffee.')
                              .setIcon(CardService.Icon.STORE)
                              .setOpenLink(CardService.newOpenLink()
                                                      .setUrl('https://buymeacoffee.com/pfaitl'))
                 )
    );

  return CardService.newNavigation()
                    .pushCard(exportCard.build());
}

function showSDIFCreatorDialog() {
  SDIFCreator.showDialog();
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
