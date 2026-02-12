/**
 * Add menu item to Custom Tools
 */
function onOpen(e) {
  const menu = SpreadsheetApp.getUi().createAddonMenu();
  if (e && e.authMode === ScriptApp.AuthMode.NONE) {
    Logger.log("Not enough permissions to create menu yet.");
    menu.addItem('Start Add-on', 'showSidebar');
  } else {
    menu.addItem('Create Sheets from Template', 'showCreateSheetsDialog')
      .addItem('Create Core Tables', 'showCreateTablesDialog')
      .addSeparator()
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
    ? tables.map(function(table) {
      return '<label><input type="checkbox" value="' + table.name + '"> ' + table.title + '</label>';
    }).join('')
    : 'No tables available.';
  const html = HtmlService.createHtmlOutput(
    `
    <style>
      body {font-family: Arial, sans-serif; padding: 14px;}
      h2 {margin: 0 0 8px;}
      fieldset {border: 1px solid #ddd; padding: 10px; margin-bottom: 12px;}
      legend {font-weight: bold;}
      label {display:block; margin: 4px 0; font-size:smaller;}
      .row {display:flex; gap:8px;}
      .row > div {flex:1;}
      select, button { padding: 8px; margin: 10px 0; width: 100%; font-size: 16px; }
      button { background: #4285f4; color: white; border: none; cursor: pointer; }
      button:hover { background: #3267d6; }
      .small {font-size: 12px; color: #666}
      #status {margin-top:8px; white-space: pre-wrap; font-size: smaller;}
    </style>
    <h2>Create Core Tables</h2>
    <div class="small">Select one or more table types to create or rebuild. You can override dropdown values (e.g., school years) below.</div>

    <fieldset>
      <legend>Tables</legend>
      <div id="tableList">${tableListHtml}</div>
      <div id="tableListDebug" class="small">Rendered server-side: ${tables.length} table option(s).</div>
    </fieldset>

    <fieldset>
      <legend>Dropdown values (optional)</legend>
      <div class="row">
        <div>
          <label>School Years (comma-separated)</label>
          <input type="text" id="years" value="Y5,Y6,Y7,Y8,Y9,Y10,Y11,Y12,Y13" />
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
  ).setWidth(520).setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, 'Create Core Tables');
}

function buildAdminSidebarCard() {
  const imageBytes = DriveApp.getFileById('1PT-hMqaBpVTNpkN3KiocfgNa7QRWfmXw').getBlob().getBytes();
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
        .setStartIcon(CardService.newIconImage().setMaterialIcon(
          CardService.newMaterialIcon().setName('looks_one')
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
                      .setBackgroundColor('#c23433')
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
        .setStartIcon(CardService.newIconImage().setMaterialIcon(
          CardService.newMaterialIcon().setName('looks_two')
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
            .setBackgroundColor('#34A853')
            .setOnClickAction(CardService.newAction().setFunctionName('showCreateSheetsDialog'))
        )
    );

  // Navigation to Export Card
  const navToExportSection = CardService.newCardSection()

    .addWidget(
      CardService.newTextParagraph()
        .setText('<b>Next:</b> Export entries to CSV')
    )
    .addWidget(
      CardService.newButtonSet()
        .addButton(
          CardService.newTextButton()
            .setText('Go to Export →')
            .setTextButtonStyle(CardService.TextButtonStyle.TEXT)
            .setOnClickAction(CardService.newAction().setFunctionName('buildExportCard'))
        )
    );

  templatesCard
    .addSection(createTablesSection)
    .addSection(createSheetsSection)
    .addSection(navToExportSection)
    .setFixedFooter(
      CardService.newFixedFooter()
        .setPrimaryButton(
          CardService.newTextButton()
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setBackgroundColor('#FFDD00')
            .setText('Buy me a coffee.')
            .setIcon(CardService.Icon.STORE)
            .setOpenLink(CardService.newOpenLink().setUrl('https://buymeacoffee.com/pfaitl'))
        )
    );

  return templatesCard.build();
}

function buildExportCard() {
  const imageBytes = DriveApp.getFileById('1DQZoQh-ya7I-rVsED3HJjVdobkEFwlvZ').getBlob().getBytes();
  const encodedImageURL = `data:image/jpeg;base64,${Utilities.base64Encode(imageBytes)}`;

  // ── Card 2: Export ─────────────────────────────────────────────────────
  const exportCard = CardService.newCardBuilder()
    .setHeader(
      CardService.newCardHeader()
        .setTitle('Export Entries')
        .setSubtitle('Generate CSV files for meet management software.')
        .setImageUrl(encodedImageURL)
    );

  const exportSection = CardService.newCardSection()
    .setHeader('Export Entries to CSV')
    .setCollapsible(true)
    .addWidget(
      CardService.newDecoratedText()
        .setText('<b>STEP 3</b>')
        .setStartIcon(CardService.newIconImage().setMaterialIcon(
          CardService.newMaterialIcon().setName('looks_3')
        ))
    )
    .addWidget(
      CardService.newTextParagraph()
        .setText('Process the current sheet and generate a clean CSV file with formatted entries and times.')
    )
    .addWidget(
      CardService.newButtonSet()
        .addButton(
          CardService.newTextButton()
            .setText('Export to CSV')
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setBackgroundColor('#34A853')
            .setOnClickAction(CardService.newAction().setFunctionName('exportEntriesToCSV'))
        )
    );

  const csvConverterSection = CardService.newCardSection()
   .setHeader('Convert CSV to SDIF')
   .setCollapsible(true)
   .addWidget(
     CardService.newDecoratedText()
                .setText('<b>STEP 4</b>')
                .setStartIcon(CardService.newIconImage().setMaterialIcon(
                  CardService.newMaterialIcon().setName('looks_4')
                ))
   )
   .addWidget(
     CardService.newTextParagraph()
                .setText('After exporting your CSV, go to the CSV to SDIF Converter to prepare entries for the Hy-tek Meet Manager.')
   )

   .addWidget(
     CardService.newButtonSet()
                .addButton(
                  CardService.newTextButton()
                             .setText('Convert to SDIF')
                             .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
                             .setBackgroundColor('#34A853')
                             .setOpenLink(CardService.newOpenLink().setUrl('https://csv-sdif-converter.netlify.app/'))
                )
   )
   .addWidget(
     CardService.newTextParagraph()
                .setText('(Opens in a new window.)')
   );

  // Navigation back to Templates
  const navBackSection = CardService.newCardSection()

    .addWidget(
      CardService.newButtonSet()
        .addButton(
          CardService.newTextButton()
            .setText('← Back to Templates')
            .setTextButtonStyle(CardService.TextButtonStyle.TEXT)
            .setOnClickAction(CardService.newAction().setFunctionName('buildAdminSidebarCard'))
        )
    );

  exportCard
    .addSection(exportSection)
    .addSection(csvConverterSection)
    .addSection(navBackSection)
    .setFixedFooter(
      CardService.newFixedFooter()
        .setPrimaryButton(
          CardService.newTextButton()
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setBackgroundColor('#FFDD00')
            .setText('Buy me a coffee.')
            .setIcon(CardService.Icon.STORE)
            .setOpenLink(CardService.newOpenLink().setUrl('https://buymeacoffee.com/pfaitl'))
        )

    );

  return CardService.newNavigation().pushCard(exportCard.build());
}
