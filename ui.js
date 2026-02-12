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
  const tableNames = listAvailableTables();
  Logger.log('[showCreateTablesDialog] tableNames payload: %s', JSON.stringify(tableNames));
  const tableListHtml = tableNames.length
    ? tableNames.map(function(name) {
      return '<label><input type="checkbox" value="' + name + '"> ' + name + '</label>';
    }).join('')
    : 'No tables available.';
  const html = HtmlService.createHtmlOutput(
    `
    <style>
      body {font-family: Arial, sans-serif; padding: 14px;}
      h2 {margin: 0 0 8px;}
      fieldset {border: 1px solid #ddd; padding: 10px; margin-bottom: 12px;}
      legend {font-weight: bold;}
      label {display:block; margin: 4px 0;}
      .row {display:flex; gap:8px;}
      .row > div {flex:1;}
      button {margin-top: 10px; padding: 8px 14px;}
      .small {font-size: 12px; color: #666}
      #status {margin-top:8px; white-space: pre-wrap;}
    </style>
    <h2>Create Core Tables</h2>
    <div class="small">Select one or more table types to create or rebuild. You can override dropdown values (e.g., school years) below.</div>

    <fieldset>
      <legend>Tables</legend>
      <div id="tableList">${tableListHtml}</div>
      <div id="tableListDebug" class="small">Rendered server-side: ${tableNames.length} table option(s).</div>
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

    <button onclick="create()">Create Tables</button>
    <div id="status"></div>

    <script>
      function create(){
        const names = Array.from(document.querySelectorAll('#tableList input[type=checkbox]:checked')).map(cb=>cb.value);
        if (!names.length) { alert('Select at least one table.'); return; }
        const years = document.getElementById('years').value.split(',').map(s=>s.trim()).filter(Boolean);
        const genders = document.getElementById('genders').value.split(',').map(s=>s.trim()).filter(Boolean);
        const sheetName = document.getElementById('sheetName').value.trim();
        const startCell = document.getElementById('startCell').value.trim();
        const opts = { schoolYears: years, genders: genders };
        if (sheetName) opts.sheetName = sheetName;
        if (startCell) opts.startCell = startCell;
        document.getElementById('status').textContent = 'Working…';
        google.script.run
          .withSuccessHandler(function(res){
            const lines = res.map(function(r){
              if (r.error) {
                return r.tableName + ': ERROR - ' + r.error;
              }
              const tableMeta = r.structuredTable
                ? (' [' + r.structuredTable.action + ' structured: ' + r.structuredTable.tableName + ']')
                : '';
              return r.tableName + ': ' + r.sheetName + ' ' + r.headerA1 + ' / ' + r.dataA1 + tableMeta;
            }).join('\n');
            document.getElementById('status').textContent = 'Done.\n' + lines;
          })
          .withFailureHandler(function(err){
            document.getElementById('status').textContent = 'Error: ' + err.message;
          })
          .createTablesFromDialog(names, opts);
      }
    </script>
    `
  ).setWidth(520).setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, 'Create Core Tables');
}

function buildAdminSidebarCard() {
  const cardBuilder = CardService.newCardBuilder()
    .setHeader(
      CardService.newCardHeader()
        .setTitle('Swim Entries Admin Tools')
        .setSubtitle('Register swimmers to swim events, export entries for swim meets')
        .setImageUrl('https://www.gstatic.com/images/branding/product/1x/admin_48dp.png') // optional icon
        .setImageStyle(CardService.ImageStyle.CIRCLE)
    );

  // ── Section 1: Create Sheets from Template ─────────────────────────────
  const createSection = CardService.newCardSection()
    .setHeader('Create School Entry Sheets')
    .setCollapsible(true)
    .addWidget(
      CardService.newTextParagraph()
        .setText('Duplicate a template sheet for each school/cluster from a list in your spreadsheet, or generate core tables.')
    )
    .addWidget(
      CardService.newButtonSet()
        .addButton(
          CardService.newTextButton()
            .setText('Create Sheets from List')
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setBackgroundColor('#34A853')
            .setOnClickAction(CardService.newAction().setFunctionName('showCreateSheetsDialog'))
        )
        .addButton(
          CardService.newTextButton()
            .setText('Create Core Tables')
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setBackgroundColor('#c23433')
            .setOnClickAction(CardService.newAction().setFunctionName('showCreateTablesDialog'))
        )
    )
    .addWidget(CardService.newDivider());


  // ── Section 2: Export Entries ──────────────────────────────────────────
  const exportSection = CardService.newCardSection()
    .setHeader('Export Entries to CSV')
    .setCollapsible(true)
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
            .setBackgroundColor('#34A853') // Google green
            .setOnClickAction(CardService.newAction().setFunctionName('exportEntriesToCSV'))
        )
    );

  // ── Assemble the card ──────────────────────────────────────────────────
  cardBuilder
    .addSection(createSection)
    .addSection(exportSection);

  // Optional: Add a footer with version or help link
  cardBuilder.setFixedFooter(
    CardService.newFixedFooter()
      .setSecondaryButton(
        CardService.newTextButton()
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setBackgroundColor('#FFDD00')
        .setText('Buy me a coffee.')
        .setIcon(CardService.Icon.STORE)
        .setOpenLink(
        CardService.newOpenLink().setUrl('https://buymeacoffee.com/pfaitl')),
      )
      .setPrimaryButton(
        CardService.newTextButton().setText('CSV to SDIF Converter').setOpenLink(
        CardService.newOpenLink().setUrl('https://csv-sdif-converter.netlify.app/')),
      )
  );

  return cardBuilder.build();
}
