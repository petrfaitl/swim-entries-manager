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
      body { font-family: 'Google Sans',Roboto,sans-serif; padding: 20px; line-height: 1.4; color: #333; }
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
      <div class="small">Applicable to <strong>Individual Events Template</strong> table</div>
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
  @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600&display=swap');

  *, *::before, *::after { box-sizing: border-box; }

  body {
    font-family: 'DM Sans', 'Google Sans', sans-serif;
    padding: 18px 16px 24px;
    line-height: 1.5;
    color: #1a1a2e;
    background: #f7f8fc;
    margin: 0;
  }

  /* Header */
  .header { margin-bottom: 16px; }
  .header h2 {
    font-size: 15px;
    font-weight: 600;
    color: #1a1a2e;
    margin: 0 0 3px;
    letter-spacing: -0.2px;
  }
  .header p {
    font-size: 11.5px;
    color: #7b8299;
    margin: 0;
  }

  /* Section cards */
  .section {
    background: #fff;
    border: 1px solid #e8eaf0;
    border-radius: 10px;
    padding: 14px;
    margin-bottom: 10px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04);
  }

  .section-label {
    font-size: 10px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.7px;
    color: #0397B3;
    margin: 0 0 10px;
  }

  /* Fields */
  .field { margin-bottom: 10px; }
  .field:last-child { margin-bottom: 0; }

  label {
    display: block;
    font-size: 11.5px;
    font-weight: 500;
    color: #4a5068;
    margin-bottom: 4px;
  }

  label .hint {
    font-weight: 400;
    color: #a0a8bf;
    font-size: 10.5px;
  }

  select, input[type="text"], input[type="number"] {
    width: 100%;
    padding: 7px 10px;
    border: 1.5px solid #e0e3ed;
    border-radius: 7px;
    font-family: inherit;
    font-size: 12.5px;
    color: #1a1a2e;
    background: #fafbfd;
    outline: none;
    transition: border-color 0.15s, box-shadow 0.15s;
    appearance: none;
    -webkit-appearance: none;
  }

  select {
    background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6' viewBox='0 0 10 6'%3E%3Cpath d='M1 1l4 4 4-4' stroke='%237b8299' stroke-width='1.5' fill='none' stroke-linecap='round'/%3E%3C/svg%3E");
    background-repeat: no-repeat;
    background-position: right 10px center;
    padding-right: 28px;
    cursor: pointer;
  }

  select:focus, input[type="text"]:focus, input[type="number"]:focus {
    border-color: #0397B3;
    box-shadow: 0 0 0 3px rgba(3, 151, 179, 0.1);
    background: #fff;
  }

  .row { display: flex; gap: 8px; }
  .row > .field { flex: 1; margin-bottom: 0; }

  /* Divider */
  .divider {
    height: 1px;
    background: #eef0f6;
    margin: 10px 0;
  }

  /* Event name inputs — dynamically added */
  #eventNamesList {
    display: flex;
    flex-direction: column;
    gap: 6px;
    margin-top: 8px;
  }

  .event-name-input {
    width: 100%;
    padding: 7px 10px;
    border: 1.5px solid #e0e3ed;
    border-radius: 7px;
    font-family: inherit;
    font-size: 12.5px;
    color: #1a1a2e;
    background: #fafbfd;
    outline: none;
    transition: border-color 0.15s, box-shadow 0.15s;
    appearance: none;
    -webkit-appearance: none;
  }

  .event-name-input:focus {
    border-color: #0397B3;
    box-shadow: 0 0 0 3px rgba(3, 151, 179, 0.1);
    background: #fff;
  }

  /* Radio options */
  .radio-group { display: flex; flex-direction: column; gap: 6px; }

  .radio-option {
    display: flex;
    align-items: center;
    gap: 9px;
    padding: 8px 10px;
    border: 1.5px solid #e0e3ed;
    border-radius: 7px;
    cursor: pointer;
    transition: border-color 0.15s, background 0.15s;
    font-size: 12.5px;
    font-weight: 500;
    color: #4a5068;
    user-select: none;
    margin: 0;
  }

  .radio-option:hover { background: #f4f6fc; }

  .radio-option.selected {
    border-color: #0397B3;
    background: #f0fafc;
    color: #1a1a2e;
  }

  .radio-option input[type="radio"] {
    width: 14px;
    height: 14px;
    accent-color: #0397B3;
    margin: 0;
    flex-shrink: 0;
    appearance: auto;
    -webkit-appearance: auto;
  }

  #sameSheetOptions { margin-top: 10px; }
  #sameSheetOptions.hidden { display: none; }

  /* Button */
  #createRelayTablesBtn {
    width: 100%;
    background: linear-gradient(135deg, #0397B3 0%, #027a92 100%);
    color: white;
    border: none;
    cursor: pointer;
    border-radius: 9px;
    padding: 10px 16px;
    font-family: inherit;
    font-size: 13px;
    font-weight: 600;
    letter-spacing: 0.1px;
    margin-top: 4px;
    transition: box-shadow 0.15s, transform 0.1s;
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 7px;
  }

  #createRelayTablesBtn:hover {
    box-shadow: 0 4px 12px rgba(3, 151, 179, 0.35);
    transform: translateY(-1px);
  }

  #createRelayTablesBtn:active { transform: translateY(0); }
  #createRelayTablesBtn:disabled {
    opacity: 0.6;
    cursor: not-allowed;
    transform: none;
    box-shadow: none;
  }

  #createRelayTablesBtn svg { width: 15px; height: 15px; flex-shrink: 0; }

  /* Status */
  #status {
    margin-top: 10px;
    white-space: pre-wrap;
    font-size: 11.5px;
    color: #4a5068;
    padding: 0 2px;
  }

  #status.error { color: #c0392b; }
  #status.success { color: #0397B3; }
</style>
</head>
<body>

<div class="header">
  <h2>Create Relay Tables</h2>
  <p>Generate tables from team names. Place on one sheet or separate sheets.</p>
</div>

<!-- Team Data -->
<div class="section">
  <div class="section-label">Team Data</div>
  <div class="field">
    <label>Sheet with team names</label>
    <select id="sourceSheet"></select>
  </div>
  <div class="field">
    <label>Range with team names <span class="hint">e.g. A2:A8</span></label>
    <input type="text" id="sourceRange" value="A2:A8" placeholder="A2:A8" />
  </div>
</div>

<!-- Relay Events -->
<div class="section">
  <div class="section-label">Relay Events</div>
  <div class="field">
    <label>Number of relay events</label>
    <input type="number" id="eventCount" value="3" min="1" />
  </div>
  <div id="eventNamesList"></div>
  <div class="divider"></div>
  <div class="field">
    <label>School years <span class="hint">comma-separated</span></label>
    <input type="text" id="years" value="Y5,Y6,Y7,Y8,Y9,Junior,Intermediate,Senior" />
  </div>
</div>

<!-- Placement -->
<div class="section">
  <div class="section-label">Placement</div>
  <div class="radio-group">
    <label class="radio-option selected" id="opt-sameSheet">
      <input type="radio" name="placement" value="sameSheet" checked />
      Place all tables on one sheet
    </label>
    <label class="radio-option" id="opt-separateSheets">
      <input type="radio" name="placement" value="separateSheets" />
      Create separate sheet per event
    </label>
  </div>

  <div id="sameSheetOptions">
    <div class="divider"></div>
    <div class="row">
      <div class="field">
        <label>Target sheet name</label>
        <input type="text" id="targetSheetName" value="Relays" />
      </div>
      <div class="field">
        <label>Start cell <span class="hint">e.g. A1</span></label>
        <input type="text" id="startCell" value="A1" />
      </div>
    </div>
  </div>
</div>

<button id="createRelayTablesBtn" type="button">
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
    <rect x="3" y="3" width="18" height="18" rx="2"/>
    <line x1="3" y1="9" x2="21" y2="9"/>
    <line x1="3" y1="15" x2="21" y2="15"/>
    <line x1="9" y1="9" x2="9" y2="21"/>
  </svg>
  Create Relay Tables
</button>

<div id="status"></div>

<script>
  let eventNames = [];
  let relayDefaults = null;

  function setStatus(text, type) {
    const el = document.getElementById('status');
    el.textContent = text;
    el.className = type || '';
  }

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
    const isSameSheet = document.querySelector('input[name="placement"]:checked').value === 'sameSheet';
    const sameSheetOptions = document.getElementById('sameSheetOptions');
    sameSheetOptions.classList.toggle('hidden', !isSameSheet);
  }

  function createRelayTables() {
    const sourceSheet = document.getElementById('sourceSheet').value.trim();
    const sourceRange = document.getElementById('sourceRange').value.trim();

    if (!sourceSheet) { alert('Please select a source sheet.'); return; }
    if (!sourceRange) { alert('Please enter a range with team names.'); return; }

    const events = [];
    document.querySelectorAll('.event-name-input').forEach(function(input) {
      const label = input.value.trim();
      if (label) events.push({ label: label });
    });

    if (events.length === 0) { alert('Please enter at least one event name.'); return; }

    const yearsStr = document.getElementById('years').value.trim();
    const years = yearsStr
      ? yearsStr.split(',').map(function(y) { return y.trim(); }).filter(function(y) { return y !== ''; })
      : [];

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

    const btn = document.getElementById('createRelayTablesBtn');
    btn.disabled = true;
    setStatus('Creating relay tables…');

    google.script.run
      .withSuccessHandler(function(results) {
        btn.disabled = false;
        const lines = results.map(function(r) {
          return r.error
            ? r.eventLabel + ': ERROR — ' + r.error
            : 'Created: ' + r.eventLabel + ' on "' + r.sheetName + '"';
        }).join('\\n');

        const successCount = results.filter(function(r) { return !r.error; }).length;
        const summary = sameSheet
          ? 'Created ' + successCount + ' relay table(s) on sheet "' + config.placement.sheetName + '".'
          : 'Created ' + successCount + ' relay table(s) on ' + successCount + ' separate sheet(s).';

        setStatus(summary + '\\n\\n' + lines, 'success');
      })
      .withFailureHandler(function(err) {
        btn.disabled = false;
        setStatus('Error: ' + err.message, 'error');
      })
      .createRelayTablesFromDialog(config);
  }

  function bindEvents() {
    const btn = document.getElementById('createRelayTablesBtn');
    if (!btn) { setStatus('Error: Button not found.', 'error'); return; }

    btn.addEventListener('click', createRelayTables);
    document.getElementById('eventCount').addEventListener('input', renderEventNameInputs);

    document.querySelectorAll('input[name="placement"]').forEach(function(radio) {
      radio.addEventListener('change', function() {
        document.querySelectorAll('.radio-option').forEach(function(el) {
          el.classList.remove('selected');
        });
        radio.closest('.radio-option').classList.add('selected');
        updatePlacementOptions();
      });
    });

    loadSourceSheets();
    loadRelayDefaults();
    updatePlacementOptions();
    setStatus('Ready.');
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
