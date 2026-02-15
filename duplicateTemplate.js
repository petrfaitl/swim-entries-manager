/**
 * Duplicates the template sheet for each name in the given range
 * and renames the copies to match the names.
 *
 * @param {string} templateSheetName - Name of the sheet to duplicate (e.g. "Template")
 * @param {Range} namesRange - Range containing the list of names (single column)
 * @customfunction (you can also call it manually or from menu)
 */
function duplicateTemplateForNames(templateSheetName, sheetName,namesRange) {
  const ss = SpreadsheetApp.getActive();
  const templateSheet = ss.getSheetByName(templateSheetName);

  if (!templateSheet) {
    throw new Error(`Template sheet "${templateSheetName}" not found!`);
  }

  // Get the list of names (single column assumed)
  const rangeValues = "";
  const names = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(namesRange).getValues()
    .flat()
    .map(name => String(name).trim()) // convert to string & trim
    .filter(name => name !== "");     // remove blanks

  if (names.length === 0) {
    SpreadsheetApp.getUi().alert("No valid names found in the range.");
    return;
  }

  // Optional: Prevent duplicates (skip if sheet with this name already exists)
  const existingSheets = new Set(ss.getSheets().map(s => s.getName()));

  let createdCount = 0;

  names.forEach(name => {
    // Skip if sheet already exists
    if (existingSheets.has(name)) {
      Logger.log(`Sheet "${name}" already exists → skipping`);
      return;
    }

    // Duplicate the template
    const newSheet = templateSheet.copyTo(ss);
    newSheet.setName(name);

    // Optional: Move the new sheet to the end (or to a specific position)
    ss.setActiveSheet(newSheet);
    ss.moveActiveSheet(ss.getSheets().length);

    createdCount++;
  });

  SpreadsheetApp.getUi().alert(
    `Success!\nCreated ${createdCount} new sheets from template "${templateSheetName}".`
  );
}

function getTemplateSheetNames() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheets()
    .map(sheet => sheet.getName())
    .filter(name => name.toUpperCase().includes("TEMPLATE"));
}

function getNonTemplateSheetNames() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheets()
    .map(sheet => sheet.getName())
    .filter(name => !name.toUpperCase().includes("TEMPLATE"));
}


/**
 * Shows a dialog to let user select template name and range
 */
function showCreateSheetsDialog() {
  const ui = SpreadsheetApp.getUi();

  const html = HtmlService.createHtmlOutput(`
    <style>
      body {font-family: Arial; padding: 15px;}
      label {display: block; margin: 10px 0 5px;}
      input, select {width: 100%; padding: 8px; margin: 10px 0;}
      button { padding: 8px; margin: 10px 0; font-size: 16px; }
      button {  background: #0397B3; color: white; border: none; cursor: pointer; border-end-end-radius: 14px;
              border-start-end-radius: 14px;
              border-start-start-radius: 14px;
              border-end-start-radius: 14px;
              padding-inline-start: 24px;
              padding-inline-end: 24px;
              min-inline-size: 200px;}
      button:hover { box-shadow: 0 2px 4px 0 rgba(0, 0, 0, 0.4);}
    </style>
    
    <label>Template Sheet Name:</label>
    <select id="templateName"></select>

    <label>Sheet name with names of Teams (e.g. Team Officials):</label>
    <select id="sheet"></select>
    
    <label>Range with names (e.g. A2:A8):</label>
    <input type="text" id="range" value="A2:A8" placeholder="A2:A8">
    
    <button onclick="run()">Create Sheets</button>
    
    <script>
      function loadTemplates() {
        google.script.run
          .withSuccessHandler(names => {
            const select = document.getElementById('templateName');
            select.innerHTML = '';

            if (!names || names.length === 0) {
              select.innerHTML = '<option value="">No template sheets found</option>';
              return;
            }

            names.forEach(name => {
              const option = document.createElement('option');
              option.value = name;
              option.textContent = name;
              select.appendChild(option);
            });
          })
          .withFailureHandler(err => alert('Error loading templates: ' + err.message))
          .getTemplateSheetNames();
      }

      function loadSourceSheets() {
        google.script.run
          .withSuccessHandler(names => {
            const select = document.getElementById('sheet');
            select.innerHTML = '';

            if (!names || names.length === 0) {
              select.innerHTML = '<option value="">No source sheets found</option>';
              return;
            }

            names.forEach(name => {
              const option = document.createElement('option');
              option.value = name;
              option.textContent = name;
              select.appendChild(option);
            });
          })
          .withFailureHandler(err => alert('Error loading source sheets: ' + err.message))
          .getNonTemplateSheetNames();
      }

      function run() {
        const template = document.getElementById('templateName').value.trim();
        const sheetStr = document.getElementById('sheet').value.trim();
        const rangeStr = document.getElementById('range').value.trim();
        
        if (!template || !rangeStr) {
          alert('Please fill all fields');
          return;
        }
        
        google.script.run
          .withSuccessHandler(() => google.script.host.close())
          .withFailureHandler(err => alert('Error: ' + err.message))
          .duplicateTemplateForNames(template, sheetStr, rangeStr);
      }

      loadTemplates();
      loadSourceSheets();
    </script>
  `)
  .setWidth(380)
  .setHeight(380);

  ui.showModalDialog(html, 'Create Sheets from List');
}
