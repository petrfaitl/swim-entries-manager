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


/**
 * Shows a dialog to let user select template name and range
 */
function showCreateSheetsDialog() {
  const ui = SpreadsheetApp.getUi();

  const html = HtmlService.createHtmlOutput(`
    <style>
      body {font-family: Arial; padding: 15px;}
      label {display: block; margin: 10px 0 5px;}
      input, select {width: 100%; padding: 6px;}
      button {margin-top: 15px; padding: 8px 16px;}
    </style>
    
    <label>Template Sheet Name:</label>
    <input type="text" id="templateName" value="INDIVIDUAL_EVENTS_TEMPLATE" placeholder="e.g. INDIVIDUAL_EVENTS_TEMPLATE">

    <label>Sheet name with names of Teams (e.g. Team Officials):</label>
    <input type="text" id="sheet" value="Team Officials" placeholder="Team Officials">
    
    <label>Range with names (e.g. A2:A8):</label>
    <input type="text" id="range" value="A2:A8" placeholder="A2:A8">
    
    <button onclick="run()">Create Sheets</button>
    
    <script>
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
    </script>
  `)
  .setWidth(380)
  .setHeight(280);

  ui.showModalDialog(html, 'Create Sheets from List');
}
