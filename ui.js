/**
 * Add menu item to Custom Tools
 */
function onOpen(e) {
  const menu = SpreadsheetApp.getUi().createAddonMenu();
  if (e && e.authMode === ScriptApp.AuthMode.NONE) {
    Logger.log("Not enough permissions to create menu yet.");
    menu.addItem('Start Add-on', 'showSidebar');
  }else{
    menu.addItem('Create Sheets from Template', 'showCreateSheetsDialog')
      .addSeparator()
      .addItem('Export to CSV', 'exportEntriesToCSV')
  }
  menu.addToUi();

  
}

function onHomepage(e) {
  onOpen(e);
  return buildAdminSidebarCard();
  
  
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
    .setCollapsible(true) // Optional: collapses if not needed
    .addWidget(
      CardService.newTextParagraph()
        .setText('Duplicate a template sheet for each school/cluster from a list in your spreadsheet.')
    )
    .addWidget(
      CardService.newButtonSet()
        .addButton(
          CardService.newTextButton()
            .setText('Create Sheets from List')
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setBackgroundColor('#4285F4') // Google blue
            .setOnClickAction(CardService.newAction().setFunctionName('showCreateSheetsDialog'))
        )
    )
    .addWidget(CardService.newDivider()); // Horizontal separator

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