function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('SheetAlerts')
    .addItem('Open Configuration Panel', 'showSidePanel')
    .addSeparator()
    .addItem('Help & Documentation', 'showHelp')
    .addToUi();
}

function showSidePanel() {
  // In GWAO, the homepageTrigger handles the UI. 
  // We can't programmatically open the GWAO side panel from a menu in all cases,
  // but we can show an alert or use legacy sidebar as a fallback.
  // Best practice: tell user to click the icon in the side panel.
  var ui = SpreadsheetApp.getUi();
  ui.alert('SheetAlerts is now a Workspace Add-on. Please click the SheetAlerts icon in the right-side panel to configure your alerts.');
}

function showHelp() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('SheetAlerts Help'))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph().setText('SheetAlerts monitors your sheets and sends notifications to Slack and Gmail.'))
      .addWidget(CardService.newTextButton()
        .setText('View Documentation')
        .setOpenLink(CardService.newOpenLink().setUrl('https://github.com/imthenachoman/Nacho-Auto-Vacation-for-Gmail#readme')))) // Placeholder
    .build();
  
  // This won't work from a menu item in a GWAO as easily, 
  // but we can return it if it was a Card action. 
  // For menu, we use alert.
  SpreadsheetApp.getUi().alert('Visit https://github.com/imthenachoman/GoogleSlackAlert for documentation.');
}
