function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('SheetAlerts')
    .addItem('Configure Alerts', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('SheetAlerts Configuration')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}