function getSheetsInfo() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sheets  = ss.getSheets();
  var results = [];

  sheets.forEach(function(sheet) {
    var lastCol = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    var headers = [];
    var name    = sheet.getName();
    var score   = 0;

    if (lastCol > 0) {
      headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
                  .filter(function(h) { return h !== ''; });
    }

    score += headers.length * 2;
    if (lastRow > 1) score += 5;
    if (/^sheet\d*$/i.test(name)) score -= 3;

    results.push({
      name:    name,
      headers: headers,
      hasData: lastRow > 1,
      isEmpty: headers.length === 0,
      score:   score
    });
  });

  results.sort(function(a, b) { return b.score - a.score; });

  var savedSheetName = PropertiesService.getDocumentProperties()
                         .getProperty('SHEET_NAME');

  return { sheets: results, savedSheetName: savedSheetName || null };
}

function getSheetState(sheetName) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();

  if (!sheet) return { state: 'not_found', headers: [] };

  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();

  if (lastCol === 0) return { state: 'empty', headers: [] };

  var headers        = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var hasRealHeaders = headers.some(function(h) { return h !== ''; });

  if (!hasRealHeaders) return { state: 'empty', headers: [] };

  return {
    state:   lastRow > 1 ? 'has_data' : 'has_headers_only',
    headers: headers
  };
}

function createSheetHeaders(sheetName, headers) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) return { success: false, error: 'Sheet not found.' };

  var columns = [
    headers.nameCol      || 'Name',
    headers.emailCol     || 'Email',
    headers.statusCol    || 'Status',
    headers.extraInfoCol || 'Extra Info'
  ].filter(function(c) { return c && c.trim() !== ''; });

  sheet.getRange(1, 1, 1, columns.length).setValues([columns]);

  var headerRange = sheet.getRange(1, 1, 1, columns.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f3f3f3');
  sheet.setFrozenRows(1);

  return { success: true, columns: columns };
}

function getAllSheetNames() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .map(function(s) { return s.getName(); });
}

function getHeadersForSheet(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet || sheet.getLastColumn() === 0) return [];
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function getConfig() {
  var props = PropertiesService.getDocumentProperties();
  return {
    SHEET_NAME:        props.getProperty('SHEET_NAME'),
    SLACK_CHANNEL:     props.getProperty('SLACK_CHANNEL'),
    NAME_COL:          props.getProperty('NAME_COL'),
    EMAIL_COL:         props.getProperty('EMAIL_COL'),
    EXTRA_INFO_COL:    props.getProperty('EXTRA_INFO_COL'),
    STATUS_COL:        props.getProperty('STATUS_COL'),
    TRIGGER_VALUE:     props.getProperty('TRIGGER_VALUE'),
    DOWNSTREAM_CONFIG: props.getProperty('DOWNSTREAM_CONFIG')
  };
}

function saveConfig(config) {
  var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var docProps      = PropertiesService.getDocumentProperties();
  var scriptProps   = PropertiesService.getScriptProperties();

  docProps.setProperties({
    'SHEET_NAME':     config.SHEET_NAME     || '',
    'SLACK_CHANNEL':  config.SLACK_CHANNEL  || '',
    'NAME_COL':       config.NAME_COL,
    'EMAIL_COL':      config.EMAIL_COL,
    'EXTRA_INFO_COL': config.EXTRA_INFO_COL,
    'STATUS_COL':     config.STATUS_COL,
    'TRIGGER_VALUE':  config.TRIGGER_VALUE,
    'SPREADSHEET_ID': spreadsheetId
  });

  // Write SPREADSHEET_ID to script props so time-based triggers can find it
  scriptProps.setProperty('SPREADSHEET_ID', spreadsheetId);
  // Set dynamic key separately to avoid linter error
  scriptProps.setProperty('SLACK_CHANNEL_' + spreadsheetId, config.SLACK_CHANNEL || '');

  installTrigger();
}

function saveDownstreamConfig(config) {
  PropertiesService.getDocumentProperties()
    .setProperty('DOWNSTREAM_CONFIG', JSON.stringify(config));
}

function getDownstreamConfig() {
  var raw = PropertiesService.getDocumentProperties()
              .getProperty('DOWNSTREAM_CONFIG');
  return raw ? JSON.parse(raw) : null;
}

function installTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'runDailyAlerts') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('runDailyAlerts')
    .timeBased().everyDays(1).atHour(8).create();
}

function getSlackConnectionStatus() {
  var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var token = PropertiesService.getScriptProperties()
                .getProperty('SLACK_TOKEN_' + spreadsheetId);
  return { connected: !!token };
}

function getSlackOAuthUrl() {
  var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var clientId      = PropertiesService.getScriptProperties()
                        .getProperty('SLACK_CLIENT_ID');
  var redirectUri   = ScriptApp.getService().getUrl() + "?action=oauth_callback";
  var scopes        = 'chat:write,channels:read,channels:join,app_mentions:read';

  return 'https://slack.com/oauth/v2/authorize' +
    '?client_id='     + encodeURIComponent(clientId) +
    '&scope='         + encodeURIComponent(scopes) +
    '&redirect_uri='  + encodeURIComponent(redirectUri) +
    '&state='         + encodeURIComponent(spreadsheetId);
}

function disconnectSlack() {
  var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  PropertiesService.getScriptProperties()
    .deleteProperty('SLACK_TOKEN_' + spreadsheetId);
}

function setupDeveloperCredentials() {
  PropertiesService.getScriptProperties().setProperties({
    'SLACK_CLIENT_ID':     'YOUR_CLIENT_ID_HERE',
    'SLACK_CLIENT_SECRET': 'YOUR_CLIENT_SECRET_HERE'
  });
  Logger.log('Developer credentials saved.');
}