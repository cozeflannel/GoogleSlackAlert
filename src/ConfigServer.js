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
    DATE_COL:          props.getProperty('DATE_COL'),
    TRIGGER_VALUE:     props.getProperty('TRIGGER_VALUE'),
    DOWNSTREAM_CONFIG: props.getProperty('DOWNSTREAM_CONFIG')
  };
}

function saveConfig(config) {
  var spreadsheetId   = SpreadsheetApp.getActiveSpreadsheet().getId();
  var installerEmail  = Session.getActiveUser().getEmail();
  var docProps        = PropertiesService.getDocumentProperties();
  var scriptProps     = PropertiesService.getScriptProperties();

  docProps.setProperties({
    'SHEET_NAME':        config.SHEET_NAME     || '',
    'SLACK_CHANNEL':     config.SLACK_CHANNEL  || '',
    'NAME_COL':          config.NAME_COL,
    'EMAIL_COL':         config.EMAIL_COL,
    'EXTRA_INFO_COL':    config.EXTRA_INFO_COL,
    'STATUS_COL':        config.STATUS_COL,
    'DATE_COL':          config.DATE_COL       || '-1',
    'TRIGGER_VALUE':     config.TRIGGER_VALUE,
    'SPREADSHEET_ID':    spreadsheetId,
    'INSTALLER_EMAIL':   installerEmail
  });

  // Mirror keys that time-based / webhook handlers need via script props
  scriptProps.setProperty('SPREADSHEET_ID', spreadsheetId);
  scriptProps.setProperty('SLACK_CHANNEL_' + spreadsheetId, config.SLACK_CHANNEL || '');
  scriptProps.setProperty('INSTALLER_EMAIL_' + spreadsheetId, installerEmail);

  installTriggers();
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

/**
 * Disconnects the app completely from this sheet.
 * Clears all properties and deletes all triggers.
 */
function disconnectApp() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheetId = ss ? ss.getId() : null;
  
  if (!spreadsheetId) {
    var docProps = PropertiesService.getDocumentProperties();
    spreadsheetId = docProps.getProperty('SPREADSHEET_ID');
  }

  // 1. Delete triggers
  var handlersToClean = ['onSheetEdit', 'onSheetChange', 'sendWeeklyDigest', 'runDailyAlerts'];
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (handlersToClean.indexOf(t.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 2. Clear Document Properties
  PropertiesService.getDocumentProperties().deleteAllProperties();

  // 3. Clear Script Properties keyed to this sheet
  if (spreadsheetId) {
    var scriptProps = PropertiesService.getScriptProperties();
    scriptProps.deleteProperty('SLACK_TOKEN_' + spreadsheetId);
    scriptProps.deleteProperty('SLACK_CHANNEL_' + spreadsheetId);
    scriptProps.deleteProperty('INSTALLER_EMAIL_' + spreadsheetId);
    scriptProps.deleteProperty('PENDING_ALERTS_' + spreadsheetId);
    // Note: purposefully not deleting the global Slack Client ID/Secret here
  }

  return { success: true };
}

/**
 * Installs all three triggers for this spreadsheet, removing stale copies first.
 * 1. onEdit   (installable) → onSheetEdit
 * 2. onChange (installable) → onSheetChange
 * 3. Weekly time-based      → sendWeeklyDigest (Sunday 08:00)
 */
function installTriggers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // If run manually from the editor, active spreadsheet might be null.
  // Fall back to the ID saved in properties.
  if (!ss) {
    var docProps = PropertiesService.getDocumentProperties();
    var ssId = docProps.getProperty('SPREADSHEET_ID');
    if (ssId) {
      ss = SpreadsheetApp.openById(ssId);
    }
  }

  if (!ss) {
    Logger.log('Could not find active spreadsheet. Please click "Save Configuration" in the Sheet sidebar instead.');
    return;
  }
  var handlersToClean = ['onSheetEdit', 'onSheetChange', 'sendWeeklyDigest', 'runDailyAlerts'];

  // Remove any existing triggers for the handlers we manage
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (handlersToClean.indexOf(t.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 1. Installable onEdit trigger — fires when a user manually edits a cell
  ScriptApp.newTrigger('onSheetEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  // 2. Installable onChange trigger — fires when any cell value changes,
  //    including formula recalculations (e.g. a date formula flipping to "Late")
  ScriptApp.newTrigger('onSheetChange')
    .forSpreadsheet(ss)
    .onChange()
    .create();

  // 3. Weekly digest — every Sunday at 08:00 in the spreadsheet's timezone
  ScriptApp.newTrigger('sendWeeklyDigest')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(8)
    .create();

  Logger.log('Triggers installed: onSheetEdit, onSheetChange, sendWeeklyDigest');
}

// Legacy alias kept so any existing time-based trigger on runDailyAlerts
// continues to work until it is cleaned up by the next saveConfig().
function installTrigger() {
  installTriggers();
}

function getSlackConnectionStatus() {
  var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var token = PropertiesService.getScriptProperties()
                .getProperty('SLACK_TOKEN_' + spreadsheetId);
  return { connected: !!token };
}

function getSlackOAuthUrl() {
  var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var scriptProps   = PropertiesService.getScriptProperties();
  var clientId      = scriptProps.getProperty('SLACK_CLIENT_ID');

  // Use the stable /exec deployment URL stored at setup time, not the /dev URL
  var deployedUrl   = scriptProps.getProperty('DEPLOYED_WEBAPP_URL');
  if (!deployedUrl) {
    throw new Error('DEPLOYED_WEBAPP_URL not set in Script Properties. Run setupDeveloperCredentials() after deploying.');
  }

  var redirectUri = deployedUrl + '?action=oauth_callback';
  var scopes      = 'chat:write,chat:write.public,channels:read,channels:join,app_mentions:read';

  return 'https://slack.com/oauth/v2/authorize' +
    '?client_id='    + encodeURIComponent(clientId) +
    '&scope='        + encodeURIComponent(scopes) +
    '&redirect_uri=' + encodeURIComponent(redirectUri) +
    '&state='        + encodeURIComponent(spreadsheetId);
}

function disconnectSlack() {
  var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  PropertiesService.getScriptProperties()
    .deleteProperty('SLACK_TOKEN_' + spreadsheetId);
}

// Run ONCE from the Apps Script editor to store your Slack app credentials
function setupDeveloperCredentials() {
  PropertiesService.getScriptProperties().setProperties({
    'SLACK_CLIENT_ID':     'YOUR_CLIENT_ID_HERE',
    'SLACK_CLIENT_SECRET': 'YOUR_CLIENT_SECRET_HERE',
    'DEPLOYED_WEBAPP_URL': 'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec'
  });
  Logger.log('Developer credentials saved.');
}