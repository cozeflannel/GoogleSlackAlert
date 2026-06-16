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
  var ss            = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheetId = ss.getId();
  var installerEmail = Session.getActiveUser().getEmail();
  var docProps      = PropertiesService.getDocumentProperties();
  var scriptProps   = PropertiesService.getScriptProperties();

  // ── DocumentProperties (scoped to this spreadsheet) ──────────────────────
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

  // ── ScriptProperties (global — keyed per spreadsheet to support multi-tenant) ──
  // SLACK_CHANNEL and INSTALLER_EMAIL are keyed so webhook handlers can look
  // them up without a DocumentProperties context (e.g. doPost, publishAppHome).
  // We also mirror sheet config here so time-based triggers and doPost can
  // run runConditionCheck() reliably (DocumentProperties = null in those contexts).
  scriptProps.setProperty('SPREADSHEET_ID', spreadsheetId);
  scriptProps.setProperty('SLACK_CHANNEL_'    + spreadsheetId, config.SLACK_CHANNEL || '');
  scriptProps.setProperty('INSTALLER_EMAIL_'  + spreadsheetId, installerEmail);
  scriptProps.setProperty('SHEET_NAME_'       + spreadsheetId, config.SHEET_NAME     || '');
  scriptProps.setProperty('STATUS_COL_'       + spreadsheetId, config.STATUS_COL     || '-1');
  scriptProps.setProperty('TRIGGER_VALUE_'    + spreadsheetId, config.TRIGGER_VALUE  || '');
  scriptProps.setProperty('NAME_COL_'         + spreadsheetId, config.NAME_COL       || '-1');
  scriptProps.setProperty('EMAIL_COL_'        + spreadsheetId, config.EMAIL_COL      || '-1');
  scriptProps.setProperty('EXTRA_INFO_COL_'   + spreadsheetId, config.EXTRA_INFO_COL || '-1');
  scriptProps.setProperty('DATE_COL_'         + spreadsheetId, config.DATE_COL       || '-1');

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
 * Completely disconnects the add-on from this sheet.
 * Deletes all managed triggers, clears all properties.
 */
function disconnectApp() {
  var ss            = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheetId = ss ? ss.getId() : null;

  if (!spreadsheetId) {
    var docProps = PropertiesService.getDocumentProperties();
    spreadsheetId = docProps.getProperty('SPREADSHEET_ID');
  }

  // 1. Delete managed triggers
  var handlersToClean = [
    'onSheetEdit', 'onSheetChange', 'sendWeeklyDigest', 'runDailyAlerts'
  ];
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (handlersToClean.indexOf(t.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 2. Clear DocumentProperties
  PropertiesService.getDocumentProperties().deleteAllProperties();

  // 3. Clear per-sheet ScriptProperties
  if (spreadsheetId) {
    var scriptProps = PropertiesService.getScriptProperties();
    scriptProps.deleteProperty('SLACK_TOKEN_'    + spreadsheetId);
    scriptProps.deleteProperty('SLACK_CHANNEL_'  + spreadsheetId);
    scriptProps.deleteProperty('INSTALLER_EMAIL_'+ spreadsheetId);
    scriptProps.deleteProperty('PENDING_ALERTS_' + spreadsheetId);
    // Intentionally not deleting global SLACK_CLIENT_ID/SECRET or DEPLOYED_WEBAPP_URL
  }

  return { success: true };
}

/**
 * Installs all triggers for this spreadsheet, removing stale copies first.
 *
 * Triggers installed:
 *   1. onEdit    (installable) → onSheetEdit
 *   2. onChange  (installable) → onSheetChange
 *   3. Weekly time-based       → sendWeeklyDigest (Sunday 08:00)
 */
function installTriggers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    var ssId = PropertiesService.getDocumentProperties()
                 .getProperty('SPREADSHEET_ID');
    if (ssId) {
      ss = SpreadsheetApp.openById(ssId);
    }
  }

  if (!ss) {
    Logger.log('installTriggers: no active spreadsheet found.');
    return;
  }

  var handlersToClean = [
    'onSheetEdit', 'onSheetChange', 'sendWeeklyDigest', 'runDailyAlerts'
  ];

  // Remove any existing triggers for our handlers
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (handlersToClean.indexOf(t.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 1. Installable onEdit — fires when a user manually edits a cell
  ScriptApp.newTrigger('onSheetEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  // 2. Installable onChange — fires on any value change including formula recalcs
  ScriptApp.newTrigger('onSheetChange')
    .forSpreadsheet(ss)
    .onChange()
    .create();

  // 3. Weekly digest — every Sunday at 08:00
  ScriptApp.newTrigger('sendWeeklyDigest')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(8)
    .create();

  Logger.log('installTriggers: onSheetEdit, onSheetChange, sendWeeklyDigest installed.');
}

// Legacy alias
function installTrigger() { installTriggers(); }

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

  var deployedUrl   = scriptProps.getProperty('DEPLOYED_WEBAPP_URL');
  if (!deployedUrl) {
    throw new Error(
      'DEPLOYED_WEBAPP_URL is not set in Script Properties. ' +
      'Run setupDeveloperCredentials() after deploying the web app.'
    );
  }

  var redirectUri = deployedUrl + '?action=oauth_callback';
  var scopes      = 'chat:write,chat:write.public,channels:read,channels:join,app_mentions:read,im:write,im:history,incoming-webhook';

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

/**
 * Run ONCE from the Apps Script editor after your first web app deployment.
 * Replace the placeholder values with your real credentials.
 */
function setupDeveloperCredentials() {
  PropertiesService.getScriptProperties().setProperties({
    'SLACK_CLIENT_ID':     'YOUR_CLIENT_ID_HERE',
    'SLACK_CLIENT_SECRET': 'YOUR_CLIENT_SECRET_HERE',
    'DEPLOYED_WEBAPP_URL': 'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec'
  });
  Logger.log('Developer credentials saved to ScriptProperties.');
}