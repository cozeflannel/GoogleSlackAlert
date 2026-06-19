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

/**
 * Intelligently guesses column mappings based on header names.
 */
function guessColumnMapping(sheetName) {
  var headers = getHeadersForSheet(sheetName);
  var mapping = {
    NAME_COL: -1,
    EMAIL_COL: -1,
    EXTRA_INFO_COL: -1,
    STATUS_COL: -1,
    DATE_COL: -1
  };

  headers.forEach(function(h, i) {
    var name = String(h).toLowerCase();
    if (mapping.NAME_COL === -1 && /name|client|customer|user|id/i.test(name)) mapping.NAME_COL = i;
    if (mapping.EMAIL_COL === -1 && /email|mail|contact/i.test(name)) mapping.EMAIL_COL = i;
    if (mapping.STATUS_COL === -1 && /status|condition|stage|state/i.test(name)) mapping.STATUS_COL = i;
    if (mapping.DATE_COL === -1 && /date|due|time/i.test(name)) mapping.DATE_COL = i;
    if (mapping.EXTRA_INFO_COL === -1 && /info|note|description|memo/i.test(name)) mapping.EXTRA_INFO_COL = i;
  });

  return mapping;
}

function getConfig() {
  var props = PropertiesService.getDocumentProperties();
  return {
    SHEET_NAME:                 props.getProperty('SHEET_NAME'),
    SLACK_CHANNEL:              props.getProperty('SLACK_CHANNEL'),
    NAME_COL:                   props.getProperty('NAME_COL'),
    EMAIL_COL:                  props.getProperty('EMAIL_COL'),
    EXTRA_INFO_COL:             props.getProperty('EXTRA_INFO_COL'),
    STATUS_COL:                 props.getProperty('STATUS_COL'),
    DATE_COL:                   props.getProperty('DATE_COL'),
    TRIGGER_VALUE:              props.getProperty('TRIGGER_VALUE'),
    DOWNSTREAM_CONFIG:          props.getProperty('DOWNSTREAM_CONFIG'),
    AUTO_TIMESTAMP_COL:         props.getProperty('AUTO_TIMESTAMP_COL'),
    AUTO_TIMESTAMP_TRIGGER_COL: props.getProperty('AUTO_TIMESTAMP_TRIGGER_COL'),
    AUTO_TIMESTAMP_VALUE:       props.getProperty('AUTO_TIMESTAMP_VALUE'),
    DUE_DATE_COL:               props.getProperty('DUE_DATE_COL'),
    FINAL_STATUS_COL:           props.getProperty('FINAL_STATUS_COL'),
    ACTIONABLE_COLS:            props.getProperty('ACTIONABLE_COLS')
  };
}

/**
 * Mirrors the current bot configuration to the Supabase database.
 * This eliminates the need for the Supabase function to call GAS during modal opening,
 * preventing the Slack 3-second trigger_id timeout.
 */
function syncConfigToSupabase(spreadsheetId) {
  try {
    var scriptProps = PropertiesService.getScriptProperties();
    var token       = scriptProps.getProperty('SLACK_TOKEN_' + spreadsheetId);
    var sheetName   = scriptProps.getProperty('SHEET_NAME_' + spreadsheetId);
    var actionable  = scriptProps.getProperty('ACTIONS_' + spreadsheetId) || '[]';

    if (!token) {
      Logger.log('syncConfigToSupabase: No token found for ' + spreadsheetId + '. Skipping.');
      return;
    }

    // Get actual headers from the sheet
    var headers = [];
    try {
      var ss = SpreadsheetApp.openById(spreadsheetId);
      var sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      }
    } catch (e) {
      Logger.log('syncConfigToSupabase: Error reading headers: ' + e.toString());
    }

    // Use Supabase REST API (Upsert)
    var supabaseBaseUrl = PropertiesService.getScriptProperties().getProperty('SUPABASE_URL');
    // Supabase secret key (formerly service_role) — grants full DB access, bypasses RLS. Never expose client-side.
    var serviceKey      = PropertiesService.getScriptProperties().getProperty('SUP_SECRET_KEY');
    var supabaseUrl     = supabaseBaseUrl + '/rest/v1/bot_configs';

    var payload = {
      spreadsheet_id: spreadsheetId,
      token: token,
      headers: JSON.stringify(headers),
      actionable_cols: actionable
    };

    var options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'apikey': serviceKey,
        'Authorization': 'Bearer ' + serviceKey,
        'Prefer': 'resolution=merge-duplicates' // Upsert
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(supabaseUrl, options);
    if (response.getResponseCode() !== 201 && response.getResponseCode() !== 200) {
      Logger.log('syncConfigToSupabase error: ' + response.getContentText());
    } else {
      Logger.log('syncConfigToSupabase: Successfully mirrored config for ' + spreadsheetId);
    }
  } catch (err) {
    Logger.log('syncConfigToSupabase CRITICAL: ' + err.toString());
  }
}

function saveConfig(config) {
  var ss            = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheetId = ss.getId();
  var installerEmail = Session.getActiveUser().getEmail();
  var docProps      = PropertiesService.getDocumentProperties();
  var scriptProps   = PropertiesService.getScriptProperties();

  // ── DocumentProperties (scoped to this spreadsheet) ──────────────────────
  docProps.setProperties({
    'SHEET_NAME':                 config.SHEET_NAME     || '',
    'SLACK_CHANNEL':              config.SLACK_CHANNEL  || '',
    'NAME_COL':                   config.NAME_COL,
    'EMAIL_COL':                  config.EMAIL_COL,
    'EXTRA_INFO_COL':             config.EXTRA_INFO_COL,
    'STATUS_COL':                 config.STATUS_COL,
    'DATE_COL':                   config.DATE_COL       || '-1',
    'TRIGGER_VALUE':              config.TRIGGER_VALUE,
    'SPREADSHEET_ID':             spreadsheetId,
    'INSTALLER_EMAIL':            installerEmail,
    'AUTO_TIMESTAMP_COL':         config.AUTO_TIMESTAMP_COL         || '-1',
    'AUTO_TIMESTAMP_TRIGGER_COL': config.AUTO_TIMESTAMP_TRIGGER_COL || '-1',
    'AUTO_TIMESTAMP_VALUE':       config.AUTO_TIMESTAMP_VALUE       || '',
    'DUE_DATE_COL':               config.DUE_DATE_COL               || '-1',
    'FINAL_STATUS_COL':           config.FINAL_STATUS_COL           || '-1',
    'ACTIONABLE_COLS':            config.ACTIONABLE_COLS            || '[]'
  });

  // ── ScriptProperties (global — keyed per spreadsheet to support multi-tenant) ──
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

  // New sync properties for multi-tenant triggers
  scriptProps.setProperty('AT_COL_'         + spreadsheetId, config.AUTO_TIMESTAMP_COL         || '-1');
  scriptProps.setProperty('AT_TRIG_COL_'    + spreadsheetId, config.AUTO_TIMESTAMP_TRIGGER_COL || '-1');
  scriptProps.setProperty('AT_VAL_'         + spreadsheetId, config.AUTO_TIMESTAMP_VALUE       || '');
  scriptProps.setProperty('ACTIONS_'        + spreadsheetId, config.ACTIONABLE_COLS            || '[]');

  installTriggers();
  
  // Sync to Supabase mirror to prevent modal timeout
  syncConfigToSupabase(spreadsheetId);
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
  var docProps      = PropertiesService.getDocumentProperties();
  var scriptProps   = PropertiesService.getScriptProperties();
  var token = docProps.getProperty('SLACK_TOKEN') || scriptProps.getProperty('SLACK_TOKEN_' + spreadsheetId);
  return { connected: !!token };
}

function getSlackOAuthUrl() {
  var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var docProps      = PropertiesService.getDocumentProperties();
  var scriptProps   = PropertiesService.getScriptProperties();
  var clientId      = docProps.getProperty('SLACK_CLIENT_ID') || scriptProps.getProperty('SLACK_CLIENT_ID_' + spreadsheetId);

  // This is the SUPABASE function URL, not a Google Apps Script URL — used for Slack OAuth redirect_uri construction.
  var deployedUrl   = docProps.getProperty('DEPLOYED_WEBAPP_URL') || scriptProps.getProperty('DEPLOYED_WEBAPP_URL_' + spreadsheetId);
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
  deleteSupabaseConfigMirror(spreadsheetId);
}

function deleteSupabaseConfigMirror(spreadsheetId) {
  try {
    var scriptProps = PropertiesService.getScriptProperties();
    // Supabase secret key (formerly service_role) — grants full DB access, bypasses RLS. Never expose client-side.
    var secretKey   = scriptProps.getProperty('SUP_SECRET_KEY');
    if (!secretKey) {
      Logger.log('deleteSupabaseConfigMirror: No SUP_SECRET_KEY set. Skipping.');
      return;
    }

    var supabaseBaseUrl = scriptProps.getProperty('SUPABASE_URL');
    var supabaseUrl     = supabaseBaseUrl + '/rest/v1/bot_configs?spreadsheet_id=eq.' + encodeURIComponent(spreadsheetId);

    var options = {
      method: 'delete',
      headers: {
        'apikey': secretKey,
        'Authorization': 'Bearer ' + secretKey
      },
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(supabaseUrl, options);
    if (response.getResponseCode() >= 300) {
      Logger.log('deleteSupabaseConfigMirror error: ' + response.getContentText());
    } else {
      Logger.log('deleteSupabaseConfigMirror: Removed mirror for ' + spreadsheetId);
    }
  } catch (err) {
    Logger.log('deleteSupabaseConfigMirror CRITICAL: ' + err.toString());
  }
}

/**
 * Run ONCE from the Apps Script editor after your first web app deployment.
 * Replace the placeholder values with your real credentials.
 */
function setupDeveloperCredentials() {
  PropertiesService.getScriptProperties().setProperties({
    'SLACK_CLIENT_ID':     'YOUR_SLACK_CLIENT_ID_HERE',
    'SLACK_CLIENT_SECRET': 'YOUR_SLACK_CLIENT_SECRET_HERE',
    // This is the SUPABASE function URL, not a Google Apps Script URL — used for Slack OAuth redirect_uri construction.
    'DEPLOYED_WEBAPP_URL': 'YOUR_SUPABASE_FUNCTION_URL_HERE',
    'SUPABASE_URL':        'YOUR_SUPABASE_PROJECT_URL_HERE',
    // Supabase secret key (formerly service_role) — grants full DB access, bypasses RLS. Never expose client-side.
    'SUP_SECRET_KEY': 'YOUR_SUP_SECRET_KEY_HERE'
  });
  Logger.log('Developer credentials saved to ScriptProperties.');
}