var SUPABASE_FUNCTION_URL = 'https://apjftvnmskckrhgrdpbk.supabase.co/functions/v1/bot';

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
 * Saves the bot configuration to the Supabase installations table.
 * This is now the primary write path.
 */
function saveConfig(config) {
  var ss            = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheetId = ss.getId();
  var installerEmail = Session.getActiveUser().getEmail();
  var docProps      = PropertiesService.getDocumentProperties();

  // ── DocumentProperties (Keep local copy for fast UI access) ────────────────
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
    'ACTIONABLE_COLS':            config.ACTIONABLE_COLS            || '[]',
    'PENDING_ALERTS':             config.PENDING_ALERTS || '[]'
  });

  installTriggers();

  // ── Supabase Write (The Source of Truth) ──────────────────────────────────
  try {
    var scriptProps = PropertiesService.getScriptProperties();
    var supabaseUrl = scriptProps.getProperty('SUPABASE_URL') + '/rest/v1/installations';
    var serviceKey  = scriptProps.getProperty('SUP_SECRET_KEY');

    // Get actual headers from the sheet for the mirror
    var headers = [];
    var sheet = ss.getSheetByName(config.SHEET_NAME);
    if (sheet) {
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    }

    var payload = {
      spreadsheet_id: spreadsheetId,
      installer_email: installerEmail,
      config: JSON.stringify({
        sheet_name: config.SHEET_NAME,
        slack_channel: config.SLACK_CHANNEL,
        status_col: config.STATUS_COL,
        trigger_value: config.TRIGGER_VALUE,
        actionable_cols: config.ACTIONABLE_COLS
      }),
      headers: JSON.stringify(headers)
    };

    var options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'apikey': serviceKey,
        'Authorization': 'Bearer ' + serviceKey,
        'Prefer': 'resolution=merge-duplicates'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(supabaseUrl, options);
    if (response.getResponseCode() >= 300) {
      Logger.log('saveConfig Supabase error: ' + response.getContentText());
    }
  } catch (err) {
    Logger.log('saveConfig Supabase CRITICAL: ' + err.toString());
  }
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
 */
function disconnectApp() {
  var ss            = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheetId = ss ? ss.getId() : null;

  if (!spreadsheetId) {
    var docProps = PropertiesService.getDocumentProperties();
    spreadsheetId = docProps.getProperty('SPREADSHEET_ID');
  }

  var handlersToClean = [
    'onSheetEdit', 'onSheetChange', 'sendWeeklyDigest', 'runDailyAlerts'
  ];
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (handlersToClean.indexOf(t.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(t);
    }
  });

  PropertiesService.getDocumentProperties().deleteAllProperties();

  // Also remove from Supabase
  disconnectSlack();

  return { success: true };
}

function installTriggers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;

  var handlersToClean = ['onSheetEdit', 'onSheetChange', 'sendWeeklyDigest', 'runDailyAlerts'];
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (handlersToClean.indexOf(t.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('onSheetEdit').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('onSheetChange').forSpreadsheet(ss).onChange().create();
  ScriptApp.newTrigger('sendWeeklyDigest').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(8).create();
}

function getSlackConnectionStatus() {
  var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var scriptProps = PropertiesService.getScriptProperties();
  var supabaseUrl = scriptProps.getProperty('SUPABASE_URL') + '/rest/v1/installations?spreadsheet_id=eq.' + spreadsheetId + '&select=slack_bot_token';
  var serviceKey  = scriptProps.getProperty('SUP_SECRET_KEY');

  try {
    var response = UrlFetchApp.fetch(supabaseUrl, {
      headers: { 'apikey': serviceKey, 'Authorization': 'Bearer ' + serviceKey },
      muteHttpExceptions: true
    });
    var data = JSON.parse(response.getContentText());
    return { connected: data && data.length > 0 };
  } catch (e) {
    return { connected: false };
  }
}

function getSlackOAuthUrl() {
  var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  return SUPABASE_FUNCTION_URL + '?action=slack_oauth&state=' + encodeURIComponent(spreadsheetId);
}

function disconnectSlack() {
  var ss            = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheetId = ss ? ss.getId() : null;
  if (!spreadsheetId) {
    spreadsheetId = PropertiesService.getDocumentProperties().getProperty('SPREADSHEET_ID');
  }

  PropertiesService.getDocumentProperties().deleteProperty('SLACK_CHANNEL');

  if (spreadsheetId) {
    var scriptProps = PropertiesService.getScriptProperties();
    var supabaseUrl = scriptProps.getProperty('SUPABASE_URL') + '/rest/v1/installations?spreadsheet_id=eq.' + encodeURIComponent(spreadsheetId);
    var serviceKey  = scriptProps.getProperty('SUP_SECRET_KEY');

    UrlFetchApp.fetch(supabaseUrl, {
      method: 'delete',
      headers: { 'apikey': serviceKey, 'Authorization': 'Bearer ' + serviceKey },
      muteHttpExceptions: true
    });
  }
}
