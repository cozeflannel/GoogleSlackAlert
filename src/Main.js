// ─────────────────────────────────────────────────────────────────────────────
// TRIGGER HANDLERS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Installable onEdit trigger handler.
 *
 * For a standalone add-on, installable triggers DO receive e.source correctly
 * because the trigger is bound to a specific spreadsheet at creation time.
 *
 * Skips edits by the installer (they'd spam themselves).
 * Skips the internal AlertsLog sheet.
 */
function onSheetEdit(e) {
  try {
    // e.source IS available on installable onEdit triggers even in standalone scripts
    var ss            = e.source;
    var spreadsheetId = ss.getId();

    var editorEmail = e.user ? e.user.email : '';
    if (!editorEmail) {
      Logger.log('onSheetEdit: anonymous edit, skipping.');
      return;
    }

    var docProps       = PropertiesService.getDocumentProperties();
    var installerEmail = docProps.getProperty('INSTALLER_EMAIL') || '';

    // Installer editing their own sheet — no self-notification
    if (installerEmail && editorEmail.toLowerCase() === installerEmail.toLowerCase()) {
      Logger.log('onSheetEdit: edit by installer (' + editorEmail + '), skipping.');
      return;
    }

    var range     = e.range;
    var sheetName = range.getSheet().getName();

    if (sheetName === 'AlertsLog') return;

    var editData = {
      editorEmail:   editorEmail,
      sheetName:     sheetName,
      row:           range.getRow(),
      column:        range.getColumn(),
      columnLetter:  columnToLetter(range.getColumn()),
      oldValue:      (e.oldValue !== undefined && e.oldValue !== null) ? String(e.oldValue) : '(empty)',
      newValue:      (e.value    !== undefined && e.value    !== null) ? String(e.value)    : '(empty)',
      spreadsheetId: spreadsheetId,
      timestamp:     new Date()
    };

    Logger.log('onSheetEdit: editor=' + editorEmail +
               ' cell=' + editData.columnLetter + editData.row +
               ' sheet=' + sheetName +
               ' spreadsheetId=' + spreadsheetId);

    var scriptProps    = PropertiesService.getScriptProperties();
    var slackConnected = !!scriptProps.getProperty('SLACK_TOKEN_' + spreadsheetId);

    var slackResult = false;
    if (slackConnected) {
      try {
        slackResult = sendEditNotification(editData);
        Logger.log('onSheetEdit: Slack result=' + slackResult);
      } catch (slackErr) {
        Logger.log('onSheetEdit: Slack error: ' + slackErr.toString());
      }
    } else {
      Logger.log('onSheetEdit: Slack not connected for spreadsheetId=' + spreadsheetId);
    }

    var emailResult = false;
    if (installerEmail) {
      try {
        emailResult = sendGmailEditAlert(editData, installerEmail);
        Logger.log('onSheetEdit: email result=' + emailResult);
      } catch (mailErr) {
        Logger.log('onSheetEdit: email error: ' + mailErr.toString());
      }
    }

    var logSheet = ss.getSheetByName('AlertsLog') || _createLogSheet(ss);
    logSheet.appendRow([
      new Date(),
      editData.row,
      editorEmail,
      'Manual Edit',
      generateUUID(),
      emailResult,
      slackResult,
      false, '', '', '',
      sheetName,
      false, '',
      editData.columnLetter,
      editData.oldValue,
      editData.newValue
    ]);

  } catch (err) {
    Logger.log('onSheetEdit UNHANDLED ERROR: ' + err.toString() + '\n' + err.stack);
  }
}

/**
 * Installable onChange trigger handler.
 *
 * IMPORTANT — standalone script limitation:
 * The onChange event object does NOT reliably carry e.source in a standalone
 * add-on. We therefore resolve the spreadsheet from DocumentProperties
 * (SPREADSHEET_ID), which is always correctly scoped to the container
 * spreadsheet because DocumentProperties are per-document.
 *
 * This is safe because a single installed copy of the add-on has exactly one
 * DocumentProperties store — the one for the sheet it was installed on.
 */
function onSheetChange(e) {
  try {
    Logger.log('onSheetChange: fired. changeType=' +
               (e && e.changeType ? e.changeType : 'unknown'));
    runConditionCheck();
  } catch (err) {
    Logger.log('onSheetChange UNHANDLED ERROR: ' + err.toString() + '\n' + err.stack);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// CONDITION CHECK
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Scans the monitored sheet for rows matching the trigger condition and fires
 * Slack + email alerts for newly matching rows.
 *
 * Spreadsheet resolution order:
 *   1. DocumentProperties.SPREADSHEET_ID  ← always correct for a given install
 *   2. Try SpreadsheetApp.getActiveSpreadsheet() as last resort (editor runs only)
 */
function runConditionCheck() {
  // ── 1. Resolve spreadsheet ID ─────────────────────────────────────────────
  var docProps      = PropertiesService.getDocumentProperties();
  var spreadsheetId = docProps.getProperty('SPREADSHEET_ID');

  if (!spreadsheetId) {
    // Last resort: editor / test run context
    try {
      var active = SpreadsheetApp.getActiveSpreadsheet();
      if (active) spreadsheetId = active.getId();
    } catch (e) { /* no active spreadsheet */ }
  }

  if (!spreadsheetId) {
    Logger.log('runConditionCheck: SPREADSHEET_ID not set. ' +
               'Open the SheetAlerts sidebar and click Save Configuration.');
    return;
  }

  // ── 2. Load config ────────────────────────────────────────────────────────
  var sheetName    = docProps.getProperty('SHEET_NAME');
  var nameCol      = parseInt(docProps.getProperty('NAME_COL')       || '-1');
  var emailCol     = parseInt(docProps.getProperty('EMAIL_COL')      || '-1');
  var extraInfoCol = parseInt(docProps.getProperty('EXTRA_INFO_COL') || '-1');
  var statusCol    = parseInt(docProps.getProperty('STATUS_COL')     || '-1');
  var triggerValue = docProps.getProperty('TRIGGER_VALUE');

  Logger.log('runConditionCheck: spreadsheetId=' + spreadsheetId +
             ' sheetName=' + sheetName +
             ' statusCol=' + statusCol +
             ' triggerValue=' + triggerValue);

  if (!sheetName) {
    Logger.log('runConditionCheck: SHEET_NAME not configured.');
    return;
  }
  if (statusCol === -1 || !triggerValue) {
    Logger.log('runConditionCheck: alert condition not fully configured.');
    return;
  }

  // ── 3. Open spreadsheet and sheet ────────────────────────────────────────
  var ss;
  try {
    ss = SpreadsheetApp.openById(spreadsheetId);
  } catch (openErr) {
    Logger.log('runConditionCheck: cannot open spreadsheet ' +
               spreadsheetId + ' — ' + openErr.toString());
    return;
  }

  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log('runConditionCheck: sheet "' + sheetName + '" not found.');
    return;
  }

  // ── 4. Scan rows ──────────────────────────────────────────────────────────
  var scriptProps    = PropertiesService.getScriptProperties();
  var slackConnected = !!scriptProps.getProperty('SLACK_TOKEN_' + spreadsheetId);

  Logger.log('runConditionCheck: slackConnected=' + slackConnected +
             ' rows to scan=' + (sheet.getLastRow() - 1));

  var data             = sheet.getDataRange().getValues();
  var pendingAlertsRaw = docProps.getProperty('PENDING_ALERTS');
  var pendingAlerts    = pendingAlertsRaw ? JSON.parse(pendingAlertsRaw) : [];

  for (var i = 1; i < data.length; i++) {
    var cellValue = String(data[i][statusCol]).trim();
    var trigger   = String(triggerValue).trim();

    if (cellValue !== trigger) continue;

    var rowData = {
      rowId:         data[i][0],
      clientName:    nameCol      >= 0 ? data[i][nameCol]      : 'Unknown',
      extraInfo:     extraInfoCol >= 0 ? data[i][extraInfoCol] : '',
      status:        data[i][statusCol],
      email:         emailCol     >= 0 ? data[i][emailCol]     : '',
      rowIndex:      i + 1,
      spreadsheetId: spreadsheetId,
      sheetName:     sheetName
    };

    var existingToken = getExistingUnresolvedToken(ss, rowData.rowIndex, sheetName);
    if (existingToken) {
      Logger.log('runConditionCheck: row ' + rowData.rowIndex +
                 ' already has unresolved alert, skipping.');
      continue;
    }

    var cascadeToken = getCascadeTokenForRow(ss, sheetName, rowData.rowIndex);
    if (cascadeToken) {
      Logger.log('runConditionCheck: row ' + rowData.rowIndex +
                 ' already notified via cascade, skipping.');
      continue;
    }

    rowData.token = generateUUID();

    Logger.log('runConditionCheck: FIRING alert for row ' + rowData.rowIndex +
               ' name=' + rowData.clientName + ' status=' + rowData.status);

    var slackResult = false;
    if (slackConnected) {
      try {
        slackResult = sendSlackAlert(rowData);
        Logger.log('runConditionCheck: Slack alert result=' + slackResult);
      } catch (slackErr) {
        Logger.log('runConditionCheck: Slack error on row ' +
                   rowData.rowIndex + ': ' + slackErr.toString());
      }
    }

    var emailResult = false;
    if (rowData.email) {
      try {
        emailResult = sendGmailAlert(rowData);
        Logger.log('runConditionCheck: email result=' + emailResult);
      } catch (mailErr) {
        Logger.log('runConditionCheck: email error on row ' +
                   rowData.rowIndex + ': ' + mailErr.toString());
      }
    }

    logAlert(ss, rowData, emailResult, slackResult, sheetName, false);
    pendingAlerts.push(rowData);
  }

  var alertsJson = JSON.stringify(pendingAlerts);
  docProps.setProperty('PENDING_ALERTS', alertsJson);
  scriptProps.setProperty('PENDING_ALERTS_' + spreadsheetId, alertsJson);

  Logger.log('runConditionCheck: complete.');
}

// Legacy alias — keeps any old time-based trigger on runDailyAlerts working
function runDailyAlerts() { runConditionCheck(); }

// ─────────────────────────────────────────────────────────────────────────────
// WEEKLY DIGEST
// ─────────────────────────────────────────────────────────────────────────────

function sendWeeklyDigest() {
  var docProps      = PropertiesService.getDocumentProperties();
  var spreadsheetId = docProps.getProperty('SPREADSHEET_ID');

  if (!spreadsheetId) {
    Logger.log('sendWeeklyDigest: SPREADSHEET_ID not set.');
    return;
  }

  var scriptProps    = PropertiesService.getScriptProperties();
  var slackConnected = !!scriptProps.getProperty('SLACK_TOKEN_' + spreadsheetId);

  // Always run a condition check first to catch any new rows
  runConditionCheck();

  if (!slackConnected) {
    Logger.log('sendWeeklyDigest: Slack not connected, skipping digest post.');
    return;
  }

  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var logSheet    = spreadsheet.getSheetByName('AlertsLog');

  if (!logSheet) {
    sendWeeklyDigestSlack(spreadsheetId, {
      totalAlerts: 0, resolved: 0, pending: 0, manualEdits: 0, newEditors: []
    });
    return;
  }

  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 7);

  var data        = logSheet.getDataRange().getValues();
  var totalAlerts = 0, resolved = 0, pending = 0, manualEdits = 0;
  var editorSet   = {};

  for (var i = 1; i < data.length; i++) {
    var ts = data[i][0];
    if (!(ts instanceof Date) || ts < cutoff) continue;

    var changeType = data[i][3] || '';
    if (changeType === 'Manual Edit') {
      manualEdits++;
      if (data[i][2]) editorSet[data[i][2]] = true;
    } else {
      totalAlerts++;
      if (data[i][7] === true) resolved++;
      else pending++;
    }
  }

  sendWeeklyDigestSlack(spreadsheetId, {
    totalAlerts: totalAlerts,
    resolved:    resolved,
    pending:     pending,
    manualEdits: manualEdits,
    newEditors:  Object.keys(editorSet)
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// CASCADE
// ─────────────────────────────────────────────────────────────────────────────

function fireCascadeIfConfigured(spreadsheetId, resolvedRowData) {
  var docProps      = PropertiesService.getDocumentProperties();
  var downstreamRaw = docProps.getProperty('DOWNSTREAM_CONFIG');
  if (!downstreamRaw) return;

  var downstream = JSON.parse(downstreamRaw);
  if (!downstream.enabled) return;

  var spreadsheet     = SpreadsheetApp.openById(spreadsheetId);
  var downstreamSheet = spreadsheet.getSheetByName(downstream.sheetName);
  if (!downstreamSheet) {
    Logger.log('fireCascadeIfConfigured: downstream sheet "' +
               downstream.sheetName + '" not found.');
    return;
  }

  var data = downstreamSheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][downstream.watchCol]).trim() !==
        String(downstream.triggerValue).trim()) continue;

    var rowIndex = i + 1;

    if (getExistingUnresolvedToken(spreadsheet, rowIndex, downstream.sheetName)) {
      Logger.log('fireCascadeIfConfigured: row ' + rowIndex + ' already has alert, skipping.');
      continue;
    }
    if (getCascadeTokenForRow(spreadsheet, downstream.sheetName, rowIndex)) {
      Logger.log('fireCascadeIfConfigured: row ' + rowIndex + ' already cascaded, skipping.');
      continue;
    }

    var cascadeRowData = {
      clientName:     downstream.nameCol >= 0 ? data[i][downstream.nameCol] : 'Unknown',
      status:         data[i][downstream.watchCol],
      email:          downstream.notifyEmail   || '',
      rowIndex:       rowIndex,
      spreadsheetId:  spreadsheetId,
      sheetName:      downstream.sheetName,
      token:          generateUUID(),
      cascadeFrom:    resolvedRowData.sheetName,
      cascadeMessage: downstream.cascadeMessage || 'Updated by ' + resolvedRowData.sheetName
    };

    var slackConnected = !!PropertiesService.getScriptProperties()
                            .getProperty('SLACK_TOKEN_' + spreadsheetId);

    var slackResult = (slackConnected && downstream.notifySlackUser)
      ? sendCascadeSlackMessage(cascadeRowData, downstream.notifySlackUser)
      : false;

    var emailResult = cascadeRowData.email ? sendGmailAlert(cascadeRowData) : false;

    logAlert(spreadsheet, cascadeRowData, emailResult, slackResult,
             downstream.sheetName, true);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────────────────────────────────────

function getExistingUnresolvedToken(spreadsheet, rowIndex, sheetName) {
  var logSheet = spreadsheet.getSheetByName('AlertsLog');
  if (!logSheet) return null;

  var data = logSheet.getDataRange().getValues();
  for (var i = data.length - 1; i > 0; i--) {
    if (data[i][11] === sheetName &&
        data[i][1]  === rowIndex  &&
        data[i][12] !== true      &&
        data[i][7]  !== true) {
      return data[i][4];
    }
  }
  return null;
}

function getCascadeTokenForRow(spreadsheet, sheetName, rowIndex) {
  var logSheet = spreadsheet.getSheetByName('AlertsLog');
  if (!logSheet) return null;

  var data = logSheet.getDataRange().getValues();
  for (var i = data.length - 1; i > 0; i--) {
    if (data[i][11] === sheetName &&
        data[i][1]  === rowIndex  &&
        data[i][12] === true      &&
        data[i][7]  !== true) {
      return data[i][4];
    }
  }
  return null;
}

function logAlert(spreadsheet, rowData, emailSent, slackSent, sheetName, isCascade) {
  var logSheet = spreadsheet.getSheetByName('AlertsLog') || _createLogSheet(spreadsheet);
  logSheet.appendRow([
    new Date(),
    rowData.rowIndex,
    rowData.clientName,
    rowData.status,
    rowData.token,
    emailSent,
    slackSent,
    false, '', '', '',
    sheetName           || '',
    isCascade           || false,
    rowData.cascadeFrom || '',
    '', '', ''
  ]);
}

function _createLogSheet(spreadsheet) {
  var logSheet = spreadsheet.insertSheet('AlertsLog');
  logSheet.appendRow([
    'Timestamp', 'RowIndex', 'Name/Editor', 'Condition/Type', 'Token',
    'EmailSent', 'SlackSent', 'Resolved', 'ResolvedAt', 'ResolvedBy',
    'Notes', 'SheetName', 'IsCascade', 'CascadeFrom',
    'EditColumn', 'OldValue', 'NewValue'
  ]);
  return logSheet;
}

function generateUUID() {
  return Utilities.getUuid();
}

function columnToLetter(col) {
  var letter = '';
  while (col > 0) {
    var rem = (col - 1) % 26;
    letter  = String.fromCharCode(65 + rem) + letter;
    col     = Math.floor((col - 1) / 26);
  }
  return letter;
}

// ─────────────────────────────────────────────────────────────────────────────
// DEBUG — run manually from the Apps Script editor to verify wiring
// ─────────────────────────────────────────────────────────────────────────────
function debugConfig() {
  var docProps  = PropertiesService.getDocumentProperties().getProperties();
  var sp        = PropertiesService.getScriptProperties();
  var ssId      = docProps['SPREADSHEET_ID'];

  Logger.log('=== DocumentProperties ===');
  Object.keys(docProps).sort().forEach(function(k) {
    var v = (k.indexOf('TOKEN') !== -1 || k.indexOf('SECRET') !== -1)
      ? '***hidden***' : docProps[k];
    Logger.log('  ' + k + ' = ' + v);
  });

  Logger.log('\n=== Key ScriptProperties ===');
  ['SPREADSHEET_ID', 'SLACK_CLIENT_ID', 'DEPLOYED_WEBAPP_URL'].forEach(function(k) {
    Logger.log('  ' + k + ' = ' + (sp.getProperty(k) || '(not set)'));
  });

  if (ssId) {
    Logger.log('\n=== Per-spreadsheet ScriptProperties ===');
    Logger.log('  SLACK_TOKEN_'   + ssId + ' = ' +
      (sp.getProperty('SLACK_TOKEN_' + ssId) ? 'SET ✅' : 'NOT SET ❌'));
    Logger.log('  SLACK_CHANNEL_' + ssId + ' = ' +
      (sp.getProperty('SLACK_CHANNEL_' + ssId) || '(not set)'));
    Logger.log('  INSTALLER_EMAIL_' + ssId + ' = ' +
      (sp.getProperty('INSTALLER_EMAIL_' + ssId) || '(not set)'));
  }

  Logger.log('\n=== Installed triggers ===');
  ScriptApp.getProjectTriggers().forEach(function(t) {
    Logger.log('  handler=' + t.getHandlerFunction() +
               ' type='    + t.getEventType() +
               ' sourceId=' + t.getTriggerSourceId());
  });
}