// ─────────────────────────────────────────────────────────────────────────────
// TRIGGER HANDLERS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Installable onEdit trigger handler.
 */
function onSheetEdit(e) {
  try {
    var ss            = e.source;
    var spreadsheetId = ss.getId();
    var range         = e.range;
    var sheet         = range.getSheet();
    var sheetName     = sheet.getName();

    if (sheetName === 'AlertsLog') return;

    // ── 1. Handle Auto-Timestamp Logic ──────────────────────────────────────
    var scriptProps = PropertiesService.getScriptProperties();
    var atTrigCol   = parseInt(scriptProps.getProperty('AT_TRIG_COL_' + spreadsheetId) || '-1');
    var atDateCol   = parseInt(scriptProps.getProperty('AT_COL_'      + spreadsheetId) || '-1');
    var atVal       = scriptProps.getProperty('AT_VAL_'               + spreadsheetId);

    if (atTrigCol >= 0 && atDateCol >= 0 && range.getColumn() === (atTrigCol + 1)) {
      var editValue = String(e.value).trim();
      var dateCell  = sheet.getRange(range.getRow(), atDateCol + 1);
      
      if (editValue === atVal) {
        if (dateCell.getValue() === '') {
          dateCell.setValue(new Date());
          Logger.log('onSheetEdit: Auto-timestamped row ' + range.getRow());
        }
      } else if (editValue === '') {
        dateCell.clearContent();
      }
    }

    // ── 2. Standard Alert Notification Logic (Skip for installer) ───────────
    var editorEmail = e.user ? e.user.email : '';
    if (!editorEmail) return;

    var _docPropsRaw   = PropertiesService.getDocumentProperties();
    var docProps       = _docPropsRaw || { getProperty: function() { return null; } };
    var installerEmail = docProps.getProperty('INSTALLER_EMAIL') || '';

    if (!installerEmail) {
      installerEmail = scriptProps.getProperty('INSTALLER_EMAIL_' + spreadsheetId) || '';
    }

    if (installerEmail && editorEmail.toLowerCase() === installerEmail.toLowerCase()) {
      return;
    }

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

    var slackConnected = !!scriptProps.getProperty('SLACK_TOKEN_' + spreadsheetId);

    var slackResult = false;
    if (slackConnected) {
      try { slackResult = sendEditNotification(editData); } catch (slackErr) {}
    }

    var emailResult = false;
    if (installerEmail) {
      try { emailResult = sendGmailEditAlert(editData, installerEmail); } catch (mailErr) {}
    }

    var logSheet = ss.getSheetByName('AlertsLog') || _createLogSheet(ss);
    logSheet.appendRow([
      new Date(), editData.row, editorEmail, 'Manual Edit', generateUUID(),
      emailResult, slackResult, false, '', '', '', sheetName, false, '',
      editData.columnLetter, editData.oldValue, editData.newValue
    ]);

  } catch (err) {
    Logger.log('onSheetEdit ERROR: ' + err.toString());
  }
}

/**
 * Installable onChange trigger handler.
 */
function onSheetChange(e) {
  try {
    runConditionCheck();
  } catch (err) {
    Logger.log('onSheetChange ERROR: ' + err.toString());
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// CONDITION CHECK
// ─────────────────────────────────────────────────────────────────────────────

function runConditionCheck() {
  var _docPropsRaw  = PropertiesService.getDocumentProperties();
  var docProps      = _docPropsRaw || { getProperty: function() { return null; } };
  var scriptProps   = PropertiesService.getScriptProperties();
  var spreadsheetId = docProps.getProperty('SPREADSHEET_ID') || scriptProps.getProperty('SPREADSHEET_ID');

  if (!spreadsheetId) {
    try {
      var active = SpreadsheetApp.getActiveSpreadsheet();
      if (active) spreadsheetId = active.getId();
    } catch (e) {}
  }

  if (!spreadsheetId) return;

  var sheetName      = docProps.getProperty('SHEET_NAME') || scriptProps.getProperty('SHEET_NAME_' + spreadsheetId);
  var nameCol        = parseInt(docProps.getProperty('NAME_COL') || scriptProps.getProperty('NAME_COL_' + spreadsheetId) || '-1');
  var emailCol       = parseInt(docProps.getProperty('EMAIL_COL') || scriptProps.getProperty('EMAIL_COL_' + spreadsheetId) || '-1');
  var statusCol      = parseInt(docProps.getProperty('STATUS_COL') || scriptProps.getProperty('STATUS_COL_' + spreadsheetId) || '-1');
  var triggerValue   = docProps.getProperty('TRIGGER_VALUE') || scriptProps.getProperty('TRIGGER_VALUE_' + spreadsheetId);
  var dueDateCol     = parseInt(docProps.getProperty('DUE_DATE_COL')     || scriptProps.getProperty('DUE_DATE_COL_'     + spreadsheetId) || '-1');
  var finalStatusCol = parseInt(docProps.getProperty('FINAL_STATUS_COL') || scriptProps.getProperty('FINAL_STATUS_COL_' + spreadsheetId) || '-1');
  var tsCol          = parseInt(docProps.getProperty('AUTO_TIMESTAMP_COL') || scriptProps.getProperty('AT_COL_'           + spreadsheetId) || '-1');

  var ss;
  try { ss = SpreadsheetApp.openById(spreadsheetId); } catch (e) { return; }
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;

  var slackConnected = !!scriptProps.getProperty('SLACK_TOKEN_' + spreadsheetId);
  var data           = sheet.getDataRange().getValues();
  var today          = new Date();
  
  var pendingAlertsRaw = docProps.getProperty('PENDING_ALERTS') || scriptProps.getProperty('PENDING_ALERTS_' + spreadsheetId);
  var pendingAlerts    = pendingAlertsRaw ? JSON.parse(pendingAlertsRaw) : [];

  for (var i = 1; i < data.length; i++) {
    var rowIndex  = i + 1;
    var cellValue = String(data[i][statusCol]).trim();
    var trigger   = String(triggerValue).trim();
    
    var isTriggerMatched = (statusCol >= 0 && cellValue === trigger);
    var isUrgent         = false;

    if (finalStatusCol >= 0) {
      var currentStatus = '';
      var dueDate       = (dueDateCol >= 0) ? data[i][dueDateCol] : null;
      var completionDate = (tsCol >= 0) ? data[i][tsCol] : null;

      if (completionDate instanceof Date) {
        if (dueDate instanceof Date) {
          currentStatus = (completionDate <= dueDate) ? 'On Time' : 'Late';
        } else {
          currentStatus = 'Completed';
        }
      } else {
        if (dueDate instanceof Date && today > dueDate && cellValue !== trigger) {
          currentStatus = 'Urgent Action Needed';
          isUrgent      = true;
        }
      }
      
      if (currentStatus !== String(data[i][finalStatusCol])) {
        sheet.getRange(rowIndex, finalStatusCol + 1).setValue(currentStatus);
      }
    }

    if (!isTriggerMatched && !isUrgent) continue;

    var rowData = {
      clientName:    nameCol >= 0 ? data[i][nameCol] : 'Unknown',
      status:        isUrgent ? 'Urgent Action Needed' : cellValue,
      email:         emailCol >= 0 ? data[i][emailCol] : '',
      rowIndex:      rowIndex,
      spreadsheetId: spreadsheetId,
      sheetName:     sheetName
    };

    var existingToken = getExistingUnresolvedToken(ss, rowIndex, sheetName);
    if (existingToken) continue;

    rowData.token = generateUUID();

    var slackResult = false;
    if (slackConnected) {
      try { slackResult = sendSlackAlert(rowData); } catch (e) {}
    }

    var emailResult = false;
    if (rowData.email) {
      try { emailResult = sendGmailAlert(rowData); } catch (e) {}
    }

    logAlert(ss, rowData, emailResult, slackResult, sheetName, false);
    pendingAlerts.push(rowData);
  }

  var alertsJson = JSON.stringify(pendingAlerts);
  if (_docPropsRaw) _docPropsRaw.setProperty('PENDING_ALERTS', alertsJson);
  scriptProps.setProperty('PENDING_ALERTS_' + spreadsheetId, alertsJson);
}

// Legacy alias
function runDailyAlerts() { runConditionCheck(); }

// ─────────────────────────────────────────────────────────────────────────────
// WEEKLY DIGEST
// ─────────────────────────────────────────────────────────────────────────────

function sendWeeklyDigest() {
  var _docPropsRaw  = PropertiesService.getDocumentProperties();
  var docProps      = _docPropsRaw || { getProperty: function() { return null; } };
  var spreadsheetId = docProps.getProperty('SPREADSHEET_ID');

  if (!spreadsheetId) return;

  var scriptProps    = PropertiesService.getScriptProperties();
  var slackConnected = !!scriptProps.getProperty('SLACK_TOKEN_' + spreadsheetId);

  runConditionCheck();

  if (!slackConnected) return;

  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var logSheet    = spreadsheet.getSheetByName('AlertsLog');

  if (!logSheet) {
    sendWeeklyDigestSlack(spreadsheetId, { totalAlerts: 0, resolved: 0, pending: 0, manualEdits: 0, newEditors: [] });
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
  var _docPropsRaw  = PropertiesService.getDocumentProperties();
  var docProps      = _docPropsRaw || { getProperty: function() { return null; } };
  var downstreamRaw = docProps.getProperty('DOWNSTREAM_CONFIG');
  if (!downstreamRaw) return;

  var downstream = JSON.parse(downstreamRaw);
  if (!downstream.enabled) return;

  var spreadsheet     = SpreadsheetApp.openById(spreadsheetId);
  var downstreamSheet = spreadsheet.getSheetByName(downstream.sheetName);
  if (!downstreamSheet) return;

  var data = downstreamSheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][downstream.watchCol]).trim() !== String(downstream.triggerValue).trim()) continue;

    var rowIndex = i + 1;
    if (getExistingUnresolvedToken(spreadsheet, rowIndex, downstream.sheetName)) continue;
    if (getCascadeTokenForRow(spreadsheet, downstream.sheetName, rowIndex)) continue;

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

    var slackConnected = !!PropertiesService.getScriptProperties().getProperty('SLACK_TOKEN_' + spreadsheetId);
    var slackResult = (slackConnected && downstream.notifySlackUser) ? sendCascadeSlackMessage(cascadeRowData, downstream.notifySlackUser) : false;
    var emailResult = cascadeRowData.email ? sendGmailAlert(cascadeRowData) : false;

    logAlert(spreadsheet, cascadeRowData, emailResult, slackResult, downstream.sheetName, true);
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
    if (data[i][11] === sheetName && data[i][1] === rowIndex && data[i][12] !== true && data[i][7] !== true) {
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
    if (data[i][11] === sheetName && data[i][1] === rowIndex && data[i][12] === true && data[i][7] !== true) {
      return data[i][4];
    }
  }
  return null;
}

function logAlert(spreadsheet, rowData, emailSent, slackSent, sheetName, isCascade) {
  var logSheet = spreadsheet.getSheetByName('AlertsLog') || _createLogSheet(spreadsheet);
  logSheet.appendRow([
    new Date(), rowData.rowIndex, rowData.clientName, rowData.status, rowData.token,
    emailSent, slackSent, false, '', '', '', sheetName || '', isCascade || false,
    rowData.cascadeFrom || '', '', '', ''
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

function generateUUID() { return Utilities.getUuid(); }

function columnToLetter(col) {
  var letter = '';
  while (col > 0) {
    var rem = (col - 1) % 26;
    letter  = String.fromCharCode(65 + rem) + letter;
    col     = Math.floor((col - 1) / 26);
  }
  return letter;
}

function debugConfig() {
  var _docPropsRaw = PropertiesService.getDocumentProperties();
  var docProps     = _docPropsRaw ? _docPropsRaw.getProperties() : {};
  var sp           = PropertiesService.getScriptProperties();
  var ssId         = docProps['SPREADSHEET_ID'];

  Logger.log('=== DocumentProperties ===');
  Object.keys(docProps).sort().forEach(function(k) {
    var v = (k.indexOf('TOKEN') !== -1 || k.indexOf('SECRET') !== -1) ? '***hidden***' : docProps[k];
    Logger.log('  ' + k + ' = ' + v);
  });

  Logger.log('\n=== Key ScriptProperties ===');
  ['SPREADSHEET_ID', 'SLACK_CLIENT_ID', 'DEPLOYED_WEBAPP_URL'].forEach(function(k) {
    Logger.log('  ' + k + ' = ' + (sp.getProperty(k) || '(not set)'));
  });

  if (ssId) {
    Logger.log('\n=== Per-spreadsheet ScriptProperties ===');
    Logger.log('  SLACK_TOKEN_'   + ssId + ' = ' + (sp.getProperty('SLACK_TOKEN_' + ssId) ? 'SET ✅' : 'NOT SET ❌'));
    Logger.log('  SLACK_CHANNEL_' + ssId + ' = ' + (sp.getProperty('SLACK_CHANNEL_' + ssId) || '(not set)'));
    Logger.log('  INSTALLER_EMAIL_' + ssId + ' = ' + (sp.getProperty('INSTALLER_EMAIL_' + ssId) || '(not set)'));
  }

  Logger.log('\n=== Installed triggers ===');
  ScriptApp.getProjectTriggers().forEach(function(t) {
    Logger.log('  handler=' + t.getHandlerFunction() + ' type=' + t.getEventType() + ' sourceId=' + t.getTriggerSourceId());
  });
}
