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

    // ── 1. Handle Auto-Timestamp Logic ──────────────────────────────────────
    var docProps    = PropertiesService.getDocumentProperties();
    var atTrigCol   = parseInt(docProps.getProperty('AT_TRIG_COL') || '-1');
    var atDateCol   = parseInt(docProps.getProperty('AT_COL')      || '-1');
    var atVal       = docProps.getProperty('AT_VAL');

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

    var installerEmail = docProps.getProperty('INSTALLER_EMAIL') || '';

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

    var scriptProps = PropertiesService.getScriptProperties();
    var slackConnected = !!scriptProps.getProperty('SLACK_TOKEN_' + spreadsheetId);

    var slackResult = false;
    if (slackConnected) {
      try { slackResult = sendEditNotification(editData); } catch (slackErr) {}
    }

    var emailResult = false;
    if (installerEmail) {
      try { emailResult = sendGmailEditAlert(editData, installerEmail); } catch (mailErr) {}
    }

    // Manual edits are no longer logged to the alerts table (which is for condition alerts).
    // They remain as transient notifications.

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
  var docProps      = PropertiesService.getDocumentProperties();
  var scriptProps   = PropertiesService.getScriptProperties();
  var spreadsheetId = docProps.getProperty('SPREADSHEET_ID');

  if (!spreadsheetId) {
    try {
      var active = SpreadsheetApp.getActiveSpreadsheet();
      if (active) spreadsheetId = active.getId();
    } catch (e) {}
  }

  if (!spreadsheetId) return;

  var sheetName      = docProps.getProperty('SHEET_NAME');
  var nameCol        = parseInt(docProps.getProperty('NAME_COL') || '-1');
  var emailCol       = parseInt(docProps.getProperty('EMAIL_COL') || '-1');
  var statusCol      = parseInt(docProps.getProperty('STATUS_COL') || '-1');
  var triggerValue   = docProps.getProperty('TRIGGER_VALUE');
  var dueDateCol     = parseInt(docProps.getProperty('DUE_DATE_COL')     || '-1');
  var finalStatusCol = parseInt(docProps.getProperty('FINAL_STATUS_COL') || '-1');
  var tsCol          = parseInt(docProps.getProperty('AUTO_TIMESTAMP_COL') || '-1');

  var ss;
  try { ss = SpreadsheetApp.openById(spreadsheetId); } catch (e) { return; }
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;

  var slackConnected = !!scriptProps.getProperty('SLACK_TOKEN_' + spreadsheetId);
  var data           = sheet.getDataRange().getValues();
  var today          = new Date();
  
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
  }
}

// Legacy alias
function runDailyAlerts() { runConditionCheck(); }

// ─────────────────────────────────────────────────────────────────────────────
// WEEKLY DIGEST
// ─────────────────────────────────────────────────────────────────────────────

function sendWeeklyDigest() {
  var docProps      = PropertiesService.getDocumentProperties();
  var spreadsheetId = docProps.getProperty('SPREADSHEET_ID');

  if (!spreadsheetId) return;

  var scriptProps    = PropertiesService.getScriptProperties();
  var slackConnected = !!scriptProps.getProperty('SLACK_TOKEN_' + spreadsheetId);

  runConditionCheck();

  if (!slackConnected) return;

  // Query Supabase for alert stats for the last 7 days
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 7);
  var cutoffIso = cutoff.toISOString();

  var supabaseUrl = scriptProps.getProperty('SUPABASE_URL') + 
                    '/rest/v1/alerts?spreadsheet_id=eq.' + spreadsheetId + 
                    '&created_at=gte.' + cutoffIso + '&select=*';
  var serviceKey  = scriptProps.getProperty('SUP_SECRET_KEY');

  try {
    var response = UrlFetchApp.fetch(supabaseUrl, {
      headers: { 'apikey': serviceKey, 'Authorization': 'Bearer ' + serviceKey },
      muteHttpExceptions: true
    });
    var alerts = JSON.parse(response.getContentText());

    var totalAlerts = 0, resolved = 0, pending = 0;
    alerts.forEach(function(a) {
      totalAlerts++;
      if (a.resolved) resolved++;
      else pending++;
    });

    sendWeeklyDigestSlack(spreadsheetId, {
      totalAlerts: totalAlerts,
      resolved:    resolved,
      pending:     pending,
      manualEdits: 0, // Manual edits are no longer in the alerts table
      newEditors:  []
    });
  } catch (e) {
    Logger.log('sendWeeklyDigest Supabase error: ' + e.toString());
  }
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
  var spreadsheetId = spreadsheet.getId();
  var scriptProps = PropertiesService.getScriptProperties();
  var supabaseUrl = scriptProps.getProperty('SUPABASE_URL') + 
                    '/rest/v1/alerts?spreadsheet_id=eq.' + spreadsheetId + 
                    '&sheet_name=eq.' + encodeURIComponent(sheetName) + 
                    '&row_index=eq.' + rowIndex + 
                    '&resolved=eq.false&select=token';
  var serviceKey  = scriptProps.getProperty('SUP_SECRET_KEY');

  try {
    var response = UrlFetchApp.fetch(supabaseUrl, {
      headers: { 'apikey': serviceKey, 'Authorization': 'Bearer ' + serviceKey },
      muteHttpExceptions: true
    });
    var data = JSON.parse(response.getContentText());
    if (data && data.length > 0) return data[0].token;
  } catch (e) {
    Logger.log('getExistingUnresolvedToken Supabase error: ' + e.toString());
  }
  return null;
}

function getCascadeTokenForRow(spreadsheet, sheetName, rowIndex) {
  var spreadsheetId = spreadsheet.getId();
  var scriptProps = PropertiesService.getScriptProperties();
  var supabaseUrl = scriptProps.getProperty('SUPABASE_URL') + 
                    '/rest/v1/alerts?spreadsheet_id=eq.' + spreadsheetId + 
                    '&sheet_name=eq.' + encodeURIComponent(sheetName) + 
                    '&row_index=eq.' + rowIndex + 
                    '&resolved=eq.false&select=token';
  var serviceKey  = scriptProps.getProperty('SUP_SECRET_KEY');

  try {
    var response = UrlFetchApp.fetch(supabaseUrl, {
      headers: { 'apikey': serviceKey, 'Authorization': 'Bearer ' + serviceKey },
      muteHtttpExceptions: true
    });
    var data = JSON.parse(response.getContentText());
    if (data && data.length > 0) return data[0].token;
  } catch (e) {
    Logger.log('getCascadeTokenForRow Supabase error: ' + e.toString());
  }
  return null;
}

function logAlert(spreadsheet, rowData, emailSent, slackSent, sheetName, isCascade) {
  var spreadsheetId = spreadsheet.getId();
  var scriptProps = PropertiesService.getScriptProperties();
  var supabaseUrl = scriptProps.getProperty('SUPABASE_URL') + '/rest/v1/alerts';
  var serviceKey  = scriptProps.getProperty('SUP_SECRET_KEY');

  var payload = {
    spreadsheet_id: spreadsheetId,
    sheet_name:     sheetName || rowData.sheetName,
    row_index:      rowData.rowIndex,
    client_name:    rowData.clientName,
    status:         rowData.status,
    token:          rowData.token,
    email_sent:     emailSent,
    slack_sent:     slackSent,
    resolved:       false,
    is_cascade:     isCascade || false,
    cascade_from:   rowData.cascadeFrom || ''
  };

  try {
    UrlFetchApp.fetch(supabaseUrl, {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'apikey': serviceKey,
        'Authorization': 'Bearer ' + serviceKey
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (err) {
    Logger.log('logAlert Supabase CRITICAL: ' + err.toString());
  }
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

  Logger.log('
=== Key ScriptProperties ===');
  ['SPREADSHEET_ID', 'SLACK_CLIENT_ID', 'DEPLOYED_WEBAPP_URL'].forEach(function(k) {
    Logger.log('  ' + k + ' = ' + (sp.getProperty(k) || '(not set)'));
  });

  if (ssId) {
    Logger.log('
=== Per-spreadsheet ScriptProperties ===');
    ['SLACK_TOKEN_'   + ssId, 'SLACK_CHANNEL_' + ssId, 'INSTALLER_EMAIL_' + ssId].forEach(function(k) {
      Logger.log('  ' + k + ' = ' + (sp.getProperty(k) ? 'SET ✅' : 'NOT SET ❌'));
    });
  }

  Logger.log('
=== Installed triggers ===');
  ScriptApp.getProjectTriggers().forEach(function(t) {
    Logger.log('  handler=' + t.getHandlerFunction() + ' type=' + t.getEventType() + ' sourceId=' + t.getTriggerSourceId());
  });
}
