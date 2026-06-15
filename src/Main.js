// ─────────────────────────────────────────────────────────────────────────────
// TRIGGER HANDLERS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Installable onEdit trigger handler.
 * Fires when a user manually edits any cell in the spreadsheet.
 * If the editor is NOT the add-on installer → send edit notification.
 * If the editor IS the installer → silently skip.
 */
function onSheetEdit(e) {
  try {
    var editorEmail = e.user ? e.user.email : '';
    if (!editorEmail) return; // anonymous edit — skip

    var docProps       = PropertiesService.getDocumentProperties();
    var installerEmail = docProps.getProperty('INSTALLER_EMAIL') || '';
    var spreadsheetId  = docProps.getProperty('SPREADSHEET_ID');

    // Installer editing their own sheet — no notification
    if (installerEmail && editorEmail.toLowerCase() === installerEmail.toLowerCase()) return;

    var range      = e.range;
    var sheetName  = range.getSheet().getName();

    // Skip the internal AlertsLog sheet
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

    var scriptProps   = PropertiesService.getScriptProperties();
    var slackConnected = spreadsheetId
      ? !!scriptProps.getProperty('SLACK_TOKEN_' + spreadsheetId)
      : false;

    var slackResult = slackConnected ? sendEditNotification(editData) : false;

    // Email notification to installer
    var emailResult = false;
    if (installerEmail) {
      emailResult = sendGmailEditAlert(editData, installerEmail);
    }

    // Log to AlertsLog
    var ss       = e.source;
    var logSheet = ss.getSheetByName('AlertsLog') ||
                   _createLogSheet(ss);
    logSheet.appendRow([
      new Date(),
      editData.row,
      editorEmail,
      'Manual Edit',
      generateUUID(),
      emailResult,
      slackResult,
      false,
      '',
      '',
      '',
      sheetName,
      false,
      '',
      editData.columnLetter,
      editData.oldValue,
      editData.newValue
    ]);

  } catch (err) {
    Logger.log('onSheetEdit error: ' + err.toString());
  }
}

/**
 * Installable onChange trigger handler.
 * Fires when any cell value changes — including formula recalculations.
 * Runs the same condition check as the weekly job so formula-driven status
 * flips (e.g. "Late" computed from a due-date formula) fire alerts immediately.
 */
function onSheetChange(e) {
  try {
    // onChange gives us a change type but no range/value details —
    // we re-scan the configured status column for newly matching rows.
    runConditionCheck();
  } catch (err) {
    Logger.log('onSheetChange error: ' + err.toString());
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// CONDITION CHECK (was runDailyAlerts)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Scans the monitored sheet for rows matching the configured trigger condition
 * and sends Slack + email alerts for any newly matching rows.
 * Called by: onSheetChange (real-time) and sendWeeklyDigest (weekly).
 */
function runConditionCheck() {
  var scriptProps   = PropertiesService.getScriptProperties();
  var spreadsheetId = scriptProps.getProperty('SPREADSHEET_ID');

  if (!spreadsheetId) {
    try {
      var active = SpreadsheetApp.getActiveSpreadsheet();
      if (active) spreadsheetId = active.getId();
    } catch(e) {
      Logger.log('No spreadsheet context available.');
    }
  }

  if (!spreadsheetId) {
    Logger.log('Not configured. Open SheetAlerts sidebar and save configuration.');
    return;
  }

  var docProps  = PropertiesService.getDocumentProperties();
  var sheetName = docProps.getProperty('SHEET_NAME');

  if (!sheetName) {
    Logger.log('SHEET_NAME not set. Open SheetAlerts sidebar and save configuration.');
    return;
  }

  var nameCol      = parseInt(docProps.getProperty('NAME_COL')      || '-1');
  var emailCol     = parseInt(docProps.getProperty('EMAIL_COL')      || '-1');
  var extraInfoCol = parseInt(docProps.getProperty('EXTRA_INFO_COL') || '-1');
  var statusCol    = parseInt(docProps.getProperty('STATUS_COL')     || '-1');
  var triggerValue = docProps.getProperty('TRIGGER_VALUE');

  if (statusCol === -1 || !triggerValue) {
    Logger.log('Alert condition not configured.');
    return;
  }

  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet       = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log('Sheet "' + sheetName + '" not found. Was it renamed or deleted?');
    return;
  }

  var data           = sheet.getDataRange().getValues();
  var slackConnected = !!scriptProps.getProperty('SLACK_TOKEN_' + spreadsheetId);

  var pendingAlertsRaw = docProps.getProperty('PENDING_ALERTS');
  var pendingAlerts    = pendingAlertsRaw ? JSON.parse(pendingAlertsRaw) : [];

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][statusCol]).trim() === String(triggerValue).trim()) {
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

      var existingToken = getExistingUnresolvedToken(spreadsheet, rowData.rowIndex, sheetName);
      if (existingToken) continue;

      var cascadeToken = getCascadeTokenForRow(spreadsheet, sheetName, rowData.rowIndex);
      if (cascadeToken) {
        Logger.log('Row ' + rowData.rowIndex + ' already notified via cascade. Skipping.');
        continue;
      }

      rowData.token = generateUUID();

      var slackResult = slackConnected ? sendSlackAlert(rowData) : false;
      var emailResult = rowData.email  ? sendGmailAlert(rowData) : false;

      logAlert(spreadsheet, rowData, emailResult, slackResult, sheetName, false);
      pendingAlerts.push(rowData);
    }
  }

  var alertsJson = JSON.stringify(pendingAlerts);
  docProps.setProperty('PENDING_ALERTS', alertsJson);
  scriptProps.setProperty('PENDING_ALERTS_' + spreadsheetId, alertsJson);
}

// Legacy name alias so any existing time-based trigger on runDailyAlerts still works
function runDailyAlerts() {
  runConditionCheck();
}

// ─────────────────────────────────────────────────────────────────────────────
// WEEKLY DIGEST
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Weekly summary trigger handler (Sunday 08:00).
 * Reads AlertsLog for the past 7 days and posts a digest to Slack.
 * Also runs the condition check so any newly-overdue rows get alerted.
 */
function sendWeeklyDigest() {
  var scriptProps   = PropertiesService.getScriptProperties();
  var spreadsheetId = scriptProps.getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) {
    Logger.log('sendWeeklyDigest: SPREADSHEET_ID not set.');
    return;
  }

  var slackConnected = !!scriptProps.getProperty('SLACK_TOKEN_' + spreadsheetId);

  // First run a fresh condition check to catch any new rows
  runConditionCheck();

  if (!slackConnected) return;

  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var logSheet    = spreadsheet.getSheetByName('AlertsLog');
  if (!logSheet) {
    sendWeeklyDigestSlack(spreadsheetId, {
      totalAlerts: 0, resolved: 0, pending: 0,
      manualEdits: 0, newEditors: []
    });
    return;
  }

  var cutoff  = new Date();
  cutoff.setDate(cutoff.getDate() - 7);

  var data        = logSheet.getDataRange().getValues();
  var totalAlerts = 0;
  var resolved    = 0;
  var pending     = 0;
  var manualEdits = 0;
  var editorSet   = {};

  for (var i = 1; i < data.length; i++) {
    var ts = data[i][0];
    if (!(ts instanceof Date) || ts < cutoff) continue;

    var changeType = data[i][3] || '';
    if (changeType === 'Manual Edit') {
      manualEdits++;
      var editor = data[i][2] || '';
      if (editor) editorSet[editor] = true;
    } else {
      totalAlerts++;
      if (data[i][7] === true) {
        resolved++;
      } else {
        pending++;
      }
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
    Logger.log('Downstream sheet "' + downstream.sheetName + '" not found.');
    return;
  }

  var data = downstreamSheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][downstream.watchCol] == downstream.triggerValue) {
      var rowIndex = i + 1;

      var ownNotification = getExistingUnresolvedToken(spreadsheet, rowIndex, downstream.sheetName);
      if (ownNotification) {
        Logger.log('Row ' + rowIndex + ' on "' + downstream.sheetName +
          '" already has its own notification. Skipping cascade.');
        continue;
      }

      var priorCascade = getCascadeTokenForRow(spreadsheet, downstream.sheetName, rowIndex);
      if (priorCascade) {
        Logger.log('Row ' + rowIndex + ' on "' + downstream.sheetName +
          '" already notified via prior cascade. Skipping.');
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

      logAlert(spreadsheet, cascadeRowData, emailResult, slackResult, downstream.sheetName, true);
    }
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

/**
 * Creates or appends to the AlertsLog sheet.
 * Columns 0–13 are the original set; 14–16 are new edit-specific columns.
 */
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
    false,
    '',
    '',
    '',
    sheetName           || '',
    isCascade           || false,
    rowData.cascadeFrom || '',
    '',   // columnLetter (edit-only)
    '',   // oldValue (edit-only)
    ''    // newValue (edit-only)
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

/**
 * Converts a 1-based column number to a letter (1→A, 26→Z, 27→AA, …).
 */
function columnToLetter(col) {
  var letter = '';
  while (col > 0) {
    var rem = (col - 1) % 26;
    letter  = String.fromCharCode(65 + rem) + letter;
    col     = Math.floor((col - 1) / 26);
  }
  return letter;
}