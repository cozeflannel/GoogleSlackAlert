function runDailyAlerts() {
  // Time-based triggers have no document context, so we must
  // read SPREADSHEET_ID from script properties first, then
  // get all other config from document properties via the spreadsheet.
  var scriptProps   = PropertiesService.getScriptProperties();
  var spreadsheetId = scriptProps.getProperty('SPREADSHEET_ID');

  // Fallback: if not in script props, check if we have an active spreadsheet
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
    if (data[i][statusCol] == triggerValue) {
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

      var existingToken = getExistingUnresolvedToken(
        spreadsheet, rowData.rowIndex, sheetName
      );
      if (existingToken) continue;

      var cascadeToken = getCascadeTokenForRow(
        spreadsheet, sheetName, rowData.rowIndex
      );
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

      var ownNotification = getExistingUnresolvedToken(
        spreadsheet, rowIndex, downstream.sheetName
      );
      if (ownNotification) {
        Logger.log('Row ' + rowIndex + ' on "' + downstream.sheetName +
          '" already has its own notification. Skipping cascade.');
        continue;
      }

      var priorCascade = getCascadeTokenForRow(
        spreadsheet, downstream.sheetName, rowIndex
      );
      if (priorCascade) {
        Logger.log('Row ' + rowIndex + ' on "' + downstream.sheetName +
          '" already notified via prior cascade. Skipping.');
        continue;
      }

      var cascadeRowData = {
        clientName:     downstream.nameCol >= 0
                          ? data[i][downstream.nameCol] : 'Unknown',
        status:         data[i][downstream.watchCol],
        email:          downstream.notifyEmail   || '',
        rowIndex:       rowIndex,
        spreadsheetId:  spreadsheetId,
        sheetName:      downstream.sheetName,
        token:          generateUUID(),
        cascadeFrom:    resolvedRowData.sheetName,
        cascadeMessage: downstream.cascadeMessage ||
                        'Updated by ' + resolvedRowData.sheetName
      };

      var slackConnected = !!PropertiesService.getScriptProperties()
                              .getProperty('SLACK_TOKEN_' + spreadsheetId);

      var slackResult = (slackConnected && downstream.notifySlackUser)
        ? sendCascadeSlackMessage(cascadeRowData, downstream.notifySlackUser)
        : false;

      var emailResult = cascadeRowData.email
        ? sendGmailAlert(cascadeRowData) : false;

      logAlert(spreadsheet, cascadeRowData, emailResult, slackResult,
               downstream.sheetName, true);
    }
  }
}

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
  var logSheet = spreadsheet.getSheetByName('AlertsLog');
  if (!logSheet) {
    logSheet = spreadsheet.insertSheet('AlertsLog');
    logSheet.appendRow([
      'Timestamp', 'RowIndex', 'Name', 'Condition', 'Token',
      'EmailSent', 'SlackSent', 'Resolved', 'ResolvedAt', 'ResolvedBy',
      'Notes', 'SheetName', 'IsCascade', 'CascadeFrom'
    ]);
  }

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
    rowData.cascadeFrom || ''
  ]);
}

function generateUUID() {
  return Utilities.getUuid();
}

function createDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'runDailyAlerts') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('runDailyAlerts')
    .timeBased().everyDays(1).atHour(8).create();
  Logger.log('Daily trigger created.');
}