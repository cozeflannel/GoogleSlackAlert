function sendGmailAlert(rowData) {
  var scriptUrl = ScriptApp.getService().getUrl();
  var actionUrl = scriptUrl +
    "?action=resolve" +
    "&token="         + encodeURIComponent(rowData.token) +
    "&rowIndex="      + rowData.rowIndex +
    "&spreadsheetId=" + rowData.spreadsheetId;

  var subject = "[Action Required] " + rowData.clientName + " — " + rowData.status;

  // Include cascade context in the email if this was triggered by an upstream sheet
  var cascadeNote = rowData.cascadeMessage
    ? "<p style='background:#fff8e1;border-left:4px solid #f9a825;padding:8px 12px;" +
      "margin-bottom:12px;'><b>Context:</b> " + rowData.cascadeMessage + "</p>"
    : "";

  var htmlBody =
    "<h2>Alert: Action Required</h2>" +
    cascadeNote +
    "<p>The following item requires your attention:</p>" +
    "<table border='1' cellpadding='5' style='border-collapse:collapse;'>" +
    "<tr><td><b>Sheet</b></td><td>"       + (rowData.sheetName  || '') + "</td></tr>" +
    "<tr><td><b>Row</b></td><td>"         + rowData.rowIndex           + "</td></tr>" +
    "<tr><td><b>Name</b></td><td>"        + rowData.clientName         + "</td></tr>" +
    "<tr><td><b>Status</b></td><td>"      + rowData.status             + "</td></tr>" +
    "</table><br/>" +
    "<a href='" + actionUrl + "' style='background-color:#0052cc;color:white;" +
    "padding:10px 20px;text-decoration:none;border-radius:5px;'>Take Action</a>";

  try {
    MailApp.sendEmail({ to: rowData.email, subject: subject, htmlBody: htmlBody });
    return true;
  } catch (err) {
    Logger.log("Failed to send email: " + err.toString());
    return false;
  }
}

/**
 * Sends an email to the installer (add-on owner) when another user
 * manually edits the spreadsheet.
 *
 * @param {Object} editData  - edit context from onSheetEdit
 * @param {string} toEmail   - installer's email address
 */
function sendGmailEditAlert(editData, toEmail) {
  var sheetUrl = 'https://docs.google.com/spreadsheets/d/' +
                 editData.spreadsheetId + '/edit';
  var subject  = '[SheetAlerts] Edit by ' + editData.editorEmail +
                 ' on ' + editData.sheetName;

  var htmlBody =
    '<h2 style="color:#333;">✏️ Sheet Edit Notification</h2>' +
    '<p>Someone edited your spreadsheet. Here are the details:</p>' +
    '<table border="1" cellpadding="8" style="border-collapse:collapse;font-size:14px;">' +
    '<tr><td><b>Sheet</b></td><td>'    + editData.sheetName    + '</td></tr>' +
    '<tr><td><b>Editor</b></td><td>'   + editData.editorEmail  + '</td></tr>' +
    '<tr><td><b>Cell</b></td><td>'     + editData.columnLetter + editData.row + '</td></tr>' +
    '<tr><td><b>Old Value</b></td><td>' + editData.oldValue    + '</td></tr>' +
    '<tr><td><b>New Value</b></td><td>' + editData.newValue    + '</td></tr>' +
    '<tr><td><b>Time</b></td><td>'     + editData.timestamp.toLocaleString() + '</td></tr>' +
    '</table><br/>' +
    '<a href="' + sheetUrl + '" style="background-color:#0052cc;color:white;' +
    'padding:10px 20px;text-decoration:none;border-radius:5px;">View Spreadsheet</a>';

  try {
    MailApp.sendEmail({ to: toEmail, subject: subject, htmlBody: htmlBody });
    return true;
  } catch (err) {
    Logger.log('sendGmailEditAlert failed: ' + err.toString());
    return false;
  }
}

function getActionFormHtml(rowData) {
  var html =
    "<html><head><title>Resolve Alert</title>" +
    "<style>body{font-family:sans-serif;padding:20px;max-width:500px;margin:auto;}</style>" +
    "</head><body>" +
    "<h2>Resolve Alert: " + rowData.clientName + "</h2>" +
    "<p>Current Status: <b>" + rowData.status + "</b></p>" +
    (rowData.cascadeMessage
      ? "<p style='background:#fff8e1;border-left:4px solid #f9a825;padding:8px 12px;'>" +
        rowData.cascadeMessage + "</p>"
      : "") +
    "<div id='formContainer'>" +
    "<label>New Status:</label><br/>" +
    "<select id='status' style='margin-bottom:15px;padding:5px;width:100%;'>" +
    "<option value='Paid'>Paid</option>" +
    "<option value='Resolved'>Resolved</option>" +
    "<option value='Ignored'>Ignored</option>" +
    "</select><br/>" +
    "<label>Notes:</label><br/>" +
    "<textarea id='notes' rows='4' " +
    "style='margin-bottom:15px;padding:5px;width:100%;'></textarea><br/>" +
    "<button onclick='submitForm()' " +
    "style='padding:10px 20px;background-color:#0052cc;color:white;" +
    "border:none;cursor:pointer;'>Submit</button>" +
    "</div>" +
    "<div id='message' " +
    "style='display:none;color:green;font-weight:bold;margin-top:20px;'></div>" +
    "<script>" +
    "function submitForm(){" +
    "  document.getElementById('formContainer').style.display='none';" +
    "  document.getElementById('message').innerText='Submitting...';" +
    "  document.getElementById('message').style.display='block';" +
    "  var status=document.getElementById('status').value;" +
    "  var notes=document.getElementById('notes').value;" +
    "  google.script.run.withSuccessHandler(function(res){" +
    "    if(res.success){" +
    "      document.getElementById('message').innerText='Row successfully updated!';}" +
    "    else{document.getElementById('message').innerText='Error: '+res.error;" +
    "         document.getElementById('message').style.color='red';}" +
    "  }).resolveAlert('" + rowData.token + "'," +
    rowData.rowIndex + ",status,notes,'" + rowData.spreadsheetId + "');" +
    "}" +
    "<\/script></body></html>";

  return HtmlService.createHtmlOutput(html);
}

function resolveAlert(token, rowIndex, newStatus, notes, providedSpreadsheetId) {
  var docProps      = PropertiesService.getDocumentProperties();
  var spreadsheetId = providedSpreadsheetId || docProps.getProperty('SPREADSHEET_ID');
  var sheetName     = docProps.getProperty('SHEET_NAME');
  var statusCol     = parseInt(docProps.getProperty('STATUS_COL') || '-1') + 1;

  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet       = spreadsheet.getSheetByName(sheetName);
  var logSheet    = spreadsheet.getSheetByName('AlertsLog');

  if (!logSheet || !sheet) {
    return { success: false, error: 'Sheet not found.' };
  }

  var logData     = logSheet.getDataRange().getValues();
  var tokenValid  = false;
  var logRowIndex = -1;
  var resolvedRowData = {};

  for (var i = 1; i < logData.length; i++) {
    if (logData[i][4] === token &&
        logData[i][1] === rowIndex &&
        logData[i][7] !== true) {
      tokenValid  = true;
      logRowIndex = i + 1;
      resolvedRowData = {
        sheetName:  logData[i][11] || sheetName,
        rowIndex:   rowIndex,
        clientName: logData[i][2]
      };
      break;
    }
  }

  if (!tokenValid) {
    return { success: false, error: 'Invalid or already resolved token.' };
  }

  // Update the monitored sheet
  sheet.getRange(rowIndex, statusCol).setValue(newStatus);

  // Update the log
  logSheet.getRange(logRowIndex, 8).setValue(true);
  logSheet.getRange(logRowIndex, 9).setValue(new Date());
  logSheet.getRange(logRowIndex, 10)
    .setValue(Session.getActiveUser().getEmail() || 'Slack/Web User');
  logSheet.getRange(logRowIndex, 11).setValue(notes || '');

  // Update pending alerts in both property stores
  var pendingAlertsRaw = docProps.getProperty('PENDING_ALERTS');
  if (pendingAlertsRaw) {
    var updated     = JSON.parse(pendingAlertsRaw)
                        .filter(function(a) { return a.token !== token; });
    var updatedJson = JSON.stringify(updated);
    docProps.setProperty('PENDING_ALERTS', updatedJson);
    PropertiesService.getScriptProperties()
      .setProperty('PENDING_ALERTS_' + spreadsheetId, updatedJson);
  }

  // Fire cascade to downstream sheet if configured
  fireCascadeIfConfigured(spreadsheetId, resolvedRowData);

  return { success: true };
}