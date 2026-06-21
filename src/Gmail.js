function sendGmailAlert(rowData) {
  try {
    // The resolve link now points to the Supabase edge function, not the GAS web app
    var actionUrl = 'https://apjftvnmskckrhgrdpbk.supabase.co/functions/v1/bot' +
      '?action=resolve' +
      '&token='         + encodeURIComponent(rowData.token) +
      '&rowIndex='      + rowData.rowIndex +
      '&spreadsheetId=' + rowData.spreadsheetId;

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

    MailApp.sendEmail({ to: rowData.email, subject: subject, htmlBody: htmlBody });
    return true;
  } catch (err) {
    Logger.log("Failed to send Gmail alert: " + err.toString());
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
