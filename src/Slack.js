function _getToken(spreadsheetId) {
  if (!spreadsheetId) {
    throw new Error("spreadsheetId is required to look up the Slack token.");
  }
  var token = PropertiesService.getScriptProperties()
                .getProperty("SLACK_TOKEN_" + spreadsheetId);
  if (!token) {
    throw new Error(
      "No Slack token found for spreadsheet " + spreadsheetId +
      ". Complete the OAuth install flow first."
    );
  }
  return token;
}

function sendSlackAlert(rowData) {
  var token = _getToken(rowData.spreadsheetId);
  var scriptProps = PropertiesService.getScriptProperties();
  var channel = scriptProps.getProperty('SLACK_CHANNEL_' + rowData.spreadsheetId) || '#general';
  var sheetUrl = "https://docs.google.com/spreadsheets/d/" + rowData.spreadsheetId + "/edit#gid=0";
  var url = "https://slack.com/api/chat.postMessage";

  var payload = {
    "channel": channel,
    "blocks": [
      {
        "type": "header",
        "text": { "type": "plain_text", "text": "🚨 Action Required: " + rowData.clientName }
      },
      {
        "type": "section",
        "fields": [
          { "type": "mrkdwn", "text": "*Row Index:*\n" + rowData.rowIndex },
          { "type": "mrkdwn", "text": "*Status:*\n" + rowData.status },
          { "type": "mrkdwn", "text": "*Email:*\n" + (rowData.email || "N/A") }
        ]
      },
      {
        "type": "actions",
        "elements": [
          {
            "type": "button",
            "text": { "type": "plain_text", "text": "✅ Take Action" },
            "style": "primary",
            "action_id": "open_action_modal",
            "value": JSON.stringify({
              token: rowData.token,
              rowIndex: rowData.rowIndex,
              spreadsheetId: rowData.spreadsheetId
            })
          },
          {
            "type": "button",
            "text": { "type": "plain_text", "text": "🔗 View in Sheet" },
            "url": sheetUrl
          }
        ]
      }
    ]
  };

  var response = UrlFetchApp.fetch(url, {
    "method": "post",
    "headers": { "Authorization": "Bearer " + token },
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  });

  var result = JSON.parse(response.getContentText());
  Logger.log("Slack response: " + response.getContentText());
  return result.ok;
}

function openActionModal(triggerId, actionValue) {
  var parsedValue = JSON.parse(actionValue);
  var token = _getToken(parsedValue.spreadsheetId);
  var url = "https://slack.com/api/views.open";
  var rowIndex = parsedValue.rowIndex;
  var actionToken = parsedValue.token;

  var payload = {
    "trigger_id": triggerId,
    "view": {
      "type": "modal",
      "callback_id": "resolve_alert_modal",
      "private_metadata": actionValue,
      "title": { "type": "plain_text", "text": "Resolve Alert" },
      "submit": { "type": "plain_text", "text": "Submit" },
      "close": { "type": "plain_text", "text": "Cancel" },
      "blocks": [
        {
          "type": "section",
          "text": { "type": "mrkdwn", "text": "You are updating row *" + rowIndex + "*." }
        },
        {
          "type": "input",
          "block_id": "status_block",
          "element": {
            "type": "static_select",
            "action_id": "status_select",
            "options": [
              { "text": { "type": "plain_text", "text": "Paid" }, "value": "Paid" },
              { "text": { "type": "plain_text", "text": "Resolved" }, "value": "Resolved" },
              { "text": { "type": "plain_text", "text": "Ignored" }, "value": "Ignored" }
            ]
          },
          "label": { "type": "plain_text", "text": "New Status" }
        },
        {
          "type": "input",
          "block_id": "notes_block",
          "optional": true,
          "element": {
            "type": "plain_text_input",
            "action_id": "notes_input",
            "multiline": true
          },
          "label": { "type": "plain_text", "text": "Notes" }
        }
      ]
    }
  };

  UrlFetchApp.fetch(url, {
    "method": "post",
    "headers": { "Authorization": "Bearer " + token },
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  });
}

function publishAppHome(userId, spreadsheetId) {
  var token = _getToken(spreadsheetId);
  var scriptProps = PropertiesService.getScriptProperties();
  var pendingAlertsRaw = scriptProps.getProperty('PENDING_ALERTS_' + spreadsheetId);
  var pendingAlerts = pendingAlertsRaw ? JSON.parse(pendingAlertsRaw) : [];

  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var logSheet = spreadsheet.getSheetByName('AlertsLog');
  var historyBlocks = [];

  if (logSheet) {
    var data = logSheet.getDataRange().getValues();
    var historyCount = 0;
    for (var i = data.length - 1; i > 0 && historyCount < 20; i--) {
      if (data[i][7] === true) {
        historyCount++;
        historyBlocks.push({
          "type": "section",
          "text": {
            "type": "mrkdwn",
            "text": "✅ *Row " + data[i][1] + "* - " + data[i][2] +
                    " (Status: " + data[i][3] + ")\nResolved by: " + data[i][9] +
                    "\nNotes: " + data[i][10]
          }
        });
        historyBlocks.push({ "type": "divider" });
      }
    }
  }

  var blocks = [
    {
      "type": "header",
      "text": { "type": "plain_text", "text": "SheetAlerts Dashboard" }
    },
    {
      "type": "section",
      "text": { "type": "mrkdwn", "text": "Here are your pending operational tasks and recent history." }
    },
    { "type": "divider" },
    {
      "type": "header",
      "text": { "type": "plain_text", "text": "🔴 Active Alerts" }
    }
  ];

  if (pendingAlerts.length === 0) {
    blocks.push({
      "type": "section",
      "text": { "type": "mrkdwn", "text": "No active alerts! 🎉" }
    });
  } else {
    pendingAlerts.forEach(function(alert) {
      blocks.push({
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": "*Row " + alert.rowIndex + ":* " + alert.clientName +
                  " (Status: " + alert.status + ")"
        },
        "accessory": {
          "type": "button",
          "text": { "type": "plain_text", "text": "Take Action" },
          "action_id": "open_action_modal",
          "value": JSON.stringify({
            token: alert.token,
            rowIndex: alert.rowIndex,
            spreadsheetId: alert.spreadsheetId
          })
        }
      });
    });
  }

  blocks.push({ "type": "divider" });
  blocks.push({
    "type": "header",
    "text": { "type": "plain_text", "text": "📜 History (Last 20)" }
  });

  if (historyBlocks.length === 0) {
    blocks.push({
      "type": "section",
      "text": { "type": "mrkdwn", "text": "No history available." }
    });
  } else {
    blocks = blocks.concat(historyBlocks);
  }

  var payload = {
    "user_id": userId,
    "view": {
      "type": "home",
      "blocks": blocks
    }
  };

  UrlFetchApp.fetch("https://slack.com/api/views.publish", {
    "method": "post",
    "headers": { "Authorization": "Bearer " + token },
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  });
}