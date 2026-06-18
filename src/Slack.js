// ─────────────────────────────────────────────────────────────────────────────
// TOKEN HELPER
// ─────────────────────────────────────────────────────────────────────────────

function _getToken(spreadsheetId) {
  if (!spreadsheetId) {
    throw new Error('spreadsheetId is required to look up the Slack token.');
  }
  var token = PropertiesService.getScriptProperties()
                .getProperty('SLACK_TOKEN_' + spreadsheetId);
  if (!token) {
    throw new Error(
      'No Slack token found for spreadsheet ' + spreadsheetId +
      '. Complete the OAuth install flow first.'
    );
  }
  return token;
}

function _getChannel(spreadsheetId) {
  return PropertiesService.getScriptProperties()
           .getProperty('SLACK_CHANNEL_' + spreadsheetId) || '#general';
}

// ─────────────────────────────────────────────────────────────────────────────
// CONDITION ALERT (triggered by formula / daily check)
// Rich block message with "Take Action" modal button + "View in Sheet" link
// ─────────────────────────────────────────────────────────────────────────────

function getSlackChannels(spreadsheetId) {
  var token = _getToken(spreadsheetId);
  var result = _fetchSlack('https://slack.com/api/conversations.list?types=public_channel,private_channel', {
    method: 'get',
    headers: { Authorization: 'Bearer ' + token }
  });

  if (!result.ok) {
    Logger.log('Error fetching channels: ' + result.error);
    return [];
  }

  return result.channels.map(function(c) {
    return { name: '#' + c.name, id: c.id };
  });
}

function _fetchSlack(url, options) {
  try {
    var response = UrlFetchApp.fetch(url, options);
    var result   = JSON.parse(response.getContentText());
    if (!result.ok) {
      Logger.log('Slack API Error: ' + result.error + ' URL: ' + url);
    }
    return result;
  } catch (err) {
    Logger.log('Slack Fetch Exception: ' + err.toString() + ' URL: ' + url);
    return { ok: false, error: err.toString() };
  }
}

function sendSlackAlert(rowData) {
  var token    = _getToken(rowData.spreadsheetId);
  var channel  = _getChannel(rowData.spreadsheetId);
  var sheetUrl = 'https://docs.google.com/spreadsheets/d/' + rowData.spreadsheetId + '/edit#gid=0';

  var payload = {
    channel: channel,
    blocks: [
      {
        type: 'header',
        text: { type: 'plain_text', text: '🚨 Action Required: ' + rowData.clientName }
      },
      {
        type: 'section',
        fields: [
          { type: 'mrkdwn', text: '*Sheet:*\n'     + rowData.sheetName },
          { type: 'mrkdwn', text: '*Row:*\n'       + rowData.rowIndex  },
          { type: 'mrkdwn', text: '*Condition:*\n' + rowData.status    },
          { type: 'mrkdwn', text: '*Email:*\n'     + (rowData.email || 'N/A') }
        ]
      },
      {
        type: 'section',
        text: {
          type: 'mrkdwn',
          text: '*<' + sheetUrl + '|📄 Open Spreadsheet>*'
        }
      },
      {
        type: 'actions',
        elements: [
          {
            type:      'button',
            text:      { type: 'plain_text', text: '✅ Take Action' },
            style:     'primary',
            action_id: 'open_action_modal',
            value:     JSON.stringify({
              token:         rowData.token,
              rowIndex:      rowData.rowIndex,
              spreadsheetId: rowData.spreadsheetId
            })
          },
          {
            type:      'button',
            text:      { type: 'plain_text', text: '🔗 View in Sheet' },
            url:       sheetUrl,
            action_id: 'view_sheet_link'
          }
        ]
      }
    ]
  };

  var result = _fetchSlack('https://slack.com/api/chat.postMessage', {
    method:      'post',
    headers:     { Authorization: 'Bearer ' + token },
    contentType: 'application/json',
    payload:     JSON.stringify(payload)
  });

  return result.ok;
}

// ─────────────────────────────────────────────────────────────────────────────
// EDIT NOTIFICATION (triggered when another user manually edits a cell)
// Rich block message with "Acknowledge" button — clicking threads a reply
// ─────────────────────────────────────────────────────────────────────────────

function sendEditNotification(editData) {
  var token    = _getToken(editData.spreadsheetId);
  var channel  = _getChannel(editData.spreadsheetId);
  var sheetUrl = 'https://docs.google.com/spreadsheets/d/' + editData.spreadsheetId + '/edit';

  var changeText =
    '*' + editData.columnLetter + editData.row + '*: ' +
    '`' + editData.oldValue + '` → `' + editData.newValue + '`';

  var payload = {
    channel: channel,
    blocks: [
      {
        type: 'header',
        text: { type: 'plain_text', text: '✏️ Sheet Edit by ' + editData.editorEmail }
      },
      {
        type: 'section',
        fields: [
          { type: 'mrkdwn', text: '*Sheet:*\n'   + editData.sheetName },
          { type: 'mrkdwn', text: '*Editor:*\n'  + editData.editorEmail },
          { type: 'mrkdwn', text: '*Change:*\n'  + changeText },
          { type: 'mrkdwn', text: '*Time:*\n'    + editData.timestamp.toLocaleString() }
        ]
      },
      {
        type: 'section',
        text: {
          type: 'mrkdwn',
          text: '*<' + sheetUrl + '|📄 View Spreadsheet>*'
        }
      },
      {
        type: 'actions',
        elements: [
          {
            type:      'button',
            text:      { type: 'plain_text', text: '👍 Acknowledge' },
            action_id: 'acknowledge_edit',
            value:     JSON.stringify({
              spreadsheetId: editData.spreadsheetId,
              editorEmail:   editData.editorEmail,
              sheetName:     editData.sheetName,
              row:           editData.row,
              columnLetter:  editData.columnLetter
            })
          },
          {
            type:      'button',
            text:      { type: 'plain_text', text: '🔗 View in Sheet' },
            url:       sheetUrl,
            action_id: 'view_sheet_edit_link'
          }
        ]
      }
    ]
  };

  var result = _fetchSlack('https://slack.com/api/chat.postMessage', {
    method:      'post',
    headers:     { Authorization: 'Bearer ' + token },
    contentType: 'application/json',
    payload:     JSON.stringify(payload)
  });

  // Store the message ts so acknowledge can post a thread reply
  if (result.ok && result.ts) {
    PropertiesService.getScriptProperties()
      .setProperty('EDIT_MSG_TS_' + editData.spreadsheetId + '_' + result.ts, result.ts);
  }

  return result.ok;
  }

  /**
  * Posts a threaded reply under an edit notification when the user clicks Acknowledge.
  * Called from Webhooks.js block_action handler.
  *
  * @param {string} spreadsheetId
  * @param {string} channel - Slack channel ID (from the original message payload)
  * @param {string} messageTs - timestamp of the original message (thread_ts)
  * @param {string} acknowledgedBy - Slack user ID who clicked
  */
  function postAcknowledgeThread(spreadsheetId, channel, messageTs, acknowledgedBy) {
  var token = _getToken(spreadsheetId);

  var payload = {
    channel:   channel,
    thread_ts: messageTs,
    text:      '✅ Acknowledged by <@' + acknowledgedBy + '>'
  };

  var result = _fetchSlack('https://slack.com/api/chat.postMessage', {
    method:      'post',
    headers:     { Authorization: 'Bearer ' + token },
    contentType: 'application/json',
    payload:     JSON.stringify(payload)
  });

  return result.ok;
  }

  // ─────────────────────────────────────────────────────────────────────────────
  // WEEKLY DIGEST
  // ─────────────────────────────────────────────────────────────────────────────

  function sendWeeklyDigestSlack(spreadsheetId, digestData) {
  var token    = _getToken(spreadsheetId);
  var channel  = _getChannel(spreadsheetId);
  var sheetUrl = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/edit';

  var editorsText = digestData.newEditors.length
    ? digestData.newEditors.join(', ')
    : 'None';

  var payload = {
    channel: channel,
    blocks: [
      {
        type: 'header',
        text: { type: 'plain_text', text: '📊 Weekly SheetAlerts Digest' }
      },
      {
        type: 'section',
        text: {
          type: 'mrkdwn',
          text: 'Here\'s a summary of everything that happened in your spreadsheet this week.'
        }
      },
      { type: 'divider' },
      {
        type: 'section',
        fields: [
          { type: 'mrkdwn', text: '*🚨 Condition Alerts Fired:*\n' + digestData.totalAlerts },
          { type: 'mrkdwn', text: '*✅ Alerts Resolved:*\n'        + digestData.resolved    },
          { type: 'mrkdwn', text: '*⏳ Alerts Still Pending:*\n'   + digestData.pending     },
          { type: 'mrkdwn', text: '*✏️ Manual Edits by Others:*\n' + digestData.manualEdits }
        ]
      },
      {
        type: 'section',
        text: {
          type: 'mrkdwn',
          text: '*Editors who made changes this week:*\n' + editorsText
        }
      },
      { type: 'divider' },
      {
        type: 'section',
        text: {
          type: 'mrkdwn',
          text: '*<' + sheetUrl + '|📄 Open Spreadsheet>*'
        }
      }
    ]
  };

  var result = _fetchSlack('https://slack.com/api/chat.postMessage', {
    method:      'post',
    headers:     { Authorization: 'Bearer ' + token },
    contentType: 'application/json',
    payload:     JSON.stringify(payload)
  });

  return result.ok;
  }

  // ─────────────────────────────────────────────────────────────────────────────
  // MODAL
  // ─────────────────────────────────────────────────────────────────────────────

  function openActionModal(triggerId, actionValue) {
  var parsedValue = JSON.parse(actionValue);
  var token       = _getToken(parsedValue.spreadsheetId);
  var rowIndex    = parsedValue.rowIndex;

  var payload = {
    trigger_id: triggerId,
    view: {
      type:             'modal',
      callback_id:      'resolve_alert_modal',
      private_metadata: actionValue,
      title:  { type: 'plain_text', text: 'Resolve Alert'  },
      submit: { type: 'plain_text', text: 'Submit'         },
      close:  { type: 'plain_text', text: 'Cancel'         },
      blocks: [
        {
          type: 'section',
          text: { type: 'mrkdwn', text: 'You are updating row *' + rowIndex + '*.' }
        },
        {
          type:     'input',
          block_id: 'status_block',
          element: {
            type:      'static_select',
            action_id: 'status_select',
            options: [
              { text: { type: 'plain_text', text: 'Paid'     }, value: 'Paid'     },
              { text: { type: 'plain_text', text: 'Resolved' }, value: 'Resolved' },
              { text: { type: 'plain_text', text: 'Ignored'  }, value: 'Ignored'  }
            ]
          },
          label: { type: 'plain_text', text: 'New Status' }
        },
        {
          type:     'input',
          block_id: 'notes_block',
          optional: true,
          element: {
            type:      'plain_text_input',
            action_id: 'notes_input',
            multiline: true
          },
          label: { type: 'plain_text', text: 'Notes' }
        }
      ]
    }
  };

  _fetchSlack('https://slack.com/api/views.open', {
    method:      'post',
    headers:     { Authorization: 'Bearer ' + token },
    contentType: 'application/json',
    payload:     JSON.stringify(payload)
  });
  }

  // ─────────────────────────────────────────────────────────────────────────────
  // APP HOME DASHBOARD
  // ─────────────────────────────────────────────────────────────────────────────

  function publishAppHome(userId, spreadsheetId) {
    var token               = _getToken(spreadsheetId);
    var scriptProps         = PropertiesService.getScriptProperties();
    var pendingAlertsRaw    = scriptProps.getProperty('PENDING_ALERTS_' + spreadsheetId);
    var pendingAlerts       = pendingAlertsRaw ? JSON.parse(pendingAlertsRaw) : [];
    var sheetUrl            = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/edit';

    var spreadsheet  = SpreadsheetApp.openById(spreadsheetId);
    var logSheet     = spreadsheet.getSheetByName('AlertsLog');

    var stats = { sent: 0, resolved: 0, pending: pendingAlerts.length };
    var historyBlocks = [];
    var editBlocks    = [];

    if (logSheet) {
      var data = logSheet.getDataRange().getValues();
      stats.sent = data.filter(function(r) { return r[3] !== 'Manual Edit'; }).length - 1; // minus header
      stats.resolved = data.filter(function(r) { return r[7] === true; }).length;

      var historyCount = 0;
      for (var i = data.length - 1; i > 0 && historyCount < 20; i--) {
        var changeType = data[i][3] || '';
        if (changeType === 'Manual Edit') {
          editBlocks.push({
            type: 'section',
            text: {
              type: 'mrkdwn',
              text: '✏️ *' + data[i][2] + '* edited *' + (data[i][14] || '?') + data[i][1] + '* (' + data[i][11] + ')\n' +
                    '`' + (data[i][15] || '') + '` → `' + (data[i][16] || '') + '`'
            }
          });
          historyCount++;
        } else if (data[i][7] === true) {
          historyBlocks.push({
            type: 'section',
            text: {
              type: 'mrkdwn',
              text: '✅ *Row ' + data[i][1] + '* — ' + data[i][2] +
                    '\nStatus: ' + data[i][3] + ' | Resolved by: ' + (data[i][9] || 'N/A')
            }
          });
          historyCount++;
        }
      }
    }

    var blocks = [
      {
        type: 'header',
        text: { type: 'plain_text', text: '📊 SheetAlerts Dashboard' }
      },
      {
        type: 'section',
        fields: [
          { type: 'mrkdwn', text: '*Total Alerts Sent:*\n' + stats.sent },
          { type: 'mrkdwn', text: '*Actions Taken:*\n' + stats.resolved },
          { type: 'mrkdwn', text: '*Outstanding:*\n' + stats.pending }
        ]
      },
      {
        type: 'actions',
        elements: [
          { type: 'button', text: { type: 'plain_text', text: '📄 Open Spreadsheet' }, url: sheetUrl }
        ]
      },
      { type: 'divider' },
      { type: 'header', text: { type: 'plain_text', text: '🔴 Active Alerts' } }
    ];

    if (pendingAlerts.length === 0) {
      blocks.push({ type: 'section', text: { type: 'mrkdwn', text: 'No active alerts! 🎉' } });
    } else {
      pendingAlerts.forEach(function(alert) {
        blocks.push({
          type: 'section',
          text: { type: 'mrkdwn', text: '*Row ' + alert.rowIndex + ':* ' + alert.clientName + '\nCondition: `' + alert.status + '`' },
          accessory: {
            type: 'button',
            text: { type: 'plain_text', text: '✅ Resolve' },
            style: 'primary',
            action_id: 'open_action_modal',
            value: JSON.stringify({ token: alert.token, rowIndex: alert.rowIndex, spreadsheetId: alert.spreadsheetId })
          }
        });
      });
    }

    blocks.push({ type: 'divider' }, { type: 'header', text: { type: 'plain_text', text: '📜 Recent Activity' } });
    blocks = blocks.concat(editBlocks.slice(0, 5)).concat(historyBlocks.slice(0, 5));

    _fetchSlack('https://slack.com/api/views.publish', {
      method: 'post',
      headers: { Authorization: 'Bearer ' + token },
      contentType: 'application/json',
      payload: JSON.stringify({ user_id: userId, view: { type: 'home', blocks: blocks } })
    });
  }
  // ─────────────────────────────────────────────────────────────────────────────
  // CASCADE SLACK MESSAGE
  // ─────────────────────────────────────────────────────────────────────────────

  function sendCascadeSlackMessage(cascadeRowData, slackUserId) {
  var token    = _getToken(cascadeRowData.spreadsheetId);
  var sheetUrl = 'https://docs.google.com/spreadsheets/d/' + cascadeRowData.spreadsheetId + '/edit';

  // Open a DM with the specific Slack user
  var dmResult = _fetchSlack('https://slack.com/api/conversations.open', {
    method:      'post',
    headers:     { Authorization: 'Bearer ' + token },
    contentType: 'application/json',
    payload:     JSON.stringify({ users: slackUserId })
  });
  if (!dmResult.ok) return false;
  var dmChannel = dmResult.channel.id;

  var payload = {
    channel: dmChannel,
    blocks: [
      {
        type: 'header',
        text: { type: 'plain_text', text: '🔗 Cascade Notification' }
      },
      {
        type: 'section',
        text: {
          type: 'mrkdwn',
          text: cascadeRowData.cascadeMessage + '\n\n' +
                '*Row ' + cascadeRowData.rowIndex + ':* ' + cascadeRowData.clientName +
                ' (' + cascadeRowData.sheetName + ')'
        }
      },
      {
        type: 'section',
        text: { type: 'mrkdwn', text: '*<' + sheetUrl + '|📄 Open Spreadsheet>*' }
      }
    ]
  };

  var result = _fetchSlack('https://slack.com/api/chat.postMessage', {
    method:      'post',
    headers:     { Authorization: 'Bearer ' + token },
    contentType: 'application/json',
    payload:     JSON.stringify(payload)
  });

  return result.ok;
  }