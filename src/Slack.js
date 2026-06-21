// ─────────────────────────────────────────────────────────────────────────────
// TOKEN HELPER
// ─────────────────────────────────────────────────────────────────────────────

function _getToken(spreadsheetId) {
  if (!spreadsheetId) {
    throw new Error('spreadsheetId is required to look up the Slack token.');
  }
  
  // Tokens are now stored in Supabase. This helper is used for initial alerts 
  // fired from GAS. It will now query the Supabase installations table.
  var scriptProps = PropertiesService.getScriptProperties();
  var supabaseUrl = scriptProps.getProperty('SUPABASE_URL') + '/rest/v1/installations?spreadsheet_id=eq.' + spreadsheetId + '&select=slack_bot_token';
  var serviceKey  = scriptProps.getProperty('SUP_SECRET_KEY');

  try {
    var response = UrlFetchApp.fetch(supabaseUrl, {
      headers: { 'apikey': serviceKey, 'Authorization': 'Bearer ' + serviceKey },
      muteHttpExceptions: true
    });
    var data = JSON.parse(response.getContentText());
    if (data && data.length > 0) {
      return data[0].slack_bot_token;
    }
  } catch (e) {
    Logger.log('Error fetching token from Supabase: ' + e.toString());
  }

  throw new Error('No Slack token found in Supabase for spreadsheet ' + spreadsheetId + '. Complete OAuth first.');
}

function _getChannel(spreadsheetId) {
  var docProps = PropertiesService.getDocumentProperties();
  var channel = docProps.getProperty('SLACK_CHANNEL');
  return channel || '#general';
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
// CASCADE SLACK MESSAGE
// ─────────────────────────────────────────────────────────────────────────────

function sendCascadeSlackMessage(cascadeRowData, slackUserId) {
  var token    = _getToken(cascadeRowData.spreadsheetId);
  var sheetUrl = 'https://docs.google.com/spreadsheets/d/' + cascadeRowData.spreadsheetId + '/edit';

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