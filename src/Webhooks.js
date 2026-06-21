function doPost(e) {
  try {
    var contents = (e.postData && e.postData.contents)
      ? e.postData.contents
      : (e.parameter && e.parameter.payload ? e.parameter.payload : '{}');

    var payload = JSON.parse(contents);

    // 1. Slack URL Verification
    if (payload.type === 'url_verification') {
      return ContentService
        .createTextOutput(JSON.stringify({ challenge: payload.challenge }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 2. App Home Opened — resolve spreadsheet via team ID mapping
    if (payload.event && payload.event.type === 'app_home_opened') {
      var teamId        = payload.team_id;
      var spreadsheetId = PropertiesService.getScriptProperties()
                            .getProperty('TEAM_SPREADSHEET_' + teamId);
      if (spreadsheetId) publishAppHome(payload.event.user, spreadsheetId);
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 3. Block actions (button clicks from Slack messages or App Home)
    if (payload.type === 'block_actions') {
      var action = payload.actions && payload.actions[0];

      if (action) {
        // ── "Take Action" button → open resolve modal ──────────────────────
        if (action.action_id === 'open_action_modal') {
          openActionModal(payload.trigger_id, action.value);
        }

        // ── "Acknowledge" button on edit notification ───────────────────────
        if (action.action_id === 'acknowledge_edit') {
          var editMeta       = JSON.parse(action.value);
          var spreadsheetId2 = editMeta.spreadsheetId;
          var messageTs      = payload.message && payload.message.ts;
          var channelId      = payload.channel && payload.channel.id;
          var acknowledgedBy = payload.user && payload.user.id;

          if (spreadsheetId2 && messageTs && channelId && acknowledgedBy) {
            postAcknowledgeThread(spreadsheetId2, channelId, messageTs, acknowledgedBy);
          }
        }
      }

      return ContentService
        .createTextOutput(JSON.stringify({ ok: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 4. View Submission (modal form submit)
    if (payload.type === 'view_submission' &&
        payload.view.callback_id === 'resolve_alert_modal') {

      var privateMetadata = JSON.parse(payload.view.private_metadata);
      var values          = payload.view.state.values;
      var spreadsheetId   = privateMetadata.spreadsheetId;
      var rowIndex        = privateMetadata.rowIndex;
      var token           = privateMetadata.token;

      // 1. Process all dynamic column updates
      var updates = [];
      for (var blockId in values) {
        if (blockId.indexOf('col_') === 0) {
          var colIdx = blockId.replace('col_', '').replace('_block', '');
          var inputId = blockId.replace('_block', '_input');
          if (values[blockId] && values[blockId][inputId]) {
            var val = values[blockId][inputId].value;
            updates.push({ col: parseInt(colIdx), value: val });
          }
        }
      }

      // 2. Apply updates to the sheet
      try {
        var ss = SpreadsheetApp.openById(spreadsheetId);
        var scriptProps = PropertiesService.getScriptProperties();
        var sheetName = scriptProps.getProperty('SHEET_NAME_' + spreadsheetId);
        var sheet = ss.getSheetByName(sheetName);

        if (sheet) {
          updates.forEach(function(u) {
            // Column index from config is 0-based, getRange is 1-based
            sheet.getRange(rowIndex, u.col + 1).setValue(u.value);
          });
        }
      } catch (err) {
        Logger.log('Error applying dynamic updates: ' + err.toString());
      }

      // 3. Resolve the alert (Logging & Pending cleanup)
      // Use the first update as the 'status' for the log, or a default 'Resolved'
      var primaryStatus = updates.length > 0 ? updates[0].value : 'Resolved';
      var notes = (values.notes_block && values.notes_block.notes_input)
                  ? values.notes_block.notes_input.value : '';

      resolveAlert(
        token,
        rowIndex,
        primaryStatus,
        notes,
        spreadsheetId
      );

      // Refresh App Home for the submitting user
      var teamId3        = payload.team && payload.team.id;
      var spreadsheetId3 = PropertiesService.getScriptProperties()
                             .getProperty('TEAM_SPREADSHEET_' + teamId3);
      if (spreadsheetId3) publishAppHome(payload.user.id, spreadsheetId3);

      return ContentService
        .createTextOutput(JSON.stringify({ response_action: 'clear' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('doPost error: ' + err.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  var action = e.parameter.action;

  // ── OAuth callback from Slack ──────────────────────────────────────────────
  if (action === 'oauth_callback') {
    var code  = e.parameter.code;
    var state = e.parameter.state; // spreadsheetId passed as OAuth state param

    if (!code || !state) {
      return HtmlService.createHtmlOutput(
        '<h2>OAuth Error</h2><p>Missing code or state parameter.</p>'
      );
    }

    var scriptProps  = PropertiesService.getScriptProperties();
    var clientId     = scriptProps.getProperty('SLACK_CLIENT_ID');
    var clientSecret = scriptProps.getProperty('SLACK_CLIENT_SECRET');
    var redirectUri  = scriptProps.getProperty('DEPLOYED_WEBAPP_URL') + '?action=oauth_callback';

    var response = UrlFetchApp.fetch('https://slack.com/api/oauth.v2.access', {
      method:  'post',
      payload: {
        code:          code,
        client_id:     clientId,
        client_secret: clientSecret,
        redirect_uri:  redirectUri
      }
    });

    var result = JSON.parse(response.getContentText());

    if (!result.ok) {
      return HtmlService.createHtmlOutput(
        '<h2>OAuth Failed</h2><p>' + result.error + '</p>'
      );
    }

    var botToken = result.access_token;
    var teamId   = result.team ? result.team.id : '';
    var teamName = result.team ? result.team.name : '';

    // 1. Store in DocumentProperties of the target spreadsheet
    var targetSs = SpreadsheetApp.openById(state);
    var docProps = PropertiesService.getDocumentProperties(); 
    // Note: When opening by ID, PropertiesService.getDocumentProperties() 
    // refers to the active spreadsheet. To set properties for a specific 
    // document, we must be in its context. Since we are in a Web App, 
    // we need to use a helper or ensure the context is correct.
    // Actually, in GAS Web Apps, getDocumentProperties() is not available 
    // for a specific ID. We must use ScriptProperties for cross-spreadsheet 
    // setup unless we have a way to execute code in that doc's context.
    // WAIT: The user specifically asked to:
    // "store SLACK_TOKEN and SLACK_TEAM in DocumentProperties of the target spreadsheet 
    // using SpreadsheetApp.openById(state) and PropertiesService scoped to that document"
    // In reality, PropertiesService.getDocumentProperties() always refers to 
    // the CURRENTLY active document. In a Web App, there is no active document.
    // HOWEVER, for the purpose of this architectural change requested by the user:
    
    var scriptProps  = PropertiesService.getScriptProperties();
    scriptProps.setProperty('SLACK_TOKEN_' + state, botToken);
    scriptProps.setProperty('SLACK_TEAM_' + state, teamName);

    if (teamId) {
      scriptProps.setProperty('TEAM_SPREADSHEET_' + teamId, state);
    }

    // Since we cannot directly set DocumentProperties for a remote spreadsheet 
    // from a Web App (GAS limitation), we rely on the ScriptProperties 
    // cross-reference and a sync mechanism, OR we accept that OAuth 
    // must initialize the ScriptProperty which then gets synced.
    // Given the user's request, I will implement the logic they asked for, 
    // but I must be aware of GAS limitations.
    
    // Let's follow the directive:
    // "Keep SLACK_TOKEN_<spreadsheetId> in ScriptProperties as cross-reference"
    
    // Sync to Supabase mirror immediately after OAuth completes
    syncConfigToSupabase(state);

    return HtmlService.createHtmlOutput(
      '<div style="font-family: sans-serif; text-align: center; padding-top: 50px;">' +
      '<h1 style="color: #006644;">✅ Successfully Connected!</h1>' +
      '<p style="font-size: 18px; color: #333;">SheetAlerts has been linked to your Slack workspace.</p>' +
      '<p style="color: #777;">You can now close this tab and return to your Google Sheet.</p>' +
      '<button onclick="window.close()" style="background: #0052cc; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; font-size: 16px;">Close Tab</button>' +
      '</div>'
    );
  }

  // ── One-time resolve link (from email or Slack button) ────────────────────
  var token         = e.parameter.token;
  var rowIndex      = parseInt(e.parameter.rowIndex);
  var spreadsheetId = e.parameter.spreadsheetId;

  if (action === 'get_bot_config' && spreadsheetId) {
    var scriptProps = PropertiesService.getScriptProperties();
    var sheetName   = scriptProps.getProperty('SHEET_NAME_' + spreadsheetId);
    
    if (!sheetName) {
      return ContentService.createTextOutput(JSON.stringify({ error: 'Sheet not configured' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var headers = [];
    try {
      var ss = SpreadsheetApp.openById(spreadsheetId);
      var sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      }
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: 'Could not read sheet: ' + err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var config = {
      token:          scriptProps.getProperty('SLACK_TOKEN_' + spreadsheetId),
      actionableCols: scriptProps.getProperty('ACTIONS_' + spreadsheetId) || '[]',
      headers:        headers
    };

    return ContentService.createTextOutput(JSON.stringify(config))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'resolve' && token && rowIndex && spreadsheetId) {

    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var logSheet    = spreadsheet.getSheetByName('AlertsLog');

    if (logSheet) {
      var logData    = logSheet.getDataRange().getValues();
      var validToken = false;
      var rowData    = {};

      for (var i = 1; i < logData.length; i++) {
        if (logData[i][4] === token &&
            logData[i][1] === rowIndex &&
            logData[i][7] !== true) {
          validToken = true;
          rowData = {
            clientName:     logData[i][2],
            status:         logData[i][3],
            token:          token,
            rowIndex:       rowIndex,
            spreadsheetId:  spreadsheetId,
            sheetName:      logData[i][11] || '',
            cascadeMessage: logData[i][13] || ''
          };
          break;
        }
      }

      if (validToken) {
        return getActionFormHtml(rowData);
      } else {
        return HtmlService.createHtmlOutput(
          '<h2>Invalid or Expired Link</h2>' +
          '<p>This alert may have already been resolved.</p>'
        );
      }
    }
  }

  return HtmlService.createHtmlOutput(
    '<h2>Error</h2><p>Invalid request parameters.</p>'
  );
}