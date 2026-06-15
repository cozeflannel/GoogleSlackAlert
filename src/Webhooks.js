function doPost(e) {
  try {
    var contents = (e.postData && e.postData.contents)
      ? e.postData.contents
      : (e.parameter && e.parameter.payload ? e.parameter.payload : "{}");

    var payload = JSON.parse(contents);

    // 1. Slack URL Verification
    if (payload.type === "url_verification") {
      return ContentService
        .createTextOutput(JSON.stringify({ challenge: payload.challenge }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 2. App Home Opened — resolve spreadsheet via team ID mapping
    if (payload.event && payload.event.type === "app_home_opened") {
      var teamId        = payload.team_id;
      var spreadsheetId = PropertiesService.getScriptProperties()
                            .getProperty('TEAM_SPREADSHEET_' + teamId);
      if (spreadsheetId) publishAppHome(payload.event.user, spreadsheetId);
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 3. Block actions
    if (payload.type === "block_actions") {
      var action = payload.actions && payload.actions[0];
      if (action && action.action_id === "open_action_modal") {
        openActionModal(payload.trigger_id, action.value);
      }
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 4. View Submission
    if (payload.type === "view_submission" &&
        payload.view.callback_id === "resolve_alert_modal") {
      var privateMetadata = JSON.parse(payload.view.private_metadata);
      var values          = payload.view.state.values;
      var newStatus       = values.status_block.status_select.selected_option.value;
      var notes           = values.notes_block && values.notes_block.notes_input
                              ? values.notes_block.notes_input.value : "";

      resolveAlert(
        privateMetadata.token,
        privateMetadata.rowIndex,
        newStatus,
        notes,
        privateMetadata.spreadsheetId
      );

      var teamId2        = payload.team && payload.team.id;
      var spreadsheetId2 = PropertiesService.getScriptProperties()
                             .getProperty('TEAM_SPREADSHEET_' + teamId2);
      if (spreadsheetId2) publishAppHome(payload.user.id, spreadsheetId2);

      return ContentService
        .createTextOutput(JSON.stringify({ response_action: "clear" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  var action = e.parameter.action;

  // ── OAuth callback from Slack ──────────────────────────────────────────────
  if (action === "oauth_callback") {
    var code  = e.parameter.code;
    var state = e.parameter.state; // spreadsheetId passed as OAuth state param

    if (!code || !state) {
      return HtmlService.createHtmlOutput(
        "<h2>OAuth Error</h2><p>Missing code or state parameter.</p>"
      );
    }

    var scriptProps  = PropertiesService.getScriptProperties();
    var clientId     = scriptProps.getProperty("SLACK_CLIENT_ID");
    var clientSecret = scriptProps.getProperty("SLACK_CLIENT_SECRET");
    var redirectUri  = ScriptApp.getService().getUrl() + "?action=oauth_callback";

    var response = UrlFetchApp.fetch("https://slack.com/api/oauth.v2.access", {
      method:  "post",
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
        "<h2>OAuth Failed</h2><p>" + result.error + "</p>"
      );
    }

    var botToken = result.access_token;
    var teamId   = result.team ? result.team.id : "";

    // Store token keyed by spreadsheet ID
    scriptProps.setProperty("SLACK_TOKEN_" + state, botToken);
    // Store team name for display
    scriptProps.setProperty("SLACK_TEAM_"  + state,
      result.team ? result.team.name : "");
    // Map team ID → spreadsheet ID so app_home_opened can find the right sheet
    if (teamId) {
      scriptProps.setProperty("TEAM_SPREADSHEET_" + teamId, state);
    }

    // Do NOT call SpreadsheetApp here — just return success HTML
    return HtmlService.createHtmlOutput(
      "<h2>✅ Connected!</h2>" +
      "<p>Slack has been successfully connected to your spreadsheet.<br>" +
      "You can close this tab and return to your sheet.</p>"
    );
  }

  // ── One-time resolve link (from email or Slack button) ────────────────────
  var token         = e.parameter.token;
  var rowIndex      = parseInt(e.parameter.rowIndex);
  var spreadsheetId = e.parameter.spreadsheetId;

  if (action === "resolve" && token && rowIndex && spreadsheetId) {
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var logSheet    = spreadsheet.getSheetByName("AlertsLog");

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
          "<h2>Invalid or Expired Link</h2>" +
          "<p>This alert may have already been resolved.</p>"
        );
      }
    }
  }

  return HtmlService.createHtmlOutput(
    "<h2>Error</h2><p>Invalid request parameters.</p>"
  );
}