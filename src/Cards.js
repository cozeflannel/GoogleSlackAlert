/**
 * Entry point for the Google Workspace Add-on.
 * Triggered when the user opens the add-on sidebar.
 */
function onHomepage(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    return CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('No Spreadsheet Found'))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newTextParagraph().setText('Please open this add-on from within a Google Sheet.')))
      .build();
  }
  return buildSettingsCard(ss.getId());
}

/**
 * Main UI builder using CardService.
 * Replaces the legacy Sidebar.html.
 */
function buildSettingsCard(spreadsheetId, selectedSheetName) {
  var config      = getConfig();
  var sheetName   = selectedSheetName || config.SHEET_NAME || SpreadsheetApp.openById(spreadsheetId).getSheets()[0].getName();
  var slackStatus = getSlackConnectionStatus();
  var dsConfig    = getDownstreamConfig() || { enabled: false };

  // Auto-map columns if this is a new sheet or if we're force-refreshing
  var mapping = config.STATUS_COL ? config : guessColumnMapping(sheetName);

  var card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('SheetAlerts Setup')
      .setSubtitle('Fast, automatic monitoring'));

  // ── Section 1: Connection Status ─────────────────────────────────────────
  var statusSection = CardService.newCardSection();
  if (!slackStatus.connected) {
    statusSection.addWidget(CardService.newTextParagraph().setText('<b>Step 1: Connect to Slack</b>'));
    statusSection.addWidget(CardService.newTextButton()
      .setText('Connect Slack Workspace')
      .setOpenLink(CardService.newOpenLink()
        .setUrl(getSlackOAuthUrl())
        .setOpenAs(CardService.OpenAs.FULL_SIZE)
        .setOnClose(CardService.OnClose.RELOAD_ADD_ON)));
  } else {
    statusSection.addWidget(CardService.newTextParagraph().setText('✅ <b>Connected to Slack</b>'));
  }
  card.addSection(statusSection);

  // ── Section 2: Sheet & Trigger ───────────────────────────────────────────
  var configSection = CardService.newCardSection()
    .setHeader('Step 2: What to monitor?');
  
  var sheetPicker = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle('Select Sheet')
    .setFieldName('SHEET_NAME')
    .setOnChangeAction(CardService.newAction().setFunctionName('handleSheetChange'));

  var sheets = SpreadsheetApp.openById(spreadsheetId).getSheets();
  sheets.forEach(function(s) {
    var name = s.getName();
    sheetPicker.addItem(name, name, name === sheetName);
  });
  configSection.addWidget(sheetPicker);

  var headers = getHeadersForSheet(sheetName);
  var statusPicker = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle('Condition Column')
    .setFieldName('STATUS_COL');
  
  headers.forEach(function(h, i) {
    if (h) statusPicker.addItem(h, String(i), String(i) === String(mapping.STATUS_COL));
  });
  configSection.addWidget(statusPicker);

  configSection.addWidget(CardService.newTextInput()
    .setFieldName('TRIGGER_VALUE')
    .setTitle('Alert when value equals...')
    .setHint('e.g. Late, Overdue')
    .setValue(config.TRIGGER_VALUE || ''));

  if (slackStatus.connected) {
    var channelPicker = CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.DROPDOWN)
      .setTitle('Post to Slack Channel')
      .setFieldName('SLACK_CHANNEL');

    var channels = getSlackChannels(spreadsheetId);
    channels.forEach(function(c) {
      channelPicker.addItem(c.name, c.name, c.name === config.SLACK_CHANNEL);
    });
    configSection.addWidget(channelPicker);
  }
  
  card.addSection(configSection);

  // ── Section 3: Advanced (Collapsible-ish) ────────────────────────────────
  var advancedSection = CardService.newCardSection()
    .setHeader('Advanced Mapping (Auto-detected)')
    .setCollapsible(true);

  var advFields = [
    { name: 'NAME_COL',       label: 'Name Column' },
    { name: 'EMAIL_COL',      label: 'Email Column' },
    { name: 'EXTRA_INFO_COL', label: 'Extra Info' }
  ];

  advFields.forEach(function(f) {
    var sel = CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.DROPDOWN)
      .setTitle(f.label)
      .setFieldName(f.name);
    sel.addItem('-- None --', '-1', mapping[f.name] === -1);
    headers.forEach(function(h, i) {
      if (h) sel.addItem(h, String(i), String(i) === String(mapping[f.name]));
    });
    advancedSection.addWidget(sel);
  });
  card.addSection(advancedSection);

  // ── Section 4: Auto-Timestamp & Status ───────────────────────────────────
  var logicSection = CardService.newCardSection()
    .setHeader('Auto-Date & Status Logic')
    .setCollapsible(true);

  logicSection.addWidget(CardService.newTextParagraph()
    .setText('Automatically write a completion date and track Late/On-Time status.'));

  var triggerColPicker = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle('When this column...')
    .setFieldName('AUTO_TIMESTAMP_TRIGGER_COL');
  
  var dateColPicker = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle('...changes, write date to...')
    .setFieldName('AUTO_TIMESTAMP_COL');

  triggerColPicker.addItem('-- Disabled --', '-1', config.AUTO_TIMESTAMP_TRIGGER_COL === '-1');
  dateColPicker.addItem('-- Disabled --', '-1', config.AUTO_TIMESTAMP_COL === '-1');

  headers.forEach(function(h, i) {
    if (h) {
      triggerColPicker.addItem(h, String(i), String(i) === String(config.AUTO_TIMESTAMP_TRIGGER_COL));
      dateColPicker.addItem(h, String(i), String(i) === String(config.AUTO_TIMESTAMP_COL));
    }
  });
  logicSection.addWidget(triggerColPicker);

  logicSection.addWidget(CardService.newTextInput()
    .setFieldName('AUTO_TIMESTAMP_VALUE')
    .setTitle('...equals this value (e.g. Done)')
    .setValue(config.AUTO_TIMESTAMP_VALUE || ''));

  logicSection.addWidget(dateColPicker);

  var dueColPicker = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle('Due Date column')
    .setFieldName('DUE_DATE_COL');
  
  var finalStatusPicker = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle('Write Late/On-Time status to...')
    .setFieldName('FINAL_STATUS_COL');

  dueColPicker.addItem('-- None --', '-1', config.DUE_DATE_COL === '-1');
  finalStatusPicker.addItem('-- None --', '-1', config.FINAL_STATUS_COL === '-1');

  headers.forEach(function(h, i) {
    if (h) {
      dueColPicker.addItem(h, String(i), String(i) === String(config.DUE_DATE_COL));
      finalStatusPicker.addItem(h, String(i), String(i) === String(config.FINAL_STATUS_COL));
    }
  });
  logicSection.addWidget(dueColPicker);
  logicSection.addWidget(finalStatusPicker);

  card.addSection(logicSection);

  // ── Section 5: Slack Action Modal Mapping ────────────────────────────────
  var actionSection = CardService.newCardSection()
    .setHeader('Slack "Take Action" Modal')
    .setCollapsible(true);

  actionSection.addWidget(CardService.newTextParagraph()
    .setText('Choose which columns should appear as editable fields in the Slack modal.'));

  var currentActions = JSON.parse(config.ACTIONABLE_COLS || '[]');
  var actionPicker   = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setTitle('Actionable Columns')
    .setFieldName('ACTIONABLE_COLS');

  headers.forEach(function(h, i) {
    if (h) {
      actionPicker.addItem(h, String(i), currentActions.indexOf(String(i)) !== -1);
    }
  });
  actionSection.addWidget(actionPicker);
  card.addSection(actionSection);

  // ── Footer ───────────────────────────────────────────────────────────────
  card.setFixedFooter(CardService.newFixedFooter()
    .setPrimaryButton(CardService.newTextButton()
      .setText('Save & Start Monitoring')
      .setOnClickAction(CardService.newAction().setFunctionName('handleSaveSettings'))));

  return card.build();
}

// ── Action Handlers ────────────────────────────────────────────────────────

function handleSheetChange(e) {
  var sheetName = e.formInput.SHEET_NAME;
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(buildSettingsCard(SpreadsheetApp.getActiveSpreadsheet().getId(), sheetName)))
    .build();
}

function handleToggleDownstream(e) {
  var enabled = e.formInput.ds_enabled === 'true';
  // We don't save yet, just refresh UI to show/hide fields
  // In a real app, you might want to temporarily store this in privateMetadata
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(buildSettingsCard(SpreadsheetApp.getActiveSpreadsheet().getId(), e.formInput.SHEET_NAME)))
    .build();
}

function handleSaveSettings(e) {
  try {
    var config = {
      SHEET_NAME:                 e.formInput.SHEET_NAME,
      SLACK_CHANNEL:              e.formInput.SLACK_CHANNEL || '',
      NAME_COL:                   e.formInput.NAME_COL,
      EMAIL_COL:                  e.formInput.EMAIL_COL,
      EXTRA_INFO_COL:             e.formInput.EXTRA_INFO_COL,
      STATUS_COL:                 e.formInput.STATUS_COL,
      DATE_COL:                   e.formInput.DATE_COL || '-1',
      TRIGGER_VALUE:              e.formInput.TRIGGER_VALUE,
      AUTO_TIMESTAMP_COL:         e.formInput.AUTO_TIMESTAMP_COL,
      AUTO_TIMESTAMP_TRIGGER_COL: e.formInput.AUTO_TIMESTAMP_TRIGGER_COL,
      AUTO_TIMESTAMP_VALUE:       e.formInput.AUTO_TIMESTAMP_VALUE,
      DUE_DATE_COL:               e.formInput.DUE_DATE_COL,
      FINAL_STATUS_COL:           e.formInput.FINAL_STATUS_COL,
      ACTIONABLE_COLS:            JSON.stringify(e.formInputs.ACTIONABLE_COLS || [])
    };

    saveConfig(config);

    // Warp Speed: Trigger an immediate scan so the user sees results instantly
    try {
      runConditionCheck();
    } catch (e) {
      Logger.log('Immediate scan error: ' + e.toString());
    }

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('✅ Configuration saved! Monitoring started.'))
      .build();
  } catch (err) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('❌ Error saving: ' + err.toString()))
      .build();
  }
}

function handleDisconnectSlack(e) {
  disconnectSlack();
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(buildSettingsCard(SpreadsheetApp.getActiveSpreadsheet().getId())))
    .setNotification(CardService.newNotification().setText('Slack disconnected.'))
    .build();
}

function handleDisconnectApp(e) {
  disconnectApp();
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText('App disconnected and triggers removed.'))
    .setNavigation(CardService.newNavigation().popToRoot())
    .build();
}

function handleShowPrivacy(e) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(buildPrivacyCard()))
    .build();
}

function handleShowTerms(e) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(buildTermsCard()))
    .build();
}
