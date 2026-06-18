function buildPrivacyCard() {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Privacy Policy'))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph().setText('SheetAlerts respects your privacy. We only access the data strictly necessary to provide the alert services you configure.'))
      .addWidget(CardService.newTextParagraph().setText('1. Data Access: We access your spreadsheets to monitor for changes.'))
      .addWidget(CardService.newTextParagraph().setText('2. Data Storage: Configuration is stored in your Google Account properties.'))
      .addWidget(CardService.newTextParagraph().setText('3. Third-party: We send data to Slack and Gmail only as directed by your configuration.'))
      .addWidget(CardService.newTextParagraph().setText('For our full policy, visit: https://example.com/privacy'))) // Placeholder
    .build();
}

function buildTermsCard() {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Terms of Service'))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph().setText('By using SheetAlerts, you agree to the following terms:'))
      .addWidget(CardService.newTextParagraph().setText('1. Use: You will use this tool for lawful purposes only.'))
      .addWidget(CardService.newTextParagraph().setText('2. Reliability: While we strive for 100% uptime, Google Apps Script quotas and third-party API availability may affect performance.'))
      .addWidget(CardService.newTextParagraph().setText('3. Liability: We are not responsible for any data loss or missed alerts.')))
    .build();
}
