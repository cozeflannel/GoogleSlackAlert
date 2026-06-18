function diagCheckProperties() {
  var sp    = PropertiesService.getScriptProperties();
  var props = sp.getProperties();
  var keys  = Object.keys(props).sort();

  Logger.log('ALL SCRIPT PROPERTIES (' + keys.length + ' total)');
  keys.forEach(function(k) {
    var v = (k.indexOf('TOKEN') !== -1 || k.indexOf('SECRET') !== -1)
      ? '***hidden***' : props[k];
    Logger.log(k + ' = ' + JSON.stringify(v));
  });

  var ssId = props['SPREADSHEET_ID'] || '';
  Logger.log('\n--- KEY CHECKS (Modern Architecture) ---');
  [
    ['SPREADSHEET_ID (Global)',    !!ssId],
    ['SLACK_CLIENT_ID',            !!props['SLACK_CLIENT_ID']],
    ['DEPLOYED_WEBAPP_URL',        !!props['DEPLOYED_WEBAPP_URL']],
    ['SLACK_TOKEN (Multi-tenant)', !!props['SLACK_TOKEN_' + ssId]],
    ['SHEET_NAME (Multi-tenant)',  !!props['SHEET_NAME_' + ssId]]
  ].forEach(function(c) { Logger.log((c[1] ? '✅ ' : '❌ ') + c[0]); });
}

function diagTestSlackToken() {
  var sp    = PropertiesService.getScriptProperties();
  var ssId  = sp.getProperty('SPREADSHEET_ID');
  if (!ssId) { Logger.log('NO SPREADSHEET_ID found.'); return; }
  var token = sp.getProperty('SLACK_TOKEN_' + ssId);
  if (!token) { Logger.log('NO TOKEN for ' + ssId); return; }
  
  var result = _fetchSlack('https://slack.com/api/auth.test', {
    method: 'post', headers: { Authorization: 'Bearer ' + token }
  });
  Logger.log('auth.test: ' + (result.ok ? 'VALID (bot=' + result.user + ')' : 'INVALID: ' + result.error));
}

/**
 * DANGER: Clears ALL script properties.
 * Use this to get out of the "sauce" and start fresh.
 */
function diagVerifyBridge() {
  var props = PropertiesService.getScriptProperties().getProperties();
  var ssId  = props['SPREADSHEET_ID'];
  var url   = props['DEPLOYED_WEBAPP_URL'] || '';

  Logger.log('--- BRIDGE VERIFICATION ---');
  Logger.log('1. Spreadsheet ID: ' + (ssId || '❌ MISSING (Run Save Config in Sidebar)'));
  Logger.log('2. Deployed URL:   ' + (url || '❌ MISSING'));
  
  if (url.indexOf('script.google.com') !== -1) {
    Logger.log('⚠️ WARNING: DEPLOYED_WEBAPP_URL points directly to Google, not Supabase.');
  }

  if (url.indexOf('apjftvnmskckrhgrdpbk') !== -1) {
    Logger.log('✅ Bridge points to Supabase.');
  }

  Logger.log('\nNext Step: Fix the space in your Supabase Secret "GAS_WEBAPP_URL" via the web dashboard.');
}
