/**
 * Returns system domain information for the login page.
 * @returns {Object} adminDomain or error
 */
function getSystemDomainInfo() {
  try {
    var adminEmail = PropertiesService.getScriptProperties().getProperty('deployUser');
    if (!adminEmail) {
      throw new Error('システム管理者が設定されていません。');
    }
    var adminDomain = adminEmail.split('@')[1];
    return { adminDomain: adminDomain };
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * Serves the LoginPage HTML with client ID injected.
 * @returns {HtmlOutput}
 */
function doGetLoginPage() {
  var htmlOutput = HtmlService.createTemplateFromFile('LoginPage').evaluate();
  var clientId = PropertiesService.getScriptProperties().getProperty('GOOGLE_CLIENT_ID');
  htmlOutput.GOOGLE_CLIENT_ID = clientId;
  htmlOutput.setTitle('みんなの回答ボード ログイン');
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return htmlOutput;
}
