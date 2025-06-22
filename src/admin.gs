// In the Apps Script environment, handleError is defined globally.
// When running tests in Node, import it from the module.
if (typeof module !== 'undefined') {
  var handleError = require('./ErrorHandling.gs').handleError;
}

function safeGetUserEmail() {
  try {
    return Session.getActiveUser().getEmail();
  } catch (e) {
    return '';
  }
}

function getAdminEmails() {
  const props = PropertiesService.getScriptProperties();
  const str = props.getProperty('ADMIN_EMAILS') || '';
  return str.split(',').map(e => e.trim()).filter(Boolean);
}

function isUserAdmin(email) {
  const userEmail = email || safeGetUserEmail();
  return getAdminEmails().includes(userEmail);
}

function checkAdmin() {
  return isUserAdmin();
}

function getSheets() {
  try {
    const allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    const visibleSheets = allSheets.filter(sheet => !sheet.isSheetHidden());
    const filtered = visibleSheets.filter(sheet => {
      const name = sheet.getName();
      return name !== 'Config' && name !== ROSTER_CONFIG.SHEET_NAME;
    });
    return filtered.map(sheet => sheet.getName());
  } catch (error) {
    handleError('getSheets', error);
  }
}

function getAppSettings() {
  const properties = PropertiesService.getScriptProperties() || {};
  const getProp = typeof properties.getProperty === 'function' ? (k) => properties.getProperty(k) : () => null;
  const published = getProp(APP_PROPERTIES.IS_PUBLISHED);
  const sheet = getProp(APP_PROPERTIES.ACTIVE_SHEET);
  const showDetailsProp = getProp(APP_PROPERTIES.SHOW_DETAILS);
  let activeName = sheet;
  if (!activeName) {
    try {
      const sheets = getSheets();
      activeName = sheets[0] || '';
    } catch (error) {
      console.error('getAppSettings Error:', error);
      activeName = '';
    }
  }
  return {
    isPublished: published === null ? true : published === 'true',
    activeSheetName: activeName,
    showDetails: showDetailsProp === 'true'
  };
}

function getAdminSettings() {
  const props = PropertiesService.getScriptProperties();
  const allSheets = getSheets();
  const currentUserEmail = safeGetUserEmail();
  const adminEmails = getAdminEmails();
  const appSettings = getAppSettings();
  return {
    allSheets: allSheets,
    currentUserEmail: currentUserEmail,
    deployId: props.getProperty('DEPLOY_ID'),
    webAppUrl: getWebAppUrl(),
    adminEmails: adminEmails,
    isUserAdmin: adminEmails.includes(currentUserEmail),
    isPublished: appSettings.isPublished,
    activeSheetName: appSettings.activeSheetName,
    showDetails: appSettings.showDetails
  };
}

function publishApp(sheetName) {
  if (!checkAdmin()) {
    throw new Error('権限がありません。');
  }
  if (!sheetName) {
    throw new Error('シート名が指定されていません。');
  }
  if (typeof require === 'function') {
    require('./Code.gs').prepareSheetForBoard(sheetName);
  } else {
    prepareSheetForBoard(sheetName);
  }
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty(APP_PROPERTIES.IS_PUBLISHED, 'true');
  properties.setProperty(APP_PROPERTIES.ACTIVE_SHEET, sheetName);
  return `「${sheetName}」を公開しました。`;
}

function unpublishApp() {
  if (!checkAdmin()) {
    throw new Error('権限がありません。');
  }
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty(APP_PROPERTIES.IS_PUBLISHED, 'false');
  return 'アプリを非公開にしました。';
}

function setShowDetails(flag) {
  if (!checkAdmin()) {
    throw new Error('権限がありません。');
  }
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty(APP_PROPERTIES.SHOW_DETAILS, String(flag));
  return `詳細表示を${flag ? '有効' : '無効'}にしました。`;
}

function saveWebAppUrl(url) {
  const props = PropertiesService.getScriptProperties();
  props.setProperties({ WEB_APP_URL: (url || '').trim() });
}

function getWebAppUrl() {
  const props = PropertiesService.getScriptProperties();
  return (props.getProperty('WEB_APP_URL') || '').trim();
}

function saveDeployId(id) {
  const props = PropertiesService.getScriptProperties();
  props.setProperties({ DEPLOY_ID: (id || '').trim() });
}

// functions not exported previously are omitted from exports
if (typeof module !== 'undefined') {
  module.exports = {
    safeGetUserEmail,
    getAdminEmails,
    isUserAdmin,
    checkAdmin,
    getSheets,
    getAppSettings,
    getAdminSettings,
    publishApp,
    unpublishApp,
    setShowDetails,
    saveWebAppUrl,
    getWebAppUrl,
    saveDeployId
  };
}
