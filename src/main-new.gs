/**
 * @fileoverview メインエントリーポイント - シンプルで責任範囲を限定
 */

/**
 * HTMLファイルをインクルード
 * @param {string} filename - ファイル名
 * @returns {string} HTML内容
 */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (error) {
    logError(error, 'include', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM, { filename });
    return `<!-- Error loading ${filename} -->`;
  }
}

/**
 * Webアプリケーションのメインエントリーポイント
 * @param {Object} e - リクエストパラメータ
 * @returns {GoogleAppsScript.HTML.HtmlOutput}
 */
function doGet(e) {
  try {
    // システム初期化チェック
    if (!isSystemInitialized()) {
      return showSetupPage();
    }
    
    // ユーザー認証チェック
    const user = authenticateUser();
    if (!user) {
      return showLoginPage();
    }
    
    // リクエストパラメータ解析
    const params = parseRequestParams(e);
    
    // ルーティング
    switch (params.mode) {
      case 'setup':
        return showAppSetupPage(params.userId);
      case 'admin':
        return showAdminPanel(user, params);
      case 'unpublished':
        return showUnpublishedPage(params);
      default:
        return showMainPage(user, params);
    }
    
  } catch (error) {
    logError(error, 'doGet', ERROR_SEVERITY.CRITICAL, ERROR_CATEGORIES.SYSTEM);
    return showErrorPage(error);
  }
}

/**
 * POSTリクエストのエントリーポイント
 * @param {Object} e - リクエストパラメータ
 * @returns {GoogleAppsScript.Content.TextOutput}
 */
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    
    // APIルーティング
    const result = handleApiRequest(action, params);
    
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    logError(error, 'doPost', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM);
    return ContentService
      .createTextOutput(JSON.stringify({ 
        success: false, 
        error: error.message 
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * システムが初期化されているかチェック
 * @returns {boolean}
 */
function isSystemInitialized() {
  const dbId = Properties.getScriptProperty('DATABASE_SPREADSHEET_ID');
  const adminEmail = Properties.getScriptProperty('ADMIN_EMAIL');
  return !!(dbId && adminEmail);
}

/**
 * ユーザー認証
 * @returns {Object|null} ユーザー情報
 */
function authenticateUser() {
  try {
    const email = SessionService.getActiveUserEmail();
    if (!email) return null;
    
    const userId = getUserIdFromEmail(email);
    return getUserInfo(userId);
  } catch (error) {
    logError(error, 'authenticateUser', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.AUTHENTICATION);
    return null;
  }
}

/**
 * リクエストパラメータを解析
 * @param {Object} e - リクエストイベント
 * @returns {Object} パラメータ
 */
function parseRequestParams(e) {
  const params = e.parameter || {};
  return {
    mode: params.mode || '',
    userId: params.userId || '',
    action: params.action || '',
    data: params.data || ''
  };
}

/**
 * セットアップページを表示
 * @returns {GoogleAppsScript.HTML.HtmlOutput}
 */
function showSetupPage() {
  const template = Html.createTemplateFromFile('SetupPage');
  return template.evaluate()
    .setTitle('みんなの回答ボード - 初期設定')
    .setFaviconUrl('https://www.google.com/s2/favicons?domain=google.com')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * ログインページを表示
 * @returns {GoogleAppsScript.HTML.HtmlOutput}
 */
function showLoginPage() {
  const template = Html.createTemplateFromFile('LoginPage');
  return template.evaluate()
    .setTitle('みんなの回答ボード - ログイン')
    .setFaviconUrl('https://www.google.com/s2/favicons?domain=google.com')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * アプリセットアップページを表示
 * @param {string} userId - ユーザーID
 * @returns {GoogleAppsScript.HTML.HtmlOutput}
 */
function showAppSetupPage(userId) {
  const template = Html.createTemplateFromFile('AppSetupPage');
  template.userId = userId;
  return template.evaluate()
    .setTitle('みんなの回答ボード - アプリ設定')
    .setFaviconUrl('https://www.google.com/s2/favicons?domain=google.com')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 管理パネルを表示
 * @param {Object} user - ユーザー情報
 * @param {Object} params - パラメータ
 * @returns {GoogleAppsScript.HTML.HtmlOutput}
 */
function showAdminPanel(user, params) {
  const template = Html.createTemplateFromFile('AdminPanel');
  template.user = user;
  template.params = params;
  template.isDeployUser = isDeployUser();
  template.escapeJavaScript = escapeJavaScript;
  template.shouldEnableDebugMode = shouldEnableDebugMode;
  
  return template.evaluate()
    .setTitle('みんなの回答ボード - 管理パネル')
    .setFaviconUrl('https://www.google.com/s2/favicons?domain=google.com')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

/**
 * メインページを表示
 * @param {Object} user - ユーザー情報
 * @param {Object} params - パラメータ
 * @returns {GoogleAppsScript.HTML.HtmlOutput}
 */
function showMainPage(user, params) {
  const template = Html.createTemplateFromFile('Page');
  template.user = user;
  template.params = params;
  template.escapeJavaScript = escapeJavaScript;
  
  return template.evaluate()
    .setTitle('みんなの回答ボード')
    .setFaviconUrl('https://www.google.com/s2/favicons?domain=google.com')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

/**
 * 未公開ページを表示
 * @param {Object} params - パラメータ
 * @returns {GoogleAppsScript.HTML.HtmlOutput}
 */
function showUnpublishedPage(params) {
  const template = Html.createTemplateFromFile('Unpublished');
  template.params = params;
  
  return template.evaluate()
    .setTitle('みんなの回答ボード - 未公開')
    .setFaviconUrl('https://www.google.com/s2/favicons?domain=google.com')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * エラーページを表示
 * @param {Error} error - エラー
 * @returns {GoogleAppsScript.HTML.HtmlOutput}
 */
function showErrorPage(error) {
  const template = Html.createTemplateFromFile('ErrorBoundary');
  template.errorMessage = error.message || '予期しないエラーが発生しました';
  template.errorDetails = error.stack || '';
  
  return template.evaluate()
    .setTitle('みんなの回答ボード - エラー')
    .setFaviconUrl('https://www.google.com/s2/favicons?domain=google.com')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * JavaScript文字列をエスケープ
 * @param {string} str - エスケープする文字列
 * @returns {string}
 */
function escapeJavaScript(str) {
  if (!str) return '';
  
  return String(str)
    .replace(/\\/g, '\\\\')
    .replace(/'/g, "\\'")
    .replace(/"/g, '\\"')
    .replace(/\n/g, '\\n')
    .replace(/\r/g, '\\r')
    .replace(/\t/g, '\\t');
}

/**
 * APIリクエストをハンドリング
 * @param {string} action - アクション
 * @param {Object} params - パラメータ
 * @returns {Object} レスポンス
 */
function handleApiRequest(action, params) {
  // APIハンドラーに委譲（Core.gsまたは新しいAPIモジュールで処理）
  if (typeof handleCoreApiRequest === 'function') {
    return handleCoreApiRequest(action, params);
  }
  
  throw new Error(`Unknown API action: ${action}`);
}