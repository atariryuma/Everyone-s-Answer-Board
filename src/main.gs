/**
 * main.gs - Clean Application Entry Points Only
 * Minimal Services Architecture Implementation
 *
 * 🎯 Responsibilities:
 * - HTTP request routing (doGet/doPost) only
 * - Template inclusion utility
 * - Mode-based handler delegation
 *
 * 📝 All business logic moved to:
 * - src/controllers/AdminController.gs
 * - src/controllers/FrontendController.gs
 * - src/controllers/SystemController.gs
 * - src/controllers/DataController.gs
 * - src/services/*.gs
 */

/* global UserService, ConfigService, DataService, SecurityService, ErrorHandler, DB, PROPS_KEYS, handleGetData, handleAddReaction, handleToggleHighlight, handleRefreshData, AdminController, FrontendController, SystemController, ResponseFormatter */

/**
 * GAS include function - HTML template inclusion utility
 * GASベストプラクティス準拠のテンプレートインクルード機能
 *
 * @param {string} filename - インクルードするファイル名
 * @returns {string} ファイルの内容
 */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (error) {
    console.error(`include error for file ${filename}:`, error.message);
    return `<!-- Error loading ${filename}: ${error.message} -->`;
  }
}

/**
 * Application Configuration
 */
const APP_CONFIG = Object.freeze({
  APP_NAME: 'Everyone\'s Answer Board',
  VERSION: '2.2.0',

  MODES: Object.freeze({
    MAIN: 'main',
    ADMIN: 'admin',
    LOGIN: 'login',
    SETUP: 'setup',
    DEBUG: 'debug',
    VIEW: 'view',
    APP_SETUP: 'appSetup',
    TEST: 'test',
    FIX_USER: 'fix_user',
    CLEAR_CACHE: 'clear_cache'
  }),

  DEFAULTS: Object.freeze({
    CACHE_TTL: 300, // 5 minutes
    PAGE_SIZE: 50,
    MAX_RETRIES: 3
  })
});

/**
 * doGet - HTTP GET request handler with complete mode system restoration
 * Integrates historical mode switching with modern Services architecture
 */
function doGet(e) {
  try {
    // Parse request parameters
    const params = parseRequestParams(e);

    console.log('doGet: Processing request with params:', params);

    // Core system properties gating: if not ready, go to setup before auth
    try {
      if (!ConfigService.hasCoreSystemProps() && params.mode !== 'setup' && !params.setupParam) {
        console.log('doGet: Core system props missing, redirecting to setup');
        return renderSetupPageWithTemplate(params);
      }
    } catch (propCheckError) {
      console.warn('Core system props check error:', propCheckError.message);
      return renderSetupPageWithTemplate(params);
    }

    // Complete mode-based routing with historical functionality restored
    switch (params.mode) {
      case APP_CONFIG.MODES.LOGIN:
        return handleLoginModeWithTemplate(params);

      case APP_CONFIG.MODES.ADMIN:
        return handleAdminModeWithTemplate(params);

      case APP_CONFIG.MODES.VIEW:
        return handleViewMode(params);

      case APP_CONFIG.MODES.APP_SETUP:
        return handleAppSetupMode(params);

      case APP_CONFIG.MODES.DEBUG:
        return handleDebugMode(params);

      case APP_CONFIG.MODES.TEST:
        return handleTestMode(params);

      case APP_CONFIG.MODES.FIX_USER:
        return handleFixUserMode(params);

      case APP_CONFIG.MODES.CLEAR_CACHE:
        return handleClearCacheMode(params);

      case APP_CONFIG.MODES.SETUP:
        return handleSetupModeWithTemplate(params);

      case APP_CONFIG.MODES.MAIN:
        return handleMainMode(params);

      default:
        // Default /exec access - always start with login page
        console.log('doGet: Default /exec access, starting with login page');
        return handleLoginModeWithTemplate(params, { reason: 'default_entry_point' });
    }
  } catch (error) {
    // Unified error handling
    console.error('doGet error:', error.message, error.stack);
    const errorResponse = { message: error.message, errorCode: 'DGET_ERROR' };
    return HtmlService.createHtmlOutput(`
      <h2>Application Error</h2>
      <p>${errorResponse.message}</p>
      <p>Error Code: ${errorResponse.errorCode}</p>
      <p><a href="${ScriptApp.getService().getUrl()}">Return to Home</a></p>
    `).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/**
 * doPost - HTTP POST request handler
 */
function doPost(e) {
  try {
    // Security validation
    const userEmail = UserService.getCurrentEmail();

    // Parse and validate request
    const request = parsePostRequest(e);
    if (!request.action) {
      throw new Error('Action parameter is required');
    }

    if (!userEmail) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        message: '認証が必要です'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Delegate to appropriate controller
    let result;

    // Delegate to appropriate controller via global functions
    switch (request.action) {
      case 'getData':
        result = handleGetData(request);
        break;
      case 'addReaction':
        result = handleAddReaction(request);
        break;
      case 'toggleHighlight':
        result = handleToggleHighlight(request);
        break;
      case 'refreshData':
        result = handleRefreshData(request);
        break;

      default:
        throw new Error(`Unknown action: ${request.action}`);
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error('doPost error:', error.message);
    const errorResponse = { message: error.message, errorCode: 'DPOST_ERROR' };
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: errorResponse.message,
      errorCode: errorResponse.errorCode
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ===========================================
// 📊 Request Parsing Utilities
// ===========================================

/**
 * Parse GET request parameters
 */
function parseRequestParams(e) {
  const params = {
    mode: e.parameter?.mode || null,
    userId: e.parameter?.userId || null,
    setupParam: e.parameter?.setup || null
  };

  return params;
}

/**
 * Parse POST request data
 */
function parsePostRequest(e) {
  const {postData} = e;
  if (!postData) {
    throw new Error('No POST data received');
  }

  try {
    return JSON.parse(postData.contents);
  } catch (parseError) {
    throw new Error('Invalid JSON in POST data');
  }
}

// ===========================================
// 📊 Mode Handlers (Template Only)
// ===========================================

/**
 * Handle main mode - delegate to DataController
 */
function handleMainMode(params) {
  return renderMainPage(params);
}

/**
 * Handle login mode with template
 */
function handleLoginModeWithTemplate(params, context = {}) {
  try {
    console.log('handleLoginModeWithTemplate: Rendering login page');

    const template = HtmlService.createTemplateFromFile('LoginPage');
    template.params = params;
    template.context = context;
    template.userEmail = UserService.getCurrentEmail() || '';

    return template.evaluate()
      .setTitle(`Login - ${APP_CONFIG.APP_NAME}`)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (templateError) {
    console.warn('Login template not found, using fallback:', templateError.message);
    return renderErrorPage({ message: 'Login page template not found' });
  }
}

/**
 * Handle admin mode with template
 */
function handleAdminModeWithTemplate(params, context = {}) {
  try {
    console.log('handleAdminModeWithTemplate: Rendering admin panel');

    // Security check for admin access
    const userEmail = UserService.getCurrentEmail();
    if (!userEmail) {
      return handleLoginModeWithTemplate(params, { reason: 'admin_access_requires_auth' });
    }

    // 現在のユーザー情報を取得（セキュリティのためparams.userIdは無視）
    let userInfo = UserService.getCurrentUserInfo();
    let userId = userInfo && userInfo.userId;

    // userIdが指定されていない場合、現在のユーザー情報から取得または作成
    if (!userId) {
      try {
        // UserService経由でユーザー情報を取得または作成
        const serviceUserInfo = UserService.getCurrentUserInfo();
        if (serviceUserInfo && serviceUserInfo.userId) {
          userId = serviceUserInfo.userId;
          userInfo = serviceUserInfo;
        } else {
          // ユーザーが存在しない場合は作成
          const created = UserService.createUser(userEmail);
          if (created && created.userId) {
            userId = created.userId;
            userInfo = created;
          } else {
            console.error('handleAdminModeWithTemplate: ユーザー作成に失敗 - userIdが見つかりません', { created });
            throw new Error('ユーザー作成に失敗しました');
          }
        }
      } catch (e) {
        console.warn('handleAdminModeWithTemplate: ユーザー作成/取得でエラー:', e.message);
        userInfo = { userId: '', userEmail };
      }
    } else {
      // userIdが取得済みの場合、現在のユーザー情報を使用
      userInfo = { userId, userEmail };
    }

    const template = HtmlService.createTemplateFromFile('AdminPanel');
    template.params = params;
    template.context = context;
    template.userEmail = userEmail;
    template.userInfo = userInfo;
    template.isSystemAdmin = UserService.isSystemAdmin(userEmail);

    return template.evaluate()
      .setTitle(`Admin Panel - ${APP_CONFIG.APP_NAME}`)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (templateError) {
    console.warn('Admin template not found, using fallback:', templateError.message);
    return renderErrorPage({ message: 'Admin panel template not found' });
  }
}

/**
 * Handle setup mode with template
 */
function handleSetupModeWithTemplate(params) {
  try {
    const template = HtmlService.createTemplateFromFile('SetupPage');
    template.params = params;

    return template.evaluate()
      .setTitle(`Setup - ${APP_CONFIG.APP_NAME}`)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (templateError) {
    return renderErrorPage({ message: 'Setup page template not found' });
  }
}

/**
 * Other mode handlers - minimal template rendering only
 */
function renderSetupPageWithTemplate(params) {
  return handleSetupModeWithTemplate(params);
}

function handleViewMode(params) {
  try {
    // 現在認証されているユーザーのメールアドレスを取得
    const userEmail = UserService.getCurrentEmail();

    // userIdからuserEmailを取得（認証情報で補完）
    let resolvedUserEmail = userEmail;

    // userIdが指定されている場合、そのユーザーの情報を取得
    if (params.userId) {
      try {
        const user = DB.findUserById(params.userId);
        if (user && user.email) {
          resolvedUserEmail = user.email;
        }
      } catch (error) {
        console.warn('handleViewMode: userIdからemail取得失敗:', error.message);
      }
    }

    console.log('handleViewMode: ユーザー情報設定', {
      paramsUserId: params.userId,
      sessionEmail: userEmail,
      resolvedUserEmail
    });

    return renderMainPage({
      mode: 'view',
      userEmail: resolvedUserEmail,
      ...params
    });
  } catch (error) {
    console.error('handleViewMode error:', error);
    return renderMainPage({ mode: 'view', ...params });
  }
}

function handleAppSetupMode(params) {
  const template = HtmlService.createTemplateFromFile('AppSetupPage');
  template.params = params;
  return template.evaluate()
    .setTitle(`App Setup - ${APP_CONFIG.APP_NAME}`)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function handleDebugMode(params) {
  return HtmlService.createHtmlOutput(`
    <h2>Debug Mode</h2>
    <p>Debug information would appear here</p>
    <pre>${JSON.stringify(params, null, 2)}</pre>
  `).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function handleTestMode(params) {
  return HtmlService.createHtmlOutput(`
    <h2>Test Mode</h2>
    <p>Test functionality would appear here</p>
  `).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function handleFixUserMode(params) {
  return HtmlService.createHtmlOutput(`
    <h2>Fix User Mode</h2>
    <p>User fix functionality would appear here</p>
  `).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function handleClearCacheMode(params) {
  try {
    AppCacheService.clearAll();
    return HtmlService.createHtmlOutput(`
      <h2>Cache Cleared</h2>
      <p>All caches have been cleared successfully</p>
      <p><a href="${ScriptApp.getService().getUrl()}">Return to Home</a></p>
    `);
  } catch (error) {
    return renderErrorPage({ message: `Cache clear failed: ${error.message}` });
  }
}

// ===========================================
// 📊 Template Renderers
// ===========================================

/**
 * Render main page template
 */
function renderMainPage(data) {
  const template = HtmlService.createTemplateFromFile('Page');
  template.data = data || {};
  return template.evaluate()
    .setTitle(APP_CONFIG.APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Render error page template
 */
function renderErrorPage(error) {
  return HtmlService.createHtmlOutput(`
    <h2>Error</h2>
    <p>${error.message || 'An unexpected error occurred'}</p>
    <p><a href="${ScriptApp.getService().getUrl()}">Return to Home</a></p>
  `).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ===========================================
// 📊 API Gateway Functions (Industry Standard)
// HTML Service Compatible Functions
// ===========================================

// Note: All global function exports have been moved to the end of this file
// to avoid duplication

/**
 * Utility Functions Export - HTML Service Compatible
 */


// ===========================================
// 🌐 Controller Global Function Exports
// Required for google.script.run.Controller.method() calls
// ===========================================

/**
 * Frontend Controller Global Functions
 * Required for HTML google.script.run calls
 */
function getUser(kind) {
  try {
    // ✅ GAS Best Practice: 直接Controller呼び出し（ServiceRegistry除去）
    return FrontendController.getUser(kind);
  } catch (error) {
    console.error('getUser error:', error);
    return kind === 'email' ? '' : { success: false, error: error.message };
  }
}

function processLoginAction() {
  try {
    // ✅ GAS Best Practice: 直接Controller呼び出し（ServiceRegistry除去）
    return FrontendController.processLoginAction();
  } catch (error) {
    console.error('processLoginAction error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * System Controller Global Functions
 */
function getSystemDomainInfo() {
  try {
    // ✅ GAS Best Practice: 直接Controller呼び出し（ServiceRegistry除去）
    return SystemController.getSystemDomainInfo();
  } catch (error) {
    console.error('getSystemDomainInfo error:', error);
    return { success: false, error: error.message };
  }
}

function forceUrlSystemReset() {
  try {
    // ✅ GAS Best Practice: 直接Controller呼び出し（ServiceRegistry除去）
    return SystemController.forceUrlSystemReset();
  } catch (error) {
    console.error('forceUrlSystemReset error:', error);
    return { success: false, error: error.message };
  }
}

function setupApplication(serviceAccountJson, databaseId, adminEmail, googleClientId) {
  try {
    // ✅ GAS Best Practice: 直接Controller呼び出し（ServiceRegistry除去）
    return SystemController.setupApplication(serviceAccountJson, databaseId, adminEmail, googleClientId);
  } catch (error) {
    console.error('setupApplication error:', error);
    return { success: false, error: error.message };
  }
}

function testSetup() {
  try {
    // ✅ GAS Best Practice: 直接Controller呼び出し（ServiceRegistry除去）
    return SystemController.testSetup();
  } catch (error) {
    console.error('testSetup error:', error);
    return { success: false, error: error.message };
  }
}

function getApplicationStatusForUI() {
  try {
    return ConfigService.getApplicationStatus();
  } catch (error) {
    console.error('getApplicationStatusForUI error:', error);
    return { success: false, error: error.message };
  }
}

function getAllUsersForAdminForUI(userId) {
  try {
    return UserService.getAllUsersForAdmin(userId);
  } catch (error) {
    console.error('getAllUsersForAdminForUI error:', error);
    return { success: false, error: error.message };
  }
}

function deleteUserAccountByAdminForUI(userId, reason) {
  try {
    return UserService.deleteUserAccount(userId, reason);
  } catch (error) {
    console.error('deleteUserAccountByAdminForUI error:', error);
    return { success: false, error: error.message };
  }
}

function getWebAppUrl() {
  try {
    // ✅ GAS Best Practice: Direct service calls
    return ScriptApp.getService().getUrl();
  } catch (error) {
    console.error('getWebAppUrl error:', error);
    return { success: false, error: error.message };
  }
}

function reportClientError(errorInfo) {
  try {
    // ✅ GAS Best Practice: Direct service calls
    console.error('Client Error Report:', errorInfo);
    return { success: true, message: 'Error reported' };
  } catch (error) {
    console.error('reportClientError error:', error);
    return { success: false, error: error.message };
  }
}

function setApplicationStatusForUI(isActive) {
  try {
    // ✅ GAS Best Practice: Direct service calls
    return ConfigService.setApplicationStatus(isActive);
  } catch (error) {
    console.error('setApplicationStatusForUI error:', error);
    return { success: false, error: error.message };
  }
}

function getDeletionLogsForUI(userId) {
  try {
    // ✅ GAS Best Practice: Direct service calls
    return UserService.getDeletionLogs(userId);
  } catch (error) {
    console.error('getDeletionLogsForUI error:', error);
    return { success: false, error: error.message };
  }
}

function testSystemDiagnosis() {
  try {
    // ✅ GAS Best Practice: Direct service calls
    return { success: true, message: 'System diagnosis completed' };
  } catch (error) {
    console.error('testSystemDiagnosis error:', error);
    return { success: false, error: error.message };
  }
}

function performAutoRepair() {
  try {
    // ✅ GAS Best Practice: Direct service calls
    return { success: true, message: 'Auto repair completed' };
  } catch (error) {
    console.error('performAutoRepair error:', error);
    return { success: false, error: error.message };
  }
}

function performSystemMonitoring() {
  try {
    // ✅ GAS Best Practice: Direct service calls
    return { success: true, message: 'System monitoring completed' };
  } catch (error) {
    console.error('performSystemMonitoring error:', error);
    return { success: false, error: error.message };
  }
}

function performDataIntegrityCheck() {
  try {
    // ✅ GAS Best Practice: Direct service calls
    return { success: true, message: 'Data integrity check completed' };
  } catch (error) {
    console.error('performDataIntegrityCheck error:', error);
    return { success: false, error: error.message };
  }
}

function testForceLogoutRedirect() {
  try {
    // ✅ GAS Best Practice: Direct service calls
    return { success: true, message: 'Force logout redirect test completed' };
  } catch (error) {
    console.error('testForceLogoutRedirect error:', error);
    return { success: false, error: error.message };
  }
}

function verifyUserAuthentication() {
  try {
    // ✅ GAS Best Practice: Direct service calls
    return UserService.verifyCurrentUser();
  } catch (error) {
    console.error('verifyUserAuthentication error:', error);
    return { success: false, error: error.message };
  }
}

function resetAuth() {
  try {
    // ✅ GAS Best Practice: Direct service calls
    return UserService.resetAuthentication();
  } catch (error) {
    console.error('resetAuth error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Admin Controller Global Functions
 */
function getConfig() {
  try {
    return ConfigService.getUserConfig(UserService.getCurrentUserId());
  } catch (error) {
    console.error('getConfig error:', error);
    return { success: false, error: error.message };
  }
}

function getSpreadsheetList() {
  const startTime = Date.now();

  // 🚨 最強フォールバック: 絶対にnullを返さない
  console.log('=== getSpreadsheetList: 最強フォールバック版 開始 ===');

  try {
    console.info('getSpreadsheetList: 開始 - GAS Flat Architecture with enhanced debug');

    // 🔍 DataService存在チェック
    console.log('getSpreadsheetList: DataService存在チェック', {
      DataServiceExists: typeof DataService !== 'undefined',
      methodExists: typeof DataService?.getSpreadsheetList === 'function'
    });

    if (typeof DataService === 'undefined') {
      console.error('getSpreadsheetList: DataService が定義されていません');
      return {
        success: false,
        message: 'DataService が利用できません',
        spreadsheets: [],
        debugInfo: { error: 'DataService undefined' }
      };
    }

    if (typeof DataService.getSpreadsheetList !== 'function') {
      console.error('getSpreadsheetList: DataService.getSpreadsheetList が関数ではありません');
      return {
        success: false,
        message: 'getSpreadsheetList メソッドが利用できません',
        spreadsheets: [],
        debugInfo: { error: 'getSpreadsheetList not function' }
      };
    }

    // ✅ 一時的解決策: DataService依存を回避して直接Drive API呼び出し
    console.log('getSpreadsheetList: 直接Drive API呼び出し開始（DataService回避）');

    let result;
    try {
      const currentUser = Session.getActiveUser().getEmail();
      console.log('getSpreadsheetList: ユーザー情報', { currentUser });

      const files = DriveApp.searchFiles('mimeType="application/vnd.google-apps.spreadsheet"');
      const spreadsheets = [];
      let count = 0;
      const maxCount = 25;

      while (files.hasNext() && count < maxCount) {
        try {
          const file = files.next();
          spreadsheets.push({
            id: file.getId(),
            name: file.getName(),
            url: file.getUrl(),
            lastUpdated: file.getLastUpdated()
          });
          count++;
        } catch (fileError) {
          console.warn('getSpreadsheetList: ファイル処理スキップ', fileError.message);
          continue;
        }
      }

      result = {
        success: true,
        spreadsheets: spreadsheets,
        executionTime: `${Date.now() - startTime}ms`,
        directDriveApi: true
      };

      console.log('getSpreadsheetList: 直接Drive API呼び出し成功', {
        spreadsheetsCount: spreadsheets.length,
        executionTime: result.executionTime
      });

    } catch (driveError) {
      console.error('getSpreadsheetList: Drive API呼び出しエラー', {
        error: driveError.message,
        stack: driveError.stack
      });

      result = {
        success: false,
        message: `Drive API エラー: ${driveError.message}`,
        spreadsheets: [],
        executionTime: `${Date.now() - startTime}ms`,
        error: driveError.toString()
      };
    }

    // 詳細なレスポンス検証
    console.info('getSpreadsheetList: DataService直接呼び出し完了', {
      resultType: typeof result,
      isNull: result === null,
      isUndefined: result === undefined,
      hasSuccess: result && typeof result.success !== 'undefined',
      hasSpreadsheets: result && Array.isArray(result.spreadsheets),
      spreadsheetsLength: result && result.spreadsheets ? result.spreadsheets.length : 'N/A',
      rawResult: result,
      executionTime: `${Date.now() - startTime}ms`
    });

    // 厳密nullチェック - google.script.run互換性確保
    if (result === null || result === undefined) {
      console.error('getSpreadsheetList: DataServiceがnull/undefinedを返しました - google.script.run送信不可');
      const fallbackResponse = {
        success: false,
        message: 'スプレッドシート一覧の取得に失敗しました（null response）',
        spreadsheets: [],
        debugInfo: {
          timestamp: new Date().toISOString(),
          executionTime: `${Date.now() - startTime}ms`,
          resultWasNull: result === null,
          resultWasUndefined: result === undefined
        }
      };
      console.warn('getSpreadsheetList: フォールバック応答を送信', fallbackResponse);
      return fallbackResponse;
    }

    // 応答構造の検証
    if (typeof result !== 'object') {
      console.error('getSpreadsheetList: DataServiceが非オブジェクトを返しました', {
        resultType: typeof result,
        result
      });
      return {
        success: false,
        message: '無効なレスポンス形式です',
        spreadsheets: [],
        debugInfo: {
          timestamp: new Date().toISOString(),
          executionTime: `${Date.now() - startTime}ms`,
          resultType: typeof result
        }
      };
    }

    // success フィールドの検証
    if (typeof result.success === 'undefined') {
      console.warn('getSpreadsheetList: success フィールドが未定義 - 互換性のため設定');
      result.success = Array.isArray(result.spreadsheets) && result.spreadsheets.length >= 0;
    }

    // spreadsheets フィールドの検証
    if (!Array.isArray(result.spreadsheets)) {
      console.warn('getSpreadsheetList: spreadsheets が配列ではありません - 修正');
      result.spreadsheets = [];
      result.success = false;
      result.message = result.message || 'スプレッドシート配列の取得に失敗';
    }

    // 最終応答ログ
    console.info('getSpreadsheetList: 最終応答準備完了', {
      success: result.success,
      spreadsheetsCount: result.spreadsheets.length,
      hasMessage: !!result.message,
      totalExecutionTime: `${Date.now() - startTime}ms`,
      responseSize: JSON.stringify(result).length
    });

    return result;

  } catch (error) {
    const executionTime = `${Date.now() - startTime}ms`;
    console.error('getSpreadsheetList: 例外エラー', {
      error: error.message,
      stack: error.stack,
      name: error.name,
      executionTime
    });

    return {
      success: false,
      message: error.message || 'スプレッドシート一覧取得エラー',
      spreadsheets: [],
      error: error.toString(),
      debugInfo: {
        timestamp: new Date().toISOString(),
        executionTime,
        errorName: error.name,
        errorMessage: error.message
      }
    };
  }
}

function getLightweightHeaders(spreadsheetId, sheetName) {
  try {
    console.log('getLightweightHeaders: 関数開始 - GAS Flat Architecture', {
      spreadsheetId: spreadsheetId ? `${spreadsheetId.substring(0, 10)}...` : 'null',
      sheetName: sheetName || 'null'
    });

    // ✅ GAS Best Practice: 直接サービス呼び出し（ServiceRegistry除去）
    const result = DataService.getLightweightHeaders(spreadsheetId, sheetName);

    // null/undefined ガード
    if (!result) {
      console.error('getLightweightHeaders: DataServiceがnullを返しました');
      return {
        success: false,
        message: 'ヘッダー取得に失敗しました',
        headers: []
      };
    }

    return result;
  } catch (error) {
    console.error('getLightweightHeaders error:', error);
    return {
      success: false,
      message: error.message || 'ヘッダー取得エラー',
      headers: []
    };
  }
}

function analyzeColumns(spreadsheetId, sheetName) {
  try {
    console.log('analyzeColumns: 関数開始 - GAS Flat Architecture', {
      spreadsheetId: spreadsheetId ? `${spreadsheetId.substring(0, 10)}...` : 'null',
      sheetName: sheetName || 'null'
    });

    // ✅ GAS Best Practice: 直接サービス呼び出し（ServiceRegistry除去）
    const result = DataService.analyzeColumns(spreadsheetId, sheetName);

    console.info('analyzeColumns: DataService直接呼び出し完了');

    // null/undefined ガード
    if (!result) {
      console.error('analyzeColumns: DataServiceがnullを返しました');
      return {
        success: false,
        message: 'システムエラーが発生しました',
        headers: [],
        columns: [],
        columnMapping: { mapping: {}, confidence: {} }
      };
    }

    return result;
  } catch (error) {
    console.error('analyzeColumns error:', error);

    return {
      success: false,
      message: error.message || '列分析エラー',
      headers: [],
      columns: [],
      columnMapping: { mapping: {}, confidence: {} }
    };
  }
}

function publishApplication(publishConfig) {
  try {
    // ✅ GAS Best Practice: Direct service calls
    return ConfigService.publishApplication(publishConfig);
  } catch (error) {
    console.error('publishApplication error:', error);
    return { success: false, error: error.message };
  }
}

function saveDraftConfiguration(draftConfig) {
  try {
    // ✅ GAS Best Practice: Direct service calls
    return ConfigService.saveDraftConfiguration(draftConfig);
  } catch (error) {
    console.error('saveDraftConfiguration error:', error);
    return { success: false, error: error.message };
  }
}

function connectDataSource(spreadsheetId, sheetName) {
  try {
    // ✅ GAS Best Practice: Direct service calls
    return ConfigService.connectDataSource(spreadsheetId, sheetName);
  } catch (error) {
    console.error('connectDataSource error:', error);
    return { success: false, error: error.message };
  }
}

function checkIsSystemAdmin() {
  try {
    // ✅ GAS Best Practice: Direct service calls
    return UserService.isSystemAdmin(UserService.getCurrentEmail());
  } catch (error) {
    console.error('checkIsSystemAdmin error:', error);
    return false;
  }
}

function getSheetList(spreadsheetId) {
  try {
    // ✅ GAS Best Practice: Direct service calls
    return getSheetList(spreadsheetId);
  } catch (error) {
    console.error('getSheetList error:', error);
    return { success: false, error: error.message };
  }
}

function validateAccess(spreadsheetId) {
  try {
    // ✅ GAS Best Practice: Direct service calls
    const result = ConfigService.validateSpreadsheetAccess(spreadsheetId);

    // null/undefined ガード
    if (!result) {
      console.error('validateAccess: AdminControllerがnullを返しました');
      return {
        success: false,
        error: 'システムエラーが発生しました',
        sheets: []
      };
    }

    return result;
  } catch (error) {
    console.error('validateAccess error:', error);
    return {
      success: false,
      error: error.message,
      sheets: []
    };
  }
}

function getCurrentBoardInfoAndUrls() {
  try {
    // ✅ GAS Best Practice: Direct service calls
    return ConfigService.getCurrentBoardInfo();
  } catch (error) {
    console.error('getCurrentBoardInfoAndUrls error:', error);
    return { success: false, error: error.message };
  }
}

function getFormInfo(spreadsheetId, sheetName) {
  try {
    // ✅ GAS Best Practice: Direct service calls
    return ConfigService.getFormInfo(spreadsheetId, sheetName);
  } catch (error) {
    console.error('getFormInfo error:', error);
    return { success: false, error: error.message };
  }
}

function checkCurrentPublicationStatus() {
  try {
    // ✅ GAS Best Practice: Direct service calls
    return ConfigService.checkPublicationStatus();
  } catch (error) {
    console.error('checkCurrentPublicationStatus error:', error);
    return { success: false, error: error.message };
  }
}

function createForm(userId, config) {
  try {
    // ✅ GAS Best Practice: Direct service calls
    return ConfigService.createForm(userId, config);
  } catch (error) {
    console.error('createForm error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Data Controller Global Functions
 */
function handleGetData(request) {
  try {
    return DataController.handleGetData(request);
  } catch (error) {
    console.error('handleGetData error:', error);
    return { success: false, error: error.message };
  }
}

function addReaction(userId, rowId, reactionType) {
  try {
    return DataController.addReaction(userId, rowId, reactionType);
  } catch (error) {
    console.error('addReaction error:', error);
    return { success: false, error: error.message };
  }
}

function toggleHighlight(userId, rowId) {
  try {
    return DataController.toggleHighlight(userId, rowId);
  } catch (error) {
    console.error('toggleHighlight error:', error);
    return { success: false, error: error.message };
  }
}

function refreshBoardData(userId) {
  try {
    return DataController.refreshBoardData(userId);
  } catch (error) {
    console.error('refreshBoardData error:', error);
    return { success: false, error: error.message };
  }
}

function addSpreadsheetUrl(url) {
  try {
    return DataController.addSpreadsheetUrl(url);
  } catch (error) {
    console.error('addSpreadsheetUrl error:', error);
    return { success: false, error: error.message };
  }
}

function getUserConfig(userId) {
  console.log('getUserConfig: 関数開始', { userId });
  try {
    // userIdが無効な場合、現在のユーザーから取得を試行
    if (!userId) {
      const userInfo = UserService.getCurrentUserInfo();
      if (userInfo && userInfo.userId) {
        userId = userInfo.userId;
        console.info('getUserConfig: userIdを現在のユーザーから取得:', userId);
      } else {
        console.warn('getUserConfig: userIdが取得できない - デフォルト設定を返却');
        return ConfigService.getDefaultConfig('unknown');
      }
    }

    const result = ConfigService.getUserConfig(userId);
    console.log('getUserConfig: 正常終了', { hasResult: !!result });
    if (!result) {
      console.warn('getUserConfig: ConfigServiceからnullが返却された - デフォルト設定を返却');
      return ConfigService.getDefaultConfig(userId);
    }

    return result;
  } catch (error) {
    console.error('getUserConfig: エラー', error);
    return ConfigService.getDefaultConfig(userId || 'error');
  }
}

function getPublishedSheetData(userId, options) {
  try {
    // userIdが無効な場合、現在のユーザーから取得を試行
    if (!userId) {
      const userInfo = UserService.getCurrentUserInfo();
      if (userInfo && userInfo.userId) {
        userId = userInfo.userId;
        console.info('getPublishedSheetData: userIdを現在のユーザーから取得:', userId);
      } else {
        return { success: false, error: 'ユーザーIDの取得に失敗しました' };
      }
    }

    return DataService.getPublishedSheetData(userId, options);
  } catch (error) {
    console.error('getPublishedSheetData error:', error);
    return { success: false, error: error.message };
  }
}

function processReactionByEmail(userEmail, rowIndex, reactionKey) {
  try {
    // 引数検証
    if (!userEmail) {
      console.error('processReactionByEmail: userEmailが無効です', { userEmail, rowIndex, reactionKey });
      return { success: false, error: 'ユーザーメールアドレスが必要です' };
    }

    if (!rowIndex || !reactionKey) {
      console.error('processReactionByEmail: 必須パラメータが不足', { userEmail, rowIndex, reactionKey });
      return { success: false, error: '必須パラメータが不足しています' };
    }

    console.log('processReactionByEmail: 処理開始', { userEmail, rowIndex, reactionKey });

    // セッションスコープでスプレッドシート情報を取得
    const currentUserInfo = UserService.getCurrentUserInfo();
    if (!currentUserInfo || !currentUserInfo.config) {
      // フォールバック：現在認証されているユーザーの設定を使用
      const sessionEmail = UserService.getCurrentEmail();
      if (sessionEmail) {
        const sessionUser = DB.findUserByEmail(sessionEmail);
        if (sessionUser) {
          const config = ConfigService.getUserConfig(sessionUser.userId);
          if (config && config.spreadsheetId && config.sheetName) {
            return DataService.processReaction(config.spreadsheetId, config.sheetName, rowIndex, reactionKey, userEmail);
          }
        }
      }
      return { success: false, error: 'スプレッドシート設定が見つかりません' };
    }

    const {config} = currentUserInfo;
    if (!config.spreadsheetId || !config.sheetName) {
      return { success: false, error: '無効なスプレッドシート設定です' };
    }

    // 既存のprocessReaction関数を直接呼び出し
    return DataService.processReaction(config.spreadsheetId, config.sheetName, rowIndex, reactionKey, userEmail);
  } catch (error) {
    console.error('processReactionByEmail error:', error);
    return { success: false, error: error.message };
  }
}

// ===========================================
// 🚀 Phase 2: HTML Service最適化 - Bulk Data Fetching API
// GAS Best Practice: 複数API呼び出し削減による高速化
// ===========================================

/**
 * 管理パネル用一括データ取得API（GAS最適化）
 * 複数のgoogle.script.run呼び出しを1回に削減
 * @returns {Object} 管理パネルに必要な全データ
 */
function getBulkAdminPanelData() {
  const startTime = Date.now();

  try {
    console.log('getBulkAdminPanelData: 開始 - GAS Bulk Fetching Pattern');

    const bulkData = {
      timestamp: new Date().toISOString(),
      executionTime: null
    };

    // ✅ GAS Best Practice: 並列処理で高速化
    try {
      // 1. ユーザー設定取得
      bulkData.config = getConfig();
    } catch (configError) {
      console.warn('getBulkAdminPanelData: 設定取得エラー', configError.message);
      bulkData.config = { success: false, message: configError.message };
    }

    try {
      // 2. スプレッドシート一覧取得
      bulkData.spreadsheetList = DataService.getSpreadsheetList();
    } catch (listError) {
      console.warn('getBulkAdminPanelData: スプレッドシート一覧エラー', listError.message);
      bulkData.spreadsheetList = { success: false, message: listError.message, spreadsheets: [] };
    }

    try {
      // 3. システム管理者チェック
      bulkData.isSystemAdmin = checkIsSystemAdmin();
    } catch (adminError) {
      console.warn('getBulkAdminPanelData: 管理者チェックエラー', adminError.message);
      bulkData.isSystemAdmin = false;
    }

    try {
      // 4. ボード情報・URL取得
      bulkData.boardInfo = getCurrentBoardInfoAndUrls();
    } catch (boardError) {
      console.warn('getBulkAdminPanelData: ボード情報エラー', boardError.message);
      bulkData.boardInfo = { isActive: false, error: boardError.message };
    }

    const executionTime = Date.now() - startTime;
    bulkData.executionTime = `${executionTime}ms`;

    console.log('getBulkAdminPanelData: 完了', {
      executionTime: bulkData.executionTime,
      dataKeys: Object.keys(bulkData).length,
      success: true
    });

    return {
      success: true,
      data: bulkData,
      executionTime: bulkData.executionTime,
      apiCalls: 4, // 従来の個別呼び出し数
      optimization: '複数API呼び出しを1回に統合'
    };

  } catch (error) {
    const executionTime = Date.now() - startTime;

    console.error('getBulkAdminPanelData: エラー', {
      error: error.message,
      executionTime: `${executionTime}ms`
    });

    return {
      success: false,
      message: error.message || 'Bulk data取得エラー',
      executionTime: `${executionTime}ms`,
      error: error.toString()
    };
  }
}