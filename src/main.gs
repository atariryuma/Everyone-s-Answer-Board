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

/* global UserService, ConfigService, DataService, SecurityService, ErrorHandler, DB, PROPS_KEYS, handleGetData, handleAddReaction, handleToggleHighlight, handleRefreshData, AdminController, FrontendController, SystemController, ResponseFormatter, AppCacheService, getAdminSpreadsheetList, addDataReaction, toggleDataHighlight, getConfig, checkIsSystemAdmin, getCurrentBoardInfoAndUrls */

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
// getUser is implemented in FrontendController.gs

// processLoginAction is implemented in FrontendController.gs

/**
 * System Controller Global Functions
 */
// getSystemDomainInfo is implemented in SystemController.gs

// forceUrlSystemReset is implemented in SystemController.gs

// setupApplication is implemented in SystemController.gs

// testSetup is implemented in SystemController.gs

function getApplicationStatusForUI() {
  try {
    return ConfigService.getApplicationStatus();
  } catch (error) {
    console.error('getApplicationStatusForUI error:', error);
    return { success: false, error: error.message };
  }
}

// getAllUsersForAdminForUI is implemented in DataController.gs

// deleteUserAccountByAdminForUI is implemented in DataController.gs

// getWebAppUrl is implemented in SystemController.gs

// reportClientError is implemented in FrontendController.gs

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

// testSystemDiagnosis is implemented in SystemController.gs

// performAutoRepair is implemented in SystemController.gs

// performSystemMonitoring is implemented in SystemController.gs

// performDataIntegrityCheck is implemented in SystemController.gs

// testForceLogoutRedirect is implemented in FrontendController.gs

// verifyUserAuthentication is implemented in FrontendController.gs

// resetAuth is implemented in FrontendController.gs

/**
 * Admin Controller Global Functions
 */
// getConfig is implemented in AdminController.gs

function getSpreadsheetList() {
  try {
    return getAdminSpreadsheetList();
  } catch (error) {
    console.error('getSpreadsheetList error:', error);
    return {
      success: false,
      message: error.message || 'スプレッドシート一覧取得エラー',
      spreadsheets: []
    };
  }
}

// getLightweightHeaders is implemented in AdminController.gs

// analyzeColumns is implemented in AdminController.gs

// publishApplication is implemented in AdminController.gs

// saveDraftConfiguration is implemented in AdminController.gs

function connectDataSource(spreadsheetId, sheetName) {
  try {
    // ✅ GAS Best Practice: Direct service calls
    return ConfigService.connectDataSource(spreadsheetId, sheetName);
  } catch (error) {
    console.error('connectDataSource error:', error);
    return { success: false, error: error.message };
  }
}

// checkIsSystemAdmin is implemented in AdminController.gs

function getSheetList(spreadsheetId) {
  try {
    // ✅ GAS Best Practice: Direct service calls
    return getSheetList(spreadsheetId);
  } catch (error) {
    console.error('getSheetList error:', error);
    return { success: false, error: error.message };
  }
}

// validateAccess is implemented in AdminController.gs

// getCurrentBoardInfoAndUrls is implemented in AdminController.gs

// getFormInfo is implemented in AdminController.gs

// checkCurrentPublicationStatus is implemented in AdminController.gs

// createForm is implemented in AdminController.gs

/**
 * Data Controller Global Functions
 */
// handleGetData is implemented in DataController.gs

function addReaction(userId, rowId, reactionType) {
  try {
    return addDataReaction(userId, rowId, reactionType);
  } catch (error) {
    console.error('addReaction error:', error);
    return { success: false, error: error.message };
  }
}

function toggleHighlight(userId, rowId) {
  try {
    return toggleDataHighlight(userId, rowId);
  } catch (error) {
    console.error('toggleHighlight error:', error);
    return { success: false, error: error.message };
  }
}

// refreshBoardData is implemented in DataController.gs

// addSpreadsheetUrl is implemented in DataController.gs

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

// getPublishedSheetData is implemented in DataController.gs

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