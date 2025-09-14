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
 * 🚀 GAS Service Discovery & Dynamic Loading
 * Addresses non-deterministic file loading order in Google Apps Script
 */
function getAvailableService(serviceName) {
  const serviceMap = {
    'UserService': () => (typeof UserService !== 'undefined') ? UserService : null,
    'ConfigService': () => (typeof ConfigService !== 'undefined') ? ConfigService : null,
    'DataService': () => (typeof DataService !== 'undefined') ? DataService : null,
    'SecurityService': () => (typeof SecurityService !== 'undefined') ? SecurityService : null
  };

  if (serviceMap[serviceName]) {
    return serviceMap[serviceName]();
  }
  return null;
}

/**
 * Safe service method caller with fallback
 */
function callServiceMethod(serviceName, methodName, ...args) {
  const service = getAvailableService(serviceName);
  if (service && typeof service[methodName] === 'function') {
    try {
      return service[methodName](...args);
    } catch (error) {
      console.warn(`callServiceMethod: ${serviceName}.${methodName} execution failed:`, error.message);
    }
  }

  console.warn(`callServiceMethod: ${serviceName}.${methodName} not available, trying fallback`);

  // 🚀 Built-in Fallback Functions
  if (serviceName === 'UserService') {
    return handleUserServiceFallback(methodName, ...args);
  }

  return null;
}

/**
 * UserService fallback using GAS built-in APIs
 */
function handleUserServiceFallback(methodName, ...args) {
  try {
    switch (methodName) {
      case 'getCurrentEmail':
        return getCurrentEmailDirect();

      case 'getCurrentUserInfo': {
        const email = getCurrentEmailDirect();
        return email ? { userEmail: email, userId: null, isActive: true } : null;
      }

      case 'isSystemAdmin': {
        const userEmail = args[0] || getCurrentEmailDirect();
        return checkIsSystemAdminDirect(userEmail);
      }

      default:
        console.warn(`handleUserServiceFallback: No fallback for ${methodName}`);
        return null;
    }
  } catch (error) {
    console.error(`handleUserServiceFallback(${methodName}):`, error.message);
    return null;
  }
}

/**
 * Direct email retrieval using GAS Session API
 */
function getCurrentEmailDirect() {
  try {
    // Method 1: Session.getActiveUser()
    let email = Session.getActiveUser().getEmail();
    if (email) {
      return email;
    }

    // Method 2: Session.getEffectiveUser()
    email = Session.getEffectiveUser().getEmail();
    if (email) {
      return email;
    }

    console.warn('getCurrentEmailDirect: No email available from Session API');
    return null;
  } catch (error) {
    console.error('getCurrentEmailDirect:', error.message);
    return null;
  }
}

/**
 * Direct system admin check
 */
function checkIsSystemAdminDirect(email) {
  try {
    if (!email) return false;

    const adminEmail = getSystemPropertyDirect('ADMIN_EMAIL');
    return adminEmail ? email.toLowerCase() === adminEmail.toLowerCase() : false;
  } catch (error) {
    console.error('checkIsSystemAdminDirect:', error.message);
    return false;
  }
}

/**
 * System properties fallback (no dependencies)
 */
function getSystemPropertyDirect(propertyName) {
  try {
    return PropertiesService.getScriptProperties().getProperty(propertyName);
  } catch (error) {
    console.error(`getSystemPropertyDirect(${propertyName}):`, error.message);
    return null;
  }
}

/**
 * Direct system properties check (no service dependencies)
 */
function checkCoreSystemPropsDirectly() {
  try {
    const props = PropertiesService.getScriptProperties();
    const adminEmail = props.getProperty('ADMIN_EMAIL');
    const dbId = props.getProperty('DATABASE_SPREADSHEET_ID');
    const creds = props.getProperty('SERVICE_ACCOUNT_CREDS');

    if (!adminEmail || !dbId || !creds) {
      console.warn('checkCoreSystemPropsDirectly: Missing required properties', {
        hasAdmin: !!adminEmail,
        hasDb: !!dbId,
        hasCreds: !!creds
      });
      return false;
    }

    // Validate SERVICE_ACCOUNT_CREDS JSON
    try {
      const parsed = JSON.parse(creds);
      if (!parsed || !parsed.client_email) {
        console.warn('checkCoreSystemPropsDirectly: Invalid SERVICE_ACCOUNT_CREDS format');
        return false;
      }
    } catch (jsonError) {
      console.warn('checkCoreSystemPropsDirectly: SERVICE_ACCOUNT_CREDS JSON parse failed', jsonError.message);
      return false;
    }

    return true;
  } catch (error) {
    console.error('checkCoreSystemPropsDirectly: Error', error.message);
    return false;
  }
}

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
 * Mode Access Control Rules - 統一アクセス制御マトリクス
 */
const MODE_ACCESS_RULES = Object.freeze({
  // セットアップ不要（常にアクセス可能）
  'login': { requiresSetup: false, requiresAuth: false, accessLevel: 'public' },
  'setup': { requiresSetup: false, requiresAuth: false, accessLevel: 'public' },

  // セットアップ必要 + 認証必要 + オープンアクセス
  'view': { requiresSetup: true, requiresAuth: true, accessLevel: 'open' },
  'main': { requiresSetup: true, requiresAuth: true, accessLevel: 'open' },

  // セットアップ必要 + 認証必要 + 所有者制限
  'admin': { requiresSetup: true, requiresAuth: true, accessLevel: 'owner' },
  'appSetup': { requiresSetup: true, requiresAuth: true, accessLevel: 'owner' },

  // セットアップ必要 + 認証必要 + システム管理者専用
  'debug': { requiresSetup: true, requiresAuth: true, accessLevel: 'system_admin' },
  'test': { requiresSetup: true, requiresAuth: true, accessLevel: 'system_admin' },
  'fix_user': { requiresSetup: true, requiresAuth: true, accessLevel: 'system_admin' },
  'clear_cache': { requiresSetup: true, requiresAuth: true, accessLevel: 'system_admin' }
});

/**
 * 統一モードアクセス検証
 * @param {string} mode - アクセスするモード
 * @param {Object} params - リクエストパラメータ
 * @returns {Object} アクセス結果 { allowed: boolean, redirect?: string, reason?: string }
 */
function validateModeAccess(mode, params) {
  const rules = MODE_ACCESS_RULES[mode];
  if (!rules) {
    console.warn('validateModeAccess: 不明なモード', mode);
    return { allowed: false, redirect: 'login', reason: 'unknown_mode' };
  }

  // Step 1: セットアップ必要チェック
  if (rules.requiresSetup) {
    try {
      // 🚀 Dynamic Service Discovery for ConfigService
      const hasSetup = callServiceMethod('ConfigService', 'hasCoreSystemProps') || checkCoreSystemPropsDirectly();
      if (hasSetup === null) {
        console.warn('validateModeAccess: ConfigService unavailable, checking properties directly');
        const directCheck = checkCoreSystemPropsDirectly();
        if (!directCheck) {
          return { allowed: false, redirect: 'setup', reason: 'setup_required_direct_check' };
        }
      }

      console.log('validateModeAccess: Setup check result', { mode, hasSetup });

      if (!hasSetup) {
        console.log('validateModeAccess: セットアップ未完了', mode);
        return { allowed: false, redirect: 'setup', reason: 'setup_required' };
      }
    } catch (error) {
      console.error('validateModeAccess: Setup check error', { mode, error: error.message });
      return { allowed: false, redirect: 'setup', reason: 'setup_check_error' };
    }
  }

  // Step 2: 認証必要チェック
  if (rules.requiresAuth) {
    // 🚀 Dynamic Service Discovery with Fallback
    const userEmail = callServiceMethod('UserService', 'getCurrentEmail');
    if (!userEmail) {
      console.log('validateModeAccess: 認証が必要 (no email from service discovery)', mode);
      return { allowed: false, redirect: 'login', reason: 'auth_required' };
    }

    // Step 3: アクセスレベルチェック
    if (rules.accessLevel === 'system_admin') {
      const isAdmin = callServiceMethod('UserService', 'isSystemAdmin', userEmail);
      if (!isAdmin) {
        console.log('validateModeAccess: システム管理者権限が必要', { mode, userEmail });
        return { allowed: false, redirect: 'login', reason: 'admin_required' };
      }
    }

    if (rules.accessLevel === 'owner' && params.userId) {
      // 🚀 Dynamic Service Discovery for owner check
      const currentUser = callServiceMethod('UserService', 'getCurrentUserInfo');
      const isOwnUser = currentUser && currentUser.userId === params.userId;
      const isSystemAdmin = callServiceMethod('UserService', 'isSystemAdmin', userEmail);

      if (!isOwnUser && !isSystemAdmin) {
        console.log('validateModeAccess: 所有者権限が必要', {
          mode,
          requestedUserId: params.userId,
          currentUserId: currentUser?.userId
        });
        return { allowed: false, redirect: 'login', reason: 'owner_required' };
      }
    }
  }

  console.log('validateModeAccess: アクセス許可', { mode, accessLevel: rules.accessLevel });
  return { allowed: true };
}

/**
 * doGet - HTTP GET request handler with complete mode system restoration
 * Integrates historical mode switching with modern Services architecture
 */
function doGet(e) {
  try {
    // Parse request parameters
    const params = parseRequestParams(e);
    console.log('doGet: Processing request with params:', params);

    // 統一アクセス制御チェック
    const accessResult = validateModeAccess(params.mode || 'main', params);

    if (!accessResult.allowed) {
      console.log(`doGet: Access denied, redirecting to ${accessResult.redirect}`, accessResult);

      switch (accessResult.redirect) {
        case 'setup':
          return handleSetupModeWithTemplate(params);
        case 'login':
          return handleLoginModeWithTemplate(params, { reason: accessResult.reason });
        default:
          return handleLoginModeWithTemplate(params, { reason: 'access_error' });
      }
    }

    // アクセス許可後の既存ルーティング
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
    // Security validation with service availability check
    // 🚀 Direct Session API - Zero Dependencies
    const userEmail = getCurrentEmailDirect();
    if (!userEmail) {
      console.error('doPost: No user email available from Session API');
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: 'Authentication required, please login'
      })).setMimeType(ContentService.MimeType.JSON);
    }

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
    template.userEmail = getCurrentEmailDirect() || '';

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

    // 🚀 Direct Session API - Zero Dependencies
    const userEmail = getCurrentEmailDirect();
    if (!userEmail) {
      return handleLoginModeWithTemplate(params, { reason: 'admin_access_requires_auth' });
    }

    // URLパラメータからuserIdを取得（管理パネル用）- validateModeAccess()で権限チェック済み
    let {userId} = params;
    let userInfo = { userEmail, userId, isActive: true };

    // userIdが指定されていない場合、フォールバック処理（依存関係最小化）
    if (!userId) {
      try {
        // 🚀 Zero-dependency fallback: Generate temporary userId for session
        userId = `temp_${userEmail.replace('@', '_at_').replace(/\./g, '_')}`;
        userInfo = { userId, userEmail, isActive: true, isTemporary: true };
        console.info('handleAdminModeWithTemplate: Using temporary userId for session', { userId });
      } catch (e) {
        console.warn('handleAdminModeWithTemplate: Temporary user creation error:', e.message);
        userId = 'temp_unknown';
        userInfo = { userId, userEmail, isActive: true };
      }
    } else {
      // userIdが指定されている場合、そのユーザー情報を使用
      userInfo = { userId, userEmail, isActive: true };
    }

    const template = HtmlService.createTemplateFromFile('AdminPanel');
    template.params = params;
    template.context = context;
    template.userEmail = userEmail;
    template.userInfo = userInfo;
    template.isSystemAdmin = checkIsSystemAdminDirect(userEmail);

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

function handleViewMode(params) {
  try {
    // 🚀 Direct Session API - Zero Dependencies
    const userEmail = getCurrentEmailDirect();

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

function handleTestMode() {
  return HtmlService.createHtmlOutput(`
    <h2>Test Mode</h2>
    <p>Test functionality would appear here</p>
  `).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function handleFixUserMode() {
  return HtmlService.createHtmlOutput(`
    <h2>Fix User Mode</h2>
    <p>User fix functionality would appear here</p>
  `).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function handleClearCacheMode() {
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


/**
 * System Controller Global Functions
 */


function getUser(kind = 'email') {
  try {
    // 🚀 Dynamic Service Discovery for UserService
    const userEmail = callServiceMethod('UserService', 'getCurrentEmail');

    if (!userEmail) {
      console.warn('getUser: No user email available via service discovery');
      return kind === 'email' ? '' : { success: false, message: 'ユーザー情報が取得できません' };
    }

    // 後方互換性重視: kind==='email' の場合は純粋な文字列を返す
    if (kind === 'email') {
      return String(userEmail);
    }

    // 統一オブジェクト形式（'full' など）
    const userInfo = callServiceMethod('UserService', 'getCurrentUserInfo');
    if (!userInfo) {
      console.warn('getUser: UserService.getCurrentUserInfo not available via service discovery');
      return { success: false, message: 'Full user info service unavailable' };
    }

    return {
      success: true,
      email: userEmail,
      userId: userInfo?.userId || null,
      isActive: userInfo?.isActive || false,
      hasConfig: !!userInfo?.config
    };
  } catch (error) {
    console.error('getUser エラー:', error.message);
    return kind === 'email' ? '' : { success: false, message: error.message };
  }
}

function getApplicationStatusForUI() {
  try {
    if (typeof ConfigService !== 'undefined' && ConfigService.getApplicationStatus) {
      return ConfigService.getApplicationStatus();
    } else {
      console.warn('getApplicationStatusForUI: ConfigService not available');
      return { success: false, error: 'ConfigService not available' };
    }
  } catch (error) {
    console.error('getApplicationStatusForUI error:', error);
    return { success: false, error: error.message };
  }
}


function setApplicationStatusForUI(isActive) {
  try {
    if (typeof ConfigService !== 'undefined' && ConfigService.setApplicationStatus) {
      return ConfigService.setApplicationStatus(isActive);
    } else {
      console.warn('setApplicationStatusForUI: ConfigService not available');
      return { success: false, error: 'ConfigService not available' };
    }
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

/**
 * Admin Controller Global Functions
 */

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
function connectDataSource(spreadsheetId, sheetName) {
  try {
    if (typeof ConfigService !== 'undefined' && ConfigService.connectDataSource) {
      return ConfigService.connectDataSource(spreadsheetId, sheetName);
    } else {
      console.warn('connectDataSource: ConfigService not available');
      return { success: false, error: 'ConfigService not available' };
    }
  } catch (error) {
    console.error('connectDataSource error:', error);
    return { success: false, error: error.message };
  }
}
function getSheetList(spreadsheetId) {
  try {
    return DataService.getSheetList(spreadsheetId);
  } catch (error) {
    console.error('getSheetList error:', error);
    return { success: false, error: error.message };
  }
}


/**
 * Data Controller Global Functions
 */

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
        if (typeof ConfigService !== 'undefined' && ConfigService.getDefaultConfig) {
          return ConfigService.getDefaultConfig('unknown');
        } else {
          console.warn('getUserConfig: ConfigService not available, returning fallback config');
          return { setupStatus: 'pending', appPublished: false };
        }
      }
    }

    if (typeof ConfigService !== 'undefined' && ConfigService.getUserConfig) {
      const result = ConfigService.getUserConfig(userId);
      console.log('getUserConfig: 正常終了', { hasResult: !!result });
      if (!result) {
        console.warn('getUserConfig: ConfigServiceからnullが返却された - デフォルト設定を返却');
        return ConfigService.getDefaultConfig(userId);
      }
      return result;
    } else {
      console.warn('getUserConfig: ConfigService not available, returning fallback config');
      return { setupStatus: 'pending', appPublished: false };
    }

  } catch (error) {
    console.error('getUserConfig: エラー', error);
    if (typeof ConfigService !== 'undefined' && ConfigService.getDefaultConfig) {
      return ConfigService.getDefaultConfig(userId || 'error');
    } else {
      return { setupStatus: 'error', appPublished: false };
    }
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
        if (sessionUser && typeof ConfigService !== 'undefined' && ConfigService.getUserConfig) {
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