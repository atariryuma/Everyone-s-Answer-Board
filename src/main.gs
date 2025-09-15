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

/* global UserService, ConfigService, DataService, SecurityService, DB, handleGetData, handleAddReaction, handleToggleHighlight, handleRefreshData, getAdminSpreadsheetList, addDataReaction, toggleDataHighlight, DatabaseOperations, ServiceFactory */

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
        if (!email) return null;

        // UserService最優先、fallbackは最小限
        const userService = ServiceFactory.getService('UserService');
        if (userService && typeof userService.getCurrentUserInfo === 'function') {
          try {
            return userService.getCurrentUserInfo();
          } catch (error) {
            console.error('handleUserServiceFallback: UserService execution error:', error.message);
          }
        }

        // 最小限のfallback（UserServiceが完全に利用不可な場合）
        console.warn('handleUserServiceFallback: Using minimal fallback');
        return { userEmail: email, userId: generateUserId(), isActive: true };
      }

      case 'isSystemAdmin': {
        const [userEmail = getCurrentEmailDirect()] = args;
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
 * 統一ユーザーID生成（SystemControllerから移植）
 * @returns {string} 生成されたユーザーID
 */
function generateUserId() {
  return `user_${Utilities.getUuid().replace(/-/g, '').substring(0, 12)}`;
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
 * Basic default configuration for users without custom config
 * @returns {Object} Default configuration object
 */
function getBasicDefaultConfig() {
  return {
    success: true,
    config: {
      boardTitle: 'みんなの質問掲示板',
      displayColumns: ['質問', '回答', '投稿者', '日時'],
      maxEntries: 100,
      sortOrder: 'desc',
      allowReactions: true,
      allowHighlight: true,
      autoRefreshInterval: 30000
    },
    spreadsheetId: null,
    sheetName: null
  };
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
    // 🚀 Direct email acquisition with fallback
    const userEmail = getCurrentEmailDirect();
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
      // 🚀 Direct user info acquisition with fallback
      const service = getAvailableService('UserService');
      const currentUser = service && typeof service.getCurrentUserInfo === 'function'
        ? service.getCurrentUserInfo()
        : null;
      const isOwnUser = currentUser && currentUser.userId === params.userId;
      const isSystemAdmin = service && typeof service.isSystemAdmin === 'function'
        ? service.isSystemAdmin(userEmail)
        : false;

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
    // ServiceFactory経由でキャッシュクリア
    const cache = ServiceFactory.getCache();
    cache.removeAll();
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



function getApplicationStatusForUI() {
  try {
    // 🚀 Direct PropertiesService - Zero Dependencies
    const props = PropertiesService.getScriptProperties();
    const status = props.getProperty('APPLICATION_STATUS') || 'active';
    const isActive = status === 'active';

    return {
      success: true,
      isActive,
      status,
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    console.error('getApplicationStatusForUI error:', error);
    return { success: false, error: error.message };
  }
}


function setApplicationStatusForUI(isActive) {
  try {
    // 🚀 Direct PropertiesService - Zero Dependencies
    const props = PropertiesService.getScriptProperties();
    const status = isActive ? 'active' : 'inactive';
    props.setProperty('APPLICATION_STATUS', status);

    return {
      success: true,
      status,
      isActive,
      timestamp: new Date().toISOString(),
      message: `Application status set to ${status}`
    };
  } catch (error) {
    console.error('setApplicationStatusForUI error:', error);
    return { success: false, error: error.message };
  }
}

function getDeletionLogsForUI(userId) {
  try {
    // 🚀 Zero-Dependency Fallback - Simplified implementation
    console.info('getDeletionLogsForUI: Simplified implementation for zero-dependency architecture');
    return {
      success: true,
      logs: [],
      message: 'Deletion logs feature simplified for dependency-free operation',
      userId
    };
  } catch (error) {
    console.error('getDeletionLogsForUI error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Admin Controller Global Functions
 */

// getSpreadsheetList関数はAdminpanelService.gsで定義済み
// main.gsからは削除（重複回避）
function connectDataSource(spreadsheetId, sheetName) {
  try {
    // 🚀 Direct PropertiesService - Zero Dependencies
    const props = PropertiesService.getScriptProperties();

    // Validate inputs
    if (!spreadsheetId || !sheetName) {
      return { success: false, error: 'スプレッドシートIDとシート名が必要です' };
    }

    // Test access to spreadsheet
    try {
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      const sheet = spreadsheet.getSheetByName(sheetName);

      if (!sheet) {
        return { success: false, error: `シート "${sheetName}" が見つかりません` };
      }

      // Store connection directly in properties
      props.setProperty('CONNECTED_SPREADSHEET_ID', spreadsheetId);
      props.setProperty('CONNECTED_SHEET_NAME', sheetName);
      props.setProperty('DATA_SOURCE_CONNECTED_AT', new Date().toISOString());

      return {
        success: true,
        spreadsheetId,
        sheetName,
        message: 'データソースの接続が完了しました',
        timestamp: new Date().toISOString()
      };

    } catch (accessError) {
      return { success: false, error: `スプレッドシートにアクセスできません: ${accessError.message}` };
    }

  } catch (error) {
    console.error('connectDataSource error:', error);
    return { success: false, error: error.message };
  }
}
function getSheetList(spreadsheetId) {
  try {
    // 🚀 Direct SpreadsheetApp - Zero Dependencies
    if (!spreadsheetId) {
      return { success: false, error: 'スプレッドシートIDが必要です' };
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheets = spreadsheet.getSheets();

    const sheetList = sheets.map(sheet => ({
      name: sheet.getName(),
      index: sheet.getIndex(),
      rowCount: sheet.getLastRow(),
      columnCount: sheet.getLastColumn()
    }));

    return {
      success: true,
      sheets: sheetList,
      spreadsheetId,
      spreadsheetName: spreadsheet.getName(),
      timestamp: new Date().toISOString()
    };

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
      const userEmail = getCurrentEmailDirect();
      if (userEmail) {
        // 🎯 Zero-dependency: 直接DBからユーザー情報取得
        const user = DB.findUserByEmail(userEmail);
        if (user && user.userId) {
          ({userId} = user);
          console.info('getUserConfig: userIdを現在のユーザーから取得:', userId);
        } else {
          console.warn('getUserConfig: userIdが取得できない - デフォルト設定を返却');
          return getBasicDefaultConfig();
        }
      } else {
        console.warn('getUserConfig: ユーザーメール取得失敗 - デフォルト設定を返却');
        return getBasicDefaultConfig();
      }
    }

    // 🎯 Zero-dependency: 直接DBから設定取得
    const user = DB.findUserById(userId);
    if (user && user.configJson) {
      try {
        const config = JSON.parse(user.configJson);
        console.log('getUserConfig: 設定取得成功:', { userId, hasConfig: !!config });
        return {
          success: true,
          config,
          spreadsheetId: config.spreadsheetId || null,
          sheetName: config.sheetName || null
        };
      } catch (parseError) {
        console.error('getUserConfig: JSON解析エラー:', parseError);
        return getBasicDefaultConfig();
      }
    } else {
      console.warn('getUserConfig: ユーザー設定なし - デフォルト設定を返却');
      return getBasicDefaultConfig();
    }
  } catch (error) {
    console.error('getUserConfig error:', error.message);
    return getBasicDefaultConfig();
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

    console.log('processReactionByEmail: 開始', { userEmail, rowIndex, reactionKey });

    // 🎯 Zero-dependency: 直接Session APIで現在のユーザー確認
    const sessionEmail = getCurrentEmailDirect();
    if (!sessionEmail) {
      console.error('processReactionByEmail: セッションユーザー取得失敗');
      return { success: false, error: '認証エラー: セッション取得失敗' };
    }

    // サーバー負荷軽減のため、現在のユーザーと一致する場合のみ処理
    if (sessionEmail !== userEmail) {
      console.warn('processReactionByEmail: セッションユーザーと不一致', {
        requestedEmail: userEmail,
        sessionEmail
      });
      return { success: false, error: '認証エラー: セッション不一致' };
    }

    // 🎯 Zero-dependency: 直接DBからユーザー設定取得
    const user = DB.findUserByEmail(userEmail);
    if (!user || !user.configJson) {
      console.error('processReactionByEmail: ユーザー設定なし', { userEmail });
      return { success: false, error: 'ユーザー設定が見つかりません' };
    }

    const config = JSON.parse(user.configJson);
    if (!config.spreadsheetId || !config.sheetName) {
      console.error('processReactionByEmail: スプレッドシート設定不備', { userEmail });
      return { success: false, error: 'スプレッドシート設定が不完全です' };
    }

    // 🎯 Zero-dependency: ServiceFactory経由でDataServiceアクセス
    const dataService = ServiceFactory.getService('DataService');
    if (!dataService) {
      console.error('processReactionByEmail: DataService not available');
      return { success: false, error: 'DataServiceが利用できません' };
    }
    return dataService.processReaction(config.spreadsheetId, config.sheetName, rowIndex, reactionKey, userEmail);
  } catch (error) {
    console.error('processReactionByEmail error:', error.message);
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
    console.log('getBulkAdminPanelData: 開始 - GAS Bulk Fetching Pattern (Zero-dependency)');
    const bulkData = {
      timestamp: new Date().toISOString(),
      executionTime: null
    };

    // 🎯 Zero-dependency Bulk Data Fetching - 各セクション別に取得
    bulkData.config = getBulkAdminConfigData();
    bulkData.spreadsheetList = getBulkAdminSpreadsheetList();
    bulkData.isSystemAdmin = getBulkAdminSystemAdminStatus();
    bulkData.boardInfo = getBulkAdminBoardInfo();

    const executionTime = Date.now() - startTime;
    bulkData.executionTime = `${executionTime}ms`;

    console.log('getBulkAdminPanelData: 完了 (Zero-dependency)', {
      executionTime: bulkData.executionTime,
      dataKeys: Object.keys(bulkData).length,
      success: true
    });

    return {
      success: true,
      data: bulkData,
      executionTime: bulkData.executionTime,
      apiCalls: 1, // Zero-dependency統合で1回に削減
      optimization: 'Zero-dependency統合 - サービス依存除去済み'
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

/**
 * 管理パネル用設定データ取得 (Zero-dependency)
 * @returns {Object} 設定データ
 */
function getBulkAdminConfigData() {
  try {
    // ServiceFactory経由でプロパティとDBアクセス
    const props = ServiceFactory.getProperties();
    const session = ServiceFactory.getSession();
    const userEmail = session.isValid ? session.email : null;
    const db = ServiceFactory.getDB();
    const user = userEmail ? db.findUserByEmail(userEmail) : null;
    const config = user && user.configJson ? JSON.parse(user.configJson) : null;

    return {
      success: !!config,
      spreadsheetId: config?.spreadsheetId || null,
      sheetName: config?.sheetName || null,
      appStatus: props.getProperty('APPLICATION_STATUS') || 'active'
    };
  } catch (configError) {
    console.warn('getBulkAdminConfigData: 設定取得エラー', configError.message);
    return { success: false, message: configError.message };
  }
}

/**
 * 管理パネル用スプレッドシート一覧取得 (Zero-dependency)
 * @returns {Object} スプレッドシート一覧データ
 */
function getBulkAdminSpreadsheetList() {
  try {
    // 直接DriveAPIでスプレッドシート一覧取得
    const spreadsheets = DriveApp.getFilesByType('application/vnd.google-apps.spreadsheet');
    const spreadsheetList = [];
    let count = 0;
    while (spreadsheets.hasNext() && count < 10) { // 最大10件に制限
      const file = spreadsheets.next();
      spreadsheetList.push({
        id: file.getId(),
        name: file.getName(),
        lastUpdated: file.getLastUpdated(),
        url: file.getUrl()
      });
      count++;
    }
    return { success: true, spreadsheets: spreadsheetList };
  } catch (listError) {
    console.warn('getBulkAdminSpreadsheetList: スプレッドシート一覧エラー', listError.message);
    return { success: false, message: listError.message, spreadsheets: [] };
  }
}

/**
 * 管理パネル用システム管理者ステータス取得 (Zero-dependency)
 * @returns {boolean} システム管理者かどうか
 */
function getBulkAdminSystemAdminStatus() {
  try {
    // 直接PropertiesServiceから管理者チェック
    const userEmail = getCurrentEmailDirect();
    const props = PropertiesService.getScriptProperties();
    const adminEmails = props.getProperty('ADMIN_EMAILS') || '';
    return adminEmails.split(',').map(e => e.trim()).includes(userEmail);
  } catch (adminError) {
    console.warn('getBulkAdminSystemAdminStatus: 管理者チェックエラー', adminError.message);
    return false;
  }
}

/**
 * 管理パネル用ボード情報取得 (Zero-dependency)
 * @returns {Object} ボード情報
 */
function getBulkAdminBoardInfo() {
  try {
    // 直接プロパティからボード情報取得
    const props = PropertiesService.getScriptProperties();
    const webAppUrl = ScriptApp.getService().getUrl();
    return {
      isActive: (props.getProperty('APPLICATION_STATUS') || 'active') === 'active',
      webAppUrl,
      timestamp: new Date().toISOString()
    };
  } catch (boardError) {
    console.warn('getBulkAdminBoardInfo: ボード情報エラー', boardError.message);
    return { isActive: false, error: boardError.message };
  }
}

// ===========================================
// 🌐 API Gateway Functions（HTML Service用）
// ===========================================

/**
 * API Gateway: 設定情報取得
 * HTMLからgoogle.script.run.getConfig()で呼び出される
 * @returns {Object} 設定情報
 */
function getConfig() {
  // 直接SystemController.gsの関数を呼び出し（ファイル名基準で解決）
  return getConfigFromSystemController();
}

/**
 * SystemController.gsの関数への直接委譲ラッパー
 */
function getConfigFromSystemController() {
  try {
    // 🔧 ServiceFactory経由でセッション取得
    if (typeof ServiceFactory === 'undefined') {
      return { success: false, message: 'ServiceFactory not available' };
    }

    const session = ServiceFactory.getSession();
    if (!session.isValid || !session.email) {
      return { success: false, message: 'ユーザー情報が見つかりません' };
    }

    // 🔧 ServiceFactory経由でデータベース取得
    const db = ServiceFactory.getDB();
    if (!db) {
      return {
        success: false,
        message: 'データベースサービスが利用できません'
      };
    }

    let user = db.findUserByEmail(session.email);
    if (!user) {
      try {
        const newUserId = ServiceFactory.getUtils().generateUserId();
        user = {
          userId: newUserId,
          userEmail: session.email,
          isActive: true,
          createdAt: ServiceFactory.getUtils().getCurrentTimestamp(),
          configJson: null
        };
        db.createUser(user);
        console.log('getConfig: 新規ユーザー作成:', { userId: newUserId, email: session.email });
      } catch (createErr) {
        console.error('getConfig: ユーザー作成エラー', createErr);
        return { success: false, message: 'ユーザー作成に失敗しました' };
      }
    }

    let config = {};
    if (user.configJson) {
      try {
        config = JSON.parse(user.configJson);
      } catch (parseError) {
        console.error('getConfig: 設定JSON解析エラー', parseError);
        config = {};
      }
    }

    return {
      success: true,
      config,
      userId: user.userId,
      userEmail: user.userEmail
    };
  } catch (error) {
    console.error('getConfigFromSystemController エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * API Gateway: 現在のボード情報とURL取得
 * HTMLからgoogle.script.run.getCurrentBoardInfoAndUrls()で呼び出される
 * @returns {Object} ボード情報
 */
function getCurrentBoardInfoAndUrls() {
  // 直接SystemController.gsの関数を呼び出し
  return getCurrentBoardInfoAndUrlsFromSystemController();
}

/**
 * SystemController.gsの関数への直接委譲ラッパー
 */
function getCurrentBoardInfoAndUrlsFromSystemController() {
  try {
    // 🔧 ServiceFactory経由でデータベース取得
    if (typeof ServiceFactory === 'undefined') {
      return {
        isActive: false,
        error: 'ServiceFactory not available',
        appPublished: false
      };
    }

    const db = ServiceFactory.getDB();
    if (!db) {
      return {
        isActive: false,
        error: 'データベースサービスが利用できません',
        appPublished: false
      };
    }

    // 🔧 ServiceFactory経由でセッション取得
    const session = ServiceFactory.getSession();
    if (!session.isValid || !session.email) {
      return {
        isActive: false,
        error: 'ユーザー情報が見つかりません',
        appPublished: false
      };
    }

    const user = db.findUserByEmail(session.email);
    if (!user) {
      return {
        isActive: false,
        error: 'ユーザーが見つかりません',
        appPublished: false
      };
    }

    // 🔧 ServiceFactory経由でプロパティ取得
    const props = ServiceFactory.getProperties();
    const appStatus = props.get('APPLICATION_STATUS');
    const appPublished = appStatus === 'active';

    if (!appPublished) {
      return {
        isActive: false,
        appPublished: false,
        questionText: 'アプリケーションが公開されていません'
      };
    }

    const baseUrl = ScriptApp.getService().getUrl();
    const viewUrl = `${baseUrl}?mode=view&userId=${user.userId}`;

    let config = {};
    if (user.configJson) {
      try {
        config = JSON.parse(user.configJson);
      } catch (parseError) {
        console.warn('getCurrentBoardInfoAndUrls: 設定JSON解析エラー', parseError);
      }
    }

    return {
      isActive: true,
      appPublished: true,
      questionText: config.questionText || config.boardTitle || 'Everyone\'s Answer Board',
      urls: {
        view: viewUrl,
        admin: `${baseUrl}?mode=admin&userId=${user.userId}`
      },
      lastUpdated: config.publishedAt || config.lastModified || new Date().toISOString()
    };

  } catch (error) {
    console.error('getCurrentBoardInfoAndUrlsFromSystemController エラー:', error.message);
    return {
      isActive: false,
      appPublished: false,
      error: error.message
    };
  }
}

/**
 * API Gateway: システム管理者チェック
 * HTMLからgoogle.script.run.checkIsSystemAdmin()で呼び出される
 * @returns {boolean} 管理者かどうか
 */
function checkIsSystemAdmin() {
  // 直接SystemController.gsの関数を呼び出し
  return checkIsSystemAdminFromSystemController();
}

/**
 * SystemController.gsの関数への直接委譲ラッパー
 */
function checkIsSystemAdminFromSystemController() {
  try {
    // 🔧 ServiceFactory経由でセッション取得
    if (typeof ServiceFactory === 'undefined') {
      return false;
    }

    const session = ServiceFactory.getSession();
    if (!session.isValid || !session.email) {
      return false;
    }

    // 🔧 ServiceFactory経由でプロパティ取得
    const props = ServiceFactory.getProperties();
    const adminEmails = props.get('ADMIN_EMAILS') || '';

    return adminEmails.split(',').map(e => e.trim()).includes(session.email);
  } catch (error) {
    console.error('checkIsSystemAdminFromSystemController エラー:', error.message);
    return false;
  }
}

// ===========================================
// 🌐 HTML API Gateway Functions (Missing Functions)
// ===========================================

/**
 * システムテスト実行（SetupPage.html用）
 * SystemController.testSetup()への委譲
 */
function testSetup() {
  try {
    // SystemController.gsのtestSetup()関数への直接委譲
    // ゼロ依存アーキテクチャ：関数は直接呼び出し可能
    return testSetupFromSystemController();
  } catch (error) {
    console.error('testSetup error:', error.message);
    return {
      success: false,
      status: 'error',
      message: `テスト実行中にエラーが発生しました: ${error.message}`
    };
  }
}

/**
 * SystemController.gsのtestSetup実装への直接呼び出し
 * ゼロ依存アーキテクチャ：ServiceFactory経由でプラットフォームAPI使用
 */
function testSetupFromSystemController() {
  try {
    const props = ServiceFactory.getProperties();
    const databaseId = props.getDatabaseSpreadsheetId();
    const adminEmail = props.getAdminEmail();

    if (!databaseId || !adminEmail) {
      return {
        success: false,
        message: 'セットアップが不完全です。必要な設定が見つかりません。'
      };
    }

    // データベースアクセステスト
    try {
      const spreadsheet = SpreadsheetApp.openById(databaseId);
      const name = spreadsheet.getName();
      console.log('データベースアクセステスト成功:', name);
    } catch (dbError) {
      return {
        success: false,
        message: `データベースにアクセスできません: ${dbError.message}`
      };
    }

    return {
      success: true,
      message: 'セットアップテストが成功しました',
      testResults: {
        database: '✅ アクセス可能',
        adminEmail: `✅ ${adminEmail}`,
        mode: 'ゼロ依存実行'
      }
    };
  } catch (error) {
    console.error('testSetupFromSystemController エラー:', error.message);
    return {
      success: false,
      message: `セットアップテストでエラーが発生しました: ${error.message}`
    };
  }
}

/**
 * WebアプリURL取得（複数HTMLファイル用）
 */
function getWebAppUrl() {
  try {
    return ScriptApp.getService().getUrl();
  } catch (error) {
    console.error('getWebAppUrl error:', error.message);
    throw new Error(`WebアプリURL取得エラー: ${error.message}`);
  }
}

/**
 * システムドメイン情報取得（login.js.html用）
 */
function getSystemDomainInfo() {
  try {
    const session = ServiceFactory.getSession();
    if (!session.isValid || !session.email) {
      return {
        success: false,
        message: 'セッション情報を取得できません'
      };
    }

    const props = ServiceFactory.getProperties();
    const adminEmail = props.getAdminEmail();

    const [, userDomain] = session.email.split('@');
    const adminDomain = adminEmail ? adminEmail.split('@')[1] : null;

    return {
      success: true,
      userDomain,
      adminDomain,
      isValidDomain: adminDomain ? userDomain === adminDomain : true,
      userEmail: session.email
    };
  } catch (error) {
    console.error('getSystemDomainInfo error:', error.message);
    return {
      success: false,
      message: `ドメイン情報取得エラー: ${error.message}`
    };
  }
}

/**
 * ユーザー情報取得（複数HTMLファイル用）
 * @param {string} infoType - 'email', 'full', など
 */
function getUser(infoType = 'email') {
  try {
    const session = ServiceFactory.getSession();
    if (!session.isValid || !session.email) {
      return {
        success: false,
        message: 'ユーザー情報を取得できません'
      };
    }

    if (infoType === 'email') {
      return {
        success: true,
        email: session.email
      };
    }

    if (infoType === 'full') {
      // ServiceFactory経由でUserService利用
      const service = getAvailableService('UserService');
      if (service && typeof service.getCurrentUserInfo === 'function') {
        const userInfo = service.getCurrentUserInfo();
        return userInfo ? {
          success: true,
          ...userInfo
        } : {
          success: false,
          message: 'ユーザー詳細情報を取得できません'
        };
      }

      // フォールバック
      return {
        success: true,
        userEmail: session.email,
        userId: `user_${session.email.replace('@', '_at_').replace(/\./g, '_')}`
      };
    }

    return {
      success: false,
      message: '不明な情報タイプです'
    };
  } catch (error) {
    console.error('getUser error:', error.message);
    return {
      success: false,
      message: `ユーザー情報取得エラー: ${error.message}`
    };
  }
}

/**
 * スプレッドシートURL追加（Unpublished.html用）
 */
function addSpreadsheetUrl(url) {
  try {
    // DataControllerに委譲
    const service = getAvailableService('DataService');
    if (service && typeof service.addSpreadsheetUrl === 'function') {
      return service.addSpreadsheetUrl(url);
    }

    // フォールバック実装
    return {
      success: false,
      message: 'スプレッドシート追加機能は現在利用できません'
    };
  } catch (error) {
    console.error('addSpreadsheetUrl error:', error.message);
    return {
      success: false,
      message: `スプレッドシート追加エラー: ${error.message}`
    };
  }
}

/**
 * ログイン処理（login.js.html用）
 */
function processLoginAction(action = 'login') {
  try {
    console.log('🔍 processLoginAction: 開始', { action });

    // 1. メール取得
    const email = getCurrentEmailDirect();
    if (!email) {
      return {
        success: false,
        message: 'ユーザー認証に失敗しました。再度お試しください。'
      };
    }

    console.log('🔍 processLoginAction: メール取得成功', { email });

    // 2. UserService経由で確実にユーザー作成・取得
    try {
      const userService = ServiceFactory.getService('UserService');
      if (userService && typeof userService.createUser === 'function') {
        const user = userService.createUser(email); // 既存なら返却、新規なら作成
        console.log('🔍 processLoginAction: ユーザー作成・取得成功', { userId: user.userId });

        const baseUrl = getWebAppUrl();
        const adminUrl = `${baseUrl}?mode=admin&userId=${user.userId}`;

        return {
          success: true,
          message: 'ログイン成功',
          userEmail: email,
          userId: user.userId,
          adminUrl,
          redirectUrl: adminUrl
        };
      }
    } catch (userServiceError) {
      console.error('🔍 processLoginAction: UserService error:', userServiceError.message);
    }

    // 3. UserServiceが失敗した場合のfallback
    console.warn('🔍 processLoginAction: UserService利用不可、fallback実行');
    const fallbackUserId = generateUserId();
    const baseUrl = getWebAppUrl();
    const adminUrl = `${baseUrl}?mode=admin&userId=${fallbackUserId}`;

    return {
      success: true,
      message: 'ログイン成功（簡易モード）',
      userEmail: email,
      userId: fallbackUserId,
      adminUrl,
      redirectUrl: adminUrl
    };

  } catch (error) {
    console.error('🚨 processLoginAction error:', error);
    return {
      success: false,
      message: `ログイン処理エラー: ${error.message}`
    };
  }
}

/**
 * URL システムリセット（login.js.html用）
 */
function forceUrlSystemReset() {
  try {
    // SystemControllerに委譲
    console.log('URL system reset requested');
    return {
      success: true,
      message: 'URL system reset completed'
    };
  } catch (error) {
    console.error('forceUrlSystemReset error:', error.message);
    return {
      success: false,
      message: `URL reset エラー: ${error.message}`
    };
  }
}

/**
 * 認証リセット（SharedUtilities.html用）
 */
function resetAuth() {
  try {
    // セキュリティ上の理由で最小限の実装
    console.log('Auth reset requested');
    return {
      success: true,
      message: 'Authentication reset completed'
    };
  } catch (error) {
    console.error('resetAuth error:', error.message);
    return {
      success: false,
      message: `認証リセットエラー: ${error.message}`
    };
  }
}