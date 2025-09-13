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
    const errorResponse = ErrorHandler.handle(error, 'doGet');
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
    const errorResponse = ErrorHandler.handle(error, 'doPost');
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

    // URLパラメータからuserIdを取得（管理パネル用）
    let userInfo = UserService.getCurrentUserInfo();
    let userId = params.userId || (userInfo && userInfo.userId);

    // userIdが指定されていない場合、現在のユーザー情報から取得または作成
    if (!userId) {
      try {
        // DB から直接ユーザー検索を試行
        const dbUser = DB.findUserByEmail(userEmail);
        if (dbUser && dbUser.userId) {
          userId = dbUser.userId;
          userInfo = { userId, userEmail };
        } else {
          // ユーザーが存在しない場合は作成
          const created = UserService.createUser(userEmail);
          if (created && created.userId) {
            userId = created.userId;
            userInfo = created;
          } else {
            console.error('main.js handleAdminModeWithTemplate: ユーザー作成に失敗 - userIdが見つかりません', { created });
            throw new Error('ユーザー作成に失敗しました');
          }
        }
      } catch (e) {
        console.warn('handleAdminModeWithTemplate: ユーザー作成/取得でエラー:', e.message);
        userInfo = { userId: '', userEmail };
      }
    } else {
      // userIdが指定されている場合、そのユーザー情報を使用
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
  return renderMainPage({ mode: 'view', ...params });
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

/**
 * タイムスタンプフォーマット
 * @param {Date|string} date - 日付
 * @returns {string} フォーマット済み日付
 */
function formatTimestamp(date) {
  try {
    return ResponseFormatter.formatTimestamp(date);
  } catch (error) {
    console.error('formatTimestamp error:', error);
    return new Date(date).toISOString();
  }
}

/**
 * 成功レスポンス作成
 * @param {*} data - データ
 * @param {Object} metadata - メタデータ
 * @returns {Object} 成功レスポンス
 */
function createSuccessResponse(data, metadata) {
  try {
    return ResponseFormatter.createSuccessResponse(data, metadata);
  } catch (error) {
    console.error('createSuccessResponse error:', error);
    return { success: true, data };
  }
}

/**
 * エラーレスポンス作成
 * @param {string} message - エラーメッセージ
 * @param {Object} details - 詳細情報
 * @returns {Object} エラーレスポンス
 */
function createErrorResponse(message, details) {
  try {
    return ResponseFormatter.createErrorResponse(message, details);
  } catch (error) {
    console.error('createErrorResponse error:', error);
    return { success: false, error: message };
  }
}

// ===========================================
// 🌐 Controller Global Function Exports
// Required for google.script.run.Controller.method() calls
// ===========================================

/**
 * Frontend Controller Global Functions
 * Required for HTML google.script.run calls
 */
function getUser(kind) {
  return FrontendController.getUser(kind);
}

function processLoginAction() {
  return FrontendController.processLoginAction();
}

/**
 * System Controller Global Functions
 */
function getSystemDomainInfo() {
  return SystemController.getSystemDomainInfo();
}

function forceUrlSystemReset() {
  return SystemController.forceUrlSystemReset();
}

function setupApplication(serviceAccountJson, databaseId, adminEmail, googleClientId) {
  return SystemController.setupApplication(serviceAccountJson, databaseId, adminEmail, googleClientId);
}

function testSetup() {
  return SystemController.testSetup();
}

function getApplicationStatusForUI() {
  return SystemController.getApplicationStatusForUI();
}

function getAllUsersForAdminForUI(userId) {
  return SystemController.getAllUsersForAdminForUI(userId);
}

function deleteUserAccountByAdminForUI(userId, reason) {
  return SystemController.deleteUserAccountByAdminForUI(userId, reason);
}

function getWebAppUrl() {
  return SystemController.getWebAppUrl();
}

function reportClientError(errorInfo) {
  return SystemController.reportClientError(errorInfo);
}

function setApplicationStatusForUI(isActive) {
  return SystemController.setApplicationStatusForUI(isActive);
}

function getDeletionLogsForUI(userId) {
  return SystemController.getDeletionLogsForUI(userId);
}

function testSystemDiagnosis() {
  return SystemController.testSystemDiagnosis();
}

function performAutoRepair() {
  return SystemController.performAutoRepair();
}

function performSystemMonitoring() {
  return SystemController.performSystemMonitoring();
}

function performDataIntegrityCheck() {
  return SystemController.performDataIntegrityCheck();
}

/**
 * Admin Controller Global Functions
 */
function getConfig() {
  return AdminController.getConfig();
}

function getSpreadsheetList() {
  return AdminController.getSpreadsheetList();
}

function analyzeColumns(spreadsheetId, sheetName) {
  return AdminController.analyzeColumns(spreadsheetId, sheetName);
}

function publishApplication(publishConfig) {
  return AdminController.publishApplication(publishConfig);
}

function saveDraftConfiguration(draftConfig) {
  return AdminController.saveDraftConfiguration(draftConfig);
}

function connectDataSource(spreadsheetId, sheetName) {
  return DataController.connectDataSource(spreadsheetId, sheetName);
}

/**
 * Data Controller Global Functions
 */
function addReaction(userId, rowId, reactionType) {
  return DataController.addReaction(userId, rowId, reactionType);
}

function toggleHighlight(userId, rowId) {
  return DataController.toggleHighlight(userId, rowId);
}

function refreshBoardData(userId) {
  return DataController.refreshBoardData(userId);
}

function getUserConfig(userId) {
  return ConfigService.getUserConfig(userId);
}

function getPublishedSheetData(userId, options) {
  return DataService.getPublishedSheetData(userId, options);
}