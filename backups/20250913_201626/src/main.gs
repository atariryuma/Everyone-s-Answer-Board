/**
 * main.gs - Clean Application Entry Points
 * New Services Architecture Implementation
 *
 * 🎯 Responsibilities:
 * - HTTP request routing (doGet/doPost)
 * - Service layer coordination
 * - Error handling & user feedback
 */

/* global UserService, ConfigService, DataService, SecurityService, ErrorHandler, DB, PROPS_KEYS, URL */

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
  VERSION: '2.1.0',
  
  MODES: Object.freeze({
    MAIN: 'main',
    ADMIN: 'admin', 
    LOGIN: 'login',
    SETUP: 'setup',
    DEBUG: 'debug'
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
      case 'login':
        return handleLoginModeWithTemplate(params);

      case 'admin':
        return handleAdminModeWithTemplate(params);

      case 'view':
        return handleViewMode(params);

      case 'appSetup':
        return handleAppSetupMode(params);

      case 'debug':
        return handleDebugMode(params);

      case 'test':
        return handleTestMode(params);

      case 'fix_user':
        return handleFixUserMode(params);

      case 'clear_cache':
        return handleClearCacheMode(params);

      case 'setup':
        return handleSetupModeWithTemplate(params);

      default:
        // Handle setupParam=true for legacy compatibility
        if (params.setupParam === 'true' || params.setupParam === true) {
          return renderSetupPageWithTemplate(params);
        }

        // Default flow: authentication check for main app
        const userEmail = UserService.getCurrentEmail();
        if (!userEmail) {
          console.log('doGet: No authentication, redirecting to login');
          return handleLoginModeWithTemplate(params, { reason: 'authentication_required' });
        }

        // Access control verification
        const accessResult = SecurityService.checkUserPermission(null, 'authenticated_user');
        if (!accessResult.hasPermission) {
          return renderErrorPage({
            success: false,
            message: 'アクセスが拒否されました',
            details: accessResult.reason,
            canRetry: false
          });
        }

        // System initialization check for authenticated users
        let systemReady = false;
        try {
          systemReady = DataService.isSystemSetup() || ConfigService.isSystemSetup();
        } catch (setupCheckError) {
          console.warn('System setup check error:', setupCheckError.message);
          systemReady = false;
        }

        if (!systemReady) {
          console.log('doGet: System not ready, redirecting to admin for setup');
          return handleAdminModeWithTemplate(params, { reason: 'setup_required' });
        }

        // Main application view (equivalent to historical mode=view)
        return handleViewMode(params);
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

    // Route to appropriate handler
    switch (request.action) {
      case 'getData':
        return handleGetData(request);
      
      case 'addReaction':
        return handleAddReaction(request);
      
      case 'toggleHighlight':
        return handleToggleHighlight(request);
      
      case 'refreshData':
        return handleRefreshData(request);
      
      default:
        throw new Error(`Unknown action: ${request.action}`);
    }
  } catch (error) {
    const errorResponse = ErrorHandler.handle(error, 'doPost');
    return ContentService.createTextOutput(JSON.stringify(errorResponse))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Request parameter parsing - enhanced for complete mode system
 */
function parseRequestParams(e) {
  const params = e.parameter || {};
  return {
    mode: params.mode || 'main',
    userId: params.userId,
    classFilter: params.classFilter,
    sortOrder: params.sortOrder || 'desc',
    adminMode: params.adminMode === 'true',
    // Historical compatibility parameters
    setupParam: params.setupParam,
    forced: params.forced,
    returnUrl: params.returnUrl,
    // Debug and test parameters
    debugMode: params.debugMode === 'true',
    testId: params.testId,
    // Additional utility parameters
    action: params.action,
    target: params.target
  };
}

/**
 * POST request parsing
 */
function parsePostRequest(e) {
  try {
    const postData = e.postData?.contents;
    if (!postData) {
      throw new Error('No POST data received');
    }
    
    return JSON.parse(postData);
  } catch (error) {
    throw new Error('Invalid JSON in POST data');
  }
}

/**
 * Main mode handler
 */
function handleMainMode(params) {
  try {
    const userInfo = UserService.getCurrentUserInfo();
    if (!userInfo) {
      return handleLoginMode(params, { reason: 'user_info_required' });
    }

    // Get bulk data using new DataService
    const bulkResult = DataService.getBulkData(userInfo.userId, {
      includeSheetData: true,
      includeSystemInfo: true
    });

    if (!bulkResult.success) {
      throw new Error(bulkResult.error);
    }

    // Render main page
    return renderMainPage(bulkResult.data);
  } catch (error) {
    console.error('handleMainMode error:', error.message);
    throw error;
  }
}

/**
 * Data retrieval handler
 */
function handleGetData(request) {
  try {
    // Get current user
    const userInfo = UserService.getCurrentUserInfo();
    if (!userInfo) {
      throw new Error('User information not found');
    }

    // Use new DataService instead of legacy getPublishedSheetData
    const options = {
      classFilter: request.classFilter,
      sortOrder: request.sortOrder,
      adminMode: request.adminMode
    };
    
    const result = DataService.getSheetData(userInfo.userId, options);
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.error('handleGetData error:', error.message);
    const errorResponse = ErrorHandler.handle(error, 'getData');
    return ContentService.createTextOutput(JSON.stringify(errorResponse))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Reaction handler
 */
function handleAddReaction(request) {
  try {
    const result = DataService.addReaction(
      (UserService.getCurrentUserInfo() || {}).userId,
      request.rowId,
      request.reactionType
    );

    return ContentService.createTextOutput(JSON.stringify({
      success: result,
      message: result ? 'Reaction added successfully' : 'Failed to add reaction'
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    const errorResponse = ErrorHandler.handle(error, 'addReaction');
    return ContentService.createTextOutput(JSON.stringify(errorResponse))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Highlight toggle handler
 */
function handleToggleHighlight(request) {
  try {
    const userInfo = UserService.getCurrentUserInfo();
    if (!userInfo) {
      throw new Error('User authentication required');
    }

    const result = DataService.toggleHighlight(userInfo.userId, request.rowId);

    return ContentService.createTextOutput(JSON.stringify({
      success: result,
      message: result ? 'Highlight toggled successfully' : 'Failed to toggle highlight'
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    const errorResponse = ErrorHandler.handle(error, 'toggleHighlight');
    return ContentService.createTextOutput(JSON.stringify(errorResponse))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Data refresh handler
 */
function handleRefreshData(request) {
  try {
    const userInfo = UserService.getCurrentUserInfo();
    if (!userInfo) {
      throw new Error('User authentication required');
    }

    // Clear cache and get fresh data
    const options = {
      classFilter: request.classFilter,
      sortOrder: request.sortOrder,
      adminMode: request.adminMode,
      useCache: false
    };
    
    const result = DataService.getSheetData(userInfo.userId, options);
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    const errorResponse = ErrorHandler.handle(error, 'refreshData');
    return ContentService.createTextOutput(JSON.stringify(errorResponse))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Simple page renderers
 */
function renderMainPage(data) {
  const template = HtmlService.createTemplateFromFile('Page');
  template.data = data;
  return template.evaluate()
    .setTitle(APP_CONFIG.APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function renderErrorPage(error) {
  return HtmlService.createHtmlOutput(`
    <h2>Error</h2>
    <p>${error.message}</p>
    ${error.details ? `<p>Details: ${error.details}</p>` : ''}
    <p><a href="${ScriptApp.getService().getUrl()}">Return to Home</a></p>
  `)
    .setTitle(`Error - ${  APP_CONFIG.APP_NAME}`)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Setup page renderer using proper HTML templates
 */
function renderSetupPageWithTemplate(params) {
  try {
    console.log('renderSetupPageWithTemplate: Rendering setup page');
    const template = HtmlService.createTemplateFromFile('SetupPage');
    template.params = params;
    return template.evaluate()
      .setTitle(`Setup - ${APP_CONFIG.APP_NAME}`)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (templateError) {
    console.warn('Setup template not found, using fallback:', templateError.message);
    return HtmlService.createHtmlOutput(`
      <h2>System Setup Required</h2>
      <p>The system needs to be configured before use.</p>
      <p><a href="?mode=setup">Start Setup</a></p>
      <p><a href="?setupParam=true">Legacy Setup</a></p>
    `)
      .setTitle(`Setup - ${APP_CONFIG.APP_NAME}`)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/**
 * Login mode handler using proper HTML templates
 */
function handleLoginModeWithTemplate(params, context = {}) {
  try {
    console.log('handleLoginModeWithTemplate: Rendering login page');
    const template = HtmlService.createTemplateFromFile('LoginPage');
    template.params = params;
    template.context = context;
    template.reason = context.reason || 'Authentication required';
    template.forced = params.forced === 'true';
    return template.evaluate()
      .setTitle(`Login - ${APP_CONFIG.APP_NAME}`)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (templateError) {
    console.warn('Login template not found, using fallback:', templateError.message);
    return HtmlService.createHtmlOutput(`
      <h2>Login Required</h2>
      <p>Please sign in to access ${APP_CONFIG.APP_NAME}.</p>
      <p>Reason: ${context.reason || 'Authentication required'}</p>
      <p><a href="${ScriptApp.getService().getUrl()}">Try Again</a></p>
    `)
      .setTitle(`Login - ${APP_CONFIG.APP_NAME}`)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/**
 * Admin mode handler using proper HTML templates
 */
function handleAdminModeWithTemplate(params, context = {}) {
  try {
    console.log('handleAdminModeWithTemplate: Rendering admin panel');

    // Security check for admin access
    const userEmail = UserService.getCurrentEmail();
    if (!userEmail) {
      return handleLoginModeWithTemplate(params, { reason: 'admin_access_requires_auth' });
    }

    const template = HtmlService.createTemplateFromFile('AdminPanel');
    template.params = params;
    template.context = context;
    template.userEmail = userEmail;
    template.isSystemAdmin = UserService.isSystemAdmin(userEmail);

    return template.evaluate()
      .setTitle(`Admin Panel - ${APP_CONFIG.APP_NAME}`)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (templateError) {
    console.warn('Admin template not found, using fallback:', templateError.message);
    return HtmlService.createHtmlOutput(`
      <h2>Admin Panel</h2>
      <p>Configuration and management interface</p>
      <p>User: ${UserService.getCurrentEmail() || 'Not authenticated'}</p>
      <p><a href="?mode=view">Return to Main App</a></p>
      <p><a href="?mode=setup">System Setup</a></p>
    `)
      .setTitle(`Admin - ${APP_CONFIG.APP_NAME}`)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/**
 * Setup mode handler using proper HTML templates
 */
function handleSetupModeWithTemplate(params) {
  try {
    console.log('handleSetupModeWithTemplate: Rendering setup interface');
    const template = HtmlService.createTemplateFromFile('SetupPage');
    template.params = params;
    template.mode = 'setup';
    return template.evaluate()
      .setTitle(`System Setup - ${APP_CONFIG.APP_NAME}`)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (templateError) {
    console.warn('Setup template not found, using fallback:', templateError.message);
    return HtmlService.createHtmlOutput(`
      <h2>System Setup</h2>
      <p>Initial system configuration interface</p>
      <p><a href="${ScriptApp.getService().getUrl()}">Return to Main</a></p>
    `)
      .setTitle(`Setup - ${APP_CONFIG.APP_NAME}`)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/**
 * Debug mode handler - system diagnostics and troubleshooting
 */
function handleDebugMode(params) {
  try {
    console.log('handleDebugMode: Running system diagnostics');
    const diagnostics = {
      timestamp: new Date().toISOString(),
      user: {
        email: UserService.getCurrentEmail(),
        userInfo: UserService.getCurrentUserInfo(),
        isSystemAdmin: UserService.isSystemAdmin()
      },
      services: {
        UserService: UserService.diagnose(),
        ConfigService: ConfigService.diagnose(),
        DataService: DataService.diagnose(),
        SecurityService: SecurityService.diagnose()
      },
      system: {
        hasCoreProps: ConfigService.hasCoreSystemProps(),
        isSystemSetup: ConfigService.isSystemSetup(),
        webAppUrl: getWebAppUrl(),
        scriptId: ScriptApp.getScriptId()
      },
      parameters: params
    };

    return HtmlService.createHtmlOutput(`
      <h2>System Diagnostics</h2>
      <div style="font-family: monospace; white-space: pre-wrap; background: #f5f5f5; padding: 10px; border: 1px solid #ddd;">${JSON.stringify(diagnostics, null, 2)}</div>
      <p><a href="${ScriptApp.getService().getUrl()}">Return to Main</a></p>
      <p><a href="?mode=admin">Admin Panel</a></p>
    `)
      .setTitle(`Debug - ${APP_CONFIG.APP_NAME}`)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('handleDebugMode error:', error.message);
    return renderErrorPage({ message: error.message });
  }
}

/**
 * View mode handler - the main application interface (historical mode=view)
 */
function handleViewMode(params) {
  try {
    console.log('handleViewMode: Rendering main application view');

    const userInfo = UserService.getCurrentUserInfo();
    if (!userInfo) {
      return handleLoginModeWithTemplate(params, { reason: 'user_info_required' });
    }

    // Get bulk data using new DataService
    const bulkResult = DataService.getBulkData(userInfo.userId, {
      includeSheetData: true,
      includeSystemInfo: true,
      classFilter: params.classFilter,
      sortOrder: params.sortOrder,
      adminMode: params.adminMode
    });

    if (!bulkResult.success) {
      console.error('handleViewMode: Data retrieval failed:', bulkResult.error);
      return renderErrorPage({
        success: false,
        message: 'データの取得に失敗しました',
        details: bulkResult.error
      });
    }

    // Render main application page
    const template = HtmlService.createTemplateFromFile('Page');
    template.data = bulkResult.data;
    template.params = params;
    template.userInfo = userInfo;

    return template.evaluate()
      .setTitle(`${APP_CONFIG.APP_NAME} - Answer Board`)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('handleViewMode error:', error.message);
    throw error;
  }
}

/**
 * App setup mode handler - application configuration interface
 */
function handleAppSetupMode(params) {
  try {
    console.log('handleAppSetupMode: Rendering app setup interface');

    // Require authentication for app setup
    const userEmail = UserService.getCurrentEmail();
    if (!userEmail) {
      return handleLoginModeWithTemplate(params, { reason: 'setup_requires_auth' });
    }

    const template = HtmlService.createTemplateFromFile('AppSetupPage');
    template.params = params;
    template.userEmail = userEmail;
    template.isSystemAdmin = UserService.isSystemAdmin(userEmail);

    return template.evaluate()
      .setTitle(`App Setup - ${APP_CONFIG.APP_NAME}`)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (templateError) {
    console.warn('AppSetupPage template not found, using fallback:', templateError.message);
    return HtmlService.createHtmlOutput(`
      <h2>App Setup</h2>
      <p>Application configuration interface</p>
      <p>User: ${UserService.getCurrentEmail() || 'Not authenticated'}</p>
      <p><a href="?mode=admin">Admin Panel</a></p>
      <p><a href="?mode=view">Return to Main App</a></p>
    `)
      .setTitle(`App Setup - ${APP_CONFIG.APP_NAME}`)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/**
 * Test mode handler - system testing and validation
 */
function handleTestMode(params) {
  try {
    console.log('handleTestMode: Running system tests');

    // Basic system tests
    const testResults = {
      timestamp: new Date().toISOString(),
      tests: [],
      passed: 0,
      failed: 0
    };

    // Test 1: User authentication
    try {
      const userEmail = UserService.getCurrentEmail();
      testResults.tests.push({
        name: 'User Authentication',
        status: userEmail ? 'PASS' : 'FAIL',
        result: userEmail || 'No user email found'
      });
      if (userEmail) testResults.passed++; else testResults.failed++;
    } catch (e) {
      testResults.tests.push({
        name: 'User Authentication',
        status: 'ERROR',
        result: e.message
      });
      testResults.failed++;
    }

    // Test 2: Services availability
    const services = ['UserService', 'ConfigService', 'DataService', 'SecurityService'];
    services.forEach(serviceName => {
      try {
        const serviceObj = eval(serviceName);
        const diagnose = serviceObj.diagnose();
        testResults.tests.push({
          name: serviceName,
          status: diagnose.status === 'healthy' ? 'PASS' : 'FAIL',
          result: diagnose.message || 'Available'
        });
        if (diagnose.status === 'healthy') testResults.passed++; else testResults.failed++;
      } catch (e) {
        testResults.tests.push({
          name: serviceName,
          status: 'ERROR',
          result: e.message
        });
        testResults.failed++;
      }
    });

    return HtmlService.createHtmlOutput(`
      <h2>System Test Results</h2>
      <p>Test ID: ${params.testId || 'manual'}</p>
      <p>Status: ${testResults.failed === 0 ? '✅ All tests passed' : '❌ Some tests failed'}</p>
      <p>Results: ${testResults.passed} passed, ${testResults.failed} failed</p>
      <div style="font-family: monospace; white-space: pre-wrap; background: #f5f5f5; padding: 10px; border: 1px solid #ddd;">${JSON.stringify(testResults, null, 2)}</div>
      <p><a href="${ScriptApp.getService().getUrl()}">Return to Main</a></p>
      <p><a href="?mode=debug">Debug Mode</a></p>
    `)
      .setTitle(`Test Results - ${APP_CONFIG.APP_NAME}`)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('handleTestMode error:', error.message);
    return renderErrorPage({ message: error.message });
  }
}

/**
 * Fix user mode handler - user account repair utilities
 */
function handleFixUserMode(params) {
  try {
    console.log('handleFixUserMode: User account repair mode');

    const userEmail = UserService.getCurrentEmail();
    if (!userEmail) {
      return handleLoginModeWithTemplate(params, { reason: 'fix_user_requires_auth' });
    }

    const fixResults = {
      timestamp: new Date().toISOString(),
      userEmail: userEmail,
      fixes: [],
      successful: 0,
      failed: 0
    };

    // Fix 1: Ensure user exists in database
    try {
      let user = UserService.getCurrentUserInfo();
      if (!user) {
        const createResult = UserService.createUser(userEmail);
        if (createResult.success) {
          fixResults.fixes.push('✅ Created missing user record');
          fixResults.successful++;
          user = UserService.getCurrentUserInfo();
        } else {
          fixResults.fixes.push('❌ Failed to create user record: ' + createResult.error);
          fixResults.failed++;
        }
      } else {
        fixResults.fixes.push('✅ User record exists');
        fixResults.successful++;
      }
    } catch (e) {
      fixResults.fixes.push('❌ User fix error: ' + e.message);
      fixResults.failed++;
    }

    // Fix 2: Validate and repair config
    try {
      const userInfo = UserService.getCurrentUserInfo();
      if (userInfo) {
        const config = ConfigService.getUserConfig(userInfo.userId) || {};

        if (!config.version) {
          config.version = '1.0.0';
          fixResults.fixes.push('✅ Added missing version to config');
        }

        if (!config.lastModified) {
          config.lastModified = new Date().toISOString();
          fixResults.fixes.push('✅ Added missing lastModified to config');
        }

        const saveResult = ConfigService.saveUserConfig(userInfo.userId, config);
        if (saveResult.success) {
          fixResults.fixes.push('✅ Config validation and repair completed');
          fixResults.successful++;
        } else {
          fixResults.fixes.push('❌ Config save failed: ' + saveResult.error);
          fixResults.failed++;
        }
      }
    } catch (e) {
      fixResults.fixes.push('❌ Config fix error: ' + e.message);
      fixResults.failed++;
    }

    return HtmlService.createHtmlOutput(`
      <h2>User Account Repair</h2>
      <p>User: ${userEmail}</p>
      <p>Status: ${fixResults.failed === 0 ? '✅ All repairs successful' : '⚠️ Some repairs failed'}</p>
      <h3>Repair Results:</h3>
      <ul>${fixResults.fixes.map(fix => '<li>' + fix + '</li>').join('')}</ul>
      <div style="font-family: monospace; white-space: pre-wrap; background: #f5f5f5; padding: 10px; border: 1px solid #ddd; margin-top: 10px;">${JSON.stringify(fixResults, null, 2)}</div>
      <p><a href="${ScriptApp.getService().getUrl()}">Return to Main</a></p>
      <p><a href="?mode=admin">Admin Panel</a></p>
    `)
      .setTitle(`User Repair - ${APP_CONFIG.APP_NAME}`)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('handleFixUserMode error:', error.message);
    return renderErrorPage({ message: error.message });
  }
}

/**
 * Clear cache mode handler - cache management utilities
 */
function handleClearCacheMode(params) {
  try {
    console.log('handleClearCacheMode: Cache clearing utilities');

    const clearResults = {
      timestamp: new Date().toISOString(),
      operations: [],
      successful: 0,
      failed: 0
    };

    // Clear script cache
    try {
      const scriptCache = CacheService.getScriptCache();
      scriptCache.removeAll(['user_session', 'app_state', 'config_cache', 'data_cache']);
      clearResults.operations.push('✅ Script cache cleared');
      clearResults.successful++;
    } catch (e) {
      clearResults.operations.push('❌ Script cache clear failed: ' + e.message);
      clearResults.failed++;
    }

    // Clear user cache if authenticated
    try {
      const userEmail = UserService.getCurrentEmail();
      if (userEmail) {
        const userCache = CacheService.getUserCache();
        userCache.removeAll(['user_data', 'sheet_data', 'permissions']);
        clearResults.operations.push('✅ User cache cleared');
        clearResults.successful++;
      } else {
        clearResults.operations.push('ℹ️ No user authenticated, user cache not cleared');
      }
    } catch (e) {
      clearResults.operations.push('❌ User cache clear failed: ' + e.message);
      clearResults.failed++;
    }

    // Clear document cache (if available)
    try {
      if (typeof CacheService.getDocumentCache === 'function') {
        const docCache = CacheService.getDocumentCache();
        docCache.removeAll(['temp_data']);
        clearResults.operations.push('✅ Document cache cleared');
        clearResults.successful++;
      } else {
        clearResults.operations.push('ℹ️ Document cache not available');
      }
    } catch (e) {
      clearResults.operations.push('❌ Document cache clear failed: ' + e.message);
      clearResults.failed++;
    }

    return HtmlService.createHtmlOutput(`
      <h2>Cache Management</h2>
      <p>Target: ${params.target || 'all'}</p>
      <p>Status: ${clearResults.failed === 0 ? '✅ All operations successful' : '⚠️ Some operations failed'}</p>
      <h3>Operations:</h3>
      <ul>${clearResults.operations.map(op => '<li>' + op + '</li>').join('')}</ul>
      <div style="font-family: monospace; white-space: pre-wrap; background: #f5f5f5; padding: 10px; border: 1px solid #ddd; margin-top: 10px;">${JSON.stringify(clearResults, null, 2)}</div>
      <p><a href="${ScriptApp.getService().getUrl()}">Return to Main</a></p>
      <p><a href="?mode=debug">Debug Mode</a></p>
    `)
      .setTitle(`Cache Management - ${APP_CONFIG.APP_NAME}`)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('handleClearCacheMode error:', error.message);
    return renderErrorPage({ message: error.message });
  }
}

// ==============================================================================
// フロントエンドAPI関数（HTMLから呼び出される重要な関数群）
// ==============================================================================

/**
 * 現在のユーザー情報を取得
 * login.js.html, SetupPage.html, AdminPanel.js.html から呼び出される
 * 
 * @param {string} [kind='email'] - 取得する情報の種類（'email' or 'full'）
 * @returns {Object|null} 統一されたユーザー情報オブジェクト
 */
function getUser(kind = 'email') {
  try {
    const userEmail = UserService.getCurrentEmail();
    
    if (!userEmail) {
      return null;
    }
    
    // ✅ 修正: 戻り値型を統一（常にオブジェクト）
    const userInfo = UserService.getCurrentUserInfo();
    const result = {
      email: userEmail,
      userId: userInfo?.userId || null,
      isActive: userInfo?.isActive || false,
      hasConfig: !!userInfo?.config
    };
    
    // 後方互換性: emailのみ必要な場合は email フィールドに文字列が入っている
    if (kind === 'email') {
      // フロントエンドで user.email でアクセス可能
      result.value = userEmail; // レガシー対応用
    }
    
    return result;
  } catch (error) {
    console.error('getUser エラー:', error.message);
    return null;
  }
}

/**
 * WebアプリケーションのURLを取得
 * 複数のHTMLファイルから呼び出される基本機能
 * 
 * @returns {string} WebアプリのURL
 */
function getWebAppUrl() {
  try {
    return ScriptApp.getService().getUrl();
  } catch (error) {
    console.error('getWebAppUrl エラー:', error.message);
    // フォールバック URL を返す
    return 'https://script.google.com';
  }
}

/**
 * ログイン処理を実行
 * login.js.html から呼び出される統合ログイン処理
 * 
 * @returns {Object} ログイン結果
 */
function processLoginAction() {
  try {
    const userEmail = UserService.getCurrentEmail();
    
    if (!userEmail) {
      return {
        success: false,
        message: 'ログインが必要です',
        requireAuth: true
      };
    }
    
    // ユーザー情報を取得または作成
    let userInfo = UserService.getCurrentUserInfo();
    
    if (!userInfo) {
      // 新規ユーザーの場合は作成
      const createResult = UserService.createUser(userEmail);
      if (!createResult.success) {
        return {
          success: false,
          message: createResult.error || 'ユーザー作成に失敗しました'
        };
      }
      userInfo = UserService.getCurrentUserInfo();
    }
    
    // セットアップ状態の確認
    const config = ConfigService.getUserConfig(userInfo.userId);
    const needsSetup = !config?.spreadsheetId || config?.setupStatus !== 'completed';
    
    return {
      success: true,
      userId: userInfo.userId,
      email: userEmail,
      needsSetup,
      redirectUrl: needsSetup ? `${getWebAppUrl()}?mode=admin` : getWebAppUrl()
    };
  } catch (error) {
    console.error('processLoginAction エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * URL内部状態をリセット
 * login.js.html から呼び出される（主にログ記録用）
 * 
 * @returns {Object} リセット結果
 */
function forceUrlSystemReset() {
  try {
    console.log('URL システムリセットが要求されました');
    
    // キャッシュのクリア
    try {
      const cache = CacheService.getScriptCache();
      cache.removeAll(['user_session', 'app_state']);
    } catch (cacheError) {
      console.warn('キャッシュクリアエラー:', cacheError.message);
    }
    
    return {
      success: true,
      message: 'システムがリセットされました',
      redirectUrl: getWebAppUrl()
    };
  } catch (error) {
    console.error('forceUrlSystemReset エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

// ==============================================================================
// セットアップ・管理関数（SetupPage.html, AppSetupPage.html用）
// ==============================================================================

/**
 * アプリケーションの初期セットアップ
 * SetupPage.html から呼び出される
 * 
 * @param {string} serviceAccountJson - Service Account JSON文字列
 * @param {string} databaseId - データベーススプレッドシートID
 * @param {string} adminEmail - 管理者メールアドレス
 * @param {string} [googleClientId] - Google Client ID (オプション)
 * @returns {Object} セットアップ結果
 */
function setupApplication(serviceAccountJson, databaseId, adminEmail, googleClientId) {
  try {
    // 権限チェック（初回セットアップまたは管理者のみ）
    const currentEmail = UserService.getCurrentEmail();
    const props = PropertiesService.getScriptProperties();
    const existingAdmin = props.getProperty(PROPS_KEYS.ADMIN_EMAIL);
    
    if (existingAdmin && currentEmail !== existingAdmin) {
      return {
        success: false,
        message: '管理者権限が必要です'
      };
    }
    
    // Service Account資格情報の検証と保存
    try {
      const saCredentials = JSON.parse(serviceAccountJson);
      if (!saCredentials.private_key || !saCredentials.client_email) {
        throw new Error('Service Account資格情報が不正です');
      }
      props.setProperty(PROPS_KEYS.SERVICE_ACCOUNT_CREDS, serviceAccountJson);
    } catch (parseError) {
      return {
        success: false,
        message: `Service Account JSON の形式が不正です: ${  parseError.message}`
      };
    }
    
    // データベースIDの検証と保存
    if (!databaseId || databaseId.length < 10) {
      return {
        success: false,
        message: 'データベーススプレッドシートIDが不正です'
      };
    }
    props.setProperty(PROPS_KEYS.DATABASE_SPREADSHEET_ID, databaseId);
    
    // 管理者メールの検証と保存
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(adminEmail)) {
      return {
        success: false,
        message: '管理者メールアドレスの形式が不正です'
      };
    }
    props.setProperty(PROPS_KEYS.ADMIN_EMAIL, adminEmail);
    
    // Google Client IDの保存（オプション）
    if (googleClientId) {
      props.setProperty(PROPS_KEYS.GOOGLE_CLIENT_ID, googleClientId);
    }
    
    // システム初期化フラグを設定
    props.setProperty('SYSTEM_INITIALIZED', 'true');
    
    return {
      success: true,
      message: 'セットアップが完了しました',
      redirectUrl: `${getWebAppUrl()  }?mode=login`
    };
  } catch (error) {
    console.error('setupApplication エラー:', error.message);
    return {
      success: false,
      message: `セットアップ中にエラーが発生しました: ${  error.message}`
    };
  }
}

/**
 * セットアップのテスト実行
 * SetupPage.html から呼び出される
 * 
 * @returns {Object} テスト結果
 */
function testSetup() {
  try {
    const results = {
      serviceAccount: false,
      database: false,
      adminEmail: false,
      webAppUrl: false
    };
    
    const props = PropertiesService.getScriptProperties();
    
    // Service Account 確認
    try {
      const saCreds = props.getProperty(PROPS_KEYS.SERVICE_ACCOUNT_CREDS);
      results.serviceAccount = !!saCreds;
    } catch (e) {
      console.warn('Service Account チェックエラー:', e.message);
    }
    
    // データベース接続確認
    try {
      const dbId = props.getProperty(PROPS_KEYS.DATABASE_SPREADSHEET_ID);
      if (dbId) {
        const sheet = SpreadsheetApp.openById(dbId);
        results.database = !!sheet;
      }
    } catch (e) {
      console.warn('データベースチェックエラー:', e.message);
    }
    
    // 管理者メール確認
    try {
      const adminEmail = props.getProperty(PROPS_KEYS.ADMIN_EMAIL);
      results.adminEmail = !!adminEmail;
    } catch (e) {
      console.warn('管理者メールチェックエラー:', e.message);
    }
    
    // WebApp URL 確認
    try {
      results.webAppUrl = !!getWebAppUrl();
    } catch (e) {
      console.warn('WebApp URL チェックエラー:', e.message);
    }
    
    const allPassed = Object.values(results).every(v => v === true);
    
    return {
      success: allPassed,
      results,
      message: allPassed ? 'すべてのテストに合格しました' : '一部のテストに失敗しました'
    };
  } catch (error) {
    console.error('testSetup エラー:', error.message);
    return {
      success: false,
      message: `テスト中にエラーが発生しました: ${  error.message}`
    };
  }
}

// ==============================================================================
// 管理パネル用API関数（AdminPanel.js.html, AppSetupPage.html用）
// ==============================================================================

/**
 * 現在の設定を取得
 * AdminPanel.js.html から呼び出される
 * 
 * @returns {Object} 設定情報
 */
function getConfig() {
  try {
    const userInfo = UserService.getCurrentUserInfo();
    if (!userInfo) {
      return {
        success: false,
        message: 'ユーザー情報が見つかりません'
      };
    }
    
    const config = ConfigService.getUserConfig(userInfo.userId);
    return {
      success: true,
      config: config || {}
    };
  } catch (error) {
    console.error('getConfig エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * スプレッドシート一覧を取得
 * AdminPanel.js.html, AppSetupPage.html から呼び出される
 * 
 * @returns {Object} スプレッドシート一覧
 */
function getSpreadsheetList() {
  try {
    const files = DriveApp.getFilesByType('application/vnd.google-apps.spreadsheet');
    const spreadsheets = [];
    let count = 0;
    const maxCount = 50; // 最大50件まで
    
    while (files.hasNext() && count < maxCount) {
      const file = files.next();
      spreadsheets.push({
        id: file.getId(),
        name: file.getName(),
        url: file.getUrl(),
        lastUpdated: file.getLastUpdated()
      });
      count++;
    }
    
    return {
      success: true,
      spreadsheets
    };
  } catch (error) {
    console.error('getSpreadsheetList エラー:', error.message);
    return {
      success: false,
      message: error.message,
      spreadsheets: []
    };
  }
}

/**
 * シート一覧を取得
 * AdminPanel.js.html, AppSetupPage.html から呼び出される
 * 
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {Object} シート一覧
 */
function getSheetList(spreadsheetId) {
  try {
    if (!spreadsheetId) {
      return {
        success: false,
        message: 'スプレッドシートIDが指定されていません'
      };
    }
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheets = spreadsheet.getSheets();
    const sheetList = sheets.map(sheet => ({
      name: sheet.getName(),
      index: sheet.getIndex(),
      rowCount: sheet.getMaxRows(),
      columnCount: sheet.getMaxColumns()
    }));
    
    return {
      success: true,
      sheets: sheetList
    };
  } catch (error) {
    console.error('getSheetList エラー:', error.message);
    return {
      success: false,
      message: error.message,
      sheets: []
    };
  }
}

/**
 * 列を分析
 * AdminPanel.js.html, AppSetupPage.html から呼び出される
 * 
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} 列分析結果
 */
function analyzeColumns(spreadsheetId, sheetName) {
  try {
    if (!spreadsheetId || !sheetName) {
      return {
        success: false,
        message: 'スプレッドシートIDとシート名が必要です'
      };
    }
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      return {
        success: false,
        message: 'シートが見つかりません'
      };
    }
    
    // ヘッダー行を取得（1行目）
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // サンプルデータを取得（最大5行）
    const sampleRowCount = Math.min(5, sheet.getLastRow() - 1);
    let sampleData = [];
    if (sampleRowCount > 0) {
      sampleData = sheet.getRange(2, 1, sampleRowCount, sheet.getLastColumn()).getValues();
    }
    
    // 列情報を分析
    const columns = headers.map((header, index) => {
      const samples = sampleData.map(row => row[index]).filter(v => v);
      
      // 列タイプを推測
      let type = 'text';
      if (header.toLowerCase().includes('timestamp') || header.toLowerCase().includes('日時')) {
        type = 'datetime';
      } else if (header.toLowerCase().includes('email') || header.toLowerCase().includes('メール')) {
        type = 'email';
      } else if (header.toLowerCase().includes('class') || header.toLowerCase().includes('クラス')) {
        type = 'class';
      } else if (header.toLowerCase().includes('name') || header.toLowerCase().includes('名前')) {
        type = 'name';
      } else if (samples.length > 0 && samples.every(s => !isNaN(s))) {
        type = 'number';
      }
      
      return {
        index,
        header,
        type,
        samples: samples.slice(0, 3) // 最大3つのサンプル
      };
    });
    
    return {
      success: true,
      columns,
      totalRows: sheet.getLastRow(),
      totalColumns: sheet.getLastColumn()
    };
  } catch (error) {
    console.error('analyzeColumns エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * フォーム情報を取得
 * AdminPanel.js.html, AppSetupPage.html から呼び出される
 * 
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} フォーム情報
 */
function getFormInfo(_spreadsheetId, _sheetName) {
  // 現在はフォーム連携機能は未実装
  return {
    success: true,
    hasForm: false,
    formUrl: null,
    formTitle: null,
    message: 'フォーム連携機能は将来のアップデートで提供予定です'
  };
}

/**
 * 設定の下書きを保存
 * AdminPanel.js.html から呼び出される
 * 
 * @param {Object} config - 保存する設定
 * @returns {Object} 保存結果
 */
function saveDraftConfiguration(config) {
  try {
    const userInfo = UserService.getCurrentUserInfo();
    if (!userInfo) {
      return {
        success: false,
        message: 'ユーザー情報が見つかりません'
      };
    }
    
    // ConfigServiceを使用して設定を保存
    const result = ConfigService.saveUserConfig(userInfo.userId, config);
    
    if (result.success) {
      return {
        success: true,
        message: '設定を保存しました'
      };
    } else {
      return {
        success: false,
        message: result.error || '設定の保存に失敗しました'
      };
    }
  } catch (error) {
    console.error('saveDraftConfiguration エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * アプリケーションを公開
 * AdminPanel.js.html から呼び出される
 * 
 * @param {Object} publishConfig - 公開設定
 * @returns {Object} 公開結果
 */
function publishApplication(publishConfig) {
  try {
    const userInfo = UserService.getCurrentUserInfo();
    if (!userInfo) {
      return {
        success: false,
        message: 'ユーザー情報が見つかりません'
      };
    }
    
    // 現在の設定を取得
    const currentConfig = ConfigService.getUserConfig(userInfo.userId) || {};
    
    // 公開設定をマージ
    const updatedConfig = {
      ...currentConfig,
      ...publishConfig,
      appPublished: true,
      publishedAt: new Date().toISOString(),
      setupStatus: 'completed'
    };
    
    // 設定を保存
    const result = ConfigService.saveUserConfig(userInfo.userId, updatedConfig);
    
    if (result.success) {
      return {
        success: true,
        message: 'アプリケーションを公開しました',
        appUrl: getWebAppUrl()
      };
    } else {
      return {
        success: false,
        message: result.error || '公開に失敗しました'
      };
    }
  } catch (error) {
    console.error('publishApplication エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * システム管理者かどうかを確認
 * AdminPanel.js.html から呼び出される
 * 
 * @returns {boolean} システム管理者かどうか
 */
function checkIsSystemAdmin() {
  try {
    return UserService.isSystemAdmin();
  } catch (error) {
    console.error('checkIsSystemAdmin エラー:', error.message);
    return false;
  }
}

/**
 * 現在のボード情報とURLを取得
 * AdminPanel.js.html から呼び出される
 * 
 * @returns {Object} ボード情報
 */
function getCurrentBoardInfoAndUrls() {
  try {
    const userInfo = UserService.getCurrentUserInfo();
    if (!userInfo) {
      return {
        success: false,
        message: 'ユーザー情報が見つかりません'
      };
    }
    
    const config = ConfigService.getUserConfig(userInfo.userId);
    const baseUrl = getWebAppUrl();
    
    return {
      success: true,
      boardInfo: {
        userId: userInfo.userId,
        email: userInfo.userEmail || UserService.getCurrentEmail(),
        spreadsheetId: config?.spreadsheetId,
        sheetName: config?.sheetName,
        isPublished: config?.appPublished || false,
        publishedAt: config?.publishedAt
      },
      urls: {
        viewUrl: baseUrl,
        adminUrl: `${baseUrl}?mode=admin`,
        debugUrl: `${baseUrl}?mode=debug`
      }
    };
  } catch (error) {
    console.error('getCurrentBoardInfoAndUrls エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * スプレッドシートへのアクセスを検証
 * AdminPanel.js.html から呼び出される
 * 
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {Object} アクセス検証結果
 */
function validateAccess(spreadsheetId) {
  try {
    if (!spreadsheetId) {
      return {
        success: false,
        hasAccess: false,
        message: 'スプレッドシートIDが指定されていません'
      };
    }
    
    // アクセスを試行
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const name = spreadsheet.getName();
    
    return {
      success: true,
      hasAccess: true,
      spreadsheetName: name,
      message: 'アクセス可能です'
    };
  } catch (error) {
    console.error('validateAccess エラー:', error.message);
    return {
      success: false,
      hasAccess: false,
      message: `スプレッドシートにアクセスできません: ${  error.message}`
    };
  }
}

// ==============================================================================
// その他のユーティリティ関数
// ==============================================================================

/**
 * クライアントエラーを報告
 * ErrorBoundary.html から呼び出される
 * 
 * @param {Object} errorInfo - エラー情報
 * @returns {Object} 報告結果
 */
function reportClientError(errorInfo) {
  try {
    console.error('クライアントエラー報告:', errorInfo);
    
    // エラーログを記録（将来的にはSecurityServiceや専用のログサービスに委譲）
    const logEntry = {
      timestamp: new Date().toISOString(),
      type: 'client_error',
      userEmail: UserService.getCurrentEmail() || 'unknown',
      errorInfo
    };
    
    // コンソールにログ出力（将来的には永続化）
    console.log('Error Log Entry:', JSON.stringify(logEntry));
    
    return {
      success: true,
      message: 'エラーが報告されました'
    };
  } catch (error) {
    console.error('reportClientError エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * 強制ログアウトとリダイレクトのテスト
 * ErrorBoundary.html から呼び出される
 * 
 * @returns {Object} テスト結果
 */
function testForceLogoutRedirect() {
  try {
    console.log('強制ログアウトテストが実行されました');
    
    return {
      success: true,
      message: 'ログアウトテスト完了',
      redirectUrl: `${getWebAppUrl()  }?mode=login`
    };
  } catch (error) {
    console.error('testForceLogoutRedirect エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * 認証をリセット
 * SharedUtilities.html から呼び出される
 * 
 * @returns {Object} リセット結果
 */
function resetAuth() {
  try {
    console.log('認証リセットが要求されました');
    
    // キャッシュをクリア
    const cache = CacheService.getScriptCache();
    cache.removeAll(['user_session', 'auth_token']);
    
    return {
      success: true,
      message: '認証がリセットされました',
      redirectUrl: `${getWebAppUrl()  }?mode=login`
    };
  } catch (error) {
    console.error('resetAuth エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * ユーザー認証を検証
 * SharedUtilities.html から呼び出される
 * 
 * @returns {Object} 認証検証結果
 */
function verifyUserAuthentication() {
  try {
    const userEmail = UserService.getCurrentEmail();
    const userInfo = UserService.getCurrentUserInfo();
    
    return {
      success: true,
      authenticated: !!userEmail,
      email: userEmail,
      hasUserInfo: !!userInfo,
      userId: userInfo?.userId
    };
  } catch (error) {
    console.error('verifyUserAuthentication エラー:', error.message);
    return {
      success: false,
      authenticated: false,
      message: error.message
    };
  }
}

/**
 * スプレッドシートURLを追加
 * Unpublished.html から呼び出される
 * 
 * @param {string} url - スプレッドシートURL
 * @returns {Object} 追加結果
 */
function addSpreadsheetUrl(url) {
  try {
    // URLからスプレッドシートIDを抽出
    const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (!match) {
      return {
        success: false,
        message: '無効なスプレッドシートURLです'
      };
    }
    
    const spreadsheetId = match[1];
    
    // アクセスを確認
    try {
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      const name = spreadsheet.getName();
    } catch (accessError) {
      return {
        success: false,
        message: 'スプレッドシートにアクセスできません'
      };
    }
    
    // ユーザー情報を取得
    const userInfo = UserService.getCurrentUserInfo();
    if (!userInfo) {
      return {
        success: false,
        message: 'ユーザー情報が見つかりません'
      };
    }
    
    // 設定を更新
    const config = ConfigService.getUserConfig(userInfo.userId) || {};
    config.spreadsheetId = spreadsheetId;
    config.spreadsheetUrl = url;
    
    const result = ConfigService.saveUserConfig(userInfo.userId, config);
    
    if (result.success) {
      return {
        success: true,
        message: 'スプレッドシートを設定しました',
        spreadsheetId
      };
    } else {
      return {
        success: false,
        message: result.error || '設定の保存に失敗しました'
      };
    }
  } catch (error) {
    console.error('addSpreadsheetUrl エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * 管理者全ユーザー取得（システム管理者専用）
 * @param {Object} options - 取得オプション
 * @returns {Object} ユーザー一覧または権限エラー
 */
function getAllUsersForAdminForUI(options = {}) {
  try {
    console.log('getAllUsersForAdminForUI: システム管理者機能実行');
    
    // 🔒 二重セキュリティチェック: セッション+権限
    const currentEmail = UserService.getCurrentEmail();
    if (!currentEmail) {
      return {
        success: false,
        message: '認証が必要です'
      };
    }
    
    // システム管理者権限チェック（メールベース）
    const isAdmin = UserService.isSystemAdmin(currentEmail);
    if (!isAdmin) {
      // 🚨 権限違反をログ記録
      SecurityService.persistSecurityLog({
        event: 'UNAUTHORIZED_ADMIN_ACCESS',
        severity: 'HIGH',
        details: { attemptedBy: currentEmail, function: 'getAllUsersForAdminForUI' }
      });
      
      return {
        success: false,
        message: 'この機能にはシステム管理者権限が必要です'
      };
    }
    
    // 全ユーザー取得
    const users = DB.getAllUsers();
    const userList = users.map(user => ({
      userId: user.userId,
      userEmail: user.userEmail,
      isActive: user.isActive,
      lastModified: user.lastModified,
      configSummary: (() => {
        try {
          const config = JSON.parse(user.configJson || '{}');
          return {
            setupStatus: config.setupStatus || 'pending',
            appPublished: config.appPublished || false,
            spreadsheetId: config.spreadsheetId ? '設定済み' : '未設定',
            lastAccessedAt: config.lastAccessedAt || '不明'
          };
        } catch (e) {
          return { error: 'config解析エラー' };
        }
      })()
    }));
    
    return {
      success: true,
      users: userList,
      totalCount: userList.length,
      activeCount: userList.filter(u => u.isActive).length
    };
  } catch (error) {
    console.error('getAllUsersForAdminForUI エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * ユーザーアカウント削除（システム管理者専用）
 * @param {string} targetUserId - 削除対象のユーザーID
 * @returns {Object} 削除結果
 */
function deleteUserAccountByAdminForUI(targetUserId) {
  try {
    console.log('deleteUserAccountByAdminForUI: ユーザー削除実行:', targetUserId);
    
    // 🔒 厳格な認証チェック
    const currentEmail = UserService.getCurrentEmail();
    if (!currentEmail) {
      return {
        success: false,
        message: '認証が必要です'
      };
    }
    
    // システム管理者権限チェック（メールベース）
    const isAdmin = UserService.isSystemAdmin(currentEmail);
    if (!isAdmin) {
      // 🚨 重大な権限違反をログ記録
      SecurityService.persistSecurityLog({
        event: 'UNAUTHORIZED_USER_DELETE_ATTEMPT',
        severity: 'CRITICAL',
        details: { 
          attemptedBy: currentEmail, 
          targetUserId,
          timestamp: new Date().toISOString()
        }
      });
      
      return {
        success: false,
        message: 'この機能にはシステム管理者権限が必要です'
      };
    }
    
    if (!targetUserId) {
      return {
        success: false,
        message: '削除対象のユーザーIDが指定されていません'
      };
    }
    
    // 自分自身の削除防止
    const currentUserInfo = UserService.getCurrentUserInfo();
    if (currentUserInfo && currentUserInfo.userId === targetUserId) {
      return {
        success: false,
        message: '自分自身のアカウントは削除できません'
      };
    }
    
    // 対象ユーザー存在確認
    const targetUser = DB.findUserById(targetUserId);
    if (!targetUser) {
      return {
        success: false,
        message: '削除対象のユーザーが見つかりません'
      };
    }
    
    // ユーザー削除実行
    const deleteResult = DB.deleteUser(targetUserId);
    if (deleteResult) {
      // 削除ログの記録
      SecurityService.persistSecurityLog({
        event: 'USER_DELETION',
        severity: 'HIGH',
        details: {
          deletedUserId: targetUserId,
          deletedUserEmail: targetUser.userEmail,
          deletedBy: currentUserInfo?.userEmail || 'unknown'
        }
      });
      
      return {
        success: true,
        message: `ユーザー ${targetUser.userEmail} を削除しました`,
        deletedUser: {
          userId: targetUserId,
          userEmail: targetUser.userEmail
        }
      };
    } else {
      return {
        success: false,
        message: 'ユーザー削除に失敗しました'
      };
    }
  } catch (error) {
    console.error('deleteUserAccountByAdminForUI エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * 強制ログアウト・リダイレクト（システム管理者専用）
 * @returns {HtmlOutput} ログアウト処理のHTML
 */
function forceLogoutAndRedirectToLogin() {
  try {
    console.log('forceLogoutAndRedirectToLogin: 強制ログアウト実行');
    
    // システム管理者権限チェック（ログあり）
    const isAdmin = UserService.isSystemAdmin();
    if (isAdmin) {
      console.log('forceLogoutAndRedirectToLogin: システム管理者による実行');
    }
    
    // セッション情報クリア（可能な範囲で）
    try {
      // GASでは完全なセッションクリアは困難だが、ログは記録
      const currentUserInfo = UserService.getCurrentUserInfo();
      if (currentUserInfo) {
        SecurityService.persistSecurityLog({
          event: 'FORCE_LOGOUT',
          severity: 'MEDIUM',
          details: {
            userEmail: currentUserInfo.userEmail,
            triggeredBy: isAdmin ? 'system_admin' : 'user'
          }
        });
      }
    } catch (logError) {
      console.warn('forceLogoutAndRedirectToLogin: ログ記録エラー:', logError.message);
    }
    
    // ログイン画面へのリダイレクト
    const loginUrl = `${getWebAppUrl()  }?mode=login&forced=true`;
    return createRedirect(loginUrl);
  } catch (error) {
    console.error('forceLogoutAndRedirectToLogin エラー:', error.message);
    return HtmlService.createHtmlOutput(`
      <h2>ログアウト処理エラー</h2>
      <p>エラー: ${error.message}</p>
      <p><a href="${getWebAppUrl()}">ホームに戻る</a></p>
    `).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/**
 * データ関連関数：削除ログ取得（システム管理者専用）
 * @param {Object} options - 取得オプション（期間、制限数など）
 * @returns {Object} 削除ログ一覧
 */
function getDeletionLogsForUI(options = {}) {
  try {
    console.log('getDeletionLogsForUI: 削除ログ取得');
    
    // システム管理者権限チェック
    const isAdmin = UserService.isSystemAdmin();
    if (!isAdmin) {
      return {
        success: false,
        message: 'この機能にはシステム管理者権限が必要です'
      };
    }
    
    // SecurityService経由でログ取得
    const logs = SecurityService.getSecurityLogs({
      eventType: 'USER_DELETION',
      limit: options.limit || 50,
      startDate: options.startDate,
      endDate: options.endDate
    });
    
    return {
      success: true,
      logs,
      totalCount: logs.length
    };
  } catch (error) {
    console.error('getDeletionLogsForUI エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * ユーザーログインステータス取得
 * @returns {Object} ログイン状態情報
 */
function getLoginStatus() {
  try {
    const currentUserInfo = UserService.getCurrentUserInfo();
    
    if (!currentUserInfo) {
      return {
        isLoggedIn: false,
        message: 'ログインしていません'
      };
    }
    
    return {
      isLoggedIn: true,
      userEmail: currentUserInfo.userEmail,
      userId: currentUserInfo.userId,
      isActive: currentUserInfo.isActive,
      isSystemAdmin: UserService.isSystemAdmin()
    };
  } catch (error) {
    console.error('getLoginStatus エラー:', error.message);
    return {
      isLoggedIn: false,
      message: 'ログイン状態の確認に失敗しました',
      error: error.message
    };
  }
}

/**
 * システムドメイン情報取得
 * @returns {Object} ドメイン情報
 */
function getSystemDomainInfo() {
  try {
    const webAppUrl = getWebAppUrl();
    const domain = new URL(webAppUrl).hostname;
    
    return {
      success: true,
      webAppUrl,
      domain,
      scriptId: ScriptApp.getScriptId()
    };
  } catch (error) {
    console.error('getSystemDomainInfo エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * システムステータス取得
 * @returns {Object} システム全体の状態
 */
function getSystemStatus() {
  try {
    // 基本システム情報
    const currentUser = UserService.getCurrentUserInfo();
    const isAdmin = UserService.isSystemAdmin();
    
    // システム設定状況
    const hasSystemProps = ConfigService.hasCoreSystemProps();
    const systemSetup = ConfigService.isSystemSetup();
    
    // 統計情報
    const userCount = DB.getAllUsers().length;
    const activeUserCount = DB.getAllUsers().filter(u => u.isActive).length;
    
    return {
      success: true,
      system: {
        hasSystemProps,
        systemSetup,
        scriptId: ScriptApp.getScriptId(),
        webAppUrl: getWebAppUrl()
      },
      user: {
        isLoggedIn: !!currentUser,
        userEmail: currentUser?.userEmail,
        isSystemAdmin: isAdmin
      },
      statistics: {
        totalUsers: userCount,
        activeUsers: activeUserCount
      },
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    console.error('getSystemStatus エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * ヘッダーインデックス取得（列マッピング支援）
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} ヘッダー情報
 */
function getHeaderIndices(spreadsheetId, sheetName) {
  try {
    console.log('getHeaderIndices: ヘッダー解析開始:', { spreadsheetId, sheetName });
    
    if (!spreadsheetId || !sheetName) {
      return {
        success: false,
        message: 'スプレッドシートIDとシート名が必要です'
      };
    }
    
    // DataService経由でヘッダー情報取得
    const headerInfo = DataService.analyzeSheetStructure(spreadsheetId, sheetName);
    
    if (!headerInfo.success) {
      return headerInfo;
    }
    
    return {
      success: true,
      headers: headerInfo.headers,
      indices: headerInfo.indices,
      suggestions: headerInfo.suggestions
    };
  } catch (error) {
    console.error('getHeaderIndices エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * 公開シートデータ取得
 * @param {string} userId - ユーザーID
 * @param {Object} options - 取得オプション
 * @returns {Object} 公開データまたはエラー
 */
function getPublishedSheetData(userId, options = {}) {
  try {
    console.log('getPublishedSheetData: 公開データ取得開始:', userId);
    
    if (!userId) {
      return {
        success: false,
        message: 'ユーザーIDが指定されていません'
      };
    }
    
    // DataService経由でデータ取得
    const data = DataService.getBulkData(userId, {
      classFilter: options.classFilter,
      sortOrder: options.sortOrder || 'desc',
      includeReactions: true,
      adminMode: options.adminMode || false
    });
    
    return data;
  } catch (error) {
    console.error('getPublishedSheetData エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * ユーザー認証確認（シンプル版）
 * @returns {Object} 認証状態
 */
function confirmUserRegistration() {
  try {
    const userInfo = UserService.getCurrentUserInfo();
    
    if (!userInfo) {
      return {
        success: false,
        message: 'ユーザーが認証されていません'
      };
    }
    
    return {
      success: true,
      user: {
        userId: userInfo.userId,
        userEmail: userInfo.userEmail,
        isActive: userInfo.isActive
      }
    };
  } catch (error) {
    console.error('confirmUserRegistration エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}


/**
 * アプリケーション設定切り替え（有効/無効）
 * @param {boolean} enabled - 有効にするかどうか
 * @returns {Object} 切り替え結果
 */
function setApplicationStatusForUI(enabled) {
  try {
    console.log('setApplicationStatusForUI: アプリケーション状態変更:', enabled);
    
    const userInfo = UserService.getCurrentUserInfo();
    if (!userInfo) {
      return {
        success: false,
        message: 'ユーザー情報が見つかりません'
      };
    }
    
    // 権限チェック（所有者のみ）
    const hasPermission = SecurityService.checkUserPermission(userInfo.userId, 'OWNER');
    if (!hasPermission.hasPermission) {
      return {
        success: false,
        message: hasPermission.reason
      };
    }
    
    const config = ConfigService.getUserConfig(userInfo.userId);
    if (!config) {
      return {
        success: false,
        message: '設定が見つかりません'
      };
    }
    
    // 状態更新
    config.appPublished = Boolean(enabled);
    if (enabled) {
      config.setupStatus = 'completed';
      config.publishedAt = new Date().toISOString();
    }
    
    const result = ConfigService.saveUserConfig(userInfo.userId, config);
    
    if (result.success) {
      return {
        success: true,
        message: enabled ? 'アプリケーションを有効にしました' : 'アプリケーションを無効にしました',
        status: {
          appPublished: config.appPublished,
          setupStatus: config.setupStatus
        }
      };
    } else {
      return {
        success: false,
        message: result.error || '状態の更新に失敗しました'
      };
    }
  } catch (error) {
    console.error('setApplicationStatusForUI エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * X-Frame-Options対応のリダイレクト関数
 * セキュリティ上重要な安全なリダイレクト処理
 * 
 * @param {string} url - リダイレクト先URL
 * @returns {GoogleAppsScript.HTML.HtmlOutput} リダイレクトHTML
 */
function createRedirect(url) {
  try {
    if (!url || typeof url !== 'string') {
      throw new Error('有効なURLが必要です');
    }
    
    // URLの基本的な検証
    const urlPattern = /^https?:\/\/.+/;
    if (!urlPattern.test(url)) {
      throw new Error('HTTPまたはHTTPSのURLが必要です');
    }
    
    const script = `
      <script>
        try {
          // X-Frame-Options対応: 親フレームとサブフレームの両方に対応
          if (window.top && window.top.location && window.top !== window) {
            window.top.location.href = '${url}';
          } else {
            window.location.href = '${url}';
          }
        } catch (e) {
          // フォールバック: 通常のリダイレクト
          window.location.href = '${url}';
        }
      </script>
    `;
    
    return HtmlService.createHtmlOutput(script)
      .setTitle('Redirecting...')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('createRedirect エラー:', error.message);
    // エラー時は安全なフォールバックを返す
    return HtmlService.createHtmlOutput(`
      <h2>Redirect Error</h2>
      <p>リダイレクトに失敗しました: ${error.message}</p>
      <p><a href="${getWebAppUrl()}">ホームに戻る</a></p>
    `).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/**
 * UI用のアプリケーション状態を取得
 * AppSetupPage.html から呼び出される
 * 
 * @returns {Object} アプリケーション状態情報
 */
function getApplicationStatusForUI() {
  try {
    const userInfo = UserService.getCurrentUserInfo();
    if (!userInfo) {
      return {
        success: false,
        message: 'ユーザー情報が見つかりません'
      };
    }
    
    const config = ConfigService.getUserConfig(userInfo.userId) || {};
    const props = PropertiesService.getScriptProperties();
    
    // システム全体の状態
    const systemStatus = {
      hasServiceAccount: !!props.getProperty(PROPS_KEYS.SERVICE_ACCOUNT_CREDS),
      hasDatabaseId: !!props.getProperty(PROPS_KEYS.DATABASE_SPREADSHEET_ID),
      hasAdminEmail: !!props.getProperty(PROPS_KEYS.ADMIN_EMAIL),
      isInitialized: props.getProperty('SYSTEM_INITIALIZED') === 'true'
    };
    
    // ユーザー固有の状態
    const userStatus = {
      userId: userInfo.userId,
      userEmail: userInfo.userEmail || UserService.getCurrentEmail(),
      isOwner: true, // 現在のユーザーは自分の設定の所有者
      isSystemAdmin: UserService.isSystemAdmin(),
      
      // 設定状態
      hasSpreadsheet: !!config.spreadsheetId,
      hasSheetName: !!config.sheetName,
      setupStatus: config.setupStatus || 'pending',
      isPublished: config.appPublished || false,
      publishedAt: config.publishedAt,
      
      // 最終更新情報
      lastModified: config.lastModified || userInfo.lastModified,
      configVersion: config.version || '1.0.0'
    };
    
    // 完了率の計算
    const completionItems = [
      systemStatus.hasServiceAccount,
      systemStatus.hasDatabaseId,
      systemStatus.hasAdminEmail,
      userStatus.hasSpreadsheet,
      userStatus.hasSheetName
    ];
    const completionRate = Math.round((completionItems.filter(Boolean).length / completionItems.length) * 100);
    
    // 次のステップの提案
    const nextSteps = [];
    if (!systemStatus.hasServiceAccount) nextSteps.push('Service Account設定');
    if (!systemStatus.hasDatabaseId) nextSteps.push('データベース設定');
    if (!systemStatus.hasAdminEmail) nextSteps.push('管理者メール設定');
    if (!userStatus.hasSpreadsheet) nextSteps.push('スプレッドシート選択');
    if (!userStatus.hasSheetName) nextSteps.push('シート名設定');
    if (nextSteps.length === 0 && !userStatus.isPublished) nextSteps.push('アプリケーション公開');
    
    return {
      success: true,
      systemStatus,
      userStatus,
      completionRate,
      nextSteps,
      canPublish: completionRate >= 80,
      lastUpdated: new Date().toISOString()
    };
  } catch (error) {
    console.error('getApplicationStatusForUI エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * システム診断を実行
 * AppSetupPage.html から呼び出される
 * 
 * @returns {Object} 診断結果
 */
function testSystemDiagnosis() {
  try {
    const diagnostics = {
      timestamp: new Date().toISOString(),
      services: {},
      system: {},
      summary: { passed: 0, failed: 0, warnings: 0 }
    };
    
    // 各サービスの診断を実行
    try {
      diagnostics.services.UserService = UserService.diagnose();
      if (diagnostics.services.UserService.status === 'healthy') diagnostics.summary.passed++;
      else diagnostics.summary.failed++;
    } catch (e) {
      diagnostics.services.UserService = { status: 'error', message: e.message };
      diagnostics.summary.failed++;
    }
    
    try {
      diagnostics.services.ConfigService = ConfigService.diagnose();
      if (diagnostics.services.ConfigService.status === 'healthy') diagnostics.summary.passed++;
      else diagnostics.summary.failed++;
    } catch (e) {
      diagnostics.services.ConfigService = { status: 'error', message: e.message };
      diagnostics.summary.failed++;
    }
    
    try {
      diagnostics.services.DataService = DataService.diagnose();
      if (diagnostics.services.DataService.status === 'healthy') diagnostics.summary.passed++;
      else diagnostics.summary.failed++;
    } catch (e) {
      diagnostics.services.DataService = { status: 'error', message: e.message };
      diagnostics.summary.failed++;
    }
    
    try {
      diagnostics.services.SecurityService = SecurityService.diagnose();
      if (diagnostics.services.SecurityService.status === 'healthy') diagnostics.summary.passed++;
      else diagnostics.summary.failed++;
    } catch (e) {
      diagnostics.services.SecurityService = { status: 'error', message: e.message };
      diagnostics.summary.failed++;
    }
    
    // システムレベルの診断
    const props = PropertiesService.getScriptProperties();
    diagnostics.system = {
      hasServiceAccount: !!props.getProperty(PROPS_KEYS.SERVICE_ACCOUNT_CREDS),
      hasDatabaseId: !!props.getProperty(PROPS_KEYS.DATABASE_SPREADSHEET_ID),
      hasAdminEmail: !!props.getProperty(PROPS_KEYS.ADMIN_EMAIL),
      webAppUrl: getWebAppUrl(),
      currentUser: UserService.getCurrentEmail()
    };
    
    // 総合評価
    const overallHealth = diagnostics.summary.failed === 0 ? 'healthy' : 
                         diagnostics.summary.failed < diagnostics.summary.passed ? 'warning' : 'critical';
    
    return {
      success: true,
      health: overallHealth,
      diagnostics,
      recommendations: generateRecommendations(diagnostics)
    };
  } catch (error) {
    console.error('testSystemDiagnosis エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * 自動修復を実行
 * AppSetupPage.html から呼び出される
 * 
 * @returns {Object} 修復結果
 */
function performAutoRepair() {
  try {
    const repairResults = {
      timestamp: new Date().toISOString(),
      repairs: [],
      success: 0,
      failed: 0
    };
    
    const userInfo = UserService.getCurrentUserInfo();
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    // 1. 設定の整合性チェックと修復
    try {
      const config = ConfigService.getUserConfig(userInfo.userId) || {};
      let needsUpdate = false;
      const updates = { ...config };
      
      // setupStatus と appPublished の整合性修復
      if (config.appPublished && config.setupStatus !== 'completed') {
        updates.setupStatus = 'completed';
        needsUpdate = true;
        repairResults.repairs.push('セットアップ状態を「completed」に修復');
      }
      
      // バージョン情報の補完
      if (!config.version) {
        updates.version = '1.0.0';
        needsUpdate = true;
        repairResults.repairs.push('バージョン情報を追加');
      }
      
      // タイムスタンプの補完
      if (!config.lastModified) {
        updates.lastModified = new Date().toISOString();
        needsUpdate = true;
        repairResults.repairs.push('最終更新日時を追加');
      }
      
      if (needsUpdate) {
        const result = ConfigService.saveUserConfig(userInfo.userId, updates);
        if (result.success) {
          repairResults.success++;
        } else {
          repairResults.failed++;
          repairResults.repairs.push(`設定更新失敗: ${  result.error}`);
        }
      }
    } catch (configError) {
      repairResults.failed++;
      repairResults.repairs.push(`設定修復エラー: ${  configError.message}`);
    }
    
    // 2. キャッシュのクリア
    try {
      const cache = CacheService.getScriptCache();
      cache.removeAll(['stale_data', 'invalid_config']);
      repairResults.success++;
      repairResults.repairs.push('古いキャッシュをクリア');
    } catch (cacheError) {
      repairResults.failed++;
      repairResults.repairs.push(`キャッシュクリアエラー: ${  cacheError.message}`);
    }
    
    return {
      success: true,
      repairResults,
      message: `修復完了: ${repairResults.success}件成功, ${repairResults.failed}件失敗`
    };
  } catch (error) {
    console.error('performAutoRepair エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * システム監視を実行
 * AppSetupPage.html から呼び出される
 * 
 * @returns {Object} 監視結果
 */
function performSystemMonitoring() {
  try {
    const monitoring = {
      timestamp: new Date().toISOString(),
      metrics: {
        responseTime: Date.now(),
        memoryUsage: 'N/A', // GASでは直接取得不可
        executionTime: 0
      },
      alerts: [],
      status: 'monitoring'
    };
    
    const startTime = Date.now();
    
    // 1. 基本サービスの応答性能をチェック
    try {
      const userCheck = UserService.getCurrentEmail();
      const configCheck = ConfigService.diagnose();
      const dataCheck = DataService.diagnose();
      
      monitoring.serviceHealth = {
        userService: !!userCheck,
        configService: configCheck.status === 'healthy',
        dataService: dataCheck.status === 'healthy'
      };
    } catch (serviceError) {
      monitoring.alerts.push(`サービスヘルスチェックエラー: ${  serviceError.message}`);
    }
    
    // 2. システムリソースの監視
    try {
      const props = PropertiesService.getScriptProperties();
      const allProps = props.getProperties();
      
      monitoring.systemResources = {
        propertiesCount: Object.keys(allProps).length,
        hasRequiredProps: !!(allProps[PROPS_KEYS.ADMIN_EMAIL] && 
                            allProps[PROPS_KEYS.DATABASE_SPREADSHEET_ID]),
        cacheAvailable: !!CacheService.getScriptCache()
      };
      
      // アラートの生成
      if (monitoring.systemResources.propertiesCount > 50) {
        monitoring.alerts.push(`プロパティ数が多いです (${  monitoring.systemResources.propertiesCount  })`);
      }
      
      if (!monitoring.systemResources.hasRequiredProps) {
        monitoring.alerts.push('必須プロパティが不足です');
      }
    } catch (resourceError) {
      monitoring.alerts.push(`リソース監視エラー: ${  resourceError.message}`);
    }
    
    // 3. パフォーマンス指標の計算
    const endTime = Date.now();
    monitoring.metrics.executionTime = endTime - startTime;
    monitoring.metrics.responseTime = endTime - monitoring.metrics.responseTime;
    
    // パフォーマンスアラート
    if (monitoring.metrics.executionTime > 10000) { // 10秒以上
      monitoring.alerts.push(`実行時間が長いです (${  monitoring.metrics.executionTime  }ms)`);
    }
    
    // 総合状態の判定
    monitoring.overallStatus = monitoring.alerts.length === 0 ? 'healthy' : 
                              monitoring.alerts.length < 3 ? 'warning' : 'critical';
    
    return {
      success: true,
      monitoring,
      alertCount: monitoring.alerts.length,
      summary: `監視結果: ${monitoring.overallStatus} (アラート: ${monitoring.alerts.length}件)`
    };
  } catch (error) {
    console.error('performSystemMonitoring エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * データ整合性チェックを実行
 * AppSetupPage.html から呼び出される
 * 
 * @returns {Object} 整合性チェック結果
 */
function performDataIntegrityCheck() {
  try {
    const integrityResults = {
      timestamp: new Date().toISOString(),
      checks: [],
      issues: [],
      passed: 0,
      failed: 0
    };
    
    const userInfo = UserService.getCurrentUserInfo();
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    // 1. ユーザーデータの整合性チェック
    try {
      const config = ConfigService.getUserConfig(userInfo.userId);
      
      if (config) {
        // configJSONの構造チェック
        const requiredFields = ['spreadsheetId', 'setupStatus'];
        const missingFields = requiredFields.filter(field => !config[field]);
        
        if (missingFields.length === 0) {
          integrityResults.checks.push('設定フィールド: OK');
          integrityResults.passed++;
        } else {
          integrityResults.checks.push('設定フィールド: 不足');
          integrityResults.issues.push(`不足フィールド: ${  missingFields.join(', ')}`);
          integrityResults.failed++;
        }
        
        // スプレッドシートアクセスチェック
        if (config.spreadsheetId) {
          try {
            const sheet = SpreadsheetApp.openById(config.spreadsheetId);
            const name = sheet.getName();
            integrityResults.checks.push('スプレッドシートアクセス: OK');
            integrityResults.passed++;
          } catch (accessError) {
            integrityResults.checks.push('スプレッドシートアクセス: エラー');
            integrityResults.issues.push('スプレッドシートにアクセスできません');
            integrityResults.failed++;
          }
        }
      } else {
        integrityResults.issues.push('ユーザー設定が存在しません');
        integrityResults.failed++;
      }
    } catch (userDataError) {
      integrityResults.issues.push(`ユーザーデータチェックエラー: ${  userDataError.message}`);
      integrityResults.failed++;
    }
    
    // 2. システムデータの整合性チェック
    try {
      const props = PropertiesService.getScriptProperties();
      const requiredProps = [PROPS_KEYS.ADMIN_EMAIL, PROPS_KEYS.DATABASE_SPREADSHEET_ID];
      
      for (const prop of requiredProps) {
        const value = props.getProperty(prop);
        if (value) {
          integrityResults.checks.push(`${prop}: OK`);
          integrityResults.passed++;
        } else {
          integrityResults.checks.push(`${prop}: 不足`);
          integrityResults.issues.push(`必須プロパティが設定されていません: ${prop}`);
          integrityResults.failed++;
        }
      }
    } catch (systemDataError) {
      integrityResults.issues.push(`システムデータチェックエラー: ${  systemDataError.message}`);
      integrityResults.failed++;
    }
    
    // 総合結果
    const overallIntegrity = integrityResults.failed === 0 ? 'perfect' : 
                            integrityResults.failed < integrityResults.passed ? 'acceptable' : 'poor';
    
    return {
      success: true,
      integrity: overallIntegrity,
      results: integrityResults,
      summary: `整合性チェック: ${integrityResults.passed}件成功, ${integrityResults.failed}件失敗`
    };
  } catch (error) {
    console.error('performDataIntegrityCheck エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * 診断結果から推奨事項を生成
 * @private
 * @param {Object} diagnostics - 診断結果
 * @returns {Array} 推奨事項の配列
 */
function generateRecommendations(diagnostics) {
  const recommendations = [];
  
  if (diagnostics.summary.failed > 0) {
    recommendations.push('失敗したサービスの確認と修復を実行してください');
  }
  
  if (!diagnostics.system.hasServiceAccount) {
    recommendations.push('Service Account設定が必要です');
  }
  
  if (!diagnostics.system.hasDatabaseId) {
    recommendations.push('データベーススプレッドシートIDの設定が必要です');
  }
  
  if (!diagnostics.system.hasAdminEmail) {
    recommendations.push('管理者メールアドレスの設定が必要です');
  }
  
  if (recommendations.length === 0) {
    recommendations.push('システムは正常に動作しています');
  }
  
  return recommendations;
}
