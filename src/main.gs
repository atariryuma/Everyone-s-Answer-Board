/**
 * main.gs - Clean Application Entry Points
 * New Services Architecture Implementation
 * 
 * 🎯 Responsibilities:
 * - HTTP request routing (doGet/doPost)
 * - Service layer coordination  
 * - Error handling & user feedback
 */

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
 * doGet - HTTP GET request handler
 * Uses new services architecture
 */
function doGet(e) {
  try {
    // Parse request parameters
    const params = parseRequestParams(e);
    
    // Security check: authentication status
    const userEmail = UserService.getCurrentEmail();
    if (!userEmail && params.mode !== APP_CONFIG.MODES.LOGIN && params.mode !== APP_CONFIG.MODES.SETUP) {
      // Redirect to login for unauthenticated users
      return handleLoginMode(params, { reason: 'authentication_required' });
    }
    
    // Access control verification (authenticated users only)
    if (userEmail) {
      const accessResult = SecurityService.checkUserPermission(null, 'authenticated_user');
      if (!accessResult.hasPermission) {
        return renderErrorPage({
          success: false,
          message: 'アクセスが拒否されました',
          details: accessResult.reason,
          canRetry: false
        });
      }
    }
    
    // System initialization check
    if (!ConfigService.isSystemSetup()) {
      return renderSetupPage(params);
    }

    // Route based on mode
    switch (params.mode) {
      case APP_CONFIG.MODES.DEBUG:
        return handleDebugMode(params);
      
      case APP_CONFIG.MODES.ADMIN:
        return handleAdminMode(params);
      
      case APP_CONFIG.MODES.LOGIN:
        return handleLoginMode(params);
      
      case APP_CONFIG.MODES.SETUP:
        return handleSetupMode(params);
      
      default:
        return handleMainMode(params);
    }
  } catch (error) {
    // Unified error handling
    const errorResponse = ErrorHandler.createSafeResponse(error, 'doGet');
    return HtmlService.createHtmlOutput(`
      <h2>Application Error</h2>
      <p>${errorResponse.message}</p>
      <p>Error Code: ${errorResponse.errorCode}</p>
      <p><a href="${ScriptApp.getService().getUrl()}">Return to Home</a></p>
    `);
  }
}

/**
 * doPost - HTTP POST request handler
 */
function doPost(e) {
  try {
    // Security validation
    const userEmail = UserService.getCurrentEmail();
    if (!userEmail) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        message: '認証が必要です'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Parse and validate request
    const request = parsePostRequest(e);
    if (!request.action) {
      throw new Error('Action parameter is required');
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
    const errorResponse = ErrorHandler.createSafeResponse(error, 'doPost');
    return ContentService.createTextOutput(JSON.stringify(errorResponse))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Request parameter parsing
 */
function parseRequestParams(e) {
  const params = e.parameter || {};
  return {
    mode: params.mode || APP_CONFIG.MODES.MAIN,
    userId: params.userId,
    classFilter: params.classFilter,
    sortOrder: params.sortOrder || 'desc',
    adminMode: params.adminMode === 'true'
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
    const errorResponse = ErrorHandler.createSafeResponse(error, 'getData');
    return ContentService.createTextOutput(JSON.stringify(errorResponse))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Reaction handler
 */
function handleAddReaction(request) {
  try {
    const userInfo = UserService.getCurrentUserInfo();
    if (!userInfo) {
      throw new Error('User authentication required');
    }

    const result = DataService.addReaction(
      userInfo.userId,
      request.rowId,
      request.reactionType
    );

    return ContentService.createTextOutput(JSON.stringify({
      success: result,
      message: result ? 'Reaction added successfully' : 'Failed to add reaction'
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    const errorResponse = ErrorHandler.createSafeResponse(error, 'addReaction');
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
    const errorResponse = ErrorHandler.createSafeResponse(error, 'toggleHighlight');
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
    const errorResponse = ErrorHandler.createSafeResponse(error, 'refreshData');
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
  `).setTitle(`Error - ${  APP_CONFIG.APP_NAME}`);
}

function renderSetupPage(params) {
  return HtmlService.createHtmlOutput(`
    <h2>System Setup Required</h2>
    <p>The system needs to be configured before use.</p>
    <p><a href="?mode=setup">Start Setup</a></p>
  `).setTitle(`Setup - ${  APP_CONFIG.APP_NAME}`);
}

function handleLoginMode(params, context = {}) {
  return HtmlService.createHtmlOutput(`
    <h2>Login Required</h2>
    <p>Please sign in to access ${APP_CONFIG.APP_NAME}.</p>
    <p>Reason: ${context.reason || 'Authentication required'}</p>
    <p><a href="${ScriptApp.getService().getUrl()}">Try Again</a></p>
  `).setTitle(`Login - ${  APP_CONFIG.APP_NAME}`);
}

function handleAdminMode(params) {
  // Placeholder for admin functionality
  return HtmlService.createHtmlOutput(`
    <h2>Admin Panel</h2>
    <p>Admin functionality coming soon...</p>
    <p><a href="${ScriptApp.getService().getUrl()}">Return to Main</a></p>
  `).setTitle(`Admin - ${  APP_CONFIG.APP_NAME}`);
}

function handleSetupMode(params) {
  // Placeholder for setup functionality
  return HtmlService.createHtmlOutput(`
    <h2>System Setup</h2>
    <p>Setup functionality coming soon...</p>
    <p><a href="${ScriptApp.getService().getUrl()}">Return to Main</a></p>
  `).setTitle(`Setup - ${  APP_CONFIG.APP_NAME}`);
}

function handleDebugMode(params) {
  try {
    const diagnostics = {
      UserService: UserService.diagnose(),
      ConfigService: ConfigService.diagnose(),
      DataService: DataService.diagnose(),
      SecurityService: SecurityService.diagnose()
    };

    return HtmlService.createHtmlOutput(`
      <h2>System Diagnostics</h2>
      <pre>${JSON.stringify(diagnostics, null, 2)}</pre>
      <p><a href="${ScriptApp.getService().getUrl()}">Return to Main</a></p>
    `).setTitle(`Debug - ${  APP_CONFIG.APP_NAME}`);
  } catch (error) {
    return renderErrorPage({ message: error.message });
  }
}