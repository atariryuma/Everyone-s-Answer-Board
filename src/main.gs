/**
 * main.gs - Simplified Application Entry Points
 *
 * 🎯 Responsibilities:
 * - HTTP request routing (doGet/doPost)
 * - Simple mode validation
 * - Template serving
 *
 * Following Google Apps Script Best Practices:
 * - Direct API calls (no abstraction layers)
 * - Minimal service calls
 * - Simple, readable code
 */

/* global handleGetData, handleAddReaction, handleToggleHighlight, handleRefreshData, DatabaseOperations, DataService, SystemController, UserService, ServiceFactory, isSystemAdmin, validateConfig, createErrorResponse, createSuccessResponse, createAuthError, createUserNotFoundError, createAdminRequiredError, createExceptionResponse, hasCoreSystemProps */

// ===========================================
// 🔧 Core Utility Functions
// ===========================================

/**
 * Get current user email - simple and direct
 */
function getCurrentEmail() {
  try {
    const session = ServiceFactory.getSession();
    return session.email;
  } catch (error) {
    console.error('getCurrentEmail error:', error.message);
    return null;
  }
}


/**
 * Include HTML template
 */
function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}

// ===========================================
// 🌐 HTTP Entry Points
// ===========================================

/**
 * Handle GET requests
 */
function doGet(e) {
  try {
    const params = e ? e.parameter : {};
    const mode = params.mode || 'main';

    console.log('doGet: mode =', mode);

    // Simple routing
    switch (mode) {
      case 'login':
        return HtmlService.createTemplateFromFile('LoginPage.html')
          .evaluate();

      case 'admin': {
        const email = getCurrentEmail();
        if (!email) {
          return HtmlService.createHtmlOutput('<h1>Login Required</h1>');
        }
        if (!ServiceFactory.getUserService().isSystemAdmin(email)) {
          return HtmlService.createTemplateFromFile('AccessRestricted.html').evaluate();
        }
        // Inject minimal userInfo for template usage
        let userInfo = null;
        try {
          const db = ServiceFactory.getDB();
          const user = db && db.findUserByEmail ? db.findUserByEmail(email) : null;
          if (user) {
            userInfo = {
              userId: user.userId || null,
              userEmail: user.userEmail || email
            };
          }
        } catch (e) {
          console.warn('doGet(admin): user lookup failed', e && e.message);
        }
        const adminTmpl = HtmlService.createTemplateFromFile('AdminPanel.html');
        adminTmpl.userInfo = userInfo;
        return adminTmpl.evaluate();
      }

      case 'setup': {
        // Only allow initial setup when core properties are NOT configured (no DB, no SA creds, no admin email)
        let showSetup = false;
        try {
          if (typeof hasCoreSystemProps === 'function') {
            showSetup = !hasCoreSystemProps();
          } else {
            const props = ServiceFactory.getProperties();
            const hasAdmin = !!props.getProperty('ADMIN_EMAIL');
            const hasDb = !!props.getProperty('DATABASE_SPREADSHEET_ID');
            const hasCreds = !!props.getProperty('SERVICE_ACCOUNT_CREDS');
            showSetup = !(hasAdmin && hasDb && hasCreds);
          }
        } catch (e) {
          // Conservative: if check fails, assume setup allowed
          showSetup = true;
        }

        return showSetup
          ? HtmlService.createTemplateFromFile('SetupPage.html').evaluate()
          : HtmlService.createTemplateFromFile('AccessRestricted.html').evaluate();
      }

      case 'appSetup': {
        // System admin only setup page
        const email = getCurrentEmail();
        if (!email) {
          return HtmlService.createTemplateFromFile('LoginPage.html').evaluate();
        }

        // Check if user is system admin
        const userService = ServiceFactory.getUserService();
        if (!userService.isSystemAdmin(email)) {
          console.warn('appSetup access denied:', email);
          return HtmlService.createTemplateFromFile('AccessRestricted.html').evaluate();
        }

        console.log('appSetup access granted for system admin:', email);
        return HtmlService.createTemplateFromFile('AppSetupPage.html').evaluate();
      }

      case 'main':
      default: {
        // Minimal routing policy: do not auto-redirect to login.
        // For unspecified or unknown modes, show AccessRestricted to avoid unintended login attempts.
        return HtmlService.createTemplateFromFile('AccessRestricted.html').evaluate();
      }
    }
  } catch (error) {
    console.error('doGet error:', error.message);
    return HtmlService.createHtmlOutput(`
      <h1>Error</h1>
      <p>${error.message}</p>
    `);
  }
}

/**
 * Handle POST requests
 */
function doPost(e) {
  try {
    // Parse request
    const postData = e.postData ? e.postData.contents : '{}';
    const request = JSON.parse(postData);
    const {action} = request;

    console.log('doPost: action =', action);

    // Verify authentication
    const email = getCurrentEmail();
    if (!email) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        message: 'Authentication required'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Route to handlers (these functions are defined in DataController.gs)
    let result;
    switch (action) {
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
        result = { success: false, message: `Unknown action: ${  action}` };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error('doPost error:', error.message);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ===========================================
// 🔧 API Functions (called from HTML)
// ===========================================

/**
 * Get current user email for HTML templates
 * Direct call - GAS Best Practice
 */

/**
 * Get user information with specified type
 * @param {string} infoType - Type of info: 'email', 'full'
 * @returns {Object} User information
 */
function getUser(infoType = 'email') {
  try {
    const email = getCurrentEmail();
    if (!email) {
      return {
        success: false,
        message: 'Not authenticated'
      };
    }

    if (infoType === 'email') {
      return {
        success: true,
        email
      };
    }

    if (infoType === 'full') {
      // Get user from database if available
      const db = ServiceFactory.getDB();
      const user = db.findUserByEmail(email);

      return {
        success: true,
        email,
        userId: user ? user.userId : null,
        userInfo: user || { email }
      };
    }

    return {
      success: false,
      message: `Unknown info type: ${infoType}`
    };
  } catch (error) {
    console.error('getUser error:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * Get user configuration - unified function for current user
 */
function getConfig() {
  try {
    const email = getCurrentEmail();
    if (!email) return createAuthError();

    // Get user from database by email
    const db = ServiceFactory.getDB();
    const user = db.findUserByEmail(email);
    if (!user) {
      return createUserNotFoundError();
    }

    const config = user.configJson ? JSON.parse(user.configJson) : {};
    return { success: true, config, userId: user.userId };
  } catch (error) {
    console.error('getConfig error:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * Get web app URL
 */
function getWebAppUrl() {
  try {
    return ScriptApp.getService().getUrl();
  } catch (error) {
    console.error('getWebAppUrl error:', error.message);
    return '';
  }
}

/**
 * Test system setup
 */
function testSetup() {
  try {
    const props = ServiceFactory.getProperties();
    const dbId = props.getProperty('DATABASE_SPREADSHEET_ID');
    const adminEmail = props.getProperty('ADMIN_EMAIL');

    return {
      success: !!(dbId && adminEmail),
      message: dbId && adminEmail ? 'Setup complete' : 'Setup incomplete'
    };
  } catch (error) {
    console.error('testSetup error:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * Reset authentication (clear cache)
 */
function resetAuth() {
  try {
    ServiceFactory.getCache().removeAll();
    return { success: true, message: 'Authentication reset' };
  } catch (error) {
    console.error('resetAuth error:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * Get user configuration by userId - for compatibility with existing HTML calls
 */
function getUserConfig(userId) {
  try {
    const email = getCurrentEmail();
    if (!email) return createAuthError();

    // If no userId provided, get current user's config
    if (!userId) {
      return getConfig();
    }

    const db = ServiceFactory.getDB();
    const user = db.findUserById(userId);

    if (!user) {
      return createUserNotFoundError();
    }

    const config = user.configJson ? JSON.parse(user.configJson) : {};
    return { success: true, config, userId: user.userId };
  } catch (error) {
    console.error('getUserConfig error:', error.message);
    return createExceptionResponse(error);
  }
}


/**
 * Legacy alias for backwards compatibility
 */

/**
 * Add reaction - simplified name, direct email implementation
 */
function addReaction(userEmail, rowIndex, reactionKey) {
  try {
    // Get user and validate
    const user = typeof DatabaseOperations !== 'undefined' ?
      ServiceFactory.getDB().findUserByEmail(userEmail) : null;

    if (!user || !user.configJson) {
      return createErrorResponse('User configuration not found');
    }

    const config = JSON.parse(user.configJson);
    if (!config.spreadsheetId || !config.sheetName) {
      return createErrorResponse('Spreadsheet configuration incomplete');
    }

    // Validate reaction parameters
    if (!reactionKey || typeof rowIndex !== 'number' || rowIndex < 1) {
      return createErrorResponse('Invalid reaction parameters');
    }

    // Delegate to DataService unified implementation
    const res = DataService.processReaction(
      config.spreadsheetId,
      config.sheetName,
      rowIndex,
      reactionKey,
      userEmail
    );

    const ok = res && res.status === 'success';
    return ok
      ? { success: true, message: res.message || 'Reaction added successfully', data: { newValue: res.newValue } }
      : { success: false, message: res?.message || 'Failed to add reaction' };

  } catch (error) {
    console.error('processReactionByEmail error:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * Setup application - unified implementation from SystemController
 */
function setupApplication(serviceAccountJson, databaseId, adminEmail, googleClientId) {
  try {
    // Validation
    if (!serviceAccountJson || !databaseId || !adminEmail) {
      return {
        success: false,
        message: '必須パラメータが不足しています'
      };
    }

    // System properties setup
    const props = ServiceFactory.getProperties();
    props.setProperty('DATABASE_SPREADSHEET_ID', databaseId);
    props.setProperty('ADMIN_EMAIL', adminEmail);
    props.setProperty('SERVICE_ACCOUNT_CREDS', serviceAccountJson);

    if (googleClientId) {
      props.setProperty('GOOGLE_CLIENT_ID', googleClientId);
    }

    // Initialize database if needed
    try {
      const testAccess = ServiceFactory.getSpreadsheet().openById(databaseId);
      console.log('Database access confirmed:', testAccess.getName());
    } catch (dbError) {
      console.warn('Database access test failed:', dbError.message);
    }

    return {
      success: true,
      message: 'Application setup completed successfully',
      data: {
        databaseId,
        adminEmail,
        googleClientId: googleClientId || null
      }
    };
  } catch (error) {
    console.error('setupApplication error:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * Process login action - handles user login flow
 */
function processLoginAction() {
  try {
    const email = getCurrentEmail();
    if (!email) {
      return {
        success: false,
        message: 'Authentication failed - no user email available'
      };
    }

    // Create or get user
    const db = ServiceFactory.getDB();
    let user = db.findUserByEmail(email);
    if (!user) {
      // Create new user
      const userData = {
        userId: Utilities.getUuid(),
        userEmail: email,
        isActive: true,
        configJson: JSON.stringify({
          setupStatus: 'pending',
          appPublished: false,
          createdAt: new Date().toISOString()
        }),
        lastModified: new Date().toISOString()
      };

      db.createUser(userData);
      user = userData;
    }

    const baseUrl = ScriptApp.getService().getUrl();
    const redirectUrl = `${baseUrl}?mode=admin&userId=${user.userId}`;
    // Return redirect URL at top-level for client compatibility
    return {
      success: true,
      message: 'Login successful',
      redirectUrl,
      adminUrl: redirectUrl,
      url: redirectUrl,
      data: {
        userId: user.userId,
        email,
        redirectUrl
      }
    };
  } catch (error) {
    console.error('processLoginAction error:', error.message);
    return {
      success: false,
      message: `Login failed: ${error.message}`
    };
  }
}

/**
 * Get system domain information - unified implementation
 */
function getSystemDomainInfo() {
  try {
    const currentUser = getCurrentEmail();
    if (!currentUser) {
      return {
        success: false,
        message: 'Session information not available'
      };
    }

    let domain = 'unknown';
    if (currentUser && currentUser.includes('@')) {
      [, domain] = currentUser.split('@');
    }

    const adminEmail = ServiceFactory.getProperties().getProperty('ADMIN_EMAIL');
    const adminDomain = adminEmail ? adminEmail.split('@')[1] : null;

    return {
      success: true,
      domain,
      currentUser,
      userDomain: domain,
      adminDomain,
      isValidDomain: adminDomain ? domain === adminDomain : true,
      userEmail: currentUser,
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    console.error('getSystemDomainInfo error:', error.message);
    return {
      success: false,
      message: `Domain information error: ${error.message}`
    };
  }
}

/**
 * Get application status - system admin only
 */
function getAppStatus() {
  try {
    // Verify system admin access
    const email = getCurrentEmail();
    if (!email) {
      return createErrorResponse('Authentication required');
    }

    const userService = ServiceFactory.getUserService();
    if (!userService.isSystemAdmin(email)) {
      console.warn('getAppStatus access denied:', email);
      return createErrorResponse('System administrator access required');
    }

    // Get application status
    const props = ServiceFactory.getProperties();
    const isActive = props.getProperty('APP_STATUS') !== 'inactive';

    return {
      status: 'success',
      success: true,
      isActive,
      appStatus: isActive ? 'active' : 'inactive',
      timestamp: new Date().toISOString(),
      adminEmail: email
    };
  } catch (error) {
    console.error('getAppStatus error:', error.message);
    return createErrorResponse(error.message || 'Failed to get application status');
  }
}


/**
 * Set application status - simplified name
 */
function setAppStatus(isActive) {
  try {
    const email = getCurrentEmail();
    if (!email || !isSystemAdmin(email)) {
      return createAdminRequiredError();
    }

    const props = ServiceFactory.getProperties();
    props.setProperty('APP_STATUS', isActive ? 'active' : 'inactive');

    return {
      success: true,
      isActive,
      status: isActive ? 'active' : 'inactive',
      updatedBy: email,
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    console.error('setAppStatus error:', error.message);
    return createExceptionResponse(error);
  }
}


/**
 * Add reaction by user ID - converts userId to email for addReaction
 */

/**
 * Check if user is system admin - simplified name
 */

/**
 * Legacy alias for backwards compatibility
 */

/**
 * Get users - simplified name for admin panel
 */
function getAdminUsers(options = {}) {
  try {
    const email = getCurrentEmail();
    if (!email || !ServiceFactory.getUserService().isSystemAdmin(email)) {
      return createAdminRequiredError();
    }

    // Get all users from database
    const db = ServiceFactory.getDB();
    const users = db.getAllUsers();
    return {
      success: true,
      users: users || []
    };
  } catch (error) {
    console.error('getAdminUsers error:', error.message);
    return createExceptionResponse(error);
  }
}

function getUsers(options = {}) { return getAdminUsers(options); }


/**
 * Delete user - simplified name
 */
function deleteUser(userId, reason = '') {
  try {
    const email = getCurrentEmail();
    if (!email) {
      return {
        success: false,
        message: 'ユーザー認証が必要です'
      };
    }

    // System admin check with ServiceFactory
    if (!isSystemAdmin(email)) {
      return {
        success: false,
        message: '管理者権限が必要です'
      };
    }

    if (!userId) {
      return {
        success: false,
        message: 'ユーザーIDが指定されていません'
      };
    }

    // Get user info before deletion
    const db = ServiceFactory.getDB();
    const targetUser = db.findUserById(userId);
    if (!targetUser) {
      return {
        success: false,
        message: 'ユーザーが見つかりません'
      };
    }

    // Execute deletion
    const result = db.deleteUser(userId);

    if (result.success) {
      console.log('管理者によるユーザー削除:', {
        admin: email,
        deletedUser: targetUser.userEmail,
        deletedUserId: userId,
        reason: reason || 'No reason provided',
        timestamp: new Date().toISOString()
      });

      return {
        success: true,
        message: 'ユーザーアカウントを削除しました',
        deletedUser: {
          userId,
          email: targetUser.userEmail
        }
      };
    } else {
      return {
        success: false,
        message: result.message || '削除に失敗しました'
      };
    }
  } catch (error) {
    console.error('deleteUser error:', error.message);
    return createExceptionResponse(error);
  }
}


/**
 * Check if current user is admin - simplified name
 */
function isAdmin() {
  try {
    const email = getCurrentEmail();
    if (!email) {
      return false;
    }
    return isSystemAdmin(email);
  } catch (error) {
    console.error('isAdmin error:', error.message);
    return false;
  }
}


/**
 * Get logs - simplified name
 */
function getLogs(options = {}) {
  try {
    const email = getCurrentEmail();
    if (!email || !ServiceFactory.getUserService().isSystemAdmin(email)) {
      return createAdminRequiredError();
    }

    // For now, return empty logs (can be enhanced later)
    return {
      success: true,
      logs: [],
      message: 'Logs functionality available'
    };
  } catch (error) {
    console.error('getLogs error:', error.message);
    return createExceptionResponse(error);
  }
}


/**
 * Get sheets - simplified name for spreadsheet list
 */
function getSheets() {
  try {
    // Direct implementation for spreadsheet access
    const drive = DriveApp;
    const spreadsheets = drive.searchFiles('mimeType="application/vnd.google-apps.spreadsheet"');

    const sheets = [];
    while (spreadsheets.hasNext()) {
      const file = spreadsheets.next();
      sheets.push({
        id: file.getId(),
        name: file.getName(),
        url: file.getUrl()
      });
    }

    return {
      success: true,
      sheets
    };
  } catch (error) {
    console.error('getSheets error:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * Legacy alias for backwards compatibility
 */

/**
 * Get board info - simplified name
 */
function getBoardInfo() {
  try {
    const email = getCurrentEmail();
    if (!email) {
      return createAuthError();
    }

    const db = ServiceFactory.getDB();
    const user = db.findUserByEmail(email);
    if (!user) {
      return { success: false, message: 'User not found' };
    }

    const props = ServiceFactory.getProperties();
    const appStatus = props.getProperty('APP_STATUS');
    const appPublished = appStatus === 'active';

    const baseUrl = ScriptApp.getService().getUrl();
    let config = {};
    if (user.configJson) {
      try {
        config = JSON.parse(user.configJson);
      } catch (parseError) {
        console.warn('getBoardInfo: config parse error', parseError);
      }
    }

    return {
      success: true,
      isActive: appPublished,
      appPublished,
      questionText: config.questionText || config.boardTitle || 'Everyone\'s Answer Board',
      urls: {
        view: `${baseUrl}?mode=view&userId=${user.userId}`,
        admin: `${baseUrl}?mode=admin&userId=${user.userId}`
      },
      lastUpdated: config.publishedAt || config.lastModified || new Date().toISOString()
    };
  } catch (error) {
    console.error('getBoardInfo error:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * Column analysis - direct call to DataService
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @param {Object} options - 分析オプション
 * @returns {Object} 列分析結果
 */
function columnAnalysis(spreadsheetId, sheetName, options = {}) {
  try {
    return DataService.columnAnalysis(spreadsheetId, sheetName, options);
  } catch (error) {
    console.error('columnAnalysis error:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * Legacy alias for backwards compatibility - analyzeColumns
 */
function analyzeColumns(spreadsheetId, sheetName, options = {}) {
  return columnAnalysis(spreadsheetId, sheetName, options);
}

/**
 * Legacy alias for backwards compatibility
 */

// ===========================================
// 🔧 Unified Validation Functions
// ===========================================

/**
 * Validate URL - unified function replacing multiple duplicates
 */

/**
 * Validate email - unified function
 */

/**
 * Validate spreadsheet ID - unified function
 */



// ===========================================
// 🔧 Unified Data Operations
// ===========================================

/**
 * Toggle highlight - simplified name for highlight operations
 */

/**
 * Remove reaction - simplified name
 */


// ===========================================
// 🔧 Additional HTML-Called Functions
// ===========================================

/**
 * Get sheet list for spreadsheet - simplified name
 */
function getSheetList(spreadsheetId) {
  try {
    if (!spreadsheetId) {
      return createErrorResponse('Spreadsheet ID required');
    }

    const spreadsheet = ServiceFactory.getSpreadsheet().openById(spreadsheetId);
    const sheets = spreadsheet.getSheets();

    const sheetList = sheets.map(sheet => ({
      name: sheet.getName(),
      id: sheet.getSheetId(),
      rowCount: sheet.getLastRow(),
      columnCount: sheet.getLastColumn()
    }));

    return {
      success: true,
      sheets: sheetList
    };
  } catch (error) {
    console.error('getSheetList error:', error.message);
    return createExceptionResponse(error);
  }
}




/**
 * Save user config - simplified name
 */
function saveConfig(userId, config) {
  try {
    if (!userId || !config) {
      return { success: false, message: 'User ID and config required' };
    }

    // Validate config first
    const validation = validateConfig(config);
    if (!validation.isValid) {
      return {
        success: false,
        message: 'Invalid config',
        errors: validation.errors
      };
    }

    // Save to database
    const db = ServiceFactory.getDB();
    const user = db.findUserById(userId);
    if (!user) {
      return { success: false, message: 'User not found' };
    }

    const updatedUser = {
      ...user,
      configJson: JSON.stringify(config),
      lastModified: new Date().toISOString()
    };

    const result = db.updateUser(userId, updatedUser);

    return {
      success: !!result,
      message: result ? 'Config saved successfully' : 'Failed to save config'
    };
  } catch (error) {
    console.error('saveConfig error:', error.message);
    return createExceptionResponse(error);
  }
}

