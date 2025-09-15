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

/* global handleGetData, handleAddReaction, handleToggleHighlight, handleRefreshData, DatabaseOperations, DataService, SystemController, UserService, ServiceFactory, isSystemAdmin, validateConfig, createErrorResponse, createSuccessResponse, createAuthError, createUserNotFoundError, createAdminRequiredError, createExceptionResponse */

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

// isAdmin removed - use UserService.isSystemAdmin() instead

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
          .evaluate()
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'admin': {
        const email = getCurrentEmail();
        if (!email) {
          return HtmlService.createHtmlOutput('<h1>Login Required</h1>');
        }

        return HtmlService.createTemplateFromFile('AdminPanel.html')
          .evaluate()
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }

      case 'setup':
        return HtmlService.createTemplateFromFile('SetupPage.html')
          .evaluate()
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'main':
      default: {
        const userEmail = getCurrentEmail();
        if (!userEmail) {
          // Redirect to login
          const loginUrl = `${ScriptApp.getService().getUrl()  }?mode=login`;
          return HtmlService.createHtmlOutput(`
            <script>window.top.location.href = '${loginUrl}';</script>
          `);
        }

        return HtmlService.createTemplateFromFile('Page.html')
          .evaluate()
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
 * Get sheet data - direct name without intermediate wrapper
 */
function getSheetData(userId, options = {}) {
  try {
    return DataService.getSheetData(userId, options);
  } catch (error) {
    console.error('getSheetData error:', error.message);
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

    // Process reaction directly
    const spreadsheet = ServiceFactory.getSpreadsheet().openById(config.spreadsheetId);
    const sheet = spreadsheet.getSheetByName(config.sheetName);

    if (!sheet) {
      return createErrorResponse('Sheet not found');
    }

    // Find or create reaction column
    const [headers] = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
    let reactionColumn = headers.indexOf(reactionKey) + 1;

    if (reactionColumn === 0) {
      // Create new reaction column
      reactionColumn = sheet.getLastColumn() + 1;
      sheet.getRange(1, reactionColumn).setValue(reactionKey);
    }

    // Update reaction count
    const currentValue = sheet.getRange(rowIndex, reactionColumn).getValue() || 0;
    const newValue = Math.max(0, currentValue + 1);
    sheet.getRange(rowIndex, reactionColumn).setValue(newValue);

    console.info('processReactionByEmail: Reaction processed', {
      userEmail, rowIndex, reactionKey, oldValue: currentValue, newValue
    });

    return {
      success: true,
      message: 'Reaction added successfully',
      data: { oldValue: currentValue, newValue }
    };

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
    return {
      success: true,
      message: 'Login successful',
      data: {
        userId: user.userId,
        email,
        redirectUrl: `${baseUrl}?mode=admin&userId=${user.userId}`
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
 * Get application status - simplified name
 */
function getAppStatus() {
  try {
    const props = ServiceFactory.getProperties();
    const isActive = props.getProperty('APP_STATUS') !== 'inactive';
    return {
      success: true,
      isActive,
      status: isActive ? 'active' : 'inactive',
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    console.error('getAppStatus error:', error.message);
    return createExceptionResponse(error);
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
function addReactionById(userId, rowId, reactionType) {
  try {
    // Get user email from userId
    const db = ServiceFactory.getDB();
    const user = db.findUserById(userId);

    if (!user || !user.userEmail) {
      return createUserNotFoundError();
    }

    // Convert rowId to rowIndex (assuming 1-based indexing)
    const rowIndex = parseInt(rowId, 10);
    if (isNaN(rowIndex) || rowIndex < 1) {
      return createErrorResponse('Invalid row ID');
    }

    // Use the unified addReaction function
    return addReaction(user.userEmail, rowIndex, reactionType);

  } catch (error) {
    console.error('addReactionById error:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * Check if user is system admin - simplified name
 */

/**
 * Legacy alias for backwards compatibility
 */

/**
 * Get users - simplified name for admin panel
 */
function getUsers(options = {}) {
  try {
    const email = getCurrentEmail();
    if (!email || !UserService.isSystemAdmin(email)) {
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
    console.error('getUsers error:', error.message);
    return createExceptionResponse(error);
  }
}


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
 * Get logs - simplified name
 */
function getLogs(options = {}) {
  try {
    const email = getCurrentEmail();
    if (!email || !UserService.isSystemAdmin(email)) {
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
 * Legacy alias for backwards compatibility
 */

// ===========================================
// 🔧 Unified Validation Functions
// ===========================================

/**
 * Validate URL - unified function replacing multiple duplicates
 */
// validateUrl removed - use validators.gs validateUrl instead

/**
 * Validate email - unified function
 */
// validateEmail removed - use validators.gs validateEmail instead

/**
 * Validate spreadsheet ID - unified function
 */
// validateSpreadsheetId removed - use validators.gs validateSpreadsheetId instead

// validateConfig removed - use validateConfig from validators.gs instead

// Removed legacy aliases - use direct functions instead

// ===========================================
// 🔧 Unified Data Operations
// ===========================================

/**
 * Toggle highlight - simplified name for highlight operations
 */
function toggleHighlight(userId, rowId) {
  try {
    const email = getCurrentEmail();
    if (!email || !UserService.isSystemAdmin(email)) {
      return createAdminRequiredError();
    }

    // Direct implementation
    const db = ServiceFactory.getDB();
    const user = db.findUserById(userId);
    if (!user || !user.configJson) {
      return createErrorResponse('User configuration not found');
    }

    const config = JSON.parse(user.configJson);
    if (!config.spreadsheetId || !config.sheetName) {
      return createErrorResponse('Spreadsheet configuration incomplete');
    }

    const spreadsheet = ServiceFactory.getSpreadsheet().openById(config.spreadsheetId);
    const sheet = spreadsheet.getSheetByName(config.sheetName);

    if (!sheet) {
      return createErrorResponse('Sheet not found');
    }

    // Find or create highlight column
    const [headers] = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
    let highlightColumn = headers.indexOf('highlight') + 1;

    if (highlightColumn === 0) {
      highlightColumn = sheet.getLastColumn() + 1;
      sheet.getRange(1, highlightColumn).setValue('highlight');
    }

    // Toggle highlight value
    const currentValue = sheet.getRange(rowId, highlightColumn).getValue();
    const newValue = currentValue ? '' : 'highlighted';
    sheet.getRange(rowId, highlightColumn).setValue(newValue);

    return {
      success: true,
      message: 'Highlight toggled successfully',
      highlighted: !!newValue
    };

  } catch (error) {
    console.error('toggleHighlight error:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * Remove reaction - simplified name
 */
function removeReaction(userId, rowId, reactionType) {
  try {
    const db = ServiceFactory.getDB();
    const user = db.findUserById(userId);
    if (!user || !user.userEmail) {
      return { success: false, message: 'User not found' };
    }

    const rowIndex = parseInt(rowId, 10);
    if (isNaN(rowIndex) || rowIndex < 1) {
      return createErrorResponse('Invalid row ID');
    }

    // Get user config
    if (!user.configJson) {
      return createErrorResponse('User configuration not found');
    }

    const config = JSON.parse(user.configJson);
    if (!config.spreadsheetId || !config.sheetName) {
      return createErrorResponse('Spreadsheet configuration incomplete');
    }

    // Process reaction removal
    const spreadsheet = ServiceFactory.getSpreadsheet().openById(config.spreadsheetId);
    const sheet = spreadsheet.getSheetByName(config.sheetName);

    if (!sheet) {
      return createErrorResponse('Sheet not found');
    }

    // Find reaction column
    const [headers] = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
    const reactionColumn = headers.indexOf(reactionType) + 1;

    if (reactionColumn === 0) {
      return { success: false, message: 'Reaction type not found' };
    }

    // Decrease reaction count
    const currentValue = sheet.getRange(rowIndex, reactionColumn).getValue() || 0;
    const newValue = Math.max(0, currentValue - 1);
    sheet.getRange(rowIndex, reactionColumn).setValue(newValue);

    return {
      success: true,
      message: 'Reaction removed successfully',
      data: { oldValue: currentValue, newValue }
    };

  } catch (error) {
    console.error('removeReaction error:', error.message);
    return createExceptionResponse(error);
  }
}

// Legacy aliases removed - use direct function names

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
 * Analyze columns - simplified name
 */
function analyzeColumns(spreadsheetId, sheetName) {
  try {
    if (!spreadsheetId || !sheetName) {
      return createErrorResponse('Spreadsheet ID and sheet name required');
    }

    const spreadsheet = ServiceFactory.getSpreadsheet().openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      return createErrorResponse('Sheet not found');
    }

    // Get headers (first row)
    const lastColumn = sheet.getLastColumn();
    if (lastColumn === 0) {
      return { success: false, message: 'No data found' };
    }

    const [headers] = sheet.getRange(1, 1, 1, lastColumn).getValues();

    // Analyze data types in first few rows
    const sampleSize = Math.min(10, sheet.getLastRow() - 1);
    const sampleData = sampleSize > 0 ?
      sheet.getRange(2, 1, sampleSize, lastColumn).getValues() : [];

    const analysis = headers.map((header, index) => {
      const columnData = sampleData.map(row => row[index]);

      return {
        name: header,
        index: index + 1,
        sampleValues: columnData.slice(0, 3),
        hasData: columnData.some(value => value !== null && value !== ''),
        dataType: detectDataType(columnData)
      };
    });

    return {
      success: true,
      columns: analysis,
      totalRows: sheet.getLastRow(),
      totalColumns: lastColumn
    };
  } catch (error) {
    console.error('analyzeColumns error:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * Helper function to detect data type
 */
function detectDataType(values) {
  const nonEmptyValues = values.filter(v => v !== null && v !== '');
  if (nonEmptyValues.length === 0) return 'empty';

  const [firstValue] = nonEmptyValues;
  if (typeof firstValue === 'number') return 'number';
  if (firstValue instanceof Date) return 'date';
  if (typeof firstValue === 'boolean') return 'boolean';
  return 'text';
}

/**
 * Get form info - simplified name
 */
function getFormInfo(spreadsheetId, sheetName) {
  try {
    if (!spreadsheetId || !sheetName) {
      return createErrorResponse('Spreadsheet ID and sheet name required');
    }

    // For now, return basic info (can be enhanced later)
    return {
      success: true,
      formInfo: {
        hasForm: false,
        message: 'Form detection not implemented yet'
      }
    };
  } catch (error) {
    console.error('getFormInfo error:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * Create form - simplified name
 */
function createForm(userId, config) {
  try {
    if (!userId || !config) {
      return { success: false, message: 'User ID and config required' };
    }

    // For now, return success (can be enhanced later)
    return {
      success: true,
      message: 'Form creation not implemented yet',
      formUrl: null
    };
  } catch (error) {
    console.error('createForm error:', error.message);
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

// Removed legacy aliases - use saveConfig directly