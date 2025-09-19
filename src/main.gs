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

/* global ServiceFactory, createErrorResponse, createSuccessResponse, createAuthError, createUserNotFoundError, createAdminRequiredError, createExceptionResponse, hasCoreSystemProps, isSystemAdmin, getUserSheetData, analyzeColumnStructure, validateConfig, dsAddReaction, dsToggleHighlight, Auth, Data, Config, getConfigSafe, saveConfigSafe, cleanConfigFields, getQuestionText, DB, validateAccess, URL */

// ===========================================
// 🔧 Core Utility Functions
// ===========================================

/**
 * Get current user email - simple and direct
 * @returns {string|null} User email or null if not authenticated
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
 * @param {string} filename - Template filename to include
 * @returns {string} HTML content of the template
 */
function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}

// ===========================================
// 🌐 HTTP Entry Points
// ===========================================

/**
 * Handle GET requests
 * @param {Object} e - Event object containing request parameters
 * @returns {HtmlOutput} HTML response for the requested page
 */
function doGet(e) {
  try {
    const params = e ? e.parameter : {};
    const mode = params.mode || 'main';


    // Simple routing
    switch (mode) {
      case 'login':
        return HtmlService.createTemplateFromFile('LoginPage.html')
          .evaluate();

      case 'admin': {
        const email = getCurrentEmail();
        if (!email) {
          const errorTemplate = HtmlService.createTemplateFromFile('ErrorBoundary.html');
          errorTemplate.title = 'ログインが必要です';
          errorTemplate.message = '管理画面にアクセスするにはログインが必要です。';
          return errorTemplate.evaluate();
        }
        if (!ServiceFactory.getUserService().isSystemAdmin(email)) {
          return HtmlService.createTemplateFromFile('AccessRestricted.html').evaluate();
        }
        // Inject minimal userInfo for template usage
        let userInfo = null;
        try {
          const db = ServiceFactory.getDB();
          let user = Data.findUserByEmail(email);

          // Auto-create admin user if not exists - with enhanced error handling
          if (!user) {
            try {
              const userService = ServiceFactory.getUserService();
              if (userService && typeof userService.createUser === 'function') {
                user = Data.createUser(email);
              } else {
                console.warn('UserService.createUser not available, creating minimal user object');
                user = {
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
              }
            } catch (createError) {
              console.warn('Failed to create admin user via UserService:', createError.message);
              // Fallback: create minimal user object without database write
              user = {
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
            }
          }

          if (user) {
            userInfo = {
              userId: user.userId || null,
              userEmail: user.userEmail || email
            };
          } else {
            // Fallback userInfo for admin even without DB user
            userInfo = {
              userId: null,
              userEmail: email
            };
          }
        } catch (e) {
          console.warn('doGet(admin): user lookup failed', e && e.message);
          // Fallback userInfo for admin
          userInfo = {
            userId: null,
            userEmail: email
          };
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

        return HtmlService.createTemplateFromFile('AppSetupPage.html').evaluate();
      }

      case 'view': {
        // Public view page - requires userId parameter
        const {userId} = params;
        if (!userId) {
          return HtmlService.createTemplateFromFile('AccessRestricted.html').evaluate();
        }

        // Verify user exists
        try {
          const db = ServiceFactory.getDB();
          const user = Data.findUserById(userId);
          if (!user) {
            const errorTemplate = HtmlService.createTemplateFromFile('ErrorBoundary.html');
            errorTemplate.title = 'ユーザーが見つかりません';
            errorTemplate.message = '指定されたユーザーの回答ボードは存在しません。URLをご確認ください。';
            errorTemplate.hideLoginButton = true;
            return errorTemplate.evaluate();
          }

          // Check if user's board is published - 統一API使用
          const configResult = getConfigSafe(userId);
          const config = configResult.success ? configResult.config : {};

          if (!config.appPublished) {
            const errorTemplate = HtmlService.createTemplateFromFile('ErrorBoundary.html');
            errorTemplate.title = 'ボードは非公開です';
            errorTemplate.message = 'このユーザーの回答ボードは現在非公開設定になっています。';
            errorTemplate.hideLoginButton = true;
            return errorTemplate.evaluate();
          }

          // Board is published - serve view page
          const template = HtmlService.createTemplateFromFile('Page.html');
          template.userId = userId;
          template.userEmail = user.userEmail || null;

          // 問題文設定: getQuestionTextで取得してテンプレートに渡す
          const questionText = getQuestionText(config);
          template.questionText = questionText || '回答ボード';
          template.boardTitle = questionText || user.userEmail || '回答ボード';

          // Admin privilege detection using existing ServiceFactory pattern
          const currentEmail = Session.getActiveUser().getEmail();
          const userService = ServiceFactory.getUserService();
          const isSystemAdmin = userService.isSystemAdmin(currentEmail);
          const isOwnBoard = currentEmail === user.userEmail;

          // Set admin privileges for template
          const hasAdminPrivileges = isSystemAdmin || isOwnBoard;
          template.showAdminFeatures = hasAdminPrivileges;
          template.isAdminUser = hasAdminPrivileges;
          template.showHighlightToggle = hasAdminPrivileges;

          return template.evaluate();

        } catch (error) {
          console.error('view mode error:', error.message);
          const errorTemplate = HtmlService.createTemplateFromFile('ErrorBoundary.html');
          errorTemplate.title = 'アクセスエラー';
          errorTemplate.message = '回答ボードへのアクセス中にエラーが発生しました。';
          errorTemplate.hideLoginButton = true;
          return errorTemplate.evaluate();
        }
      }

      case 'main':
      default: {
        // Default landing is AccessRestricted to prevent unintended login/account creation.
        // Viewers must specify ?mode=view&userId=... and admins explicitly use ?mode=login.
        return HtmlService.createTemplateFromFile('AccessRestricted.html').evaluate();
      }
    }
  } catch (error) {
    console.error('doGet error:', {
      message: error.message,
      stack: error.stack,
      mode: e.parameter?.mode,
      userId: `${e.parameter?.userId?.substring(0, 8)  }***`
    });

    const errorTemplate = HtmlService.createTemplateFromFile('ErrorBoundary.html');
    errorTemplate.title = 'システムエラー';
    errorTemplate.message = 'システムで予期しないエラーが発生しました。管理者にお問い合わせください。';

    // Validate error object before template literal usage
    if (error && error.message) {
      errorTemplate.debugInfo = `Error: ${error.message}\nStack: ${error.stack || 'N/A'}`;
    } else {
      errorTemplate.debugInfo = 'An unknown error occurred during request processing.';
    }

    return errorTemplate.evaluate();
  }
}

/**
 * Handle POST requests
 * @param {Object} e - Event object containing POST data
 * @returns {TextOutput} JSON response with operation result
 */
function doPost(e) {
  try {
    // Parse request
    const postData = e.postData ? e.postData.contents : '{}';
    const request = JSON.parse(postData);
    const {action} = request;


    // Verify authentication
    const email = getCurrentEmail();
    if (!email) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        message: 'Authentication required'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 🎯 Zero-Dependency Architecture: Direct DataService calls
    let result;
    switch (action) {
      case 'getData':
        // 🎯 Zero-Dependency: Direct Data class call
        try {
          const user = Data.findUserByEmail(email);
          if (!user) {
            result = { success: false, message: 'ユーザーが登録されていません' };
          } else {
            result = { success: true, data: getUserSheetData(user.userId, request.options || {}) };
          }
        } catch (error) {
          result = { success: false, message: error.message };
        }
        break;
      case 'addReaction':
        // Direct DataService call for reactions
        result = dsAddReaction(request.userId || email, request.rowId, request.reactionType);
        break;
      case 'toggleHighlight':
        // Direct DataService call for highlights
        result = dsToggleHighlight(request.userId || email, request.rowId);
        break;
      case 'refreshData':
        // 🎯 Zero-Dependency: Direct Data class call
        try {
          const user = Data.findUserByEmail(email);
          if (!user) {
            result = { success: false, message: 'ユーザーが登録されていません' };
          } else {
            result = { success: true, data: getUserSheetData(user.userId, request.options || {}) };
          }
        } catch (error) {
          result = { success: false, message: error.message };
        }
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
  const startTime = new Date().toISOString();
  console.log('=== getConfig START ===', { timestamp: startTime });

  try {
    // 🎯 User authentication logging
    const email = getCurrentEmail();
    console.log('getConfig: User authentication', {
      emailFound: !!email,
      emailLength: email ? email.length : 0
    });

    if (!email) {
      console.error('getConfig: Authentication failed');
      return createAuthError();
    }

    // 🎯 Database access logging
    const db = ServiceFactory.getDB();
    console.log('getConfig: Database access', {
      dbAvailable: !!db,
      dbType: typeof db,
      findUserByEmailAvailable: !!(db && typeof db.findUserByEmail === 'function')
    });

    if (!db) {
      console.error('getConfig: Database not available');
      return createErrorResponse('データベースアクセスエラー');
    }

    // 🎯 User lookup logging
    console.log('getConfig: Starting user lookup', { email: `${email.substring(0, 5)  }***` });
    const user = db.findUserByEmail(email);
    console.log('getConfig: User lookup result', {
      userFound: !!user,
      userId: user ? user.userId : null,
      configJsonLength: user ? (user.configJson ? user.configJson.length : 0) : null,
      userKeys: user ? Object.keys(user) : null
    });

    // 🎯 Database diagnostic logging if user not found
    if (!user) {
      console.log('getConfig: Running database diagnostics');
      try {
        // Check all users in database
        const allUsers = db.getAllUsers ? db.getAllUsers({ limit: 10 }) : [];
        console.log('getConfig: Database users diagnostic', {
          totalUsers: allUsers.length,
          userEmails: allUsers.map(u => u.userEmail ? `${u.userEmail.substring(0, 5)  }***` : 'NO_EMAIL'),
          hasCurrentEmail: allUsers.some(u => u.userEmail === email)
        });

        // Check database structure
        const dbConfig = Config.database();
        if (dbConfig && dbConfig.isValid) {
          const dataAccess = Data.open(dbConfig.spreadsheetId);
          if (dataAccess.spreadsheet) {
            const batchResults = dataAccess.batchRead(['Users!A1:E10']);
            const rows = batchResults.length > 0 && batchResults[0].results.length > 0 ?
              batchResults[0].results[0].values : [];
          console.log('getConfig: Raw database content', {
            databaseId: `${dbConfig.spreadsheetId.substring(0, 10)}***`,
            rowCount: rows.length,
            headers: rows.length > 0 ? rows[0] : [],
            sampleEmails: rows.slice(1, 4).map(row => row[1] ? `${row[1].substring(0, 5)  }***` : 'NO_EMAIL')
          });
          }
        }
      } catch (diagnosticError) {
        console.error('getConfig: Database diagnostic failed', {
          error: diagnosticError.message
        });
      }
    }

    if (!user) {
      console.error('getConfig: User not found', { email: `${email.substring(0, 5)  }***` });
      return createUserNotFoundError();
    }

    // 🎯 Config parsing - 統一API使用
    const configResult = getConfigSafe(user.userId);
    const config = configResult.success ? configResult.config : {};

    console.log('getConfig: Config loaded via unified API', {
      success: configResult.success,
      configKeys: Object.keys(config),
      configSize: JSON.stringify(config).length,
      message: configResult.message
    });

    const endTime = new Date().toISOString();
    const processingTime = new Date(endTime) - new Date(startTime);

    console.log('=== getConfig SUCCESS ===', {
      startTime,
      endTime,
      processingTimeMs: processingTime,
      userId: user.userId,
      configKeys: Object.keys(config)
    });

    return { success: true, config, userId: user.userId };
  } catch (error) {
    const endTime = new Date().toISOString();
    const processingTime = new Date(endTime) - new Date(startTime);

    console.error('=== getConfig ERROR ===', {
      startTime,
      endTime,
      processingTimeMs: processingTime,
      errorMessage: error.message,
      errorStack: error.stack
    });

    return createExceptionResponse(error);
  }
}

/**
 * Get web app URL
 * @returns {string} The URL of the deployed web app
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
 * @returns {Object} Setup status with success flag and message
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
    let user = db.findUserById(userId);

    if (!user) {
      // Auto-create user if it's a system admin - with enhanced safety
      const userService = ServiceFactory.getUserService();
      if (userService && userService.isSystemAdmin(email)) {
        try {
          const userByEmail = db.findUserByEmail(email);
          if (userByEmail && userByEmail.userId === userId) {
            user = userByEmail;
          } else if (typeof userService.createUser === 'function') {
            user = userService.createUser(email);
          } else {
            console.warn('UserService.createUser not available for admin user creation');
          }
        } catch (createError) {
          console.warn('Failed to auto-create admin user:', createError.message);
        }
      }

      // Still no user found
      if (!user) {
        return createUserNotFoundError();
      }
    }

    // 統一API使用
    const configResult = getConfigSafe(user.userId);
    return {
      success: configResult.success,
      config: configResult.config,
      userId: user.userId,
      message: configResult.message
    };
  } catch (error) {
    console.error('getUserConfig error:', error.message);
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
      const testAccess = Data.open(databaseId).spreadsheet;
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

      Data.createUser(userData.userEmail);
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
    const email = getCurrentEmail();
    if (!email) {
      return createErrorResponse('Authentication required');
    }

    // ユーザー情報とconfigJsonを取得
    const db = ServiceFactory.getDB();
    const user = db.findUserByEmail(email);
    if (!user) {
      return createErrorResponse('User not found');
    }

    // 統一API使用: 構造化パース
    const configResult = getConfigSafe(user.userId);
    const config = configResult.success ? configResult.config : {};

    // ユーザーのボード公開状態を取得
    const isActive = Boolean(config.appPublished);

    return {
      status: 'success',
      success: true,
      isActive,
      appStatus: isActive ? 'active' : 'inactive',
      timestamp: new Date().toISOString(),
      adminEmail: email,
      userId: user.userId
    };
  } catch (error) {
    console.error('getAppStatus error:', error.message);
    return createErrorResponse(error.message || 'Failed to get application status');
  }
}

/**
 * Validate spreadsheet access with service account
 * Frontend API for validateAccess function in SystemController.gs
 */
function validateAccessAPI(spreadsheetId, autoAddEditor = true) {
  try {
    // Zero-dependency: Direct function call from SystemController
    return validateAccess(spreadsheetId, autoAddEditor);
  } catch (error) {
    console.error('validateAccessAPI error:', error.message);
    return {
      success: false,
      message: error.message,
      sheets: []
    };
  }
}

/**
 * Set application status - simplified name
 */
function setAppStatus(isActive) {
  try {
    const email = getCurrentEmail();
    if (!email) {
      return createAuthError();
    }

    // ユーザー情報を取得
    const db = ServiceFactory.getDB();
    const user = db.findUserByEmail(email);
    if (!user) {
      return createUserNotFoundError();
    }

    // 統一API使用: 設定取得・更新・保存
    const configResult = getConfigSafe(user.userId);
    const config = configResult.success ? configResult.config : {};

    // ボード公開状態を更新
    config.appPublished = Boolean(isActive);
    if (isActive) {
      config.publishedAt = config.publishedAt || new Date().toISOString();
    }
    config.lastModified = new Date().toISOString();
    config.lastAccessedAt = new Date().toISOString();

    // 統一API使用: 検証・サニタイズ・保存
    const saveResult = saveConfigSafe(user.userId, config);
    if (!saveResult.success) {
      return createErrorResponse(`Failed to update user configuration: ${saveResult.message}`);
    }

    return {
      success: true,
      isActive: Boolean(isActive),
      status: isActive ? 'active' : 'inactive',
      updatedBy: email,
      userId: user.userId,
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    console.error('setAppStatus error:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * Check if user is system admin - simplified name
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
 * Toggle user active status for admin - system admin only
 */
function toggleUserActiveStatus(targetUserId) {
  try {
    const email = getCurrentEmail();
    if (!email || !isSystemAdmin(email)) {
      return createAdminRequiredError();
    }

    const db = ServiceFactory.getDB();
    const targetUser = db.findUserById(targetUserId);
    if (!targetUser) {
      return createUserNotFoundError();
    }

    const updatedUser = {
      ...targetUser,
      isActive: !targetUser.isActive,
      lastModified: new Date().toISOString()
    };

    const result = db.updateUser(targetUserId, updatedUser);
    if (result.success) {
      return {
        success: true,
        message: `ユーザー状態を${updatedUser.isActive ? 'アクティブ' : '非アクティブ'}に変更しました`,
        userId: targetUserId,
        newStatus: updatedUser.isActive,
        timestamp: new Date().toISOString()
      };
    } else {
      return createErrorResponse('ユーザー状態の更新に失敗しました');
    }
  } catch (error) {
    console.error('toggleUserActiveStatus error:', error.message);
    return createExceptionResponse(error);
  }
}


/**
 * Toggle user board publication status for admin - system admin only
 */
function toggleUserBoardStatus(targetUserId) {
  try {
    const email = getCurrentEmail();
    if (!email || !isSystemAdmin(email)) {
      return createAdminRequiredError();
    }

    const db = ServiceFactory.getDB();
    const targetUser = db.findUserById(targetUserId);
    if (!targetUser) {
      return createUserNotFoundError();
    }

    // 統一API使用: 設定取得・更新・保存
    const configResult = getConfigSafe(targetUserId);
    const config = configResult.success ? configResult.config : {};

    // ボード公開状態を切り替え
    config.appPublished = !config.appPublished;
    if (config.appPublished) {
      config.publishedAt = config.publishedAt || new Date().toISOString();
    }
    config.lastModified = new Date().toISOString();
    config.lastAccessedAt = new Date().toISOString();

    // 統一API使用: 検証・サニタイズ・保存
    const saveResult = saveConfigSafe(targetUserId, config);
    if (!saveResult.success) {
      return createErrorResponse(`Failed to toggle board status: ${saveResult.message}`);
    }

    const result = saveResult; // 統一APIレスポンスをそのまま利用
    if (result.success) {
      return {
        success: true,
        message: `ボードを${config.appPublished ? '公開' : '非公開'}に変更しました`,
        userId: targetUserId,
        boardPublished: config.appPublished,
        timestamp: new Date().toISOString()
      };
    } else {
      return createErrorResponse('ボード状態の更新に失敗しました');
    }
  } catch (error) {
    console.error('toggleUserBoardStatus error:', error.message);
    return createExceptionResponse(error);
  }
}


/**
 * Clear active sheet publication (set board to unpublished)
 * @param {string} targetUserId - 対象ユーザーID（省略時は現在の管理者）
 * @returns {Object} 実行結果
 */
function clearActiveSheet(targetUserId) {
  try {
    const email = getCurrentEmail();
    if (!email || !isSystemAdmin(email)) {
      return createAdminRequiredError();
    }

    const db = ServiceFactory.getDB();
    if (!db) {
      return createErrorResponse('データベース接続エラー');
    }

    let targetUser = targetUserId ? db.findUserById(targetUserId) : null;
    if (!targetUser) {
      targetUser = db.findUserByEmail(email);
    }

    if (!targetUser) {
      return createUserNotFoundError();
    }

    // 統一API使用: 設定取得・更新・保存
    const configResult = getConfigSafe(targetUser.userId);
    const config = configResult.success ? configResult.config : {};

    const wasPublished = config.appPublished === true;
    config.appPublished = false;
    config.publishedAt = null;
    config.lastModified = new Date().toISOString();
    config.lastAccessedAt = new Date().toISOString();

    // 統一API使用: 検証・サニタイズ・保存
    const saveResult = saveConfigSafe(targetUser.userId, config);
    if (!saveResult.success) {
      return createErrorResponse(`ボード状態の更新に失敗しました: ${saveResult.message}`);
    }

    return {
      success: true,
      message: wasPublished ? 'ボードを非公開にしました' : 'ボードは既に非公開です',
      boardPublished: false,
      userId: targetUser.userId
    };
  } catch (error) {
    console.error('clearActiveSheet error:', error.message);
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
    const email = getCurrentEmail();
    if (!email || !ServiceFactory.getUserService().isSystemAdmin(email)) {
      return createAdminRequiredError();
    }
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
 * Validate header integrity for user's active sheet
 * @param {string} targetUserId - 対象ユーザーID（省略可能）
 * @returns {Object} 検証結果
 */
function validateHeaderIntegrity(targetUserId) {
  try {
    const session = ServiceFactory.getSession();
    const db = ServiceFactory.getDB();

    if (!db) {
      return createErrorResponse('データベース接続エラー');
    }

    let targetUser = targetUserId ? db.findUserById(targetUserId) : null;
    if (!targetUser && session?.email) {
      targetUser = db.findUserByEmail(session.email);
    }

    if (!targetUser) {
      return createUserNotFoundError();
    }

    // 統一API使用: 構造化パース
    const configResult = getConfigSafe(targetUser.userId);
    const config = configResult.success ? configResult.config : {};

    if (!config.spreadsheetId || !config.sheetName) {
      return {
        success: false,
        error: 'Spreadsheet configuration incomplete'
      };
    }

    const dataAccess = Data.open(config.spreadsheetId);
    const sheet = dataAccess.getSheet(config.sheetName);
    if (!sheet) {
      return {
        success: false,
        error: 'Sheet not found'
      };
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] || [];
    const normalizedHeaders = headers.map(header => String(header || '').trim());
    const emptyHeaderCount = normalizedHeaders.filter(header => header.length === 0).length;
    const valid = normalizedHeaders.length > 0 && emptyHeaderCount < normalizedHeaders.length;

    return {
      success: valid,
      valid,
      headerCount: normalizedHeaders.length,
      emptyHeaderCount,
      headers: normalizedHeaders,
      error: valid ? null : 'ヘッダー情報が不完全です'
    };
  } catch (error) {
    console.error('validateHeaderIntegrity error:', error.message);
    return createExceptionResponse(error);
  }
}


/**
 * Get board info - simplified name
 */
function getBoardInfo() {
  console.log('🔍 getBoardInfo START');

  try {
    const email = getCurrentEmail();
    if (!email) {
      console.error('❌ Authentication failed');
      return createAuthError();
    }

    const db = ServiceFactory.getDB();
    const user = db.findUserByEmail(email);
    if (!user) {
      console.error('❌ User not found:', email);
      return { success: false, message: 'User not found' };
    }

    // 統一API使用: 構造化パース
    const configResult = getConfigSafe(user.userId);
    const config = configResult.success ? configResult.config : {};

    const appPublished = Boolean(config.appPublished);
    const baseUrl = ScriptApp.getService().getUrl();

    console.log('✅ getBoardInfo SUCCESS:', {
      userId: user.userId,
      appPublished,
      hasConfig: !!user.configJson
    });

    return {
      success: true,
      isActive: appPublished,
      appPublished,
      questionText: getQuestionText(config),
      urls: {
        view: `${baseUrl}?mode=view&userId=${user.userId}`,
        admin: `${baseUrl}?mode=admin&userId=${user.userId}`
      },
      lastUpdated: config.publishedAt || config.lastModified || new Date().toISOString()
    };
  } catch (error) {
    console.error('❌ getBoardInfo ERROR:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * Get sheet data - API Gateway function for DataService
 * @param {string} userId - ユーザーID
 * @param {Object} options - 取得オプション
 * @returns {Object} シートデータ取得結果
 */
function getSheetData(userId, options = {}) {
  try {
    if (!userId) {
      console.warn('getSheetData: userId not provided');
      // Direct return format like admin panel getSheetList
      return { success: false, message: 'ユーザーIDが必要です', data: [], headers: [], sheetName: '' };
    }

    // Delegate to DataService using Zero-Dependency pattern
    const result = getUserSheetData(userId, options);

    // Return directly without wrapping - same pattern as admin panel getSheetList
    return result;
  } catch (error) {
    console.error('getSheetData error:', error.message);
    // Direct return format like admin panel getSheetList
    return { success: false, message: error.message || 'データ取得エラー', data: [], headers: [], sheetName: '' };
  }
}

/**
 * フロントエンド回答ボード用のデータ取得関数
 * Page.htmlから呼び出される
 * @param {string} classFilter - クラスフィルター (例: 'すべて', '3年A組')
 * @param {string} sortOrder - ソート順 (例: 'newest', 'oldest', 'random', 'score')
 * @returns {Object} フロントエンド期待形式のデータ
 */
function getPublishedSheetData(classFilter, sortOrder) {
  console.log('📊 getPublishedSheetData START:', {
    classFilter: classFilter || 'すべて',
    sortOrder: sortOrder || 'newest'
  });

  try {
    const email = getCurrentEmail();
    if (!email) {
      console.error('❌ Authentication failed');
      return {
        error: 'Authentication required',
        rows: [],
        sheetName: '',
        header: '認証エラー'
      };
    }

    // Zero-dependency: 直接DB操作でユーザー取得
    let db = null;
    try {
      if (typeof DB !== 'undefined' && DB) {
        db = DB;
      } else if (typeof Data !== 'undefined') {
        db = Data;
      }
    } catch (dbError) {
      console.error('❌ DB initialization error:', dbError.message);
    }

    if (!db) {
      console.error('❌ Database connection failed');
      return {
        error: 'Database connection failed',
        rows: [],
        sheetName: '',
        header: 'データベースエラー'
      };
    }

    const user = db.findUserByEmail(email);
    if (!user) {
      console.error('❌ User not found:', email);
      return {
        error: 'User not found',
        rows: [],
        sheetName: '',
        header: 'ユーザーエラー'
      };
    }

    // getUserSheetDataを呼び出し、フィルターオプションを渡す
    const options = {
      classFilter: classFilter !== 'すべて' ? classFilter : undefined,
      sortBy: sortOrder || 'newest',
      includeTimestamp: true
    };

    const result = getUserSheetData(user.userId, options);

    // フロントエンド期待形式に変換
    if (result && result.success && result.data) {
      const transformedData = {
        header: result.header || result.sheetName || '回答一覧',
        sheetName: result.sheetName || '不明',
        data: result.data.map(item => ({
          rowIndex: item.rowIndex || item.id,
          name: item.name || '',
          class: item.class || '',
          opinion: item.answer || item.opinion || '',
          reason: item.reason || '',
          reactions: item.reactions || {
            UNDERSTAND: { count: 0, reacted: false },
            LIKE: { count: 0, reacted: false },
            CURIOUS: { count: 0, reacted: false }
          },
          highlight: item.highlight || false
        })),
        rows: result.data.map(item => ({
          rowIndex: item.rowIndex || item.id,
          name: item.name || '',
          class: item.class || '',
          opinion: item.answer || item.opinion || '',
          reason: item.reason || '',
          reactions: item.reactions || {
            UNDERSTAND: { count: 0, reacted: false },
            LIKE: { count: 0, reacted: false },
            CURIOUS: { count: 0, reacted: false }
          },
          highlight: item.highlight || false
        })),
        showDetails: result.showDetails !== false,
        success: true
      };

      console.log('✅ getPublishedSheetData SUCCESS:', {
        userId: user.userId,
        dataCount: result.data.length,
        sheetName: result.sheetName
      });

      return transformedData;
    } else {
      console.error('❌ Data retrieval failed:', result?.message);
      return {
        error: result?.message || 'データ取得エラー',
        rows: [],
        sheetName: result?.sheetName || '',
        header: result?.header || '問題'
      };
    }

  } catch (error) {
    console.error('❌ getPublishedSheetData ERROR:', error.message);
    return {
      error: error.message || 'データ取得エラー',
      rows: [],
      sheetName: '',
      header: '問題'
    };
  }
}

/**
 * 🎯 API Gateway: 列分析（AI自動判定）
 * @param {string} spreadsheetId スプレッドシートID
 * @param {string} sheetName シート名
 * @returns {Object} 列分析結果
 */
function analyzeColumns(spreadsheetId, sheetName) {
  try {
    const result = analyzeColumnStructure(spreadsheetId, sheetName);
    return result;
  } catch (error) {
    console.error('analyzeColumns error:', error.message);
    console.error('analyzeColumns stack:', error.stack);
    const exceptionResult = createExceptionResponse(error);
    return exceptionResult;
  }
}


// ===========================================
// 🔧 Unified Validation Functions
// ===========================================

// ===========================================
// 🔧 Unified Data Operations
// ===========================================


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

    const dataAccess = Data.open(spreadsheetId);
    const {spreadsheet} = dataAccess;
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

    // 統一API使用: 検証・サニタイズ・保存を一元化
    const saveResult = saveConfigSafe(userId, config);
    if (!saveResult.success) {
      return {
        success: false,
        message: saveResult.message || 'Failed to save config',
        errors: saveResult.errors || []
      };
    }

    const result = saveResult; // 統一APIレスポンスをそのまま利用

    return {
      success: !!result,
      message: result ? 'Config saved successfully' : 'Failed to save config'
    };
  } catch (error) {
    console.error('saveConfig error:', error.message);
    return createExceptionResponse(error);
  }
}

// ===========================================
// 🎯 f0068fa復元機能 - Zero-Dependency Architecture準拠
// ===========================================

/**
 * フォームURL取得 - 簡素化実装（configJson直接アクセス）
 * @param {string} sheetId - 未使用（互換性のため残存）
 * @returns {Object} フォーム情報
 */
function detectFormUrl(sheetId = null) {
  try {
    const email = getCurrentEmail();
    if (!email) {
      return { success: false, message: 'Authentication required', formUrl: null };
    }

    const db = ServiceFactory.getDB();
    const user = db.findUserByEmail(email);
    if (!user) {
      return { success: false, message: 'User not found', formUrl: null };
    }

    // 統一API使用: 構造化パース
    const configResult = getConfigSafe(user.userId);
    const config = configResult.success ? configResult.config : {};

    if (!config.formUrl) {
      return { success: false, message: 'Form URL not configured', formUrl: null };
    }

    return {
      success: true,
      formUrl: config.formUrl,
      formTitle: config.formTitle || 'フォーム',
      source: 'config'
    };
  } catch (error) {
    console.error('detectFormUrl error:', error.message);
    return { success: false, message: error.message, formUrl: null };
  }
}

/**
 * ドメイン認証チェック - 簡素化実装（既存認証パターン利用）
 * @param {string} userEmail - ユーザーメール（オプション）
 * @returns {Object} 認証結果
 */
function checkDomainAuth(userEmail = null) {
  try {
    const email = userEmail || getCurrentEmail();
    return {
      success: true,
      authenticated: !!email,
      domainStatus: email ? 'authenticated' : 'not_authenticated',
      userEmail: email
    };
  } catch (error) {
    console.error('checkDomainAuth error:', error.message);
    return {
      success: false,
      authenticated: false,
      domainStatus: 'error',
      message: error.message
    };
  }
}

/**
 * 新着コンテンツ検出機能 - Zero-Dependency直接実装
 * @param {string|number} lastUpdateTime - 最終更新時刻（ISO文字列またはタイムスタンプ）
 * @returns {Object} 新着検出結果
 */
function detectNewContent(lastUpdateTime) {
  try {
    const email = getCurrentEmail();
    if (!email) {
      return {
        success: false,
        hasNewContent: false,
        message: 'Authentication required'
      };
    }

    // 最終更新時刻の正規化
    let lastUpdate;
    try {
      if (typeof lastUpdateTime === 'string') {
        lastUpdate = new Date(lastUpdateTime);
      } else if (typeof lastUpdateTime === 'number') {
        lastUpdate = new Date(lastUpdateTime);
      } else {
        lastUpdate = new Date(0); // 初回チェック
      }
    } catch (e) {
      console.warn('detectNewContent: timestamp parse error', e);
      lastUpdate = new Date(0);
    }

    // ユーザーデータ取得
    const db = ServiceFactory.getDB();
    const user = db.findUserByEmail(email);
    if (!user) {
      return {
        success: false,
        hasNewContent: false,
        message: 'User not found'
      };
    }

    // 現在のシートデータ取得
    const currentData = getUserSheetData(user.userId, { includeTimestamp: true });
    if (!currentData?.success || !currentData.data) {
      return {
        success: true,
        hasNewContent: false,
        message: 'No data available'
      };
    }

    // 新着コンテンツチェック
    let newItemsCount = 0;
    const newItems = [];

    currentData.data.forEach((item, index) => {
      // アイテムのタイムスタンプチェック
      let itemTimestamp = new Date(0);
      if (item.timestamp) {
        try {
          itemTimestamp = new Date(item.timestamp);
        } catch (e) {
          // タイムスタンプが無効な場合、行番号ベースで推定
          itemTimestamp = new Date();
        }
      }

      if (itemTimestamp > lastUpdate) {
        newItemsCount++;
        newItems.push({
          rowIndex: item.rowIndex || index + 1,
          name: item.name || '匿名',
          preview: `${(item.answer || item.opinion || '').substring(0, 50)  }...`,
          timestamp: itemTimestamp.toISOString()
        });
      }
    });

    return {
      success: true,
      hasNewContent: newItemsCount > 0,
      newItemsCount,
      newItems: newItems.slice(0, 5), // 最新5件まで
      totalItems: currentData.data.length,
      checkTimestamp: new Date().toISOString(),
      lastUpdateTime: lastUpdate.toISOString()
    };
  } catch (error) {
    console.error('detectNewContent error:', error.message);
    return {
      success: false,
      hasNewContent: false,
      message: error.message
    };
  }
}

/**
 * Connect to data source - API Gateway function for DataService
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} 接続結果
 */
function connectDataSource(spreadsheetId, sheetName) {
  try {
    const email = getCurrentEmail();
    if (!email || !ServiceFactory.getUserService().isSystemAdmin(email)) {
      return createAdminRequiredError();
    }

    // Direct DataService call using ServiceFactory pattern
    const dataService = ServiceFactory.getDataService();
    return dataService.connectToSheetInternal(spreadsheetId, sheetName);

  } catch (error) {
    console.error('connectDataSource error:', error.message);
    return createExceptionResponse(error);
  }
}


// ===========================================
// 🔧 Missing API Functions for Frontend Error Fix
// ===========================================

/**
 * Get deploy user domain info - フロントエンドエラー修正用
 * @returns {Object} ドメイン情報
 */
function getDeployUserDomainInfo() {
  console.log('🏢 getDeployUserDomainInfo START');

  try {
    const email = getCurrentEmail();
    if (!email) {
      console.error('❌ Authentication failed');
      return {
        success: false,
        message: 'Authentication required',
        domain: null,
        isValidDomain: false
      };
    }

    const domain = email.includes('@') ? email.split('@')[1] : 'unknown';
    const adminEmail = ServiceFactory.getProperties().getProperty('ADMIN_EMAIL');
    const adminDomain = adminEmail ? adminEmail.split('@')[1] : null;
    const isValidDomain = adminDomain ? domain === adminDomain : true;

    console.log('✅ getDeployUserDomainInfo SUCCESS:', {
      userDomain: domain,
      adminDomain: adminDomain || 'N/A',
      isValidDomain
    });

    return {
      success: true,
      domain,
      userEmail: email,
      userDomain: domain,
      adminDomain,
      isValidDomain,
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    console.error('❌ getDeployUserDomainInfo ERROR:', error.message);
    return {
      success: false,
      message: error.message,
      domain: null,
      isValidDomain: false
    };
  }
}

/**
 * Get active form info - フロントエンドエラー修正用
 * @returns {Object} フォーム情報
 */
function getActiveFormInfo() {
  console.log('📝 getActiveFormInfo START');

  try {
    const email = getCurrentEmail();
    if (!email) {
      console.error('❌ Authentication failed');
      return {
        success: false,
        message: 'Authentication required',
        formUrl: null,
        formTitle: null
      };
    }

    const db = ServiceFactory.getDB();
    const user = db.findUserByEmail(email);
    if (!user) {
      console.error('❌ User not found:', email);
      return {
        success: false,
        message: 'User not found',
        formUrl: null,
        formTitle: null
      };
    }

    // 統一API使用: 構造化パース
    const configResult = getConfigSafe(user.userId);
    const config = configResult.success ? configResult.config : {};

    // フォーム表示判定: URL存在性を優先、検証は補助的に利用
    const hasFormUrl = !!(config.formUrl && config.formUrl.trim());
    const isValidUrl = hasFormUrl && isValidFormUrl(config.formUrl);

    console.log('✅ getActiveFormInfo SUCCESS:', {
      userId: user.userId,
      hasFormUrl,
      isValidUrl,
      formUrl: config.formUrl || null,
      formTitle: config.formTitle || 'フォーム'
    });

    return {
      success: hasFormUrl,
      shouldShow: hasFormUrl,  // URL存在性ベースで表示判定
      formUrl: hasFormUrl ? config.formUrl : null,
      formTitle: config.formTitle || 'フォーム',
      isValidUrl,  // 検証結果も提供（デバッグ用）
      message: hasFormUrl ?
        (isValidUrl ? 'Valid form found' : 'Form URL found but validation failed') :
        'No form URL configured'
    };
  } catch (error) {
    console.error('❌ getActiveFormInfo ERROR:', error.message);
    return {
      success: false,
      shouldShow: false,
      message: error.message,
      formUrl: null,
      formTitle: null
    };
  }
}

/**
 * Check if URL is valid Google Forms URL
 * @param {string} url - URL to validate
 * @returns {boolean} True if valid Google Forms URL
 */
function isValidFormUrl(url) {
  if (!url || typeof url !== 'string') {
    return false;
  }

  try {
    const urlObj = new URL(url.trim());

    if (urlObj.protocol !== 'https:') {
      return false;
    }

    const validHosts = ['docs.google.com', 'forms.gle'];
    const isValidHost = validHosts.includes(urlObj.hostname);

    if (urlObj.hostname === 'docs.google.com') {
      // Google Forms URLの包括的サポート: /forms/ と /viewform を許可
      return urlObj.pathname.includes('/forms/') || urlObj.pathname.includes('/viewform');
    }

    return isValidHost;
  } catch {
    return false;
  }
}

/**
 * Get incremental sheet data - フロントエンドエラー修正用
 * @param {string} sheetName - シート名
 * @param {Object} options - 取得オプション
 * @returns {Object} 増分データ
 */
function getIncrementalSheetData(sheetName, options = {}) {
  try {
    const email = getCurrentEmail();
    if (!email) {
      return {
        success: false,
        message: 'Authentication required',
        data: [],
        hasNewData: false
      };
    }

    const db = ServiceFactory.getDB();
    const user = db.findUserByEmail(email);
    if (!user) {
      return {
        success: false,
        message: 'User not found',
        data: [],
        hasNewData: false
      };
    }

    // getUserSheetDataを使用してデータ取得
    const result = getUserSheetData(user.userId, {
      includeTimestamp: true,
      classFilter: options.classFilter,
      sortBy: options.sortOrder || 'newest'
    });

    if (!result?.success) {
      return {
        success: false,
        message: result?.message || 'Data retrieval failed',
        data: [],
        hasNewData: false
      };
    }

    // 増分データチェック（lastSeenCountベース）
    const lastSeenCount = options.lastSeenCount || 0;
    const currentCount = result.data ? result.data.length : 0;
    const hasNewData = currentCount > lastSeenCount;

    // 新着データのみ抽出（必要に応じて）
    let incrementalData = result.data || [];
    if (hasNewData && lastSeenCount > 0) {
      incrementalData = incrementalData.slice(0, currentCount - lastSeenCount);
    }

    return {
      success: true,
      data: incrementalData,
      hasNewData,
      totalCount: currentCount,
      lastSeenCount,
      newItemsCount: hasNewData ? currentCount - lastSeenCount : 0,
      sheetName: result.sheetName || sheetName,
      header: result.header,
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    console.error('getIncrementalSheetData error:', error.message);
    // Return consistent response format that polling can handle gracefully
    return {
      success: false,
      status: 'error',
      message: error.message || 'データ取得エラー',
      data: [],
      hasNewData: false,
      newCount: 0,
      totalCount: 0,
      timestamp: new Date().toISOString()
    };
  }
}
