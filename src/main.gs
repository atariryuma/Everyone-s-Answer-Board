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

/* global handleGetData, handleAddReaction, handleToggleHighlight, handleRefreshData, DatabaseOperations, DataService, SystemController, UserService, ServiceFactory, isSystemAdmin, validateConfig, createErrorResponse, createSuccessResponse, createAuthError, createUserNotFoundError, createAdminRequiredError, createExceptionResponse, hasCoreSystemProps, getUserSheetData, processReaction, columnAnalysisImpl */

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

    console.log('doGet: mode =', mode);

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
          let user = db && db.findUserByEmail ? db.findUserByEmail(email) : null;

          // Auto-create admin user if not exists - with enhanced error handling
          if (!user && db && typeof db.createUser === 'function') {
            console.log('Creating admin user in database:', email);
            try {
              const userService = ServiceFactory.getUserService();
              if (userService && typeof userService.createUser === 'function') {
                user = userService.createUser(email);
                console.log('Admin user created successfully:', user?.userId);
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

        console.log('appSetup access granted for system admin:', email);
        return HtmlService.createTemplateFromFile('AppSetupPage.html').evaluate();
      }

      case 'view': {
        // Public view page - requires userId parameter
        const {userId} = params;
        if (!userId) {
          console.warn('view mode accessed without userId parameter');
          return HtmlService.createTemplateFromFile('AccessRestricted.html').evaluate();
        }

        // Verify user exists
        try {
          const db = ServiceFactory.getDB();
          const user = db.findUserById(userId);
          if (!user) {
            console.warn('view mode: user not found:', userId);
            const errorTemplate = HtmlService.createTemplateFromFile('ErrorBoundary.html');
            errorTemplate.title = 'ユーザーが見つかりません';
            errorTemplate.message = '指定されたユーザーの回答ボードは存在しません。URLをご確認ください。';
            errorTemplate.hideLoginButton = true;
            return errorTemplate.evaluate();
          }

          // Check if user's board is published
          let config = {};
          if (user.configJson) {
            try {
              config = JSON.parse(user.configJson);
            } catch (parseError) {
              console.warn('view mode: config parse error', parseError);
            }
          }

          if (!config.appPublished) {
            console.warn('view mode: board not published for user:', userId);
            const errorTemplate = HtmlService.createTemplateFromFile('ErrorBoundary.html');
            errorTemplate.title = 'ボードは非公開です';
            errorTemplate.message = 'このユーザーの回答ボードは現在非公開設定になっています。';
            errorTemplate.hideLoginButton = true;
            return errorTemplate.evaluate();
          }

          // Board is published - serve view page
          console.log('view mode: serving published board for user:', userId);
          const template = HtmlService.createTemplateFromFile('Page.html');
          template.data = {
            userId,
            userEmail: user.userEmail || null
          };
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
        // Minimal routing policy: do not auto-redirect to login.
        // For unspecified or unknown modes, show AccessRestricted to avoid unintended login attempts.
        return HtmlService.createTemplateFromFile('AccessRestricted.html').evaluate();
      }
    }
  } catch (error) {
    console.error('doGet error:', error.message);
    const errorTemplate = HtmlService.createTemplateFromFile('ErrorBoundary.html');
    errorTemplate.title = 'システムエラー';
    errorTemplate.message = 'システムで予期しないエラーが発生しました。管理者にお問い合わせください。';
    errorTemplate.debugInfo = `Error: ${error.message}\nStack: ${error.stack || 'N/A'}`;
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
          console.log('Auto-creating admin user for getUserConfig:', userId);
          const userByEmail = db.findUserByEmail(email);
          if (userByEmail && userByEmail.userId === userId) {
            user = userByEmail;
          } else if (typeof userService.createUser === 'function') {
            user = userService.createUser(email);
          } else {
            console.warn('UserService.createUser not available for admin user creation');
          }
          console.log('Admin user auto-created:', user?.userId);
        } catch (createError) {
          console.warn('Failed to auto-create admin user:', createError.message);
          // For admin users, still provide basic access even if DB write fails
          if (userService.isSystemAdmin(email)) {
            console.log('Providing fallback access for system admin');
          }
        }
      }

      // Still no user found
      if (!user) {
        return createUserNotFoundError();
      }
    }

    const config = user.configJson ? JSON.parse(user.configJson) : {};
    return { success: true, config, userId: user.userId };
  } catch (error) {
    console.error('getUserConfig error:', error.message);
    return createExceptionResponse(error);
  }
}

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
    const res = processReaction(
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

    let config = {};
    if (user.configJson) {
      try {
        config = JSON.parse(user.configJson);
      } catch (parseError) {
        console.warn('getAppStatus: config parse error', parseError);
      }
    }

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

    // 現在の設定を取得
    let config = {};
    if (user.configJson) {
      try {
        config = JSON.parse(user.configJson);
      } catch (parseError) {
        console.warn('setAppStatus: config parse error', parseError);
      }
    }

    // ボード公開状態を更新
    config.appPublished = Boolean(isActive);
    if (isActive) {
      config.publishedAt = config.publishedAt || new Date().toISOString();
    }
    config.lastModified = new Date().toISOString();

    // データベースに保存
    const updatedUser = {
      ...user,
      configJson: JSON.stringify(config),
      lastModified: new Date().toISOString()
    };

    const result = db.updateUser(user.userId, updatedUser);
    if (!result.success) {
      return createErrorResponse('Failed to update user configuration');
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
 * Add reaction by user ID - converts userId to email for addReaction
 */

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

    let config = {};
    try {
      config = JSON.parse(targetUser.configJson || '{}');
    } catch (parseError) {
      console.warn('toggleUserBoardStatus: config parse error', parseError);
    }

    // ボード公開状態を切り替え
    config.appPublished = !config.appPublished;
    if (config.appPublished) {
      config.publishedAt = config.publishedAt || new Date().toISOString();
    }
    config.lastModified = new Date().toISOString();

    const updatedUser = {
      ...targetUser,
      configJson: JSON.stringify(config),
      lastModified: new Date().toISOString()
    };

    const result = db.updateUser(targetUserId, updatedUser);
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

    const baseUrl = ScriptApp.getService().getUrl();
    let config = {};
    if (user.configJson) {
      try {
        config = JSON.parse(user.configJson);
      } catch (parseError) {
        console.warn('getBoardInfo: config parse error', parseError);
      }
    }

    // ユーザーのconfigJsonからボード公開状態を取得
    const appPublished = Boolean(config.appPublished);

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
 * Get sheet data - API Gateway function for DataService
 * @param {string} userId - ユーザーID
 * @param {Object} options - 取得オプション
 * @returns {Object} シートデータ取得結果
 */
function getSheetData(userId, options = {}) {
  try {
    if (!userId) {
      console.warn('getSheetData: userId not provided');
      return createErrorResponse('ユーザーIDが必要です', { data: [], headers: [], sheetName: '' });
    }

    console.log('getSheetData: calling getUserSheetData', { userId, hasOptions: !!options });

    // Delegate to DataService using Zero-Dependency pattern
    const result = getUserSheetData(userId, options);

    console.log('getSheetData: DataService result', {
      success: result?.success,
      hasData: !!result?.data,
      dataLength: result?.data?.length || 0,
      hasHeaders: !!result?.headers,
      sheetName: result?.sheetName || 'undefined'
    });

    return result;
  } catch (error) {
    console.error('getSheetData error:', error.message);
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
    // Call the actual implementation in DataService.gs
    return columnAnalysisImpl(spreadsheetId, sheetName, options);
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

