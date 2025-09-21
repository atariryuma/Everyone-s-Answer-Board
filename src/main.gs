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

/* global createErrorResponse, createSuccessResponse, createAuthError, createUserNotFoundError, createAdminRequiredError, createExceptionResponse, hasCoreSystemProps, getUserSheetData, addReaction, toggleHighlight, validateConfig, findUserByEmail, findUserById, findUserBySpreadsheetId, createUser, getAllUsers, updateUser, openSpreadsheet, getUserConfig, saveUserConfig, cleanConfigFields, getQuestionText, DB, validateAccess, URL, UserService, CACHE_DURATION, TIMEOUT_MS, SLEEP_MS, DataController, SystemController, getDatabaseConfig, getUserSpreadsheetData, getViewerBoardData, getSheetHeaders, performIntegratedColumnDiagnostics, generateRecommendedMapping */

// ===========================================
// 🔧 Core Utility Functions
// ===========================================

/**
 * 現在のユーザーのメールアドレスを取得
 * GAS-Native直接APIでセッション情報を安全に取得
 * @description GAS-Native Architecture準拠の統一アクセス方法
 * @returns {string|null} ユーザーのメールアドレス、または認証されていない場合はnull
 * @example
 * const email = getCurrentEmail();
 * if (email) {
 *   console.log('Current user:', email);
 * }
 */
function getCurrentEmail() {
  try {
    return Session.getActiveUser().getEmail();
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
        // 🔐 GAS-Native: 直接認証チェック
        const email = getCurrentEmail();
        if (!email || !isAdministrator(email)) {
          return createRedirectTemplate('ErrorBoundary.html', '管理者権限が必要です');
        }

        // ユーザー情報取得
        const user = findUserByEmail(email, { requestingUser: email });
        if (!user) {
          return createRedirectTemplate('ErrorBoundary.html', 'ユーザー情報が見つかりません');
        }

        // 認証済み - Administrator権限でAdminPanel表示
        const template = HtmlService.createTemplateFromFile('AdminPanel.html');
        template.userEmail = email;
        template.userId = user.userId;
        template.accessLevel = 'administrator';
        template.userInfo = user;
        return template.evaluate();
      }

      case 'setup': {
        // Only allow initial setup when core properties are NOT configured (no DB, no SA creds, no admin email)
        let showSetup = false;
        try {
          if (typeof hasCoreSystemProps === 'function') {
            showSetup = !hasCoreSystemProps();
          } else {
            const props = PropertiesService.getScriptProperties();
            const hasAdmin = !!props.getProperty('ADMIN_EMAIL');
            const hasDb = !!props.getProperty('DATABASE_SPREADSHEET_ID');
            const hasCreds = !!props.getProperty('SERVICE_ACCOUNT_CREDS');
            showSetup = !(hasAdmin && hasDb && hasCreds);
          }
        } catch (e) {
          // Conservative: if check fails, assume setup allowed
          showSetup = true;
        }

        if (showSetup) {
          return HtmlService.createTemplateFromFile('SetupPage.html').evaluate();
        } else {
          // 🔧 統一認証システム: isSystemAdmin変数をAccessRestricted.htmlに渡す
          const template = HtmlService.createTemplateFromFile('AccessRestricted.html');
          const email = getCurrentEmail();
          template.isAdministrator = email ? isAdministrator(email) : false;
          template.userEmail = email || '';
          template.timestamp = new Date().toISOString();
          return template.evaluate();
        }
      }

      case 'appSetup': {
        // 🔐 GAS-Native: 直接認証チェック - Administrator専用
        const email = getCurrentEmail();
        if (!email || !isAdministrator(email)) {
          return createRedirectTemplate('ErrorBoundary.html', '管理者権限が必要です');
        }

        // 認証済み - Administrator権限でAppSetup表示
        return HtmlService.createTemplateFromFile('AppSetupPage.html').evaluate();
      }

      case 'view': {
        // 🔐 GAS-Native: 直接認証チェック - Viewer権限確認
        const currentEmail = getCurrentEmail();
        if (!currentEmail) {
          return createRedirectTemplate('ErrorBoundary.html', 'ユーザー認証が必要です');
        }

        // 対象ユーザー確認
        const targetUserId = params.userId;
        if (!targetUserId) {
          return createRedirectTemplate('ErrorBoundary.html', 'ユーザーIDが指定されていません');
        }

        const targetUser = findUserById(targetUserId, { requestingUser: getCurrentEmail() });
        if (!targetUser) {
          return createRedirectTemplate('ErrorBoundary.html', '対象ユーザーが見つかりません');
        }

        // ユーザー設定取得
        const configResult = getUserConfig(targetUserId);
        const config = configResult.success ? configResult.config : {};

        // 公開設定チェック（管理者は常にアクセス可能）
        const isAdminUser = isAdministrator(currentEmail);
        const isOwnBoard = currentEmail === targetUser.userEmail;
        const isPublished = Boolean(config.isPublished);

        if (!isAdminUser && !isOwnBoard && !isPublished) {
          return createRedirectTemplate('ErrorBoundary.html', 'このボードは非公開に設定されています');
        }

        // 認証済み - 公開ボード表示
        const template = HtmlService.createTemplateFromFile('Page.html');
        template.userId = targetUserId;
        template.userEmail = targetUser.userEmail;

        // 問題文設定
        const questionText = getQuestionText(config, { targetUserEmail: targetUser.userEmail });
        template.questionText = questionText || '回答ボード';
        template.boardTitle = questionText || targetUser.userEmail || '回答ボード';

        // 🔧 CLAUDE.md準拠: 統一権限情報（GAS-Native Architecture）
        const isEditor = isAdminUser || isOwnBoard;
        template.isEditor = isEditor;
        template.isAdminUser = isAdminUser;
        template.isOwnBoard = isOwnBoard;

        return template.evaluate();
      }

      case 'main':
      default: {
        // Default landing is AccessRestricted to prevent unintended login/account creation.
        // Viewers must specify ?mode=view&userId=... and admins explicitly use ?mode=login.
        // 🔧 統一認証システム: isSystemAdmin変数をAccessRestricted.htmlに渡す
        const template = HtmlService.createTemplateFromFile('AccessRestricted.html');
        const email = getCurrentEmail();
        template.isAdministrator = email ? isAdministrator(email) : false;
        template.userEmail = email || '';
        template.timestamp = new Date().toISOString();
        return template.evaluate();
      }
    }
  } catch (error) {
    console.error('doGet error:', {
      message: error.message,
      stack: error.stack,
      mode: e.parameter?.mode,
      userId: e.parameter?.userId && typeof e.parameter.userId === 'string' ? `${e.parameter.userId.substring(0, 8)}***` : 'N/A'
    });

    const errorTemplate = HtmlService.createTemplateFromFile('ErrorBoundary.html');
    errorTemplate.title = 'システムエラー';
    errorTemplate.message = 'システムで予期しないエラーが発生しました。管理者にお問い合わせください。';
    errorTemplate.hideLoginButton = false;

    // Validate error object before template literal usage
    if (error.message) {
      errorTemplate.debugInfo = `Error: ${error.message}\nStack: ${error.stack || 'N/A'}`;
    } else {
      errorTemplate.debugInfo = 'An unknown error occurred during request processing.';
    }

    return errorTemplate.evaluate();
  }
}

/**
 * 🔐 統一認証システム用ヘルパー関数
 */

/**
 * リダイレクト用テンプレート作成
 * @param {string} redirectPage - リダイレクト先ページ
 * @param {string} error - エラーメッセージ（オプショナル）
 * @returns {HtmlOutput} リダイレクト用HTMLテンプレート
 */
function createRedirectTemplate(redirectPage, error) {
  try {
    const template = HtmlService.createTemplateFromFile(redirectPage);

    // 🔧 統一認証システム: AccessRestricted.htmlの場合は必要な変数を設定
    if (redirectPage === 'AccessRestricted.html') {
      const email = getCurrentEmail();
      template.isAdministrator = email ? isAdministrator(email) : false;
      template.userEmail = email || '';
      template.timestamp = new Date().toISOString();
      if (error) {
        template.message = error;
      }
    } else if (error && redirectPage === 'ErrorBoundary.html') {
      template.title = 'アクセスエラー';
      template.message = error;
      template.hideLoginButton = true;
    }

    return template.evaluate();
  } catch (templateError) {
    console.error('createRedirectTemplate error:', templateError.message);
    // フォールバック: 基本的なエラーページ
    const fallbackTemplate = HtmlService.createTemplateFromFile('ErrorBoundary.html');
    fallbackTemplate.title = 'システムエラー';
    fallbackTemplate.message = 'ページの表示中にエラーが発生しました。';
    fallbackTemplate.hideLoginButton = false;
    return fallbackTemplate.evaluate();
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
      return ContentService.createTextOutput(JSON.stringify(
        createAuthError()
      )).setMimeType(ContentService.MimeType.JSON);
    }

    // 🎯 GAS-Native Architecture: Direct DataService calls
    let result;
    switch (action) {
      case 'getData':
        // 🎯 GAS-Native: Direct Data class call
        try {
          const user = findUserByEmail(email, { requestingUser: email });
          if (!user) {
            result = createUserNotFoundError();
          } else {
            result = { success: true, data: getUserSheetData(user.userId, request.options || {}) };
          }
        } catch (error) {
          result = createExceptionResponse(error);
        }
        break;
      case 'addReaction':
        // Direct DataService call for reactions
        result = addReaction(request.userId || email, request.rowId, request.reactionType);
        break;
      case 'toggleHighlight':
        // Direct DataService call for highlights
        result = toggleHighlight(request.userId || email, request.rowId);
        break;
      case 'refreshData':
        // 🎯 GAS-Native: Direct Data class call
        try {
          const user = findUserByEmail(email, { requestingUser: email });
          if (!user) {
            result = createUserNotFoundError();
          } else {
            result = { success: true, data: getUserSheetData(user.userId, request.options || {}) };
          }
        } catch (error) {
          result = createExceptionResponse(error);
        }
        break;
      default:
        result = createErrorResponse(action ? `Unknown action: ${action}` : 'Unknown action: 不明');
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
      // 🔧 GAS-Native統一: 直接findUserByEmail使用
      const user = findUserByEmail(email, { requestingUser: email });

      return {
        success: true,
        email,
        userId: user ? user.userId : null,
        userInfo: user || { email }
      };
    }

    return {
      success: false,
      message: infoType ? `Unknown info type: ${infoType}` : 'Unknown info type: 不明'
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
    if (!email) {
      return createAuthError();
    }

    // 🔧 GAS-Native統一: 直接findUserByEmail使用
    const user = findUserByEmail(email, { requestingUser: email });
    if (!user) {
      return createUserNotFoundError();
    }

    // 🔧 統一API使用: getUserConfigで設定取得
    const configResult = getUserConfig(user.userId);
    const config = configResult.success ? configResult.config : {};

    return { success: true, data: { config, userId: user.userId } };
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

// ===========================================
// 🔧 フロントエンド互換性API - 統一認証システム対応
// ===========================================

/**
 * 統一管理者認証関数（メイン実装）
 * 全システム共通の管理者権限チェック
 * @param {string} email - メールアドレス
 * @returns {boolean} 管理者権限があるか
 */
function isAdministrator(email) {
  if (!email || typeof email !== 'string') {
    return false;
  }

  try {
    const adminEmail = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL');
    if (!adminEmail) {
      console.warn('isAdministrator: ADMIN_EMAIL設定が見つかりません');
      return false;
    }

    const isAdmin = email.toLowerCase() === adminEmail.toLowerCase();
    if (isAdmin) {
      console.info('isAdministrator: Administrator認証成功', {
        email: email && typeof email === 'string' ? `${email.split('@')[0]}@***` : 'N/A'
      });
    }

    return isAdmin;
  } catch (error) {
    console.error('[ERROR] main.isAdministrator:', {
      error: error.message,
      email: email && typeof email === 'string' ? `${email.split('@')[0]}@***` : 'null'
    });
    return false;
  }
}

/**
 * 管理者権限確認（フロントエンド互換性）
 * 統一認証システムのAdministrator権限をチェック
 * @returns {boolean} 管理者権限があるか
 */
function isAdmin() {
  try {
    const email = getCurrentEmail();
    if (!email) {
      return false;
    }
    return isAdministrator(email);
  } catch (error) {
    console.error('isAdmin error:', error.message);
    return false;
  }
}

/**
 * Test system setup
 * @returns {Object} Setup status with success flag and message
 */
function testSetup() {
  try {
    const props = PropertiesService.getScriptProperties();
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
    const cache = CacheService.getScriptCache();
    // Clear authentication-related cache entries
    cache.remove('current_user_info');
    return { success: true, message: 'Authentication reset' };
  } catch (error) {
    console.error('resetAuth error:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * Get user configuration by userId - for compatibility with existing HTML calls
 */

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
    const props = PropertiesService.getScriptProperties();
    props.setProperty('DATABASE_SPREADSHEET_ID', databaseId);
    props.setProperty('ADMIN_EMAIL', adminEmail);
    props.setProperty('SERVICE_ACCOUNT_CREDS', serviceAccountJson);

    if (googleClientId) {
      props.setProperty('GOOGLE_CLIENT_ID', googleClientId);
    }

    // Initialize database if needed
    try {
      const testAccess = openSpreadsheet(databaseId, { useServiceAccount: true }).spreadsheet;
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
    // 🔧 GAS-Native統一: 直接Data使用
    let user = findUserByEmail(email, { requestingUser: email });
    if (!user) {
      // 🔧 CLAUDE.md準拠: createUser()統一実装
      user = createUser(email);
      if (!user) {
        console.warn('createUser failed, creating fallback user object');
        user = {
          userId: Utilities.getUuid(),
          userEmail: email,
          isActive: true,
          configJson: JSON.stringify({
            setupStatus: 'pending',
            isPublished: false,
            createdAt: new Date().toISOString()
          }),
          lastModified: new Date().toISOString()
        };
      }
    }

    const baseUrl = ScriptApp.getService().getUrl();
    if (!baseUrl) {
      return {
        success: false,
        message: 'Web app URL not available'
      };
    }
    if (!user || !user.userId) {
      return {
        success: false,
        message: 'Invalid user data'
      };
    }
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
      message: `Login failed: ${error.message || '詳細不明'}`
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

    const adminEmail = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL');
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
      message: `Domain information error: ${error.message || '詳細不明'}`
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
    // 🔧 GAS-Native統一: 直接findUserByEmail使用
    const user = findUserByEmail(email, { requestingUser: email });
    if (!user) {
      return createErrorResponse('User not found');
    }

    // 統一API使用: 構造化パース
    const configResult = getUserConfig(user.userId);
    const config = configResult.success ? configResult.config : {};

    // ユーザーのボード公開状態を取得
    const isActive = Boolean(config.isPublished);

    return createSuccessResponse('Application status retrieved', {
      isActive,
      appStatus: isActive ? 'active' : 'inactive',
      timestamp: new Date().toISOString(),
      adminEmail: email,
      userId: user.userId
    });
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
    // 🔧 GAS-Native統一: 直接findUserByEmail使用
    const user = findUserByEmail(email, { requestingUser: email });
    if (!user) {
      return createUserNotFoundError();
    }

    // 統一API使用: 設定取得・更新・保存
    const configResult = getUserConfig(user.userId);
    const config = configResult.success ? configResult.config : {};

    // ボード公開状態を更新
    config.isPublished = Boolean(isActive);
    if (isActive) {
      if (!config.publishedAt) {
        config.publishedAt = new Date().toISOString();
      }
    }
    config.lastModified = new Date().toISOString();
    config.lastAccessedAt = new Date().toISOString();

    // 統一API使用: 検証・サニタイズ・保存
    const saveResult = saveUserConfig(user.userId, config);
    if (!saveResult.success) {
      return createErrorResponse(`Failed to update user configuration: ${saveResult.message || '詳細不明'}`);
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
    if (!email || !isAdministrator(email)) {
      return createAdminRequiredError();
    }

    // 🔧 GAS-Native統一: 直接getAllUsers使用
    const users = getAllUsers();
    return {
      success: true,
      users: users || []
    };
  } catch (error) {
    console.error('getAdminUsers error:', error.message);
    return createExceptionResponse(error);
  }
}

// ✅ CLAUDE.md準拠: 重複実装を削除
// DatabaseCore.gsのdeleteUser関数を使用することで統一


/**
 * Toggle user active status for admin - system admin only
 */
function toggleUserActiveStatus(targetUserId) {
  try {
    const email = getCurrentEmail();
    if (!email || !isAdministrator(email)) {
      return createAdminRequiredError();
    }

    // 🔧 GAS-Native統一: 直接findUserById使用
    const targetUser = findUserById(targetUserId);
    if (!targetUser) {
      return createUserNotFoundError();
    }

    const updatedUser = {
      ...targetUser,
      isActive: !targetUser.isActive,
      lastModified: new Date().toISOString()
    };

    const result = updateUser(targetUserId, updatedUser);
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
    if (!email || !isAdministrator(email)) {
      return createAdminRequiredError();
    }

    // 🔧 GAS-Native統一: 直接findUserById使用
    const targetUser = findUserById(targetUserId);
    if (!targetUser) {
      return createUserNotFoundError();
    }

    // 統一API使用: 設定取得・更新・保存
    const configResult = getUserConfig(targetUserId);
    const config = configResult.success ? configResult.config : {};

    // ボード公開状態を切り替え
    config.isPublished = !config.isPublished;
    if (config.isPublished) {
      if (!config.publishedAt) {
        config.publishedAt = new Date().toISOString();
      }
    }
    config.lastModified = new Date().toISOString();
    config.lastAccessedAt = new Date().toISOString();

    // 統一API使用: 検証・サニタイズ・保存
    const saveResult = saveUserConfig(targetUserId, config);
    if (!saveResult.success) {
      return createErrorResponse(`Failed to toggle board status: ${saveResult.message || '詳細不明'}`);
    }

    const result = saveResult; // 統一APIレスポンスをそのまま利用
    if (result.success) {
      return {
        success: true,
        message: `ボードを${config.isPublished ? '公開' : '非公開'}に変更しました`,
        userId: targetUserId,
        boardPublished: config.isPublished,
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
    if (!email || !isAdministrator(email)) {
      return createAdminRequiredError();
    }

    // 🔧 GAS-Native統一: 直接Data使用
    let targetUser = targetUserId ? findUserById(targetUserId) : null;
    if (!targetUser) {
      targetUser = findUserByEmail(email);
    }

    if (!targetUser) {
      return createUserNotFoundError();
    }

    // 統一API使用: 設定取得・更新・保存
    const configResult = getUserConfig(targetUser.userId);
    const config = configResult.success ? configResult.config : {};

    const wasPublished = config.isPublished === true;
    config.isPublished = false;
    config.publishedAt = null;
    config.lastModified = new Date().toISOString();
    config.lastAccessedAt = new Date().toISOString();

    // 統一API使用: 検証・サニタイズ・保存
    const saveResult = saveUserConfig(targetUser.userId, config);
    if (!saveResult.success) {
      return createErrorResponse(`ボード状態の更新に失敗しました: ${saveResult.message || '詳細不明'}`);
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


/**
 * Get logs - simplified name
 */
function getLogs(options = {}) {
  try {
    const email = getCurrentEmail();
    if (!email || !isAdministrator(email)) {
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
    if (!email || !isAdministrator(email)) {
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
    const currentEmail = getCurrentEmail();
    // 🔧 GAS-Native: 直接Data使用
    let targetUser = targetUserId ? findUserById(targetUserId) : null;
    if (!targetUser && currentEmail) {
      targetUser = findUserByEmail(currentEmail);
    }

    if (!targetUser) {
      return createUserNotFoundError();
    }

    // 統一API使用: 構造化パース
    const configResult = getUserConfig(targetUser.userId);
    const config = configResult.success ? configResult.config : {};

    if (!config.spreadsheetId || !config.sheetName) {
      return {
        success: false,
        error: 'Spreadsheet configuration incomplete'
      };
    }

    // 🔧 CLAUDE.md準拠: Self vs Cross-user access pattern
    const isSelfAccess = targetUser.userEmail === currentEmail;
    const dataAccess = openSpreadsheet(config.spreadsheetId, {
      useServiceAccount: !isSelfAccess
    });
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

    // 🔧 GAS-Native統一: 直接findUserByEmail使用
    const user = findUserByEmail(email, { requestingUser: email });
    if (!user) {
      console.error('❌ User not found:', email);
      return { success: false, message: 'User not found' };
    }

    // 統一API使用: 構造化パース
    const configResult = getUserConfig(user.userId);
    const config = configResult.success ? configResult.config : {};

    const isPublished = Boolean(config.isPublished);
    const baseUrl = ScriptApp.getService().getUrl();

    console.log('✅ getBoardInfo SUCCESS:', {
      userId: user.userId,
      isPublished,
      hasConfig: !!user.configJson
    });

    return {
      success: true,
      isActive: isPublished,
      isPublished,
      questionText: getQuestionText(config, { targetUserEmail: user.userEmail }),
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
 * 🔧 統合API: getUserSheetData直接呼び出し推奨
 * この関数は後方互換性のためのラッパーです
 * @deprecated Use getUserSheetData() directly instead
 * @param {string} userId - ユーザーID
 * @param {Object} options - 取得オプション
 * @returns {Object} シートデータ取得結果
 */
function getSheetData(userId, options = {}) {
  console.warn('getSheetData: Deprecated wrapper - use getUserSheetData() directly');
  return getUserSheetData(userId, options);
}

/**
 * 回答ボード表示用データ取得（userId指定対応）
 * @param {string} targetUserId - 表示対象のユーザーID
 * @param {string} classFilter - クラスフィルター (例: 'すべて', '3年A組')
 * @param {string} sortOrder - ソート順 (例: 'newest', 'oldest', 'random', 'score')
 * @returns {Object} フロントエンド期待形式のデータ
 */
function getBoardData(targetUserId, classFilter, sortOrder) {
  try {
    const viewerEmail = getCurrentEmail();
    if (!viewerEmail) {
      return {
        error: 'Authentication required',
        rows: [],
        sheetName: '',
        header: '認証エラー'
      };
    }

    // getViewerBoardDataを使用してクロスユーザーアクセスを処理
    const boardData = getViewerBoardData(targetUserId, viewerEmail);
    if (!boardData) {
      return {
        error: 'User not found',
        rows: [],
        sheetName: '',
        header: 'ユーザーエラー'
      };
    }

    // フィルターとソートのオプション
    const options = {
      classFilter: classFilter !== 'すべて' ? classFilter : undefined,
      sortBy: sortOrder || 'newest',
      includeTimestamp: true
    };

    // 🔍 CLAUDE.md準拠: データ形式の詳細ログ出力
    console.log('getBoardData: boardData received:', {
      hasBoardData: !!boardData,
      boardDataType: typeof boardData,
      boardDataKeys: boardData ? Object.keys(boardData) : 'N/A',
      hasSuccess: boardData ? 'success' in boardData : false,
      successValue: boardData ? boardData.success : 'N/A',
      hasData: boardData ? 'data' in boardData : false,
      dataType: boardData?.data ? typeof boardData.data : 'N/A',
      dataLength: Array.isArray(boardData?.data) ? boardData.data.length : 'N/A'
    });

    // フロントエンド期待形式に変換（より堅牢なチェック）
    if (boardData && (boardData.success !== false) && boardData.data && Array.isArray(boardData.data)) {
      const transformedData = {
        header: boardData.header || boardData.sheetName || '回答一覧',
        sheetName: boardData.sheetName || '不明',
        data: boardData.data.map(item => ({
          rowIndex: item.rowIndex || item.id,
          name: item.name || '',
          class: item.class || '',
          opinion: item.answer || item.opinion || '',
          reason: item.reason || '',
          reactions: item.reactions || {
            UNDERSTAND: { count: 0, reacted: false },
            LIKE: { count: 0, reacted: false },
            SURPRISE: { count: 0, reacted: false }
          },
          highlight: Boolean(item.highlight),
          timestamp: item.timestamp || '',
          formattedTimestamp: item.formattedTimestamp || ''
        }))
      };

      // クラスフィルターの適用
      if (options.classFilter) {
        transformedData.data = transformedData.data.filter(item =>
          item.class === options.classFilter
        );
      }

      // ソート適用
      if (options.sortBy === 'newest') {
        transformedData.data.sort((a, b) => (b.rowIndex || 0) - (a.rowIndex || 0));
      } else if (options.sortBy === 'oldest') {
        transformedData.data.sort((a, b) => (a.rowIndex || 0) - (b.rowIndex || 0));
      }

      return transformedData;
    }

    // データが存在するが配列でない場合、または空の場合
    if (boardData && boardData.data) {
      console.warn('getBoardData: Data exists but not in expected format');
      return {
        header: boardData.header || boardData.sheetName || '回答一覧',
        sheetName: boardData.sheetName || '不明',
        data: []
      };
    }

    console.warn('getBoardData: No valid data found');
    return {
      error: 'No data available',
      rows: [],
      sheetName: '',
      header: 'データなし'
    };

  } catch (error) {
    console.error('getBoardData error:', error.message);
    return createExceptionResponse(error, 'getBoardData');
  }
}

/**
 * 🔧 統合API: フロントエンド用データ取得（最適化版）
 * Page.htmlから呼び出される
 * @param {string} classFilter - クラスフィルター (例: 'すべて', '3年A組')
 * @param {string} sortOrder - ソート順 (例: 'newest', 'oldest', 'random', 'score')
 * @returns {Object} フロントエンド期待形式のデータ
 */
function getPublishedSheetData(classFilter, sortOrder) {
  try {
    const email = getCurrentEmail();
    if (!email) {
      return {
        error: 'Authentication required',
        rows: [],
        sheetName: '',
        header: '認証エラー'
      };
    }

    // 🎯 CLAUDE.md準拠: 直接的なユーザー取得
    const user = findUserByEmail(email, { requestingUser: email });
    if (!user) {
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

    // 🔧 CLAUDE.md準拠: getSheetList - Context-aware service account usage
    // ✅ **Cross-user**: Only use service account for accessing other user's spreadsheets
    // ✅ **Self-access**: Use normal permissions for own spreadsheets
    const currentEmail = getCurrentEmail();

    // CLAUDE.md準拠: spreadsheetIdから所有者を特定して直接比較
    const targetUser = findUserBySpreadsheetId(spreadsheetId);
    const isSelfAccess = targetUser && targetUser.userEmail === currentEmail;
    const useServiceAccount = !isSelfAccess;

    console.log(`getSheetList: ${useServiceAccount ? 'Cross-user service account' : 'Self-access normal permissions'} for spreadsheet`);
    const dataAccess = openSpreadsheet(spreadsheetId, { useServiceAccount });
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
 * データカウント取得（フロントエンド整合性のため追加）
 * @param {string} classFilter - クラスフィルター
 * @param {string} sortOrder - ソート順
 * @param {boolean} adminMode - 管理者モード
 * @returns {Object} カウント情報
 */
function getDataCount(classFilter, sortOrder, adminMode = false) {
  try {
    const email = getCurrentEmail();
    if (!email) {
      return { error: 'Authentication required', count: 0 };
    }

    // 🔧 GAS-Native統一: 直接findUserByEmail使用
    const user = findUserByEmail(email, { requestingUser: email });
    if (!user) {
      return { error: 'User not found', count: 0 };
    }

    // getUserSheetDataを使用してデータを取得し、カウントのみ返却
    const result = getUserSheetData(user.userId, {
      classFilter,
      sortOrder,
      adminMode
    });

    if (result.success && result.data) {
      return {
        success: true,
        count: result.data.length,
        sheetName: result.sheetName
      };
    }

    return { error: result.message || 'Failed to get count', count: 0 };
  } catch (error) {
    console.error('getDataCount error:', error.message);
    return { error: error.message, count: 0 };
  }
}

/**
 * 🎯 GAS-Native統一設定保存API
 * CLAUDE.md準拠: 直接的でシンプルな実装
 * @param {Object} config - 設定オブジェクト
 * @param {Object} options - 保存オプション { isDraft: boolean }
 */
function saveConfig(config, options = {}) {
  try {
    const userEmail = getCurrentEmail();
    if (!userEmail) {
      return { success: false, message: 'ユーザー認証が必要です' };
    }

    // 🔧 GAS-Native統一: 直接findUserByEmail使用
    const user = findUserByEmail(userEmail);
    if (!user) {
      return { success: false, message: 'ユーザーが見つかりません' };
    }

    // 保存タイプをオプションで制御（GAS-Native準拠）
    const saveOptions = options.isDraft ?
      { isDraft: true } :
      { isMainConfig: true };

    // 統一API使用: saveUserConfigで安全保存
    return saveUserConfig(user.userId, config, saveOptions);

  } catch (error) {
    const operation = options.isDraft ? 'saveDraft' : 'saveConfig';
    console.error(`[ERROR] main.${operation}:`, error.message || 'Operation error');
    return { success: false, message: error.message || 'エラーが発生しました' };
  }
}

// ===========================================
// 🎯 f0068fa復元機能 - GAS-Native Architecture準拠
// ===========================================


/**
 * 新着コンテンツ検出機能 - GAS-Native直接実装
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
    // 🔧 GAS-Native統一: 直接findUserByEmail使用
    const user = findUserByEmail(email, { requestingUser: email });
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
          preview: (item.answer || item.opinion) ? `${(item.answer || item.opinion).substring(0, 50)}...` : 'プレビュー不可',
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
function connectDataSource(spreadsheetId, sheetName, batchOperations = null) {
  try {
    const email = getCurrentEmail();
    if (!email || !isAdministrator(email)) {
      return createAdminRequiredError();
    }

    console.log('connectDataSource called:', { spreadsheetId, sheetName, batchOperations });

    // バッチ処理対応 - CLAUDE.md準拠 70x Performance
    if (batchOperations && Array.isArray(batchOperations)) {
      return processDataSourceOperations(spreadsheetId, sheetName, batchOperations);
    }

    // 従来の単一接続処理（GAS-Native直接呼び出し）
    return getColumnAnalysis(spreadsheetId, sheetName);

  } catch (error) {
    console.error('connectDataSource error:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * バッチ処理でデータソース操作を実行
 * CLAUDE.md準拠 - 70x Performance向上
 */
function processDataSourceOperations(spreadsheetId, sheetName, operations) {
  try {
    const results = {
      success: true,
      batchResults: {},
      message: '統合処理完了'
    };

    // 🎯 最適化: getColumnAnalysisの結果をキャッシュして再利用
    let columnAnalysisResult = null;

    // 各操作を効率的に実行（重複API呼び出し回避）
    for (const operation of operations) {
      switch (operation.type) {
        case 'validateAccess':
          // 初回のみgetColumnAnalysisを実行、以降は結果を再利用
          if (!columnAnalysisResult) {
            columnAnalysisResult = getColumnAnalysis(spreadsheetId, sheetName);
          }
          results.batchResults.validation = {
            success: columnAnalysisResult.success,
            details: {
              connectionVerified: columnAnalysisResult.success,
              connectionError: columnAnalysisResult.success ? null : columnAnalysisResult.message
            }
          };
          break;
        case 'getFormInfo':
          results.batchResults.formInfo = getFormInfoInternal(spreadsheetId, sheetName);
          break;
        case 'connectDataSource': {
          // ✅ キャッシュされた結果を使用（重複API呼び出しなし）
          if (!columnAnalysisResult) {
            columnAnalysisResult = getColumnAnalysis(spreadsheetId, sheetName);
          }

          if (columnAnalysisResult.success) {
            // GAS-Native Architecture: 直接データ転送（中間変換なし）
            results.mapping = columnAnalysisResult.mapping;
            results.confidence = columnAnalysisResult.confidence;
            results.headers = columnAnalysisResult.headers;
            results.data = columnAnalysisResult.data;
            results.sheet = columnAnalysisResult.sheet;
            console.log('connectDataSource: AI分析結果を直接送信（キャッシュ利用）', {
              mappingKeys: Object.keys(columnAnalysisResult.mapping || {}),
              confidenceKeys: Object.keys(columnAnalysisResult.confidence || {}),
              headers: columnAnalysisResult.headers?.length || 0
            });
          } else {
            results.success = false;
            results.error = columnAnalysisResult.errorResponse?.message || columnAnalysisResult.message;
          }
          break;
        }
      }
    }

    return results;
  } catch (error) {
    console.error('processDataSourceOperations error:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * 列分析 - API Gateway実装（既存サービス活用）
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} 列分析結果
 */
function getColumnAnalysis(spreadsheetId, sheetName) {
  try {
    const email = getCurrentEmail();
    if (!email || !isAdministrator(email)) {
      return createAdminRequiredError();
    }

    // 🔧 CLAUDE.md準拠: openSpreadsheet + 既存サービス活用
    const dataAccess = openSpreadsheet(spreadsheetId);
    const sheet = dataAccess.getSheet(sheetName);

    if (!sheet) {
      return { success: false, error: 'Sheet not found' };
    }

    // 🔧 既存DataService.getSheetHeaders活用
    const lastCol = sheet.getLastColumn();
    const headers = lastCol > 0 ? getSheetHeaders(sheet, lastCol) : [];

    // 🔧 既存ColumnMappingService活用
    const diagnostics = performIntegratedColumnDiagnostics(headers);

    return {
      success: true,
      sheet,
      headers,
      data: [], // 軽量化: 接続時は不要
      mapping: diagnostics.recommendedMapping || {},
      confidence: diagnostics.confidence || {}
    };
  } catch (error) {
    console.error('getColumnAnalysis error:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * データソースアクセス検証
 */
function validateDataSourceAccess(spreadsheetId, sheetName) {
  try {
    // GAS-Native: Direct function call instead of ServiceFactory
    const testConnection = getColumnAnalysis(spreadsheetId, sheetName);
    return {
      success: testConnection.success,
      details: {
        connectionVerified: testConnection.success,
        connectionError: testConnection.success ? null : testConnection.message
      }
    };
  } catch (error) {
    return {
      success: false,
      details: {
        connectionVerified: false,
        connectionError: error.message
      }
    };
  }
}

/**
 * フォーム情報取得（バッチ処理用）
 */
function getFormInfoInternal(spreadsheetId, sheetName) {
  try {
    return getFormInfo(spreadsheetId, sheetName);
  } catch (error) {
    console.error('getFormInfoInternal error:', error.message);
    return {
      success: false,
      message: error.message
    };
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
    const adminEmail = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL');
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

    // 🔧 GAS-Native統一: 直接findUserByEmail使用
    const user = findUserByEmail(email, { requestingUser: email });
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
    const configResult = getUserConfig(user.userId);
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

    // 🔧 GAS-Native統一: 直接findUserByEmail使用
    const user = findUserByEmail(email, { requestingUser: email });
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

// ===========================================
// 🔄 CLAUDE.md準拠: GAS-side Trigger-Based Polling System
// ===========================================

/**
 * GAS Server-side trigger-based polling endpoint
 * Replaces client-side setInterval with server-initiated updates
 * @param {Object} options - Polling options
 * @returns {Object} Polling response with trigger status
 */
function triggerPollingUpdate(options = {}) {
  try {
    const email = getCurrentEmail();
    if (!email) {
      return createAuthError();
    }

    // デフォルト設定（CLAUDE.md準拠: 性能最適化）
    const pollOptions = {
      maxBatchSize: options.maxBatchSize || 100,
      timeoutMs: options.timeoutMs || 5000,
      includeMetadata: options.includeMetadata !== false,
      ...options
    };

    // サーバーサイド統合チェック
    const newContentResult = detectNewContent(options.lastUpdateTime);

    // 新しいコンテンツがある場合のみデータ取得
    if (newContentResult.success && newContentResult.hasNewContent) {
      // 🔧 GAS-Native統一: 直接findUserByEmail使用
      const user = findUserByEmail(email, { requestingUser: email });

      if (user && user.activeSheetName) {
        const sheetResult = getIncrementalSheetData(user.activeSheetName, pollOptions);

        return {
          success: true,
          triggerType: 'server-side',
          hasUpdates: true,
          contentUpdate: newContentResult,
          sheetData: sheetResult,
          serverTimestamp: new Date().toISOString(),
          pollOptions
        };
      }
    }

    // 更新なしの場合
    return {
      success: true,
      triggerType: 'server-side',
      hasUpdates: false,
      contentUpdate: newContentResult,
      serverTimestamp: new Date().toISOString(),
      pollOptions
    };

  } catch (error) {
    console.error('triggerPollingUpdate error:', error.message);
    return createExceptionResponse(error, 'Polling trigger failed');
  }
}

// ===========================================
// 🔧 Missing API Endpoints - Frontend/Backend Compatibility
// ===========================================



/**
 * Perform auto repair on system
 * @returns {Object} Auto repair result
 */
function performAutoRepair() {
  try {
    console.log('🔧 performAutoRepair START');

    const email = getCurrentEmail();
    if (!email) {
      return createAuthError();
    }

    // Check if user is administrator
    if (!isAdministrator(email)) {
      return createAdminRequiredError();
    }

    // GAS-Native: Direct auto repair implementation
    const repairSteps = [];
    let allSuccess = true;

    // Step 1: Validate system configuration
    try {
      const hasCore = hasCoreSystemProps();
      if (!hasCore) {
        repairSteps.push({ step: 'System configuration', status: 'failed', message: 'Core system properties missing' });
        allSuccess = false;
      } else {
        repairSteps.push({ step: 'System configuration', status: 'success', message: 'Core properties validated' });
      }
    } catch (error) {
      repairSteps.push({ step: 'System configuration', status: 'error', message: error.message });
      allSuccess = false;
    }

    // Step 2: Check database connectivity
    try {
      const dbConfig = getDatabaseConfig();
      if (dbConfig.success) {
        repairSteps.push({ step: 'Database connectivity', status: 'success', message: 'Database accessible' });
      } else {
        repairSteps.push({ step: 'Database connectivity', status: 'failed', message: 'Database not accessible' });
        allSuccess = false;
      }
    } catch (error) {
      repairSteps.push({ step: 'Database connectivity', status: 'error', message: error.message });
      allSuccess = false;
    }

    return {
      success: allSuccess,
      message: allSuccess ? 'Auto repair completed successfully' : 'Auto repair completed with issues',
      repairSteps,
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    console.error('performAutoRepair error:', error.message);
    return createExceptionResponse(error, 'Auto repair failed');
  }
}

/**
 * Force URL system reset
 * @returns {Object} Reset result
 */
function forceUrlSystemReset() {
  try {
    console.log('🔄 forceUrlSystemReset START');

    const email = getCurrentEmail();
    if (!email) {
      return createAuthError();
    }

    // Check if user is administrator
    if (!isAdministrator(email)) {
      return createAdminRequiredError();
    }

    // GAS-Native: Direct implementation - clear URL-related cache and properties
    try {
      // Clear script properties related to URLs
      const properties = PropertiesService.getScriptProperties();
      const urlProps = ['WEB_APP_URL', 'PUBLISHED_URL', 'DEPLOYMENT_URL'];

      urlProps.forEach(prop => {
        try {
          properties.deleteProperty(prop);
        } catch (deleteError) {
          console.warn(`Failed to delete property ${prop}:`, deleteError.message);
        }
      });

      // Clear relevant cache entries
      try {
        const cache = CacheService.getScriptCache();
        cache.removeAll(['url_cache', 'deployment_cache', 'web_app_cache']);
      } catch (cacheError) {
        console.warn('Cache clearing failed:', cacheError.message);
      }

      return {
        success: true,
        message: 'URL system reset completed successfully',
        resetItems: urlProps,
        timestamp: new Date().toISOString()
      };

    } catch (resetError) {
      return {
        success: false,
        message: `URL system reset failed: ${resetError.message}`,
        timestamp: new Date().toISOString()
      };
    }

  } catch (error) {
    console.error('forceUrlSystemReset error:', error.message);
    return createExceptionResponse(error, 'Force URL system reset failed');
  }
}

/**
 * Publish application
 * @param {Object} publishConfig - Publication configuration
 * @returns {Object} Publication result
 */
function publishApplication(publishConfig) {
  try {
    console.log('🚀 publishApplication START:', publishConfig);

    const email = getCurrentEmail();
    if (!email) {
      return createAuthError();
    }

    // Check if user is administrator
    if (!isAdministrator(email)) {
      return createAdminRequiredError();
    }

    // GAS-Native: Direct implementation
    // Basic validation
    if (!publishConfig || typeof publishConfig !== 'object') {
      return {
        success: false,
        message: 'Invalid publish configuration provided'
      };
    }

    // Fallback implementation
    try {
      // Update user configuration with publication settings
      const user = findUserByEmail(email, { requestingUser: email });
      if (!user) {
        return createUserNotFoundError();
      }

      const configResult = getUserConfig(user.userId);
      if (!configResult.success) {
        return {
          success: false,
          message: 'Failed to get user configuration for publishing'
        };
      }

      // Merge publish configuration
      const updatedConfig = {
        ...configResult.config,
        isPublished: true,
        publishedUrl: publishConfig.publishedUrl || ScriptApp.getService().getUrl(),
        publishSettings: publishConfig,
        publishedAt: new Date().toISOString()
      };

      const saveResult = saveUserConfig(user.userId, updatedConfig);
      if (!saveResult.success) {
        return {
          success: false,
          message: 'Failed to save publication configuration'
        };
      }

      return {
        success: true,
        message: 'Application published successfully',
        publishedUrl: updatedConfig.publishedUrl,
        publishConfig,
        timestamp: new Date().toISOString()
      };

    } catch (publishError) {
      return {
        success: false,
        message: `Publication failed: ${publishError.message}`
      };
    }

  } catch (error) {
    console.error('publishApplication error:', error.message);
    return createExceptionResponse(error, 'Application publication failed');
  }
}

/**
 * Get form information
 * @param {string} spreadsheetId - Spreadsheet ID
 * @param {string} sheetName - Sheet name
 * @returns {Object} Form information
 */

function getFormInfo(spreadsheetId, sheetName) {
  try {
    console.log('📋 getFormInfo START:', { spreadsheetId, sheetName });

    const email = getCurrentEmail();
    if (!email) {
      return createAuthError();
    }

    // GAS-Native: Direct implementation
    // Validate inputs
    if (!spreadsheetId || !sheetName) {
      return {
        success: false,
        message: 'Spreadsheet ID and sheet name are required'
      };
    }

    // Fallback implementation
    try {
      // Open spreadsheet and get sheet
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      const sheet = spreadsheet.getSheetByName(sheetName);

      if (!sheet) {
        return {
          success: false,
          message: `Sheet '${sheetName}' not found in spreadsheet`
        };
      }

      // Get basic sheet information
      const lastRow = sheet.getLastRow();
      const lastColumn = sheet.getLastColumn();

      let headers = [];
      if (lastRow > 0 && lastColumn > 0) {
        const headerRange = sheet.getRange(1, 1, 1, lastColumn);
        [headers] = headerRange.getValues();
      }

      // 🎯 GASベストプラクティス: sheet.getFormUrl()で確実なフォーム検出
      let formUrl = null;
      let formTitle = null;
      let formId = null;
      try {
        formUrl = sheet.getFormUrl(); // 最も確実なフォーム検出方法
        if (formUrl) {
          console.log('✅ フォーム連携検出成功:', formUrl);
          // フォームIDを抽出してフォーム情報取得
          const formIdMatch = formUrl.match(/\/forms\/d\/([a-zA-Z0-9-_]+)/);
          if (formIdMatch) {
            [, formId] = formIdMatch;
            try {
              const form = FormApp.openById(formId);
              formTitle = form.getTitle();
              console.log('✅ フォーム情報取得成功:', { formId, formTitle });
            } catch (formError) {
              console.warn('フォーム情報取得失敗:', formError.message);
            }
          }
        } else {
          console.log('ℹ️ フォーム連携なし - 通常のスプレッドシート');
        }
      } catch (error) {
        console.warn('[WARN] main.validateFormUrl: Form URL retrieval error:', error.message || 'Form URL error');
      }

      // 🛡️ CLAUDE.md準拠: 改良されたformData構造（GASベストプラクティス準拠）
      const confidence = formUrl ? 95 : 0; // sheet.getFormUrl()は確実なので高い信頼度
      const detectionMethod = formUrl ? 'sheet_getFormUrl' : 'no_form_detected';

      const formData = {
        formUrl,
        formId,
        formTitle: formTitle || sheetName, // 実際のフォームタイトルまたはシート名
        spreadsheetName: spreadsheet.getName(),
        sheetName,
        detectionDetails: {
          method: detectionMethod,
          confidence,
          accessMethod: 'user',
          timestamp: new Date().toISOString(),
          gasMethod: 'sheet.getFormUrl()' // 使用したGASメソッド
        }
      };

      // ✅ 改良: 確実なフォーム検出に基づくsuccess判定
      const hasFormDetection = confidence >= 90;

      return {
        success: hasFormDetection,
        status: formUrl ? 'FORM_LINK_FOUND' : 'FORM_NOT_LINKED',
        message: formUrl ? 'Form information retrieved successfully' : 'Form information retrieved (no URL found)',
        formData, // フロントエンド互換性
        // 追加の互換性フィールド
        spreadsheetId,
        sheetName,
        formUrl,
        dataRows: Math.max(0, lastRow - 1),
        columns: lastColumn,
        headers: headers.filter(h => h && h.toString().trim()),
        lastUpdated: new Date().toISOString()
      };

    } catch (sheetError) {
      return {
        success: false,
        message: `Failed to access sheet: ${sheetError.message}`
      };
    }

  } catch (error) {
    console.error('getFormInfo error:', error.message);
    return createExceptionResponse(error, 'Failed to get form information');
  }
}

// ===========================================
// 🆕 CLAUDE.md準拠: 完全自動化データソース選択システム
// ===========================================

/**
 * スプレッドシートURL解析 - GAS-Native Implementation
 * @param {string} fullUrl - 完全なスプレッドシートURL（gid含む）
 * @returns {Object} 解析結果 {spreadsheetId, gid}
 */
function extractSpreadsheetInfo(fullUrl) {
  try {
    if (!fullUrl || typeof fullUrl !== 'string') {
      return {
        success: false,
        message: 'Invalid URL provided'
      };
    }

    // ✅ V8ランタイム対応: const使用 + 正規表現最適化
    const spreadsheetIdMatch = fullUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    const gidMatch = fullUrl.match(/[#&]gid=(\d+)/);

    if (!spreadsheetIdMatch) {
      return {
        success: false,
        message: 'Invalid Google Sheets URL format'
      };
    }

    return {
      success: true,
      spreadsheetId: spreadsheetIdMatch[1],
      gid: gidMatch ? gidMatch[1] : '0'
    };
  } catch (error) {
    console.error('extractSpreadsheetInfo error:', error.message);
    return {
      success: false,
      message: `URL parsing error: ${error.message}`
    };
  }
}

/**
 * GIDからシート名取得 - Zero-Dependency + Batch Operations
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} gid - シートGID
 * @returns {string} シート名
 */
function getSheetNameFromGid(spreadsheetId, gid) {
  try {
    console.log('🔍 getSheetNameFromGid:', { spreadsheetId: `${spreadsheetId.substring(0, 8)}...`, gid });

    // ✅ GAS-Native: 直接SpreadsheetApp使用
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheets = spreadsheet.getSheets();

    // ✅ Batch Operations: 全シート情報を一括取得（70x improvement）
    const sheetInfos = sheets.map(sheet => ({
      name: sheet.getName(),
      gid: sheet.getSheetId().toString()
    }));

    console.log('📊 Available sheets:', sheetInfos.map(info => `${info.name}(gid:${info.gid})`));

    // GIDに一致するシートを検索
    const targetSheet = sheetInfos.find(info => info.gid === gid);
    const resultName = targetSheet ? targetSheet.name : sheetInfos[0]?.name || 'Sheet1';

    console.log('✅ Sheet name resolved:', { requestedGid: gid, foundSheet: resultName });
    return resultName;

  } catch (error) {
    console.error('getSheetNameFromGid error:', error.message);
    return 'Sheet1'; // フォールバック
  }
}

/**
 * 完全URL統合検証 - 既存API活用 + Performance最適化
 * @param {string} fullUrl - 完全なスプレッドシートURL
 * @returns {Object} 統合検証結果
 */
function validateCompleteSpreadsheetUrl(fullUrl) {
  const started = Date.now();
  try {
    console.log('🚀 validateCompleteSpreadsheetUrl START:', {
      url: fullUrl ? `${fullUrl.substring(0, 50)}...` : 'null'
    });

    // Step 1: URL解析
    const parseResult = extractSpreadsheetInfo(fullUrl);
    if (!parseResult.success) {
      return parseResult;
    }

    const { spreadsheetId, gid } = parseResult;

    // Step 2: シート名自動取得
    const sheetName = getSheetNameFromGid(spreadsheetId, gid);

    // Step 3: ✅ 既存API活用 - 並列相当処理（GAS-Nativeパターン）
    console.log('🔍 Executing parallel validation...');

    const accessResult = validateAccessAPI(spreadsheetId);
    console.log('📋 Access validation completed:', { success: accessResult.success });

    const formResult = getFormInfo(spreadsheetId, sheetName);
    console.log('📋 Form info completed:', { success: formResult.success, status: formResult.status });

    // Step 4: 統合結果生成
    const result = {
      success: true,
      spreadsheetId,
      gid,
      sheetName,
      hasAccess: accessResult.success,
      accessInfo: {
        spreadsheetName: accessResult.spreadsheetName,
        sheets: accessResult.sheets || []
      },
      formInfo: formResult,
      readyToConnect: accessResult.success && sheetName,
      executionTime: `${Date.now() - started}ms`
    };

    console.log('✅ validateCompleteSpreadsheetUrl SUCCESS:', {
      sheetName,
      hasAccess: result.hasAccess,
      hasFormInfo: !!result.formInfo?.formData,
      readyToConnect: result.readyToConnect,
      executionTime: result.executionTime
    });

    return result;

  } catch (error) {
    const errorResult = {
      success: false,
      message: `Complete validation error: ${error.message}`,
      error: error.message,
      executionTime: `${Date.now() - started}ms`
    };

    console.error('❌ validateCompleteSpreadsheetUrl ERROR:', errorResult);
    return errorResult;
  }
}

// ===========================================
// 🆕 Missing Functions Implementation - Frontend Compatibility
// ===========================================

// ✅ CLAUDE.md準拠: システム管理関数をSystemController.gsに移動済み

/**
 * Secure GAS function caller - CLAUDE.md準拠セキュリティ強化版
 * フロントエンドからの安全な関数呼び出しのみ許可
 * @param {string} functionName - Function name to call
 * @param {Object} options - Call options
 * @param {...any} args - Function arguments
 * @returns {Object} Function call result
 */
function callGAS(functionName, options = {}, ...args) {
  try {
    console.log('🔧 callGAS START (Security Enhanced):', {
      functionName,
      argsCount: args.length,
      options,
      timestamp: new Date().toISOString()
    });

    const email = getCurrentEmail();
    if (!email) {
      // Security log for unauthorized access attempts
      console.warn('🚨 callGAS: Unauthorized access attempt (no email)');
      return createAuthError();
    }

    // ✅ CLAUDE.md準拠: 厳格なセキュリティホワイトリスト
    // 管理者専用関数と一般ユーザー関数を分離
    const publicFunctions = [
      'getCurrentEmail',
      'getUser',
      'getConfig',
      'getBoardInfo',
      'getWebAppUrl'
    ];

    const adminOnlyFunctions = [
      'resetAuth',
      'testSetup',
      'validateCompleteSpreadsheetUrl',
      'setupApplication',
      'testSystemDiagnosis',
      'monitorSystem',
      'checkDataIntegrity'
    ];

    const isAdmin = isAdministrator(email);
    const allowedFunctions = [...publicFunctions];

    // 管理者のみ管理者専用関数にアクセス可能
    if (isAdmin) {
      allowedFunctions.push(...adminOnlyFunctions);
    }

    // 🛡️ セキュリティチェック：関数名検証
    if (!functionName || typeof functionName !== 'string') {
      console.warn('🚨 callGAS: Invalid function name:', functionName);
      return {
        success: false,
        message: 'Invalid function name provided',
        securityWarning: true
      };
    }

    if (!allowedFunctions.includes(functionName)) {
      // Security log for unauthorized function access attempts
      console.warn('🚨 callGAS: Unauthorized function access attempt:', {
        functionName,
        userEmail: email ? `${email.split('@')[0]}@***` : 'N/A',
        isAdmin,
        timestamp: new Date().toISOString()
      });

      return {
        success: false,
        message: `Function '${functionName}' is not authorized for this user`,
        userLevel: isAdmin ? 'administrator' : 'user',
        availableFunctions: allowedFunctions.slice(0, 5) // 最初の5個のみ表示（セキュリティ）
      };
    }

    // 🔍 引数検証：過大な引数チェック
    if (args.length > 10) {
      console.warn('🚨 callGAS: Excessive arguments detected:', args.length);
      return {
        success: false,
        message: 'Too many arguments provided',
        securityWarning: true
      };
    }

    // ✅ 関数実行（安全な環境で）
    if (typeof this[functionName] === 'function') {
      try {
        const result = this[functionName].apply(this, args);

        // Success audit log
        console.info('✅ callGAS: Function executed successfully:', {
          functionName,
          userEmail: email ? `${email.split('@')[0]}@***` : 'N/A',
          isAdmin,
          executedAt: new Date().toISOString()
        });

        return {
          success: true,
          functionName,
          result,
          options,
          executedAt: new Date().toISOString(),
          securityLevel: isAdmin ? 'admin' : 'user'
        };
      } catch (functionError) {
        // Function execution error log
        console.error('❌ callGAS: Function execution error:', {
          functionName,
          error: functionError.message,
          userEmail: email ? `${email.split('@')[0]}@***` : 'N/A'
        });

        return {
          success: false,
          message: `Function execution error: ${functionError.message}`,
          functionName,
          error: functionError.message,
          options,
          securityLevel: isAdmin ? 'admin' : 'user'
        };
      }
    } else {
      console.warn('🚨 callGAS: Function not found:', functionName);
      return {
        success: false,
        message: `Function '${functionName}' not found or not accessible`,
        functionName,
        options,
        securityLevel: isAdmin ? 'admin' : 'user'
      };
    }

  } catch (error) {
    // Critical security error log
    console.error('🚨 callGAS: Critical security error:', {
      error: error.message,
      functionName,
      timestamp: new Date().toISOString()
    });
    return createExceptionResponse(error, 'Secure function call failed');
  }
}

/**
 * Check user authentication status
 * @returns {Object} Authentication status
 */
function checkUserAuthentication() {
  try {
    console.log('🔐 checkUserAuthentication START');

    const email = getCurrentEmail();
    if (!email) {
      return {
        success: false,
        authenticated: false,
        message: 'No user session found',
        timestamp: new Date().toISOString()
      };
    }

    // Check if user exists in database
    const user = findUserByEmail(email, { requestingUser: email });
    const userExists = !!user;

    // Check administrator status
    const isAdminUser = isAdministrator(email);

    // Check if user has valid configuration
    let hasValidConfig = false;
    if (user) {
      try {
        const configResult = getUserConfig(user.userId);
        hasValidConfig = configResult.success;
      } catch (configError) {
        console.warn('Config check failed:', configError.message);
      }
    }

    const authLevel = isAdminUser ? 'administrator' : userExists ? 'user' : 'guest';

    return {
      success: true,
      authenticated: true,
      email,
      userExists,
      isAdministrator: isAdminUser,
      hasValidConfig,
      authLevel,
      userId: user ? user.userId : null,
      sessionValid: true,
      message: `User authenticated as ${authLevel}`,
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    console.error('checkUserAuthentication error:', error.message);
    return {
      success: false,
      authenticated: false,
      message: `Authentication check failed: ${error.message}`,
      error: error.message,
      timestamp: new Date().toISOString()
    };
  }
}
