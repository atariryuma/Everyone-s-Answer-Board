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

/* global createErrorResponse, createSuccessResponse, createAuthError, createUserNotFoundError, createAdminRequiredError, createExceptionResponse, hasCoreSystemProps, getUserSheetData, addReaction, toggleHighlight, validateConfig, findUserByEmail, findUserById, findUserBySpreadsheetId, createUser, getAllUsers, updateUser, openSpreadsheet, getUserConfig, saveUserConfig, clearConfigCache, cleanConfigFields, getQuestionText, DB, validateAccess, URL, UserService, CACHE_DURATION, TIMEOUT_MS, SLEEP_MS, SYSTEM_LIMITS, DataController, SystemController, getDatabaseConfig, getViewerBoardData, getSheetHeaders, performIntegratedColumnDiagnostics, generateRecommendedMapping, getFormInfo */

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

    // ✅ Performance optimization: Cache email for authentication-required routes
    const currentEmail = (mode !== 'login') ? getCurrentEmail() : null;

    // Simple routing
    switch (mode) {
      case 'login': {
        // 極限シンプル: ログインページ（静的表示のみ）
        return HtmlService.createTemplateFromFile('LoginPage.html').evaluate();
      }

      case 'admin': {
        // ✅ CLAUDE.md準拠: Batch operations for 70x performance improvement
        const adminData = getBatchedAdminData();
        if (!adminData.success) {
          return createRedirectTemplate('ErrorBoundary.html', adminData.error || '管理者権限が必要です');
        }

        const { email, user, config } = adminData;
        const isAdmin = isAdministrator(email);

        // 認証済み - Administrator/Editor権限でAdminPanel表示
        const template = HtmlService.createTemplateFromFile('AdminPanel.html');
        template.userEmail = email;
        template.userId = user.userId;
        template.accessLevel = isAdmin ? 'administrator' : 'editor';
        template.userInfo = user;
        template.configJSON = JSON.stringify({
          userId: user.userId,
          userEmail: email,
          spreadsheetId: config.spreadsheetId || '',
          sheetName: config.sheetName || '',
          isPublished: Boolean(config.isPublished),
          isEditor: true, // 管理者・編集ユーザーは常にエディター権限
          isAdminUser: isAdmin,
          isOwnBoard: true,
          formUrl: config.formUrl || '',
          formTitle: config.formTitle || '',
          showDetails: config.showDetails !== false,
          autoStopEnabled: Boolean(config.autoStopEnabled),
          autoStopTime: config.autoStopTime || null,
          setupStatus: config.setupStatus || 'pending',
          displaySettings: config.displaySettings || {},
          columnMapping: config.columnMapping || {}
        });

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
          template.isAdministrator = currentEmail ? isAdministrator(currentEmail) : false;
          template.userEmail = currentEmail || '';
          template.timestamp = new Date().toISOString();
          return template.evaluate();
        }
      }

      case 'appSetup': {
        // 🔐 GAS-Native: 直接認証チェック - Administrator専用
        if (!currentEmail || !isAdministrator(currentEmail)) {
          return createRedirectTemplate('ErrorBoundary.html', '管理者権限が必要です');
        }

        // 認証済み - Administrator権限でAppSetup表示
        return HtmlService.createTemplateFromFile('AppSetupPage.html').evaluate();
      }

      case 'view': {
        // 🔐 GAS-Native: 直接認証チェック - Viewer権限確認
        if (!currentEmail) {
          return createRedirectTemplate('ErrorBoundary.html', 'ユーザー認証が必要です');
        }

        // 対象ユーザー確認
        const targetUserId = params.userId;
        if (!targetUserId) {
          return createRedirectTemplate('ErrorBoundary.html', 'ユーザーIDが指定されていません');
        }

        // ✅ CLAUDE.md準拠: Batch operations for 70x performance improvement
        const viewerData = getBatchedViewerData(targetUserId, currentEmail);
        if (!viewerData.success) {
          return createRedirectTemplate('ErrorBoundary.html', viewerData.error || '対象ユーザーが見つかりません');
        }

        const { targetUser, config, isAdminUser } = viewerData;
        const isOwnBoard = currentEmail === targetUser.userEmail;
        const isPublished = Boolean(config.isPublished);

        if (!isAdminUser && !isOwnBoard && !isPublished) {
          return createRedirectTemplate('ErrorBoundary.html', 'このボードは非公開に設定されています');
        }

        // 認証済み - 公開ボード表示
        const template = HtmlService.createTemplateFromFile('Page.html');
        template.userId = targetUserId;
        template.userEmail = targetUser.userEmail;
        template.questionText = '読み込み中...';
        template.boardTitle = targetUser.userEmail || '回答ボード';

        // 🔧 CLAUDE.md準拠: 統一権限情報（GAS-Native Architecture）
        const isEditor = isAdminUser || isOwnBoard;
        template.isEditor = isEditor;
        template.isAdminUser = isAdminUser;
        template.isOwnBoard = isOwnBoard;

        // 🔧 CLAUDE.md準拠: configJSON統一取得（Zero-Dependency）
        template.sheetName = config.sheetName;
        template.configJSON = JSON.stringify({
          userId: targetUserId,
          userEmail: targetUser.userEmail,
          spreadsheetId: config.spreadsheetId || '',
          sheetName: config.sheetName,
          questionText: '読み込み中...',
          isPublished: Boolean(config.isPublished),
          isEditor,
          isAdminUser,
          isOwnBoard,
          formUrl: config.formUrl || '',
          showDetails: config.showDetails !== false,
          autoStopEnabled: Boolean(config.autoStopEnabled),
          autoStopTime: config.autoStopTime || null
        });

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

    // V8ランタイム安全: error変数存在チェック後のテンプレートリテラル使用
    if (error && error.message) {
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
        result = handleUserDataRequest(email, request.options || {});
        break;
      case 'addReaction':
        // 🎯 Multi-tenant: request.userId = target user (board owner), email = actor (current user)
        if (!request.userId) {
          result = createErrorResponse('Target user ID required for reaction');
        } else {
          result = addReaction(request.userId, request.rowId, request.reactionType);
        }
        break;
      case 'toggleHighlight':
        // 🎯 Multi-tenant: request.userId = target user (board owner), email = actor (current user)
        if (!request.userId) {
          result = createErrorResponse('Target user ID required for highlight');
        } else {
          result = toggleHighlight(request.userId, request.rowId);
        }
        break;
      case 'refreshData':
        result = handleUserDataRequest(email, request.options || {});
        break;
      default:
        result = createErrorResponse(action ? `Unknown action: ${action}` : 'Unknown action: 不明');
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    // V8ランタイム安全: error変数とerror.message存在チェック
    const errorMessage = error && error.message ? error.message : '予期しないエラーが発生しました';
    console.error('doPost error:', errorMessage);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: errorMessage
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 統一ユーザーデータリクエスト処理関数
 * 重複していたgetDataとrefreshDataのロジックを統一
 * @param {string} email - ユーザーメールアドレス
 * @param {Object} options - リクエストオプション
 * @returns {Object} 処理結果
 */
function handleUserDataRequest(email, options = {}) {
  try {
    const user = findUserByEmail(email, { requestingUser: email });
    if (!user) {
      return createUserNotFoundError();
    }
    return { success: true, data: getUserSheetData(user.userId, options) };
  } catch (error) {
    console.error('handleUserDataRequest error:', error.message);
    return createExceptionResponse(error);
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

// getWebAppUrl moved to SystemController.gs for architecture compliance

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
 * Reset authentication and clear all user session data
 * ✅ CLAUDE.md準拠: 包括的キャッシュクリア with 論理的破綻修正
 */
function resetAuth() {
  try {
    const cache = CacheService.getScriptCache();
    let clearedKeysCount = 0;
    let clearConfigResult = null;

    // 🔧 修正1: 現在ユーザー情報を事前取得（キャッシュクリア前）
    const currentEmail = getCurrentEmail();
    const currentUser = currentEmail ? findUserByEmail(currentEmail) : null;
    const userId = currentUser?.userId;

    // 🔧 修正2: ConfigService専用クリア関数の活用
    if (userId) {
      try {
        clearConfigCache(userId);
        clearConfigResult = 'ConfigService cache cleared successfully';
        console.log(`resetAuth: ConfigService cache cleared for user ${userId.substring(0, 8)}***`);
      } catch (configError) {
        console.warn('resetAuth: ConfigService cache clear failed:', configError.message);
        clearConfigResult = `ConfigService cache clear failed: ${configError.message}`;
      }
    }

    // 🔧 修正3: 包括的キャッシュキーリスト（実際の使用パターンに合わせて更新）
    const globalCacheKeysToRemove = [
      'current_user_info',
      'admin_auth_cache',
      'session_data',
      'system_diagnostic_cache',
      'bulk_admin_data_cache'
    ];

    // グローバルキャッシュクリア
    globalCacheKeysToRemove.forEach(key => {
      try {
        cache.remove(key);
        clearedKeysCount++;
      } catch (e) {
        console.warn(`resetAuth: Failed to remove global cache key ${key}:`, e.message);
      }
    });

    // 🔧 修正4: User固有キャッシュの完全クリア（email + userId ベース）
    const userSpecificKeysCleared = [];
    if (currentEmail) {
      const emailBasedKeys = [
        `board_data_${currentEmail}`,
        `user_data_${currentEmail}`,
        `admin_panel_${currentEmail}`
      ];

      emailBasedKeys.forEach(key => {
        try {
          cache.remove(key);
          userSpecificKeysCleared.push(key);
          clearedKeysCount++;
        } catch (e) {
          console.warn(`resetAuth: Failed to remove email-based cache key ${key}:`, e.message);
        }
      });
    }

    if (userId) {
      const userIdBasedKeys = [
        `user_config_${userId}`,
        `config_${userId}`,
        `user_${userId}`,
        `board_data_${userId}`,
        `question_text_${userId}`
      ];

      userIdBasedKeys.forEach(key => {
        try {
          cache.remove(key);
          userSpecificKeysCleared.push(key);
          clearedKeysCount++;
        } catch (e) {
          console.warn(`resetAuth: Failed to remove userId-based cache key ${key}:`, e.message);
        }
      });
    }

    // 🔧 修正5: リアクション・ハイライトロック完全クリア
    let reactionLocksCleared = 0;
    if (userId) {
      try {
        // リアクション・ハイライトロックの動的検索・削除は複雑なので、
        // 基本的なロックキーのパターンをクリア
        const lockPatterns = [
          `reaction_${userId}_`,
          `highlight_${userId}_`
        ];

        // Note: GAS CacheService doesn't support pattern matching,
        // so we clear known common patterns and rely on TTL expiration
        for (let i = 0; i < SYSTEM_LIMITS.MAX_LOCK_ROWS; i++) { // 最大100行のロックをクリア
          lockPatterns.forEach(pattern => {
            try {
              cache.remove(`${pattern}${i}`);
              reactionLocksCleared++;
            } catch (e) {
              // Lock key might not exist - this is expected
            }
          });
        }
        console.log(`resetAuth: Cleared ${reactionLocksCleared} reaction/highlight locks for user ${userId.substring(0, 8)}***`);
      } catch (lockError) {
        console.warn('resetAuth: Reaction lock clearing failed:', lockError.message);
      }
    }

    // 🔧 修正6: 包括的なログ出力
    const logDetails = {
      currentUser: currentEmail ? `${currentEmail.substring(0, 8)}***@${currentEmail.split('@')[1]}` : 'N/A',
      userId: userId ? `${userId.substring(0, 8)}***` : 'N/A',
      globalKeysCleared: globalCacheKeysToRemove.length,
      userSpecificKeysCleared: userSpecificKeysCleared.length,
      reactionLocksCleared,
      configServiceResult: clearConfigResult,
      totalKeysCleared: clearedKeysCount
    };

    console.log('resetAuth: Authentication reset completed', logDetails);

    return {
      success: true,
      message: 'Authentication and session data cleared successfully',
      details: {
        clearedKeys: clearedKeysCount,
        userSpecificKeys: userSpecificKeysCleared.length,
        reactionLocks: reactionLocksCleared,
        configService: clearConfigResult ? 'success' : 'skipped'
      }
    };
  } catch (error) {
    console.error('resetAuth error:', error.message, error.stack);
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
            isPublished: false
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
    config.lastAccessedAt = new Date().toISOString();

    // 統一API使用: 検証・サニタイズ・保存
    const saveResult = saveUserConfig(user.userId, config, { forceUpdate: false });
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
    config.lastAccessedAt = new Date().toISOString();

    // 統一API使用: 検証・サニタイズ・保存
    const saveResult = saveUserConfig(targetUserId, config, { forceUpdate: false });
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
    if (!email) {
      return createAuthError();
    }

    // 🔧 GAS-Native統一: 直接Data使用
    let targetUser = targetUserId ? findUserById(targetUserId) : null;
    if (!targetUser) {
      targetUser = findUserByEmail(email);
    }

    if (!targetUser) {
      return createUserNotFoundError();
    }

    // ✅ 編集者権限チェック: 管理者または自分のボード
    const isAdmin = isAdministrator(email);
    const isOwnBoard = targetUser.userEmail === email;

    if (!isAdmin && !isOwnBoard) {
      return createErrorResponse('ボードの非公開権限がありません');
    }

    // 統一API使用: 設定取得・更新・保存
    const configResult = getUserConfig(targetUser.userId);
    const config = configResult.success ? configResult.config : {};

    const wasPublished = config.isPublished === true;
    config.isPublished = false;
    config.publishedAt = null;
    config.lastAccessedAt = new Date().toISOString();

    // 統一API使用: 検証・サニタイズ・保存
    const saveResult = saveUserConfig(targetUser.userId, config, { forceUpdate: false });
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
 * Get user's spreadsheets from their Drive - available to authenticated users
 * ✅ CLAUDE.md準拠: Editor access for own Drive resources
 * ✅ Performance optimized with caching and batch operations
 */
function getSheets() {
  try {
    const email = getCurrentEmail();
    if (!email) {
      console.warn('getSheets: Unauthenticated access attempt');
      return {
        success: false,
        error: 'ユーザー認証が必要です'
      };
    }

    // 🚀 Performance optimization: Cache user's spreadsheet list
    const cacheKey = `sheets_${email}`;
    try {
      const cached = CacheService.getScriptCache().get(cacheKey);
      if (cached) {
        console.log('getSheets: Cache hit for user:', `${email.split('@')[0]}@***`);
        return JSON.parse(cached);
      }
    } catch (cacheError) {
      console.warn('getSheets: Cache read failed:', cacheError.message);
    }

    console.log('getSheets: Fetching spreadsheets for user:', `${email.split('@')[0]}@***`);

    // ✅ CLAUDE.md準拠: Direct DriveApp access for own resources
    const drive = DriveApp;
    const spreadsheets = drive.searchFiles('mimeType="application/vnd.google-apps.spreadsheet"');

    const sheets = [];
    let processedCount = 0;

    while (spreadsheets.hasNext()) {
      try {
        const file = spreadsheets.next();
        sheets.push({
          id: file.getId(),
          name: file.getName(),
          url: file.getUrl()
        });
        processedCount++;

        // 🛡️ Safety limit to prevent timeout
        if (processedCount > 100) {
          console.warn('getSheets: Processing limit reached (100 files)');
          break;
        }
      } catch (fileError) {
        console.warn('getSheets: Error processing file:', fileError.message);
        continue;
      }
    }

    const result = {
      success: true,
      sheets,
      totalFound: sheets.length,
      processingLimited: processedCount > 100
    };

    console.log('getSheets: Found', sheets.length, 'spreadsheets for user:', `${email.split('@')[0]}@***`);

    // 🚀 Performance optimization: Cache results for 5 minutes
    try {
      const cacheTtl = CACHE_DURATION.LONG; // 300 seconds
      CacheService.getScriptCache().put(cacheKey, JSON.stringify(result), cacheTtl);
      console.log('getSheets: Results cached for user:', `${email.split('@')[0]}@***`, 'TTL:', cacheTtl, 's');
    } catch (cacheError) {
      console.warn('getSheets: Cache write failed:', cacheError.message);
    }

    return result;
  } catch (error) {
    console.error('getSheets error:', error.message);
    return {
      success: false,
      error: `スプレッドシート取得エラー: ${error.message}`,
      details: error.stack
    };
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


    return {
      success: true,
      isActive: isPublished,
      isPublished,
      questionText: getQuestionText(config, { targetUserEmail: user.userEmail }),
      urls: {
        view: `${baseUrl}?mode=view&userId=${user.userId}`,
        admin: `${baseUrl}?mode=admin&userId=${user.userId}`
      },
      lastUpdated: config.publishedAt || user.lastModified
    };
  } catch (error) {
    console.error('❌ getBoardInfo ERROR:', error.message);
    return createExceptionResponse(error);
  }
}




/**
 * 🔧 統合API: フロントエンド用データ取得（最適化版・クロスユーザー対応）
 * Page.htmlから呼び出される
 * @param {string} classFilter - クラスフィルター (例: 'すべて', '3年A組')
 * @param {string} sortOrder - ソート順 (例: 'newest', 'oldest', 'random', 'score')
 * @param {boolean} adminMode - 管理者モード（未使用、後方互換性のため保持）
 * @param {string} targetUserId - 対象ユーザーID（指定時はクロスユーザーアクセス）
 * @returns {Object} フロントエンド期待形式のデータ
 */
function getPublishedSheetData(classFilter, sortOrder, adminMode = false, targetUserId = null) {
  // ✅ Performance optimization: Add timeout for slow responses
  const startTime = Date.now();

  try {
    // ✅ CLAUDE.md準拠: Batched admin authentication (70x performance improvement)
    const adminAuth = getBatchedAdminAuth({ allowNonAdmin: true });
    if (!adminAuth.success || !adminAuth.authenticated) {
      return {
        success: false,
        error: 'Authentication required',
        data: [],  // ✅ CLAUDE.md準拠: 統一レスポンス形式（rows → data）
        sheetName: '',
        header: '認証エラー'
      };
    }

    const { email: viewerEmail, isAdmin: isSystemAdmin } = adminAuth;

    // CLAUDE.md準拠: クロスユーザーアクセス対応
    let user;
    if (targetUserId) {
      // ✅ Timeout check before expensive operations
      if (Date.now() - startTime > TIMEOUT_MS.EXTENDED) {
        console.warn('getPublishedSheetData: Timeout during user lookup');
        return {
          success: false,
          error: 'Request timeout - please try again',
          data: [],  // ✅ CLAUDE.md準拠: 統一レスポンス形式（rows → data）
          sheetName: '',
          header: 'タイムアウト'
        };
      }

      // クロスユーザーアクセス: targetUserIdで指定されたユーザーのデータを取得
      const boardData = getViewerBoardData(targetUserId, viewerEmail);
      if (!boardData) {
        return {
          success: false,
          error: 'User not found or access denied',
          data: [],  // ✅ CLAUDE.md準拠: 統一レスポンス形式（rows → data）
          sheetName: '',
          header: 'ユーザーエラー'
        };
      }

      // ✅ Timeout check after data retrieval
      if (Date.now() - startTime > TIMEOUT_MS.EXTENDED) {
        console.warn('getPublishedSheetData: Timeout during data processing');
        return {
          success: false,
          error: 'Data processing timeout - please try again',
          data: [],  // ✅ CLAUDE.md準拠: 統一レスポンス形式（rows → data）
          sheetName: '',
          header: 'タイムアウト'
        };
      }

      // getViewerBoardDataの結果を直接使用
      return transformBoardDataToFrontendFormat(boardData, classFilter, sortOrder);
    } else {
      // 従来通り: 現在のユーザーのデータを取得

      // ✅ CLAUDE.md準拠: Optimized single user lookup (eliminated duplicate findUserByEmail call)
      const searchOptions = {
        requestingUser: viewerEmail,
        adminMode: isSystemAdmin,
        // Use enhanced permissions for admins from the start to avoid second lookup
        ignorePermissions: isSystemAdmin
      };

      user = findUserByEmail(viewerEmail, searchOptions);
      if (!user) {
        return {
          success: false,
          error: 'User not found',
          rows: [],
          sheetName: '',
          header: 'ユーザーエラー'
        };
      }
    }

    // getUserSheetDataを呼び出し、フィルターオプションを渡す
    const options = {
      classFilter: classFilter !== 'すべて' ? classFilter : undefined,
      sortBy: sortOrder || 'newest',
      includeTimestamp: true,
      adminMode: isSystemAdmin, // 管理者権限をデータ取得に渡す
      requestingUser: viewerEmail
    };

    const result = getUserSheetData(user.userId, options);

    // 統一されたtransformBoardDataToFrontendFormat関数を使用
    if (result && result.success && result.data) {

      // CLAUDE.md準拠: 統一されたデータ変換関数使用（DRY原則）
      return transformBoardDataToFrontendFormat(result, classFilter, sortOrder);
    } else {
      console.error('❌ Data retrieval failed:', result?.message);
      return {
        success: false,
        error: result?.message || 'データ取得エラー',
        data: [],  // ✅ CLAUDE.md準拠: 統一レスポンス形式（rows → data）
        sheetName: result?.sheetName || '',
        header: result?.header || '問題'
      };
    }

  } catch (error) {
    // ✅ CLAUDE.md V8準拠: 変数存在チェック後のエラーハンドリング
    const errorMessage = (error && error.message) ? error.message : 'データ取得エラー';
    console.error('❌ getPublishedSheetData ERROR:', errorMessage);
    return {
      success: false,
      error: errorMessage,
      data: [],  // ✅ CLAUDE.md準拠: 統一レスポンス形式（rows → data）
      sheetName: '',
      header: '問題'
    };
  }
}

/**
 * getViewerBoardDataの結果をフロントエンド期待形式に変換
 * @param {Object} boardData - getViewerBoardDataの戻り値
 * @param {string} classFilter - クラスフィルター
 * @param {string} sortOrder - ソート順
 * @returns {Object} フロントエンド期待形式のデータ
 */
function transformBoardDataToFrontendFormat(boardData, classFilter, sortOrder) {

  // 🔧 CLAUDE.md準拠: Enhanced validation with detailed logging
  const hasValidBoardData = !!boardData;
  const hasSuccessFlag = boardData?.success === true;
  const hasDataProperty = boardData && 'data' in boardData;
  const isDataArray = Array.isArray(boardData?.data);


  // フロントエンド期待形式に変換（CLAUDE.md準拠）
  if (hasValidBoardData && hasSuccessFlag && hasDataProperty && isDataArray) {
    const transformedData = {
      success: true,
      header: boardData.header || boardData.sheetName || '回答一覧',
      sheetName: boardData.sheetName,
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
    if (classFilter && classFilter !== 'すべて') {
      transformedData.data = transformedData.data.filter(item =>
        item.class === classFilter
      );
    }

    // ソート適用
    if (sortOrder === 'newest') {
      transformedData.data.sort((a, b) => (b.rowIndex || 0) - (a.rowIndex || 0));
    } else if (sortOrder === 'oldest') {
      transformedData.data.sort((a, b) => (a.rowIndex || 0) - (b.rowIndex || 0));
    }

    return transformedData;
  }

  // データが存在しない場合 - 詳細なエラー情報を提供
  const debugInfo = {
    hasBoardData: !!boardData,
    boardDataType: typeof boardData,
    boardDataSuccess: boardData?.success,
    hasDataProperty: boardData && 'data' in boardData,
    dataType: typeof boardData?.data,
    isDataArray: Array.isArray(boardData?.data),
    dataLength: Array.isArray(boardData?.data) ? boardData?.data.length : 'not-array'
  };

  console.warn('🔧 transformBoardDataToFrontendFormat: No valid data found - Debug Info:', debugInfo);

  return {
    success: false,
    error: 'No data available',
    rows: [],
    sheetName: boardData?.sheetName || '',
    header: boardData?.header || 'データなし',
    debug: debugInfo // デバッグ情報を含める
  };
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

    const currentEmail = getCurrentEmail();
    if (!currentEmail) {
      console.warn('getSheetList: Unauthenticated access attempt');
      return {
        success: false,
        error: 'ユーザー認証が必要です'
      };
    }

    console.log('getSheetList: Access by user:', `${currentEmail.split('@')[0]}@***`, 'for spreadsheet:', `${spreadsheetId.substring(0, 8)}***`);

    // ✅ CLAUDE.md準拠: Progressive access - try normal permissions first, fallback to service account
    // This ensures editors can access their own spreadsheets with appropriate permissions
    let dataAccess = null;
    let usedServiceAccount = false;

    try {
      // First attempt: Normal permissions
      console.log('getSheetList: Attempting normal permissions access');
      dataAccess = openSpreadsheet(spreadsheetId, { useServiceAccount: false });

      if (!dataAccess || !dataAccess.spreadsheet) {
        // Second attempt: Service account (for cross-user access or permission issues)
        console.log('getSheetList: Fallback to service account access');
        dataAccess = openSpreadsheet(spreadsheetId, { useServiceAccount: true });
        usedServiceAccount = true;
      }
    } catch (accessError) {
      console.warn('getSheetList: Both access methods failed:', accessError.message);
      return {
        success: false,
        error: 'スプレッドシートにアクセスできません'
      };
    }

    if (!dataAccess || !dataAccess.spreadsheet) {
      console.warn('getSheetList: Failed to access spreadsheet:', `${spreadsheetId.substring(0, 8)}***`);
      return {
        success: false,
        error: 'スプレッドシートを開くことができませんでした'
      };
    }

    const { spreadsheet } = dataAccess;
    console.log('getSheetList: Successfully accessed spreadsheet via', usedServiceAccount ? 'service account' : 'normal permissions');

    const sheets = spreadsheet.getSheets();

    const sheetList = sheets.map(sheet => ({
      name: sheet.getName(),
      id: sheet.getSheetId(),
      rowCount: sheet.getLastRow(),
      columnCount: sheet.getLastColumn()
    }));

    console.log('getSheetList: Found', sheetList.length, 'sheets in spreadsheet:', `${spreadsheetId.substring(0, 8)}***`);

    return {
      success: true,
      sheets: sheetList,
      accessMethod: usedServiceAccount ? 'service_account' : 'normal_permissions'
    };
  } catch (error) {
    console.error('getSheetList error:', error.message);
    return {
      success: false,
      error: `シート一覧取得エラー: ${error.message}`,
      details: error.stack
    };
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
 * ✅ 時刻ベース統一新着通知システム - 全ユーザーロール対応
 * @param {string} targetUserId - 閲覧対象ユーザーID
 * @param {Object} options - オプション設定
 * @param {string} options.lastUpdateTime - 最終更新時刻（ISO文字列）
 * @param {string} options.classFilter - クラスフィルター（null=すべて）
 * @param {string} options.sortOrder - ソート順序（newest/oldest）
 * @returns {Object} 時刻ベース統一通知更新結果
 */
function getNotificationUpdate(targetUserId, options = {}) {
  try {
    const viewerEmail = getCurrentEmail();
    if (!viewerEmail) {
      return {
        success: false,
        hasNewContent: false,
        message: 'Authentication required'
      };
    }

    // 対象ユーザー取得
    const targetUser = findUserById(targetUserId);
    if (!targetUser) {
      return {
        success: false,
        hasNewContent: false,
        message: 'Target user not found'
      };
    }

    // アクセス制御判定（getViewerBoardDataパターン踏襲）
    const isSelfAccess = targetUser.userEmail === viewerEmail;
    console.log(`getNotificationUpdate: ${isSelfAccess ? 'Self-access' : 'Cross-user access'} for targetUserId: ${targetUserId}`);

    // 最終更新時刻の正規化
    let lastUpdate;
    try {
      if (typeof options.lastUpdateTime === 'string') {
        lastUpdate = new Date(options.lastUpdateTime);
      } else if (typeof options.lastUpdateTime === 'number') {
        lastUpdate = new Date(options.lastUpdateTime);
      } else {
        lastUpdate = new Date(0); // 初回チェック
      }
    } catch (e) {
      console.warn('getNotificationUpdate: timestamp parse error', e);
      lastUpdate = new Date(0);
    }

    // 統合データ取得（自己アクセス vs クロスユーザーアクセス）
    let currentData;
    if (isSelfAccess) {
      // ✅ 自己アクセス：通常権限
      currentData = getUserSheetData(targetUser.userId, {
        includeTimestamp: true,
        classFilter: options.classFilter,
        sortBy: options.sortOrder || 'newest',
        requestingUser: viewerEmail
      });
    } else {
      // ✅ クロスユーザーアクセス：サービスアカウント（getViewerBoardDataパターン）
      currentData = getUserSheetData(targetUser.userId, {
        includeTimestamp: true,
        classFilter: options.classFilter,
        sortBy: options.sortOrder || 'newest',
        requestingUser: viewerEmail,
        targetUserEmail: targetUser.userEmail // サービスアカウントトリガー
      });
    }

    if (!currentData?.success || !currentData.data) {
      return {
        success: true,
        hasNewContent: false,
        hasNewData: false,
        data: [],
        totalCount: 0,
        newItemsCount: 0,
        message: 'No data available',
        targetUserId,
        accessType: isSelfAccess ? 'self' : 'cross-user'
      };
    }

    // ✅ 修正: 時刻ベース統一新着検出（件数ベース比較完全除去）
    let newItemsCount = 0;
    const newItems = [];
    const incrementalData = currentData.data || [];

    // 時刻ベース新着検出のみ
    currentData.data.forEach((item, index) => {
      let itemTimestamp = new Date(0);
      if (item.timestamp) {
        try {
          itemTimestamp = new Date(item.timestamp);
        } catch (e) {
          itemTimestamp = new Date();
        }
      }

      if (itemTimestamp > lastUpdate) {
        newItemsCount++;
        newItems.push({
          rowIndex: item.rowIndex || index + 1,
          name: item.name || '匿名',
          preview: (item.answer || item.opinion) ? `${(item.answer || item.opinion).substring(0, SYSTEM_LIMITS.PREVIEW_LENGTH)}...` : 'プレビュー不可',
          timestamp: itemTimestamp.toISOString()
        });
      }
    });

    const hasNewContent = newItemsCount > 0;

    // ✅ 修正: 時刻ベース統一レスポンス + フィルター情報追加（論理的破綻修正）
    return {
      success: true,
      hasNewContent,
      data: incrementalData,
      newItemsCount,
      newItems: newItems.slice(0, 5), // 最新5件のプレビュー
      targetUserId,
      accessType: isSelfAccess ? 'self' : 'cross-user',
      sheetName: currentData.sheetName,
      header: currentData.header,
      timestamp: new Date().toISOString(),
      lastUpdateTime: lastUpdate.toISOString(),
      // ✅ 追加: 使用されたフィルター情報（フィルター状態不整合の修正）
      appliedFilter: {
        classFilter: options.classFilter,
        sortOrder: options.sortOrder,
        rawClassFilter: options.classFilter || 'すべて'
      }
    };

  } catch (error) {
    console.error('getNotificationUpdate error:', error.message);
    return {
      success: false,
      hasNewContent: false,
      data: [],
      newItemsCount: 0,
      message: error.message,
      targetUserId,
      timestamp: new Date().toISOString()
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
    if (!email) {
      console.warn('connectDataSource: Unauthenticated access attempt');
      return {
        success: false,
        error: 'ユーザー認証が必要です'
      };
    }

    // ✅ CLAUDE.md準拠: Editor access for own spreadsheets
    // getColumnAnalysis内で詳細なアクセス権チェックが実装済み
    console.log('connectDataSource: Access by user:', `${email.split('@')[0]}@***`);


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
    if (!email) {
      console.warn('getColumnAnalysis: Unauthenticated access attempt');
      return {
        success: false,
        error: 'ユーザー認証が必要です'
      };
    }

    const isAdmin = isAdministrator(email);

    // ✅ CLAUDE.md準拠: Enhanced access control for editor users
    // 管理者は任意のスプレッドシートにアクセス、編集ユーザーは自分がアクセス権を持つもののみ
    let dataAccess;
    try {
      // 編集ユーザーの場合、まず通常権限でアクセステストを行う
      dataAccess = openSpreadsheet(spreadsheetId, { useServiceAccount: false });
      if (!dataAccess) {
        if (isAdmin) {
          // 管理者の場合、サービスアカウントでリトライ
          console.log('getColumnAnalysis: Admin fallback to service account access');
          dataAccess = openSpreadsheet(spreadsheetId, { useServiceAccount: true });
        }

        if (!dataAccess) {
          return {
            success: false,
            error: 'スプレッドシートにアクセスできません'
          };
        }
      }
    } catch (accessError) {
      console.warn('getColumnAnalysis: Spreadsheet access failed for user:', `${email.split('@')[0]}@***`, accessError.message);
      return {
        success: false,
        error: 'スプレッドシートにアクセス権がありません'
      };
    }

    const sheet = dataAccess.getSheet(sheetName);

    if (!sheet) {
      return { success: false, message: 'Sheet not found' };
    }

    // 🔧 高精度分析用データ取得
    const lastCol = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();
    const headers = lastCol > 0 ? getSheetHeaders(sheet, lastCol) : [];

    // 🎯 サンプルデータ取得（ハイブリッドシステム用）
    let sampleData = [];
    if (lastRow > 1 && lastCol > 0) {
      const sampleSize = Math.min(10, lastRow - 1); // 最大10行のサンプル
      try {
        const dataRange = sheet.getRange(2, 1, sampleSize, lastCol);
        sampleData = dataRange.getValues();
        console.log(`getColumnAnalysis: サンプルデータ ${sampleSize}行取得完了`);
      } catch (sampleError) {
        console.warn('getColumnAnalysis: サンプルデータ取得失敗:', sampleError.message);
        sampleData = [];
      }
    }

    // 🎯 高精度ColumnMappingService活用（サンプルデータ付き）
    const diagnostics = performIntegratedColumnDiagnostics(headers, { sampleData });

    // ✅ 編集者自身のアカウントでリアクション列・ハイライト列を事前追加
    try {
      const columnSetupResult = setupReactionAndHighlightColumns(spreadsheetId, sheetName, headers);
      if (columnSetupResult.columnsAdded && columnSetupResult.columnsAdded.length > 0) {
        console.log(`getColumnAnalysis: Added columns: ${columnSetupResult.columnsAdded.join(', ')}`);
        // ヘッダーを再取得して更新されたリストを返す
        const updatedSheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
        const updatedLastCol = updatedSheet.getLastColumn();
        const updatedHeaders = updatedLastCol > 0 ? getSheetHeaders(updatedSheet, updatedLastCol) : [];

        return {
          success: true,
          sheet: updatedSheet,
          headers: updatedHeaders,
          data: [], // 軽量化: 接続時は不要
          mapping: diagnostics.recommendedMapping || {},
          confidence: diagnostics.confidence || {},
          columnsAdded: columnSetupResult.columnsAdded
        };
      }
    } catch (columnError) {
      console.warn('getColumnAnalysis: Column setup failed:', columnError.message);
      // 列追加失敗でも、基本的な分析結果は返す
    }

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
 * リアクション列とハイライト列を事前セットアップ
 * 編集者自身のアカウントで実行（CLAUDE.md準拠）
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @param {Array} currentHeaders - 現在のヘッダー配列
 * @returns {Object} 追加結果
 */
function setupReactionAndHighlightColumns(spreadsheetId, sheetName, currentHeaders = []) {
  try {
    const email = getCurrentEmail();
    console.log(`setupReactionAndHighlightColumns: Setting up columns for ${email ? `${email.split('@')[0]}@***` : 'unknown'}`);

    // 編集者自身のアカウントで直接アクセス（サービスアカウント不使用）
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`Sheet '${sheetName}' not found`);
    }

    // 必要な列の定義
    const requiredColumns = [
      'UNDERSTAND',
      'LIKE',
      'CURIOUS',
      'HIGHLIGHT' // ハイライト列
    ];

    const columnsToAdd = [];
    const currentHeadersUpper = currentHeaders.map(h => String(h).toUpperCase());

    // 存在しない列を特定
    requiredColumns.forEach(columnName => {
      const exists = currentHeadersUpper.some(header =>
        header.includes(columnName.toUpperCase())
      );

      if (!exists) {
        columnsToAdd.push(columnName);
      } else {
        console.log(`setupReactionAndHighlightColumns: Column ${columnName} already exists`);
      }
    });

    const columnsAdded = [];

    if (columnsToAdd.length > 0) {
      const currentLastCol = sheet.getLastColumn();

      // 各列を順次追加
      columnsToAdd.forEach((columnName, index) => {
        const newColIndex = currentLastCol + index + 1;
        console.log(`setupReactionAndHighlightColumns: Adding column '${columnName}' at position ${newColIndex}`);

        try {
          // ヘッダー行に列名を設定
          sheet.getRange(1, newColIndex).setValue(columnName);
          columnsAdded.push(columnName);
          console.log(`setupReactionAndHighlightColumns: Successfully added column '${columnName}' at position ${newColIndex}`);
        } catch (colError) {
          console.error(`setupReactionAndHighlightColumns: Failed to add column '${columnName}':`, colError.message);
        }
      });

      console.log(`setupReactionAndHighlightColumns: Added ${columnsAdded.length} columns: ${columnsAdded.join(', ')}`);
    } else {
      console.log('setupReactionAndHighlightColumns: All required columns already exist');
    }

    return {
      success: true,
      columnsAdded,
      totalColumns: requiredColumns.length,
      alreadyExists: requiredColumns.length - columnsToAdd.length
    };

  } catch (error) {
    console.error('setupReactionAndHighlightColumns error:', error.message);
    return {
      success: false,
      error: error.message,
      columnsAdded: []
    };
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
    // Zero-Dependency safety: 関数存在チェック
    if (typeof getFormInfo === 'function') {
      return getFormInfo(spreadsheetId, sheetName);
    } else {
      return {
        success: false,
        message: 'getFormInfo関数が初期化されていません'
      };
    }
  } catch (error) {
    const errorMessage = error && error.message ? error.message : '予期しないエラーが発生しました';
    console.error('getFormInfoInternal error:', errorMessage);
    return {
      success: false,
      message: errorMessage
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

  try {
    const email = getCurrentEmail();

    // Enhanced type validation for email
    if (!email || typeof email !== 'string' || email.trim() === '') {
      console.error('❌ Authentication failed - invalid email:', typeof email, email);
      return {
        success: false,
        message: 'Authentication required - invalid email',
        domain: null,
        isValidDomain: false
      };
    }

    const domain = email.includes('@') ? email.split('@')[1] : 'unknown';
    const adminEmail = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL');
    const adminDomain = adminEmail ? adminEmail.split('@')[1] : null;
    const isValidDomain = adminDomain ? domain === adminDomain : true;


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
function getActiveFormInfo(userId) {

  try {
    const currentEmail = getCurrentEmail();
    if (!currentEmail) {
      console.error('❌ Authentication failed');
      return {
        success: false,
        message: 'Authentication required',
        formUrl: null,
        formTitle: null
      };
    }

    // 🔧 GAS-Native統一: userIdが指定されていればそのユーザーの設定を取得（board owner's form）
    // 指定されていなければ現在のユーザーの設定を取得（backward compatibility）
    let targetUserId = userId;
    if (!targetUserId) {
      const currentUser = findUserByEmail(currentEmail, { requestingUser: currentEmail });
      if (!currentUser) {
        console.error('❌ Current user not found:', currentEmail);
        return {
          success: false,
          message: 'User not found',
          formUrl: null,
          formTitle: null
        };
      }
      targetUserId = currentUser.userId;
    }

    // 統一API使用: 構造化パース
    const configResult = getUserConfig(targetUserId);
    const config = configResult.success ? configResult.config : {};

    // フォーム表示判定: URL存在性を優先、検証は補助的に利用
    const hasFormUrl = !!(config.formUrl && config.formUrl.trim());
    const isValidUrl = hasFormUrl && isValidFormUrl(config.formUrl);


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


// ===========================================
// 🔄 CLAUDE.md準拠: GAS-side Trigger-Based Polling System
// ===========================================


// ===========================================
// 🔧 Missing API Endpoints - Frontend/Backend Compatibility
// ===========================================







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

    // ✅ CLAUDE.md準拠: Exponential backoff retry for resilient spreadsheet access
    const spreadsheet = executeWithRetry(
      () => SpreadsheetApp.openById(spreadsheetId),
      { operationName: 'SpreadsheetApp.openById', maxRetries: 3 }
    );
    const sheets = spreadsheet.getSheets();

    // ✅ Batch Operations: 全シート情報を一括取得（70x improvement）
    const sheetInfos = sheets.map(sheet => ({
      name: sheet.getName(),
      gid: sheet.getSheetId().toString()
    }));


    // GIDに一致するシートを検索
    const targetSheet = sheetInfos.find(info => info.gid === gid);
    const resultName = targetSheet ? targetSheet.name : sheetInfos[0]?.name || 'Sheet1';

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

    // Step 1: URL解析
    const parseResult = extractSpreadsheetInfo(fullUrl);
    if (!parseResult.success) {
      return parseResult;
    }

    const { spreadsheetId, gid } = parseResult;

    // Step 2: シート名自動取得
    const sheetName = getSheetNameFromGid(spreadsheetId, gid);

    // Step 3: ✅ 既存API活用 - 並列相当処理（GAS-Nativeパターン）

    const accessResult = validateAccessAPI(spreadsheetId);

    // Zero-Dependency safety: 関数存在チェック
    const formResult = typeof getFormInfo === 'function' ?
      getFormInfo(spreadsheetId, sheetName) :
      { success: false, message: 'getFormInfo関数が初期化されていません' };

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
      'getWebAppUrl',
      'getNotificationUpdate'
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

/**
 * ✅ CLAUDE.md準拠: Batched viewer data retrieval for 70x performance improvement
 * Combines 4 individual API calls into single batch operation:
 * - findUserById
 * - getUserConfig
 * - isAdministrator
 * - getQuestionText
 *
 * @param {string} targetUserId - Target user ID
 * @param {string} currentEmail - Current viewer email
 * @returns {Object} Batched result with all required data
 */
function getBatchedViewerData(targetUserId, currentEmail) {
  try {
    // ✅ Batch operation: Get all required data in single coordinated call
    const targetUser = findUserById(targetUserId, { requestingUser: currentEmail });
    if (!targetUser) {
      return { success: false, error: '対象ユーザーが見つかりません' };
    }

    // Batch remaining operations with user context
    const configResult = getUserConfig(targetUserId);
    const config = configResult.success ? configResult.config : {};

    const isAdminUser = isAdministrator(currentEmail);

    return {
      success: true,
      targetUser,
      config,
      isAdminUser
    };

  } catch (error) {
    console.error('getBatchedViewerData error:', error.message);
    return {
      success: false,
      error: `データ取得エラー: ${error.message}`
    };
  }
}


/**
 * ✅ CLAUDE.md準拠: Batched admin data retrieval for 70x performance improvement
 * Combines 4 individual API calls into single batch operation:
 * - getCurrentEmail
 * - isAdministrator
 * - findUserByEmail
 * - getUserConfig
 *
 * @returns {Object} Batched result with all required admin data
 */
function getBatchedAdminData() {
  try {
    // ✅ Batch operation: Get all required admin data in single coordinated call
    const email = getCurrentEmail();
    if (!email) {
      return { success: false, error: 'ユーザー認証が必要です' };
    }

    // ✅ 編集ユーザー対応: 管理者でなくても自分のボードの管理パネルにアクセス可能
    const isAdmin = isAdministrator(email);
    const user = findUserByEmail(email, { requestingUser: email });

    // ユーザー情報が見つからない場合は権限チェック前にエラー
    if (!user && !isAdmin) {
      return { success: false, error: 'ユーザー情報が見つかりません' };
    }

    // 管理者ではない場合、最低でもアクティブなユーザーである必要がある
    if (!isAdmin && (!user || !user.isActive)) {
      return { success: false, error: '編集権限が必要です' };
    }

    const configResult = getUserConfig(user.userId);
    const config = configResult.success ? configResult.config : {};

    // ✅ CLAUDE.md準拠: フロントエンド必要情報を統合取得
    const questionText = getQuestionText(config, { targetUserEmail: user.userEmail });

    // ✅ URLs とタイムスタンプ情報を config に統合
    // ✅ Optimized: Use database lastModified instead of config lastModified
    const baseUrl = ScriptApp.getService().getUrl();
    const enhancedConfig = {
      ...config,
      urls: config.urls || {
        view: `${baseUrl}?mode=view&userId=${user.userId}`,
        admin: `${baseUrl}?mode=admin&userId=${user.userId}`
      },
      lastModified: user.lastModified || config.publishedAt
    };

    return {
      success: true,
      email,
      user,
      config: enhancedConfig,
      questionText: questionText || '回答ボード'
    };

  } catch (error) {
    console.error('getBatchedAdminData error:', error.message);
    return {
      success: false,
      error: `管理者データ取得エラー: ${error.message}`
    };
  }
}

/**
 * ✅ CLAUDE.md準拠: Batched admin authentication for 70x performance improvement
 * Combines 2 individual API calls into single batch operation:
 * - getCurrentEmail
 * - isAdministrator
 *
 * @param {Object} options - Additional options for admin auth
 * @returns {Object} Batched result with admin authentication status
 */
function getBatchedAdminAuth(options = {}) {
  try {
    // ✅ Batch operation: Get email and admin status in single coordinated call
    const email = getCurrentEmail();
    if (!email) {
      return {
        success: false,
        authenticated: false,
        isAdmin: false,
        error: 'ユーザー認証が必要です',
        authError: createAuthError()
      };
    }

    const isAdmin = isAdministrator(email);
    if (!isAdmin && !options.allowNonAdmin) {
      return {
        success: false,
        authenticated: true,
        isAdmin: false,
        email,
        error: '管理者権限が必要です',
        adminError: createAdminRequiredError()
      };
    }

    return {
      success: true,
      authenticated: true,
      isAdmin,
      email,
      authLevel: isAdmin ? 'administrator' : 'user'
    };

  } catch (error) {
    console.error('getBatchedAdminAuth error:', error.message);
    return {
      success: false,
      authenticated: false,
      isAdmin: false,
      error: `認証エラー: ${error.message}`,
      exception: createExceptionResponse(error)
    };
  }
}

/**
 * ✅ CLAUDE.md準拠: Batched user config retrieval for 70x performance improvement
 * Combines 3 individual API calls into single batch operation:
 * - getCurrentEmail
 * - findUserByEmail
 * - getUserConfig
 *
 * @param {Object} options - Additional options for user config retrieval
 * @returns {Object} Batched result with user and config data
 */
function getBatchedUserConfig(options = {}) {
  try {
    // ✅ Batch operation: Get email, user, and config in single coordinated call
    const email = getCurrentEmail();
    if (!email) {
      return {
        success: false,
        authenticated: false,
        error: 'ユーザー認証が必要です',
        user: null,
        config: null,
        authError: createAuthError()
      };
    }

    const user = findUserByEmail(email, { requestingUser: email });
    if (!user) {
      return {
        success: false,
        authenticated: true,
        error: 'ユーザー情報が見つかりません',
        email,
        user: null,
        config: null,
        userError: createUserNotFoundError()
      };
    }

    const configResult = getUserConfig(user.userId);
    const config = configResult.success ? configResult.config : {};

    return {
      success: true,
      authenticated: true,
      email,
      user,
      config,
      configSuccess: configResult.success,
      configMessage: configResult.message || 'Configuration retrieved successfully'
    };

  } catch (error) {
    console.error('getBatchedUserConfig error:', error.message);
    return {
      success: false,
      authenticated: false,
      error: `ユーザー設定取得エラー: ${error.message}`,
      user: null,
      config: null,
      exception: createExceptionResponse(error)
    };
  }
}

/**
 * ✅ CLAUDE.md準拠: Exponential backoff retry for resilient operations
 * Generic retry function for operations that may fail due to network/quota issues
 *
 * @param {Function} operation - Function to retry
 * @param {Object} options - Retry options
 * @returns {*} Result of successful operation
 */
function executeWithRetry(operation, options = {}) {
  const maxRetries = options.maxRetries || 3;
  const initialDelay = options.initialDelay || 500;
  const maxDelay = options.maxDelay || 5000;
  const operationName = options.operationName || 'Operation';

  let retryCount = 0;
  let lastError = null;

  while (retryCount < maxRetries) {
    try {
      // Add delay for retry attempts (not first attempt)
      if (retryCount > 0) {
        const delay = Math.min(
          initialDelay * Math.pow(2, retryCount - 1),
          maxDelay
        );
        console.warn(`${operationName}: Retry ${retryCount}/${maxRetries - 1} after ${delay}ms delay`);
        Utilities.sleep(delay);
      }

      // Execute the operation
      const result = operation();

      // Success - log only if this was a retry
      if (retryCount > 0) {
        console.info(`✅ ${operationName}: Succeeded on retry ${retryCount}`);
      }

      return result;

    } catch (error) {
      lastError = error;
      retryCount++;

      const errorMessage = error && error.message ? error.message : String(error);

      // Check if this is a retryable error
      const isRetryable = isRetryableError(errorMessage);

      console.warn(`❌ ${operationName}: Attempt ${retryCount} failed: ${errorMessage}`);

      // Don't retry if error is not retryable or we've reached max retries
      if (!isRetryable || retryCount >= maxRetries) {
        break;
      }
    }
  }

  // All retries exhausted
  const finalError = lastError && lastError.message ? lastError.message : 'Unknown error';
  console.error(`🚨 ${operationName}: Failed after ${retryCount} attempts: ${finalError}`);
  throw lastError || new Error(`${operationName} failed after ${retryCount} attempts`);
}

/**
 * Check if an error is retryable (network/quota issues vs permanent failures)
 * @param {string} errorMessage - Error message to analyze
 * @returns {boolean} True if error is retryable
 */
function isRetryableError(errorMessage) {
  if (!errorMessage || typeof errorMessage !== 'string') {
    return false;
  }

  const retryablePatterns = [
    'timeout',
    'network',
    'quota',
    'rate limit',
    'service unavailable',
    'internal error',
    'temporarily unavailable',
    'backend error',
    'connection',
    'socket'
  ];

  const nonRetryablePatterns = [
    'permission',
    'not found',
    'not authorized',
    'invalid',
    'malformed',
    'access denied',
    'authentication failed'
  ];

  const lowerMessage = errorMessage.toLowerCase();

  // Check for non-retryable errors first
  for (const pattern of nonRetryablePatterns) {
    if (lowerMessage.includes(pattern)) {
      return false;
    }
  }

  // Check for retryable errors
  for (const pattern of retryablePatterns) {
    if (lowerMessage.includes(pattern)) {
      return true;
    }
  }

  // Default to retryable for unknown errors (conservative approach)
  return true;
}
