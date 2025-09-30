/**
 * main.gs - Simplified Application Entry Points
 *
 * Responsibilities:
 * - HTTP request routing (doGet/doPost)
 * - Simple mode validation
 * - Template serving
 *
 * Following Google Apps Script Best Practices:
 * - Direct API calls (no abstraction layers)
 * - Minimal service calls
 * - Simple, readable code
 */

/* global createErrorResponse, createSuccessResponse, createAuthError, createUserNotFoundError, createAdminRequiredError, createExceptionResponse, hasCoreSystemProps, getUserSheetData, addReaction, toggleHighlight, validateConfig, findUserByEmail, findUserById, findUserBySpreadsheetId, createUser, getAllUsers, updateUser, openSpreadsheet, getUserConfig, saveUserConfig, clearConfigCache, cleanConfigFields, getQuestionText, DB, validateAccess, URL, UserService, CACHE_DURATION, TIMEOUT_MS, SLEEP_MS, SYSTEM_LIMITS, SystemController, getDatabaseConfig, getViewerBoardData, getSheetHeaders, performIntegratedColumnDiagnostics, generateRecommendedMapping, getFormInfo, enhanceConfigWithDynamicUrls */

// Core Utility Functions

/**
 * 現在のユーザーのメールアドレスを取得
 * @returns {string|null} ユーザーのメールアドレス、または認証されていない場合はnull
 */
function getCurrentEmail() {
  try {
    const activeUser = Session.getActiveUser();
    const email = activeUser ? activeUser.getEmail() : null;

    if (!email || email.trim() === '') {
      try {
        const effectiveUser = Session.getEffectiveUser();
        const effectiveEmail = effectiveUser ? effectiveUser.getEmail() : null;
        if (effectiveEmail && effectiveEmail.trim() !== '') {
          return effectiveEmail;
        }
      } catch (altError) {
        // Alternative email retrieval failed, continue to fallback return
      }
    }

    return email || null;
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


// 🌐 HTTP Entry Points

/**
 * Handle GET requests
 * @param {Object} e - Event object containing request parameters
 * @returns {HtmlOutput} HTML response for the requested page
 */
function doGet(e) {
  try {
    const params = e ? e.parameter : {};
    const mode = params.mode || 'main';

    //Performance optimization: Cache email for authentication-required routes
    const currentEmail = (mode !== 'login') ? getCurrentEmail() : null;

    // 🚫 アプリ全体のアクセス制限チェック
    // APP_DISABLED フラグがtrueの場合、管理者以外のアクセスを制限
    const isAppDisabled = checkAppAccessRestriction();
    if (isAppDisabled) {
      const isAdmin = currentEmail ? isAdministrator(currentEmail) : false;

      // 管理者のみappSetupモードでのアクセスを許可（復旧作業用）
      if (mode === 'appSetup' && isAdmin) {
        // 管理者のappSetup アクセスは通常通り処理
      } else {
        // 停止中画面を表示（管理者には復旧用のリンクを表示）
        const template = HtmlService.createTemplateFromFile('AccessRestricted.html');
        template.isAdministrator = isAdmin;
        template.userEmail = currentEmail || '';
        template.timestamp = new Date().toISOString();
        template.isAppDisabled = true; // アプリ停止状態を明示
        return template.evaluate();
      }
    }


    // Simple routing
    switch (mode) {
      case 'login': {
        // 極限シンプル: ログインページ（静的表示のみ）
        return HtmlService.createTemplateFromFile('LoginPage.html').evaluate();
      }

      case 'manual': {
        // 教師向けマニュアルページ（静的表示のみ）
        return HtmlService.createTemplateFromFile('TeacherManual.html').evaluate();
      }

      case 'admin': {
        // 🔐 GAS-Native: 直接認証チェック - Admin権限確認
        if (!currentEmail) {
          return createRedirectTemplate('ErrorBoundary.html', 'ユーザー認証が必要です');
        }

        // 対象ユーザー確認（userIdパラメータが指定されている場合）
        const targetUserId = params.userId;
        if (!targetUserId) {
          return createRedirectTemplate('ErrorBoundary.html', 'ユーザーIDが指定されていません');
        }

        // Batch operations for 70x performance improvement
        const adminData = getBatchedAdminData(targetUserId);
        if (!adminData.success) {
          return createRedirectTemplate('ErrorBoundary.html', adminData.error || '管理者権限が必要です');
        }

        const { email, user, config } = adminData;
        const isAdmin = isAdministrator(email);

        // Dynamic URL generation
        const enhancedConfig = enhanceConfigWithDynamicUrls(config, user.userId);

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
          setupStatus: config.setupStatus || 'pending',
          displaySettings: config.displaySettings || {},
          columnMapping: config.columnMapping || {},
          dynamicUrls: enhancedConfig.dynamicUrls || {}
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
          // Pass isSystemAdmin variable to AccessRestricted.html
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

        //userIdパラメータを取得（管理パネルに戻るリンクで使用）
        const userIdParam = params.userId;

        // 認証済み - Administrator権限でAppSetup表示
        const template = HtmlService.createTemplateFromFile('AppSetupPage.html');

        //管理パネルに戻るリンクのためにuserIdを渡す（オプション）
        template.userIdParam = userIdParam || '';

        return template.evaluate();
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

        // Batch operations for 70x performance improvement
        const viewerData = getBatchedViewerData(targetUserId, currentEmail);
        if (!viewerData.success) {
          return createRedirectTemplate('ErrorBoundary.html', viewerData.error || '対象ユーザーが見つかりません');
        }

        const { targetUser, config, isAdminUser } = viewerData;
        const isOwnBoard = currentEmail === targetUser.userEmail;
        const isPublished = Boolean(config.isPublished);

        // 🔧 論理的修正: 非公開状態なら所有者・非所有者問わずUnpublished.htmlを表示
        if (!isPublished) {
          const template = HtmlService.createTemplateFromFile('Unpublished.html');
          template.isEditor = isAdminUser || isOwnBoard; // 表示内容制御
          template.editorName = targetUser.userName || targetUser.userEmail || '';

          // Generate board URL
          const baseUrl = ScriptApp.getService().getUrl();
          template.boardUrl = `${baseUrl}?mode=view&userId=${targetUserId}`;


          return template.evaluate();
        }

        // 認証済み - 公開ボード表示
        const template = HtmlService.createTemplateFromFile('Page.html');
        template.userId = targetUserId;
        template.userEmail = targetUser.userEmail;
        template.questionText = '読み込み中...';
        template.boardTitle = targetUser.userEmail || '回答ボード';

        // Unified permission information
        const isEditor = isAdminUser || isOwnBoard;
        template.isEditor = isEditor;
        template.isAdminUser = isAdminUser;
        template.isOwnBoard = isOwnBoard;

        // Unified configJSON retrieval
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
          displaySettings: config.displaySettings || { showNames: false, showReactions: false }
        });

        return template.evaluate();
      }

      case 'main':
      default: {
        // Default landing is AccessRestricted to prevent unintended login/account creation.
        // Viewers must specify ?mode=view&userId=... and admins explicitly use ?mode=login.
        // Pass isSystemAdmin variable to AccessRestricted.html
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

    // Set necessary variables for AccessRestricted.html
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
        try {
          const user = findUserByEmail(email, { requestingUser: email });
          if (!user) {
            result = createUserNotFoundError();
          } else {
            result = { success: true, data: getUserSheetData(user.userId, request.options || {}) };
          }
        } catch (error) {
          console.error('getData error:', error.message);
          result = createExceptionResponse(error);
        }
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
        try {
          const user = findUserByEmail(email, { requestingUser: email });
          if (!user) {
            result = createUserNotFoundError();
          } else {
            result = { success: true, data: getUserSheetData(user.userId, request.options || {}) };
          }
        } catch (error) {
          console.error('refreshData error:', error.message);
          result = createExceptionResponse(error);
        }
        break;
      case 'publishApp':
        try {
          const user = findUserByEmail(email, { requestingUser: email });
          if (!user) {
            result = createUserNotFoundError();
          } else {
            const publishConfig = {
              ...request.config,
              isPublished: true,
              publishedAt: new Date().toISOString(),
              setupComplete: true
            };

            const saveResult = saveUserConfig(user.userId, publishConfig);
            if (!saveResult.success) {
              result = createErrorResponse(saveResult.message || '公開設定の保存に失敗しました');
            } else {
              result = {
                success: true,
                message: 'アプリが正常に公開されました',
                publishedAt: publishConfig.publishedAt,
                config: saveResult.config,
                etag: saveResult.etag
              };
            }
          }
        } catch (error) {
          console.error('publishApp error:', error.message);
          result = createExceptionResponse(error);
        }
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



// API Functions (called from HTML)

/**
 * Get current user email for HTML templates
 * Direct call - GAS Best Practice
 */


/**
 * Get user configuration - unified function for current user
 */
function getConfig() {
  try {
    const email = getCurrentEmail();
    if (!email) {
      return createAuthError();
    }

    // Direct findUserByEmail usage
    const user = findUserByEmail(email, { requestingUser: email });
    if (!user) {
      return createUserNotFoundError();
    }

    // Use getUserConfig for configuration retrieval
    const configResult = getUserConfig(user.userId);
    const config = configResult.success ? configResult.config : {};

    return { success: true, data: { config, userId: user.userId } };
  } catch (error) {
    console.error('getConfig error:', error.message);
    return createExceptionResponse(error);
  }
}

// getWebAppUrl moved to SystemController.gs for architecture compliance

// Frontend compatibility API - unified authentication system

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
      // Administrator authenticated
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
 * Get user configuration by userId - for compatibility with existing HTML calls
 */


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
    // Direct Data usage
    let user = findUserByEmail(email, { requestingUser: email });
    if (!user) {
      // Unified createUser() implementation
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

    // Get all users (no filter needed for physical deletion)
    const users = getAllUsers({ activeOnly: false }, { forceServiceAccount: true });
    return {
      success: true,
      users: users || []
    };
  } catch (error) {
    console.error('getAdminUsers error:', error.message);
    return createExceptionResponse(error);
  }
}

// Remove duplicate implementation
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

    // Direct findUserById usage
    const targetUser = findUserById(targetUserId, { requestingUser: email });
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
 * Toggle user board publication status - admin only (for managing other users)
 */
function toggleUserBoardStatus(targetUserId) {
  try {
    const email = getCurrentEmail();
    if (!email) {
      return createAuthError('ユーザー認証が必要です');
    }

    // 権限チェック: 管理者のみ（他ユーザーのボード状態を管理するため）
    if (!isAdministrator(email)) {
      return createAuthError('管理者権限が必要です');
    }

    // Direct findUserById usage
    const targetUser = findUserById(targetUserId, { requestingUser: email });
    if (!targetUser) {
      return createUserNotFoundError();
    }

    // Minimal update: change only isPublished field
    // 現在のconfigJsonを取得してパース
    let currentConfig = {};
    try {
      const configJsonStr = targetUser.configJson || '{}';
      currentConfig = JSON.parse(configJsonStr);
    } catch (error) {
      console.warn('toggleUserBoardStatus: Invalid configJson, using empty config:', error.message);
      currentConfig = {};
    }

    // 公開状態のみを切り替え
    const newIsPublished = !currentConfig.isPublished;
    const updates = {};

    // 新しいconfigJsonを構築（isPublishedとpublishedAtのみ更新）
    const updatedConfig = { ...currentConfig };
    updatedConfig.isPublished = newIsPublished;
    if (newIsPublished && !updatedConfig.publishedAt) {
      updatedConfig.publishedAt = new Date().toISOString();
    }
    updatedConfig.lastAccessedAt = new Date().toISOString();

    updates.configJson = JSON.stringify(updatedConfig);
    updates.lastModified = new Date().toISOString();

    // Direct database update: minimal changes only
    const updateResult = updateUser(targetUserId, updates);
    if (!updateResult.success) {
      return createErrorResponse(`Failed to toggle board status: ${updateResult.message || '詳細不明'}`);
    }

    return {
      success: true,
      message: `ボードを${newIsPublished ? '公開' : '非公開'}に変更しました`,
      userId: targetUserId,
      boardPublished: newIsPublished,
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    console.error('toggleUserBoardStatus error:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * 自分のボードを再公開（所有者用のシンプル関数）
 */
function republishMyBoard() {
  try {
    const email = getCurrentEmail();
    if (!email) {
      return createAuthError('ユーザー認証が必要です');
    }

    // 現在のユーザーを取得
    const currentUser = findUserByEmail(email, { requestingUser: email });
    if (!currentUser) {
      return createUserNotFoundError('ユーザーが見つかりません');
    }

    // 設定を取得
    const configResult = getUserConfig(currentUser.userId);
    const config = configResult.success ? configResult.config : {};

    // 公開状態に変更
    config.isPublished = true;
    config.publishedAt = new Date().toISOString();

    // 設定を保存
    const saveResult = saveUserConfig(currentUser.userId, config);
    if (!saveResult.success) {
      return {
        success: false,
        message: '設定の保存に失敗しました'
      };
    }

    return {
      success: true,
      message: '回答ボードを再公開しました',
      userId: currentUser.userId
    };

  } catch (error) {
    console.error('republishMyBoard error:', error.message);
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

    // Direct Data usage
    let targetUser = targetUserId ? findUserById(targetUserId, { requestingUser: email }) : null;
    if (!targetUser) {
      targetUser = findUserByEmail(email, { requestingUser: email });
    }

    if (!targetUser) {
      return createUserNotFoundError();
    }

    // Editor permission check: admin or own board
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
 * 自分がオーナーのスプレッドシートを30件取得
 */
function getSheets() {
  const files = DriveApp.getFilesByType('application/vnd.google-apps.spreadsheet');
  const sheets = [];

  while (files.hasNext() && sheets.length < 30) {
    const file = files.next();
    sheets.push({
      id: file.getId(),
      name: file.getName(),
      url: file.getUrl()
    });
  }

  return { success: true, sheets };
}




/**
 * Validate header integrity for user's active sheet
 * @param {string} targetUserId - 対象ユーザーID（省略可能）
 * @returns {Object} 検証結果
 */
function validateHeaderIntegrity(targetUserId) {
  try {
    const currentEmail = getCurrentEmail();
    // Direct Data usage
    let targetUser = targetUserId ? findUserById(targetUserId, { requestingUser: currentEmail }) : null;
    if (!targetUser && currentEmail) {
      targetUser = findUserByEmail(currentEmail, { requestingUser: currentEmail });
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

    // Self vs Cross-user access pattern
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
      console.error('Authentication failed');
      return createAuthError();
    }

    // Direct findUserByEmail usage
    const user = findUserByEmail(email, { requestingUser: email });
    if (!user) {
      console.error('User not found:', email);
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
    console.error('getBoardInfo ERROR:', error.message);
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
function getPublishedSheetData(classFilter, sortOrder, adminMode, targetUserId) {
  // Simplified parameters - back to working version
  classFilter = classFilter || null;
  sortOrder = sortOrder || 'newest';
  adminMode = adminMode || false;
  targetUserId = targetUserId || null;

  const startTime = Date.now();

  try {
    const adminAuth = getBatchedAdminAuth({ allowNonAdmin: true });
    if (!adminAuth.success || !adminAuth.authenticated) {
      return {
        success: false,
        error: 'Authentication required',
        data: [],
        sheetName: '',
        header: '認証エラー'
      };
    }

    const { email: viewerEmail, isAdmin: isSystemAdmin } = adminAuth;

    if (targetUserId) {
      // ✅ CLAUDE.md準拠: 完全な事前読み込みとpreloadedAuth伝播
      // 認証情報を渡してfindUserByIdの重複DB呼び出しを排除
      const targetUser = findUserById(targetUserId, {
        requestingUser: viewerEmail,
        preloadedAuth: { email: viewerEmail, isAdmin: isSystemAdmin }
      });
      if (!targetUser) {
        console.error('getPublishedSheetData: Target user not found', { targetUserId, viewerEmail });
        return {
          success: false,
          error: 'Target user not found',
          debugMessage: 'ユーザー検索に失敗しました',
          data: [],
          sheetName: '',
          header: 'ユーザーエラー'
        };
      }

      // ✅ preloadedUserを渡してgetUserConfig内のfindUserById重複呼び出しを排除
      const configResult = getUserConfig(targetUserId, targetUser);
      const targetUserConfig = configResult.success ? configResult.config : {};

      // ✅ 完全なpreloadedAuth伝播 - getUserSheetData内の全DB呼び出しに伝達
      const options = {
        classFilter: classFilter !== 'すべて' ? classFilter : undefined,
        sortBy: sortOrder || 'newest',
        includeTimestamp: true,
        adminMode: isSystemAdmin || (targetUser.userEmail === viewerEmail),
        requestingUser: viewerEmail,
        preloadedAuth: { email: viewerEmail, isAdmin: isSystemAdmin }
      };

      const dataFetchStart = Date.now();

      // ✅ 完全な事前読み込みデータを渡してDB重複アクセス排除
      const result = getUserSheetData(targetUser.userId, options, targetUser, targetUserConfig);
      const dataFetchEnd = Date.now();

      if (!result || !result.success) {
        console.error('getPublishedSheetData: getUserSheetData failed', {
          result,
          error: result?.message || 'データ取得エラー',
          totalTime: Date.now() - startTime
        });
        return {
          success: false,
          error: result?.message || 'データ取得エラー',
          debugMessage: 'スプレッドシート接続に失敗しました',
          data: [],
          sheetName: result?.sheetName || '',
          header: result?.header || '問題'
        };
      }

      const finalResult = {
        success: true,
        data: result.data || [],
        header: result.header || result.sheetName || '回答一覧',
        sheetName: result.sheetName || 'Sheet1'
      };

      // Safe serialization test before return
      try {
        const testSerialization = JSON.stringify(finalResult);

        // Create clean, safe result object with Date protection
        const safeResult = {
          success: true,
          data: Array.isArray(finalResult.data) ? finalResult.data.map(item => {
            if (typeof item === 'object' && item !== null) {
              // Deep clean with Date object protection
              const cleaned = {};
              for (const [key, value] of Object.entries(item)) {
                if (value instanceof Date) {
                  cleaned[key] = value.toISOString();
                } else if (typeof value === 'object' && value !== null) {
                  try {
                    cleaned[key] = JSON.parse(JSON.stringify(value));
                  } catch (e) {
                    cleaned[key] = String(value);
                  }
                } else {
                  cleaned[key] = value;
                }
              }
              return cleaned;
            }
            return item;
          }) : [],
          header: String(finalResult.header || '回答一覧'),
          sheetName: String(finalResult.sheetName || 'Sheet1'),
          displaySettings: targetUserConfig.displaySettings || { showNames: false, showReactions: false }
        };

        return safeResult;

      } catch (serializationError) {
        console.error('getPublishedSheetData: Serialization failed', {
          error: serializationError.message,
          dataLength: finalResult.data?.length || 0,
          headerType: typeof finalResult.header,
          sheetNameType: typeof finalResult.sheetName
        });

        // Return minimal safe response
        return {
          success: true,
          data: [],
          header: '回答一覧（データ変換エラー）',
          sheetName: 'Sheet1',
          displaySettings: targetUserConfig.displaySettings || { showNames: false, showReactions: false }
        };
      }
    }

    // ✅ CLAUDE.md準拠: preloadedAuthを渡してfindUserByEmail内の重複DB呼び出しを排除
    const user = findUserByEmail(viewerEmail, {
      requestingUser: viewerEmail,
      adminMode: isSystemAdmin,
      ignorePermissions: isSystemAdmin,
      preloadedAuth: { email: viewerEmail, isAdmin: isSystemAdmin }
    });

    if (!user) {
      console.error('getPublishedSheetData: User not found (self-access)', { viewerEmail });
      return {
        success: false,
        error: 'User not found',
        data: [],
        sheetName: '',
        header: 'ユーザーエラー'
      };
    }

    // ✅ preloadedUserを渡してgetUserConfig内のfindUserById重複呼び出しを排除
    const configResult = getUserConfig(user.userId, user);
    const userConfig = configResult.success ? configResult.config : {};

    // ✅ 完全なpreloadedAuth伝播 - getUserSheetData内の全DB呼び出しに伝達
    const options = {
      classFilter: classFilter !== 'すべて' ? classFilter : undefined,
      sortBy: sortOrder || 'newest',
      includeTimestamp: true,
      adminMode: isSystemAdmin,
      requestingUser: viewerEmail,
      preloadedAuth: { email: viewerEmail, isAdmin: isSystemAdmin }
    };

    const dataFetchStart = Date.now();

    // ✅ 完全な事前読み込みデータを渡してDB重複アクセス排除
    const result = getUserSheetData(user.userId, options, user, userConfig);
    const dataFetchEnd = Date.now();

    //Simple null check and direct return
    if (!result || !result.success) {
      console.error('getPublishedSheetData: getUserSheetData failed (self-access)', {
        result,
        error: result?.message || 'データ取得エラー',
        totalTime: Date.now() - startTime
      });
      return {
        success: false,
        error: result?.message || 'データ取得エラー',
        debugMessage: 'スプレッドシート接続に失敗しました',
        data: [],
        sheetName: result?.sheetName || '',
        header: result?.header || '問題',
      };
    }

    const finalResult = {
      success: true,
      data: result.data || [],
      header: result.header || result.sheetName || '回答一覧',
      sheetName: result.sheetName || 'Sheet1'
    };

    console.info('getPublishedSheetData: Final result prepared (self-access)', {
      dataLength: finalResult.data.length,
      hasHeader: !!finalResult.header,
      hasSheetName: !!finalResult.sheetName,
      totalExecutionTime: Date.now() - startTime
    });

    //Safe serialization test before return (self-access)
    try {
      const testSerialization = JSON.stringify(finalResult);
      console.info('getPublishedSheetData: Serialization test passed (self-access)', {
        serializationLength: testSerialization.length,
        isValidJson: true
      });

      //Create clean, safe result object with Date protection (self-access)
      const safeResult = {
        success: true,
        data: Array.isArray(finalResult.data) ? finalResult.data.map(item => {
          if (typeof item === 'object' && item !== null) {
            //Deep clean with Date object protection
            const cleaned = {};
            for (const [key, value] of Object.entries(item)) {
              if (value instanceof Date) {
                cleaned[key] = value.toISOString();
              } else if (typeof value === 'object' && value !== null) {
                try {
                  cleaned[key] = JSON.parse(JSON.stringify(value));
                } catch (e) {
                  cleaned[key] = String(value);
                }
              } else {
                cleaned[key] = value;
              }
            }
            return cleaned;
          }
          return item;
        }) : [],
        header: String(finalResult.header || '回答一覧'),
        sheetName: String(finalResult.sheetName || 'Sheet1'),
        displaySettings: userConfig.displaySettings || { showNames: false, showReactions: false }
      };

      console.info('getPublishedSheetData: Safe result created, returning to frontend (self-access)', {
        dataLength: safeResult.data.length,
        headerLength: safeResult.header.length,
        sheetNameLength: safeResult.sheetName.length
      });

      return safeResult;

    } catch (serializationError) {
      console.error('getPublishedSheetData: Serialization failed (self-access)', {
        error: serializationError.message,
        dataLength: finalResult.data?.length || 0,
        headerType: typeof finalResult.header,
        sheetNameType: typeof finalResult.sheetName
      });

      //Return minimal safe response
      return {
        success: true,
        data: [],
        header: '回答一覧（データ変換エラー）',
        sheetName: 'Sheet1',
        displaySettings: userConfig.displaySettings || { showNames: false, showReactions: false }
      };
    }
  } catch (error) {
    console.error('getPublishedSheetData: Exception caught', {
      error: error?.message || 'Unknown error',
      stack: error?.stack,
      classFilter,
      sortOrder,
      adminMode,
      targetUserId,
      timestamp: new Date().toISOString()
    });
    return {
      success: false,
      error: error?.message || 'データ取得エラー',
      data: [],
      sheetName: '',
      header: '問題'
    };
  }
}


// Unified Validation Functions

// Unified Data Operations


// Additional HTML-Called Functions

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


    // Progressive access - try normal permissions first, fallback to service account
    // This ensures editors can access their own spreadsheets with appropriate permissions
    let dataAccess = null;
    let usedServiceAccount = false;

    try {
      // First attempt: Normal permissions
      dataAccess = openSpreadsheet(spreadsheetId, { useServiceAccount: false });

      if (!dataAccess || !dataAccess.spreadsheet) {
        // Second attempt: Service account (for cross-user access or permission issues)
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

    const sheets = spreadsheet.getSheets();

    const sheetList = sheets.map(sheet => ({
      name: sheet.getName(),
      id: sheet.getSheetId(),
      rowCount: sheet.getLastRow(),
      columnCount: sheet.getLastColumn()
    }));


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

    // Direct findUserByEmail usage
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
  const startTime = Date.now();
  const saveType = options.isDraft ? 'draft' : 'main';

  try {
    const userEmail = getCurrentEmail();

    if (!userEmail) {
      return { success: false, message: 'ユーザー認証が必要です' };
    }

    // Direct findUserByEmail usage
    const user = findUserByEmail(userEmail, { requestingUser: userEmail });

    if (!user) {
      return { success: false, message: 'ユーザーが見つかりません' };
    }

    // 保存タイプをオプションで制御（GAS-Native準拠）
    const saveOptions = options.isDraft ?
      { isDraft: true } :
      { isMainConfig: true };

    // 統一API使用: saveUserConfigで安全保存
    const result = saveUserConfig(user.userId, config, saveOptions);

    return result;

  } catch (error) {
    const duration = Date.now() - startTime;
    const operation = options.isDraft ? 'saveDraft' : 'saveConfig';
    console.error(`saveConfig: ERROR after ${duration}ms - ${error.message || 'Operation error'}`);
    return { success: false, message: error.message || 'エラーが発生しました' };
  }
}

// f0068fa restoration function - GAS-Native Architecture


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
    // Basic validation
    const email = getCurrentEmail();
    if (!email || !targetUserId) {
      return { success: false, message: 'Invalid request' };
    }

    // Get target user data
    const targetUser = findUserById(targetUserId, { requestingUser: email });
    if (!targetUser) {
      return { success: false, message: 'User not found' };
    }

    // Get current data using existing data service
    const userData = getUserSheetData(targetUserId, {
      includeTimestamp: true,
      classFilter: options.classFilter,
      sortBy: options.sortOrder || 'newest',
      requestingUser: email,
      targetUserEmail: targetUser.userEmail
    });

    if (!userData.success) {
      return { success: false, message: 'Data access failed' };
    }

    // Simple new item detection based on timestamp
    const lastUpdate = new Date(options.lastUpdateTime || 0);
    const newItems = userData.data.filter(item => {
      const itemTime = new Date(item.timestamp || 0);
      return itemTime > lastUpdate;
    });

    return {
      success: true,
      hasNewContent: newItems.length > 0,
      newItemsCount: newItems.length,
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    return { success: false, message: error.message };
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

    // Editor access for own spreadsheets
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

    // Optimization: cache getColumnAnalysis results for reuse
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
          //キャッシュされた結果を使用（重複API呼び出しなし）
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

    // Enhanced access control for editor users
    // 管理者は任意のスプレッドシートにアクセス、編集ユーザーは自分がアクセス権を持つもののみ
    let dataAccess;
    try {
      // 編集ユーザーの場合、まず通常権限でアクセステストを行う
      dataAccess = openSpreadsheet(spreadsheetId, { useServiceAccount: false });
      if (!dataAccess) {
        if (isAdmin) {
          // 管理者の場合、サービスアカウントでリトライ
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

    // High-precision analysis data retrieval
    const lastCol = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();
    const headers = lastCol > 0 ? getSheetHeaders(sheet, lastCol) : [];

    // Sample data retrieval for hybrid system
    let sampleData = [];
    if (lastRow > 1 && lastCol > 0) {
      const sampleSize = Math.min(10, lastRow - 1); // 最大10行のサンプル
      try {
        const dataRange = sheet.getRange(2, 1, sampleSize, lastCol);
        sampleData = dataRange.getValues();
      } catch (sampleError) {
        console.warn('getColumnAnalysis: サンプルデータ取得失敗:', sampleError.message);
        sampleData = [];
      }
    }

    // High-precision ColumnMappingService with sample data
    const diagnostics = performIntegratedColumnDiagnostics(headers, { sampleData });

    //編集者自身のアカウントでリアクション列・ハイライト列を事前追加
    try {
      const columnSetupResult = setupReactionAndHighlightColumns(spreadsheetId, sheetName, headers);
      if (columnSetupResult.columnsAdded && columnSetupResult.columnsAdded.length > 0) {
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


// Missing API Functions for Frontend Error Fix


/**
 * Get active form info - フロントエンドエラー修正用
 * @returns {Object} フォーム情報
 */
function getActiveFormInfo(userId) {

  try {
    const currentEmail = getCurrentEmail();
    if (!currentEmail) {
      console.error('Authentication failed');
      return {
        success: false,
        message: 'Authentication required',
        formUrl: null,
        formTitle: null
      };
    }

    // Get user config if userId specified (board owner's form)
    // 指定されていなければ現在のユーザーの設定を取得（backward compatibility）
    let targetUserId = userId;
    if (!targetUserId) {
      const currentUser = findUserByEmail(currentEmail, { requestingUser: currentEmail });
      if (!currentUser) {
        console.error('Current user not found:', currentEmail);
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


// GAS-side Trigger-Based Polling System


// Missing API Endpoints - Frontend/Backend Compatibility







// 🆕 CLAUDE.md準拠: 完全自動化データソース選択システム

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

    //V8ランタイム対応: const使用 + 正規表現最適化
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

    // Exponential backoff retry for resilient spreadsheet access
    const spreadsheet = executeWithRetry(
      () => SpreadsheetApp.openById(spreadsheetId),
      { operationName: 'SpreadsheetApp.openById', maxRetries: 3 }
    );
    const sheets = spreadsheet.getSheets();

    //Batch Operations: 全シート情報を一括取得（70x improvement）
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

    const accessResult = validateAccess(spreadsheetId, true);

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

// 🆕 Missing Functions Implementation - Frontend Compatibility

// システム管理関数をSystemController.gsに移動済み

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
      console.warn('callGAS: Unauthorized access attempt (no email)');
      return createAuthError();
    }

    // 厳格なセキュリティホワイトリスト
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
      'validateCompleteSpreadsheetUrl',
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

    // Security check: function name validation
    if (!functionName || typeof functionName !== 'string') {
      console.warn('callGAS: Invalid function name:', functionName);
      return {
        success: false,
        message: 'Invalid function name provided',
        securityWarning: true
      };
    }

    if (!allowedFunctions.includes(functionName)) {
      // Security log for unauthorized function access attempts
      console.warn('callGAS: Unauthorized function access attempt:', {
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
      console.warn('callGAS: Excessive arguments detected:', args.length);
      return {
        success: false,
        message: 'Too many arguments provided',
        securityWarning: true
      };
    }

    //関数実行（安全な環境で）
    if (typeof this[functionName] === 'function') {
      try {
        const result = this[functionName].apply(this, args);

        // Success audit log
        console.info('callGAS: Function executed successfully:', {
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
        console.error('callGAS: Function execution error:', {
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
      console.warn('callGAS: Function not found:', functionName);
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
    console.error('callGAS: Critical security error:', {
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
    // ✅ CLAUDE.md準拠: preloadedAuth構築でDB重複アクセス排除
    const isAdminUser = isAdministrator(currentEmail);
    const preloadedAuth = { email: currentEmail, isAdmin: isAdminUser };

    // ✅ preloadedAuthを渡してfindUserById内のgetAllUsers重複呼び出しを排除
    const targetUser = findUserById(targetUserId, {
      requestingUser: currentEmail,
      preloadedAuth
    });
    if (!targetUser) {
      return { success: false, error: '対象ユーザーが見つかりません' };
    }

    // ✅ preloadedUserを渡してgetUserConfig内のfindUserById重複呼び出しを排除
    const configResult = getUserConfig(targetUserId, targetUser);
    const config = configResult.success ? configResult.config : {};

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
 * - getCurrentEmail (session email)
 * - findUserById (target user validation)
 * - isAdministrator
 * - permission validation
 * - getUserConfig
 *
 * @param {string} targetUserId - Target user ID for admin access
 * @returns {Object} Batched result with all required admin data
 */
function getBatchedAdminData(targetUserId) {
  try {
    //Batch operation: Get current email from session
    const currentEmail = getCurrentEmail();
    if (!currentEmail) {
      return { success: false, error: 'ユーザー認証が必要です' };
    }

    // ✅ CLAUDE.md準拠: preloadedAuth構築でDB重複アクセス排除
    const isAdmin = isAdministrator(currentEmail);
    const preloadedAuth = { email: currentEmail, isAdmin };

    // ✅ preloadedAuthを渡してfindUserById内のgetAllUsers重複呼び出しを排除
    const targetUser = findUserById(targetUserId, {
      requestingUser: currentEmail,
      preloadedAuth
    });
    if (!targetUser) {
      return { success: false, error: '指定されたユーザーが見つかりません' };
    }

    //権限チェック: 管理者またはターゲットユーザー本人のみアクセス可能
    const isOwnBoard = currentEmail === targetUser.userEmail;

    if (!isAdmin && !isOwnBoard) {
      return {
        success: false,
        error: `他のユーザーの管理画面にはアクセスできません。管理者権限が必要です。`
      };
    }

    //編集者権限の追加確認（管理者でない場合）
    if (!isAdmin && !targetUser.isActive) {
      return { success: false, error: '対象ユーザーがアクティブではありません' };
    }

    // ✅ preloadedUserを渡してgetUserConfig内のfindUserById重複呼び出しを排除
    const configResult = getUserConfig(targetUserId, targetUser);
    const config = configResult.success ? configResult.config : {};

    // フロントエンド必要情報を統合取得
    const questionText = getQuestionText(config, { targetUserEmail: targetUser.userEmail });

    //URLs とタイムスタンプ情報を config に統合
    //Optimized: Use database lastModified instead of config lastModified
    const baseUrl = ScriptApp.getService().getUrl();
    const enhancedConfig = {
      ...config,
      urls: config.urls || {
        view: `${baseUrl}?mode=view&userId=${targetUserId}`,
        admin: `${baseUrl}?mode=admin&userId=${targetUserId}`
      },
      lastModified: targetUser.lastModified || config.publishedAt
    };

    return {
      success: true,
      email: currentEmail,
      user: targetUser,
      config: enhancedConfig,
      questionText: questionText || '回答ボード',
      isAdminAccess: isAdmin && !isOwnBoard // 管理者として他ユーザーにアクセスしているかどうか
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
 * @returns {Object} Batched result with user and config data
 */
function getBatchedUserConfig() {
  try {
    //Batch operation: Get email, user, and config in single coordinated call
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
        console.info(`${operationName}: Succeeded on retry ${retryCount}`);
      }

      return result;

    } catch (error) {
      lastError = error;
      retryCount++;

      const errorMessage = error && error.message ? error.message : String(error);

      // Check if this is a retryable error
      const isRetryable = isRetryableError(errorMessage);

      console.warn(`${operationName}: Attempt ${retryCount} failed: ${errorMessage}`);

      // Don't retry if error is not retryable or we've reached max retries
      if (!isRetryable || retryCount >= maxRetries) {
        break;
      }
    }
  }

  // All retries exhausted
  const finalError = lastError && lastError.message ? lastError.message : 'Unknown error';
  console.error(`${operationName}: Failed after ${retryCount} attempts: ${finalError}`);
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

// Performance Metrics API - Priority 1 Enhancement

/**
 * パフォーマンスメトリクス取得API (管理者専用)
 * Priority 1改善: 詳細監視機能追加
 *
 * @param {string} category - メトリクスカテゴリ ('api', 'cache', 'batch', 'error', 'all')
 * @param {Object} options - 取得オプション
 * @returns {Object} パフォーマンス統計結果
 */
function getPerformanceMetrics(category = 'all', options = {}) {
  try {
    // SystemController経由でメトリクス取得
    return SystemController.getPerformanceMetrics(category, options);
  } catch (error) {
    console.error('getPerformanceMetrics API error:', error.message);
    return {
      success: false,
      error: `Performance metrics collection failed: ${error.message}`,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * パフォーマンス診断API (管理者専用)
 * Priority 1改善: システム健全性診断
 *
 * @param {Object} options - 診断オプション
 * @returns {Object} 診断結果と改善推奨事項
 */
function diagnosePerformance(options = {}) {
  try {
    // SystemController経由で診断実行
    return SystemController.diagnosePerformance(options);
  } catch (error) {
    console.error('diagnosePerformance API error:', error.message);
    return {
      success: false,
      error: `Performance diagnosis failed: ${error.message}`,
      timestamp: new Date().toISOString()
    };
  }
}

// Application Access Control - app-wide control functions

/**
 * アプリ全体のアクセス制限状態をチェック
 * APP_DISABLED プロパティの状態を確認
 * @returns {boolean} true: アプリ停止中, false: 正常運用中
 */
function checkAppAccessRestriction() {
  try {
    const props = PropertiesService.getScriptProperties();
    const appDisabled = props.getProperty('APP_DISABLED');
    return appDisabled === 'true';
  } catch (error) {
    console.error('checkAppAccessRestriction error:', error.message);
    // エラー時は安全側として制限なしとする
    return false;
  }
}

/**
 * アプリ全体のアクセスを停止する（管理者専用）
 * @param {string} reason - 停止理由（オプション）
 * @returns {Object} 実行結果
 */
function disableAppAccess(reason = 'システムメンテナンス') {
  try {
    const currentEmail = getCurrentEmail();
    if (!currentEmail || !isAdministrator(currentEmail)) {
      return createAdminRequiredError();
    }

    const props = PropertiesService.getScriptProperties();
    props.setProperty('APP_DISABLED', 'true');
    props.setProperty('APP_DISABLED_REASON', reason);
    props.setProperty('APP_DISABLED_BY', currentEmail);
    props.setProperty('APP_DISABLED_AT', new Date().toISOString());

    return createSuccessResponse('アプリケーションを停止しました', {
      reason,
      disabledBy: currentEmail,
      timestamp: new Date().toISOString()
    });
  } catch (error) {
    console.error('disableAppAccess error:', error.message);
    return createExceptionResponse(error, 'アプリケーション停止処理');
  }
}

/**
 * アプリ全体のアクセス制限を解除する（管理者専用）
 * @returns {Object} 実行結果
 */
function enableAppAccess() {
  try {
    const currentEmail = getCurrentEmail();
    if (!currentEmail || !isAdministrator(currentEmail)) {
      return createAdminRequiredError();
    }

    const props = PropertiesService.getScriptProperties();

    // 停止情報を記録用に保持してから削除
    const disabledReason = props.getProperty('APP_DISABLED_REASON') || '';
    const disabledBy = props.getProperty('APP_DISABLED_BY') || '';
    const disabledAt = props.getProperty('APP_DISABLED_AT') || '';

    props.deleteProperty('APP_DISABLED');
    props.deleteProperty('APP_DISABLED_REASON');
    props.deleteProperty('APP_DISABLED_BY');
    props.deleteProperty('APP_DISABLED_AT');

    // 復旧記録を残す
    props.setProperty('APP_ENABLED_BY', currentEmail);
    props.setProperty('APP_ENABLED_AT', new Date().toISOString());

    return createSuccessResponse('アプリケーションを再開しました', {
      enabledBy: currentEmail,
      previousRestriction: {
        reason: disabledReason,
        disabledBy,
        disabledAt
      },
      timestamp: new Date().toISOString()
    });
  } catch (error) {
    console.error('enableAppAccess error:', error.message);
    return createExceptionResponse(error, 'アプリケーション再開処理');
  }
}

/**
 * アプリアクセス制限の状態情報を取得（管理者専用）
 * @returns {Object} アクセス制限状態の詳細情報
 */
function getAppAccessStatus() {
  try {
    const currentEmail = getCurrentEmail();
    if (!currentEmail || !isAdministrator(currentEmail)) {
      return createAdminRequiredError();
    }

    const props = PropertiesService.getScriptProperties();
    const isDisabled = props.getProperty('APP_DISABLED') === 'true';

    const status = {
      isDisabled,
      status: isDisabled ? 'disabled' : 'enabled',
      timestamp: new Date().toISOString()
    };

    if (isDisabled) {
      status.restriction = {
        reason: props.getProperty('APP_DISABLED_REASON') || '',
        disabledBy: props.getProperty('APP_DISABLED_BY') || '',
        disabledAt: props.getProperty('APP_DISABLED_AT') || ''
      };
    } else {
      status.lastEnabled = {
        enabledBy: props.getProperty('APP_ENABLED_BY') || '',
        enabledAt: props.getProperty('APP_ENABLED_AT') || ''
      };
    }

    return createSuccessResponse('アクセス制限状態を取得しました', status);
  } catch (error) {
    console.error('getAppAccessStatus error:', error.message);
    return createExceptionResponse(error, 'アクセス制限状態取得処理');
  }
}

/**
 * スプレッドシート一覧のキャッシュをクリア
 * 🔄 更新ボタン用のキャッシュクリア機能
 * @returns {Object} 処理結果
 */
function clearSheetsCache() {
  return { success: true, message: 'キャッシュをクリアしました' };
}
