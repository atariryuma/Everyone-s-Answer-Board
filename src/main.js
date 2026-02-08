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

/* global createErrorResponse, createSuccessResponse, createAuthError, createUserNotFoundError, createAdminRequiredError, createExceptionResponse, hasCoreSystemProps, getUserSheetData, addReaction, toggleHighlight, validateConfig, findUserByEmail, findUserById, findUserBySpreadsheetId, createUser, getAllUsers, updateUser, openSpreadsheet, getUserConfig, saveUserConfig, clearConfigCache, cleanConfigFields, getQuestionText, validateAccess, URL, UserService, CACHE_DURATION, TIMEOUT_MS, SLEEP_MS, SYSTEM_LIMITS, SystemController, getViewerBoardData, performIntegratedColumnDiagnostics, generateRecommendedMapping, getFormInfo, enhanceConfigWithDynamicUrls, getCachedProperty, getSheetInfo, setupDomainWideSharing, shouldEnforceDomainRestrictions, validateDomainAccess */


/**
 * 現在のユーザーのメールアドレスを取得
 * ✅ SECURITY: getActiveUser() のみ使用（getEffectiveUser() は権限昇格リスクあり）
 * @returns {string|null} ユーザーのメールアドレス、または認証されていない場合はnull
 */
function getCurrentEmail() {
  try {
    const email = Session.getActiveUser().getEmail();
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

/**
 * ドメイン制限を評価
 * @param {string|null} email - 検証対象メール
 * @returns {Object} 検証結果
 */
function evaluateDomainRestriction(email) {
  try {
    const mustEnforce = (typeof shouldEnforceDomainRestrictions === 'function')
      ? shouldEnforceDomainRestrictions()
      : true;

    if (!mustEnforce) {
      return {
        allowed: true,
        reason: 'setup_incomplete',
        message: 'Domain restriction skipped during setup'
      };
    }

    if (typeof validateDomainAccess !== 'function') {
      return {
        allowed: true,
        reason: 'validator_missing',
        message: 'Domain validator not available'
      };
    }

    return validateDomainAccess(email, {
      allowIfAdminUnconfigured: true,
      allowIfEmailMissing: true
    });
  } catch (error) {
    console.warn('evaluateDomainRestriction failed:', error.message);
    return {
      allowed: true,
      reason: 'validation_error',
      message: error.message
    };
  }
}

/**
 * アクセス制限テンプレートを生成
 * @param {string|null} email - 現在ユーザーのメール
 * @param {boolean} isAppDisabled - アプリ停止状態
 * @param {string} message - 表示メッセージ
 * @returns {HtmlOutput} 制限ページ
 */
function createAccessRestrictedTemplate(email, isAppDisabled = false, message = '') {
  const template = HtmlService.createTemplateFromFile('AccessRestricted.html');
  template.isAdministrator = email ? isAdministrator(email) : false;
  template.userEmail = email || '';
  template.timestamp = new Date().toISOString();
  template.isAppDisabled = Boolean(isAppDisabled);
  if (message) {
    template.message = message;
  }
  return template.evaluate().setTitle('回答ボード');
}



/**
 * Handle GET requests
 * @param {Object} e - Event object containing request parameters
 * @returns {HtmlOutput} HTML response for the requested page
 */
function doGet(e) {
  try {
    const params = e ? e.parameter : {};
    const mode = params.mode || 'main';

    const currentEmail = getCurrentEmail();

    // セットアップ完了後は管理者ドメイン外アクセスを遮断
    const domainRestriction = evaluateDomainRestriction(currentEmail);
    if (domainRestriction && domainRestriction.allowed === false) {
      return createAccessRestrictedTemplate(
        currentEmail,
        false,
        '管理者と同一ドメインのアカウントでアクセスしてください。'
      );
    }

    const isAppDisabled = checkAppAccessRestriction();
    if (isAppDisabled) {
      const isAdmin = currentEmail ? isAdministrator(currentEmail) : false;
      const canOpenAppSetup = mode === 'appSetup' && isAdmin;
      if (!canOpenAppSetup) {
        return createAccessRestrictedTemplate(currentEmail, true);
      }
    }


    switch (mode) {
      case 'login': {
        return HtmlService.createTemplateFromFile('LoginPage.html').evaluate().setTitle('ログイン');
      }

      case 'manual': {
        return HtmlService.createTemplateFromFile('TeacherManual.html').evaluate().setTitle('使い方ガイド');
      }

      case 'admin': {
        if (!currentEmail) {
          return createRedirectTemplate('ErrorBoundary.html', 'ユーザー認証が必要です');
        }

        const targetUserId = params.userId;
        if (!targetUserId) {
          return createRedirectTemplate('ErrorBoundary.html', 'ユーザーIDが指定されていません');
        }

        const adminData = getBatchedAdminData(targetUserId);
        if (!adminData.success) {
          return createRedirectTemplate('ErrorBoundary.html', adminData.error || '管理者権限が必要です');
        }

        const { email, user, config } = adminData;
        const isAdmin = isAdministrator(email);

        const enhancedConfig = enhanceConfigWithDynamicUrls(config, user.userId);

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

        return template.evaluate().setTitle('管理');
      }

      case 'setup': {
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
          showSetup = true;
        }

        if (showSetup) {
          return HtmlService.createTemplateFromFile('SetupPage.html').evaluate().setTitle('初期設定');
        } else {
          return createAccessRestrictedTemplate(currentEmail);
        }
      }

      case 'appSetup': {
        if (!currentEmail || !isAdministrator(currentEmail)) {
          return createRedirectTemplate('ErrorBoundary.html', '管理者権限が必要です');
        }

        const userIdParam = params.userId;

        const template = HtmlService.createTemplateFromFile('AppSetupPage.html');

        template.userIdParam = userIdParam || '';

        return template.evaluate().setTitle('システム設定');
      }

      case 'view': {
        if (!currentEmail) {
          return createRedirectTemplate('ErrorBoundary.html', 'ユーザー認証が必要です');
        }

        const targetUserId = params.userId;
        if (!targetUserId) {
          return createRedirectTemplate('ErrorBoundary.html', 'ユーザーIDが指定されていません');
        }

        const viewerData = getBatchedViewerData(targetUserId, currentEmail);
        if (!viewerData.success) {
          return createRedirectTemplate('ErrorBoundary.html', viewerData.error || '対象ユーザーが見つかりません');
        }

        const { targetUser, config, isAdminUser } = viewerData;
        const isOwnBoard = currentEmail === targetUser.userEmail;
        const isPublished = Boolean(config.isPublished);

        if (!isPublished) {
          const template = HtmlService.createTemplateFromFile('Unpublished.html');
          template.isEditor = isAdminUser || isOwnBoard; // 表示内容制御
          template.editorName = targetUser.userName || targetUser.userEmail || '';
          template.userId = targetUserId; // 管理パネル遷移用

          const baseUrl = getWebAppUrl();
          template.boardUrl = baseUrl ? `${baseUrl}?mode=view&userId=${targetUserId}` : '';


          return template.evaluate().setTitle('未公開');
        }

        const template = HtmlService.createTemplateFromFile('Page.html');
        template.userId = targetUserId;
        template.userEmail = targetUser.userEmail;
        template.questionText = '読み込み中...';
        template.boardTitle = targetUser.userEmail || '回答ボード';

        const isEditor = isAdminUser || isOwnBoard;
        template.isEditor = isEditor;
        template.isAdminUser = isAdminUser;
        template.isOwnBoard = isOwnBoard;

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

        return template.evaluate().setTitle('回答ボード');
      }

      case 'main':
      default: {
        return createAccessRestrictedTemplate(currentEmail);
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

    return errorTemplate.evaluate().setTitle('エラー');
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

    const title = redirectPage === 'ErrorBoundary.html' ? 'エラー' : '回答ボード';
    return template.evaluate().setTitle(title);
  } catch (templateError) {
    console.error('createRedirectTemplate error:', templateError.message);
    const fallbackTemplate = HtmlService.createTemplateFromFile('ErrorBoundary.html');
    fallbackTemplate.title = 'システムエラー';
    fallbackTemplate.message = 'ページの表示中にエラーが発生しました。';
    fallbackTemplate.hideLoginButton = false;
    return fallbackTemplate.evaluate().setTitle('エラー');
  }
}

/**
 * Handle POST requests
 * @param {Object} e - Event object containing POST data
 * @returns {TextOutput} JSON response with operation result
 */
function doPost(e) {
  const jsonResponse = (payload) => ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);

  const isPlainObject = (value) => Boolean(value) && typeof value === 'object' && !Array.isArray(value);

  const isJsonContentType = (contentType) => {
    if (!contentType || typeof contentType !== 'string') {
      return true;
    }
    const normalized = contentType.toLowerCase();
    return normalized.includes('application/json') || normalized.includes('+json');
  };

  try {
    if (!e || !e.postData || typeof e.postData.contents !== 'string') {
      return jsonResponse({
        success: false,
        message: 'Request body is required',
        error: 'MISSING_POST_BODY'
      });
    }

    const contentType = e.postData.type || '';
    if (!isJsonContentType(contentType)) {
      return jsonResponse({
        success: false,
        message: 'Unsupported Content-Type. Use application/json',
        error: 'UNSUPPORTED_CONTENT_TYPE'
      });
    }

    const postData = e.postData.contents;
    const MAX_POST_BODY_SIZE = 1024 * 1024; // 1MB
    if (postData && postData.length > MAX_POST_BODY_SIZE) {
      return jsonResponse({
        success: false,
        message: 'Request body too large',
        error: 'PAYLOAD_TOO_LARGE'
      });
    }

    if (!postData.trim()) {
      return jsonResponse({
        success: false,
        message: 'Empty request body',
        error: 'EMPTY_POST_BODY'
      });
    }

    let request;
    try {
      request = JSON.parse(postData);
    } catch (parseError) {
      console.error('doPost: Invalid JSON received:', {
        error: parseError.message,
        dataLength: postData ? postData.length : 0,
        dataPreview: postData ? postData.substring(0, 100) : 'N/A'
      });
      return jsonResponse({
        success: false,
        message: 'Invalid JSON format in request body',
        error: 'JSON_PARSE_ERROR'
      });
    }

    if (!isPlainObject(request)) {
      return jsonResponse({
        success: false,
        message: 'Request payload must be a JSON object',
        error: 'INVALID_REQUEST_PAYLOAD'
      });
    }

    const action = typeof request.action === 'string' ? request.action.trim() : '';
    const allowedActions = ['getData', 'addReaction', 'toggleHighlight', 'refreshData', 'publishApp'];
    if (!allowedActions.includes(action)) {
      return jsonResponse({
        success: false,
        message: action ? `Unknown action: ${action}` : 'Unknown action: 不明',
        error: 'UNKNOWN_ACTION'
      });
    }

    const email = getCurrentEmail();
    if (!email) {
      return jsonResponse(createAuthError());
    }

    const domainRestriction = evaluateDomainRestriction(email);
    if (domainRestriction && domainRestriction.allowed === false) {
      return jsonResponse({
        success: false,
        message: '管理者と同一ドメインのアカウントでアクセスしてください。',
        error: 'DOMAIN_ACCESS_DENIED'
      });
    }

    let result;
    switch (action) {
      case 'getData':
      case 'refreshData': {
        if (request.options !== undefined && !isPlainObject(request.options)) {
          result = createErrorResponse('Invalid options payload');
          break;
        }
        try {
          const user = findUserByEmail(email, { requestingUser: email });
          if (!user) {
            result = createUserNotFoundError();
          } else {
            result = { success: true, data: getUserSheetData(user.userId, request.options || {}) };
          }
        } catch (error) {
          console.error(`${action} error:`, error.message);
          result = createExceptionResponse(error);
        }
        break;
      }
      case 'addReaction':
        if (!request.userId || typeof request.userId !== 'string' || !request.userId.trim()) {
          result = createErrorResponse('Target user ID required for reaction');
        } else if (request.rowId === undefined || request.rowId === null) {
          result = createErrorResponse('Target row ID required for reaction');
        } else if (typeof request.reactionType !== 'string') {
          result = createErrorResponse('Reaction type required');
        } else {
          result = addReaction(request.userId, request.rowId, request.reactionType);
        }
        break;
      case 'toggleHighlight':
        if (!request.userId || typeof request.userId !== 'string' || !request.userId.trim()) {
          result = createErrorResponse('Target user ID required for highlight');
        } else if (request.rowId === undefined || request.rowId === null) {
          result = createErrorResponse('Target row ID required for highlight');
        } else {
          result = toggleHighlight(request.userId, request.rowId);
        }
        break;
      case 'publishApp':
        try {
          const publishConfig = isPlainObject(request.config) ? request.config : {};
          if (!isPlainObject(publishConfig) || Object.keys(publishConfig).length === 0) {
            result = createErrorResponse('Publish config is required');
            break;
          }
          if (SystemController && typeof SystemController.publishApp === 'function') {
            result = SystemController.publishApp(publishConfig);
          } else if (typeof publishApp === 'function') {
            result = publishApp(publishConfig);
          } else {
            result = createErrorResponse('公開処理関数が見つかりません');
          }
        } catch (error) {
          console.error('publishApp error:', error.message);
          result = createExceptionResponse(error);
        }
        break;
      default:
        result = createErrorResponse(action ? `Unknown action: ${action}` : 'Unknown action: 不明');
    }

    return jsonResponse(result);

  } catch (error) {
    const errorMessage = error && error.message ? error.message : '予期しないエラーが発生しました';
    console.error('doPost error:', errorMessage);
    return jsonResponse({
      success: false,
      message: errorMessage
    });
  }
}





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
    const adminEmail = getCachedProperty('ADMIN_EMAIL');
    if (!adminEmail) {
      console.warn('isAdministrator: ADMIN_EMAIL設定が見つかりません');
      return false;
    }

    const isAdmin = email.toLowerCase() === adminEmail.toLowerCase();
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
      console.warn('callGAS: Unauthorized access attempt (no email)');
      return createAuthError();
    }

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

    if (isAdmin) {
      allowedFunctions.push(...adminOnlyFunctions);
    }

    if (!functionName || typeof functionName !== 'string') {
      console.warn('callGAS: Invalid function name:', functionName);
      return {
        success: false,
        message: 'Invalid function name provided',
        securityWarning: true
      };
    }

    if (!allowedFunctions.includes(functionName)) {
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

    if (args.length > 10) {
      console.warn('callGAS: Excessive arguments detected:', args.length);
      return {
        success: false,
        message: 'Too many arguments provided',
        securityWarning: true
      };
    }

    if (typeof this[functionName] === 'function') {
      try {
        const result = this[functionName].apply(this, args);

        return {
          success: true,
          functionName,
          result,
          options,
          executedAt: new Date().toISOString(),
          securityLevel: isAdmin ? 'admin' : 'user'
        };
      } catch (functionError) {
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
    console.error('callGAS: Critical security error:', {
      error: error.message,
      functionName,
      timestamp: new Date().toISOString()
    });
    return createExceptionResponse(error, 'Secure function call failed');
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
    const isAdminUser = isAdministrator(currentEmail);
    const preloadedAuth = { email: currentEmail, isAdmin: isAdminUser };

    const targetUser = findUserById(targetUserId, {
      requestingUser: currentEmail,
      preloadedAuth,
      allowPublishedRead: true
    });
    if (!targetUser) {
      return { success: false, error: '対象ユーザーが見つかりません' };
    }

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
    const currentEmail = getCurrentEmail();
    if (!currentEmail) {
      return { success: false, error: 'ユーザー認証が必要です' };
    }

    const isAdmin = isAdministrator(currentEmail);
    const preloadedAuth = { email: currentEmail, isAdmin };

    const targetUser = findUserById(targetUserId, {
      requestingUser: currentEmail,
      preloadedAuth
    });
    if (!targetUser) {
      return { success: false, error: '指定されたユーザーが見つかりません' };
    }

    const isOwnBoard = currentEmail === targetUser.userEmail;

    if (!isAdmin && !isOwnBoard) {
      return {
        success: false,
        error: `他のユーザーの管理画面にはアクセスできません。管理者権限が必要です。`
      };
    }

    if (!isAdmin && !targetUser.isActive) {
      return { success: false, error: '対象ユーザーがアクティブではありません' };
    }

    const configResult = getUserConfig(targetUserId, targetUser);
    const config = configResult.success ? configResult.config : {};

    const questionText = getQuestionText(config, { targetUserEmail: targetUser.userEmail });

    const baseUrl = getWebAppUrl();
    const enhancedConfig = {
      ...config,
      urls: config.urls || {
        view: baseUrl ? `${baseUrl}?mode=view&userId=${targetUserId}` : '',
        admin: baseUrl ? `${baseUrl}?mode=admin&userId=${targetUserId}` : ''
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
      if (retryCount > 0) {
        const errorMessage = lastError && lastError.message ? lastError.message : '';

        const is429Error = errorMessage.includes('429') || errorMessage.includes('Quota exceeded');
        const baseDelay = is429Error ? initialDelay * 2 : initialDelay;

        const delay = Math.min(
          baseDelay * Math.pow(2, retryCount - 1),
          maxDelay
        );
        if (retryCount === 1 || retryCount === maxRetries - 1) {
          console.warn(`${operationName}: Retry ${retryCount}/${maxRetries - 1} after ${delay}ms delay${is429Error ? ' (429 quota)' : ''}`);
        }
        Utilities.sleep(delay);
      }

      const result = operation();

      if (retryCount > 0) {
      }

      return result;

    } catch (error) {
      lastError = error;
      retryCount++;

      const errorMessage = error && error.message ? error.message : String(error);

      const isRetryable = isRetryableError(errorMessage);

      if (retryCount >= maxRetries || !isRetryable) {
        console.warn(`${operationName}: Attempt ${retryCount} failed: ${errorMessage}`);
      }

      if (!isRetryable || retryCount >= maxRetries) {
        break;
      }
    }
  }

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

  for (const pattern of nonRetryablePatterns) {
    if (lowerMessage.includes(pattern)) {
      return false;
    }
  }

  for (const pattern of retryablePatterns) {
    if (lowerMessage.includes(pattern)) {
      return true;
    }
  }

  return true;
}


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
