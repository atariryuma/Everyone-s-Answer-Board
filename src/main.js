/**
 * @fileoverview main - HTTP entry points (`doGet` / `doPost`)、認証ヘルパー、
 *   isAdministrator、レトライ/バッチ認証ユーティリティ。
 */

/* global createErrorResponse, createSuccessResponse, createAuthError, createUserNotFoundError, createAdminRequiredError, createExceptionResponse, hasCoreSystemProps, getUserSheetData, addReaction, toggleHighlight, findUserByEmail, findUserById, findPublishedBoardOwner, getConfigOrDefault, getCachedProperty, enhanceConfigWithDynamicUrls, shouldEnforceDomainRestrictions, validateDomainAccess, dispatchAdminOperation, timingSafeEqual, setCachedProperty, getQuestionText, getWebAppUrl, publishApp, getLessonForReview, isBoardCollaborator */
// isAdministrator は本ファイル内で関数として定義されているため /* global */ には載せない。

/**
 * APIキー認証時にgetCurrentEmail()が管理者メールを返すためのコンテキスト
 * GASはシングルスレッドのため、グローバル変数で安全に管理できる
 */
let _apiKeyAdminEmail = null;

// 現在ユーザーのメール。getActiveUser() のみ (getEffectiveUser() は権限昇格リスク)。
//   API キー認証中は ADMIN_EMAIL を返す。
function getCurrentEmail() {
  if (_apiKeyAdminEmail) return _apiKeyAdminEmail;
  try {
    const email = Session.getActiveUser().getEmail();
    return email || null;
  } catch (error) {
    logError_('getCurrentEmail', error);
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
  // Why (v2865 / M6): ドメイン制限は組織境界の最終ゲート。 enforce が必要と確定した後の
  //   例外 / validator 欠落で fail-open すると、 transient な失敗だけで domain 外ユーザーを
  //   通してしまう。 enforcement が確定した時点以降は fail-closed (deny) にする。
  //   ただし enforcement 判定 (shouldEnforceDomainRestrictions) 自体が落ちた = 状態不明の
  //   場合のみ、 setup 中の lockout を避けるため従来どおり可用性優先で fail-open に倒す。
  let enforcementConfirmed = false;
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
      // validator 不在は単一グローバルスコープでは「コード未読込」= アプリ全体が壊れている状態で
      //   あり、 runtime のセキュリティ事象ではない。 ここは従来どおり通す (M6 の対象外)。
      return {
        allowed: true,
        reason: 'validator_missing',
        message: 'Domain validator not available'
      };
    }

    enforcementConfirmed = true;  // validator が実在し enforce 必要 → ここ以降の例外は fail-closed

    return validateDomainAccess(email, {
      allowIfAdminUnconfigured: true,
      allowIfEmailMissing: true
    });
  } catch (error) {
    console.warn('evaluateDomainRestriction failed:', error.message);
    return {
      allowed: !enforcementConfirmed,  // enforce 確定済なら deny、 状態不明なら従来どおり通す
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
  template.isRegisteredUser = email ? !!findUserByEmail(email, { requestingUser: email }) : false;
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

    // セットアップ未完了なら自動でセットアップページへ
    if (mode !== 'setup') {
      const setupDone = (typeof hasCoreSystemProps === 'function') ? hasCoreSystemProps() : true;
      if (!setupDone) {
        return HtmlService.createTemplateFromFile('SetupPage.html').evaluate().setTitle('初期設定');
      }
    }

    const handler = DO_GET_HANDLERS[mode] || handleMainMode_;
    return handler(params, currentEmail);
  } catch (error) {
    logError_('doGet', error, {
      stack: error.stack,
      mode: e.parameter?.mode,
      userId: e.parameter?.userId && typeof e.parameter.userId === 'string' ? `${e.parameter.userId.substring(0, 8)}***` : 'N/A'
    });

    const errorTemplate = HtmlService.createTemplateFromFile('ErrorBoundary.html');
    errorTemplate.title = 'システムエラー';
    errorTemplate.message = 'システムで予期しないエラーが発生しました。管理者にお問い合わせください。';
    errorTemplate.hideLoginButton = true;

    return errorTemplate.evaluate().setTitle('エラー');
  }
}

// ─── doGet mode handlers ─────────────────────────────────────
// Why: doGet を 1 ハンドラ = 1 mode の薄い dispatcher にする。
//   元の switch (~200 行) を Clean Code "Replace Conditional with Polymorphism" で table-driven 化。

function handleLoginMode_() {
  return HtmlService.createTemplateFromFile('LoginPage.html').evaluate().setTitle('ログイン');
}

function handleManualMode_() {
  return HtmlService.createTemplateFromFile('TeacherManual.html').evaluate().setTitle('使い方ガイド');
}

function handleAdminMode_(params, currentEmail) {
  if (!currentEmail) {
    return createRedirectTemplate('ErrorBoundary.html', 'ログインが必要です。トップページに戻ってログインしてください。');
  }
  const targetUserId = params.userId;
  if (!targetUserId) {
    return createRedirectTemplate('ErrorBoundary.html', 'URLが正しくありません。ログインページからやり直してください。');
  }
  const adminData = getBatchedAdminData(targetUserId);
  if (!adminData.success) {
    return createRedirectTemplate('ErrorBoundary.html', adminData.error || 'このページへのアクセス権限がありません。');
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
    isEditor: true,
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

function handleSetupMode_(params, currentEmail) {
  let showSetup = false;
  try {
    if (typeof hasCoreSystemProps === 'function') {
      showSetup = !hasCoreSystemProps();
    } else {
      // hasCoreSystemProps が未ロードの fallback。getCachedProperty で 3 回の Property fetch をまとめる。
      const hasAdmin = !!getCachedProperty('ADMIN_EMAIL');
      const hasDb = !!getCachedProperty('DATABASE_SPREADSHEET_ID');
      const hasCreds = !!getCachedProperty('SERVICE_ACCOUNT_CREDS');
      showSetup = !(hasAdmin && hasDb && hasCreds);
    }
  } catch (_) {
    showSetup = true;
  }
  if (showSetup) {
    return HtmlService.createTemplateFromFile('SetupPage.html').evaluate().setTitle('初期設定');
  }
  return createAccessRestrictedTemplate(currentEmail);
}

function handleAppSetupMode_(params, currentEmail) {
  if (!currentEmail || !isAdministrator(currentEmail)) {
    return createRedirectTemplate('ErrorBoundary.html', 'このページへのアクセス権限がありません。');
  }
  let userIdParam = params.userId || '';
  if (!userIdParam) {
    const adminUser = findUserByEmail(currentEmail, { requestingUser: currentEmail });
    if (adminUser) userIdParam = adminUser.userId;
  }
  const template = HtmlService.createTemplateFromFile('AppSetupPage.html');
  template.userIdParam = userIdParam;
  return template.evaluate().setTitle('システム設定');
}

function handleViewMode_(params, currentEmail) {
  if (!currentEmail) {
    return createRedirectTemplate('ErrorBoundary.html', 'ログインが必要です。トップページに戻ってログインしてください。');
  }
  const targetUserId = params.userId;
  if (!targetUserId) {
    return createRedirectTemplate('ErrorBoundary.html', 'URLが正しくありません。ログインページからやり直してください。');
  }
  const viewerData = getBatchedViewerData(targetUserId, currentEmail);
  if (!viewerData.success) {
    return createRedirectTemplate('ErrorBoundary.html', viewerData.error || '対象のボードが見つかりません。URLを確認してください。');
  }

  const { targetUser, config, isAdminUser } = viewerData;
  const isOwnBoard = currentEmail === targetUser.userEmail;
  const isPublished = Boolean(config.isPublished);

  // Why isActive check: 管理者がユーザーを「無効化」(isActive=false) しても、本人の
  //   board が isPublished=true のままなら domain user は依然として viewable だった
  //   (Edge case audit #8 / handleViewMode_ ACL gap)。disable は「公開を強制終了する」
  //   意図なので、isActive=false は admin/owner 以外には未公開として扱う。
  if (targetUser.isActive === false && !isAdminUser && !isOwnBoard) {
    const template = HtmlService.createTemplateFromFile('Unpublished.html');
    template.isEditor = false;
    template.editorName = '';
    template.userId = targetUserId;
    template.boardUrl = '';
    return template.evaluate().setTitle('未公開');
  }

  if (!isPublished) {
    const template = HtmlService.createTemplateFromFile('Unpublished.html');
    template.isEditor = isAdminUser || isOwnBoard;
    template.editorName = targetUser.userName || targetUser.userEmail || '';
    template.userId = targetUserId;
    const baseUrl = getWebAppUrl();
    template.boardUrl = baseUrl ? `${baseUrl}?mode=view&userId=${targetUserId}` : '';
    return template.evaluate().setTitle('未公開');
  }

  // v2855+: ボード SS の editor (writer/owner) として共有された共同教師は collaborator として
  //   owner と同じ「進行操作」 (highlight / unpublish / リアクション削除) を実行可能。
  //   ボード設定変更や lesson 編集は引き続き owner / admin 専用 (AccessControl.js 参照)。
  const isCollaborator = !isAdminUser && !isOwnBoard
    && (typeof isBoardCollaborator === 'function' && isBoardCollaborator(targetUser, currentEmail));
  const isEditor = isAdminUser || isOwnBoard || isCollaborator;
  const template = HtmlService.createTemplateFromFile('Page.html');
  template.userId = targetUserId;
  template.userEmail = targetUser.userEmail;
  template.questionText = '読み込み中...';
  template.boardTitle = targetUser.userEmail || '回答ボード';
  template.isEditor = isEditor;
  template.isAdminUser = isAdminUser;
  template.isOwnBoard = isOwnBoard;
  template.isCollaborator = isCollaborator;
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
    isCollaborator,
    formUrl: config.formUrl || '',
    showDetails: config.showDetails !== false,
    displaySettings: config.displaySettings || DEFAULT_DISPLAY_SETTINGS
  });
  return template.evaluate().setTitle('回答ボード');
}

function handleReviewMode_(params, currentEmail) {
  // Lesson 振り返り (read-only)。owner / admin のみ可、生徒からは 403。
  if (!currentEmail) {
    return createRedirectTemplate('ErrorBoundary.html', 'ログインが必要です。');
  }
  const targetUserId = params.userId;
  const lessonId = params.lessonId;
  if (!targetUserId || !lessonId) {
    return createRedirectTemplate('ErrorBoundary.html', 'URLが正しくありません (userId / lessonId が必要)。');
  }
  const reviewResult = getLessonForReview(targetUserId, lessonId);
  if (!reviewResult.success) {
    return createRedirectTemplate('ErrorBoundary.html', reviewResult.message || 'レッスンを開けませんでした。');
  }
  const lesson = reviewResult.data.lesson;
  const template = HtmlService.createTemplateFromFile('Page.html');
  template.userId = targetUserId;
  template.userEmail = currentEmail;
  // Why: heading は hydrate 完了後 client 側 (__updateReviewHeading) で該当 phase の question
  //   に上書きされる。banner 側で lesson 名を出すので、heading にレッスン名を入れると重複。
  //   ここは placeholder としてのみ機能 (skeleton 中の短時間だけ表示される)。
  template.questionText = '📖 読み込み中…';
  template.boardTitle = '📖 振り返り: ' + (lesson.name || '');
  template.isEditor = true;
  template.isAdminUser = isAdministrator(currentEmail);
  template.isOwnBoard = true;
  template.sheetName = (lesson.lessonJson && lesson.lessonJson.phases && lesson.lessonJson.phases[0] && lesson.lessonJson.phases[0].sheetName) || '';
  template.configJSON = JSON.stringify({
    userId: targetUserId,
    userEmail: currentEmail,
    spreadsheetId: '',
    sheetName: '',
    questionText: '📖 読み込み中…',
    isPublished: false,
    isEditor: true,
    isAdminUser: template.isAdminUser,
    isOwnBoard: true,
    isReviewMode: true,
    reviewLesson: lesson
  });
  return template.evaluate().setTitle('振り返り: ' + (lesson.name || ''));
}

function handleMainMode_(params, currentEmail) {
  return createAccessRestrictedTemplate(currentEmail);
}

const DO_GET_HANDLERS = {
  login: handleLoginMode_,
  manual: handleManualMode_,
  admin: handleAdminMode_,
  setup: handleSetupMode_,
  appSetup: handleAppSetupMode_,
  view: handleViewMode_,
  review: handleReviewMode_,
  main: handleMainMode_
};

// ─── Auth & redirect helpers ─────────────────────────────────────

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
      template.webAppUrl = getWebAppUrl() || '';
    }

    const title = redirectPage === 'ErrorBoundary.html' ? 'エラー' : '回答ボード';
    return template.evaluate().setTitle(title);
  } catch (templateError) {
    logError_('createRedirectTemplate', templateError);
    const fallbackTemplate = HtmlService.createTemplateFromFile('ErrorBoundary.html');
    fallbackTemplate.title = 'システムエラー';
    fallbackTemplate.message = 'ページの表示中にエラーが発生しました。';
    fallbackTemplate.hideLoginButton = true;
    fallbackTemplate.webAppUrl = getWebAppUrl() || '';
    return fallbackTemplate.evaluate().setTitle('エラー');
  }
}

// =====================================================================
// doPost action handlers — 旧来 doPost 内の switch case を個別関数化。
// 各 handler は (request, email?, ...) を受け取り、{success, message, ...} を返す純粋関数。
// テスト容易性 + 認知負荷の低減のために抽出。
// =====================================================================

const isPlainObject = (v) => Boolean(v) && typeof v === 'object' && !Array.isArray(v);

/**
 * getData / refreshData ハンドラ — 公開ボードのデータ取得。
 * @param {Object} request - doPost payload
 * @param {string} email - 認証済 viewer email
 * @param {string} action - 'getData' | 'refreshData' (ログ識別用)
 * @returns {Object} response
 */
function doPostHandleGetData(request, email, action) {
  if (request.options !== undefined && !isPlainObject(request.options)) {
    return createErrorResponse('Invalid options payload');
  }
  try {
    const user = findUserByEmail(email, { requestingUser: email });
    if (!user) {
      return createUserNotFoundError();
    }
    return { success: true, data: getUserSheetData(user.userId, request.options || {}) };
  } catch (error) {
    logError_(action, error);
    return createExceptionResponse(error);
  }
}

/**
 * addReaction ハンドラ — 児童/教師のリアクション追加。
 * @param {Object} request
 * @returns {Object} response
 */
function doPostHandleAddReaction(request) {
  if (!request.userId || typeof request.userId !== 'string' || !request.userId.trim()) {
    return createErrorResponse('Target user ID required for reaction');
  }
  if (!isValidRowIdPayload(request.rowId)) {
    return createErrorResponse('Target row ID must be a positive integer or "row_<n>" string');
  }
  if (typeof request.reactionType !== 'string') {
    return createErrorResponse('Reaction type required');
  }
  return addReaction(request.userId, request.rowId, request.reactionType);
}

/**
 * toggleHighlight ハンドラ — 教師の注目マーク切替。
 * @param {Object} request
 * @returns {Object} response
 */
function doPostHandleToggleHighlight(request) {
  if (!request.userId || typeof request.userId !== 'string' || !request.userId.trim()) {
    return createErrorResponse('Target user ID required for highlight');
  }
  if (!isValidRowIdPayload(request.rowId)) {
    return createErrorResponse('Target row ID must be a positive integer or "row_<n>" string');
  }
  return toggleHighlight(request.userId, request.rowId);
}

/**
 * publishApp ハンドラ — ボード公開（owner 認証は publishApp 内で行う）。
 */
function doPostHandlePublishApp(request) {
  try {
    if (!isPlainObject(request.config) || Object.keys(request.config).length === 0) {
      return createErrorResponse('Publish config is required');
    }
    return publishApp(request.config);
  } catch (error) {
    logError_('publishApp', error);
    return createExceptionResponse(error);
  }
}

/**
 * rowId payload validator.
 *   accepts: 42 / "42" / "row_42" (positive integer-bearing forms)
 *   rejects: NaN / negative / 0 / non-integer string / null / undefined / object
 *
 * Why: addReaction / toggleHighlight の rowId は下流で parseInt(replace('row_', '')) されて
 *   getRange(rowIndex, ...) に渡る。NaN が来ると Sheet API が黙って崩壊するので、doPost 層で reject。
 *
 * @param {*} rowId
 */
function isValidRowIdPayload(rowId) {
  if (rowId === null || rowId === undefined) return false;
  if (typeof rowId === 'number') {
    return Number.isInteger(rowId) && rowId > 0;
  }
  if (typeof rowId === 'string') {
    const cleaned = rowId.replace(/^row_/, '').trim();
    if (!cleaned) return false;
    const n = Number(cleaned);
    return Number.isInteger(n) && n > 0;
  }
  return false;
}

/**
 * Handle POST requests
 * @param {Object} e - Event object containing POST data
 * @returns {TextOutput} JSON response with operation result
 */
function doPost(e) {
  const jsonResponse = (payload) => ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);

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
      logError_('doPost.JSON.parse', parseError, {
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
    // doPost allowlist — 追加時は (1) この配列に追加 (2) 下の switch case を追加
    // (3) 各 case 内で request fields の input validation を必ず行うこと（CLAUDE.md より）。
    //   - getData / refreshData : 公開ボードの閲覧（ドメイン認証のみ、API key 不要）
    //   - addReaction / toggleHighlight : 児童/教師の interaction（rowId validate 必須）
    //   - publishApp : ボード公開（owner 認証 + etag conflict 検出）
    //   - adminApi : 管理者全 op（APIキー + timingSafeEqual 比較）
    //   - setupApiKey : 初回 APIキー設定（一度だけ、管理者のみ）
    //   - reportClientError : フロントエンドエラー → Cloud Logging へ
    const allowedActions = ['getData', 'addReaction', 'toggleHighlight', 'refreshData', 'publishApp', 'adminApi', 'setupApiKey', 'reportClientError'];
    if (!allowedActions.includes(action)) {
      return jsonResponse({
        success: false,
        message: action ? `Unknown action: ${action}` : 'Unknown action: 不明',
        error: 'UNKNOWN_ACTION'
      });
    }

    // setupApiKey: 初回APIキー設定（キー未設定時のみ動作）
    if (action === 'setupApiKey') {
      return jsonResponse(handleSetupApiKeyAction_(request));
    }

    if (action === 'adminApi') {
      return jsonResponse(handleAdminApiAction_(request));
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

    // action handler は table-driven dispatch (DO_POST_ACTION_HANDLERS)。
    //   各 handler は (request, email, action) を受け取り result を返す純粋関数。
    const actionHandler = DO_POST_ACTION_HANDLERS[action];
    const result = actionHandler
      ? actionHandler(request, email, action)
      : createErrorResponse(action ? `Unknown action: ${action}` : 'Unknown action: 不明');

    return jsonResponse(result);

  } catch (error) {
    // Why: stack trace を保持して Cloud Logging に出すことで、CLI からエラー発生箇所を
    //   即座に特定できる。クライアントへの応答メッセージは generic に保つ（情報漏洩防止）。
    logError_('doPost', error, {
      stack: error && error.stack ? error.stack.substring(0, 1000) : undefined,
      timestamp: new Date().toISOString()
    });
    return jsonResponse({
      success: false,
      message: 'リクエスト処理中にエラーが発生しました。再度お試しください。'
    });
  }
}

// ─── doPost action handlers + dispatch table ─────────────────────────────────────

// setupApiKey: 初回 ADMIN_API_KEY 設定ゲート。管理者のみ、既存キーがあれば reject。
function handleSetupApiKeyAction_(request) {
  // Why direct PropertiesService (lint suppressed): 初回キー設定ゲートは stale cache が
  //   二重設定の race を引き起こすので 30s memory cache をバイパスして必ず最新を読む。
  // lint-disable-next-line no-direct-property-fetch
  const existingKey = PropertiesService.getScriptProperties().getProperty('ADMIN_API_KEY');
  if (existingKey) {
    return createErrorResponse('ADMIN_API_KEY is already configured', null, { error: 'ALREADY_CONFIGURED' });
  }
  const setupEmail = getCurrentEmail();
  if (!setupEmail || !isAdministrator(setupEmail)) {
    return createErrorResponse('Admin authentication required', null, { error: 'ADMIN_REQUIRED' });
  }
  // Why: length >= 16 だけだと空白埋め文字列 (16 spaces) が通る silent data quality bug。
  //   trim 後の effective length と「非空白文字を含む」を併用検証。
  const rawKey = typeof request.apiKey === 'string' ? request.apiKey : '';
  const trimmedKey = rawKey.trim();
  const hasMeaningfulChars = trimmedKey.length >= 16 && /\S/.test(trimmedKey);
  if (!hasMeaningfulChars) {
    return createErrorResponse('apiKey must be a string of at least 16 non-whitespace characters', null, { error: 'INVALID_KEY_FORMAT' });
  }
  // setCachedProperty also busts the 30s helpers.js memory cache so the very
  // next adminApi call in this process sees the new key.
  setCachedProperty('ADMIN_API_KEY', rawKey);
  return createSuccessResponse('ADMIN_API_KEY has been set', { configured: true });
}

// adminApi: APIキー認証 (Session 不要 → AI エージェントからアクセス可能)
function handleAdminApiAction_(request) {
  const apiKey = typeof request.apiKey === 'string' ? request.apiKey : '';
  const storedKey = getCachedProperty('ADMIN_API_KEY');
  if (!storedKey) {
    return createErrorResponse('ADMIN_API_KEY is not configured. Use setupApiKey action first.', null, { error: 'API_KEY_NOT_CONFIGURED' });
  }
  if (!apiKey || !timingSafeEqual(apiKey, storedKey)) {
    return createErrorResponse('Invalid API key', null, { error: 'INVALID_API_KEY' });
  }
  const adminParams = request.params === undefined ? {} : request.params;
  if (!isPlainObject(adminParams)) {
    return createErrorResponse('params must be a JSON object', null, { error: 'INVALID_PARAMS' });
  }
  _apiKeyAdminEmail = getCachedProperty('ADMIN_EMAIL');
  try {
    return dispatchAdminOperation(request.operation, adminParams);
  } finally {
    _apiKeyAdminEmail = null;
  }
}

const DO_POST_ACTION_HANDLERS = {
  getData: (req, email, action) => doPostHandleGetData(req, email, action),
  refreshData: (req, email, action) => doPostHandleGetData(req, email, action),
  addReaction: (req) => doPostHandleAddReaction(req),
  toggleHighlight: (req) => doPostHandleToggleHighlight(req),
  publishApp: (req) => doPostHandlePublishApp(req),
  reportClientError: (req, email) => handleClientErrorReport(email, req.payload)
};

/**
 * Frontend error report receiver — drains client-side errors into Cloud Logging.
 *
 * Why: GAS Cloud Logging only captures server-side console.* calls. Errors thrown
 *   inside page.js.html / AdminPanel.js.html otherwise live only in the browser's
 *   DevTools console and are invisible to CLI debugging (`npm run logs:cloud`).
 *   This handler logs structured client error info via console.error so it shows
 *   up alongside backend logs.
 *
 * Limits applied (defense-in-depth, since the endpoint accepts client-controlled
 * strings from any authenticated domain user):
 *   - message ≤ 500 chars, stack ≤ 2000 chars, context fields ≤ 200 chars
 *   - max 30 reports per 5 minutes per email (CacheService backed)
 *   - level restricted to error / warn / log; anything else → error
 *
 * @param {string} email - submitter email (already auth-checked upstream)
 * @param {Object} payload - { level, message, stack, source, line, col, context }
 * @returns {Object} result envelope
 */
/**
 * Frontend orchestration entry point — google.script.run.reportClientError(payload) から呼ばれる。
 *
 * Why: SharedErrorHandling.html は google.script.run.reportClientError(...) を直叩きする。
 *   google.script.run はトップレベル関数しか呼べないため、グローバルとして expose する必要が
 *   ある。doPost action='reportClientError' とは別パス (window.error / unhandledrejection
 *   の捕捉時に doPost を経由しない軽量パス)。
 *
 *   この wrapper が無いと、フロントエンドからの error report は silently fail して
 *   Cloud Logging には届かなくなる ([client/error] prefix が無い = CLAUDE.md の
 *   「フロントエンドエラーも収集される」が破綻)。
 */
function reportClientError(payload) {
  try {
    const email = getCurrentEmail();
    return handleClientErrorReport(email, payload);
  } catch (error) {
    // Silent: client error reporter 自体の失敗で client 体験を壊さない。
    return createErrorResponse('reportClientError internal error', null, { error: 'REPORT_FAILED' });
  }
}

function handleClientErrorReport(email, payload) {
  try {
    if (!isPlainObject(payload)) {
      return createErrorResponse('payload must be an object', null, { error: 'INVALID_PAYLOAD' });
    }

    // Rate limit per email (Cache TTL = 300s).
    // Why: Cache rather than ScriptProperties to avoid write contention; a single
    //      page bug could otherwise flood logs from every viewer simultaneously.
    try {
      const cache = CacheService.getScriptCache();
      // 仮名化（メアドそのものはキーに入れない；ログ収集の incident で漏らさない）
      // Why: 旧実装は SHA-1 先頭 4 bytes (8 hex = 32 bit) で衝突確率が高く、別ユーザ間で
      //   レート制限を共有してしまう可能性があった。64 bit (8 bytes = 16 hex) に拡大し、
      //   2^32 倍の衝突耐性を持たせる。
      const bucket = `cer_${Utilities.computeDigest(
        Utilities.DigestAlgorithm.SHA_1, String(email || '')
      ).slice(0, 8).map(b => (b & 0xff).toString(16).padStart(2, '0')).join('')}`;
      const raw = cache.get(bucket);
      const count = raw ? parseInt(raw, 10) || 0 : 0;
      if (count >= 30) {
        return createErrorResponse('rate limit exceeded', null, { error: 'RATE_LIMIT' });
      }
      cache.put(bucket, String(count + 1), 300);
    } catch (cacheErr) {
      // Cache 障害時はレート制限を諦めて続行（信頼性 > 制限）
    }

    const cap = (v, n) => {
      if (v === null || v === undefined) return '';
      return String(v).replace(/[\r\n\t]+/g, ' ').substring(0, n);
    };

    const level = ['error', 'warn', 'log', 'debug'].includes(payload.level) ? payload.level : 'error';
    const message = cap(payload.message, 500);
    const stack = cap(payload.stack, 2000);
    const source = cap(payload.source, 200);
    const lineNo = Number.isFinite(payload.line) ? payload.line : null;
    const colNo = Number.isFinite(payload.col) ? payload.col : null;
    const context = cap(payload.context, 200);
    const url = cap(payload.url, 300);
    const userAgent = cap(payload.userAgent, 200);

    // Why: console.error にすることで Cloud Logging の severity が ERROR になり、
    //      `npm run logs:cloud --severity ERROR` で即座に拾える。
    //      `[client]` prefix で backend ログとの区別を可能にし、`signatureOf` の
    //      集計でも別グループになる。
    const tag = `[client/${level}]`;
    const summary = `${tag} ${context ? '(' + context + ') ' : ''}${message}`;
    // Why (domain leak fix): 旧実装は `user@***` でドメイン部のみマスクしていたが、
    //   local part をそのまま残すと naha-okinawa の児童 ID (例: t260781p) が
    //   ログから個人特定可能だった。local part も短縮 SHA1 で完全仮名化する。
    const anonReporter = (() => {
      if (!email || typeof email !== 'string' || email.indexOf('@') < 0) return 'unknown';
      try {
        const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_1, email)
          .slice(0, 4)
          .map(b => (b & 0xff).toString(16).padStart(2, '0'))
          .join('');
        return `anon-${hash}`;
      } catch (_) {
        return 'anon';
      }
    })();
    const detail = {
      reporter: anonReporter,
      url,
      source,
      line: lineNo,
      col: colNo,
      stack: stack || undefined,
      userAgent
    };

    if (level === 'error') console.error(summary, detail);
    else if (level === 'warn') console.warn(summary, detail);
    else console.log(summary, detail);

    return createSuccessResponse('reported', { accepted: true });
  } catch (error) {
    // 自分自身で例外を起こしたらサイレントで返す（ログループ回避）
    return createErrorResponse('report handler error', null, { error: error && error.message });
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

    return email.toLowerCase() === adminEmail.toLowerCase();
  } catch (error) {
    logError_('main.isAdministrator', error);
    return false;
  }
}

/**
 * @param {string} targetUserId - Target user ID
 * @param {string} currentEmail - Current viewer email
 * @returns {Object} Batched result with all required data
 */
function getBatchedViewerData(targetUserId, currentEmail) {
  try {
    const isAdminUser = isAdministrator(currentEmail);
    const preloadedAuth = { email: currentEmail, isAdmin: isAdminUser };

    const targetUser = findPublishedBoardOwner(targetUserId, currentEmail, { preloadedAuth });
    if (!targetUser) {
      return { success: false, error: '対象ユーザーが見つかりません' };
    }

    const config = getConfigOrDefault(targetUserId, targetUser);

    return {
      success: true,
      targetUser,
      config,
      isAdminUser
    };

  } catch (error) {
    logError_('getBatchedViewerData', error);
    return {
      success: false,
      error: `データ取得エラー: ${error.message}`
    };
  }
}

/**
 * @param {string} targetUserId - Target user ID for admin access
 * @returns {Object} Batched result with all required admin data
 */
function getBatchedAdminData(targetUserId) {
  try {
    const currentEmail = getCurrentEmail();
    if (!currentEmail) {
      return createAuthError();
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

    const config = getConfigOrDefault(targetUserId, targetUser);

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
    logError_('getBatchedAdminData', error);
    return {
      success: false,
      error: `管理者データ取得エラー: ${error.message}`
    };
  }
}

/**
 * 管理者認証をバッチで行う（getCurrentEmail + isAdministrator を 1 回で）。
 * @param {Object} [options]
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
    logError_('getBatchedAdminAuth', error);
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
 * 指数バックオフ付きリトライ。429 検出時はベース遅延を 2 倍に拡張。
 * @param {Function} operation
 * @param {Object} [options]
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

        // operation 側が既に adaptive backoff (= inner Utilities.sleep) を済ませている場合は
        // ここで二重に sleep しない。 二重 backoff は 429 storm 時に inner(最大60s)+outer(最大20s)
        // が積み上がり GAS 6 分実行制限を踏み抜く原因になっていた。
        const alreadyBackedOff = errorMessage.includes('adaptive backoff');

        if (!alreadyBackedOff) {
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
      }

      return operation();

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

  logError_(operationName, lastError || new Error('Unknown error'), { retryCount });
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
    'authentication failed',
    // 非冪等 write (例: :append) の失敗。 commit 済か判定不能なため retry すると重複を生む (H3)。
    'non-idempotent'
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

