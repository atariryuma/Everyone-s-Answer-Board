/**
 * @fileoverview StudyQuest - みんなの回答ボード
 * スプレッドシートのカスタムメニューを使用しない独立版。
 */

// =================================================================
// 定数定義
// =================================================================

// マルチテナント用ユーザーデータベース設定
const USER_DB_CONFIG = {
  SHEET_NAME: 'Users',
  HEADERS: [
    'userId',
    'adminEmail',
    'spreadsheetId', 
    'spreadsheetUrl',
    'createdAt',
    'accessToken',
    'configJson',
    'lastAccessedAt',
    'isActive'
  ]
};
const COLUMN_HEADERS = {
  TIMESTAMP: 'タイムスタンプ',
  EMAIL: 'メールアドレス',
  CLASS: 'クラス',
  OPINION: '回答',
  REASON: '理由',
  NAME: '名前',
  UNDERSTAND: 'なるほど！',
  LIKE: 'いいね！',
  CURIOUS: 'もっと知りたい！',
  HIGHLIGHT: 'ハイライト'
};
const SCORING_CONFIG = {
  LIKE_MULTIPLIER_FACTOR: 0.05 // 1いいね！ごとにスコアが5%増加
};

const TIME_CONSTANTS = {
  LOCK_WAIT_MS: 10000
};

const REACTION_KEYS = ["UNDERSTAND","LIKE","CURIOUS"];
const EMAIL_REGEX = /^[^\n@]+@[^\n@]+\.[^\n@]+$/;
// Debug flag. Set to true to enable verbose logging
var DEBUG = false;

function debugLog() {
  if (DEBUG && typeof console !== 'undefined' && console.log) {
    console.log.apply(console, arguments);
  }
}

/**
 * 現在のユーザーの権限を検証するヘルパー関数
 * @return {Object} {userId, userInfo} - 検証済みのユーザー情報
 * @throws {Error} 認証失敗時
 */
function validateCurrentUser() {
  const props = PropertiesService.getUserProperties();
  const userId = props.getProperty('CURRENT_USER_ID');
  if (!userId) {
    throw new Error('認証が必要です。ログインし直してください。');
  }
  
  const userInfo = getUserInfo(userId);
  const currentEmail = Session.getActiveUser().getEmail();
  if (!userInfo || userInfo.adminEmail !== currentEmail) {
    throw new Error('権限がありません。管理者アカウントでログインしてください。');
  }
  return { userId, userInfo };
}

/**
 * XSS攻撃を防ぐための入力サニタイゼーション関数
 * @param {string} input - サニタイズする入力
 * @return {string} サニタイズされた文字列
 */
function sanitizeInput(input) {
  if (typeof input !== 'string') return input;
  return input
    .replace(/[<>'"&]/g, function(match) {
      return {
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#x27;',
        '&': '&amp;'
      }[match];
    })
    .substring(0, 1000); // 長さ制限
}

/**
 * レート制限チェック関数
 * @param {string} action - 実行するアクション
 * @param {string} userEmail - ユーザーメール
 */
function checkRateLimit(action, userEmail) {
  const cache = CacheService.getScriptCache();
  const key = `rate_limit_${action}_${userEmail}`;
  const attempts = parseInt(cache.get(key) || '0');
  
  if (attempts >= 30) { // 30回/時間制限
    throw new Error('操作回数の制限に達しました。しばらく待ってから再試行してください。');
  }
  
  cache.put(key, (attempts + 1).toString(), 3600); // 1時間
}

/**
 * セキュアなエラーハンドリング関数
 * @param {string} prefix - エラープレフィックス
 * @param {Error} error - 元のエラー
 * @param {boolean} returnObject - オブジェクトで返すかどうか
 * @return {Object|void} エラーオブジェクトまたは例外スロー
 */
function secureHandleError(prefix, error, returnObject = false) {
  // ログには詳細を出力（管理者用）
  console.error(prefix + ':', error);
  
  // ユーザーには安全なメッセージのみ表示
  let userMessage = 'システムエラーが発生しました。管理者にお問い合わせください。';
  
  // 特定のエラーパターンに対する安全なメッセージ
  const errorMsg = error?.message || String(error);
  
  if (errorMsg.includes('permission') || errorMsg.includes('Permission')) {
    userMessage = '権限がありません。管理者にお問い合わせください。';
  } else if (errorMsg.includes('not found') || errorMsg.includes('見つかりません')) {
    userMessage = '指定されたデータが見つかりません。';
  } else if (errorMsg.includes('invalid') || errorMsg.includes('無効')) {
    userMessage = '無効な操作です。入力内容を確認してください。';
  }
  
  if (returnObject) {
    return { status: 'error', message: userMessage };
  }
  throw new Error(userMessage);
}

function applySecurityHeaders(output) {
  if (output && typeof output.addMetaTag === 'function') {
    output.addMetaTag(
      'Content-Security-Policy',
      "default-src 'self'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline'; img-src 'self' data:"
    );
    if (typeof output.setXFrameOptionsMode === 'function') {
      // Allow embedding within the same origin by default for enhanced security
      // unless a future embedding requirement mandates ALLOWALL.
      output.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.SAMEORIGIN);
    }
  }
  return output;
}
var getConfig;
var handleError;

// handleError関数はErrorHandling.gsで定義されています

if (typeof global !== 'undefined' && global.getConfig) {
  getConfig = global.getConfig;
}

function getCurrentSpreadsheet() {
  const props = PropertiesService.getUserProperties();
  const spreadsheetId = props.getProperty('CURRENT_SPREADSHEET_ID');
  
  if (!spreadsheetId) {
    // 従来の動作（単一テナント時）
    return SpreadsheetApp.getActiveSpreadsheet();
  }
  
  // マルチテナント時は指定されたスプレッドシートを使用
  return SpreadsheetApp.openById(spreadsheetId);
}

function safeGetUserEmail() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email || typeof email !== 'string' || !isValidEmail(email)) {
      throw new Error('Invalid user email');
    }
    return email;
  } catch (e) {
    console.error('Failed to get user email:', e);
    // In test environment, return a default test email
    if (typeof global !== 'undefined' && global.mockUserEmail) {
      return global.mockUserEmail;
    }
    throw new Error('認証が必要です。Googleアカウントでログインしてください。');
  }
}

/**
 * フロントエンドから現在のユーザーのメールアドレスを取得するためのラッパー。
 * Registration.html などから呼び出される。
 *
 * @return {string} ユーザーのメールアドレス
 */
function getActiveUserEmail() {
  return safeGetUserEmail();
}

/**
 * セキュリティ強化: ユーザー認証状態の検証
 * @return {Object} 認証状態と関連情報
 */
function verifyUserAuthentication() {
  try {
    const email = safeGetUserEmail();
    if (!email || !isValidEmail(email)) {
      return {
        authenticated: false,
        error: 'Invalid or missing email'
      };
    }
    
    return {
      authenticated: true,
      email: email,
      domain: getEmailDomain(email)
    };
  } catch (error) {
    return {
      authenticated: false,
      error: 'Authentication failed'
    };
  }
}

function isValidEmail(email) {
  if (!email || typeof email !== 'string') {
    return false;
  }
  return EMAIL_REGEX.test(email);
}

function getEmailDomain(email) {
  return (email || '').toString().split('@').pop().toLowerCase();
}

function isSameDomain(emailA, emailB) {
  if (!emailA || !emailB || typeof emailA !== 'string' || typeof emailB !== 'string') {
    return false;
  }
  if (!isValidEmail(emailA) || !isValidEmail(emailB)) {
    return false;
  }
  const domainA = getEmailDomain(emailA);
  const domainB = getEmailDomain(emailB);
  return domainA && domainB && domainA === domainB;
}

/**
 * ウェブアプリのデプロイ者のメールアドレスを取得・保存します。
 * @returns {string} デプロイ者のメールアドレス
 */
function getWebAppDeployerEmail() {
  const props = PropertiesService.getScriptProperties();
  const deployerEmail = props.getProperty('WEB_APP_DEPLOYER_EMAIL');

  if (deployerEmail) {
    return deployerEmail;
  }

  // スクリプトプロパティに保存されていない場合、現在のユーザーをデプロイ者として保存
  try {
    const newDeployerEmail = Session.getEffectiveUser().getEmail();
    if (newDeployerEmail) {
      props.setProperty('WEB_APP_DEPLOYER_EMAIL', newDeployerEmail);
      debugLog(`ウェブアプリのデプロイ者を ${newDeployerEmail} に設定しました。`);
      return newDeployerEmail;
    }
  } catch (e) {
    console.error('デプロイ者のメールアドレス取得に失敗しました。', e);
    // フォールバックとして空文字を返す
    return '';
  }
  return '';
}

function getAdminEmails(spreadsheetId) {
  const props = PropertiesService.getScriptProperties();
  let adminEmails = [];
  
  if (spreadsheetId) {
    const str = props.getProperty(`ADMIN_EMAILS_${spreadsheetId}`);
    if (str) adminEmails = str.split(',').map(e => e.trim()).filter(Boolean);
  }
  
  if (adminEmails.length === 0) {
    const str = props.getProperty('ADMIN_EMAILS') || '';
    adminEmails = str.split(',').map(e => e.trim()).filter(Boolean);
  }
  
  // 管理者が設定されていない場合、スプレッドシートの所有者を管理者とする
  if (adminEmails.length === 0) {
    try {
      const ss = getCurrentSpreadsheet();
      const owner = ss.getOwner();
      if (owner) {
        const ownerEmail = owner.getEmail();
        if (ownerEmail) {
          debugLog('管理者未設定のため、所有者を管理者に設定:', ownerEmail);
          adminEmails = [ownerEmail];
        }
      }
    } catch (e) {
      console.error('所有者情報の取得に失敗:', e);
    }
  }
  
  return adminEmails;
}

function isUserAdmin(email) {
  const userProps = PropertiesService.getUserProperties();
  const userId = userProps.getProperty('CURRENT_USER_ID');
  const userEmail = email || safeGetUserEmail();

  debugLog('isUserAdmin check:', {
    email: email, 
    userEmail: userEmail, 
    userId: userId 
  });

  if (userId) {
    const userInfo = getUserInfo(userId);
    const spreadsheetId = userInfo && userInfo.spreadsheetId;
    const admins = getAdminEmails(spreadsheetId);
    debugLog('Multi-tenant admin check:', {
      spreadsheetId: spreadsheetId, 
      admins: admins, 
      userEmail: userEmail, 
      isAdmin: admins.includes(userEmail) 
    });
    return admins.includes(userEmail);
  }

  const admins = getAdminEmails();
  debugLog('Single-tenant admin check:', {
    admins: admins, 
    userEmail: userEmail, 
    isAdmin: admins.includes(userEmail) 
  });
  return admins.includes(userEmail);
}

function checkAdmin() {
  return isUserAdmin();
}

function publishApp(sheetName) {
  if (!checkAdmin()) {
    throw new Error('権限がありません。');
  }
  const props = PropertiesService.getScriptProperties();
  props.setProperty('IS_PUBLISHED', 'true');
  props.setProperty('PUBLISHED_SHEET_NAME', sheetName);
  return `「${sheetName}」を公開しました。`;
}

function unpublishApp() {
  if (!checkAdmin()) {
    throw new Error('権限がありません。');
  }
  const props = PropertiesService.getScriptProperties();
  props.setProperty('IS_PUBLISHED', 'false');
  if (props.deleteProperty) props.deleteProperty('PUBLISHED_SHEET_NAME');
  return 'アプリを非公開にしました。';
}

function parseReactionString(val) {
  if (!val) return [];
  return val
    .toString()
    .split(',')
    .map(s => s.trim())
    .filter(Boolean);
}

function hashTimestamp(ts) {
  var str = String(ts);
  var hash = 0;
  for (var i = 0; i < str.length; i++) {
    hash = ((hash << 5) - hash) + str.charCodeAt(i);
    hash |= 0;
  }
  return hash;
}


/**
 * 管理パネルの初期化に必要なデータを取得します。
 * @returns {object} - 現在のアクティブシート情報とシートのリスト
 */
function getAdminSettings() {
  const props = PropertiesService.getScriptProperties();
  const userProps = PropertiesService.getUserProperties();
  const userId = userProps.getProperty('CURRENT_USER_ID');
  
  let adminEmails = [];
  let appSettings = {};
  
  if (userId) {
    // マルチテナントモード
    const userInfo = getUserInfo(userId);
    if (userInfo) {
      adminEmails = getAdminEmails(userInfo.spreadsheetId);
      appSettings = getAppSettingsForUser();
    }
  } else {
    // 従来のモード
    adminEmails = getAdminEmails();
    appSettings = getAppSettingsForUser();
  }
  
  const allSheets = getSheets();
  const currentUserEmail = safeGetUserEmail();
  
  // 新しいシステム：常に利用可能な回答ボード
  // activeSheetNameがあれば利用可能、なければシート選択が必要
  const isPublished = props.getProperty('IS_PUBLISHED') === 'true';
  const activeSheetName = appSettings.activeSheetName || props.getProperty('ACTIVE_SHEET_NAME') || '';
  const isAvailable = !!activeSheetName;
  
  return {
    allSheets: allSheets,
    currentUserEmail: currentUserEmail,
    deployId: props.getProperty('DEPLOY_ID'),
    webAppUrl: getWebAppUrlEnhanced(), // 簡素化された関数を呼び出す
    adminEmails: adminEmails,
    isUserAdmin: adminEmails.includes(currentUserEmail),
    activeSheetName: activeSheetName,
    showNames: appSettings.showNames,
    showCounts: appSettings.showCounts,
    isPublished: isPublished
  };
}

/**
 * アクティブシートを指定されたシートに変更します。
 * @param {string} sheetName - アクティブにするシート名。
 */
function switchActiveSheet(sheetName) {
  const props = PropertiesService.getUserProperties();
  const userId = props.getProperty('CURRENT_USER_ID');
  
  if (!userId) {
    throw new Error('ユーザー情報が見つかりません。');
  }
  
  const userInfo = getUserInfo(userId);
  
  if (!userInfo || userInfo.adminEmail !== Session.getActiveUser().getEmail()) {
    throw new Error('権限がありません。');
  }
  
  if (!sheetName) {
    throw new Error('シート名が指定されていません。');
  }
  
  // スプレッドシートの存在確認
  const spreadsheet = SpreadsheetApp.openById(userInfo.spreadsheetId);
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`シート「${sheetName}」が見つかりません。`);
  }

  prepareSheetForBoard(sheetName);
  clearHeaderCache(sheetName);

  try {
    updateUserConfig(userId, {
      activeSheetName: sheetName,
      publishedAt: new Date().toISOString() // 公開日時を記録
    });
  } catch (error) {
    throw new Error(`設定の更新に失敗しました: ${error.message}`);
  }

  auditLog('ACTIVE_SHEET_CHANGED', userId, { sheetName, userEmail: userInfo.adminEmail });
  return `アクティブシートを「${sheetName}」に変更しました。`;
}

/**
 * アクティブシートの選択をクリアします。
 */
function clearActiveSheet() {
  const props = PropertiesService.getUserProperties();
  const userId = props.getProperty('CURRENT_USER_ID');
  
  if (!userId) {
    throw new Error('ユーザー情報が見つかりません。');
  }
  
  const userInfo = getUserInfo(userId);
  if (!userInfo || userInfo.adminEmail !== Session.getActiveUser().getEmail()) {
    throw new Error('権限がありません。');
  }
  
  // アクティブシートと公開日時をクリア
  try {
    updateUserConfig(userId, {
      activeSheetName: '',
      publishedAt: '' // 公開日時をクリア
    });
  } catch (error) {
    console.error('Failed to update user config:', error);
    throw new Error(`設定の更新に失敗しました: ${error.message}`);
  }
  
  auditLog('ACTIVE_SHEET_CLEARED', userId, { userEmail: userInfo.adminEmail });
  
  return 'アクティブシートをクリアしました。';
}


function setShowDetails(flag) {
  const props = PropertiesService.getUserProperties();
  const userId = props.getProperty('CURRENT_USER_ID');
  
  if (!userId) {
    throw new Error('ユーザー情報が見つかりません。');
  }
  
  const userInfo = getUserInfo(userId);
  if (!userInfo || userInfo.adminEmail !== Session.getActiveUser().getEmail()) {
    throw new Error('権限がありません。');
  }
  
  // ユーザー設定を更新
  try {
    updateUserConfig(userId, {
      showNames: Boolean(flag),
      showCounts: Boolean(flag),
      showDetails: Boolean(flag)
    });
  } catch (error) {
    console.error('Failed to update user config:', error);
    throw new Error(`設定の更新に失敗しました: ${error.message}`);
  }

  auditLog('SHOW_DETAILS_UPDATED', userId, { showNames: Boolean(flag), showCounts: Boolean(flag), userEmail: userInfo.adminEmail });

  return `詳細表示を${flag ? '有効' : '無効'}にしました。`;
}

function setDisplayOptions(options) {
  const props = PropertiesService.getUserProperties();
  const userId = props.getProperty('CURRENT_USER_ID');

  if (!userId) {
    throw new Error('ユーザー情報が見つかりません。');
  }

  const userInfo = getUserInfo(userId);
  if (!userInfo || userInfo.adminEmail !== Session.getActiveUser().getEmail()) {
    throw new Error('権限がありません。');
  }

  try {
    updateUserConfig(userId, {
      showNames: Boolean(options.showNames),
      showCounts: Boolean(options.showCounts)
    });
  } catch (error) {
    console.error('Failed to update user config:', error);
    throw new Error(`設定の更新に失敗しました: ${error.message}`);
  }

  auditLog('DISPLAY_OPTIONS_UPDATED', userId, { options, userEmail: userInfo.adminEmail });

  return '表示設定を更新しました。';
}



// =================================================================
// GAS Webアプリケーションのエントリーポイント
// =================================================================
function validateUserId(userId) {
  if (!userId || typeof userId !== 'string') {
    throw new Error('ユーザーIDが指定されていません。');
  }
  if (!/^[a-zA-Z0-9-_]{10,}$/.test(userId)) {
    throw new Error('無効なユーザーIDの形式です。');
  }
  return userId;
}

function doGet(e) {
  try {
    e = e || {};
    const params = e.parameter || {};
    const userId = params.userId;
    const mode = params.mode;


    // ユーザーIDがない場合は登録画面を表示
    if (!userId) {
      const output = HtmlService.createTemplateFromFile('Registration').evaluate();
      applySecurityHeaders(output);
      return output
        .setTitle('StudyQuest - 新規登録')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }
    
    // ユーザーIDを検証
    const validatedUserId = validateUserId(userId);
    
    // ユーザー情報を取得
    let userInfo;
    try {
      userInfo = getUserInfo(validatedUserId);
    } catch (e) {
      userInfo = getUserInfoInternal(validatedUserId);
    }
    if (!userInfo) {
      const output = HtmlService.createHtmlOutput('無効なユーザーIDです。');
      applySecurityHeaders(output);
      return output.setTitle('エラー');
    }

    // --- 自動非公開タイマーのチェック ---
    const config = userInfo.configJson || {};
    if (config.publishedAt) {
      const publishedDate = new Date(config.publishedAt);
      const sixHoursLater = new Date(publishedDate.getTime() + 6 * 60 * 60 * 1000);
      if (new Date() > sixHoursLater) {
        // セキュリティ向上: 機密情報を含むログは削除
        console.log('ボードが6時間経過したため自動的に非公開になりました。');
        clearActiveSheet(); // サーバー側で非公開処理
        // ユーザーに通知するための専用ページを表示
        const template = HtmlService.createTemplateFromFile('Unpublished');
        template.message = 'この回答ボードは、公開から6時間が経過したため、安全のため自動的に非公開になりました。再度利用する場合は、管理者にご連絡ください。';
        const output = template.evaluate();
        applySecurityHeaders(output);
        return output.setTitle('公開期間が終了しました');
      }
    }

    // 現在のユーザーを取得（必須）
    let viewerEmail;
    try {
      viewerEmail = safeGetUserEmail();
    } catch (e) {
      if (userInfo.configJson && userInfo.configJson.isPublished === false) {
        const template = HtmlService.createTemplateFromFile('Unpublished');
        template.userId = validatedUserId;
        template.userInfo = userInfo;
        template.userEmail = userInfo.adminEmail;
        template.isOwner = false;
        template.ownerName = userInfo.adminEmail || 'Unknown';
        template.boardUrl = `${getWebAppUrlEnhanced()}?userId=${validatedUserId}`;
        auditLog('UNPUBLISHED_ACCESS_NO_LOGIN', validatedUserId, { error: e });
        const output = template.evaluate();
        applySecurityHeaders(output);
        if (output.setSandboxMode) output.setSandboxMode(HtmlService.SandboxMode.IFRAME);
        return output.setTitle('StudyQuest - 回答ボード');
      }
      if (typeof HtmlService !== 'undefined') {
        const output = HtmlService.createHtmlOutput('認証が必要です。Googleアカウントでログインしてください。');
        applySecurityHeaders(output);
        return output.setTitle('認証エラー');
      }
      throw e;
    }
    
    // ドメインベースのアクセス制御
    if (!isSameDomain(viewerEmail, userInfo.adminEmail)) {
      auditLog('ACCESS_DENIED', validatedUserId, { viewerEmail, adminEmail: userInfo.adminEmail });
      if (typeof HtmlService !== 'undefined') {
        const output = HtmlService.createHtmlOutput('システムエラーが発生しました。管理者にお問い合わせください。');
        applySecurityHeaders(output);
        return output.setTitle('エラー');
      }
      throw new Error('権限がありません。同じドメインのユーザーのみアクセスできます。');
    }

    // 現在のコンテキストを設定
    PropertiesService.getUserProperties().setProperties({
      CURRENT_USER_ID: validatedUserId,
      CURRENT_SPREADSHEET_ID: userInfo.spreadsheetId
    });
  
    // 管理モードの場合
    if (mode === 'admin') {
      // 管理者権限のチェック
      if (!(viewerEmail === userInfo.adminEmail || isUserAdmin(viewerEmail))) {
        auditLog('ADMIN_ACCESS_DENIED', validatedUserId, { viewerEmail });
        if (typeof HtmlService !== 'undefined') {
          const output = HtmlService.createHtmlOutput('権限がありません。');
          applySecurityHeaders(output);
          return output.setTitle('アクセス拒否');
        }
        throw new Error('権限がありません。');
      }

      const template = HtmlService.createTemplateFromFile('AdminPanel');
      template.userId = validatedUserId;
      template.userInfo = userInfo;
      template.adminUserEmail = userInfo.adminEmail;
      template.webAppAdminEmail = getWebAppDeployerEmail(); // お問い合わせ先メールアドレス
      auditLog('ADMIN_ACCESS', validatedUserId, { viewerEmail });
      const output = template.evaluate();
      applySecurityHeaders(output);
      if (output.setSandboxMode) output.setSandboxMode(HtmlService.SandboxMode.IFRAME);
      return output.setTitle('StudyQuest - 管理パネル');
    }
  
  // 通常表示モード - 常に回答ボードを表示

  
  // 下位互換性：古いsheetNameを新しいactiveSheetNameにマッピング
  const activeSheetName = config.activeSheetName || config.sheetName || '';
  
  // デバッグ情報をログに出力
  debugLog('Config debug:', {
    userId: validatedUserId,
    configJson: config,
    activeSheetName: activeSheetName,
    hasActiveSheet: !!activeSheetName
  });
  
  // アクティブシートが設定されていない場合は、統合された管理画面を表示
  if (!activeSheetName) {
    // 管理者の場合は統合管理画面、一般ユーザーの場合は待機画面
    if (userInfo.adminEmail === viewerEmail) {
      const template = HtmlService.createTemplateFromFile('AdminPanel');
      template.userId = validatedUserId;
      template.userInfo = userInfo;
      template.adminUserEmail = userInfo.adminEmail;
      template.webAppAdminEmail = getWebAppDeployerEmail(); // お問い合わせ先メールアドレス
      auditLog('ADMIN_ACCESS_NO_SHEET', validatedUserId, { viewerEmail });
      const output = template.evaluate();
      applySecurityHeaders(output);
      if (output.setSandboxMode) output.setSandboxMode(HtmlService.SandboxMode.IFRAME);
      return output.setTitle('StudyQuest - 管理パネル');
    } else {
      const template = HtmlService.createTemplateFromFile('Unpublished');
      template.userId = validatedUserId;
      template.userInfo = userInfo;
      template.userEmail = userInfo.adminEmail;
      template.isOwner = false;
      template.ownerName = userInfo.adminEmail || 'Unknown';
      template.boardUrl = `${getWebAppUrlEnhanced()}?userId=${validatedUserId}`;
      auditLog('UNPUBLISHED_ACCESS', validatedUserId, { viewerEmail, isOwner: false });
      const output = template.evaluate();
      applySecurityHeaders(output);
      if (output.setSandboxMode) output.setSandboxMode(HtmlService.SandboxMode.IFRAME);
      return output.setTitle('StudyQuest - 回答ボード');
    }
  }

  // 公開されていない場合は待機画面を表示
  if (!config.isPublished && userInfo.adminEmail !== viewerEmail) {
    const template = HtmlService.createTemplateFromFile('Unpublished');
    template.userId = validatedUserId;
    template.userInfo = userInfo;
    template.userEmail = userInfo.adminEmail;
    template.isOwner = false;
    template.ownerName = userInfo.adminEmail || 'Unknown';
    template.boardUrl = `${getWebAppUrlEnhanced()}?userId=${validatedUserId}`;
    auditLog('UNPUBLISHED_ACCESS', validatedUserId, { viewerEmail, isOwner: false });
    const output = template.evaluate();
    applySecurityHeaders(output);
    if (output.setSandboxMode) output.setSandboxMode(HtmlService.SandboxMode.IFRAME);
    return output.setTitle('StudyQuest - 回答ボード');
  }
  
  // 既存のPage.htmlを使用
  const template = HtmlService.createTemplateFromFile('Page');
  let mapping = {};
  try {
    if (typeof getConfig === 'function') {
      mapping = getConfig(activeSheetName);
    }
  } catch (error) {
    console.warn('Config not found for sheet:', activeSheetName, error);
    mapping = {};
  }
  
  const isCurrentUserAdmin = userInfo.adminEmail === viewerEmail;

  const cfgShowNames = typeof config.showNames !== 'undefined' ? config.showNames : (config.showDetails || false);
  const cfgShowCounts = typeof config.showCounts !== 'undefined' ? config.showCounts : (config.showDetails || false);

  Object.assign(template, {
    showAdminFeatures: false, // 管理者であっても最初は閲覧モードで起動
    showHighlightToggle: isCurrentUserAdmin, // 管理者は閲覧モードでもハイライト可能
    showScoreSort: false, // 管理者であっても最初は閲覧モードで起動
    isStudentMode: true, // 管理者であっても最初は学生モードで起動
    isAdminUser: isCurrentUserAdmin, // 管理者権限の有無はそのまま渡す
    showCounts: cfgShowCounts,
    displayMode: cfgShowNames ? 'named' : 'anonymous',
    sheetName: activeSheetName,
    mapping: mapping,
    userId: userId,
    ownerName: userInfo.adminEmail || 'Unknown'
  });
  
    template.userEmail = viewerEmail;
    
    auditLog('VIEW_ACCESS', validatedUserId, { viewerEmail, sheetName: activeSheetName });
    
    const output = template.evaluate();
    applySecurityHeaders(output);
    return output
        .setTitle('StudyQuest - みんなのかいとうボード')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (error) {
    console.error('doGet error:', error);
    if (typeof HtmlService !== 'undefined') {
      const output = HtmlService.createHtmlOutput('システムエラーが発生しました。管理者にお問い合わせください。');
      applySecurityHeaders(output);
      return output.setTitle('エラー');
    }
    throw error;
  }
}


// =================================================================
// クライアントサイドから呼び出される関数
// =================================================================

/**
 * 指定されたシートまたはアクティブシートのデータを取得します。
 * @param {string} requestedSheetName - 取得するシート名（省略時はアクティブシート）
 * @param {string} classFilter - クラスフィルター
 * @param {string} sortBy - ソート順
 */
function getPublishedSheetData(requestedSheetName, classFilter, sortBy) {
  // セキュリティ強化: 権限チェックを最初に実行
  const { userId, userInfo } = validateCurrentUser();
  
  const props = PropertiesService.getUserProperties();
  const spreadsheetId = props.getProperty('CURRENT_SPREADSHEET_ID');
  
  if (!spreadsheetId) {
    throw new Error('スプレッドシート情報が見つかりません。');
  }
  
  if (!userInfo) {
    throw new Error('無効なユーザーです。');
  }
  
  const config = userInfo.configJson || {};
  
  // 下位互換性を保ちながら、アクティブシート名を取得
  const activeSheetName = config.activeSheetName || config.sheetName || '';
  
  // リクエストされたシート名、またはアクティブシートを使用
  const targetSheetName = requestedSheetName || activeSheetName;
  
  if (!targetSheetName) {
    throw new Error('表示するシートが設定されていません。');
  }
  
  // ユーザーのスプレッドシートを開く
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getSheetByName(targetSheetName);
  
  if (!sheet) {
    throw new Error(`シート「${targetSheetName}」が見つかりません。`);
  }
  
  // 既存のgetSheetData関数のロジックを使用（スプレッドシートを指定）
  const order = sortBy || 'newest';
  const data = getSheetDataForSpreadsheet(spreadsheet, targetSheetName, classFilter, order);
  
  const cfgShowNames = typeof config.showNames !== 'undefined' ? config.showNames : (config.showDetails || false);
  const cfgShowCounts = typeof config.showCounts !== 'undefined' ? config.showCounts : (config.showDetails || false);

  return {
    sheetName: targetSheetName,
    header: data.header,
    rows: data.rows,
    showNames: cfgShowNames,
    showCounts: cfgShowCounts,
    isActiveSheet: targetSheetName === activeSheetName
  };
}

/**
 * 利用可能なシート一覧とアクティブシート情報を取得します。
 */
function getAvailableSheets() {
  // セキュリティ強化: 権限チェックを最初に実行
  const { userId, userInfo } = validateCurrentUser();
  
  const config = userInfo.configJson || {};
  const activeSheetName = config.activeSheetName || config.sheetName || '';
  
  const allSheets = getSheets();
  
  return {
    sheets: allSheets.map(sheetName => ({
      name: sheetName,
      isActive: sheetName === activeSheetName
    })),
    activeSheetName: activeSheetName
  };
}

/**
 * 新しいスプレッドシートURLを追加し、アクティブシートとして設定します。
 * @param {string} spreadsheetUrl - 追加するスプレッドシートのURL
 */
function addSpreadsheetUrl(spreadsheetUrl) {
  // セキュリティ強化: 権限チェックを最初に実行
  const { userId, userInfo } = validateCurrentUser();

  // URLからスプレッドシートIDを抽出（セキュリティ強化）
  let spreadsheetId;
  
  // URL長制限チェック
  if (!spreadsheetUrl || spreadsheetUrl.length > 500) {
    throw new Error('無効なURLです。');
  }
  
  // URLサニタイゼーション
  const sanitizedUrl = spreadsheetUrl.trim();
  
  // [修正] 正規表現リテラルのスラッシュを修正し、パターンをより正確にしました。
  const urlPattern = /^https:\/\/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9-_]{44})(?:\/.*)?$/;
  const match = sanitizedUrl.match(urlPattern);

  // [修正] 構文エラーの原因となっていた余分な括弧とelseブロックを削除しました。
  if (match) {
    spreadsheetId = match[1];
  } else {
    // [修正] Unicodeエスケープされていたエラーメッセージを通常の文字列にしました。
    throw new Error('有効なGoogleスプレッドシートのURLではありません。正しい形式: https://docs.google.com/spreadsheets/d/...');
  }

  // スプレッドシートにアクセスできるかテスト
  try {
    const testSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheets = testSpreadsheet.getSheets();

    if (sheets.length === 0) {
      throw new Error('スプレッドシートにシートが見つかりません。');
    }

    // ユーザー設定を更新：新しいスプレッドシートIDを設定
    const userDb = getDatabase().getSheetByName(USER_DB_CONFIG.SHEET_NAME);
    const data = userDb.getDataRange().getValues();
    const headers = data[0];
    const userIdIndex = headers.indexOf('userId');
    const spreadsheetIdIndex = headers.indexOf('spreadsheetId');
    const spreadsheetUrlIndex = headers.indexOf('spreadsheetUrl');

    for (let i = 1; i < data.length; i++) {
      if (data[i][userIdIndex] === userId) {
        // スプレッドシートIDとURLを更新
        userDb.getRange(i + 1, spreadsheetIdIndex + 1).setValue(spreadsheetId);
        userDb.getRange(i + 1, spreadsheetUrlIndex + 1).setValue(spreadsheetUrl);
        break;
      }
    }

    // ユーザープロパティも更新
    props.setProperty('CURRENT_SPREADSHEET_ID', spreadsheetId);

    // 最初のシートをアクティブシートとして設定
    const firstSheetName = sheets[0].getName();
    updateUserConfig(userId, {
      activeSheetName: firstSheetName
    });

    // 新しいシートの基本設定を自動作成
    try {
      // シートのヘッダーを取得して推測
      const firstSheet = testSpreadsheet.getSheetByName(firstSheetName);
      if (firstSheet && firstSheet.getLastRow() > 0) {
        const headers = firstSheet.getRange(1, 1, 1, firstSheet.getLastColumn()).getValues()[0];
        const guessedConfig = guessHeadersFromArray(headers);

        // 基本設定を保存（少なくとも1つのヘッダーが推測できた場合）
        if (guessedConfig.questionHeader || guessedConfig.answerHeader) {
          saveSheetConfig(firstSheetName, guessedConfig);
        }
      }
    } catch (configError) {
      // セキュリティ向上: エラー詳細を本番環境では表示しない
      console.warn('Failed to auto-create config for new sheet');
      // 設定作成に失敗してもスプレッドシート追加は成功とする
    }

    auditLog('SPREADSHEET_ADDED', userId, {
      spreadsheetId,
      spreadsheetUrl,
      firstSheetName,
      userEmail: userInfo.adminEmail
    });

    return {
      success: true,
      message: `スプレッドシート「${testSpreadsheet.getName()}」が追加され、シート「${firstSheetName}」がアクティブになりました。`,
      spreadsheetId: spreadsheetId,
      firstSheetName: firstSheetName
    };

  } catch (error) {
    console.error('Failed to access spreadsheet:', error);
    throw new Error(`スプレッドシートにアクセスできません。URLが正しいか、共有設定を確認してください。エラー: ${error.message}`);
  }
}

/**
 * AdminPanel用のステータス取得関数
 */
function getStatus() {
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    const correctWebAppUrl = 'https://script.google.com/a/naha-okinawa.ed.jp/macros/s/AKfycbzFF3psxBRUja1DsrVDkleOGrUxar1QqxqGYwBVKmpcZybrtNddH5iKD-nbqmYWEZKK/exec';
    const correctDeployId = 'AKfycbzFF3psxBRUja1DsrVDkleOGrUxar1QqxqGYwBVKmpcZybrtNddH5iKD-nbqmYWEZKK';
    
    if (scriptProps.getProperty('WEB_APP_URL') !== correctWebAppUrl) {
      scriptProps.setProperty('WEB_APP_URL', correctWebAppUrl);
    }
    if (scriptProps.getProperty('DEPLOY_ID') !== correctDeployId) {
      scriptProps.setProperty('DEPLOY_ID', correctDeployId);
    }

    const settings = getAdminSettings();
    const props = PropertiesService.getUserProperties();
    const userId = props.getProperty('CURRENT_USER_ID');
    
    if (!userId) {
      return {
        activeSheetName: '',
        allSheets: [],
        answerCount: 0,
        totalReactions: 0,
        webAppUrl: getWebAppUrlEnhanced(),
        spreadsheetUrl: '',
        showNames: false,
        showCounts: false
      };
    }
    
    const userInfo = getUserInfo(userId);
    const config = userInfo ? userInfo.configJson || {} : {};
    const activeSheetName = config.activeSheetName || '';
    
    // Get available sheets
    let allSheets = [];
    try {
      allSheets = getSheets() || [];
    } catch (error) {
      console.warn('Failed to get sheets:', error);
    }
    
    // Get answer count if active sheet exists
    let answerCount = 0;
    let totalReactions = 0;
    
    if (activeSheetName) {
      try {
        const sheetData = getPublishedSheetData(activeSheetName);
        answerCount = sheetData.rows ? sheetData.rows.length : 0;
        totalReactions = sheetData.rows ? 
          sheetData.rows.reduce((sum, row) => {
            const reactions = row.reactions || {};
            return sum + (reactions.UNDERSTAND?.count || 0) + 
                       (reactions.LIKE?.count || 0) + 
                       (reactions.CURIOUS?.count || 0);
          }, 0) : 0;
      } catch (error) {
        console.warn('Failed to get sheet data for status:', error);
      }
    }
    
    return {
      activeSheetName: activeSheetName,
      allSheets: allSheets,
      answerCount: answerCount,
      totalReactions: totalReactions,
      webAppUrl: getWebAppUrlEnhanced(),
      spreadsheetUrl: userInfo ? userInfo.spreadsheetUrl || '' : '',
      showNames: typeof config.showNames !== 'undefined' ? config.showNames : (config.showDetails || false),
      showCounts: typeof config.showCounts !== 'undefined' ? config.showCounts : (config.showDetails || false)
    };
  } catch (error) {
    console.error('getStatus error:', error);
    throw new Error('ステータスの取得に失敗しました: ' + error.message);
  }
}

/**
 * AdminPanel用のシート切り替え関数（エイリアス）
 */
function switchToSheet(sheetName) {
  return switchActiveSheet(sheetName);
}


// =================================================================
// 内部処理関数
// =================================================================
function getSheets() {
  try {
    const allSheets = getCurrentSpreadsheet().getSheets();
    const visibleSheets = allSheets.filter(sheet => !sheet.isSheetHidden());
    const filtered = visibleSheets.filter(sheet => {
      const name = sheet.getName();
      return name !== 'Config';
    });
    return filtered.map(sheet => sheet.getName());
  } catch (error) {
    handleError('getSheets', error);
  }
}

function getSheetHeaders(sheetName) {
  const sheet = getCurrentSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`シート '${sheetName}' が見つかりません。`);
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headerRow.map(v => (v !== null && v !== undefined) ? String(v) : '');
}

/**
 * シートのヘッダー情報と自動設定を同時に返す新しい関数
 * フロントエンドでの重複したヘッダー推測ロジックを削除し、バックエンドで統一処理
 */
function getSheetHeadersWithAutoConfig(sheetName) {
  try {
    debugLog('Getting headers and auto config for sheet:', sheetName);
    
    const headers = getSheetHeaders(sheetName);
    if (!headers || headers.length === 0) {
      return {
        status: 'error',
        message: 'ヘッダー情報を取得できませんでした。'
      };
    }
    
    const autoConfig = guessHeadersFromArray(headers);
    
    debugLog('Headers:', headers);
    debugLog('Auto config:', autoConfig);
    
    return {
      status: 'success',
      headers: headers,
      autoConfig: autoConfig
    };
  } catch (error) {
    console.error('Error in getSheetHeadersWithAutoConfig:', error);
    return handleError('getSheetHeadersWithAutoConfig', error, true);
  }
}

/**
 * ヘッダー配列から設定項目を推測する関数
 *
 * Keywords cover common variants (e.g. コメント vs 回答) and include
 * both full/half-width characters. Header strings are normalized by
 * trimming/collapsing whitespace and converting full-width ASCII before
 * searching.
 *
 * @param {Array} headers - ヘッダー名の配列
 * @returns {Object} 推測された設定オブジェクト
 */
function guessHeadersFromArray(headers) {
  // Normalize header strings: convert full-width ASCII, collapse whitespace and lowercase
  const normalize = (str) => String(str || '')
    .replace(/[\uFF01-\uFF5E]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

  const normalizedHeaders = headers.map(h => ({ original: String(h), normalized: normalize(h) }));

  // Find the first header that contains any of the provided keywords
  const find = (keys) => {
    const keyNorm = keys.map(k => normalize(k));
    const found = normalizedHeaders.find(h => keyNorm.some(k => h.normalized.includes(k)));
    return found ? found.original : '';
  };
  
  debugLog('Available headers:', headers);
  
  // Googleフォーム特有のヘッダー構造に対応
  const isGoogleForm = normalizedHeaders.some(h =>
    h.normalized.includes(normalize('タイムスタンプ')) ||
    h.normalized.includes(normalize('メールアドレス'))
  );
  
  let question = '';
  let answer = '';
  let reason = '';
  let name = '';
  let classHeader = '';
  
  if (isGoogleForm) {
    debugLog('Detected Google Form response sheet');
    
    // Googleフォームの一般的な構造: タイムスタンプ, メールアドレス, [質問1], [質問2], ...
    const nonMetaHeaders = normalizedHeaders.filter(h => {
      const hStr = h.normalized;
      return !hStr.includes(normalize('タイムスタンプ')) &&
             !hStr.includes('timestamp') &&
             !hStr.includes(normalize('メールアドレス')) &&
             !hStr.includes('email');
    });
    
    debugLog('Non-meta headers:', nonMetaHeaders.map(h => h.original));
    
    // より柔軟な推測ロジック
    for (let i = 0; i < nonMetaHeaders.length; i++) {
      const { original: header, normalized: hStr } = nonMetaHeaders[i];
      
      // 質問を含む長いテキストを質問ヘッダーとして推測
      if (!question && (hStr.includes('だろうか') || hStr.includes('ですか') || hStr.includes('でしょうか') || hStr.length > 20)) {
        question = header;
        // 同じ内容が複数列にある場合、回答用として2番目を使用
        if (i + 1 < nonMetaHeaders.length && nonMetaHeaders[i + 1].original === header) {
          answer = header;
          continue;
        }
      }
      
      // 回答・意見に関するヘッダー
      if (!answer && (hStr.includes('回答') || hStr.includes('答え') || hStr.includes('意見') || hStr.includes('考え'))) {
        answer = header;
      }
      
      // 理由に関するヘッダー
      if (!reason && (hStr.includes('理由') || hStr.includes('詳細') || hStr.includes('説明'))) {
        reason = header;
      }
      
      // 名前に関するヘッダー
      if (!name && (hStr.includes('名前') || hStr.includes('氏名') || hStr.includes('学生'))) {
        name = header;
      }
      
      // クラスに関するヘッダー
      if (!classHeader && (hStr.includes('クラス') || hStr.includes('組') || hStr.includes('学級'))) {
        classHeader = header;
      }
    }
    
    // フォールバック: まだ回答が見つからない場合
    if (!answer && nonMetaHeaders.length > 0) {
      // 最初の非メタヘッダーを回答として使用
      answer = nonMetaHeaders[0].original;
    }
    
  } else {
    // 通常のシート用の推測ロジック
    // Keyword lists include common variants and full/half-width characters
    question = find(['質問', '質問内容', '問題', '問い', 'お題', 'question', 'question text', 'q', 'Ｑ']);
    answer = find(['回答', '解答', '答え', '返答', 'answer', 'a', 'Ａ', '意見', 'コメント', 'opinion', 'response']);
    reason = find(['理由', '訳', 'わけ', '根拠', '詳細', '説明', 'comment', 'why', 'reason', 'detail']);
    name = find(['名前', '氏名', 'name', '学生', 'student', 'ペンネーム', 'ニックネーム', 'author']);
    classHeader = find(['クラス', 'class', 'classroom', '組', '学級', '班', 'グループ']);
  }
  
  const guessed = {
    questionHeader: question,
    answerHeader: answer,
    reasonHeader: reason,
    nameHeader: name,
    classHeader: classHeader
  };
  
  debugLog('Guessed headers:', guessed);
  
  // 最終フォールバック: 何も推測できない場合
  if (!question && !answer && headers.length > 0) {
    debugLog('No specific headers found, using positional mapping');
    
    // タイムスタンプとメールを除外して最初の列を回答として使用
    const usableHeaders = normalizedHeaders
      .filter(h => {
        const hStr = h.normalized;
        return !hStr.includes(normalize('タイムスタンプ')) &&
               !hStr.includes('timestamp') &&
               !hStr.includes(normalize('メールアドレス')) &&
               !hStr.includes('email') &&
               h.normalized !== '';
      })
      .map(h => h.original);
    
    if (usableHeaders.length > 0) {
      guessed.answerHeader = usableHeaders[0];
      if (usableHeaders.length > 1) guessed.reasonHeader = usableHeaders[1];
      if (usableHeaders.length > 2) guessed.nameHeader = usableHeaders[2];
      if (usableHeaders.length > 3) guessed.classHeader = usableHeaders[3];
    }
  }
  
  return guessed;
}

/**
 * Minimal board data builder used in tests
 */
function buildBoardData(sheetName) {
  const sheet = getCurrentSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`指定されたシート「${sheetName}」が見つかりません。`);

  const values = sheet.getDataRange().getValues();
  if (values.length < 1) return { header: '', entries: [] };

  const cfgFunc = (typeof global !== 'undefined' && global.getConfig)
    ? global.getConfig
    : (typeof getConfig === 'function' ? getConfig : null);
  const cfg = cfgFunc ? cfgFunc(sheetName) : {};

  const answerHeader = cfg.answerHeader || cfg.questionHeader || COLUMN_HEADERS.OPINION;
  const reasonHeader = cfg.reasonHeader || COLUMN_HEADERS.REASON;
  const classHeader = cfg.classHeader || COLUMN_HEADERS.CLASS;
  const nameHeader = cfg.nameHeader || '';

  const required = [answerHeader, reasonHeader, classHeader, COLUMN_HEADERS.EMAIL];
  if (nameHeader) required.push(nameHeader);

  const headerIndices = findHeaderIndices(values[0], required);

  const entries = values.slice(1).map(row => ({
    answer: row[headerIndices[answerHeader]] || '',
    reason: row[headerIndices[reasonHeader]] || '',
    name: nameHeader ? row[headerIndices[nameHeader]] || '' : '',
    class: row[headerIndices[classHeader]] || ''
  }));

  return { header: cfg.questionHeader || answerHeader, entries: entries };
}

function getSheetData(sheetName, classFilter, sortBy) {
  try {
    const sheet = getCurrentSpreadsheet().getSheetByName(sheetName);
    if (!sheet) throw new Error(`指定されたシート「${sheetName}」が見つかりません。`);

    const allValues = sheet.getDataRange().getValues();
    if (allValues.length < 1) return { header: "シートにデータがありません", rows: [] };

    const cfgFunc = (typeof global !== 'undefined' && global.getConfig)
      ? global.getConfig
      : (typeof getConfig === 'function' ? getConfig : null);
    
    let cfg = {};
    try {
      cfg = cfgFunc ? cfgFunc(sheetName) : {};
    } catch (configError) {
      debugLog(`Config not found for sheet ${sheetName} in getSheetData, creating default config:`, configError.message);
      
      // ヘッダーから自動推測して基本設定を作成
      const headers = allValues[0] || [];
      const guessedConfig = guessHeadersFromArray(headers);
      
      // 少なくとも1つのヘッダーが推測できた場合、デフォルト設定を保存
      if (guessedConfig.questionHeader || guessedConfig.answerHeader) {
        try {
          saveSheetConfig(sheetName, guessedConfig);
          cfg = guessedConfig;
          debugLog('Auto-created and saved config for sheet in getSheetData:', sheetName, cfg);
        } catch (saveError) {
          console.warn('Failed to save auto-created config in getSheetData:', saveError.message);
          cfg = guessedConfig; // 保存に失敗してもメモリ上のconfigは使用
        }
      }
    }
    
    const answerHeader = cfg.answerHeader || cfg.questionHeader || COLUMN_HEADERS.OPINION;
    const reasonHeader = cfg.reasonHeader || COLUMN_HEADERS.REASON;
    const classHeaderGuess = cfg.classHeader || COLUMN_HEADERS.CLASS;
    const nameHeader = cfg.nameHeader || '';

    const sheetHeaders = allValues[0].map(h => String(h || '').replace(/\s+/g,''));
    const normalized = header => String(header || '').replace(/\s+/g,'');

    const required = [
      COLUMN_HEADERS.EMAIL,
      answerHeader,
      reasonHeader,
      COLUMN_HEADERS.TIMESTAMP,
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ];
    if (nameHeader && sheetHeaders.includes(normalized(nameHeader))) {
      required.push(nameHeader);
    }

    if (sheetHeaders.includes(normalized(classHeaderGuess))) {
      required.splice(1,0,classHeaderGuess);
    }

    const headerIndices = findHeaderIndices(allValues[0], required);

    const classHeader = sheetHeaders.includes(normalized(classHeaderGuess)) ? classHeaderGuess : '';
    const dataRows = allValues.slice(1);
    const userEmail = safeGetUserEmail();
    const isAdmin = isUserAdmin(userEmail);

    const rows = dataRows.map((row, index) => {
      if (classHeader && classFilter && classFilter !== 'すべて') {
        const className = row[headerIndices[classHeader]];
        if (className !== classFilter) return null;
      }
      const email = row[headerIndices[COLUMN_HEADERS.EMAIL]];
      const opinion = row[headerIndices[answerHeader]];

      if (email && opinion) {
        const understandArr = parseReactionString(row[headerIndices[COLUMN_HEADERS.UNDERSTAND]]);
        const likeArr = parseReactionString(row[headerIndices[COLUMN_HEADERS.LIKE]]);
        const curiousArr = parseReactionString(row[headerIndices[COLUMN_HEADERS.CURIOUS]]);
        const reason = row[headerIndices[reasonHeader]] || '';
        const highlightVal = row[headerIndices[COLUMN_HEADERS.HIGHLIGHT]];
        const timestampRaw = row[headerIndices[COLUMN_HEADERS.TIMESTAMP]];
        const timestamp = timestampRaw ? new Date(timestampRaw).toISOString() : '';
        const likes = likeArr.length;
        const baseScore = reason.length;
        const likeMultiplier = 1 + (likes * SCORING_CONFIG.LIKE_MULTIPLIER_FACTOR);
        const totalScore = baseScore * likeMultiplier;
        let name = '';
        if (nameHeader && row[headerIndices[nameHeader]]) {
          name = sanitizeInput(row[headerIndices[nameHeader]]);
        } else if (email) {
          name = sanitizeInput(email.split('@')[0]);
        }
        return {
          rowIndex: index + 2,
          timestamp: timestamp,
          name: name, // 常に名前を含める（クライアント側で制御）
          email: email, // 常にemailを含める（クライアント側で制御）
          class: sanitizeInput(row[headerIndices[classHeader]] || '未分類'),
          opinion: sanitizeInput(opinion),
          reason: sanitizeInput(reason),
          reactions: {
            UNDERSTAND: { count: understandArr.length, reacted: userEmail ? understandArr.includes(userEmail) : false },
            LIKE: { count: likeArr.length, reacted: userEmail ? likeArr.includes(userEmail) : false },
            CURIOUS: { count: curiousArr.length, reacted: userEmail ? curiousArr.includes(userEmail) : false }
          },
          highlight: highlightVal === true || String(highlightVal).toLowerCase() === 'true',
          score: totalScore
        };
      }
      return null;
    }).filter(Boolean);

    if (sortBy === 'newest') {
      rows.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    } else if (sortBy === 'random') {
      rows.sort((a, b) => hashTimestamp(a.timestamp) - hashTimestamp(b.timestamp));
    } else {
      rows.sort((a, b) => b.score - a.score);
    }
    const header = cfg.questionHeader || answerHeader;
    return { header: header, rows: rows };
  } catch(e) {
    handleError(`getSheetDataForSpreadsheet for ${sheetName}`, e);
  }
}

/**
 * 指定したスプレッドシートからシートデータを取得します。
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet 対象スプレッドシート
 * @param {string} sheetName シート名
 * @param {string} classFilter クラスフィルター
 * @param {string} sortBy ソート順
 */
function getSheetDataForSpreadsheet(spreadsheet, sheetName, classFilter, sortBy) {
  try {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) throw new Error(`指定されたシート「${sheetName}」が見つかりません。`);

    const allValues = sheet.getDataRange().getValues();
    if (allValues.length < 1) return { header: "シートにデータがありません", rows: [] };

    const cfgFunc = (typeof global !== 'undefined' && global.getConfig)
      ? global.getConfig
      : (typeof getConfig === 'function' ? getConfig : null);

    let cfg = {};
    try {
      cfg = cfgFunc ? cfgFunc(sheetName) : {};
    } catch (configError) {
      debugLog(`Config not found for sheet ${sheetName} in getSheetDataForSpreadsheet, creating default config:`, configError.message);

      const headers = allValues[0] || [];
      const guessedConfig = guessHeadersFromArray(headers);

      if (guessedConfig.questionHeader || guessedConfig.answerHeader) {
        try {
          saveSheetConfig(sheetName, guessedConfig);
          cfg = guessedConfig;
          debugLog('Auto-created and saved config for sheet in getSheetDataForSpreadsheet:', sheetName, cfg);
        } catch (saveError) {
          console.warn('Failed to save auto-created config in getSheetDataForSpreadsheet:', saveError.message);
          cfg = guessedConfig;
        }
      }
    }

    const answerHeader = cfg.answerHeader || cfg.questionHeader || COLUMN_HEADERS.OPINION;
    const reasonHeader = cfg.reasonHeader || COLUMN_HEADERS.REASON;
    const classHeaderGuess = cfg.classHeader || COLUMN_HEADERS.CLASS;
    const nameHeader = cfg.nameHeader || '';

    const sheetHeaders = allValues[0].map(h => String(h || '').replace(/\s+/g,''));
    const normalized = header => String(header || '').replace(/\s+/g,'');

    const required = [
      COLUMN_HEADERS.EMAIL,
      answerHeader,
      reasonHeader,
      COLUMN_HEADERS.TIMESTAMP,
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ];
    if (nameHeader && sheetHeaders.includes(normalized(nameHeader))) {
      required.push(nameHeader);
    }

    if (sheetHeaders.includes(normalized(classHeaderGuess))) {
      required.splice(1,0,classHeaderGuess);
    }

    const headerIndices = findHeaderIndices(allValues[0], required);

    const classHeader = sheetHeaders.includes(normalized(classHeaderGuess)) ? classHeaderGuess : '';
    const dataRows = allValues.slice(1);
    const userEmail = safeGetUserEmail();

    const rows = dataRows.map((row, index) => {
      if (classHeader && classFilter && classFilter !== 'すべて') {
        const className = row[headerIndices[classHeader]];
        if (className !== classFilter) return null;
      }
      const email = row[headerIndices[COLUMN_HEADERS.EMAIL]];
      const opinion = row[headerIndices[answerHeader]];

      if (email && opinion) {
        const understandArr = parseReactionString(row[headerIndices[COLUMN_HEADERS.UNDERSTAND]]);
        const likeArr = parseReactionString(row[headerIndices[COLUMN_HEADERS.LIKE]]);
        const curiousArr = parseReactionString(row[headerIndices[COLUMN_HEADERS.CURIOUS]]);
        const reason = row[headerIndices[reasonHeader]] || '';
        const highlightVal = row[headerIndices[COLUMN_HEADERS.HIGHLIGHT]];
        const timestampRaw = row[headerIndices[COLUMN_HEADERS.TIMESTAMP]];
        const timestamp = timestampRaw ? new Date(timestampRaw).toISOString() : '';
        const likes = likeArr.length;
        const baseScore = reason.length;
        const likeMultiplier = 1 + (likes * SCORING_CONFIG.LIKE_MULTIPLIER_FACTOR);
        const totalScore = baseScore * likeMultiplier;
        let name = '';
        if (nameHeader && row[headerIndices[nameHeader]]) {
          name = row[headerIndices[nameHeader]];
        } else if (email) {
          name = email.split('@')[0];
        }
        return {
          rowIndex: index + 2,
          timestamp: timestamp,
          name: name,
          email: email,
          class: row[headerIndices[classHeader]] || '未分類',
          opinion: opinion,
          reason: reason,
          reactions: {
            UNDERSTAND: { count: understandArr.length, reacted: userEmail ? understandArr.includes(userEmail) : false },
            LIKE: { count: likeArr.length, reacted: userEmail ? likeArr.includes(userEmail) : false },
            CURIOUS: { count: curiousArr.length, reacted: userEmail ? curiousArr.includes(userEmail) : false }
          },
          highlight: highlightVal === true || String(highlightVal).toLowerCase() === 'true',
          score: totalScore
        };
      }
      return null;
    }).filter(Boolean);

    if (sortBy === 'newest') {
      rows.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    } else if (sortBy === 'random') {
      rows.sort((a, b) => hashTimestamp(a.timestamp) - hashTimestamp(b.timestamp));
    } else {
      rows.sort((a, b) => b.score - a.score);
    }

    const header = cfg.questionHeader || answerHeader;
    return { header: header, rows: rows };
  } catch(e) {
    handleError(`getSheetDataForSpreadsheet for ${sheetName}`, e);
  }
}


function addReaction(rowIndex, reactionKey, sheetName) {
  // セキュリティ強化: 権限チェックを最初に実行
  try {
    const { userId, userInfo } = validateCurrentUser();
    // レート制限チェック
    checkRateLimit('addReaction', userInfo.adminEmail);
  } catch (error) {
    return { status: 'error', message: error.message };
  }
  
  if (!rowIndex || !reactionKey || !COLUMN_HEADERS[reactionKey]) {
    return { status: 'error', message: '無効なパラメータです。' };
  }
  
  // Use LockService only in production environment
  const lock = (typeof LockService !== 'undefined') ? LockService.getScriptLock() : null;
  try {
    if (lock && !lock.tryLock(TIME_CONSTANTS.LOCK_WAIT_MS)) {
      return { status: 'error', message: '他のユーザーが操作中です。しばらく待ってから再試行してください。' };
    }
    const userEmail = safeGetUserEmail();
    if (!userEmail) {
      return { status: 'error', message: 'ログインしていないため、操作できません。' };
    }
    
    // マルチテナント対応の修正
    const spreadsheet = getCurrentSpreadsheet();
    const settings = getAppSettingsForUser();
    const targetSheet = sheetName || settings.activeSheetName;
    const sheet = spreadsheet.getSheetByName(targetSheet);
    
    if (!sheet) throw new Error(`シート '${targetSheet}' が見つかりません。`);

  const reactionHeaders = REACTION_KEYS.map(k => COLUMN_HEADERS[k]);
  const headerIndices = getHeaderIndices(targetSheet);
  const startCol = headerIndices[reactionHeaders[0]] + 1;
  const reactionRange = sheet.getRange(rowIndex, startCol, 1, REACTION_KEYS.length);
  const values = reactionRange.getValues()[0];
  const lists = {};
  REACTION_KEYS.forEach((k, idx) => {
    lists[k] = { arr: parseReactionString(values[idx]) };
  });

    const wasReacted = lists[reactionKey].arr.includes(userEmail);
    Object.keys(lists).forEach(key => {
      const idx = lists[key].arr.indexOf(userEmail);
      if (idx > -1) lists[key].arr.splice(idx, 1);
    });
    if (!wasReacted) {
      lists[reactionKey].arr.push(userEmail);
    }
    reactionRange.setValues([
      REACTION_KEYS.map(k => lists[k].arr.join(','))
    ]);

    const reactions = REACTION_KEYS.reduce((obj, k) => {
      obj[k] = {
        count: lists[k].arr.length,
        reacted: lists[k].arr.includes(userEmail)
      };
      return obj;
    }, {});
    return { status: 'ok', reactions: reactions };
  } catch (error) {
    return handleError('addReaction', error, true);
  } finally {
    if (lock) {
      try { lock.releaseLock(); } catch (e) {}
    }
  }
}

function toggleHighlight(rowIndex, sheetName, userObject = null) {
  // セキュリティ強化: 権限チェックと rate limiting
  if (!checkAdmin()) {
    return { status: 'error', message: '権限がありません。' };
  }

  // ユーザーメール取得（ログチェック用）
  const userEmail = safeGetUserEmail();
  if (!userEmail) {
    return { status: 'error', message: 'ログインしていないため、操作できません。' };
  }

  // レート制限チェック
  try {
    checkRateLimit('toggleHighlight', userEmail);
  } catch (error) {
    return { status: 'error', message: error.message };
  }

  debugLog('toggleHighlight processing', { rowIndex: rowIndex, sheetName: sheetName, userEmail: userEmail });

  const lock = (typeof LockService !== 'undefined') ? LockService.getScriptLock() : null;
  try {
    if (lock) {
      lock.waitLock(TIME_CONSTANTS.LOCK_WAIT_MS);
    }
    const settings = getAppSettingsForUser();
    const targetSheet = sheetName || settings.activeSheetName;
    const sheet = getCurrentSpreadsheet().getSheetByName(targetSheet);
    if (!sheet) throw new Error(`シート '${targetSheet}' が見つかりません。`);

    const headerIndices = getHeaderIndices(targetSheet);
    let highlightColIndex = headerIndices[COLUMN_HEADERS.HIGHLIGHT];
    
    // ハイライト列が存在しない場合は追加する
    if (highlightColIndex === undefined || highlightColIndex === -1) {
      prepareSheetForBoard(targetSheet);
      // 再度ヘッダーインデックスを取得
      const updatedHeaderIndices = getHeaderIndices(targetSheet);
      highlightColIndex = updatedHeaderIndices[COLUMN_HEADERS.HIGHLIGHT];
      
      if (highlightColIndex === undefined || highlightColIndex === -1) {
        throw new Error('ハイライト列が見つかりません。スプレッドシートにハイライト列が存在することを確認してください。');
      }
    }
    
    const colIndex = highlightColIndex + 1;

    const cell = sheet.getRange(rowIndex, colIndex);
    const current = !!cell.getValue();
    const newValue = !current;
    cell.setValue(newValue);
    debugLog('toggleHighlight updated', { rowIndex: rowIndex, sheetName: targetSheet, highlight: newValue });
    return { status: 'ok', highlight: newValue };
  } catch (error) {
    console.error('toggleHighlight failed:', error);
    return secureHandleError('toggleHighlight', error, true);
  } finally {
    if (lock) {
      lock.releaseLock();
    }
  }
}



function getAppSettingsForUser() {
  const props = PropertiesService.getUserProperties();
  const userId = props.getProperty('CURRENT_USER_ID');
  
  if (!userId) {
    // 従来の動作のフォールバック
    return {
      activeSheetName: '',
      showNames: false,
      showCounts: false
    };
  }
  
  // マルチテナント時
  const userInfo = getUserInfo(userId);
  if (!userInfo || !userInfo.configJson) {
    return {
      activeSheetName: '',
      showNames: false,
      showCounts: false
    };
  }
  
  // 下位互換性：古いsheetNameを新しいactiveSheetNameにマッピング
  const activeSheetName = userInfo.configJson.activeSheetName || userInfo.configJson.sheetName || '';
  
  return {
    activeSheetName: activeSheetName,
    showNames: typeof userInfo.configJson.showNames !== 'undefined' ? userInfo.configJson.showNames : (userInfo.configJson.showDetails || false),
    showCounts: typeof userInfo.configJson.showCounts !== 'undefined' ? userInfo.configJson.showCounts : (userInfo.configJson.showDetails || false)
  };
}

function getHeaderIndices(sheetName) {
  const cache = (typeof CacheService !== 'undefined') ? CacheService.getScriptCache() : null;
  const cacheKey = `headers_${sheetName}`;
  
  try {
    if (cache) {
      const cached = cache.get(cacheKey);
      if (cached) {
        return JSON.parse(cached);
      }
    }
  } catch (e) {
    console.error('Cache retrieval failed:', e);
  }
  
  const sheet = getCurrentSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`シート '${sheetName}' が見つかりません。`);
  
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const indices = findHeaderIndices(headerRow, [
    COLUMN_HEADERS.TIMESTAMP,
    COLUMN_HEADERS.UNDERSTAND,
    COLUMN_HEADERS.LIKE,
    COLUMN_HEADERS.CURIOUS,
    COLUMN_HEADERS.HIGHLIGHT
  ]);
  
  try {
    if (cache) {
      cache.put(cacheKey, JSON.stringify(indices), 21600); // 6 hours cache
    }
  } catch (e) {
    console.error('Cache storage failed:', e);
  }
  
  return indices;
}

function clearHeaderCache(sheetName) {
  const cache = (typeof CacheService !== 'undefined') ? CacheService.getScriptCache() : null;
  if (cache) {
    try {
      cache.remove(`headers_${sheetName}`);
    } catch (e) {
      console.error('Cache removal failed:', e);
    }
  }
}


function prepareSheetForBoard(sheetName) {
  const sheet = getCurrentSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`シート '${sheetName}' が見つかりません。`);
  const headers = getSheetHeaders(sheetName);
  let lastCol = headers.length;
  const required = [
    COLUMN_HEADERS.TIMESTAMP,
    COLUMN_HEADERS.UNDERSTAND,
    COLUMN_HEADERS.LIKE,
    COLUMN_HEADERS.CURIOUS,
    COLUMN_HEADERS.HIGHLIGHT
  ];
  required.forEach(h => {
    if (headers.indexOf(h) === -1) {
      sheet.insertColumnAfter(lastCol);
      lastCol++;
      sheet.getRange(1, lastCol).setValue(h);
      headers.push(h);
    }
  });
  clearHeaderCache(sheetName);
}

function findHeaderIndices(sheetHeaders, requiredHeaders) {
  const indices = {};
  const normalized = sheetHeaders.map(h => (typeof h === 'string' ? h.replace(/\s+/g, '') : h));
  const missingHeaders = [];
  requiredHeaders.forEach(name => {
    const idx = normalized.indexOf(name.replace(/\s+/g, ''));
    if (idx !== -1) { indices[name] = idx; } else { missingHeaders.push(name); }
  });
  if (missingHeaders.length > 0) {
    throw new Error(`必須ヘッダーが見つかりません: [${missingHeaders.join(', ')}]`);
  }
  return indices;
}

function saveWebAppUrl(url) {
  const props = PropertiesService.getScriptProperties();
  props.setProperties({ WEB_APP_URL: (url || '').trim() });
}

function getUrlOrigin(url) {
  const match = String(url).match(/^(https?://[^/]+)/);
  return match ? match[1] : '';
}

function extractDeployIdFromUrl(url) {
  if (!url) return null;
  
  // script.google.com/macros/s/{DEPLOY_ID}/exec形式からDEPLOY_IDを抽出
  const macrosMatch = url.match(/\/macros\/s\/([a-zA-Z0-9_-]+)/);
  if (macrosMatch) {
    return macrosMatch[1];
  }
  
  // script.googleusercontent.com形式のURLの場合、URL全体から推測を試行
  if (/script\.googleusercontent\.com/.test(url)) {
    // URLのハッシュ部分やパラメータからDEPLOY_IDの可能性があるものを探す
    const hashMatch = url.match(/#gid=([a-zA-Z0-9_-]+)/);
    if (hashMatch && hashMatch[1].length > 10) {
      return hashMatch[1];
    }
  }
  
  return null;
}

// 古いgetWebAppUrl関数は削除されました。getWebAppUrlEnhanced()を使用してください。



function getWebAppUrlEnhanced() {
  const props = PropertiesService.getScriptProperties();
  let stored = (props.getProperty('WEB_APP_URL') || '').trim();
  if (stored) {
    return stored;
  }
  try {
    if (typeof ScriptApp !== 'undefined') {
      return ScriptApp.getService().getUrl();
    }
  } catch (e) {
    // ScriptAppが利用できない環境（例: テスト環境）の場合
    return '';
  }
  return '';
}

// レガシー互換のためのエイリアス関数
function getWebAppUrl() {
  return getWebAppUrlEnhanced();
}

function saveDeployId(id) {
  const props = PropertiesService.getScriptProperties();
  const cleanId = (id || '').trim();
  
  // DEPLOY_ID形式の検証
  if (cleanId && !/^[a-zA-Z0-9_-]{10,}$/.test(cleanId)) {
    console.warn('Invalid DEPLOY_ID format:', cleanId);
    throw new Error('無効なDEPLOY_ID形式です');
  }
  
  if (cleanId) {
    props.setProperties({ DEPLOY_ID: cleanId });
    debugLog('Saved DEPLOY_ID:', cleanId);
    
    // DEPLOY_IDが設定された後、WebAppURLを再評価
    const currentUrl = getWebAppUrlEnhanced();
    debugLog('Updated WebApp URL after DEPLOY_ID save:', currentUrl);
    
    const storedUrl = props.getProperty('WEB_APP_URL');
    if (storedUrl && /script\.googleusercontent\.com/.test(storedUrl)) {
      props.setProperties({ WEB_APP_URL: getWebAppUrlEnhanced() });
    }
  } else {
    console.warn('Cannot save empty DEPLOY_ID');
  }
}

// =================================================================
// マルチテナント管理機能
// =================================================================

/**
 * 新規ユーザーを登録
 */
function checkRateLimit(action, userEmail) {
  try {
    const key = `rateLimit_${action}_${userEmail}`;
    const cache = (typeof CacheService !== 'undefined') ? CacheService.getScriptCache() : null;
    const attempts = cache ? parseInt(cache.get(key) || '0') : 0;
    
    if (attempts > 10) { // 10 attempts per hour
      throw new Error('レート制限に達しました。しばらく待ってから再試行してください。');
    }
    
    if (cache) {
      cache.put(key, String(attempts + 1), 3600);
    }
  } catch (error) {
    // Cache service error - continue without rate limiting
    console.warn('Rate limiting failed:', error);
  }
}

function createStudyQuestForm(userEmail, userId) {
  try {
    // FormAppとDriveAppの利用可能性をチェック
    if (typeof FormApp === 'undefined' || typeof DriveApp === 'undefined') {
      throw new Error('Google Forms API または Drive API がこの環境で利用できません。');
    }
    
    // 新しいGoogleフォームを作成（作成日時を含む）
    const now = new Date();
    const dateTimeString = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
    const form = FormApp.create(`StudyQuest - みんなの回答ボード - ${userEmail.split('@')[0]} - ${dateTimeString}`);
    
    // メールアドレス収集を有効にする
    form.setCollectEmail(true);
    form.setRequireLogin(true);
    
    // フォームの説明と回答後のメッセージを設定
    form.setDescription('StudyQuestで使用する回答フォームです。質問に対する回答を入力してください。');
    const boardUrl = `${getWebAppUrlEnhanced()}?userId=${userId}`;
    form.setConfirmationMessage(`回答ありがとうございます！\n\nみんなの回答ボードで、他の人の意見も見てみましょう。\n${boardUrl}`);

    // クラス入力項目を追加し、入力形式を制限
    const classItem = form.addTextItem();
    classItem.setTitle('クラス名');
    classItem.setRequired(true);
    const pattern = '^[A-Za-z0-9]+-[A-Za-z0-9]+$';

    const helpText = "「G1-1」のように、学年と組を半角ハイフンで区切って入力してください。";
    const textValidation = FormApp.createTextValidation()
      .setHelpText(helpText)
      .requireTextMatchesPattern(pattern)
      .build();
    classItem.setValidation(textValidation);

    // その他の項目を追加
    form.addTextItem().setTitle('名前').setRequired(true);
    form.addParagraphTextItem().setTitle('回答').setHelpText('質問に対するあなたの回答を入力してください').setRequired(true);
    form.addParagraphTextItem().setTitle('理由').setHelpText('その回答を選んだ理由を教えてください').setRequired(false);
    
    // フォームの回答先スプレッドシートを作成（作成日時を含む）
    const spreadsheet = SpreadsheetApp.create(`StudyQuest - みんなの回答ボード - 回答データ - ${userEmail.split('@')[0]} - ${dateTimeString}`);
    form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());

    // 共有設定を自動化
    const userDomain = getEmailDomain(userEmail);
    if (userDomain) {
      const formFile = DriveApp.getFileById(form.getId());
      formFile.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.EDIT);
      const spreadsheetFile = DriveApp.getFileById(spreadsheet.getId());
      spreadsheetFile.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.EDIT);
      debugLog(`フォームとスプレッドシートをドメイン「${userDomain}」で共有しました。`);
    }

    const sheet = spreadsheet.getSheets()[0];
    
    // フォームとスプレッドシートの連携を一度確立してからヘッダーを調整
    // これによりGoogleフォームが自動的に基本ヘッダー（タイムスタンプ、メールアドレス、その他の質問項目）を作成する
    Utilities.sleep(2000); // フォーム連携の完了を待つ
    
    // 現在のヘッダーを取得（Googleフォームによって自動生成されたもの）
    const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    debugLog('Current headers after form linking:', currentHeaders);
    
    // StudyQuest用の追加列を準備
    const additionalHeaders = [
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ];
    
    // 追加の列を既存のヘッダーの後に挿入
    const startCol = currentHeaders.length + 1;
    sheet.getRange(1, startCol, 1, additionalHeaders.length).setValues([additionalHeaders]);
    
    // ヘッダー行のフォーマット
    const allHeadersRange = sheet.getRange(1, 1, 1, currentHeaders.length + additionalHeaders.length);
    allHeadersRange.setFontWeight('bold').setBackground('#E3F2FD');
    
    // 列幅を調整
    try {
      sheet.autoResizeColumns(1, allHeadersRange.getNumColumns());
    } catch (e) {
      console.warn('Auto-resize failed:', e);
    }
    
    // Configシートも作成
    prepareSpreadsheetForStudyQuest(spreadsheet);
    
    return {
      formId: form.getId(),
      formUrl: form.getPublishedUrl(),
      spreadsheetId: spreadsheet.getId(),
      spreadsheetUrl: spreadsheet.getUrl(),
      editFormUrl: form.getEditUrl()
    };

  } catch (error) {
    console.error('Failed to create form and spreadsheet:', error);
    
    // FormAppが利用できない場合はスプレッドシートのみ作成
    if (error.message.includes('Google Forms API is not available') || error.message.includes('FormApp')) {
      console.warn('FormApp not available, creating spreadsheet only');
      return createStudyQuestSpreadsheetFallback(userEmail);
    }
    
    // 権限エラーの場合
    if (error.message.includes('permission') || error.message.includes('Permission')) {
      throw new Error('Googleフォーム作成の権限がありません。管理者にお問い合わせください。');
    }
    
    // その他のエラー詳細を含める
    throw new Error(`Googleフォームとスプレッドシートの作成に失敗しました。詳細: ${error.message}`);
  }
 * Admin Panel用のボード作成関数
 * 現在ログイン中のユーザーの新しいボードを作成し、公開・アクティブ化まで行います。
 */
function createBoardFromAdmin() {
  if (!checkAdmin()) {
    throw new Error('権限がありません。');
  }
  try {
    const currentUserEmail = safeGetUserEmail();
    const props = PropertiesService.getUserProperties();
    const userId = props.getProperty('CURRENT_USER_ID');

    if (!userId) {
      throw new Error('ユーザーIDが見つかりません。ページをリロードして再試行してください。');
    }

    // 1. Googleフォームとスプレッドシートを作成（共有設定、回答後URL設定済み）
    const result = createStudyQuestForm(currentUserEmail, userId);
    
    // 2. 作成されたスプレッドシートを現在のユーザーのメインスプレッドシートとして設定
    if (result.spreadsheetId && result.spreadsheetUrl) {
      const addResult = addSpreadsheetUrl(result.spreadsheetUrl);
      debugLog('Spreadsheet added to user:', addResult);

      // 3. 新しく作成したシートをアクティブ化・公開
      const newSheetName = addResult.firstSheetName;
      if (newSheetName) {
        // アクティブシートに設定
        switchActiveSheet(newSheetName);
        // 公開状態に設定
        updateUserConfig(userId, { isPublished: true });
        debugLog(`New board '${newSheetName}' has been created and published.`);
      }

      return {
        ...result,
        message: '新しいボードが作成され、公開されました！管理ページを更新します。',
        autoCreated: true,
        newSheetName: newSheetName
      };
    }

    throw new Error('スプレッドシートの作成または追加に失敗しました。');

  } catch (error) {
    console.error('Failed to create board from admin:', error);
    return handleError('createBoardFromAdmin', error, true);
  }
}

function createStudyQuestSpreadsheetFallback(userEmail) {
  try {
    const spreadsheet = SpreadsheetApp.create(`StudyQuest - 回答データ - ${userEmail.split('@')[0]}`);
    const sheet = spreadsheet.getActiveSheet();
    
    const headers = [
      COLUMN_HEADERS.TIMESTAMP,
      COLUMN_HEADERS.EMAIL,
      COLUMN_HEADERS.CLASS,
      COLUMN_HEADERS.NAME,
      COLUMN_HEADERS.OPINION,
      COLUMN_HEADERS.REASON,
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ];
    sheet.appendRow(headers);
    
    const allHeadersRange = sheet.getRange(1, 1, 1, headers.length);
    allHeadersRange.setFontWeight('bold');
    allHeadersRange.setBackground('#E3F2FD');
    
    try {
      sheet.autoResizeColumns(1, headers.length);
    } catch (e) {
      console.warn('Auto-resize failed:', e);
    }
    
    prepareSpreadsheetForStudyQuest(spreadsheet);
    
    return {
      formId: null, // フォームは作成されない
      formUrl: null,
      spreadsheetId: spreadsheet.getId(),
      spreadsheetUrl: spreadsheet.getUrl(),
      editFormUrl: null,
      message: 'Googleフォームの作成に失敗したため、スプレッドシートのみ作成しました。'
    };
  } catch (error) {
    console.error('Failed to create spreadsheet fallback:', error);
    throw new Error(`スプレッドシートの作成に失敗しました。詳細: ${error.message}`);
  }
}


function prepareSpreadsheetForStudyQuest(spreadsheet) {
  // Configシートを作成または取得
  let configSheet = spreadsheet.getSheetByName('Config');
  if (!configSheet) {
    configSheet = spreadsheet.insertSheet('Config');
  }
  
  // Configシートに設定を書き込む
  const configData = [
    ['Key', 'Value'],
    ['WEB_APP_URL', getWebAppUrlEnhanced()], // 本番URLを保存
    ['DEPLOY_ID', PropertiesService.getScriptProperties().getProperty('DEPLOY_ID') || '']
  ];
  
  configSheet.getRange(1, 1, configData.length, configData[0].length).setValues(configData);
  
  // ヘッダーのスタイル設定
  configSheet.getRange('A1:B1').setFontWeight('bold').setBackground('#E3F2FD');
  
  // 列幅の自動調整
  try {
    configSheet.autoResizeColumns(1, 2);
  } catch (e) {
    console.warn('Config sheet auto-resize failed:', e);
  }
  
  // Configシートを非表示にする
  configSheet.hideSheet();
}

function getDatabase() {
  const props = PropertiesService.getScriptProperties();
  const dbId = props.getProperty('DATABASE_ID') || props.getProperty('USER_DATABASE_ID'); // 後方互換性
  if (!dbId) {
    throw new Error('ユーザーデータベースが設定されていません。管理者に連絡してください。');
  }
  return SpreadsheetApp.openById(dbId);
}

function getUserInfo(userId) {
  const userDb = getDatabase().getSheetByName(USER_DB_CONFIG.SHEET_NAME);
  const data = userDb.getDataRange().getValues();
  const headers = data[0];
  const userIdIndex = headers.indexOf('userId');
  const adminEmailIndex = headers.indexOf('adminEmail');
  const spreadsheetIdIndex = headers.indexOf('spreadsheetId');
  const spreadsheetUrlIndex = headers.indexOf('spreadsheetUrl');
  const configJsonIndex = headers.indexOf('configJson');
  const lastAccessedAtIndex = headers.indexOf('lastAccessedAt');
  const isActiveIndex = headers.indexOf('isActive');

  for (let i = 1; i < data.length; i++) {
    if (data[i][userIdIndex] === userId) {
      const userInfo = {
        userId: data[i][userIdIndex],
        adminEmail: data[i][adminEmailIndex],
        spreadsheetId: data[i][spreadsheetIdIndex],
        spreadsheetUrl: data[i][spreadsheetUrlIndex],
        configJson: {},
        lastAccessedAt: data[i][lastAccessedAtIndex],
        isActive: data[i][isActiveIndex] === true
      };
      try {
        userInfo.configJson = JSON.parse(data[i][configJsonIndex] || '{}');
      } catch (e) {
        console.error('Failed to parse configJson for user:', userId, e);
        userInfo.configJson = {};
      }
      
      // 最終アクセス日時を更新
      userDb.getRange(i + 1, lastAccessedAtIndex + 1).setValue(new Date());
      
      return userInfo;
    }
  }
  return null;
}

function getUserInfoInternal(userId) {
  const userDb = getDatabase().getSheetByName(USER_DB_CONFIG.SHEET_NAME);
  const data = userDb.getDataRange().getValues();
  const headers = data[0];
  const userIdIndex = headers.indexOf('userId');
  const adminEmailIndex = headers.indexOf('adminEmail');
  const spreadsheetIdIndex = headers.indexOf('spreadsheetId');
  const spreadsheetUrlIndex = headers.indexOf('spreadsheetUrl');
  const configJsonIndex = headers.indexOf('configJson');
  const lastAccessedAtIndex = headers.indexOf('lastAccessedAt');
  const isActiveIndex = headers.indexOf('isActive');

  for (let i = 1; i < data.length; i++) {
    if (data[i][userIdIndex] === userId) {
      const userInfo = {
        userId: data[i][userIdIndex],
        adminEmail: data[i][adminEmailIndex],
        spreadsheetId: data[i][spreadsheetIdIndex],
        spreadsheetUrl: data[i][spreadsheetUrlIndex],
        configJson: {},
        lastAccessedAt: data[i][lastAccessedAtIndex],
        isActive: data[i][isActiveIndex] === true
      };
      try {
        userInfo.configJson = JSON.parse(data[i][configJsonIndex] || '{}');
      } catch (e) {
        console.error('Failed to parse configJson for user:', userId, e);
        userInfo.configJson = {};
      }
      
      // 最終アクセス日時を更新
      userDb.getRange(i + 1, lastAccessedAtIndex + 1).setValue(new Date());
      
      return userInfo;
    }
  }
  return null;
}


function updateUserConfig(userId, newConfig) {
  const userDb = getDatabase().getSheetByName(USER_DB_CONFIG.SHEET_NAME);
  const data = userDb.getDataRange().getValues();
  const headers = data[0];
  const userIdIndex = headers.indexOf('userId');
  const configJsonIndex = headers.indexOf('configJson');

  for (let i = 1; i < data.length; i++) {
    if (data[i][userIdIndex] === userId) {
      let currentConfig = {};
      try {
        currentConfig = JSON.parse(data[i][configJsonIndex] || '{}');
      } catch (e) {
        console.error('Failed to parse existing configJson for user:', userId, e);
      }
      const updatedConfig = Object.assign({}, currentConfig, newConfig);
      userDb.getRange(i + 1, configJsonIndex + 1).setValue(JSON.stringify(updatedConfig));
      return true;
    }
  }
  return false;
}

function saveSheetConfig(sheetName, config) {
  const props = PropertiesService.getUserProperties();
  const userId = props.getProperty('CURRENT_USER_ID');
  
  if (!userId) {
    throw new Error('ユーザー情報が見つかりません。');
  }
  
  const userInfo = getUserInfo(userId);
  if (!userInfo || userInfo.adminEmail !== Session.getActiveUser().getEmail()) {
    throw new Error('権限がありません。');
  }
  
  // ユーザー設定を更新
  try {
    const currentConfig = userInfo.configJson || {};
    const updatedConfig = Object.assign({}, currentConfig, {
      sheetConfigs: {
        ...(currentConfig.sheetConfigs || {}),
        [sheetName]: config
      }
    });
    updateUserConfig(userId, updatedConfig);
  } catch (error) {
    console.error('Failed to update user config with sheet config:', error);
    throw new Error(`シート設定の保存に失敗しました: ${error.message}`);
  }
  
  auditLog('SHEET_CONFIG_SAVED', userId, { sheetName, config, userEmail: userInfo.adminEmail });
  
  return 'シート設定を保存しました。';
}

function getConfig(sheetName) {
  const props = PropertiesService.getUserProperties();
  const userId = props.getProperty('CURRENT_USER_ID');
  
  if (!userId) {
    throw new Error('ユーザー情報が見つかりません。');
  }
  
  const userInfo = getUserInfo(userId);
  if (!userInfo) {
    throw new Error('無効なユーザーです。');
  }
  
  const sheetConfigs = userInfo.configJson && userInfo.configJson.sheetConfigs;
  if (sheetConfigs && sheetConfigs[sheetName]) {
    return sheetConfigs[sheetName];
  }
  
  throw new Error(`シート「${sheetName}」の設定が見つかりません。`);
}

function auditLog(action, userId, details = {}) {
  try {
    const logSheet = getDatabase().getSheetByName('AuditLog');
    if (!logSheet) {
      const ss = getDatabase();
      const newLogSheet = ss.insertSheet('AuditLog');
      newLogSheet.appendRow(['Timestamp', 'UserId', 'Action', 'Details']);
      newLogSheet.getRange('A1:D1').setFontWeight('bold').setBackground('#F0F4C3');
      newLogSheet.autoResizeColumns(1, 4);
    }
    
    const row = [new Date(), userId, action, JSON.stringify(details)];
    logSheet.appendRow(row);
  } catch (e) {
    console.error('Failed to write audit log:', e);
  }
}

function registerNewUser(adminEmail) {
  checkRateLimit('registerNewUser', adminEmail);
  
  const userDb = getDatabase().getSheetByName(USER_DB_CONFIG.SHEET_NAME);
  const data = userDb.getDataRange().getValues();
  const headers = data[0];
  const adminEmailIndex = headers.indexOf('adminEmail');
  
  // 既に登録済みの管理者メールアドレスかチェック
  for (let i = 1; i < data.length; i++) {
    if (data[i][adminEmailIndex] === adminEmail) {
      throw new Error('このメールアドレスは既に登録されています。');
    }
  }
  
  // 新しいユーザーIDを生成
  const userId = Utilities.getUuid();
  
  // Googleフォームとスプレッドシートを作成
  const formAndSsInfo = createStudyQuestForm(adminEmail);
  
  const newRow = [
    userId,
    adminEmail,
    formAndSsInfo.spreadsheetId,
    formAndSsInfo.spreadsheetUrl,
    new Date(),
    '', // accessToken (未使用)
    JSON.stringify({}), // configJson
    new Date(),
    true // isActive
  ];
  
  userDb.appendRow(newRow);
  
  auditLog('NEW_USER_REGISTERED', userId, { adminEmail, spreadsheetId: formAndSsInfo.spreadsheetId });
  
  return {
    userId: userId,
    spreadsheetId: formAndSsInfo.spreadsheetId,
    spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
    message: '新しいユーザーが登録されました。管理画面に移動します。'
  };
}

function getSpreadsheetUrlForUser(userId) {
  const userInfo = getUserInfo(userId);
  if (userInfo && userInfo.spreadsheetUrl) {
    return userInfo.spreadsheetUrl;
  }
  throw new Error('ユーザーのスプレッドシートURLが見つかりません。');
}

function openActiveSpreadsheet() {
  const props = PropertiesService.getUserProperties();
  const userId = props.getProperty('CURRENT_USER_ID');
  if (!userId) {
    throw new Error('ユーザー情報が見つかりません。');
  }
  const userInfo = getUserInfo(userId);
  if (!userInfo || !userInfo.spreadsheetId) {
    throw new Error('アクティブなスプレッドシートが見つかりません。');
  }
  const spreadsheet = SpreadsheetApp.openById(userInfo.spreadsheetId);
  return spreadsheet.getUrl();
}

/**
 * 現在のユーザーに紐づく既存ボード情報を取得
 * @return {Object|null} 既存ボードがあればURL等を含むオブジェクト、なければnull
 */
function getExistingBoard() {
  const email = safeGetUserEmail();
  const userDb = getDatabase().getSheetByName(USER_DB_CONFIG.SHEET_NAME);
  const data = userDb.getDataRange().getValues();
  const headers = data[0];
  const emailIdx = headers.indexOf('adminEmail');
  const userIdIdx = headers.indexOf('userId');
  const urlIdx = headers.indexOf('spreadsheetUrl');

  for (let i = 1; i < data.length; i++) {
    if (data[i][emailIdx] === email) {
      const userId = data[i][userIdIdx];
      const base = getWebAppUrlEnhanced();
      return {
        userId: userId,
        adminUrl: base ? `${base}?userId=${userId}&mode=admin` : '',
        viewUrl: base ? `${base}?userId=${userId}` : '',
        spreadsheetUrl: data[i][urlIdx] || ''
      };
    }
  }
  return null;
}

// =================================================================
// CommonJS exports for unit tests
// =================================================================
if (typeof module !== 'undefined') {
  module.exports = {
    addReaction,
    buildBoardData,
    checkAdmin,
    COLUMN_HEADERS,
    doGet,
    findHeaderIndices,
    getAdminSettings,
    getHeaderIndices,
    getSheetData,
    getSheetDataForSpreadsheet,
    getSheetHeaders,
    getWebAppUrl,
    getWebAppUrlEnhanced,
    guessHeadersFromArray,
    parseReactionString,
    isUserAdmin,
    prepareSheetForBoard,
    saveDeployId,
    saveSheetConfig,
    saveWebAppUrl,
    toggleHighlight,
    registerNewUser,
    getSpreadsheetUrlForUser,
    openActiveSpreadsheet,
    getExistingBoard
  };
}
