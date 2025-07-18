/**
 * @fileoverview メインエントリーポイント - 2024年最新技術の結集
 * V8ランタイム、最新パフォーマンス技術、安定性強化を統合
 */

/**
 * HTML ファイルを読み込む include ヘルパー
 * @param {string} path ファイルパス
 * @return {string} HTML content
 */
function include(path) {
  try {
    const tmpl = HtmlService.createTemplateFromFile(path);
    tmpl.include = include;
    return tmpl.evaluate().getContent();
  } catch (error) {
    console.error(`Error including file ${path}:`, error);
    return `<!-- Error including ${path}: ${error.message} -->`;
  }
}

/**
 * JavaScript文字列エスケープ関数
 * @param {string} str エスケープする文字列
 * @return {string} エスケープされた文字列
 */
function escapeJavaScript(str) {
  if (!str) return '';
  return str.toString()
    .replace(/\\/g, '\\\\')
    .replace(/'/g, "\\'")
    .replace(/"/g, '\\"')
    .replace(/\n/g, '\\n')
    .replace(/\r/g, '\\r')
    .replace(/\t/g, '\\t');
}

// グローバル定数の定義
var SCRIPT_PROPS_KEYS = {
  SERVICE_ACCOUNT_CREDS: 'SERVICE_ACCOUNT_CREDS',
  DATABASE_SPREADSHEET_ID: 'DATABASE_SPREADSHEET_ID',
  ADMIN_EMAIL: 'ADMIN_EMAIL'
};

var DB_SHEET_CONFIG = {
  SHEET_NAME: 'Users',
  HEADERS: [
    'userId', 'adminEmail', 'spreadsheetId', 'spreadsheetUrl',
    'createdAt', 'configJson', 'lastAccessedAt', 'isActive'
  ]
};

var LOG_SHEET_CONFIG = {
  SHEET_NAME: 'Logs',
  HEADERS: ['timestamp', 'userId', 'action', 'details']
};

var COLUMN_HEADERS = {
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

// 表示モード定数
var DISPLAY_MODES = {
  ANONYMOUS: 'anonymous',
  NAMED: 'named'
};

// リアクション関連定数
var REACTION_KEYS = ['UNDERSTAND', 'LIKE', 'CURIOUS'];

// スコア計算設定
var SCORING_CONFIG = {
  LIKE_MULTIPLIER_FACTOR: 0.1,
  RANDOM_SCORE_FACTOR: 0.01
};

var EMAIL_REGEX = /^[^\n@]+@[^\n@]+\.[^\n@]+$/;
var DEBUG = PropertiesService.getScriptProperties()
  .getProperty('DEBUG_MODE') === 'true';

/**
 * Determine if a value represents boolean true.
 * Accepts boolean true, 'true', or 'TRUE'.
 * @param {any} value
 * @returns {boolean}
 */
function isTrue(value) {
  if (typeof value === 'boolean') return value === true;
  if (value === undefined || value === null) return false;
  return String(value).toLowerCase() === 'true';
}

/**
 * HTMLエスケープ関数（Utilities.htmlEncodeの代替）
 * @param {string} text - エスケープするテキスト
 * @returns {string} エスケープされたテキスト
 */
function htmlEncode(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * HtmlOutputに安全にX-Frame-Optionsヘッダーを設定するユーティリティ
 * @param {HtmlOutput} htmlOutput - 設定対象のHtmlOutput
 * @returns {HtmlOutput} 設定後のHtmlOutput
 */
function safeSetXFrameOptionsDeny(htmlOutput) {
  try {
    if (htmlOutput && typeof htmlOutput.setXFrameOptionsMode === 'function' &&
        HtmlService && HtmlService.XFrameOptionsMode &&
        HtmlService.XFrameOptionsMode.DENY) {
      htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
    }
  } catch (e) {
    console.warn('Failed to set XFrameOptionsMode:', e.message);
  }
  return htmlOutput;
}

// getSecurityHeaders function removed - not used in current implementation

// 安定性を重視してvarを使用
var ULTRA_CONFIG = {
  EXECUTION_LIMITS: {
    MAX_TIME: 330000, // 5.5分（安全マージン）
    BATCH_SIZE: 100,
    API_RATE_LIMIT: 90 // 100秒間隔での制限
  },
  
  CACHE_STRATEGY: {
    L1_TTL: 300,     // Level 1: 5分
    L2_TTL: 3600,    // Level 2: 1時間  
    L3_TTL: 21600    // Level 3: 6時間（最大）
  }
};

/**
 * 簡素化されたエラーハンドリング関数群
 */

// ログ出力の最適化
function log(level, message, details) {
  try {
    if (typeof globalProfiler !== 'undefined') {
      globalProfiler.start('logging');
    }

    if (level === 'info' && !DEBUG) {
      return;
    }

    switch (level) {
      case 'error':
        console.error(message, details || '');
        break;
      case 'warn':
        console.warn(message, details || '');
        break;
      default:
        console.log(message, details || '');
    }
    
    if (typeof globalProfiler !== 'undefined') {
      globalProfiler.end('logging');
    }
  } catch (e) {
    // ログ出力自体が失敗した場合は無視
  }
}

function debugLog() {
  if (!DEBUG) return;
  try {
    console.log.apply(console, arguments);
  } catch (e) {
    // ignore logging errors
  }
}

/**
 * デプロイされたWebアプリのドメイン情報と現在のユーザーのドメイン情報を取得
 * AdminPanel.html と Login.html から共通で呼び出される
 */
function getDeployUserDomainInfo() {
  try {
    var activeUserEmail = Session.getActiveUser().getEmail();
    var currentDomain = getEmailDomain(activeUserEmail);
    var webAppUrl = ScriptApp.getService().getUrl();
    var deployDomain = ''; // 個人アカウント/グローバルアクセスの場合、デフォルトで空

    if (webAppUrl) {
      // Google WorkspaceアカウントのドメインをURLから抽出
      var domainMatch =
        webAppUrl.match(/\/a\/([a-zA-Z0-9\-.]+)\/macros\//) ||
        webAppUrl.match(/\/a\/macros\/([a-zA-Z0-9\-.]+)\//);
      if (domainMatch && domainMatch[1]) {
        deployDomain = domainMatch[1];
      }
      // ドメインが抽出されなかった場合、deployDomainは空のままとなり、個人アカウント/グローバルアクセスを示す
    }

    // 現在のユーザーのドメインと抽出された/デフォルトのデプロイドメインを比較
    // deployDomainが空の場合、特定のドメインが強制されていないため、一致とみなす（グローバルアクセス）
    var isDomainMatch = (currentDomain === deployDomain) || (deployDomain === '');

    debugLog('Domain info:', {
      currentDomain: currentDomain,
      deployDomain: deployDomain,
      isDomainMatch: isDomainMatch,
      webAppUrl: webAppUrl
    });

    return {
      currentDomain: currentDomain,
      deployDomain: deployDomain,
      isDomainMatch: isDomainMatch,
      webAppUrl: webAppUrl
    };
  } catch (e) {
    console.error('getDeployUserDomainInfo エラー: ' + e.message);
    return {
      currentDomain: '不明',
      deployDomain: '不明',
      isDomainMatch: false,
      error: e.message
    };
  }
}


// PerformanceOptimizer.gsでglobalProfilerが既に定義されているため、
// フォールバックの宣言は不要

/**
 * ユーザーがデータベースに登録済みかを確認するヘルパー関数
 * @param {string} email 確認するユーザーのメールアドレス
 * @param {string} spreadsheetId データベースのスプレッドシートID
 * @return {boolean} 登録されていればtrue
 */

/**
 * システムの初期セットアップが完了しているかを確認するヘルパー関数
 * @returns {boolean} セットアップが完了していればtrue
 */
function isSystemSetup() {
  var props = PropertiesService.getScriptProperties();
  var dbSpreadsheetId = props.getProperty('DATABASE_SPREADSHEET_ID');
  var adminEmail = props.getProperty('ADMIN_EMAIL');
  var serviceAccountCreds = props.getProperty('SERVICE_ACCOUNT_CREDS');
  return !!dbSpreadsheetId && !!adminEmail && !!serviceAccountCreds;
}

/**
 * 登録ページを表示する関数
 */
function showRegistrationPage() {
  try {
    var template = HtmlService.createTemplateFromFile('LoginPage');
    template.include = include;
    
  
    // No template variable processing - client will get GOOGLE_CLIENT_ID via server function
    
    var output = template.evaluate()
      .setTitle('ログイン - StudyQuest');
    return safeSetXFrameOptionsDeny(output);
  } catch (error) {
    console.error('Error in showRegistrationPage:', error);
    return HtmlService.createHtmlOutput('<h1>システムエラー</h1><p>アプリケーションの読み込みに失敗しました。</p>');
  }
}

/**
 * Get Google Client ID for fallback authentication
 * @return {Object} Object containing client ID
 */
function getGoogleClientId() {
  try {
    console.log('Getting GOOGLE_CLIENT_ID from script properties...');
    var properties = PropertiesService.getScriptProperties();
    var clientId = properties.getProperty('GOOGLE_CLIENT_ID');
    
    console.log('GOOGLE_CLIENT_ID retrieved:', clientId ? 'Found' : 'Not found');
    
    if (!clientId) {
      console.warn('GOOGLE_CLIENT_ID not found in script properties');
      
      // Try to get all properties to see what's available
      var allProperties = properties.getProperties();
      console.log('Available properties:', Object.keys(allProperties));
      
      return { 
        clientId: '', 
        error: 'GOOGLE_CLIENT_ID not found in script properties',
        setupInstructions: 'Please set GOOGLE_CLIENT_ID in Google Apps Script project settings under Properties > Script Properties'
      };
    }
    
    return { clientId: clientId, success: true };
  } catch (error) {
    console.error('Error getting GOOGLE_CLIENT_ID:', error);
    return { clientId: '', error: error.toString() };
  }
}

/**
 * システム設定の詳細チェック
 * @return {Object} システム設定の詳細情報
 */
function checkSystemConfiguration() {
  try {
    var properties = PropertiesService.getScriptProperties();
    var allProperties = properties.getProperties();
    
    var requiredProperties = [
      'GOOGLE_CLIENT_ID',
      'DATABASE_SPREADSHEET_ID', 
      'ADMIN_EMAIL',
      'SERVICE_ACCOUNT_CREDS'
    ];
    
    var configStatus = {};
    var missingProperties = [];
    
    requiredProperties.forEach(function(prop) {
      var value = allProperties[prop];
      configStatus[prop] = {
        exists: !!value,
        hasValue: !!(value && value.trim()),
        length: value ? value.length : 0
      };
      
      if (!value || !value.trim()) {
        missingProperties.push(prop);
      }
    });
    
    return {
      isFullyConfigured: missingProperties.length === 0,
      configStatus: configStatus,
      missingProperties: missingProperties,
      availableProperties: Object.keys(allProperties),
      setupComplete: isSystemSetup()
    };
  } catch (error) {
    console.error('Error checking system configuration:', error);
    return {
      isFullyConfigured: false,
      error: error.toString()
    };
  }
}

/**
 * Retrieves the administrator domain for the login page.
 * @returns {{adminDomain: string}|{error: string}} Domain info or error message.
 */
function getSystemDomainInfo() {
  try {
    var props = PropertiesService.getScriptProperties();
    var adminEmail = props.getProperty(SCRIPT_PROPS_KEYS.ADMIN_EMAIL);
    if (!adminEmail) {
      throw new Error('システム管理者が設定されていません。');
    }

    var adminDomain = adminEmail.split('@')[1];
    return { adminDomain: adminDomain };
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * Webアプリケーションのメインエントリーポイント
 * @param {Object} e - URLパラメータを含むイベントオブジェクト
 * @returns {HtmlOutput} 表示するHTMLコンテンツ
 */
function doGet(e) {
  try {
    // パラメータとユーザー情報を取得
    const params = parseRequestParams(e);
    const userEmail = Session.getActiveUser().getEmail();

    // リクエストをルーティング
    return routeRequest(params, userEmail);
  } catch (error) {
    console.error(`doGetで致命的なエラー: ${error.stack}`);
    return showErrorPage('致命的なエラー', 'アプリケーションの処理中に予期せぬエラーが発生しました。', error);
  }
}

/**
 * リクエストを適切なハンドラに振り分ける
 * @param {Object} params - 解析されたリクエストパラメータ
 * @param {string} userEmail - 現在のユーザーのメールアドレス
 * @returns {HtmlOutput}
 */
function routeRequest(params, userEmail) {
  // 1. システムセットアップが最優先
  if (!isSystemSetup()) {
    return showSetupPage();
  }

  // 2. パラメータなしのルートアクセスは常にログインページを表示
  if (!params.mode) {
    return showLoginPage();
  }

  // 3. ログインしていないユーザーはログインページへ（パラメータがあっても）
  if (!userEmail) {
    return showLoginPage();
  }

  // 4. ユーザー情報を取得（キャッシュ活用）
  const userInfo = getUserInfo(userEmail, params.userId);

  // 5. ユーザー情報がない場合は、どのモードであってもログインページを表示
  if (!userInfo) {
    console.warn('No user info found for email:', userEmail, 'or userId:', params.userId, '. Showing login page.');
    return showLoginPage();
  }

  // 6. ルーティング決定
  switch (params.mode) {
    case 'admin':
      return handleAdminRoute(userInfo, params, userEmail);
    case 'view':
      return renderAnswerBoard(userInfo, params);
    case 'appsetup':
      if (params.setupParam === 'true' && hasSetupPageAccess()) {
        return showAppSetupPage();
      }
      return showErrorPage('アクセス拒否', 'このページにアクセスする権限がありません。');
    default:
      // 不明なモードの場合はログインページへ
      return showLoginPage();
  }
}

/**
 * 管理者ページのルートを処理
 * @param {Object} userInfo - ユーザー情報
 * @param {Object} params - リクエストパラメータ
 * @param {string} userEmail - 現在のユーザーのメールアドレス
 * @returns {HtmlOutput}
 */
function handleAdminRoute(userInfo, params, userEmail) {
  // この関数が呼ばれる時点でuserInfoはnullではないことが保証されている

  // セキュリティチェック: アクセスしようとしているuserIdが自分のものでなければ、自分の管理画面にリダイレクト
  if (params.userId && params.userId !== userInfo.userId) {
    console.warn(`不正アクセス試行: ${userEmail} が userId ${params.userId} にアクセスしようとしました。`);
    const correctUrl = buildUserAdminUrl(userInfo.userId);
    return createSecureRedirect(correctUrl, '要求された管理パネルへのアクセス権がありません。');
  }

  return renderAdminPanel(userInfo, params.mode);
}


/**
 * ユーザー情報を取得（キャッシュ対応）
 * @param {string} email - ユーザーのメールアドレス
 * @param {string} userId - URLパラメータから取得したユーザーID
 * @returns {Object|null} ユーザー情報オブジェクト、見つからない場合はnull
 */
function getUserInfo(email, userId) {
  const cache = CacheService.getUserCache();
  const cacheKey = `user_info_${userId || email}`;
  
  const cachedInfo = cache.get(cacheKey);
  if (cachedInfo) {
    return JSON.parse(cachedInfo);
  }

  let userInfo = null;
  if (userId) {
    userInfo = findUserById(userId);
    // セキュリティチェック: 取得した情報のemailが現在のユーザーと一致するか確認
    if (userInfo && userInfo.adminEmail !== email) {
      return null; // 他人の情報へのアクセスは許可しない
    }
  } else {
    userInfo = findUserByEmail(email);
  }

  if (userInfo) {
    cache.put(cacheKey, JSON.stringify(userInfo), 300); // 5分キャッシュ
  }

  return userInfo;
}

/**
 * ログインページを表示
 * @returns {HtmlOutput}
 */
function showLoginPage() {
  const template = HtmlService.createTemplateFromFile('LoginPage');
  return template.evaluate()
    .setTitle('StudyQuest - ログイン')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 初回セットアップページを表示
 * @returns {HtmlOutput}
 */
function showSetupPage() {
  const template = HtmlService.createTemplateFromFile('SetupPage');
  return template.evaluate()
    .setTitle('StudyQuest - 初回セットアップ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
}

/**
 * アプリ設定ページを表示
 * @returns {HtmlOutput}
 */
function showAppSetupPage() {
    const appSetupTemplate = HtmlService.createTemplateFromFile('AppSetupPage');
    return appSetupTemplate.evaluate()
      .setTitle('アプリ設定 - StudyQuest')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
}


/**
 * エラーページを表示
 * @param {string} title - エラータイトル
 * @param {string} message - エラーメッセージ
 * @param {Error} [error] - (オプション) エラーオブジェクト
 * @returns {HtmlOutput}
 */
function showErrorPage(title, message, error) {
  const template = HtmlService.createTemplateFromFile('ErrorBoundary');
  template.title = title;
  template.message = message;
  template.mode = 'admin'; // エラーテンプレートが依存するmode変数にデフォルト値を提供
  if (DEBUG && error) {
    template.debugInfo = error.stack;
  } else {
    template.debugInfo = '';
  }
  return template.evaluate()
    .setTitle(`エラー - ${title}`)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
}

/**
 * ユーザー専用の一意の管理パネルURLを構築
 * @param {string} userId ユーザーID
 * @return {string}
 */
function buildUserAdminUrl(userId) {
  const baseUrl = getWebAppUrl();
  return `${baseUrl}?mode=admin&userId=${encodeURIComponent(userId)}`;
}


/**
 * セキュアなリダイレクトHTMLを作成 (テンプレート使用)
 * @param {string} targetUrl リダイレクト先URL
 * @param {string} message 表示メッセージ
 * @return {HtmlOutput}
 */
function createSecureRedirect(targetUrl, message) {
  const template = HtmlService.createTemplateFromFile('Redirect');
  template.targetUrl = targetUrl;
  template.message = message;
  return template.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 正しいWeb App URLを取得 (キャッシュ対応)
 * @return {string}
 */
function getWebAppUrl() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'web_app_url';
  
  const cachedUrl = cache.get(cacheKey);
  if (cachedUrl) {
    return cachedUrl;
  }

  try {
    const url = ScriptApp.getService().getUrl();
    if (url && url.includes('/macros/')) {
      cache.put(cacheKey, url, 21600); // 6時間キャッシュ
      return url;
    }
    throw new Error('Invalid URL from getService()');
  } catch (error) {
    console.error('getWebAppUrl error:', error);
    // 緊急時のフォールバックURL
    const fallbackUrl = 'https://script.google.com/a/macros/naha-okinawa.ed.jp/s/AKfycby5oABLEuyg46OvwVqt2flUKz15zocFhH-kLD0IuNWm8akKMXiKrOS5kqGCQ7V4DQ-2/exec';
    cache.put(cacheKey, fallbackUrl, 21600);
    return fallbackUrl;
  }
}


/**
 * doGet のリクエストパラメータを解析
 * @param {Object} e Event object
 * @return {{mode:string,userId:string|null,setupParam:string|null,spreadsheetId:string|null,sheetName:string|null,isDirectPageAccess:boolean}}
 */
function parseRequestParams(e) {
  const p = (e && e.parameter) || {};
  const mode = p.mode || null; // デフォルトを'admin'からnullに変更し、パラメータの有無を明確化
  const userId = p.userId || null;
  const setupParam = p.setup || null;
  const spreadsheetId = p.spreadsheetId || null;
  const sheetName = p.sheetName || null;
  const isDirectPageAccess = !!(userId && mode === 'view');
  return { mode, userId, setupParam, spreadsheetId, sheetName, isDirectPageAccess };
}



/**
 * マルチテナント対応: ステートレスなユーザーセッション検証
 * 全ての処理をuserIdベースで実行し、emailは認証確認のみに使用
 * @param {string} currentUserEmail - 現在の認証済みユーザーメール（認証確認のみ）
 * @param {Object} params - リクエストパラメータ（userIdを含む）
 * @returns {Object} userEmailとuserInfoを含むオブジェクト
 */
function validateUserSession(currentUserEmail, params) {
  console.log('validateUserSession - email:', currentUserEmail, 'params:', params);
  
  // 基本的な認証チェック（Googleアカウントにログインしているか）
  if (!currentUserEmail) {
    console.warn('validateUserSession - no authenticated user');
    return { userEmail: null, userInfo: null };
  }
  
  let userInfo = null;
  
  // マルチテナント対応: 常にuserIdベースでユーザー情報を取得
  if (params.userId) {
    console.log('validateUserSession - looking up userId:', params.userId);
    userInfo = findUserById(params.userId);
    
    if (userInfo) {
      // セキュリティチェック: リクエストされたuserIdが認証済みユーザーのものか確認
      if (userInfo.adminEmail !== currentUserEmail) {
        console.warn('validateUserSession - security violation:', 
                    'authenticated:', currentUserEmail, 
                    'requested:', userInfo.adminEmail);
        // 本人以外のデータへのアクセス試行は拒否
        return { userEmail: currentUserEmail, userInfo: null };
      }
      
      console.log('validateUserSession - valid user found:', userInfo.userId);
    } else {
      console.warn('validateUserSession - userId not found:', params.userId);
    }
  } else if (params.isDirectPageAccess) {
    // 直接ページアクセスの場合は、emailベースでフォールバック
    console.log('validateUserSession - direct page access, looking up by email');
    userInfo = findUserByEmail(currentUserEmail);
  } else {
    // userIdパラメータがない場合は、emailベースでフォールバック
    console.log('validateUserSession - no userId param, looking up by email');
    userInfo = findUserByEmail(currentUserEmail);
  }
  
  return { userEmail: currentUserEmail, userInfo: userInfo };
}

/**
 * セットアップページ関連の処理
 * @param {Object} params リクエストパラメータ
 * @param {string} userEmail 現在のユーザーのメール
 * @return {HtmlOutput|null} 表示するHTMLがあれば返す
 */
function handleSetupPages(params, userEmail) {
  // セットアップが未完了の場合は、userEmailの有無に関係なくセットアップページを優先
  if (!isSystemSetup() && !params.isDirectPageAccess) {
    const t = HtmlService.createTemplateFromFile('SetupPage');
    t.include = include;
    return safeSetXFrameOptionsDeny(t.evaluate().setTitle('初回セットアップ - StudyQuest'));
  }

  if (params.setupParam === 'true' && params.mode === 'appsetup') {
    const domainInfo = getDeployUserDomainInfo();
    if (domainInfo.deployDomain && domainInfo.deployDomain !== '' && !domainInfo.isDomainMatch) {
      debugLog('Domain access warning. Current:', domainInfo.currentDomain, 'Deploy:', domainInfo.deployDomain);
    }
    if (!hasSetupPageAccess()) {
      const errorHtml = HtmlService.createHtmlOutput(
        '<h1>アクセス拒否</h1>' +
        '<p>アプリ設定ページにアクセスする権限がありません。</p>' +
        '<p>編集者として登録され、かつアクティブ状態である必要があります。</p>' +
        '<p>現在のドメイン: ' + (domainInfo.currentDomain || '不明') + '</p>'
      );
      return safeSetXFrameOptionsDeny(errorHtml);
    }
    const appSetupTemplate = HtmlService.createTemplateFromFile('AppSetupPage');
    appSetupTemplate.include = include;
    return safeSetXFrameOptionsDeny(appSetupTemplate.evaluate().setTitle('アプリ設定 - StudyQuest'));
  }

  if (params.setupParam === 'true') {
    const explicit = HtmlService.createTemplateFromFile('SetupPage');
    explicit.include = include;
    return safeSetXFrameOptionsDeny(explicit.evaluate().setTitle('StudyQuest - サービスアカウント セットアップ'));
  }

  // システムセットアップが完了している場合のみ、userEmailをチェック
  if (!userEmail && !params.isDirectPageAccess) {
    return showRegistrationPage();
  }

  return null;
}

/**
 * 管理パネルを表示
 * @param {Object} userInfo ユーザー情報
 * @param {string} mode 表示モード
 * @return {HtmlOutput} HTMLコンテンツ
 */
function renderAdminPanel(userInfo, mode) {
  // ガード節: userInfoが存在しない場合はエラーページを表示して処理を中断
  if (!userInfo) {
    console.error('renderAdminPanelにuserInfoがnullで渡されました。これは予期せぬ状態です。');
    return showErrorPage('エラー', 'ユーザー情報の読み込みに失敗したため、管理パネルを表示できません。');
  }

  const adminTemplate = HtmlService.createTemplateFromFile('AdminPanel');
  adminTemplate.include = include;
  adminTemplate.userInfo = userInfo;
  adminTemplate.userId = userInfo.userId;
  adminTemplate.mode = mode || 'admin'; // 安全のためのフォールバック
  adminTemplate.displayMode = 'named';
  adminTemplate.showAdminFeatures = true;
  const deployUserResult = isDeployUser();
  console.log('renderAdminPanel - isDeployUser() result:', deployUserResult);
  console.log('renderAdminPanel - current user email:', Session.getActiveUser().getEmail());
  adminTemplate.isDeployUser = deployUserResult;
  adminTemplate.DEBUG_MODE = shouldEnableDebugMode();
  
  // URL更新用の情報を追加
  const correctUrl = buildUserAdminUrl(userInfo.userId);
  adminTemplate.correctUrl = correctUrl;
  adminTemplate.shouldUpdateUrl = true;
  
  return adminTemplate.evaluate()
    .setTitle('みんなの回答ボード 管理パネル')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.NATIVE);
}

/**
 * 回答ボードまたは未公開ページを表示
 * @param {Object} userInfo ユーザー情報
 * @param {Object} params リクエストパラメータ
 * @return {HtmlOutput} HTMLコンテンツ
 */
function renderAnswerBoard(userInfo, params) {
  let config = {};
  try {
    config = JSON.parse(userInfo.configJson || '{}');
  } catch (e) {
    console.warn('Invalid configJson:', e.message);
  }
  const isPublished = !!(config.appPublished && config.publishedSpreadsheetId && config.publishedSheetName);
  const sheetConfigKey = 'sheet_' + (config.publishedSheetName || params.sheetName);
  const sheetConfig = config[sheetConfigKey] || {};
  const showBoard = params.isDirectPageAccess || isPublished;
  const file = showBoard ? 'Page' : 'Unpublished';
  const template = HtmlService.createTemplateFromFile(file);
  template.include = include;

  if (showBoard) {
    try {
      if (userInfo.spreadsheetId) {
        try { addServiceAccountToSpreadsheet(userInfo.spreadsheetId); } catch (err) { console.warn('アクセス権設定警告:', err.message); }
      }
      template.userId = userInfo.userId;
      template.spreadsheetId = userInfo.spreadsheetId;
      template.ownerName = userInfo.adminEmail;
      template.sheetName = escapeJavaScript(config.publishedSheetName || params.sheetName);
      template.DEBUG_MODE = shouldEnableDebugMode();
      const rawOpinionHeader = sheetConfig.opinionHeader || config.publishedSheetName || 'お題';
      template.opinionHeader = escapeJavaScript(rawOpinionHeader);
      template.cacheTimestamp = Date.now();
      template.displayMode = config.displayMode || 'anonymous';
      template.showCounts = config.showCounts !== undefined ? config.showCounts : false;
      template.showScoreSort = template.showCounts;
      const currentUserEmail = Session.getActiveUser().getEmail();
      const isOwner = currentUserEmail === userInfo.adminEmail;
      template.showAdminFeatures = isOwner;
      template.isAdminUser = isOwner;
    } catch (e) {
      template.opinionHeader = escapeJavaScript('お題の読込エラー');
      template.cacheTimestamp = Date.now();
      template.userId = userInfo.userId;
      template.spreadsheetId = userInfo.spreadsheetId;
      template.ownerName = userInfo.adminEmail;
      template.sheetName = escapeJavaScript(params.sheetName);
      template.displayMode = 'anonymous';
      template.showCounts = false;
      template.showScoreSort = false;
      const currentUserEmail = Session.getActiveUser().getEmail();
      const isOwner = currentUserEmail === userInfo.adminEmail;
      template.showAdminFeatures = isOwner;
      template.isAdminUser = isOwner;
    }
    return template.evaluate()
      .setTitle('StudyQuest -みんなの回答ボード-')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else {
    try {
      template.userId = userInfo.userId;
      template.spreadsheetId = userInfo.spreadsheetId;
      template.ownerName = userInfo.adminEmail;
      template.isOwner = true;
      template.adminEmail = userInfo.adminEmail;
      template.cacheTimestamp = Date.now();
      const appUrls = generateAppUrls(userInfo.userId);
      template.adminPanelUrl = appUrls.adminUrl;
      template.boardUrl = appUrls.viewUrl;
    } catch (e) {
      console.error('Unpublished template setup error:', e);
      template.ownerName = 'システム管理者';
      template.isOwner = true;
      template.adminEmail = userInfo.adminEmail || 'admin@example.com';
      template.cacheTimestamp = Date.now();
    }
    return template.evaluate()
      .setTitle('StudyQuest - 準備中')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}


/**
 * JavaScript エスケープのテスト関数
 */

/**
 * パフォーマンス監視エンドポイント（簡易版）
 */

/**
 * escapeJavaScript関数のテスト
 */

/**
 * Base64エンコード/デコードのテスト
 */
