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
 * JavaScript文字列エスケープ関数 (URL対応版)
 * @param {string} str エスケープする文字列
 * @return {string} エスケープされた文字列
 */
function escapeJavaScript(str) {
  if (!str) return '';
  
  const strValue = str.toString();
  
  // URL判定: HTTP/HTTPSで始まり、すでに適切にエスケープされている場合は最小限の処理
  if (strValue.match(/^https?:\/\/[^\s<>"']+$/)) {
    // URLの場合はバックスラッシュと改行文字のみエスケープ
    return strValue
      .replace(/\\/g, '\\\\')
      .replace(/\n/g, '\\n')
      .replace(/\r/g, '\\r')
      .replace(/\t/g, '\\t');
  }
  
  // 通常のテキストの場合は従来通りの完全エスケープ
  return strValue
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
    
    // 統一されたURL取得システムを使用（開発URL除去機能付き）
    var webAppUrl = getWebAppUrlCached();
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
    
    return { status: 'success', message: 'Google Client IDを取得しました', data: { clientId: clientId } };
  } catch (error) {
    console.error('Error getting GOOGLE_CLIENT_ID:', error);
    return { status: 'error', message: 'Google Client IDの取得に失敗しました: ' + error.toString(), data: { clientId: '' } };
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
 * Retrieves the administrator domain for the login page with domain match status.
 * @returns {{adminDomain: string, isDomainMatch: boolean}|{error: string}} Domain info or error message.
 */
function getSystemDomainInfo() {
  try {
    var props = PropertiesService.getScriptProperties();
    var adminEmail = props.getProperty(SCRIPT_PROPS_KEYS.ADMIN_EMAIL);
    if (!adminEmail) {
      throw new Error('システム管理者が設定されていません。');
    }

    var adminDomain = adminEmail.split('@')[1];
    
    // 現在のユーザーのドメイン一致状況を取得
    var domainInfo = getDeployUserDomainInfo();
    var isDomainMatch = domainInfo.isDomainMatch !== undefined ? domainInfo.isDomainMatch : false;
    
    return { 
      adminDomain: adminDomain,
      isDomainMatch: isDomainMatch,
      currentDomain: domainInfo.currentDomain || '不明',
      deployDomain: domainInfo.deployDomain || adminDomain
    };
  } catch (e) {
    console.error('getSystemDomainInfo エラー:', e.message);
    return { error: e.message };
  }
}

/**
 * Webアプリケーションのメインエントリーポイント
 * リダイレクト処理強化版：認証成功時に直接管理パネルへリダイレクト
 * @param {Object} e - URLパラメータを含むイベントオブジェクト
 * @returns {HtmlOutput} 表示するHTMLコンテンツ
 */
function doGet(e) {
  try {
    // 0. 実行レベルキャッシュクリア（新しいリクエスト開始）
    clearAllExecutionCache();
    
    // 1. システムセットアップの確認 (最優先)
    if (!isSystemSetup()) {
      return showSetupPage();
    }

    // 2. URLパラメータの解析（統一されたURL解析関数を使用）
    const params = parseUrlParameters(e);

    // 3. ログイン状態の確認
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      return showLoginPage();
    }

    // 4. アプリ設定ページリクエストの処理
    if (params.setup === 'true') {
      return showAppSetupPage(params.userId);
    }

    // 5. 認証成功後の自動リダイレクト処理
    if (!params.mode && userEmail) {
      // モードが指定されていない場合、ログインユーザーの管理パネルに自動リダイレクト
      console.log('認証成功、管理パネルに自動リダイレクト開始:', userEmail);
      
      try {
        // 既存ユーザーかチェック
        const existingUser = findUserByEmail(userEmail);
        if (existingUser && existingUser.userId) {
          const adminUrl = buildAdminPanelUrl(existingUser.userId);
          console.log('既存ユーザー、管理パネルへリダイレクト:', adminUrl);
          return createRedirectResponse(adminUrl);
        } else {
          // 新規ユーザーの場合はログインページを表示（初回登録のため）
          console.log('新規ユーザー、ログインページを表示');
          return showLoginPage();
        }
      } catch (redirectError) {
        console.error('自動リダイレクト処理エラー:', redirectError.message);
        return showLoginPage(); // エラー時はログインページにフォールバック
      }
    }

    // 6. mode=admin の場合
    if (params.mode === 'admin') {
      if (!params.userId) {
        return showErrorPage('不正なリクエスト', 'ユーザーIDが指定されていません。');
      }
      // 本人確認
      if (!verifyAdminAccess(params.userId)) {
        return showErrorPage('アクセス拒否', 'この管理パネルにアクセスする権限がありません。');
      }
      const userInfo = findUserById(params.userId);
      if (!userInfo) {
        return showErrorPage('エラー', 'ユーザー情報が見つかりません。');
      }
      return renderAdminPanel(userInfo, 'admin');
    }

    // 7. mode=view の場合
    if (params.mode === 'view') {
      if (!params.userId) {
        return showErrorPage('不正なリクエスト', 'ユーザーIDが指定されていません。');
      }
      const userInfo = findUserById(params.userId);
      if (!userInfo) {
        return showErrorPage('エラー', '指定されたユーザーが見つかりません。');
      }
      return renderAnswerBoard(userInfo, params);
    }
    
    // 8. 不明なモードの場合はログインページへ
    console.log('不明なモードまたはパラメータなし、ログインページを表示');
    return showLoginPage();

  } catch (error) {
    console.error(`doGetで致命的なエラー: ${error.stack}`);
    return showErrorPage('致命的なエラー', 'アプリケーションの処理中に予期せぬエラーが発生しました。', error);
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

  // 強化されたセキュリティ検証: 指定されたIDの登録メールアドレスと現在ログイン中のGoogleアカウントが一致するかを検証
  if (params.userId) {
    const isVerified = verifyAdminAccess(params.userId);
    if (!isVerified) {
      console.warn(`セキュリティ検証失敗: userId ${params.userId} への不正アクセス試行をブロックしました。`);
      const correctUrl = buildUserAdminUrl(userInfo.userId);
      return createSecureRedirect(correctUrl, 'セキュリティ検証に失敗しました。正しい管理パネルにリダイレクトします。');
    }
    console.log(`✅ セキュリティ検証成功: userId ${params.userId} への正当なアクセスを確認しました。`);
  }

  return renderAdminPanel(userInfo, params.mode);
}


/**
 * 統合されたユーザー情報取得関数 (Phase 2: 統合完了)
 * キャッシュ戦略とセキュリティチェックを統合
 * @param {string|Object} identifier - メールアドレス、ユーザーID、または設定オブジェクト
 * @param {string} [type] - 'email' | 'userId' | null (auto-detect)
 * @param {Object} [options] - キャッシュオプション
 * @returns {Object|null} ユーザー情報オブジェクト、見つからない場合はnull
 */
function getOrFetchUserInfo(identifier, type = null, options = {}) {
  // オプションのデフォルト値
  const opts = {
    ttl: options.ttl || 300, // 5分キャッシュ
    enableSecurityCheck: options.enableSecurityCheck !== false, // デフォルトで有効
    currentUserEmail: options.currentUserEmail || null,
    useExecutionCache: options.useExecutionCache || false,
    ...options
  };

  // 引数の正規化
  let email = null;
  let userId = null;
  
  if (typeof identifier === 'object' && identifier !== null) {
    // オブジェクト形式の場合（後方互換性）
    email = identifier.email;
    userId = identifier.userId;
  } else if (typeof identifier === 'string') {
    // 文字列の場合、typeに基づいて判定
    if (type === 'email' || (!type && identifier.includes('@'))) {
      email = identifier;
    } else {
      userId = identifier;
    }
  }

  // キャッシュキーの生成
  const cacheKey = `unified_user_info_${userId || email}`;
  
  // 実行レベルキャッシュの確認（オプション）
  if (opts.useExecutionCache && _executionUserInfoCache && 
      _executionUserInfoCache.userId === userId) {
    return _executionUserInfoCache.userInfo;
  }

  // 統合キャッシュマネージャーを使用（キャッシュ miss 時は自動でデータベースから取得）
  let userInfo = null;
  
  try {
    userInfo = cacheManager.get(cacheKey, () => {
      console.log('🔍 キャッシュmiss - データベースから取得:', userId || email);
      
      let dbUserInfo = null;
      if (userId) {
        dbUserInfo = findUserById(userId);
        // セキュリティチェック: 取得した情報のemailが現在のユーザーと一致するか確認
        if (opts.enableSecurityCheck && dbUserInfo && opts.currentUserEmail && 
            dbUserInfo.adminEmail !== opts.currentUserEmail) {
          console.warn('セキュリティチェック失敗: 他人の情報へのアクセス試行');
          return null;
        }
      } else if (email) {
        dbUserInfo = findUserByEmail(email);
      }
      
      return dbUserInfo;
    }, { 
      ttl: opts.ttl || 300,
      enableMemoization: opts.enableMemoization || false 
    });
    
    // 実行レベルキャッシュにも保存（オプション）
    if (userInfo && opts.useExecutionCache && (userId || userInfo.userId)) {
      _executionUserInfoCache = { userId: userId || userInfo.userId, userInfo };
      console.log('✅ 実行レベルキャッシュに保存:', userId || userInfo.userId);
    }
    
  } catch (cacheError) {
    console.error('統合キャッシュアクセスでエラー:', cacheError.message);
    // フォールバック: 直接データベースから取得
    if (userId) {
      userInfo = findUserById(userId);
    } else if (email) {
      userInfo = findUserByEmail(email);
    }
  }

  return userInfo;
}

/**
 * レガシー getUserInfo 関数（後方互換性のため保持）
 * @deprecated getOrFetchUserInfo を使用してください
 */
function getUserInfo(email, userId) {
  return getOrFetchUserInfo(userId || email, userId ? 'userId' : 'email', {
    currentUserEmail: email,
    enableSecurityCheck: true
  });
}

/**
 * ログインページを表示
 * @returns {HtmlOutput}
 */
function showLoginPage() {
  const template = HtmlService.createTemplateFromFile('LoginPage');
  const htmlOutput = template.evaluate()
    .setTitle('StudyQuest - ログイン');
  
  // XFrameOptionsMode を安全に設定
  try {
    if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.ALLOWALL) {
      htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  } catch (e) {
    console.warn('XFrameOptionsMode設定エラー:', e.message);
  }
  
  return htmlOutput;
}

/**
 * 初回セットアップページを表示
 * @returns {HtmlOutput}
 */
function showSetupPage() {
  const template = HtmlService.createTemplateFromFile('SetupPage');
  const htmlOutput = template.evaluate()
    .setTitle('StudyQuest - 初回セットアップ');
  
  // XFrameOptionsMode を安全に設定
  try {
    if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.DENY) {
      htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
    }
  } catch (e) {
    console.warn('XFrameOptionsMode設定エラー:', e.message);
  }
  
  return htmlOutput;
}

/**
 * アプリ設定ページを表示
 * @returns {HtmlOutput}
 */
function showAppSetupPage(userId) {
    // システム管理者権限チェック
    try {
      if (!isDeployUser()) {
        console.warn('Unauthorized access attempt to app setup page:', Session.getActiveUser().getEmail());
        return showErrorPage('アクセス権限がありません', 'この機能にアクセスする権限がありません。システム管理者にお問い合わせください。');
      }
    } catch (error) {
      console.error('Error checking deploy user permissions:', error);
      return showErrorPage('認証エラー', '権限確認中にエラーが発生しました。');
    }
    
    const appSetupTemplate = HtmlService.createTemplateFromFile('AppSetupPage');
    appSetupTemplate.userId = userId;
    const htmlOutput = appSetupTemplate.evaluate()
      .setTitle('アプリ設定 - StudyQuest');
    
    // XFrameOptionsMode を安全に設定
    try {
      if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.DENY) {
        htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
      }
    } catch (e) {
      console.warn('XFrameOptionsMode設定エラー:', e.message);
    }
    
    return htmlOutput;
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
  const htmlOutput = template.evaluate()
    .setTitle(`エラー - ${title}`);
  
  // XFrameOptionsMode を安全に設定
  try {
    if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.DENY) {
      htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
    }
  } catch (e) {
    console.warn('XFrameOptionsMode設定エラー:', e.message);
  }
  
  return htmlOutput;
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
 * 指定されたURLへリダイレクトするサーバーサイド関数
 * @param {string} url - リダイレクト先のURL
 * @returns {HtmlOutput} リダイレクトを実行するHTML出力
 */
function redirectToUrl(url) {
  // XSS攻撃を防ぐため、URLをサニタイズ
  const sanitizedUrl = sanitizeRedirectUrl(url);
  return HtmlService.createHtmlOutput().setContent(`<script>window.top.location.href = '${sanitizedUrl}';</script>`);
}


/**
 * セキュアなリダイレクトHTMLを作成 (シンプル版)
 * @param {string} targetUrl リダイレクト先URL
 * @param {string} message 表示メッセージ
 * @return {HtmlOutput}
 */
function createSecureRedirect(targetUrl, message) {
  // URL検証とサニタイゼーション
  const sanitizedUrl = sanitizeRedirectUrl(targetUrl);
  
  console.log('createSecureRedirect - Original URL:', targetUrl);
  console.log('createSecureRedirect - Sanitized URL:', sanitizedUrl);
  
  // ユーザーアクティベーション必須のHTMLアンカー方式（サンドボックス制限準拠）
  const userActionRedirectHtml = `
    <!DOCTYPE html>
    <html>
    <head>
      <title>${message || 'アクセス確認'}</title>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <style>
        body { 
          font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
          background: linear-gradient(135deg, #1e3a8a 0%, #7c3aed 100%);
          min-height: 100vh;
          margin: 0;
          display: flex;
          align-items: center;
          justify-content: center;
          padding: 20px;
        }
        .container {
          background: rgba(31, 41, 55, 0.95);
          border-radius: 16px;
          padding: 40px;
          max-width: 500px;
          width: 100%;
          text-align: center;
          box-shadow: 0 25px 50px rgba(0, 0, 0, 0.5);
          border: 1px solid rgba(75, 85, 99, 0.3);
        }
        .icon { font-size: 64px; margin-bottom: 20px; }
        .title { 
          color: #10b981; 
          font-size: 24px; 
          font-weight: bold; 
          margin-bottom: 16px; 
        }
        .subtitle { 
          color: #d1d5db; 
          margin-bottom: 32px; 
          line-height: 1.5;
        }
        .main-button {
          display: inline-block;
          background: linear-gradient(135deg, #10b981 0%, #3b82f6 100%);
          color: white;
          font-weight: bold;
          padding: 16px 32px;
          border-radius: 12px;
          text-decoration: none;
          font-size: 18px;
          transition: all 0.3s ease;
          margin-bottom: 24px;
          box-shadow: 0 4px 15px rgba(16, 185, 129, 0.4);
        }
        .main-button:hover {
          transform: translateY(-2px);
          box-shadow: 0 8px 25px rgba(16, 185, 129, 0.6);
        }
        .url-info {
          background: rgba(17, 24, 39, 0.8);
          border-radius: 8px;
          padding: 16px;
          margin: 20px 0;
          border: 1px solid rgba(75, 85, 99, 0.5);
        }
        .url-text {
          color: #60a5fa;
          font-family: 'Courier New', monospace;
          font-size: 12px;
          word-break: break-all;
          line-height: 1.4;
        }
        .note {
          color: #9ca3af;
          font-size: 14px;
          margin-top: 20px;
          line-height: 1.4;
        }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="icon">🔐</div>
        <h1 class="title">${message || 'アクセス確認'}</h1>
        <p class="subtitle">セキュリティのため、下のボタンをクリックして続行してください</p>
        
        <a href="${sanitizedUrl}" target="_top" class="main-button">
          🚀 続行する
        </a>
        
        <div class="url-info">
          <div class="url-text">${sanitizedUrl}</div>
        </div>
        
        <div class="note">
          ✓ このリンクは安全です<br>
          ✓ Google Apps Script公式のセキュリティガイドラインに準拠
        </div>
      </div>
    </body>
    </html>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(userActionRedirectHtml);
  
  // XFrameOptionsMode を安全に設定
  try {
    if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.ALLOWALL) {
      htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  } catch (e) {
    console.warn('XFrameOptionsMode設定エラー:', e.message);
  }
  
  return htmlOutput;
}

/**
 * リダイレクト用URLの検証とサニタイゼーション
 * @param {string} url 検証対象のURL
 * @return {string} サニタイズされたURL
 */
function sanitizeRedirectUrl(url) {
  if (!url) {
    return getWebAppUrlCached();
  }
  
  try {
    let cleanUrl = String(url).trim();
    
    // 複数レベルのクォート除去（JSON文字列化による多重クォートに対応）
    let previousUrl = '';
    while (cleanUrl !== previousUrl) {
      previousUrl = cleanUrl;
      
      // 先頭と末尾のクォートを除去
      if ((cleanUrl.startsWith('"') && cleanUrl.endsWith('"')) ||
          (cleanUrl.startsWith("'") && cleanUrl.endsWith("'"))) {
        cleanUrl = cleanUrl.slice(1, -1);
      }
      
      // エスケープされたクォートを除去
      cleanUrl = cleanUrl.replace(/\\"/g, '"').replace(/\\'/g, "'");
      
      // URL内に埋め込まれた別のURLを検出
      const embeddedUrlMatch = cleanUrl.match(/https?:\/\/[^\s<>"']+/);
      if (embeddedUrlMatch && embeddedUrlMatch[0] !== cleanUrl) {
        console.log('Extracting embedded URL:', embeddedUrlMatch[0]);
        cleanUrl = embeddedUrlMatch[0];
      }
    }
    
    // 基本的なURL形式チェック
    if (!cleanUrl.match(/^https?:\/\/[^\s<>"']+$/)) {
      console.warn('Invalid URL format after sanitization:', cleanUrl);
      return getWebAppUrlCached();
    }
    
    // 開発モードURLのチェック
    if (cleanUrl.includes('googleusercontent.com') || cleanUrl.includes('userCodeAppPanel')) {
      console.warn('Development URL detected in redirect, using fallback:', cleanUrl);
      return getWebAppUrlCached();
    }
    
    // 最終的な URL 妥当性チェック
    if (!cleanUrl.includes('script.google.com') && !cleanUrl.includes('localhost')) {
      console.warn('Suspicious URL detected:', cleanUrl);
      return getWebAppUrlCached();
    }
    
    return cleanUrl;
  } catch (e) {
    console.error('URL sanitization error:', e.message);
    return getWebAppUrlCached();
  }
}

/**
 * 正しいWeb App URLを取得 (url.gsのgetWebAppUrlCachedを使用)
 * @return {string}
 */
function getWebAppUrl() {
  try {
    // url.gsの統一されたURL取得関数を使用
    return getWebAppUrlCached();
  } catch (error) {
    console.error('getWebAppUrl error:', error);
    // 緊急時のフォールバックURL
    const fallbackUrl = 'https://script.google.com/a/macros/naha-okinawa.ed.jp/s/AKfycby5oABLEuyg46OvwVqt2flUKz15zocFhH-kLD0IuNWm8akKMXiKrOS5kqGCQ7V4DQ-2/exec';
    return fallbackUrl;
  }
}


/**
 * doGet のリクエストパラメータを解析（レガシー互換版）
 * @deprecated parseUrlParameters()を使用してください
 * @param {Object} e Event object
 * @return {{mode:string,userId:string|null,setupParam:string|null,spreadsheetId:string|null,sheetName:string|null,isDirectPageAccess:boolean}}
 */
function parseRequestParams(e) {
  console.warn('parseRequestParams()は非推奨です。parseUrlParameters()を使用してください。');
  
  // 新しい統一された解析関数を使用
  const parsed = parseUrlParameters(e);
  
  // レガシー形式に変換
  return {
    mode: parsed.mode || null,
    userId: parsed.userId || null,
    setupParam: parsed.setup || null,
    spreadsheetId: parsed.spreadsheetId || null,
    sheetName: parsed.sheetName || null,
    isDirectPageAccess: !!(parsed.userId && parsed.mode === 'view')
  };
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
  
  
  const htmlOutput = adminTemplate.evaluate()
    .setTitle('みんなの回答ボード 管理パネル')
    .setSandboxMode(HtmlService.SandboxMode.NATIVE);
  
  // XFrameOptionsMode を安全に設定
  try {
    if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.ALLOWALL) {
      htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  } catch (e) {
    console.warn('XFrameOptionsMode設定エラー:', e.message);
  }
  
  return htmlOutput;
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
  // publishedSheetNameの型安全性確保（'true'問題の修正）
  let safePublishedSheetName = '';
  if (config.publishedSheetName) {
    if (typeof config.publishedSheetName === 'string') {
      safePublishedSheetName = config.publishedSheetName;
    } else {
      console.error('❌ main.gs: publishedSheetNameが不正な型です:', typeof config.publishedSheetName, config.publishedSheetName);
      console.warn('🔧 main.gs: publishedSheetNameを空文字にリセットしました');
      safePublishedSheetName = '';
    }
  }

  const isPublished = !!(config.appPublished && config.publishedSpreadsheetId && safePublishedSheetName);
  const sheetConfigKey = 'sheet_' + (safePublishedSheetName || params.sheetName);
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
      template.sheetName = escapeJavaScript(safePublishedSheetName || params.sheetName);
      template.DEBUG_MODE = shouldEnableDebugMode();
      // setupStatus未完了時の安全なopinionHeader取得
      const setupStatus = config.setupStatus || 'pending';
      let rawOpinionHeader;
      
      if (setupStatus === 'pending') {
        rawOpinionHeader = 'セットアップ中...';
      } else {
        rawOpinionHeader = sheetConfig.opinionHeader || safePublishedSheetName || 'お題';
      }
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
