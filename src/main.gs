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

// グローバル定数の定義（書き換え不可）
const SCRIPT_PROPS_KEYS = {
  SERVICE_ACCOUNT_CREDS: 'SERVICE_ACCOUNT_CREDS',
  DATABASE_SPREADSHEET_ID: 'DATABASE_SPREADSHEET_ID',
  ADMIN_EMAIL: 'ADMIN_EMAIL'
};

const DB_SHEET_CONFIG = {
  SHEET_NAME: 'Users',
  HEADERS: [
    'userId', 'adminEmail', 'spreadsheetId', 'spreadsheetUrl',
    'createdAt', 'configJson', 'lastAccessedAt', 'isActive'
  ]
};

const LOG_SHEET_CONFIG = {
  SHEET_NAME: 'Logs',
  HEADERS: ['timestamp', 'userId', 'action', 'details']
};

// 実行中のユーザー情報キャッシュ（パフォーマンス最適化用）
var _executionUserInfoCache = null;

/**
 * 実行中のユーザー情報キャッシュをクリア
 */
function clearExecutionUserInfoCache() {
  _executionUserInfoCache = null;
}

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

// 表示モード定数
const DISPLAY_MODES = {
  ANONYMOUS: 'anonymous',
  NAMED: 'named'
};

// リアクション関連定数
const REACTION_KEYS = ['UNDERSTAND', 'LIKE', 'CURIOUS'];

// スコア計算設定
const SCORING_CONFIG = {
  LIKE_MULTIPLIER_FACTOR: 0.1,
  RANDOM_SCORE_FACTOR: 0.01
};

const EMAIL_REGEX = /^[^\n@]+@[^\n@]+\.[^\n@]+$/;
const DEBUG = PropertiesService.getScriptProperties()
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

    // 2. アプリケーション有効/無効チェック（システム管理者は除外）
    const accessCheck = checkApplicationAccess();
    if (!accessCheck.hasAccess) {
      console.log('アプリケーションアクセス拒否:', accessCheck.accessReason);
      return showAccessRestrictedPage(accessCheck);
    }

    // 3. URLパラメータの解析
    const params = parseRequestParams(e);

    // 4. ログイン状態の確認
  const userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    return showLoginPage();
  }

  // 5. アプリ設定ページリクエストの処理
  if (params.setupParam === 'true') {
    return showAppSetupPage(params.userId);
  }

  // 6. パラメータ検証とデフォルト処理
    if (!params || !params.mode) {
      // パラメータなしの場合、前回の管理パネル状態を確認
      console.log('No mode parameter, checking previous admin session');
      
      var activeUserEmail = Session.getActiveUser().getEmail();
      if (activeUserEmail) {
        var userProperties = PropertiesService.getUserProperties();
        var lastAdminUserId = userProperties.getProperty('lastAdminUserId');
        
        if (lastAdminUserId) {
          console.log('Found previous admin session, redirecting to admin panel:', lastAdminUserId);
          // 前回の管理パネルセッションが存在する場合、そこにリダイレクト
          if (verifyAdminAccess(lastAdminUserId)) {
            const userInfo = findUserById(lastAdminUserId);
            return renderAdminPanel(userInfo, 'admin');
          } else {
            // 権限がない場合は状態をクリア
            userProperties.deleteProperty('lastAdminUserId');
          }
        }
      }
      
      // 前回のセッションがない場合はログインページを表示
      console.log('No previous admin session, showing login page');
      return showLoginPage();
    }

    // mode=login の場合
    if (params.mode === 'login') {
      console.log('Login mode requested, showing login page');
      return showLoginPage();
    }

    // mode=appSetup の場合
    if (params.mode === 'appSetup') {
      console.log('AppSetup mode requested');
      
      // 管理パネルへのアクセス権限確認（前回のセッションから）
      var userProperties = PropertiesService.getUserProperties();
      var lastAdminUserId = userProperties.getProperty('lastAdminUserId');
      
      if (!lastAdminUserId) {
        console.log('No admin session found, redirecting to login');
        return showErrorPage('認証が必要です', 'アプリ設定にアクセスするにはログインが必要です。');
      }
      
      // 管理者権限の確認
      if (!verifyAdminAccess(lastAdminUserId)) {
        console.log('Admin access denied for userId:', lastAdminUserId);
        return showErrorPage('アクセス拒否', 'アプリ設定にアクセスする権限がありません。');
      }
      
      console.log('Showing app setup page for userId:', lastAdminUserId);
      return showAppSetupPage(lastAdminUserId);
    }

    // mode=admin の場合
    if (params.mode === 'admin') {
      if (!params.userId) {
        return showErrorPage('不正なリクエスト', 'ユーザーIDが指定されていません。');
      }
      // 本人確認
      if (!verifyAdminAccess(params.userId)) {
        return showErrorPage('アクセス拒否', 'この管理パネルにアクセスする権限がありません。');
      }
      
      // 管理パネルアクセス時に状態を保存
      var userProperties = PropertiesService.getUserProperties();
      userProperties.setProperty('lastAdminUserId', params.userId);
      console.log('Saved admin session state:', params.userId);
      
      const userInfo = findUserById(params.userId);
      return renderAdminPanel(userInfo, 'admin');
    }

    // mode=view の場合（キャッシュバスティング対応強化）
    if (params.mode === 'view') {
      if (!params.userId) {
        return showErrorPage('不正なリクエスト', 'ユーザーIDが指定されていません。');
      }
      
      // ユーザー情報を強制的に最新状態で取得（キャッシュバイパス）
      const userInfo = findUserById(params.userId, { 
        useExecutionCache: false,
        forceRefresh: true 
      });
      
      if (!userInfo) {
        return showErrorPage('エラー', '指定されたユーザーが見つかりません。');
      }
      
      // パブリケーション状態の事前検証
      let config = {};
      try {
        config = JSON.parse(userInfo.configJson || '{}');
      } catch (e) {
        console.warn('Config JSON parse error during publication check:', e.message);
      }
      
      // 非公開の場合は確実にUnpublishedページに誘導（ErrorBoundary起動前に処理）
      const isCurrentlyPublished = !!(
        config.appPublished === true && 
        config.publishedSpreadsheetId && 
        config.publishedSheetName &&
        typeof config.publishedSheetName === 'string' &&
        config.publishedSheetName.trim() !== ''
      );
      
      console.log('🔍 Publication status check:', {
        appPublished: config.appPublished,
        hasSpreadsheetId: !!config.publishedSpreadsheetId,
        hasSheetName: !!config.publishedSheetName,
        isCurrentlyPublished: isCurrentlyPublished
      });
      
      // 非公開の場合は即座にUnpublishedページを表示
      if (!isCurrentlyPublished) {
        console.log('🚫 Board is unpublished, redirecting to Unpublished page immediately');
        return renderUnpublishedPage(userInfo, params);
      }
      
      return renderAnswerBoard(userInfo, params);
    }
    
    // 不明なモードの場合の処理
    console.log('Unknown mode received:', params.mode);
    console.log('Available modes: login, appSetup, admin, view');
    
    // 不明なモードでもuserId付きの場合は適切にリダイレクト
    if (params.userId && verifyAdminAccess(params.userId)) {
      console.log('Redirecting unknown mode to admin panel for valid user:', params.userId);
      const userInfo = findUserById(params.userId);
      return renderAdminPanel(userInfo, 'admin');
    }
    
    // その他の場合はログインページへ
    console.log('Redirecting unknown mode to login page');
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
    return redirectToUrl(correctUrl);
  }

  // 強化されたセキュリティ検証: 指定されたIDの登録メールアドレスと現在ログイン中のGoogleアカウントが一致するかを検証
  if (params.userId) {
    const isVerified = verifyAdminAccess(params.userId);
    if (!isVerified) {
      console.warn(`セキュリティ検証失敗: userId ${params.userId} への不正アクセス試行をブロックしました。`);
      const correctUrl = buildUserAdminUrl(userInfo.userId);
      return redirectToUrl(correctUrl);
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
      // キャッシュに存在しない場合はデータベースから取得する
      // 通常フローのためエラーレベルでは記録しない
      console.log('cache miss - fetching from database');

      const props = PropertiesService.getScriptProperties();
      if (!props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID)) {
        console.error('DATABASE_SPREADSHEET_ID not set');
        return null;
      }
      
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
      console.log('showAppSetupPage: Checking deploy user permissions...');
      const currentUserEmail = Session.getActiveUser().getEmail();
      console.log('showAppSetupPage: Current user email:', currentUserEmail);
      const deployUserCheckResult = isDeployUser();
      console.log('showAppSetupPage: isDeployUser() result:', deployUserCheckResult);

      if (!deployUserCheckResult) {
        console.warn('Unauthorized access attempt to app setup page:', currentUserEmail);
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
 * 最後にアクセスした管理ユーザーIDを取得
 * リダイレクト時の戻り先決定に使用
 * @returns {string|null} 管理ユーザーID、存在しない場合はnull
 */
function getLastAdminUserId() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const lastAdminUserId = userProperties.getProperty('lastAdminUserId');
    
    // ユーザーIDが存在し、かつ有効な管理者権限を持つかチェック
    if (lastAdminUserId && verifyAdminAccess(lastAdminUserId)) {
      console.log('Found valid admin user ID:', lastAdminUserId);
      return lastAdminUserId;
    } else {
      console.log('No valid admin user ID found');
      return null;
    }
  } catch (error) {
    console.error('Error getting last admin user ID:', error.message);
    return null;
  }
}

/**
 * アプリ設定ページのURLを取得（フロントエンドから呼び出し用）
 * @returns {string} アプリ設定ページのURL
 */
function getAppSetupUrl() {
  try {
    // システム管理者権限チェック
    console.log('getAppSetupUrl: Checking deploy user permissions...');
    const currentUserEmail = Session.getActiveUser().getEmail();
    console.log('getAppSetupUrl: Current user email:', currentUserEmail);
    const deployUserCheckResult = isDeployUser();
    console.log('getAppSetupUrl: isDeployUser() result:', deployUserCheckResult);

    if (!deployUserCheckResult) {
      console.warn('Unauthorized access attempt to get app setup URL:', currentUserEmail);
      throw new Error('アクセス権限がありません');
    }

    // WebアプリのベースURLを取得
    const baseUrl = ScriptApp.getService().getUrl();
    if (!baseUrl) {
      throw new Error('WebアプリのURLを取得できませんでした');
    }

    // アプリ設定ページのURLを生成
    const appSetupUrl = baseUrl + '?mode=appSetup';
    console.log('getAppSetupUrl: Generated URL:', appSetupUrl);
    
    return appSetupUrl;
  } catch (error) {
    console.error('Error getting app setup URL:', error);
    throw new Error('アプリ設定URLの取得に失敗しました: ' + error.message);
  }
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
 * アクセス制限ページを表示
 * @param {object} accessCheck - アクセスチェック結果
 * @returns {HtmlOutput}
 */
function showAccessRestrictedPage(accessCheck) {
  try {
    const template = HtmlService.createTemplateFromFile('AccessRestricted');
    template.include = include;
    template.accessCheck = accessCheck;
    template.isSystemAdmin = accessCheck.isSystemAdmin || false;
    template.userEmail = accessCheck.userEmail || '';
    template.timestamp = new Date().toISOString();
    
    const htmlOutput = template.evaluate()
      .setTitle('アクセス制限 - StudyQuest');
    
    // XFrameOptionsMode を安全に設定
    try {
      if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.DENY) {
        htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
      }
    } catch (e) {
      console.warn('XFrameOptionsMode設定エラー:', e.message);
    }
    
    return htmlOutput;
  } catch (error) {
    console.error('Error in showAccessRestrictedPage:', error);
    // フォールバック: シンプルなHTMLを返す
    return HtmlService.createHtmlOutput(`
      <html>
        <head><title>アクセス制限</title></head>
        <body style="font-family: Arial, sans-serif; text-align: center; padding: 50px;">
          <h1>アクセスが制限されています</h1>
          <p>このアプリケーションは現在利用できません。</p>
          <p>管理者にお問い合わせください。</p>
        </body>
      </html>
    `).setTitle('アクセス制限');
  }
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
 * 標準化されたページURL生成のヘルパー関数群
 */
const URLBuilder = {
  /**
   * ログインページのURLを生成
   * @returns {string} ログインページURL
   */
  login: function() {
    const baseUrl = getWebAppUrl();
    return `${baseUrl}?mode=login`;
  },
  
  /**
   * 管理パネルのURLを生成
   * @param {string} userId - ユーザーID
   * @returns {string} 管理パネルURL
   */
  admin: function(userId) {
    const baseUrl = getWebAppUrl();
    return `${baseUrl}?mode=admin&userId=${encodeURIComponent(userId)}`;
  },
  
  /**
   * アプリ設定ページのURLを生成
   * @returns {string} アプリ設定ページURL
   */
  appSetup: function() {
    const baseUrl = getWebAppUrl();
    return `${baseUrl}?mode=appSetup`;
  },
  
  /**
   * 回答ボードのURLを生成
   * @param {string} userId - ユーザーID
   * @returns {string} 回答ボードURL
   */
  view: function(userId) {
    const baseUrl = getWebAppUrl();
    return `${baseUrl}?mode=view&userId=${encodeURIComponent(userId)}`;
  },
  
  /**
   * パラメータ付きURLを安全に生成
   * @param {string} mode - モード ('login', 'admin', 'view', 'appSetup')
   * @param {Object} params - 追加パラメータ
   * @returns {string} 生成されたURL
   */
  build: function(mode, params = {}) {
    const baseUrl = getWebAppUrl();
    const url = new URL(baseUrl);
    url.searchParams.set('mode', mode);
    
    Object.keys(params).forEach(key => {
      if (params[key] !== null && params[key] !== undefined) {
        url.searchParams.set(key, params[key]);
      }
    });
    
    return url.toString();
  }
};

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
    
    // 開発モードURLのチェック（googleusercontent.comは有効なデプロイURLも含むため調整）
    if (cleanUrl.includes('userCodeAppPanel')) {
      console.warn('Development URL detected in redirect, using fallback:', cleanUrl);
      return getWebAppUrlCached();
    }
    
    // 最終的な URL 妥当性チェック（googleusercontent.comも有効URLとして認識）
    var isValidUrl = cleanUrl.includes('script.google.com') || 
                     cleanUrl.includes('googleusercontent.com') || 
                     cleanUrl.includes('localhost');
    
    if (!isValidUrl) {
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
 * doGet のリクエストパラメータを解析
 * @param {Object} e Event object
 * @return {{mode:string,userId:string|null,setupParam:string|null,spreadsheetId:string|null,sheetName:string|null,isDirectPageAccess:boolean}}
 */
function parseRequestParams(e) {
  // 引数の安全性チェック
  if (!e || typeof e !== 'object') {
    console.log('parseRequestParams: 無効なeventオブジェクト');
    return { mode: null, userId: null, setupParam: null, spreadsheetId: null, sheetName: null, isDirectPageAccess: false };
  }

  const p = e.parameter || {};
  const mode = p.mode || null;
  const userId = p.userId || null;
  const setupParam = p.setup || null;
  const spreadsheetId = p.spreadsheetId || null;
  const sheetName = p.sheetName || null;
  const isDirectPageAccess = !!(userId && mode === 'view');
  
  // デバッグログを追加
  console.log('parseRequestParams - Received parameters:', JSON.stringify(p));
  console.log('parseRequestParams - Parsed mode:', mode, 'setupParam:', setupParam);
  
  return { mode, userId, setupParam, spreadsheetId, sheetName, isDirectPageAccess };
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
    .setTitle('StudyQuest - 管理パネル')
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
/**
 * 非公開ボード専用のレンダリング関数
 * ErrorBoundaryを回避して確実にUnpublished.htmlを表示
 * @param {Object} userInfo - ユーザー情報
 * @param {Object} params - URLパラメータ
 * @returns {HtmlOutput} Unpublished.htmlテンプレート
 */
function renderUnpublishedPage(userInfo, params) {
  try {
    console.log('🚫 renderUnpublishedPage: Rendering unpublished page for userId:', userInfo.userId);
    
    const template = HtmlService.createTemplateFromFile('Unpublished');
    template.include = include;
    
    // 基本情報の設定
    template.userId = userInfo.userId;
    template.spreadsheetId = userInfo.spreadsheetId;
    template.ownerName = userInfo.adminEmail;
    template.isOwner = true;
    template.adminEmail = userInfo.adminEmail;
    template.cacheTimestamp = Date.now();
    
    // URL生成
    const appUrls = generateUserUrls(userInfo.userId);
    template.adminPanelUrl = appUrls.adminUrl;
    template.boardUrl = appUrls.viewUrl;
    
    console.log('✅ renderUnpublishedPage: Template setup completed');
    
    // キャッシュを無効化して確実なリダイレクトを保証
    return template.evaluate()
      .setTitle('StudyQuest - 準備中')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .addMetaTag('cache-control', 'no-cache, no-store, must-revalidate')
      .addMetaTag('pragma', 'no-cache')
      .addMetaTag('expires', '0');
      
  } catch (error) {
    console.error('❌ renderUnpublishedPage error:', error);
    // フォールバック: 基本的なエラーページ
    return showErrorPage('準備中', 'ボードの準備が完了していません。管理者にお問い合わせください。');
  }
}

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

  // 強化されたパブリケーション状態検証（キャッシュバスティング対応）
  const isPublished = !!(config.appPublished && config.publishedSpreadsheetId && safePublishedSheetName);
  
  // リアルタイム検証: 非公開状態の場合は確実に検出
  const isCurrentlyPublished = isPublished && 
    config.appPublished === true && 
    config.publishedSpreadsheetId && 
    safePublishedSheetName;
  
  const sheetConfigKey = 'sheet_' + (safePublishedSheetName || params.sheetName);
  const sheetConfig = config[sheetConfigKey] || {};
  
  // この関数は公開ボード専用（非公開判定は呼び出し前に完了）
  console.log('✅ renderAnswerBoard: Rendering published board for userId:', userInfo.userId);
  
  const template = HtmlService.createTemplateFromFile('Page');
  template.include = include;
  
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
    
    // 公開ボード: 通常のキャッシュ設定
    return template.evaluate()
      .setTitle('StudyQuest -みんなの回答ボード-')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
      
  } catch (error) {
    console.error('❌ renderAnswerBoard error:', error);
    // フォールバック: 基本的なエラーページ
    return showErrorPage('エラー', 'ボードの表示でエラーが発生しました。管理者にお問い合わせください。');
  }
}

/**
 * クライアントサイドからのパブリケーション状態チェック
 * キャッシュバスティング対応のため、リアルタイムで状態を確認
 * @param {string} userId - ユーザーID
 * @returns {Object} パブリケーション状態情報
 */
function checkCurrentPublicationStatus(userId) {
  try {
    console.log('🔍 checkCurrentPublicationStatus called for userId:', userId);
    
    if (!userId) {
      console.warn('userId is required for publication status check');
      return { error: 'userId is required', isPublished: false };
    }
    
    // ユーザー情報を強制的に最新状態で取得（キャッシュバイパス）
    const userInfo = findUserById(userId, { 
      useExecutionCache: false,
      forceRefresh: true 
    });
    
    if (!userInfo) {
      console.warn('User not found for publication status check:', userId);
      return { error: 'User not found', isPublished: false };
    }
    
    // 設定情報を解析
    let config = {};
    try {
      config = JSON.parse(userInfo.configJson || '{}');
    } catch (e) {
      console.warn('Config JSON parse error during publication status check:', e.message);
      return { error: 'Config parse error', isPublished: false };
    }
    
    // 現在のパブリケーション状態を厳密にチェック
    const isCurrentlyPublished = !!(
      config.appPublished === true && 
      config.publishedSpreadsheetId && 
      config.publishedSheetName &&
      typeof config.publishedSheetName === 'string' &&
      config.publishedSheetName.trim() !== ''
    );
    
    console.log('📊 Publication status check result:', {
      userId: userId,
      appPublished: config.appPublished,
      hasSpreadsheetId: !!config.publishedSpreadsheetId,
      hasSheetName: !!config.publishedSheetName,
      isCurrentlyPublished: isCurrentlyPublished,
      timestamp: new Date().toISOString()
    });
    
    return {
      isPublished: isCurrentlyPublished,
      publishedSheetName: config.publishedSheetName || null,
      publishedSpreadsheetId: config.publishedSpreadsheetId || null,
      lastChecked: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('❌ Error in checkCurrentPublicationStatus:', error);
    return { 
      error: error.message, 
      isPublished: false,
      lastChecked: new Date().toISOString()
    };
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
