/**
 * @fileoverview メインエントリーポイント - 2024年最新技術の結集
 * V8ランタイム、最新パフォーマンス技術、安定性強化を統合
 */

// グローバル定数の定義
var SCRIPT_PROPS_KEYS = {
  SERVICE_ACCOUNT_CREDS: 'SERVICE_ACCOUNT_CREDS',
  DATABASE_SPREADSHEET_ID: 'DATABASE_SPREADSHEET_ID'
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
var DEBUG = true;

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

/**
 * デプロイされたWebアプリのドメイン情報と現在のユーザーのドメイン情報を取得
 * AdminPanel.html と Registration.html から共通で呼び出される
 */
function getDeployUserDomainInfo() {
  try {
    var activeUserEmail = Session.getActiveUser().getEmail();
    var currentDomain = getEmailDomain(activeUserEmail);
    var webAppUrl = ScriptApp.getService().getUrl();
    var deployDomain = ''; // 個人アカウント/グローバルアクセスの場合、デフォルトで空

    if (webAppUrl) {
      // Google WorkspaceアカウントのドメインをURLから抽出を試みる (例: /a/domain.com/)
      var domainMatch = webAppUrl.match(/\/a\/([a-zA-Z0-9\-\.]+)\/macros/);
      if (domainMatch && domainMatch[1]) {
        deployDomain = domainMatch[1];
      }
      // ドメインが抽出されなかった場合、deployDomainは空のままとなり、個人アカウント/グローバルアクセスを示す
    }

    // 現在のユーザーのドメインと抽出された/デフォルトのデプロイドメインを比較
    // deployDomainが空の場合、特定のドメインが強制されていないため、一致とみなす（グローバルアクセス）
    var isDomainMatch = (currentDomain === deployDomain) || (deployDomain === '');

    console.log('Domain info:', {
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

function performSimpleCleanup() {
  try {
    // 期限切れPropertiesServiceエントリの削除
    var props = PropertiesService.getScriptProperties();
    var allProps = props.getProperties();
    var now = Date.now();
    
    Object.keys(allProps).forEach(function(key) {
      if (key.startsWith('CACHE_')) {
        try {
          var data = JSON.parse(allProps[key]);
          if (data.expiresAt && data.expiresAt < now) {
            props.deleteProperty(key);
          }
        } catch (e) {
          // 無効なデータは削除
          props.deleteProperty(key);
        }
      }
    });
  } catch (e) {
    console.warn('Cleanup failed:', e.message);
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
function isUserRegistered(email, spreadsheetId) {
  try {
    if (!email || !spreadsheetId) {
      return false;
    }
    
    // findUserByEmail関数を使用してユーザーを検索
    var userInfo = findUserByEmail(email);
    return !!userInfo; // userInfoが存在すればtrue、存在しなければfalse
    
  } catch(e) {
    console.error(`ユーザー登録確認中にエラー: ${e.message}`);
    // エラー時は安全のため「未登録」として扱う
    return false;
  }
}

/**
 * システムの初期セットアップが完了しているかを確認するヘルパー関数
 * @returns {boolean} セットアップが完了していればtrue
 */
function isSystemSetup() {
  var dbSpreadsheetId = PropertiesService.getScriptProperties().getProperty('DATABASE_SPREADSHEET_ID');
  return !!dbSpreadsheetId;
}

/**
 * 登録ページを表示する関数
 */
function showRegistrationPage() {
  var output = HtmlService.createTemplateFromFile('Registration')
    .evaluate()
    .setTitle('新規ユーザー登録 - StudyQuest');
  return safeSetXFrameOptionsDeny(output);
}

/**
 * Webアプリケーションのメインエントリーポイント
 * @param {Object} e - URLパラメータを含むイベントオブジェクト
 * @returns {HtmlOutput} 表示するHTMLコンテンツ
 */
function doGet(e) {
  // ① パラメータe全体をログに出力して、どのようなリクエストか確認する
  console.log(`doGet called with event object: ${JSON.stringify(e)}`);

  try {
    // ② 'mode' パラメータの存在を確認し、なければデフォルト値を設定
    // e.parameter が null または undefined の場合を考慮
    const mode = (e && e.parameter && e.parameter.mode) ? e.parameter.mode : 'view'; // デフォルトは 'view' モード
    const userId = (e && e.parameter && e.parameter.userId) ? e.parameter.userId : null;
    const setupParam = (e && e.parameter && e.parameter.setup) ? e.parameter.setup : null;
    const spreadsheetId = (e && e.parameter && e.parameter.spreadsheetId) ? e.parameter.spreadsheetId : null;
    const sheetName = (e && e.parameter && e.parameter.sheetName) ? e.parameter.sheetName : null;
    
    // Page.html への直接アクセス判定（特定パラメータがある場合）
    const isDirectPageAccess = userId && spreadsheetId && sheetName;

    console.log(`Request parameters validated. Mode: ${mode}, UserID: ${userId}, SetupParam: ${setupParam}, IsDirectPageAccess: ${isDirectPageAccess}`);

    // どこからのアクセスかを知るためのログを追加
    const userEmail = Session.getActiveUser().getEmail();
    console.log(`Access by user: ${userEmail}`);

    const webAppUrl = ScriptApp.getService().getUrl();
    console.log(`Current Web App URL: ${webAppUrl}`);

    if (e && e.headers && e.headers.referer) {
      console.log(`Referer: ${e.headers.referer}`);
    }

    // 1. システムの初期セットアップが完了しているか確認（Page.html直接アクセス時は除く）
    if (!isSystemSetup() && !isDirectPageAccess) {
      console.log('DEBUG: System not set up. Redirecting to SetupPage.');
      var setupHtml = HtmlService.createTemplateFromFile('SetupPage')
        .evaluate()
        .setTitle('初回セットアップ - StudyQuest');
      return safeSetXFrameOptionsDeny(setupHtml);
    }

    // セットアップページの明示的な表示要求
    if (setupParam === 'true') {
      console.log('DEBUG: Explicit setup request. Redirecting to SetupPage.');
      var explicitHtml = HtmlService.createTemplateFromFile('SetupPage')
        .evaluate()
        .setTitle('StudyQuest - サービスアカウント セットアップ');
      return safeSetXFrameOptionsDeny(explicitHtml);
    }

    // 2. ユーザー認証と情報取得（Page.html直接アクセス時は除く）
    // currentUserEmail is already defined above as userEmail
    if (!userEmail && !isDirectPageAccess) {
      console.log('DEBUG: No current user email. Redirecting to RegistrationPage.');
      return showRegistrationPage();
    }

    // 3. ユーザーがデータベースに登録済みか確認
    let userInfo = null;
    if (isDirectPageAccess) {
      // Page.html直接アクセス時は、userIdで直接検索
      userInfo = findUserById(userId);
      console.log('DEBUG: Direct page access, User info by ID:', JSON.stringify(userInfo));
    } else {
      userInfo = findUserByEmail(userEmail);
      console.log('DEBUG: User info by email:', JSON.stringify(userInfo));
    }

    if (userInfo) {
      var configJson = {};
      try {
        configJson = JSON.parse(userInfo.configJson || '{}');
      } catch (jsonErr) {
        console.warn('Invalid configJson:', jsonErr.message);
      }

      var isPublished = !!(configJson.appPublished &&
        configJson.publishedSpreadsheetId &&
        configJson.publishedSheetName);

      // Page.html直接アクセス時は、パラメータ指定されたページを表示
      if (isDirectPageAccess) {
        console.log('DEBUG: Direct page access detected. Showing specified page.');
        var pageTemplate = HtmlService.createTemplateFromFile('Page');
        pageTemplate.userId = userId;
        pageTemplate.sheetName = sheetName;
        pageTemplate.spreadsheetId = spreadsheetId;
        pageTemplate.displayMode = configJson.displayMode || DISPLAY_MODES.ANONYMOUS;
        pageTemplate.showCounts = configJson.showCounts !== false;
        var key = 'sheet_' + spreadsheetId + '_' + sheetName;
        pageTemplate.mapping = configJson[key] || {};
        pageTemplate.ownerName = userInfo.adminEmail;
        var isOwner = (userEmail === userInfo.adminEmail);
        pageTemplate.showAdminFeatures = isOwner;
        pageTemplate.showHighlightToggle = isOwner;
        pageTemplate.showScoreSort = true;
        pageTemplate.isStudentMode = !isOwner;
        pageTemplate.isAdminUser = isOwner;
        var pageHtml = pageTemplate.evaluate()
          .setTitle('回答ボード - みんなの回答ボード');
        return safeSetXFrameOptionsDeny(pageHtml);
      }

      if (mode === 'admin') {
        // 管理パネル表示
        var adminTemplate = HtmlService.createTemplateFromFile('AdminPanel');
        adminTemplate.userInfo = userInfo;
        adminTemplate.userId = userInfo.userId;
        adminTemplate.mode = mode;
        adminTemplate.displayMode = 'named';
        adminTemplate.showAdminFeatures = true;
        console.log('DEBUG: AdminPanel template properties:', {
          userInfo: adminTemplate.userInfo,
          userId: adminTemplate.userId,
          mode: adminTemplate.mode,
          displayMode: adminTemplate.displayMode,
          showAdminFeatures: adminTemplate.showAdminFeatures
        });
        var adminHtml = adminTemplate.evaluate()
          .setTitle('管理パネル - みんなの回答ボード');
        return safeSetXFrameOptionsDeny(adminHtml);
      }

      if (isPublished) {
        // 公開ボード表示
        var pageTemplate = HtmlService.createTemplateFromFile('Page');
        var sheetName = configJson.publishedSheetName;
        var spreadsheetId = configJson.publishedSpreadsheetId;
        pageTemplate.userId = userInfo.userId;
        pageTemplate.sheetName = sheetName;
        pageTemplate.spreadsheetId = spreadsheetId;
        pageTemplate.displayMode = configJson.displayMode || DISPLAY_MODES.ANONYMOUS;
        pageTemplate.showCounts = configJson.showCounts !== false;
        var key = 'sheet_' + spreadsheetId + '_' + sheetName;
        pageTemplate.mapping = configJson[key] || {};
        pageTemplate.ownerName = userInfo.adminEmail;
        var isOwner = (userEmail === userInfo.adminEmail);
        pageTemplate.showAdminFeatures = isOwner;
        pageTemplate.showHighlightToggle = isOwner;
        pageTemplate.showScoreSort = true;
        pageTemplate.isStudentMode = !isOwner;
        pageTemplate.isAdminUser = isOwner;
        var pageHtml = pageTemplate.evaluate()
          .setTitle('回答ボード - みんなの回答ボード');
        return safeSetXFrameOptionsDeny(pageHtml);
      }

      // 未公開ボード（管理パネルにリダイレクト）
      console.log('DEBUG: Board not published. Redirecting to admin panel.');
      var adminTemplate = HtmlService.createTemplateFromFile('AdminPanel');
      adminTemplate.userInfo = userInfo;
      adminTemplate.userId = userInfo.userId;
      adminTemplate.mode = 'admin';
      adminTemplate.displayMode = 'named';
      adminTemplate.showAdminFeatures = true;
      var adminHtml = adminTemplate.evaluate()
        .setTitle('管理パネル - みんなの回答ボード');
      return safeSetXFrameOptionsDeny(adminHtml);

    } else {
      // 5. 【未登録ユーザーの処理】
      //    Page.html直接アクセス時でもユーザーが見つからない場合は登録ページ
      console.log('DEBUG: User not registered. Redirecting to RegistrationPage.');
      return showRegistrationPage();
    }

  } catch (error) {
    console.error(`doGetで致命的なエラー: ${error.stack}`);
    var errorHtml = HtmlService.createHtmlOutput(
      '<h1>エラー</h1>' +
      '<p>予期せぬエラーが発生しました。管理者にお問い合わせください。</p>' +
      '<p>エラー詳細: ' + htmlEncode(error.message) + '</p>'
    );
    return safeSetXFrameOptionsDeny(errorHtml);
  }
}

/**
 * HTMLエンコードを行うヘルパー関数
 * @param {string} str - エンコードする文字列
 * @returns {string} エンコードされた文字列
 */
function htmlEncode(str) {
  if (typeof str !== 'string') return '';
  return str.replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
}

/**
 * パフォーマンス監視エンドポイント（簡易版）
 */
function getPerformanceMetrics() {
  try {
    var metrics = {};
    
    // プロファイラー情報
    if (typeof globalProfiler !== 'undefined') {
      metrics.profiler = globalProfiler.getReport();
    }
    
    // キャッシュ健康状態
    metrics.cache = cacheManager.getHealth();
    
    // 基本システム情報
    metrics.system = {
      timestamp: Date.now(),
      status: 'running'
    };
    
    return metrics;
  } catch (e) {
    return {
      error: e.message,
      timestamp: Date.now()
    };
  }
}