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
  const tmpl = HtmlService.createTemplateFromFile(path);
  tmpl.include = include;
  return tmpl.evaluate().getContent();
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
  var dbSpreadsheetId = PropertiesService.getScriptProperties().getProperty('DATABASE_SPREADSHEET_ID');
  return !!dbSpreadsheetId;
}

/**
 * 登録ページを表示する関数
 */
function showRegistrationPage() {
  var template = HtmlService.createTemplateFromFile('Registration');
  template.include = include;
  var output = template.evaluate()
    .setTitle('新規ユーザー登録 - StudyQuest');
  return safeSetXFrameOptionsDeny(output);
}

/**
 * Webアプリケーションのメインエントリーポイント
 * @param {Object} e - URLパラメータを含むイベントオブジェクト
 * @returns {HtmlOutput} 表示するHTMLコンテンツ
 */
function doGet(e) {
  var requestKey = `doGet_${JSON.stringify(e?.parameter || {})}`;
  return cacheManager.get(requestKey, () => {
    console.log(`doGet called with event object: ${JSON.stringify(e)}`);
    try {
      const params = parseRequestParams(e);
      const { userEmail, userInfo } = validateUserSession(Session.getActiveUser().getEmail(), params);
      const setupOutput = handleSetupPages(params, userEmail);
      if (setupOutput) return setupOutput;

      if (userInfo) {
        if (params.isDirectPageAccess) {
          return renderAnswerBoard(userInfo, params);
        }
        if (params.mode === 'admin') {
          return renderAdminPanel(userInfo, params.mode);
        }
        if (params.mode === 'view') {
          return renderAnswerBoard(userInfo, params);
        }
        if (!params.userId || params.userId !== userInfo.userId) {
          const correctUrl = ScriptApp.getService().getUrl();
          return HtmlService.createHtmlOutput(`
            <script>window.top.location.href='${correctUrl}'</script>
            <p>管理パネルにリダイレクトしています...</p>`);
        }
        return renderAdminPanel(userInfo, 'admin');
      }

      return showRegistrationPage();
    } catch (error) {
      console.error(`doGetで致命的なエラー: ${error.stack}`);
      var errorHtml = HtmlService.createHtmlOutput(
        '<h1>デバッグ：致命的エラー</h1>' +
        '<p>doGet関数内でエラーが発生しました</p>' +
        '<p>エラー詳細: ' + htmlEncode(error.message) + '</p>' +
        '<p>スタック: ' + htmlEncode(error.stack || 'スタック情報なし') + '</p>' +
        '<p>時刻: ' + new Date().toISOString() + '</p>' +
        '<p>executeAs設定: USER_DEPLOYING (テスト中)</p>'
      );
      return safeSetXFrameOptionsDeny(errorHtml);
    }
  }, { ttl: 60 });
}

/**
 * doGet のリクエストパラメータを解析
 * @param {Object} e Event object
 * @return {{mode:string,userId:string|null,setupParam:string|null,spreadsheetId:string|null,sheetName:string|null,isDirectPageAccess:boolean}}
 */
function parseRequestParams(e) {
  const p = (e && e.parameter) || {};
  const mode = p.mode || 'admin';
  const userId = p.userId || null;
  const setupParam = p.setup || null;
  const spreadsheetId = p.spreadsheetId || null;
  const sheetName = p.sheetName || null;
  const isDirectPageAccess = !!(userId && mode === 'view');
  return { mode, userId, setupParam, spreadsheetId, sheetName, isDirectPageAccess };
}

/**
 * セッション検証とユーザー情報取得
 * @param {string} userEmail 現在のユーザーのメール
 * @param {Object} params リクエストパラメータ
 * @return {{userEmail:string|null,userInfo:Object|null}}
 */
function validateUserSession(userEmail, params) {
  if (userEmail && !params.isDirectPageAccess) {
    try {
      validateAndRepairSession(userEmail);
    } catch (sessionError) {
      console.error('セッション管理エラー: ' + sessionError.message);
    }
  }

  let userInfo = null;
  if (params.isDirectPageAccess) {
    userInfo = findUserById(params.userId);
  } else if (userEmail) {
    userInfo = findUserByEmail(userEmail);
    if (userInfo && userInfo.spreadsheetId) {
      try {
        const testAccess = SpreadsheetApp.openById(userInfo.spreadsheetId);
        testAccess.getName();
      } catch (accessError) {
        console.warn('スプレッドシートアクセスエラー: ' + accessError.message);
        try {
          const repair = repairUserSpreadsheetAccess(userEmail, userInfo.spreadsheetId);
          if (repair.success) console.log('スプレッドシートアクセス権限を修復しました');
        } catch (repairError) {
          console.error('権限修復失敗: ' + repairError.message);
        }
      }
    }
  }
  return { userEmail, userInfo };
}

/**
 * セットアップページ関連の処理
 * @param {Object} params リクエストパラメータ
 * @param {string} userEmail 現在のユーザーのメール
 * @return {HtmlOutput|null} 表示するHTMLがあれば返す
 */
function handleSetupPages(params, userEmail) {
  if (!isSystemSetup() && !params.isDirectPageAccess) {
    const t = HtmlService.createTemplateFromFile('SetupPage');
    t.include = include;
    return safeSetXFrameOptionsDeny(t.evaluate().setTitle('初回セットアップ - StudyQuest'));
  }

  if (params.setupParam === 'true' && params.mode === 'appsetup') {
    const domainInfo = getDeployUserDomainInfo();
    if (domainInfo.deployDomain && domainInfo.deployDomain !== '' && !domainInfo.isDomainMatch) {
      console.log('Domain access warning. Current:', domainInfo.currentDomain, 'Deploy:', domainInfo.deployDomain);
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
  const adminTemplate = HtmlService.createTemplateFromFile('AdminPanel');
  adminTemplate.include = include;
  adminTemplate.userInfo = userInfo;
  adminTemplate.userId = userInfo.userId;
  adminTemplate.mode = mode;
  adminTemplate.displayMode = 'named';
  adminTemplate.showAdminFeatures = true;
  return safeSetXFrameOptionsDeny(
    adminTemplate.evaluate()
      .setTitle('みんなの回答ボード 管理パネル')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
  );
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
