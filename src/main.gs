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
  try {
    debugLog('doGet called with event object:', e);
    
    const params = parseRequestParams(e);
    const currentUserEmail = Session.getActiveUser().getEmail();
    
    // マルチテナント対応: 認証確認のみでセッション依存を削除
    console.log('doGet - currentUserEmail:', currentUserEmail, 'requestedUserId:', params.userId);
    
    // /exec直接アクセス時の認証チェック（管理パネルへのアクセスを完全に防止）
    if (isDirectExecAccess(e)) {
      return handleDirectExecAccess(currentUserEmail);
    }
    
    // マルチテナント対応: userIdベースでユーザー情報を取得
    const { userEmail, userInfo } = validateUserSession(currentUserEmail, params);
    const setupOutput = handleSetupPages(params, userEmail);
    if (setupOutput) return setupOutput;

    if (userInfo) {
      // 管理パネルアクセス時の厳格なセキュリティチェック
      if (params.mode === 'admin') {
        console.log('Admin panel access - params.userId:', params.userId);
        console.log('Admin panel access - userInfo.userId:', userInfo.userId);
        console.log('Admin panel access - currentUserEmail:', currentUserEmail);
        console.log('Admin panel access - userInfo.adminEmail:', userInfo.adminEmail);
        
        if (!params.userId) {
          // userIdパラメータが無い場合は自分のIDでリダイレクト
          const correctUrl = buildUserAdminUrl(userInfo.userId);
          console.log('Missing userId, redirecting to:', correctUrl);
          return createSecureRedirect(correctUrl, 'ユーザー専用管理パネルにリダイレクトしています...');
        }
        
        if (params.userId !== userInfo.userId) {
          // 他人のuserIdでアクセスしようとした場合
          console.warn(`Unauthorized access attempt: ${currentUserEmail} tried to access userId: ${params.userId}`);
          const correctUrl = buildUserAdminUrl(userInfo.userId);
          console.log('Unauthorized access, redirecting to:', correctUrl);
          return createSecureRedirect(correctUrl, '正しい管理パネルにリダイレクトしています...');
        }
        
        // 正当なアクセス（正しいuserIdでのアクセス）
        console.log('Valid admin panel access for userId:', params.userId);
        return renderAdminPanel(userInfo, params.mode);
      }
      
      if (params.isDirectPageAccess) {
        return renderAnswerBoard(userInfo, params);
      }
      
      if (params.mode === 'view') {
        return renderAnswerBoard(userInfo, params);
      }
      
      // デフォルトは管理パネル
      const correctUrl = buildUserAdminUrl(userInfo.userId);
      return createSecureRedirect(correctUrl, '管理パネルにリダイレクトしています...');
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
}

/**
 * /exec直接アクセスかどうかを判定
 * @param {Object} e Event object
 * @return {boolean}
 */
function isDirectExecAccess(e) {
  const params = (e && e.parameter) || {};
  // パラメータが空の場合のみ直接アクセス（管理パネルアクセスを完全防止）
  return Object.keys(params).length === 0;
}

/**
 * /exec直接アクセス時の処理
 * @param {string} userEmail ユーザーメールアドレス
 * @return {HtmlOutput}
 */
function handleDirectExecAccess(userEmail) {
  try {
    debugLog('Direct /exec access detected for user:', userEmail);
    
    // システムセットアップ未完了の場合は、セットアップページを優先表示
    if (!isSystemSetup()) {
      console.log('handleDirectExecAccess - System not setup, showing setup page');
      const t = HtmlService.createTemplateFromFile('SetupPage');
      t.include = include;
      return safeSetXFrameOptionsDeny(t.evaluate().setTitle('初回セットアップ - StudyQuest'));
    }
    
    if (!userEmail) {
      return showRegistrationPage();
    }
    
    // サービスアカウント経由でユーザーがデータベースに登録されているかチェック
    const userInfo = findUserByEmail(userEmail);
    console.log('handleDirectExecAccess - userInfo:', userInfo);
    console.log('handleDirectExecAccess - userEmail:', userEmail);
    
    if (userInfo && userInfo.userId) {
      // 登録済みユーザー: 管理パネルに自動遷移（リダイレクトではなく直接遷移）
      console.log('handleDirectExecAccess - Found user, transitioning to admin panel for userId:', userInfo.userId);
      
      // ここで直接管理パネルを表示する（リダイレクトしない）
      return renderAdminPanel(userInfo, 'admin');
    } else {
      // 未登録ユーザー: 新規登録画面表示
      console.log('handleDirectExecAccess - Unregistered user, showing registration page');
      debugLog('Unregistered user, showing registration page');
      return showRegistrationPage();
    }
  } catch (error) {
    console.error('handleDirectExecAccess error:', error);
    return showRegistrationPage();
  }
}

/**
 * ユーザー専用の一意の管理パネルURLを構築
 * @param {string} userId ユーザーID
 * @return {string}
 */
function buildUserAdminUrl(userId) {
  const baseUrl = getWebAppUrl();
  const adminUrl = `${baseUrl}?mode=admin&userId=${encodeURIComponent(userId)}`;
  console.log('buildUserAdminUrl - baseUrl:', baseUrl);
  console.log('buildUserAdminUrl - userId:', userId);
  console.log('buildUserAdminUrl - adminUrl:', adminUrl);
  return adminUrl;
}

/**
 * セキュアなリダイレクトHTMLを作成
 * @param {string} targetUrl リダイレクト先URL
 * @param {string} message 表示メッセージ
 * @return {HtmlOutput}
 */
function createSecureRedirect(targetUrl, message) {
  return HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta http-equiv="refresh" content="2;url=${targetUrl}">
      <title>リダイレクト中...</title>
    </head>
    <body style="text-align:center; padding:50px; font-family:sans-serif; background-color:#f5f5f5;">
      <div style="max-width:600px; margin:0 auto; background:white; padding:40px; border-radius:10px; box-shadow:0 2px 10px rgba(0,0,0,0.1);">
        <h2 style="color:#333; margin-bottom:20px;">${message}</h2>
        <p style="color:#666; margin-bottom:30px;">2秒後に自動的にリダイレクトされます...</p>
        <a href="${targetUrl}" 
           style="display:inline-block; background:#4CAF50; color:white; padding:12px 24px; text-decoration:none; border-radius:5px; font-weight:bold;"
           onclick="handleRedirectClick(event)">
          管理パネルへ移動
        </a>
        <p style="color:#999; font-size:14px; margin-top:20px;">
          自動的にリダイレクトされない場合は上のボタンをクリックしてください。
        </p>
      </div>
      <script>
        function handleRedirectClick(event) {
          event.preventDefault();
          try {
            // ユーザーアクションによるリダイレクト
            window.top.location.href = '${targetUrl}';
          } catch(e) {
            window.location.href = '${targetUrl}';
          }
        }
        
        // 2秒後の自動リダイレクト（ユーザーアクション後）
        setTimeout(function() {
          try {
            window.top.location.href = '${targetUrl}';
          } catch(e) {
            window.location.href = '${targetUrl}';
          }
        }, 2000);
      </script>
    </body>
    </html>
  `).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 正しいWeb App URLを取得
 * @return {string}
 */
function getWebAppUrl() {
  try {
    // まずScriptApp.getService().getUrl()を試す
    let url = ScriptApp.getService().getUrl();
    
    // URLが正しい形式かチェック
    if (url && url.includes('/macros/') && url.endsWith('/exec')) {
      return url;
    }
    
    // フォールバック: 現在のリクエストURLから構築
    const currentUrl = Session.getActiveUser().getEmail() ? 
      'https://script.google.com/a/macros/naha-okinawa.ed.jp/s/AKfycby5oABLEuyg46OvwVqt2flUKz15zocFhH-kLD0IuNWm8akKMXiKrOS5kqGCQ7V4DQ-2/exec' :
      url;
    
    console.log('Using Web App URL:', currentUrl);
    return currentUrl;
    
  } catch (error) {
    console.error('getWebAppUrl error:', error);
    // ハードコードされたフォールバック
    return 'https://script.google.com/a/macros/naha-okinawa.ed.jp/s/AKfycby5oABLEuyg46OvwVqt2flUKz15zocFhH-kLD0IuNWm8akKMXiKrOS5kqGCQ7V4DQ-2/exec';
  }
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
    
    // AppSetupPage用のuserIdを設定
    const currentUserEmail = Session.getActiveUser().getEmail();
    const currentUserInfo = findUserByEmail(currentUserEmail);
    appSetupTemplate.userId = currentUserInfo ? currentUserInfo.userId : '';
    
    console.log('AppSetupPage - currentUserEmail:', currentUserEmail);
    console.log('AppSetupPage - userId set to:', appSetupTemplate.userId);
    
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
  const adminTemplate = HtmlService.createTemplateFromFile('AdminPanel');
  adminTemplate.include = include;
  adminTemplate.userInfo = userInfo;
  adminTemplate.userId = userInfo.userId;
  adminTemplate.mode = mode;
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
