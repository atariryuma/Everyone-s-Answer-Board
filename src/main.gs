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
    'createdAt', 'configJson', 'setupStatus', 'lastAccessedAt', 'isActive'
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

// セットアップ状態定数
var SETUP_STATUS = {
  UNREGISTERED: 'unregistered',
  ACCOUNT_CREATED: 'account_created',    // 旧: basic
  RESOURCES_PENDING: 'resources_pending', // リソース作成中
  DATA_PREPARED: 'data_prepared',        // データ準備完了
  CONFIGURED: 'configured',              // 設定完了
  PUBLISHED: 'published'                 // 公開済み（旧: complete）
};

// セットアップステップ定数
var SETUP_STEPS = {
  ACCOUNT: 1,      // アカウント作成
  DATA_PREP: 2,    // データ準備（スプレッドシート・フォーム作成）
  DATA_INPUT: 3,   // データ投入（シート選択・列設定）
  PUBLISH_CONFIG: 4, // 公開設定（表示オプション設定）
  PREVIEW: 5,      // プレビュー確認
  PUBLISHED: 6     // 回答ボード公開
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
function safeSetXFrameOptionsAllowAll(htmlOutput) {
  try {
    if (htmlOutput && typeof htmlOutput.setXFrameOptionsMode === 'function' &&
        HtmlService && HtmlService.XFrameOptionsMode &&
        HtmlService.XFrameOptionsMode.ALLOWALL) {
      htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.SAMEORIGIN);
    }
    if (htmlOutput && typeof htmlOutput.setSandboxMode === 'function' &&
        HtmlService && HtmlService.SandboxMode &&
        HtmlService.SandboxMode.IFRAME) {
      htmlOutput.setSandboxMode(HtmlService.SandboxMode.IFRAME);
    }
  } catch (e) {
    console.warn('Failed to set frame options and sandbox mode:', e.message);
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
 * ログインページを表示する関数
 */
function showLoginPage() {
  var template = HtmlService.createTemplateFromFile('LoginPage');
  template.include = include;
  return safeSetXFrameOptionsAllowAll(template.evaluate().setTitle('StudyQuest - ログイン'));
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

    if (params.isDirectPageAccess) {
      const publicInfo = findUserById(params.userId);
      if (publicInfo) {
        return renderAnswerBoard(publicInfo, params);
      }
    }

    return showLoginPage();
  } catch (error) {
    console.error(`doGetで致命的なエラー: ${error.stack}`);
    
    var errorMessage = error.message;
    var isPermissionError = errorMessage && (
      errorMessage.includes('permission') || 
      errorMessage.includes('権限') ||
      errorMessage.includes('access')
    );
    
    var errorHtml;
    if (isPermissionError) {
      errorHtml = HtmlService.createHtmlOutput(
        '<h1>データベースアクセスエラー</h1>' +
        '<p>データベースへのアクセスに失敗しました。</p>' +
        '<p>このアプリではデータベース操作をサービスアカウント経由でのみ行います。</p>' +
        '<h2>管理者へ</h2>' +
        '<ul>' +
        '<li>サービスアカウントの認証情報が正しく設定されているか確認してください</li>' +
        '<li>データベーススプレッドシートがサービスアカウントと共有されているか確認してください</li>' +
        '<li>サービスアカウントがデータベーススプレッドシートの編集者権限を持っているか確認してください</li>' +
        '</ul>' +
        '<p><small>セキュリティ設定: データベースはサービスアカウント専用</small></p>'
      );
    } else {
      errorHtml = HtmlService.createHtmlOutput(
        '<h1>システムエラー</h1>' +
        '<p>アプリの動作中にエラーが発生しました</p>' +
        '<p>エラー詳細: ' + htmlEncode(error.message) + '</p>' +
        '<p>時刻: ' + new Date().toISOString() + '</p>' +
        '<p>管理者にお問い合わせください。</p>' +
        '<p><small>executeAs設定: USER_ACCESSING / データベース: サービスアカウント専用</small></p>'
      );
    }
    
    return safeSetXFrameOptionsAllowAll(errorHtml);
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
      return safeSetXFrameOptionsAllowAll(t.evaluate().setTitle('初回セットアップ - StudyQuest'));
    }
    
    if (!userEmail) {
      return showLoginPage();
    }
    
    // サービスアカウント経由でユーザーがデータベースに登録されているかチェック
    // 認証済みユーザーは常に登録ページを表示（管理パネルへのアクセスはボタン経由）
    console.log('handleDirectExecAccess - Authenticated user, showing registration page');
    debugLog('Authenticated user, showing registration page');
    return showLoginPage();
  } catch (error) {
    console.error('handleDirectExecAccess error:', error);
    return showLoginPage();
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
 * サーバーサイドでのナビゲーション処理
 * @param {string} targetUrl リダイレクト先URL
 * @param {string} message 表示メッセージ
 * @return {HtmlOutput}
 */
function createServerSideNavigation(targetUrl, message) {
  return HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta http-equiv="refresh" content="1;url=${targetUrl}">
      <title>移動中...</title>
      <style>
        body {
          font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
          background: linear-gradient(135deg, #1a1b26 0%, #0f172a 50%, #1e1b4b 100%);
          color: white;
          display: flex;
          align-items: center;
          justify-content: center;
          min-height: 100vh;
          margin: 0;
          text-align: center;
        }
        .container {
          background: rgba(255, 255, 255, 0.08);
          backdrop-filter: blur(20px);
          border-radius: 16px;
          border: 1px solid rgba(255, 255, 255, 0.2);
          padding: 40px;
          max-width: 400px;
        }
        .spinner {
          width: 40px;
          height: 40px;
          border: 3px solid rgba(255, 255, 255, 0.3);
          border-top: 3px solid #06b6d4;
          border-radius: 50%;
          animation: spin 1s linear infinite;
          margin: 0 auto 20px;
        }
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
        .btn {
          display: inline-block;
          background: linear-gradient(135deg, #06b6d4 0%, #a855f7 100%);
          color: white;
          padding: 12px 24px;
          text-decoration: none;
          border-radius: 8px;
          margin-top: 20px;
          transition: transform 0.2s;
        }
        .btn:hover {
          transform: translateY(-1px);
        }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="spinner"></div>
        <h2>${message}</h2>
        <p>自動的に移動しない場合は下のボタンをクリックしてください</p>
        <a href="${targetUrl}" class="btn">続行</a>
      </div>
      <script>
        // Fallback navigation after 2 seconds
        setTimeout(function() {
          window.location.href = '${targetUrl}';
        }, 2000);
      </script>
    </body>
    </html>
  `)
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.SAMEORIGIN)
  .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

/**
 * セキュアなリダイレクトHTMLを作成（レガシー用）
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
            // Use iframe-safe navigation
            if (window.sharedUtilities && window.sharedUtilities.navigation) {
              window.sharedUtilities.navigation.safeNavigate('${targetUrl}', { method: 'redirect' });
            } else {
              window.location.href = '${targetUrl}';
            }
          } catch(e) {
            window.location.href = '${targetUrl}';
          }
        }
        
        // 2秒後の自動リダイレクト（iframe対応）
        setTimeout(function() {
          try {
            if (window.sharedUtilities && window.sharedUtilities.navigation) {
              window.sharedUtilities.navigation.safeNavigate('${targetUrl}', { method: 'redirect' });
            } else {
              window.location.href = '${targetUrl}';
            }
          } catch(e) {
            window.location.href = '${targetUrl}';
          }
        }, 2000);
      </script>
    </body>
    </html>
  `).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.SAMEORIGIN);
}

/**
 * サーバーサイドナビゲーション処理関数
 * フロントエンドから呼び出し可能
 * @param {string} targetUrl 移動先URL
 * @param {string} message 表示メッセージ
 * @return {Object} 移動情報
 */
function performServerSideNavigation(targetUrl, message = '移動中...') {
  try {
    // URLの妥当性チェック
    if (!targetUrl) {
      throw new Error('移動先URLが指定されていません');
    }
    
    // 相対URLの場合は絶対URLに変換
    if (targetUrl.startsWith('?') || targetUrl.startsWith('/')) {
      const baseUrl = getWebAppUrl();
      targetUrl = targetUrl.startsWith('?') ? baseUrl + targetUrl : baseUrl + targetUrl;
    }
    
    console.log('performServerSideNavigation:', { targetUrl, message });
    
    return {
      success: true,
      redirectUrl: targetUrl,
      message: message,
      method: 'server_navigation'
    };
  } catch (error) {
    console.error('performServerSideNavigation error:', error);
    return {
      success: false,
      error: error.message,
      method: 'server_navigation'
    };
  }
}

/**
 * 登録後のナビゲーション処理
 * @param {string} userId ユーザーID
 * @return {Object} 移動情報
 */
function navigateToAdminPanel(userId) {
  try {
    const adminUrl = buildUserAdminUrl(userId);
    return performServerSideNavigation(adminUrl, '管理パネルに移動中...');
  } catch (error) {
    console.error('navigateToAdminPanel error:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * セットアップ状態とステップを判定
 * @param {Object} userInfo ユーザー情報
 * @return {Object} 状態とステップ情報
 */
function determineSetupProgress(userInfo) {
  try {
    if (!userInfo) {
      return {
        status: SETUP_STATUS.UNREGISTERED,
        step: SETUP_STEPS.ACCOUNT,
        nextAction: 'アカウントを作成してください',
        description: 'まずはアカウント登録から始めましょう'
      };
    }

    const config = JSON.parse(userInfo.configJson || '{}');
    const hasSpreadsheet = !!(userInfo.spreadsheetId && userInfo.spreadsheetUrl);
    const hasForm = !!(config.formUrl);
    const hasSheetConfig = !!(config.publishedSheetName);
    const isPublished = !!(config.appPublished && config.publishedSpreadsheetId);

    console.log('determineSetupProgress: 状態分析', {
      setupStatus: userInfo.setupStatus,
      hasSpreadsheet,
      hasForm,
      hasSheetConfig,
      isPublished
    });

    // 状態判定ロジック
    if (isPublished) {
      return {
        status: SETUP_STATUS.PUBLISHED,
        step: SETUP_STEPS.PUBLISHED,
        nextAction: '回答ボードが公開中です',
        description: 'ボードの管理と監視を行えます'
      };
    }
    
    if (hasSheetConfig && hasSpreadsheet && hasForm) {
      return {
        status: SETUP_STATUS.CONFIGURED,
        step: SETUP_STEPS.PREVIEW,
        nextAction: 'プレビューを確認して公開する',
        description: '設定が完了しました。プレビューを確認して公開しましょう'
      };
    }
    
    if (hasSpreadsheet && hasForm) {
      return {
        status: SETUP_STATUS.DATA_PREPARED,
        step: SETUP_STEPS.DATA_INPUT,
        nextAction: 'シートを選択して列を設定する',
        description: 'データの準備ができました。シートと列の設定を行いましょう'
      };
    }
    
    if (userInfo.setupStatus === 'basic' || userInfo.setupStatus === 'account_created') {
      return {
        status: SETUP_STATUS.ACCOUNT_CREATED,
        step: SETUP_STEPS.DATA_PREP,
        nextAction: 'データ準備を完了する',
        description: 'スプレッドシートとフォームを作成しましょう'
      };
    }
    
    // フォールバック
    return {
      status: userInfo.setupStatus || SETUP_STATUS.ACCOUNT_CREATED,
      step: SETUP_STEPS.DATA_PREP,
      nextAction: 'セットアップを続行する',
      description: 'セットアップを完了しましょう'
    };

  } catch (error) {
    console.error('determineSetupProgress error:', error);
    return {
      status: SETUP_STATUS.ACCOUNT_CREATED,
      step: SETUP_STEPS.DATA_PREP,
      nextAction: 'セットアップを確認する',
      description: 'セットアップ状況を確認してください'
    };
  }
}

/**
 * ユーザーの存在確認
 * @param {string} userId ユーザーID
 * @return {Object} 確認結果
 */
function verifyUserExists(userId) {
  try {
    console.log('verifyUserExists: 確認開始', { userId });
    
    if (!userId) {
      return { found: false, error: 'ユーザーIDが指定されていません' };
    }
    
    // ID検索とメール検索の両方で確認
    const userById = findUserById(userId);
    
    if (userById) {
      // さらにメールアドレスでの逆引き確認
      const userByEmail = findUserByEmailNonBlocking(userById.adminEmail);
      
      const verification = {
        found: true,
        userId: userById.userId,
        adminEmail: userById.adminEmail,
        consistency: userByEmail?.userId === userId,
        details: {
          foundById: !!userById,
          foundByEmail: !!userByEmail,
          idMatch: userByEmail?.userId === userId
        }
      };
      
      console.log('verifyUserExists: 確認完了', verification);
      return verification;
    } else {
      console.log('verifyUserExists: ユーザーが見つかりません', { userId });
      return { 
        found: false, 
        userId: userId,
        error: 'ユーザーが見つかりません' 
      };
    }
  } catch (error) {
    console.error('verifyUserExists error:', error);
    return {
      found: false,
      error: error.message,
      userId: userId
    };
  }
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
    return safeSetXFrameOptionsAllowAll(t.evaluate().setTitle('初回セットアップ - StudyQuest'));
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
      return safeSetXFrameOptionsAllowAll(errorHtml);
    }
    const appSetupTemplate = HtmlService.createTemplateFromFile('AppSetupPage');
    appSetupTemplate.include = include;
    
    // AppSetupPage用のuserIdを設定
    const currentUserEmail = Session.getActiveUser().getEmail();
    const currentUserInfo = findUserByEmail(currentUserEmail);
    appSetupTemplate.userId = currentUserInfo ? currentUserInfo.userId : '';
    
    console.log('AppSetupPage - currentUserEmail:', currentUserEmail);
    console.log('AppSetupPage - userId set to:', appSetupTemplate.userId);
    
    return safeSetXFrameOptionsAllowAll(appSetupTemplate.evaluate().setTitle('アプリ設定 - StudyQuest'));
  }

  if (params.setupParam === 'true') {
    const explicit = HtmlService.createTemplateFromFile('SetupPage');
    explicit.include = include;
    return safeSetXFrameOptionsAllowAll(explicit.evaluate().setTitle('StudyQuest - サービスアカウント セットアップ'));
  }

  // LoginPageページのリクエストを処理
  if (params.page === 'LoginPage') {
    return showLoginPage();
  }

  // システムセットアップが完了している場合のみ、userEmailをチェック
  if (!userEmail && !params.isDirectPageAccess) {
    return showLoginPage();
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
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.SAMEORIGIN)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
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
  const currentUserEmail = Session.getActiveUser().getEmail();
  const isOwner = currentUserEmail === userInfo.adminEmail;
  let hasDomainAccess = false;
  if (!isOwner) {
    try {
      AuthorizationService.verifyBoardAccess(userInfo.adminEmail);
      hasDomainAccess = true;
    } catch (e) {
      console.warn('Domain access denied:', e.message);
    }
  }
  const showBoard = isOwner || (isPublished && hasDomainAccess);
  const file = showBoard ? 'Page' : 'Unpublished';
  const template = HtmlService.createTemplateFromFile(file);
  template.include = include;

  if (showBoard) {
    try {
      // 🔧 レンダリング前に設定の整合性を検証・修復
      const configValidation = validateAndRepairUserConfig(userInfo.userId);
      if (configValidation.repaired) {
        console.log('🔧 [RENDER REPAIR] レンダリング前に設定が自動修復されました:', configValidation.repairs);
        // 修復後の最新情報を再取得
        const freshUserInfo = findUserById(userInfo.userId);
        if (freshUserInfo) {
          userInfo = freshUserInfo;
          config = JSON.parse(userInfo.configJson || '{}');
        }
      }
      
      if (userInfo.spreadsheetId) {
        try { addServiceAccountToSpreadsheet(userInfo.spreadsheetId); } catch (err) { console.warn('アクセス権設定警告:', err.message); }
      }
      template.userId = userInfo.userId;
      template.spreadsheetId = userInfo.spreadsheetId;
      template.ownerName = userInfo.adminEmail;
      template.sheetName = escapeJavaScript(config.publishedSheetName || params.sheetName);
      template.DEBUG_MODE = shouldEnableDebugMode();
      
      // 🔍 デバッグ: テンプレート変数解決の詳細をログ出力
      console.log('🔍 [TEMPLATE DEBUG] config.publishedSheetName:', config.publishedSheetName);
      console.log('🔍 [TEMPLATE DEBUG] sheetConfigKey:', sheetConfigKey);
      console.log('🔍 [TEMPLATE DEBUG] sheetConfig:', JSON.stringify(sheetConfig));
      console.log('🔍 [TEMPLATE DEBUG] sheetConfig.opinionHeader:', sheetConfig.opinionHeader);
      
      // 強化されたopinionHeader解決ロジック
      let rawOpinionHeader = sheetConfig.opinionHeader;
      
      // フォールバック1: publishedSheetNameを使用
      if (!rawOpinionHeader || rawOpinionHeader.trim() === '') {
        rawOpinionHeader = config.publishedSheetName;
        console.log('🔄 [FALLBACK] Using publishedSheetName as opinionHeader:', rawOpinionHeader);
      }
      
      // フォールバック2: 他のシート設定を探索
      if (!rawOpinionHeader || rawOpinionHeader.trim() === '') {
        const allSheetKeys = Object.keys(config).filter(key => key.startsWith('sheet_'));
        console.log('🔍 [FALLBACK] Searching in all sheet configs:', allSheetKeys);
        
        for (const key of allSheetKeys) {
          if (config[key] && config[key].opinionHeader) {
            rawOpinionHeader = config[key].opinionHeader;
            console.log('🔄 [FALLBACK] Found opinionHeader in', key, ':', rawOpinionHeader);
            break;
          }
        }
      }
      
      // フォールバック3: デフォルト値
      if (!rawOpinionHeader || rawOpinionHeader.trim() === '') {
        rawOpinionHeader = 'お題';
        console.log('🔄 [FALLBACK] Using default opinionHeader:', rawOpinionHeader);
      }
      
      console.log('🔍 [TEMPLATE DEBUG] rawOpinionHeader resolved to:', rawOpinionHeader);
      
      template.opinionHeader = escapeJavaScript(rawOpinionHeader);
      console.log('🔍 [TEMPLATE DEBUG] final template.opinionHeader:', template.opinionHeader);
      template.cacheTimestamp = Date.now();
      template.displayMode = config.displayMode || 'anonymous';
      template.showCounts = config.showCounts !== undefined ? config.showCounts : false;
      template.showScoreSort = template.showCounts;
      const isOwner = currentUserEmail === userInfo.adminEmail;
      template.showAdminFeatures = isOwner;
      template.isAdminUser = isOwner;
      template.formUrl = getFormUrlSafely(config, userInfo.spreadsheetId);
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
      const isOwner = currentUserEmail === userInfo.adminEmail;
      template.showAdminFeatures = isOwner;
      template.isAdminUser = isOwner;
      template.formUrl = '';
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
