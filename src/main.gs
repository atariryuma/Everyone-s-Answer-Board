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
  // ① パラメータe全体をログに出力して、どのようなリクエストか確認する
  console.log(`doGet called with event object: ${JSON.stringify(e)}`);

  try {
    // ★★★ 修正点 ★★★
    // パラメータがない場合のデフォルトモードを 'view' から 'admin' に変更。
    // これにより、登録済みユーザーはデフォルトで管理パネルにアクセスするようになります。
    const mode = (e && e.parameter && e.parameter.mode) ? e.parameter.mode : 'admin';
    const userId = (e && e.parameter && e.parameter.userId) ? e.parameter.userId : null;
    const setupParam = (e && e.parameter && e.parameter.setup) ? e.parameter.setup : null;
    const spreadsheetId = (e && e.parameter && e.parameter.spreadsheetId) ? e.parameter.spreadsheetId : null;
    const sheetName = (e && e.parameter && e.parameter.sheetName) ? e.parameter.sheetName : null;

    // Page.html への直接アクセス判定
    // userId が指定されて mode=view の場合は、登録状態に関わらず回答ボードを表示する
    const isDirectPageAccess = userId && mode === 'view';

    console.log(`Request parameters validated. Mode: ${mode}, UserID: ${userId}, SetupParam: ${setupParam}, IsDirectPageAccess: ${isDirectPageAccess}`);

    // どこからのアクセスかを知るためのログを追加
    const userEmail = Session.getActiveUser().getEmail();
    console.log(`Access by user: ${userEmail}`);
    
    // セッション分離の強化：アカウント切り替え検出と整合性保証
    if (userEmail && !isDirectPageAccess) {
      try {
        // アカウント切り替えを検出
        const switchInfo = detectAccountSwitch(userEmail);
        if (switchInfo.isAccountSwitch) {
          console.log('アカウント切り替えを検出しました: ' + switchInfo.previousEmail + ' → ' + switchInfo.currentEmail);
        }
        
        // セッション整合性を検証・修復
        const sessionValid = validateAndRepairSession(userEmail);
        if (!sessionValid) {
          console.warn('セッション整合性の修復に失敗しました。強制初期化を実行します。');
          forceSessionReset(userEmail);
        }
        
      } catch (sessionError) {
        console.error('セッション管理でエラーが発生しました: ' + sessionError.message);
        // 重大なエラーの場合は強制初期化
        try {
          forceSessionReset(userEmail);
        } catch (resetError) {
          console.error('強制セッション初期化も失敗しました: ' + resetError.message);
        }
      }
    }

    const webAppUrl = ScriptApp.getService().getUrl();
    console.log(`Current Web App URL: ${webAppUrl}`);

    if (e && e.headers && e.headers.referer) {
      console.log(`Referer: ${e.headers.referer}`);
    }

    // 1. システムの初期セットアップが完了しているか確認（Page.html直接アクセス時は除く）
    if (!isSystemSetup() && !isDirectPageAccess) {
      console.log('DEBUG: System not set up. Redirecting to SetupPage.');
      var setupTemplate = HtmlService.createTemplateFromFile('SetupPage');
      setupTemplate.include = include;
      var setupHtml = setupTemplate.evaluate()
        .setTitle('初回セットアップ - StudyQuest');
      console.log('DEBUG: Serving SetupPage HTML');
      return safeSetXFrameOptionsDeny(setupHtml);
    }

    // アプリ設定ページの表示要求（?setup=true&mode=appsetup）
    if (setupParam === 'true' && mode === 'appsetup') {
      console.log('DEBUG: App setup page request. Checking access permissions.');
      
      // アクセス権限を確認
      if (!hasSetupPageAccess()) {
        console.log('DEBUG: Access denied to app setup page.');
        var errorHtml = HtmlService.createHtmlOutput(
          '<h1>アクセス拒否</h1>' +
          '<p>アプリ設定ページにアクセスする権限がありません。</p>' +
          '<p>編集者として登録され、かつアクティブ状態である必要があります。</p>'
        );
        return safeSetXFrameOptionsDeny(errorHtml);
      }
      
      console.log('DEBUG: Serving AppSetupPage HTML');
      var appSetupTemplate = HtmlService.createTemplateFromFile('AppSetupPage');
      appSetupTemplate.include = include;
      var appSetupHtml = appSetupTemplate.evaluate()
        .setTitle('アプリ設定 - StudyQuest');
      return safeSetXFrameOptionsDeny(appSetupHtml);
    }

    // セットアップページの明示的な表示要求
    if (setupParam === 'true') {
      console.log('DEBUG: Explicit setup request. Redirecting to SetupPage.');
      var explicitTemplate = HtmlService.createTemplateFromFile('SetupPage');
      explicitTemplate.include = include;
      var explicitHtml = explicitTemplate.evaluate()
        .setTitle('StudyQuest - サービスアカウント セットアップ');
      console.log('DEBUG: Serving explicit SetupPage HTML');
      return safeSetXFrameOptionsDeny(explicitHtml);
    }

    // 2. ユーザー認証と情報取得（Page.html直接アクセス時は除く）
    if (!userEmail && !isDirectPageAccess) {
      console.log('DEBUG: No current user email. Redirecting to RegistrationPage.');
      console.log('DEBUG: Serving RegistrationPage HTML');
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
        console.log('DEBUG: Parsed configJson:', JSON.stringify(configJson));
      } catch (jsonErr) {
        console.warn('Invalid configJson:', jsonErr.message);
      }

      var isPublished = !!(configJson.appPublished &&
        configJson.publishedSpreadsheetId &&
        configJson.publishedSheetName);
      console.log(`DEBUG: isPublished = ${isPublished}`);

      // Page.html直接アクセス時は、パラメータ指定されたページを表示
      if (isDirectPageAccess) {
        console.log('DEBUG: Direct page access detected. Showing specified page.');
        const template = HtmlService.createTemplateFromFile('Page');
        template.include = include;

        try {
          const config = JSON.parse(userInfo.configJson || '{}');
          const sheetConfigKey = 'sheet_' + (config.publishedSheetName || sheetName);
          const sheetConfig = config[sheetConfigKey] || {};
          
          template.userId = userInfo.userId;
          template.spreadsheetId = userInfo.spreadsheetId;
          template.ownerName = userInfo.adminEmail;
          template.sheetName = escapeJavaScript(config.publishedSheetName || sheetName);
          const rawOpinionHeader = sheetConfig.opinionHeader || config.publishedSheetName || 'お題';
          
          // 直接rawOpinionHeaderを使用（Base64エンコード削除）
          template.opinionHeader = escapeJavaScript(rawOpinionHeader);
          template.cacheTimestamp = Date.now(); // キャッシュバスター
          
          template.displayMode = config.displayMode || 'anonymous';
          template.showCounts = config.showCounts !== undefined ? config.showCounts : false;
          
          // 現在のユーザーがボードの所有者かどうかをチェック
          const currentUserEmail = Session.getActiveUser().getEmail();
          const isOwner = currentUserEmail === userInfo.adminEmail;
          template.showAdminFeatures = isOwner; // 所有者のみに管理機能を提供
          template.isAdminUser = isOwner; // 所有者のみに管理者権限を付与
          
          console.log('DEBUG: Owner check for direct page access:', {
            currentUserEmail: currentUserEmail,
            ownerEmail: userInfo.adminEmail,
            isOwner: isOwner
          });
          
          // デバッグログ
          console.log('Template variables for direct page access:', {
            sheetName: template.sheetName,
            opinionHeader: template.opinionHeader,
            rawOpinionHeader: rawOpinionHeader,
            displayMode: template.displayMode,
            cacheTimestamp: template.cacheTimestamp
          });

        } catch (e) {
          template.opinionHeader = escapeJavaScript('お題の読込エラー');
          template.cacheTimestamp = Date.now();
          template.userId = userInfo.userId;
          template.spreadsheetId = userInfo.spreadsheetId;
          template.ownerName = userInfo.adminEmail;
          template.sheetName = escapeJavaScript(sheetName);
          template.displayMode = 'anonymous';
          template.showCounts = false;
          
          // エラー時も所有者チェックを実行
          const currentUserEmail = Session.getActiveUser().getEmail();
          const isOwner = currentUserEmail === userInfo.adminEmail;
          template.showAdminFeatures = isOwner;
          template.isAdminUser = isOwner;
        }
        
        return template.evaluate()
            .setTitle('StudyQuest -みんなの回答ボード-')
            .addMetaTag('viewport', 'width=device-width, initial-scale=1');
      }

      // 要件に従ったリダイレクトロジック：
      // - データベースに名前があれば管理パネル
      // - mode=adminまたはmode=viewで明示的に指定された場合はそれに従う
      
      if (mode === 'admin') {
        // 明示的な管理パネル要求
        console.log('DEBUG: Explicit admin mode request. Showing AdminPanel.');
        var adminTemplate = HtmlService.createTemplateFromFile('AdminPanel');
        adminTemplate.include = include;
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
      .setTitle('みんなの回答ボード 管理パネル')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
        console.log('DEBUG: Serving AdminPanel HTML');
        return safeSetXFrameOptionsDeny(adminHtml);
      }
      
      if (mode === 'view') {
        if (isPublished) {
          // 明示的な回答ボード表示要求（公開済みの場合）
          console.log('DEBUG: Explicit view mode request for published board. Showing Page.');
          const template = HtmlService.createTemplateFromFile('Page');
          template.include = include;
          
          try {
            const config = JSON.parse(userInfo.configJson || '{}');
            const sheetConfigKey = 'sheet_' + (config.publishedSheetName || sheetName);
            const sheetConfig = config[sheetConfigKey] || {};
            
            template.userId = userInfo.userId;
            template.spreadsheetId = userInfo.spreadsheetId;
            template.ownerName = userInfo.adminEmail;
            template.sheetName = escapeJavaScript(config.publishedSheetName || sheetName);
            const rawOpinionHeader = sheetConfig.opinionHeader || config.publishedSheetName || 'お題';
            
            // 直接rawOpinionHeaderを使用（Base64エンコード削除）
            template.opinionHeader = escapeJavaScript(rawOpinionHeader);
            template.cacheTimestamp = Date.now(); // キャッシュバスター
            
            template.displayMode = config.displayMode || 'anonymous';
            template.showCounts = config.showCounts !== undefined ? config.showCounts : false;
            
            // 現在のユーザーがボードの所有者かどうかをチェック
            const currentUserEmail = Session.getActiveUser().getEmail();
            const isOwner = currentUserEmail === userInfo.adminEmail;
            template.showAdminFeatures = isOwner; // 所有者のみに管理機能を提供
            template.isAdminUser = isOwner; // 所有者のみに管理者権限を付与

          } catch (e) {
            template.opinionHeader = escapeJavaScript('お題の読込エラー');
            template.cacheTimestamp = Date.now();
            template.userId = userInfo.userId;
            template.spreadsheetId = userInfo.spreadsheetId;
            template.ownerName = userInfo.adminEmail;
            template.sheetName = escapeJavaScript(sheetName);
            template.displayMode = 'anonymous';
            template.showCounts = false;
            
            // エラー時も所有者チェックを実行
            const currentUserEmail = Session.getActiveUser().getEmail();
            const isOwner = currentUserEmail === userInfo.adminEmail;
            template.showAdminFeatures = isOwner;
            template.isAdminUser = isOwner;
          }
          
          return template.evaluate()
              .setTitle('StudyQuest -みんなの回答ボード-')
              .addMetaTag('viewport', 'width=device-width, initial-scale=1');
        } else {
          // 未公開状態での表示要求 - Unpublished.htmlを表示
          console.log('DEBUG: View mode request for unpublished board. Showing Unpublished.');
          const template = HtmlService.createTemplateFromFile('Unpublished');
          template.include = include;
          
          try {
            const config = JSON.parse(userInfo.configJson || '{}');
            
            template.userId = userInfo.userId;
            template.spreadsheetId = userInfo.spreadsheetId;
            template.ownerName = userInfo.adminEmail;
            template.isOwner = true; // 管理者であることを示す
            template.adminEmail = userInfo.adminEmail;
            template.cacheTimestamp = Date.now();
            
            // 管理パネルと回答ボードのURLを設定
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

      // デフォルト：パラメータなしのアクセスは管理パネルへ
      // (modeのデフォルトが'admin'になったため、このルートが正しく機能する)
      console.log('DEBUG: Default access - user registered. Redirecting to admin panel.');
      var adminTemplate = HtmlService.createTemplateFromFile('AdminPanel');
      adminTemplate.include = include;
      adminTemplate.userInfo = userInfo;
      adminTemplate.userId = userInfo.userId;
      adminTemplate.mode = 'admin';
      adminTemplate.displayMode = 'named';
      adminTemplate.showAdminFeatures = true;
      var adminHtml = adminTemplate.evaluate()
        .setTitle('管理パネル - みんなの回答ボード');
      console.log('DEBUG: Serving default AdminPanel HTML');
      return safeSetXFrameOptionsDeny(adminHtml);

    } else {
      // 5. 【未登録ユーザーの処理】
      //    Page.html直接アクセス時でもユーザーが見つからない場合は登録ページ
      console.log('DEBUG: User not registered. Redirecting to RegistrationPage.');
      console.log('DEBUG: Serving RegistrationPage HTML for unregistered user');
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
