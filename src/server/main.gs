/**
 * @fileoverview メインエントリーポイント - 新しいファイル構成対応版
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

/**
 * 他のHTMLファイル（CSS, JS, コンポーネント）をテンプレートに埋め込むためのヘルパー関数
 * @param {string} path - client/ からの相対パス (例: 'styles/main.css.html')
 * @returns {string} ファイルのコンテンツ
 */
function include(path) {
  return HtmlService.createHtmlOutputFromFile('client/' + path).getContent();
}

/**
 * 指定されたビュー名でページをレンダリングする
 * @param {string} viewName - /src/client/views/ 内のファイル名 (拡張子なし)
 * @param {object} data - テンプレートに渡すデータ
 * @returns {HtmlOutput} 評価されたHTMLテンプレート
 */
function renderPage(viewName, data = {}) {
  const template = HtmlService.createTemplateFromFile(`client/views/${viewName}`);
  // Pass data to the template
  for (const key in data) {
    template[key] = data[key];
  }
  template.include = include;
  const output = template.evaluate()
    .setTitle(data.title || 'みんなの回答ボード') // Use title from data or default
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return safeSetXFrameOptionsDeny(output);
}

/**
 * Webアプリケーションのメインエントリーポイント
 * @param {Object} e - URLパラメータを含むイベントオブジェクト
 * @returns {HtmlOutput} 表示するHTMLコンテンツ
 */
function doGet(e) {
  console.log(`doGet called with event object: ${JSON.stringify(e)}`);

  try {
    const mode = (e && e.parameter && e.parameter.mode) ? e.parameter.mode : 'admin';
    const userId = (e && e.parameter && e.parameter.userId) ? e.parameter.userId : null;
    const setupParam = (e && e.parameter && e.parameter.setup) ? e.parameter.setup : null;
    const spreadsheetId = (e && e.parameter && e.parameter.spreadsheetId) ? e.parameter.spreadsheetId : null;
    const sheetName = (e && e.parameter && e.parameter.sheetName) ? e.parameter.sheetName : null;
    
    const isDirectPageAccess = userId && spreadsheetId && sheetName;

    console.log(`Request parameters validated. Mode: ${mode}, UserID: ${userId}, SetupParam: ${setupParam}, IsDirectPageAccess: ${isDirectPageAccess}`);

    const userEmail = Session.getActiveUser().getEmail();
    console.log(`Access by user: ${userEmail}`);

    const webAppUrl = ScriptApp.getService().getUrl();
    console.log(`Current Web App URL: ${webAppUrl}`);

    if (e && e.headers && e.headers.referer) {
      console.log(`Referer: ${e.headers.referer}`);
    }

    // 1. システムの初期セットアップが完了しているか確認
    if (!isSystemSetup() && !isDirectPageAccess) {
      console.log('DEBUG: System not set up. Redirecting to SetupPage.');
      return renderPage('SetupPage', { title: '初回セットアップ - StudyQuest' });
    }

    // セットアップページの明示的な表示要求
    if (setupParam === 'true') {
      console.log('DEBUG: Explicit setup request. Redirecting to SetupPage.');
      return renderPage('SetupPage', { title: 'StudyQuest - サービスアカウント セットアップ' });
    }

    // 2. ユーザー認証と情報取得
    if (!userEmail && !isDirectPageAccess) {
      console.log('DEBUG: No current user email. Redirecting to RegistrationPage.');
      console.log('DEBUG: Serving RegistrationPage HTML');
      return renderPage('Registration', { title: '新規ユーザー登録 - StudyQuest' });
    }

    // 3. ユーザーがデータベースに登録済みか確認
    let userInfo = null;
    if (isDirectPageAccess) {
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
        const pageData = {
          userId: userInfo.userId,
          spreadsheetId: userInfo.spreadsheetId,
          ownerName: userInfo.adminEmail,
          sheetName: configJson.publishedSheetName || sheetName,
          opinionHeader: (configJson[`sheet_${configJson.publishedSheetName || sheetName}`] && configJson[`sheet_${configJson.publishedSheetName || sheetName}`].opinionHeader) || configJson.publishedSheetName || 'お題',
          displayMode: configJson.displayMode || 'anonymous',
          showCounts: configJson.showCounts !== undefined ? configJson.showCounts : true,
          showAdminFeatures: false,
          isAdminUser: false,
          title: 'StudyQuest -みんなの回答ボード-'
        };
        return renderPage('Page', pageData);
      }

      if (mode === 'admin') {
        console.log('DEBUG: Explicit admin mode request. Showing AdminPanel.');
        const adminData = {
          userInfo: userInfo,
          userId: userInfo.userId,
          mode: mode,
          displayMode: 'named',
          showAdminFeatures: true,
          title: '管理パネル - みんなの回答ボード'
        };
        console.log('DEBUG: AdminPanel template properties:', {
          userInfo: adminData.userInfo,
          userId: adminData.userId,
          mode: adminData.mode,
          displayMode: adminData.displayMode,
          showAdminFeatures: adminData.showAdminFeatures
        });
        return renderPage('AdminPanel', adminData);
      }
      
      if (mode === 'view' && isPublished) {
        console.log('DEBUG: Explicit view mode request for published board. Showing Page.');
        const pageData = {
          userId: userInfo.userId,
          spreadsheetId: userInfo.spreadsheetId,
          ownerName: userInfo.adminEmail,
          sheetName: configJson.publishedSheetName || sheetName,
          opinionHeader: (configJson[`sheet_${configJson.publishedSheetName || sheetName}`] && configJson[`sheet_${configJson.publishedSheetName || sheetName}`].opinionHeader) || configJson.publishedSheetName || 'お題',
          displayMode: configJson.displayMode || 'anonymous',
          showCounts: configJson.showCounts !== undefined ? configJson.showCounts : true,
          showAdminFeatures: false,
          isAdminUser: false,
          title: 'StudyQuest -みんなの回答ボード-'
        };
        return renderPage('Page', pageData);
      }

      // デフォルト：パラメータなしのアクセスは管理パネルへ
      console.log('DEBUG: Default access - user registered. Redirecting to admin panel.');
      const adminData = {
        userInfo: userInfo,
        userId: userInfo.userId,
        mode: 'admin',
        displayMode: 'named',
        showAdminFeatures: true,
        title: '管理パネル - みんなの回答ボード'
      };
      console.log('DEBUG: Serving default AdminPanel HTML');
      return renderPage('AdminPanel', adminData);

    } else {
      // 5. 【未登録ユーザーの処理】
      console.log('DEBUG: User not registered. Redirecting to RegistrationPage.');
      console.log('DEBUG: Serving RegistrationPage HTML for unregistered user');
      return renderPage('Registration', { title: '新規ユーザー登録 - StudyQuest' });
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
 * ユーザーがデータベースに登録済みかを確認するヘルパー関数
 */
function isUserRegistered(email, spreadsheetId) {
  try {
    if (!email || !spreadsheetId) {
      return false;
    }
    
    var userInfo = findUserByEmail(email);
    return !!userInfo;
    
  } catch(e) {
    console.error(`ユーザー登録確認中にエラー: ${e.message}`);
    return false;
  }
}

/**
 * システムの初期セットアップが完了しているかを確認するヘルパー関数
 */
function isSystemSetup() {
  var dbSpreadsheetId = PropertiesService.getScriptProperties().getProperty('DATABASE_SPREADSHEET_ID');
  return !!dbSpreadsheetId;
}