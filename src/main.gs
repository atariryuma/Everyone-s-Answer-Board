/**
 * @fileoverview メインエントリーポイント - 2024年最新技術の結集
 * V8ランタイム、最新パフォーマンス技術、安定性強化を統合
 */

// グローバル定数の定義
var SCRIPT_PROPS_KEYS = {
  SERVICE_ACCOUNT_CREDS: 'SERVICE_ACCOUNT_CREDS',
  DATABASE_SPREADSHEET_ID: 'DATABASE_SPREADSHEET_ID',
  DEPLOY_DOMAIN: 'DEPLOY_DOMAIN'
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
 * 公開エントリーポイント（インテリジェントルーティング版）
 * システムの状態とユーザーの登録状況に応じて、適切な画面を返す
 */
function doGet(e) {
  try {
    // パラメータの安全な初期化
    e = e || {};
    e.parameter = e.parameter || {};
    
    var userId = e.parameter.userId || '';
    var requestMode = e.parameter.mode || '';
    var setup = e.parameter.setup || '';
    
    // 【チェック1】システムの初期設定が完了しているか？
    var dbSpreadsheetId = PropertiesService.getScriptProperties().getProperty('DATABASE_SPREADSHEET_ID');
    
    if (!dbSpreadsheetId) {
      // データベースIDがなければ、セットアップ画面を返す
      console.log('システムが未設定です。セットアップ画面にリダイレクトします。');
      return HtmlService.createTemplateFromFile('SetupPage')
        .evaluate()
        .setTitle('初回セットアップ - StudyQuest')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
    }

    // セットアップページの明示的な表示要求
    if (setup === 'true') {
      return HtmlService.createTemplateFromFile('SetupPage')
        .evaluate()
        .setTitle('StudyQuest - サービスアカウント セットアップ')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
    }

    // 【チェック2】ユーザーが登録済みか？
    var userEmail = Session.getActiveUser().getEmail();
    
    if (!isUserRegistered(userEmail, dbSpreadsheetId)) {
      // ユーザーがデータベースに登録されていなければ、新規登録画面を返す
      console.log(`ユーザー(${userEmail})は未登録です。登録画面にリダイレクトします。`);
      return HtmlService.createTemplateFromFile('Registration')
        .evaluate()
        .setTitle('新規ユーザー登録 - StudyQuest')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
    }

    // 【チェック3】登録済みユーザーには適切な画面を表示
    console.log(`登録済みユーザー(${userEmail})です。`);
    
    // アカウント切り替え後のリダイレクト処理
    // パラメータにuserIdがない場合は、管理パネルにリダイレクト
    if (!userId) {
      var userInfo = findUserByEmail(userEmail);
      if (userInfo) {
        console.log('登録済みユーザーを管理パネルにリダイレクトします。');
        var currentUrl = ScriptApp.getService().getUrl();
        return HtmlService.createHtmlOutput(`
          <script>
            window.location.href = '${currentUrl}?userId=${encodeURIComponent(userInfo.userId)}&mode=admin';
          </script>
          <p>管理パネルにリダイレクト中...</p>
        `).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
      }
    }
    
    // ユーザー情報を取得（userIdが指定されている場合はそれを使用、なければemailから検索）
    var userInfo;
    if (userId) {
      userInfo = findUserById(userId);
      if (!userInfo) {
        console.log('指定されたuserIdが無効です。emailから検索します。');
        userInfo = findUserByEmail(userEmail);
      }
    } else {
      userInfo = findUserByEmail(userEmail);
      if (userInfo) {
        userId = userInfo.userId;
      }
    }

    if (!userInfo) {
      console.error('ユーザー情報の取得に失敗しました。');
      return HtmlService.createHtmlOutput(
        '<h2>エラー</h2>' +
        '<p>ユーザー情報が見つかりません。システム管理者に連絡してください。</p>'
      ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
    }
    
    // ユーザーの最終アクセス日時を更新（非同期）
    try {
      updateUser(userId, { lastAccessedAt: new Date().toISOString() });
    } catch (updateError) {
      console.error('最終アクセス日時の更新に失敗: ' + updateError.message);
    }

    // ユーザー情報をプロパティに保存（リアクション機能で使用）
    PropertiesService.getUserProperties().setProperty('CURRENT_USER_ID', userId);

    // 表示モードの決定：管理者は管理パネル、それ以外は回答ボード
    if (requestMode === 'admin') {
      var template = HtmlService.createTemplateFromFile('AdminPanel');
      template.userInfo = userInfo;
      template.userId = userId;
      template.mode = requestMode;
      template.displayMode = 'named';
      template.showAdminFeatures = true;
      return template.evaluate()
        .setTitle('管理パネル - みんなの回答ボード')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
    } else {
      // デフォルトは回答ボード表示
      var template = HtmlService.createTemplateFromFile('Page');
      template.userInfo = userInfo;
      template.userId = userId;
      template.mode = requestMode;
      template.displayMode = 'anonymous';
      template.showAdminFeatures = false;
      return template.evaluate()
        .setTitle('みんなの回答ボード')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
    }
    
  } catch (error) {
    console.error('doGetで致命的なエラーが発生しました: ' + error.message);
    console.error('Stack trace:', error.stack);
    
    return HtmlService.createHtmlOutput(
      '<h1>エラー</h1>' +
      '<p>アプリケーションの起動に失敗しました。</p>' +
      '<p>エラー詳細: ' + htmlEncode(error.message) + '</p>' +
      '<p>しばらく待ってから再度お試しください。</p>'
    ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
  }
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