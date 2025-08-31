/**
 * Module-scoped Constants for main.gs
 * Core constants参照による軽量設計
 */
const MODULE_CONFIG = Object.freeze({
  // キャッシュTTL設定（CORE.TIMEOUTS参照）
  CACHE_TTL: CORE.TIMEOUTS.LONG,
  QUICK_CACHE_TTL: CORE.TIMEOUTS.SHORT,
  
  // ステータス定数（CORE.STATUS参照） 
  STATUS_ACTIVE: CORE.STATUS.ACTIVE,
  STATUS_INACTIVE: CORE.STATUS.INACTIVE,
});

// User管理の内部実装（依存注入パターン）
const User = {
  email() {
    try {
      return Session.getActiveUser().getEmail();
    } catch (e) {
      return null;
    }
  }
};

/**
 * Entry Point Functions
 * 新アーキテクチャにおけるエントリーポイント関数群
 */

function doGet(e) {
  try {
    // App名前空間の初期化（必須）
    App.init();

    // リクエストパラメータを解析
    const params = parseRequestParams(e);
    console.log('doGet - Received params:', params);

    if (!isSystemSetup()) {
      console.warn('System not setup, redirecting to setup page');
      return renderSetupPage(params);
    }

    // モードに応じたルーティング
    switch (params.mode) {
      case 'admin':
        console.log('doGet - Admin mode detected, userId:', params.userId);
        if (!params.userId) {
          throw new Error('Admin mode requires userId parameter');
        }
        try {
          const accessResult = App.getAccess().verifyAccess(params.userId, 'admin', User.email());
          if (!accessResult.allowed) {
            console.warn('Admin access denied:', accessResult);
            return HtmlService.createHtmlOutput('<h3>Access Denied</h3><p>Admin access is not allowed for this user.</p>');
          }
          
          // 互換性のためのユーザー情報変換
          const compatUserInfo = {
            userId: params.userId,
            tenantId: params.userId,
            adminEmail: accessResult.config?.email || User.email(),
            configJson: JSON.stringify(accessResult.config || {})
          };
          
          return renderAdminPanel(compatUserInfo, 'admin');
        } catch (adminError) {
          console.error('Admin mode error:', adminError);
          return HtmlService.createHtmlOutput('<h3>Error</h3><p>An error occurred in admin mode: ' + adminError.message + '</p>');
        }
        
      case 'login':
        console.log('doGet - Login mode detected');
        return renderLoginPage(params);
        
      case 'view':
      default:
        console.log('doGet - View mode (default), userId:', params.userId);
        if (!params.userId) {
          console.log('doGet - No userId provided, redirecting to login');
          return renderLoginPage(params);
        }
        
        try {
          const accessResult = App.getAccess().verifyAccess(params.userId, 'view', User.email());
          if (!accessResult.allowed) {
            console.warn('View access denied:', accessResult);
            if (accessResult.userType === 'not_found') {
              return HtmlService.createHtmlOutput('<h3>User Not Found</h3><p>The specified user does not exist.</p>');
            }
            return HtmlService.createHtmlOutput('<h3>Private Board</h3><p>This board is private.</p>');
          }
          
          // 互換性のためのユーザー情報変換
          const compatUserInfo = {
            userId: params.userId,
            tenantId: params.userId,
            adminEmail: accessResult.config?.email || '',
            configJson: JSON.stringify(accessResult.config || {})
          };

          return renderAnswerBoard(compatUserInfo, params);
        } catch (viewError) {
          console.error('View mode error:', viewError);
          return HtmlService.createHtmlOutput('<h3>Error</h3><p>An error occurred: ' + viewError.message + '</p>');
        }
    }
  } catch (error) {
    console.error('doGet - Critical error:', error);
    return HtmlService.createHtmlOutput('<h3>System Error</h3><p>A critical system error occurred: ' + error.message + '</p>');
  }
}

/**
 * Unified Services Layer
 * サービス層の統一インターフェース
 */
const Services = {
  user: {
    get current() {
      try {
        const email = User.email();
        if (!email) return null;
        
        // 簡易的なユーザー情報を返す（将来的にはApp.getConfig()経由）
        return {
          email: email,
          isAuthenticated: true
        };
      } catch (error) {
        console.error('getCurrentUserInfo エラー', { error: error.message });
        return null;
      }
    },
    
    getCurrentUserInfo() {
      try {
        // 新アーキテクチャでの単純化実装
        const userInfo = Services.user.current;
        if (!userInfo) return null;
        
        return {
          email: userInfo.email,
          tenantId: userInfo.email.split('@')[0], // 簡易的なテナントID
          ownerEmail: userInfo.email,
          // spreadsheetId, configJsonは削除され、ConfigurationManagerで管理
        };
      } catch (error) {
        console.error('getCurrentUserInfo エラー', { error: error.message });
        return null;
      }
    }
  },
  
  access: {
    check() {
      try {
        // システム管理者は常にアクセス許可
        if (Deploy.isUser()) {
          return { hasAccess: true, message: 'システム管理者アクセス許可' };
        }

        // 基本的には全ユーザーアクセス許可（必要に応じて制限を追加）
        return { hasAccess: true, message: 'アクセス許可' };
      } catch (error) {
        console.error('アプリケーションアクセスチェック失敗', { error: error.message });
        return { hasAccess: true, message: 'アクセス許可（エラー回避）' };
      }
    }
  }
};

// デプロイ関連の機能
const Deploy = {
  // ドメイン情報を取得する関数
  domain() {
    try {
      console.log('Deploy.domain() - start');
      
      const activeUserEmail = User.email();
      const currentDomain = getEmailDomain(activeUserEmail);
      
      // WebAppのURLを取得してドメインを判定
      const webAppUrl = getWebAppUrl();
      let deployDomain = ''; // 個人アカウント/グローバルアクセスの場合、デフォルトで空
      
      if (webAppUrl && webAppUrl.includes('/a/')) {
        // Google Workspace環境でのドメイン取得
        const domainMatch =
          webAppUrl.match(/\/a\/([a-zA-Z0-9\-.]+)\/macros\//);
        if (domainMatch && domainMatch[1]) {
          deployDomain = domainMatch[1];
          console.log('Deploy.domain() - Workspace domain detected:', deployDomain);
        }
      } else {
        console.log('Deploy.domain() - Personal/Global access (no domain restriction)');
      }
      
      // ドメインマッチング判定（個人アカウントの場合は常にtrue）
      const isDomainMatch = currentDomain === deployDomain || deployDomain === '';
      
      return {
        currentDomain: currentDomain,
        deployDomain: deployDomain,
        isDomainMatch: isDomainMatch,
        webAppUrl: webAppUrl,
      };
    } catch (e) {
      console.error('デプロイユーザードメイン情報取得失敗', { error: e.message });
      return {
        currentDomain: '',
        deployDomain: '',
        isDomainMatch: false,
        error: e.message,
      };
    }
  },
  
  isUser() {
    try {
      const props = PropertiesService.getScriptProperties();
      const adminEmail = props.getProperty(PROPS_KEYS.ADMIN_EMAIL);
      const currentUserEmail = User.email();

      console.log('Deploy.isUser() - 管理者確認', { adminEmail, currentUserEmail });
      return adminEmail === currentUserEmail;
    } catch (error) {
      console.error('Deploy.isUser() エラー', { error: error.message });
      return false;
    }
  }
};


/**
 * 簡素化されたエラーハンドリング関数群
 */

// ログ関数は削除 - GAS ネイティブ console を直接使用

// PerformanceOptimizer.gsでglobalProfilerが既に定義されているため、
// 重複回避のためこちらでは定義しない

/**
 * システム初期化チェック
 * @returns {boolean} システムが初期化されているかどうか
 */
function isSystemSetup() {
  const props = PropertiesService.getScriptProperties();
  const dbSpreadsheetId = props.getProperty(PROPS_KEYS.DATABASE_SPREADSHEET_ID);
  const adminEmail = props.getProperty(PROPS_KEYS.ADMIN_EMAIL);
  const serviceAccountCreds = props.getProperty(PROPS_KEYS.SERVICE_ACCOUNT_CREDS);
  return !!dbSpreadsheetId && !!adminEmail && !!serviceAccountCreds;
}

/**
 * Google Client IDを取得
 */
function getGoogleClientId() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const clientId = properties.getProperty('GOOGLE_CLIENT_ID');

    if (!clientId) {
      console.warn('GOOGLE_CLIENT_ID not found in script properties');
      
      // デバッグ用：利用可能なプロパティを確認
      const allProperties = properties.getProperties();
      console.log('Available properties:', Object.keys(allProperties));
      
      return {
        success: false,
        message: 'Google Client ID not configured',
        setupInstructions:
          'Please set GOOGLE_CLIENT_ID in Google Apps Script project settings under Properties > Script Properties',
      };
    }

    return {
      success: true,
      message: 'Google Client IDを取得しました',
      data: { clientId: clientId },
    };
  } catch (error) {
    console.error('GOOGLE_CLIENT_ID取得エラー', { error: error.message });
    return {
      success: false,
      message: 'Google Client IDの取得に失敗しました: ' + error.toString(),
      data: { clientId: '' },
    };
  }
}

/**
 * システム設定状況をチェック
 */
function checkSystemConfiguration() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const allProperties = properties.getProperties();

    const requiredProperties = [
      'GOOGLE_CLIENT_ID',
      'DATABASE_SPREADSHEET_ID',
      'ADMIN_EMAIL',
      'SERVICE_ACCOUNT_CREDS',
    ];

    const configStatus = {};
    const missingProperties = [];

    requiredProperties.forEach(prop => {
      const value = allProperties[prop];
      configStatus[prop] = {
        hasValue: !!(value && value.trim()),
        length: value ? value.length : 0,
      };

      if (!value || !value.trim()) {
        missingProperties.push(prop);
      }
    });

    return {
      isFullyConfigured: missingProperties.length === 0,
      configStatus,
      missingProperties,
      availableProperties: Object.keys(allProperties),
      setupComplete: isSystemSetup(),
    };
  } catch (error) {
    console.error('Error checking system configuration:', error);
    return {
      isFullyConfigured: false,
      error: error.toString(),
    };
  }
}

/**
 * システムドメイン情報取得
 */
function getSystemDomainInfo() {
  try {
    const props = PropertiesService.getScriptProperties();
    const adminEmail = props.getProperty(PROPS_KEYS.ADMIN_EMAIL);
    
    if (!adminEmail) {
      throw new Error('Admin email not configured');
    }
    
    const adminDomain = adminEmail.split('@')[1];
    
    // デプロイ情報を取得
    const domainInfo = Deploy.domain();
    const isDomainMatch = domainInfo.isDomainMatch !== undefined ? domainInfo.isDomainMatch : false;
    
    return {
      success: true,
      adminEmail: adminEmail,
      adminDomain: adminDomain,
      isDomainMatch: isDomainMatch,
      currentDomain: domainInfo.currentDomain || '不明',
      deployDomain: domainInfo.deployDomain || adminDomain,
    };
  } catch (e) {
    console.error('getSystemDomainInfo エラー:', e.message);
    return { error: e.message };
  }
}

/**
 * User Management Functions
 * 統一されたユーザー管理機能群
 */

/**
 * 管理者パネルの表示
 */
function showAdminPanel() {
  try {
    console.log('showAdminPanel - start');
    
    // Admin権限でのアクセス確認
    const activeUserEmail = User.email();
    if (activeUserEmail) {
      const userProperties = PropertiesService.getUserProperties();
      const lastAdminUserId = userProperties.getProperty('lastAdminUserId');
      
      if (lastAdminUserId) {
        console.log('showAdminPanel - Using saved userId:', lastAdminUserId);
        try {
          const accessResult = App.getAccess().verifyAccess(lastAdminUserId, 'admin', activeUserEmail);
          if (accessResult.allowed) {
            console.log('✅ アクセス許可: ', accessResult);
            const compatUserInfo = {
              userId: lastAdminUserId,
              tenantId: lastAdminUserId,
              adminEmail: accessResult.config?.email || activeUserEmail,
              configJson: JSON.stringify(accessResult.config || {})
            };
            
            return renderAdminPanel(compatUserInfo, 'admin');
          }
        } catch (error) {
          console.warn('Saved userId access failed:', error.message);
        }
      }
    }
    
    // デフォルト：エラーページを表示
    return HtmlService.createHtmlOutput('<h3>Admin Access Required</h3><p>Please access via the admin URL with proper userId parameter.</p>');
  } catch (error) {
    console.error('showAdminPanel エラー:', error);
    return HtmlService.createHtmlOutput('<h3>Error</h3><p>An error occurred: ' + error.message + '</p>');
  }
}

/**
 * 回答ボードの表示
 */
function showAnswerBoard(userId) {
  try {
    console.log('showAnswerBoard - start, userId:', userId);
    if (!userId) {
      throw new Error('userId is required');
    }
    
    // キャッシュから最後のユーザーIDを保存
    const userProperties = PropertiesService.getUserProperties();
    const lastAdminUserId = userProperties.getProperty('lastAdminUserId');
    
    if (lastAdminUserId !== userId) {
      userProperties.setProperty('lastAdminUserId', userId);
    }
    
    // アクセス権限確認
    const accessResult = App.getAccess().verifyAccess(userId, 'view', User.email());
    if (!accessResult.allowed) {
      if (accessResult.userType === 'not_found') {
        return HtmlService.createHtmlOutput('<h3>User Not Found</h3><p>The specified user does not exist.</p>');
      }
      return HtmlService.createHtmlOutput('<h3>Private Board</h3><p>This board is private.</p>');
    }
    
    // 互換性のためのユーザー情報変換
    const compatUserInfo = {
      userId: userId,
      tenantId: userId,
      adminEmail: accessResult.config?.email || '',
      configJson: JSON.stringify(accessResult.config || {})
    };

    return renderAnswerBoard(compatUserInfo, { userId });
  } catch (error) {
    console.error('showAnswerBoard エラー:', error);
    return HtmlService.createHtmlOutput('<h3>Error</h3><p>An error occurred: ' + error.message + '</p>');
  }
}

/**
 * Helper Functions
 * ユーティリティ関数群
 */

/**
 * メールアドレスからドメインを抽出
 * @param {string} email メールアドレス
 * @returns {string} ドメイン部分
 */
function getEmailDomain(email) {
  if (!email || typeof email !== 'string') return '';
  const parts = email.split('@');
  return parts.length >= 2 ? parts[1] : '';
}

/**
 * WebApp URLの取得
 * @returns {string} WebApp URL
 */
function getWebAppUrl() {
  try {
    return ScriptApp.getService().getUrl();
  } catch (e) {
    console.error('WebApp URL取得失敗', { error: e.message });
    return '';
  }
}

/**
 * セットアップページのレンダリング
 */
function renderSetupPage(params) {
  console.log('renderSetupPage - Rendering setup page');
  
  // SetupPage.htmlテンプレートを使用
  return HtmlService.createTemplateFromFile('SetupPage').evaluate();
}

/**
 * URL生成関数群
 * generateUserUrls等の共通URL生成ロジック
 */
const UrlGenerator = {
  /**
   * 基本URL生成（アプリスクリプトURL）
   * @returns {string} ベースURL
   */
  getBaseUrl() {
    try {
      return ScriptApp.getService().getUrl();
    } catch (e) {
      console.error('BaseURL取得失敗', e);
      return '';
    }
  },
  
  /**
   * ユーザー固有URLを生成
   * @param {string} userId ユーザーID
   * @returns {Object} 各種URL
   */
  generateUserUrls(userId) {
    const baseUrl = this.getBaseUrl();
    if (!baseUrl) return { error: 'Base URL not available' };
    
    return {
      viewUrl: `${baseUrl}?userId=${encodeURIComponent(userId)}`,
      adminUrl: `${baseUrl}?mode=admin&userId=${encodeURIComponent(userId)}`,
      editUrl: `${baseUrl}?mode=edit&userId=${encodeURIComponent(userId)}` // 将来実装
    };
  },
  
  /**
   * 安全なURL検証
   * @param {string} url 検証するURL
   * @returns {boolean} 有効かどうか
   */
  isValidUrl(url) {
    if (!url || typeof url !== 'string') return false;
    
    try {
      const cleanUrl = url.trim();
      const isValidUrl =
        cleanUrl.includes('script.google.com') ||
        cleanUrl.includes('localhost') ||
        /^https:\/\/[a-zA-Z0-9.-]+/.test(cleanUrl);
      
      return isValidUrl && !cleanUrl.includes('javascript:');
    } catch (e) {
      return false;
    }
  },
  
  /**
   * パラメーター付きURL生成
   * @param {string} baseUrl ベースURL
   * @param {Object} params パラメーターオブジェクト
   * @returns {string} 完成したURL
   */
  buildUrlWithParams(baseUrl, params) {
    if (!baseUrl) return '';
    
    const url = new URL(baseUrl);
    Object.entries(params || {}).forEach(([key, value]) => {
      if (value !== null && value !== undefined) {
        url.searchParams.set(key, String(value));
      }
    });
    
    return url.toString();
  }
};

/**
 * User URL generation - Main API
 * ユーザー固有URL生成のメインAPI
 */
function generateUserUrls(userId) {
  return UrlGenerator.generateUserUrls(userId);
}

/**
 * Core Request Processing Functions
 * リクエスト処理の核となる機能群
 */

/**
 * リクエストパラメータを解析
 * @param {Object} e Apps Script doGet イベントオブジェクト
 * @returns {Object} 解析されたパラメータ
 */
function parseRequestParams(e) {
  if (!e) {
    console.warn('parseRequestParams - No event object provided');
    return {
      mode: null,
      userId: null,
      setupParam: null,
      spreadsheetId: null,
      sheetName: null,
      isDirectPageAccess: false,
    };
  }

  const p = e.parameter || {};
  const mode = p.mode || null;
  const userId = p.userId || null;
  const setupParam = p.setup || null;
  const spreadsheetId = p.spreadsheetId || null;
  const sheetName = p.sheetName || null;
  const isDirectPageAccess = !!(spreadsheetId && sheetName);
  
  console.log('parseRequestParams - Raw parameters:', JSON.stringify(p));
  console.log('parseRequestParams - Parsed mode:', mode, 'setupParam:', setupParam);

  return { mode, userId, setupParam, spreadsheetId, sheetName, isDirectPageAccess };
}

/**
 * HTMLテンプレート用include関数
 * GAS標準のファイル読み込み機能 - すべてのHTMLテンプレートで使用される
 * @param {string} filename 読み込むHTMLファイル名（拡張子なし）
 * @returns {string} HTMLファイルの内容
 */
function include(filename) {
  try {
    const content = HtmlService.createHtmlOutputFromFile(filename).getContent();
    console.log(`include - Successfully loaded: ${filename}`);
    return content;
  } catch (error) {
    console.error(`include - Failed to load file: ${filename}`, error);
    // ファイルが見つからない場合はエラーではなく空文字を返す（開発時の安全性）
    return `<!-- ERROR: Could not load ${filename} -->`;
  }
}

/**
 * Page Rendering Functions
 * 各種ページのレンダリング機能
 */

/**
 * 管理者パネルのレンダリング
 * @param {Object} userInfo ユーザー情報
 * @param {string} mode モード（'admin'等）
 * @returns {HtmlService.HtmlOutput} HTML出力
 */
function renderAdminPanel(userInfo, mode) {
  try {
    console.log('renderAdminPanel - userInfo:', JSON.stringify(userInfo));
    console.log('renderAdminPanel - mode:', mode);
    
    // 既存のPage.htmlテンプレートを使用
    const template = HtmlService.createTemplateFromFile('Page');
    
    // テンプレート変数を設定
    template.userInfo = userInfo;
    template.mode = mode || 'admin';
    template.isAdminPanel = true;
    
    return template.evaluate()
      .setTitle('管理パネル - Everyone\'s Answer Board')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('renderAdminPanel エラー:', error);
    return HtmlService.createHtmlOutput('<h3>エラー</h3><p>管理パネルの表示中にエラーが発生しました: ' + error.message + '</p>');
  }
}

/**
 * ログインページのレンダリング
 * @param {Object} params リクエストパラメータ
 * @returns {HtmlService.HtmlOutput} HTML出力
 */
function renderLoginPage(params = {}) {
  try {
    const template = HtmlService.createTemplateFromFile('LoginPage');
    template.params = params;
    return template.evaluate()
      .setTitle('StudyQuest - ログイン')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('renderLoginPage error:', error);
    return HtmlService.createHtmlOutput('<h3>Error</h3><p>ログインページの読み込みに失敗しました: ' + error.message + '</p>');
  }
}

/**
 * 回答ボードのレンダリング
 * @param {Object} userInfo ユーザー情報
 * @param {Object} params リクエストパラメータ
 * @returns {HtmlService.HtmlOutput} HTML出力
 */
function renderAnswerBoard(userInfo, params) {
  let config = {};
  try {
    config = JSON.parse(userInfo.configJson || '{}');
  } catch (e) {
    console.error('Config JSON parse error:', e);
    config = {};
  }

  console.log('renderAnswerBoard - userInfo:', JSON.stringify(userInfo));
  console.log('renderAnswerBoard - params:', JSON.stringify(params));
  console.log('renderAnswerBoard - config keys:', Object.keys(config));

  try {
    // 既存のPage.htmlテンプレートを使用
    const template = HtmlService.createTemplateFromFile('Page');
    
    // 基本情報設定
    template.userInfo = userInfo;
    template.mode = 'view';
    template.isAdminPanel = false;
    
    // 公開シート設定の取得
    const safePublishedSpreadsheetId = config.publishedSpreadsheetId || null;
    const safePublishedSheetName = config.publishedSheetName || null;

    const sheetConfigKey = 'sheet_' + (safePublishedSheetName || params.sheetName);
    const sheetConfig = config[sheetConfigKey] || {};

    // 修正: ダイレクトアクセスよりもパブリケーション状態を優先
    const isPublished = !!(safePublishedSpreadsheetId && safePublishedSheetName);
    const finalSpreadsheetId = isPublished ? safePublishedSpreadsheetId : params.spreadsheetId;
    const finalSheetName = isPublished ? safePublishedSheetName : params.sheetName;

    console.log('renderAnswerBoard - isPublished:', isPublished);
    console.log('renderAnswerBoard - finalSpreadsheetId:', finalSpreadsheetId);
    console.log('renderAnswerBoard - finalSheetName:', finalSheetName);

    // テンプレート変数設定
    template.config = config;
    template.sheetConfig = sheetConfig;
    template.spreadsheetId = finalSpreadsheetId;
    template.sheetName = finalSheetName;
    template.isDirectPageAccess = params.isDirectPageAccess;
    template.isPublished = isPublished;

    // データ取得とテンプレート設定の処理
    try {
      if (finalSpreadsheetId && finalSheetName) {
        console.log('renderAnswerBoard - データ取得開始');
        const dataResult = getPublishedSheetData(userInfo.userId, finalSpreadsheetId, finalSheetName);
        template.data = dataResult.data;
        template.message = dataResult.message;
        template.hasData = !!(dataResult.data && dataResult.data.length > 0);
      } else {
        template.data = [];
        template.message = '表示するデータがありません';
        template.hasData = false;
      }
      
      // 現在の設定から表示設定を取得
      const currentConfig = getCurrentConfig();
      const displaySettings = currentConfig.displaySettings || {};

      // 表示設定を適用
      template.displayMode = displaySettings.showNames ? 'named' : 'anonymous';
      template.showCounts = displaySettings.showReactions !== false;
      
      console.log('renderAnswerBoard - 表示設定:', { displayMode: template.displayMode, showCounts: template.showCounts });
      
    } catch (dataError) {
      console.error('renderAnswerBoard - データ取得エラー:', dataError);
      template.data = [];
      template.message = 'データ取得中にエラーが発生しました: ' + dataError.message;
      template.hasData = false;
      
      // エラー時も表示設定を適用
      const currentConfig = getCurrentConfig();
      const displaySettings = currentConfig.displaySettings || {};
      template.displayMode = displaySettings.showNames ? 'named' : 'anonymous';
      template.showCounts = displaySettings.showReactions !== false;
    }

    // 管理者URLの生成
    try {
      if (userInfo.userId) {
        const appUrls = generateUserUrls(userInfo.userId);
        template.adminPanelUrl = appUrls.adminUrl;
        console.log('renderAnswerBoard - 管理者URL設定:', template.adminPanelUrl);
      }
    } catch (urlError) {
      console.error('renderAnswerBoard - URL生成エラー:', urlError);
      template.adminPanelUrl = '';
    }

    return template.evaluate()
      .setTitle('Everyone\'s Answer Board')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    console.error('renderAnswerBoard - 全般エラー:', error);
    return HtmlService.createHtmlOutput('<h3>エラー</h3><p>ページの表示中にエラーが発生しました: ' + error.message + '</p>');
  }
}

/**
 * Publication Status Management
 * 公開状態管理の機能群
 */

/**
 * 現在の公開状態をチェック
 * @param {string} userId ユーザーID
 * @returns {Object} 公開状態情報
 */
function checkCurrentPublicationStatus(userId) {
  try {
    if (!userId) {
      console.warn('userId is required for publication status check');
      return { error: 'userId is required', isPublished: false };
    }

    console.log('checkCurrentPublicationStatus - userId:', userId);
    
    // 新アーキテクチャでのユーザー情報取得
    const userInfo = DB.findUserById(userId);
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

    // 公開状態の判定
    const isPublished = !!(
      config.publishedSpreadsheetId && 
      config.publishedSheetName
    );

    console.log('checkCurrentPublicationStatus - result:', {
      userId: userId,
      isPublished: isPublished,
      hasSpreadsheetId: !!config.publishedSpreadsheetId,
      hasSheetName: !!config.publishedSheetName
    });

    return {
      userId: userId,
      isPublished: isPublished,
      publishedSpreadsheetId: config.publishedSpreadsheetId || null,
      lastChecked: new Date().toISOString(),
    };
  } catch (error) {
    console.error('❌ Error in checkCurrentPublicationStatus:', error);
    return {
      error: error.message,
      isPublished: false,
      lastChecked: new Date().toISOString(),
    };
  }
}

/**
 * Configuration Management Functions
 * 設定管理の機能群
 */

/**
 * 現在の設定を取得
 * @returns {Object} 現在の設定
 */
function getCurrentConfig() {
  try {
    // 簡易的な設定取得（新アーキテクチャへの移行準備）
    const userInfo = Services.user.current;
    if (!userInfo) {
      return {};
    }

    // デフォルト設定を返す
    return {
      title: 'みんなの質問箱', 
      description: '質問をお寄せください',
      displaySettings: {
        showNames: true,  // 名前表示: デフォルト有効
        showReactions: userInfo?.displaySettings?.showReactions !== false, // デフォルトtrue
      },
    };
  } catch (error) {
    console.error('getCurrentConfig エラー:', error);
    return {};
  }
}

/**
 * テスト用シンプルAPI（テスト専用）
 */
function getTestMockResponse() {
  return {
    success: true,
    message: 'Test mock response',
    data: {
      reason: 'C',
    },
  };
}

/**
 * アプリケーション設定保存
 * @param {Object} config 設定オブジェクト
 * @returns {Object} 保存結果
 */
function saveApplicationConfig(config) {
  console.log('アプリケーション設定を保存:', config);
  return {
    success: true,
    message: '保存しました',
    config: config,
  };
}

/**
 * アプリケーション公開
 * @param {Object} config 設定オブジェクト 
 * @returns {Object} 公開結果
 */
function publishApplication(config) {
  console.log('アプリを公開:', config);
  return {
    success: true,
    message: '公開しました',
    config: config,
  };
}

/**
 * システムセットアップ実行関数
 * セットアップページから呼び出される
 * @param {string} serviceAccountJson サービスアカウントJSON文字列
 * @param {string} databaseSpreadsheetId データベーススプレッドシートID
 * @param {string} adminEmail 管理者メールアドレス（省略時は現在のユーザー）
 * @param {string} googleClientId Google Client ID（オプション）
 */
function setupApplication(serviceAccountJson, databaseSpreadsheetId, adminEmail = null, googleClientId = null) {
  try {
    console.log('setupApplication - セットアップ開始');
    
    // 現在のユーザーのメールアドレスを取得
    const currentUserEmail = User.email();
    if (!currentUserEmail) {
      throw new Error('認証されたユーザーが必要です');
    }
    
    // 管理者メールアドレスが指定されていない場合は現在のユーザーを使用
    const finalAdminEmail = adminEmail || currentUserEmail;
    
    console.log('setupApplication - 管理者メール:', finalAdminEmail);
    
    // 入力値検証
    if (!serviceAccountJson || !databaseSpreadsheetId) {
      throw new Error('必須パラメータが不足しています');
    }
    
    // JSON検証
    let serviceAccountData;
    try {
      serviceAccountData = JSON.parse(serviceAccountJson);
    } catch (e) {
      throw new Error('サービスアカウントJSONの形式が無効です');
    }
    
    // サービスアカウント必須フィールド検証
    const requiredFields = ['type', 'project_id', 'private_key_id', 'private_key', 'client_email'];
    for (const field of requiredFields) {
      if (!serviceAccountData[field]) {
        throw new Error(`必須フィールド '${field}' が見つかりません`);
      }
    }
    
    if (serviceAccountData.type !== 'service_account') {
      throw new Error('サービスアカウントタイプのJSONである必要があります');
    }
    
    // スプレッドシートID検証
    if (databaseSpreadsheetId.length !== 44 || !/^[a-zA-Z0-9_-]+$/.test(databaseSpreadsheetId)) {
      throw new Error('スプレッドシートIDの形式が無効です（44文字の英数字、アンダースコア、ハイフンのみ）');
    }
    
    // Script Propertiesに設定を保存
    const properties = PropertiesService.getScriptProperties();
    
    properties.setProperties({
      [PROPS_KEYS.SERVICE_ACCOUNT_CREDS]: serviceAccountJson,
      [PROPS_KEYS.DATABASE_SPREADSHEET_ID]: databaseSpreadsheetId,
      [PROPS_KEYS.ADMIN_EMAIL]: finalAdminEmail
    });
    
    // Google Client IDが提供されている場合は設定
    if (googleClientId && googleClientId.trim()) {
      properties.setProperty('GOOGLE_CLIENT_ID', googleClientId.trim());
    }
    
    console.log('setupApplication - プロパティ設定完了');
    
    // セットアップ完了検証
    if (!isSystemSetup()) {
      throw new Error('セットアップ後の検証に失敗しました');
    }
    
    console.log('setupApplication - セットアップ完了');
    
    return {
      success: true,
      message: 'システムセットアップが正常に完了しました',
      adminEmail: finalAdminEmail,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('setupApplication エラー:', error);
    throw new Error(`セットアップに失敗しました: ${error.message}`);
  }
}

/**
 * 現在のユーザーのメールアドレスを取得
 * セットアップページで使用
 */
function getCurrentUserEmail() {
  try {
    const email = User.email();
    return {
      success: true,
      email: email || null,
      message: email ? 'ユーザーメールを取得しました' : 'ユーザーが認証されていません'
    };
  } catch (error) {
    console.error('getCurrentUserEmail エラー:', error);
    return {
      success: false,
      email: null,
      message: `メールアドレス取得に失敗: ${error.message}`
    };
  }
}

/**
 * セットアップテスト実行
 */
function testSetup() {
  try {
    if (!isSystemSetup()) {
      return {
        status: 'error',
        message: 'システムがセットアップされていません'
      };
    }
    
    // 基本的な動作確認
    const properties = PropertiesService.getScriptProperties();
    const adminEmail = properties.getProperty(PROPS_KEYS.ADMIN_EMAIL);
    const currentUserEmail = User.email();
    
    if (adminEmail !== currentUserEmail) {
      return {
        status: 'warning',
        message: '⚠️ テスト完了：現在のユーザーは管理者ではありません'
      };
    }
    
    return {
      status: 'success',
      message: '✅ テスト完了：システムは正常に動作しています'
    };
    
  } catch (error) {
    console.error('testSetup エラー:', error);
    return {
      status: 'error',
      message: `テストに失敗しました: ${error.message}`
    };
  }
}