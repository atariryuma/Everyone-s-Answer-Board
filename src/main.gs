/**
 * main.gs - Application Entry Points
 * 新しいサービス層アーキテクチャに基づくエントリーポイント
 * 
 * 🎯 責任範囲:
 * - HTTP requests routing (doGet/doPost)
 * - User interface rendering
 * - Service layer coordination
 * - Error handling & user feedback
 */

/**
 * Application Configuration
 * サービス層統合後の軽量設定
 */
const APP_CONFIG = Object.freeze({
  // アプリケーション情報
  APP_NAME: 'Everyone\'s Answer Board',
  VERSION: '2.0.0',
  
  // ページモード
  MODES: Object.freeze({
    MAIN: 'main',
    ADMIN: 'admin', 
    LOGIN: 'login',
    SETUP: 'setup',
    DEBUG: 'debug'
  }),

  // デフォルト設定
  DEFAULTS: Object.freeze({
    CACHE_TTL: 300, // 5分
    PAGE_SIZE: 50,
    MAX_RETRIES: 3
  })
});

// 廃止予定のUserオブジェクトを削除（UserManager.getCurrentEmail()を使用）

/**
 * Entry Point Functions
 * 新アーキテクチャにおけるエントリーポイント関数群
 */

/**
 * doGet - HTTP GET request handler
 * 新しいサービス層アーキテクチャを使用
 */
function doGet(e) {
  try {
    // リクエストパラメータを解析
    const params = parseRequestParams(e);
    
    // セキュリティチェック：認証状態確認
    const userEmail = UserService.getCurrentEmail();
    if (!userEmail && params.mode !== APP_CONFIG.MODES.LOGIN && params.mode !== APP_CONFIG.MODES.SETUP) {
      // 未認証の場合はログインページにリダイレクト
      return handleLoginMode(params, { reason: 'authentication_required' });
    }
    
    // アクセス権限検証（認証済みユーザーのみ）
    if (userEmail) {
      const accessResult = SecurityService.validateWebAppAccess(userEmail, params);
      if (!accessResult.hasAccess) {
        return renderErrorPage({
          success: false,
          message: 'アクセスが拒否されました',
          details: accessResult.reason,
          canRetry: false
        });
      }
    }
    
    // システム初期化確認
    if (!ConfigService.isSystemSetup()) {
      return renderSetupPage(params);
    }

    // モードに応じたルーティング
    switch (params.mode) {
      case APP_CONFIG.MODES.DEBUG:
        return handleDebugMode(params);
      
      case APP_CONFIG.MODES.ADMIN:
        return handleAdminMode(params);
      
      case APP_CONFIG.MODES.LOGIN:
        return handleLoginMode(params);
      
      case APP_CONFIG.MODES.SETUP:
        return handleSetupMode(params);
      
      default:
        return handleMainMode(params);
    }
  } catch (error) {
    // 統一エラーハンドリング使用
    const errorResponse = ErrorHandler.handle(error, {
      service: 'main',
      method: 'doGet',
      parameters: e?.parameter
    });
    
    return renderErrorPage(errorResponse);
  }
}

/**
 * デバッグモード処理
 * @param {Object} params - リクエストパラメータ
 * @returns {HtmlOutput} デバッグページ
 */
function handleDebugMode(params) {
  try {
    // 新しいサービス層を使用してユーザー情報取得
    const userInfo = UserService.getCurrentUserInfo();
    const sessionStatus = UserService.getSessionStatus();
    
    const debugData = {
      app: {
        name: APP_CONFIG.APP_NAME,
        version: APP_CONFIG.VERSION,
        timestamp: new Date().toISOString()
      },
      session: sessionStatus,
      user: userInfo ? {
        userId: userInfo.userId,
        userEmail: userInfo.userEmail,
        isActive: userInfo.isActive,
        hasConfig: !!userInfo.config,
        configCompleteness: userInfo.config?.completionScore || 0
      } : null,
      services: {
        user: UserService.diagnose(),
        config: ConfigService.diagnose(),
        data: DataService.diagnose(),
        security: SecurityService.diagnose()
      },
      suggestion: !userInfo 
        ? 'ユーザー登録が必要です。mode=loginでアクセスしてください。'
        : 'システムは正常に動作しています'
    };

    return HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>Debug - ${APP_CONFIG.APP_NAME}</title>
        <style>
          body { font-family: monospace; margin: 20px; background: #f5f5f5; }
          .container { max-width: 1200px; margin: 0 auto; background: white; padding: 20px; border-radius: 8px; }
          pre { background: #f8f8f8; padding: 15px; border-radius: 4px; overflow-x: auto; }
          .status-ok { color: #28a745; }
          .status-warning { color: #ffc107; }
          .status-error { color: #dc3545; }
          .nav { margin-bottom: 20px; }
          .nav a { margin-right: 15px; text-decoration: none; color: #007bff; }
        </style>
      </head>
      <body>
        <div class="container">
          <h2>🔍 ${APP_CONFIG.APP_NAME} - Debug Information</h2>
          
          <div class="nav">
            ${userInfo ? `<a href="?mode=admin&userId=${userInfo.userId}">📊 管理パネル</a>` : ''}
            <a href="?mode=main">🏠 メインページ</a>
            <a href="?mode=setup">⚙️ セットアップ</a>
          </div>
          
          <h3>📋 System Status</h3>
          <pre>${JSON.stringify(debugData, null, 2)}</pre>
          
          <h3>💡 Recommendations</h3>
          <p>${debugData.suggestion}</p>
          
          <hr>
          <small>Generated at: ${new Date().toLocaleString('ja-JP')}</small>
        </div>
      </body>
      </html>
    `);
  } catch (error) {
    const errorResponse = ErrorHandler.handle(error, {
      service: 'main',
      method: 'handleDebugMode'
    });
    return renderErrorPage(errorResponse);
  }
}

/**
 * 管理モード処理
 * @param {Object} params - リクエストパラメータ
 * @returns {HtmlOutput} 管理ページ
 */
function handleAdminMode(params) {
  try {
    const userId = params.userId;
    if (!userId) {
      throw new Error('ユーザーIDが必要です');
    }

    // 権限確認
    const permission = SecurityService.checkUserPermission(userId, 'owner');
    if (!permission.hasPermission) {
      return renderErrorPage({
        success: false,
        message: 'このページにアクセスする権限がありません',
        canRetry: false
      });
    }

    // バルクデータ取得
    const bulkResult = DataService.getBulkData(userId, {
      includeUserInfo: true,
      includeConfig: true,
      includeSheetData: true,
      includeSystemInfo: true
    });

    if (!bulkResult.success) {
      throw new Error(bulkResult.error);
    }

    // 既存のAdminPanel.htmlを使用（新しいデータ形式で）
    const template = HtmlService.createTemplateFromFile('AdminPanel');
    template.userData = bulkResult.data;
    
    return template.evaluate()
      .setTitle(`管理パネル - ${APP_CONFIG.APP_NAME}`)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    const errorResponse = ErrorHandler.handle(error, {
      service: 'main',
      method: 'handleAdminMode',
      userId: params.userId
    });
    return renderErrorPage(errorResponse);
  }
}

/**
 * ログインモード処理
 * @param {Object} params - リクエストパラメータ
 * @returns {HtmlOutput} ログインページ
 */
function handleLoginMode(params) {
  try {
    // 現在のセッション状態確認
    const sessionStatus = UserService.getSessionStatus();
    
    if (sessionStatus.isAuthenticated && sessionStatus.hasUserInfo) {
      // 既にログイン済みの場合は管理パネルにリダイレクト
      const userInfo = UserService.getCurrentUserInfo();
      if (userInfo) {
        return HtmlService.createHtmlOutput(`
          <script>
            window.location.href = '?mode=admin&userId=${userInfo.userId}';
          </script>
        `);
      }
    }

    // 新規ユーザー作成処理
    if (sessionStatus.isAuthenticated && !sessionStatus.hasUserInfo) {
      try {
        const newUser = UserService.createUser(sessionStatus.userEmail);
        if (newUser) {
          console.info('main.handleLoginMode: 新規ユーザー作成完了', {
            userEmail: sessionStatus.userEmail,
            userId: newUser.userId
          });
          
          return HtmlService.createHtmlOutput(`
            <script>
              window.location.href = '?mode=admin&userId=${newUser.userId}';
            </script>
          `);
        }
      } catch (createError) {
        console.error('main.handleLoginMode: ユーザー作成エラー', createError.message);
      }
    }

    // LoginPage.htmlを使用
    const template = HtmlService.createTemplateFromFile('LoginPage');
    template.sessionStatus = sessionStatus;
    
    return template.evaluate()
      .setTitle(`ログイン - ${APP_CONFIG.APP_NAME}`)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    const errorResponse = ErrorHandler.handle(error, {
      service: 'main',
      method: 'handleLoginMode'
    });
    return renderErrorPage(errorResponse);
  }
}

/**
 * メインモード処理
 * @param {Object} params - リクエストパラメータ
 * @returns {HtmlOutput} メインページ
 */
function handleMainMode(params) {
  try {
    const userId = params.userId;
    if (!userId) {
      // ユーザーIDが無い場合はログインに誘導
      return handleLoginMode(params);
    }

    // ユーザー設定取得
    const config = ConfigService.getUserConfig(userId);
    if (!config) {
      throw new Error('ユーザー設定が見つかりません');
    }

    // アプリが公開されていない場合
    if (!config.appPublished) {
      return renderErrorPage({
        success: false,
        message: 'このアプリケーションは現在公開されていません',
        canRetry: false
      });
    }

    // メインデータ取得
    const sheetResult = DataService.getSheetData(userId);
    
    // Page.htmlテンプレートを使用
    const template = HtmlService.createTemplateFromFile('Page');
    template.config = config;
    template.sheetData = sheetResult.success ? sheetResult.data : [];
    template.userId = userId;
    
    return template.evaluate()
      .setTitle(config.title || APP_CONFIG.APP_NAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    const errorResponse = ErrorHandler.handle(error, {
      service: 'main', 
      method: 'handleMainMode',
      userId: params.userId
    });
    return renderErrorPage(errorResponse);
  }
}

/**
 * エラーページレンダリング
 * @param {Object} errorResponse - エラーレスポンス
 * @returns {HtmlOutput} エラーページ
 */
function renderErrorPage(errorResponse) {
  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <title>エラー - ${APP_CONFIG.APP_NAME}</title>
      <style>
        body { font-family: sans-serif; margin: 40px; background: #f8f9fa; }
        .error-container { max-width: 600px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .error-icon { font-size: 48px; text-align: center; margin-bottom: 20px; }
        .error-message { font-size: 18px; margin-bottom: 20px; color: #dc3545; }
        .error-details { background: #f8f9fa; padding: 15px; border-radius: 4px; font-family: monospace; font-size: 12px; color: #6c757d; }
        .actions { margin-top: 20px; text-align: center; }
        .btn { display: inline-block; padding: 10px 20px; margin: 0 5px; text-decoration: none; border-radius: 4px; }
        .btn-primary { background: #007bff; color: white; }
        .btn-secondary { background: #6c757d; color: white; }
      </style>
    </head>
    <body>
      <div class="error-container">
        <div class="error-icon">❌</div>
        <h2>エラーが発生しました</h2>
        <div class="error-message">${errorResponse.message}</div>
        ${errorResponse.canRetry ? '<p>しばらく待ってから再試行してください。</p>' : ''}
        <div class="error-details">
          エラーID: ${errorResponse.errorId || 'unknown'}<br>
          発生時刻: ${errorResponse.timestamp || new Date().toISOString()}
        </div>
        <div class="actions">
          ${errorResponse.canRetry ? '<a href="#" onclick="location.reload()" class="btn btn-primary">再試行</a>' : ''}
          <a href="?mode=debug" class="btn btn-secondary">システム状態確認</a>
          <a href="?mode=login" class="btn btn-secondary">ログイン画面</a>
        </div>
      </div>
    </body>
    </html>
  `;
  
  return HtmlService.createHtmlOutput(html)
    .setTitle(`エラー - ${APP_CONFIG.APP_NAME}`);
            }
            <hr>
            <h3>🔧 GAS実行用関数</h3>
            <p><strong>GASエディタから直接実行してください：</strong></p>
            <ul>
              <li><code>testSystemStatus()</code> - システム診断</li>
              <li><code>fixConfigJsonNesting()</code> - configJSON重複修正</li>
              <li><code>resetCurrentUserToDefault()</code> - 現在のユーザーをリセット</li>
              <li><code>resetUserToDefault(userId)</code> - 指定ユーザーをリセット</li>
            </ul>
            ${userByEmail ? `<p><small>現在のユーザーID: ${userByEmail.userId}</small></p>` : ''}
          `);
        } catch (error) {
          return HtmlService.createHtmlOutput(`<h2>Debug Error</h2><pre>${error.message}</pre>`);
        }

      case 'test':
        // 🧪 統合システムテストモード
        try {
          const testResult = testUnifiedSystem();
          return HtmlService.createHtmlOutput(`
            <h2>🧪 統合システムテスト結果</h2>
            <div style="font-family: monospace; background: #f5f5f5; padding: 20px; border-radius: 8px;">
              <h3>テスト結果: ${testResult.success ? '✅ 成功' : '❌ 失敗'}</h3>
              <p><strong>メッセージ:</strong> ${testResult.message}</p>
              ${
                testResult.config
                  ? `
                <h3>設定情報:</h3>
                <pre>${JSON.stringify(testResult.config, null, 2)}</pre>
              `
                  : ''
              }
            </div>
            <hr>
            <p><a href="?mode=debug">デバッグ情報を見る</a></p>
            <p><a href="${WebApp.getUrl()}">ホームに戻る</a></p>
          `);
        } catch (error) {
          return HtmlService.createHtmlOutput(`
            <h2>❌ テストエラー</h2>
            <pre>${error.message}</pre>
            <p><a href="${WebApp.getUrl()}">ホームに戻る</a></p>
          `);
        }

      case 'fix_user':
        // 🔧 緊急修正：ユーザーデータの修正
        try {
          if (!params.userId) {
            return HtmlService.createHtmlOutput('<h2>Error</h2><p>userIdが必要です</p>');
          }

          const { currentUserEmail } = UserManager.getCurrentUserInfo();
          const userInfo = DB.findUserById(params.userId);

          if (!userInfo) {
            return HtmlService.createHtmlOutput('<h2>Error</h2><p>ユーザーが見つかりません</p>');
          }

          // 🚨 緊急修正：userEmailとisActiveを設定
          const updatedData = {
            userEmail: currentUserEmail,
            isActive: true,
          };

          // データベース更新（直接フィールド更新）
          DB.updateUserFields(params.userId, updatedData);

          return HtmlService.createHtmlOutput(`
            <h2>✅ ユーザーデータ修正完了</h2>
            <p>userEmail: ${currentUserEmail}</p>
            <p>isActive: true</p>
            <p><a href="?mode=admin&userId=${params.userId}">管理パネルにアクセス</a></p>
          `);
        } catch (error) {
          return HtmlService.createHtmlOutput(`<h2>修正エラー</h2><pre>${error.message}</pre>`);
        }

      case 'admin':
        // 管理パネルモード: userIdが必須
        if (!params.userId) {
          // userIdが無い場合はログイン画面へ
          return renderLoginPage(params);
        }

        try {
          const { currentUserEmail, userInfo } = UserManager.getCurrentUserInfo();
          if (!userInfo || userInfo.userId !== params.userId) {
            // ユーザーが存在しないか、userIdが一致しない場合
            return showErrorPage('アクセス拒否', '管理パネルへのアクセス権限がありません');
          }

          if (userInfo.isActive !== true) {
            return showErrorPage(
              'アカウントが無効です',
              'あなたのアカウントは現在無効化されています。'
            );
          }

          // アクセス検証
          const accessResult = App.getAccess().verifyAccess(
            params.userId,
            'admin',
            currentUserEmail
          );
          if (!accessResult.allowed) {
            return showErrorPage('アクセス拒否', '管理パネルへのアクセス権限がありません');
          }

          // ユーザー情報変換
          const compatUserInfo = {
            userId: params.userId,
            userEmail: userInfo.userEmail,
            configJson: userInfo.configJson,
          };

          return renderAdminPanel(compatUserInfo, 'admin');
        } catch (adminError) {
          console.error('Admin mode error:', adminError);
          return showErrorPage(
            '管理パネルエラー',
            `管理パネルの表示中にエラーが発生しました: ${adminError.message}`
          );
        }

      case 'clear_cache':
        // 🚨 緊急時キャッシュクリア（システム復旧用）
        try {
          SimpleCacheManager.clearAll();
          return HtmlService.createHtmlOutput(`
            <h2>✅ キャッシュクリア完了</h2>
            <p>Service Accountトークンとキャッシュをクリアしました</p>
            <p><a href="${WebApp.getUrl()}">ホームに戻る</a></p>
          `);
        } catch (clearError) {
          return HtmlService.createHtmlOutput(`
            <h2>❌ キャッシュクリアエラー</h2>
            <pre>${clearError.message}</pre>
          `);
        }

      case 'login':
        // ログインモード: 明示的にログイン画面を表示
        return renderLoginPage(params);

      case 'view':
        // ビューモード: userIdが必要
        if (!params.userId) {
          // userIdがない場合はログイン画面へ
          return renderLoginPage(params);
        }

        try {
          const accessResult = App.getAccess().verifyAccess(
            params.userId,
            'view',
            UserManager.getCurrentEmail()
          );
          if (!accessResult.allowed) {
            console.warn('View access denied:', accessResult);
            if (accessResult.userType === 'not_found') {
              return HtmlService.createHtmlOutput(
                '<h3>User Not Found</h3><p>The specified user does not exist.</p>'
              );
            }
            return HtmlService.createHtmlOutput(
              '<h3>Private Board</h3><p>This board is private.</p>'
            );
          }

          // ユーザー情報変換（DBからの完全なユーザー情報を取得）
          const dbUserInfo = DB.findUserById(params.userId);

          const compatUserInfo = {
            userId: params.userId,
            userEmail: accessResult.config?.userEmail || '',
            configJson: JSON.stringify(accessResult.config || {}),
            spreadsheetId: dbUserInfo?.spreadsheetId || null,
            sheetName: dbUserInfo?.sheetName || null,
          };

          return renderAnswerBoard(compatUserInfo, params);
        } catch (viewError) {
          console.error('View mode error:', viewError);
          return HtmlService.createHtmlOutput(
            `<h3>Error</h3><p>An error occurred: ${viewError.message}</p>`
          );
        }

      default:
        // デフォルト: ログイン画面を表示
        return renderLoginPage(params);
    }
  } catch (error) {
    console.error('doGet - Critical error:', error);
    return HtmlService.createHtmlOutput(
      `<h3>System Error</h3><p>A critical system error occurred: ${error.message}</p>`
    );
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
        // UnifiedManagerを使用（重複削除）
        const userInfo = UnifiedManager.user.getCurrentInfo();
        if (!userInfo) return null;

        return {
          email: userInfo.userEmail,
          isAuthenticated: true,
        };
      } catch (error) {
        console.error('Services.user.current エラー:', error.message);
        return null;
      }
    },

    getActiveUserInfo() {
      try {
        // UnifiedManagerを直接使用
        return UnifiedManager.user.getCurrentInfo();
      } catch (error) {
        console.error('Services.user.getActiveUserInfo エラー:', error.message);
        return null;
      }
    },
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
        console.error('アプリケーションアクセスチェック失敗:', error.message);
        return { hasAccess: true, message: 'アクセス許可（エラー回避）' };
      }
    },
  },
};

// デプロイ関連の機能
const Deploy = {
  // ドメイン情報を取得する関数
  domain() {
    try {
      console.log('Deploy.domain() - start');

      const activeUserEmail = UserManager.getCurrentUserEmail();
      const currentDomain = getEmailDomain(activeUserEmail);

      // WebAppのURLを取得してドメインを判定
      const webAppUrl = getWebAppUrl();
      let deployDomain = ''; // 個人アカウント/グローバルアクセスの場合、デフォルトで空

      if (webAppUrl && webAppUrl.includes('/a/')) {
        // Google Workspace環境でのドメイン取得
        const domainMatch = webAppUrl.match(/\/a\/([a-zA-Z0-9\-.]+)\/macros\//);
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
        currentDomain,
        deployDomain,
        isDomainMatch,
        webAppUrl,
      };
    } catch (e) {
      console.error('デプロイユーザードメイン情報取得失敗:', e.message);
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
      const currentUserEmail = UserManager.getCurrentUserEmail();

      console.log('Deploy.isUser() - 管理者確認:', adminEmail, currentUserEmail);
      return adminEmail === currentUserEmail;
    } catch (error) {
      console.error('Deploy.isUser() エラー:', error.message);
      return false;
    }
  },
};

/**
 * 簡素化されたエラーハンドリング関数群
 */

// PerformanceOptimizer.gsでglobalProfilerが既に定義されているため、
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
    const clientId = properties.getProperty(PROPS_KEYS.GOOGLE_CLIENT_ID);

    if (!clientId) {
      console.warn('GOOGLE_CLIENT_ID not found in script properties');

      // デバッグ用：利用可能なプロパティを確認
      const allProperties = properties.getProperties();
      console.log('Available properties count:', Object.keys(allProperties).length);

      return createResponse(false, 'Google Client ID not configured', {
        setupInstructions:
          'Please set GOOGLE_CLIENT_ID in Google Apps Script project settings under Properties > Script Properties',
      });
    }

    return createResponse(true, 'Google Client IDを取得しました', { clientId });
  } catch (error) {
    console.error('GOOGLE_CLIENT_ID取得エラー:', error.message);
    return createResponse(
      false,
      `Google Client IDの取得に失敗しました: ${error.toString()}`,
      { clientId: '' },
      error
    );
  }
}

/**
 * システム設定状況をチェック
 */
function checkSystem() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const allProperties = properties.getProperties();

    const requiredProperties = [
      PROPS_KEYS.GOOGLE_CLIENT_ID,
      PROPS_KEYS.DATABASE_SPREADSHEET_ID,
      PROPS_KEYS.ADMIN_EMAIL,
      PROPS_KEYS.SERVICE_ACCOUNT_CREDS,
    ];

    const configStatus = {};
    const missingProperties = [];

    requiredProperties.forEach((prop) => {
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

    return createResponse(true, 'System domain info retrieved', {
      adminEmail,
      adminDomain,
      isDomainMatch,
      currentDomain: domainInfo.currentDomain || '不明',
      deployDomain: domainInfo.deployDomain || adminDomain,
    });
  } catch (e) {
    console.error('getSystemDomainInfo エラー:', e.message);
    return createResponse(false, 'System domain info取得エラー', null, e);
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
    const activeUserEmail = UserManager.getCurrentUserEmail();
    if (activeUserEmail) {
      const userProperties = PropertiesService.getUserProperties();
      const lastAdminUserId = userProperties.getProperty('lastAdminUserId');

      if (lastAdminUserId) {
        console.log('showAdminPanel - Using saved userId:', lastAdminUserId);
        try {
          const accessResult = App.getAccess().verifyAccess(
            lastAdminUserId,
            'admin',
            activeUserEmail
          );
          if (accessResult.allowed) {
            console.log('✅ アクセス許可: ', accessResult);
            const compatUserInfo = {
              userId: lastAdminUserId,
              userId: lastAdminUserId,
              userEmail: accessResult.config?.email || activeUserEmail,
              configJson: JSON.stringify(accessResult.config || {}),
            };

            return renderAdminPanel(compatUserInfo, 'admin');
          }
        } catch (error) {
          console.warn('Saved userId access failed:', error.message);
        }
      }
    }

    // デフォルト：エラーページを表示
    return HtmlService.createHtmlOutput(
      '<h3>Admin Access Required</h3><p>Please access via the admin URL with proper userId parameter.</p>'
    );
  } catch (error) {
    console.error('showAdminPanel エラー:', error);
    return HtmlService.createHtmlOutput(`<h3>Error</h3><p>An error occurred: ${error.message}</p>`);
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
    const accessResult = App.getAccess().verifyAccess(
      userId,
      'view',
      UserManager.getCurrentEmail()
    );
    if (!accessResult.allowed) {
      if (accessResult.userType === 'not_found') {
        return HtmlService.createHtmlOutput(
          '<h3>User Not Found</h3><p>The specified user does not exist.</p>'
        );
      }
      return HtmlService.createHtmlOutput('<h3>Private Board</h3><p>This board is private.</p>');
    }

    // ユーザー情報変換
    const compatUserInfo = {
      userId,
      userEmail: accessResult.config?.userEmail || '',
      configJson: JSON.stringify(accessResult.config || {}),
    };

    return renderAnswerBoard(compatUserInfo, { userId });
  } catch (error) {
    console.error('showAnswerBoard エラー:', error);
    return HtmlService.createHtmlOutput(`<h3>Error</h3><p>An error occurred: ${error.message}</p>`);
  }
}

/**
 * 🦾 スマート診断システムサポート関数
 * ユーザーに適切なアクションを推奨する
 */
function getSuggestedAction(diagnostics) {
  try {
    if (!diagnostics.hasSpreadsheetId) {
      return '管理パネルでデータソース（スプレッドシート）を接続してください';
    }

    if (!diagnostics.hasSheetName) {
      return '管理パネルでシートを選択してください';
    }

    if (!diagnostics.appPublished) {
      return '管理パネルでアプリを公開してください';
    }

    if (diagnostics.setupStatus !== 'completed') {
      return '初期設定を完了してください';
    }

    return 'システムは正常に動作しています';
  } catch (error) {
    console.error('getSuggestedAction エラー:', error.message);
    return '管理パネルで設定を確認してください';
  }
}

/**
 * Helper Functions
 * ユーティリティ関数群
 */

/**
 * ユーザーIDを生成（UUID形式）
 * @returns {string} 新規ユーザーID
 */
function generateUserId() {
  return Utilities.getUuid();
}

// getEmailDomain: database.gsに統合済み

/**
 * WebApp URLの取得
 * @returns {string} WebApp URL
 */
function getWebAppUrl() {
  // 1. 複数の方法でWebApp URLを取得試行
  const urlMethods = [
    () => ScriptApp.getService().getUrl(),
    () => {
      const service = ScriptApp.newWebApp();
      return service ? service.getUrl() : null;
    },
  ];

  for (let i = 0; i < urlMethods.length; i++) {
    try {
      const url = urlMethods[i]();
      if (url && url.includes('script.google.com') && url.includes('exec')) {
        return url;
      }
    } catch (e) {
      console.warn(`getWebAppUrl: 方法${i + 1}失敗:`, e.message);
    }
  }

  // 2. フォールバック: PropertiesServiceから保存済みURLを確認
  try {
    const props = PropertiesService.getScriptProperties();
    const savedUrl = props.getProperty('CACHED_WEBAPP_URL');
    if (savedUrl && savedUrl.includes('script.google.com')) {
      console.log('getWebAppUrl: キャッシュされたURL使用', savedUrl);
      return savedUrl;
    }
  } catch (e) {
    console.warn('getWebAppUrl: キャッシュURL確認失敗:', e.message);
  }

  // 3. 最終フォールバック: Script IDから構築（改良版）
  try {
    const scriptId = ScriptApp.getScriptId();
    if (scriptId) {
      console.log('getWebAppUrl: Script ID確認', scriptId);

      // Google Workspace環境を考慮した動的URL構築
      const result = new ConfigurationManager().getCurrentUserInfoSafely();
      const currentUser = result?.currentUserEmail;
      const domain = getEmailDomain(currentUser);

      let baseUrl;
      if (domain && domain !== 'gmail.com' && domain !== 'googlemail.com') {
        // Google Workspace環境
        baseUrl = `https://script.google.com/a/${domain}/macros/s/${scriptId}/exec`;
      } else {
        // 個人Google環境
        baseUrl = `https://script.google.com/macros/s/${scriptId}/exec`;
      }

      // URLをキャッシュ保存
      try {
        const props = PropertiesService.getScriptProperties();
        props.setProperty('CACHED_WEBAPP_URL', baseUrl);
        console.log('getWebAppUrl: URL保存成功');
      } catch (cacheError) {
        console.warn('getWebAppUrl: URL保存失敗:', cacheError.message);
      }

      console.log('getWebAppUrl: フォールバックURL生成成功', baseUrl);
      return baseUrl;
    } else {
      console.error('getWebAppUrl: Script ID取得失敗');
    }
  } catch (e) {
    console.error('getWebAppUrl: フォールバック失敗:', e.message);
  }

  console.error('getWebAppUrl: 全ての方法が失敗');
  return '';
}

/**
 * メールアドレスからドメイン部分を抽出
 * @param {string} email - メールアドレス
 * @returns {string} ドメイン部分
 */
function getEmailDomain(email) {
  if (!email || typeof email !== 'string') return '';
  const parts = email.split('@');
  return parts.length >= 2 ? parts[1] : '';
}

/**
 * 管理パネルのURLを構築
 * @param {string} userId - ユーザーID
 * @returns {string} 管理パネルURL
 */
function buildUserAdminUrl(userId) {
  return generateUserUrls(userId).adminUrl;
}

/**
 * エラーページを表示
 * @param {string} title - エラータイトル
 * @param {string} message - エラーメッセージ
 * @param {Error} error - エラーオブジェクト（オプション）
 * @returns {HtmlOutput} エラーHTML
 */
function showErrorPage(title, message, error) {
  const template = HtmlService.createTemplate(`
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title><?= title ?></title>
      <style>
        body {
          font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
          background: linear-gradient(135deg, #ff6b6b 0%, #ee5a24 100%);
          margin: 0;
          padding: 0;
          display: flex;
          justify-content: center;
          align-items: center;
          min-height: 100vh;
          color: #333;
        }
        .error-container {
          background: white;
          border-radius: 12px;
          box-shadow: 0 10px 30px rgba(0,0,0,0.2);
          padding: 40px 30px;
          text-align: center;
          max-width: 500px;
          width: 90%;
        }
        .error-icon {
          font-size: 60px;
          color: #ff6b6b;
          margin-bottom: 20px;
        }
        h1 {
          color: #ff6b6b;
          margin-bottom: 20px;
        }
        .error-message {
          background: #fff5f5;
          border: 1px solid #fed7d7;
          border-radius: 6px;
          padding: 15px;
          margin: 20px 0;
        }
        .retry-btn {
          margin-top: 20px;
          padding: 12px 24px;
          background: #ff6b6b;
          color: white;
          text-decoration: none;
          border-radius: 6px;
          display: inline-block;
          transition: background 0.3s;
        }
        .retry-btn:hover {
          background: #ff5252;
        }
      </style>
    </head>
    <body>
      <div class="error-container">
        <div class="error-icon">⚠️</div>
        <h1><?= title ?></h1>
        <div class="error-message">
          <p><?= message ?></p>
          <? if (typeof errorDetails !== 'undefined' && errorDetails) { ?>
            <details>
              <summary>詳細情報</summary>
              <pre><?= errorDetails ?></pre>
            </details>
          <? } ?>
        </div>
        <a href="<?= getWebAppUrl() ?>?mode=login" class="retry-btn">ログインページに戻る</a>
      </div>
    </body>
    </html>
  `);

  template.title = title;
  template.message = message;
  template.errorDetails = error ? error.message : '';

  return template
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.NATIVE);
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
   * ユーザー固有URLを生成（統一版）
   * @param {string} userId ユーザーID
   * @returns {Object} 各種URL
   */
  generateUserUrls(userId) {
    const baseUrl = getWebAppUrl();
    if (!baseUrl) return { error: 'Base URL not available' };

    return {
      viewUrl: `${baseUrl}?mode=view&userId=${encodeURIComponent(userId)}`,
      adminUrl: `${baseUrl}?mode=admin&userId=${encodeURIComponent(userId)}`,
      editUrl: `${baseUrl}?mode=edit&userId=${encodeURIComponent(userId)}`, // 将来実装
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
  },
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

  console.log('parseRequestParams - mode:', mode, 'setup:', setupParam);

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
    return content;
  } catch (error) {
    console.error(`include - Failed to load file: ${filename}`, error);
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
    // AdminPanel専用テンプレートを使用（Page.htmlのStudyQuestApp読み込み回避）
    const template = HtmlService.createTemplateFromFile('AdminPanel');

    // テンプレート変数を設定
    template.userInfo = userInfo;
    template.mode = mode || 'admin';
    template.isAdminPanel = true;

    return template
      .evaluate()
      .setTitle("管理パネル - Everyone's Answer Board")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('renderAdminPanel エラー:', error);
    return HtmlService.createHtmlOutput(
      `<h3>エラー</h3><p>管理パネルの表示中にエラーが発生しました: ${error.message}</p>`
    );
  }
}

/**
 * 🎯 ログイン状態の確認とユーザー登録
 * ログインボタンから呼び出される処理
 * @returns {Object} ログイン状態とメッセージ
 */
function processLoginAction() {
  try {
    // キャッシュをクリアして最新情報取得
    UserManager.clearCache();
    const currentUserEmail = UserManager.getCurrentUserEmail();

    if (!currentUserEmail) {
      return {
        success: false,
        message: '認証されていません。Googleアカウントでログインしてください。',
      };
    }

    // DB直接検索（キャッシュバイパス）
    let userInfo = DB.findUserByEmail(currentUserEmail);

    if (!userInfo) {
      // 新規ユーザー作成
      console.log('🆕 新規ユーザー作成開始');
      const newUserData = createCompleteUser(currentUserEmail);
      DB.createUser(newUserData);

      // 作成後に再度確認
      userInfo = DB.findUserByEmail(currentUserEmail);
      if (!userInfo) {
        throw new Error('ユーザー作成後の確認に失敗しました');
      }
    }

    // アクティブチェック
    if (userInfo.isActive !== true) {
      return {
        success: false,
        message: 'アカウントが無効化されています。管理者にお問い合わせください。',
      };
    }

    // 最終アクセス時刻更新
    updateUserLastAccess(userInfo.userId);

    // 管理パネルURLを生成
    const adminUrl = buildUserAdminUrl(userInfo.userId);

    return {
      success: true,
      message: 'ログインが完了しました',
      adminUrl,
      userId: userInfo.userId,
    };
  } catch (error) {
    console.error('ログインアクションエラー:', error);
    return {
      success: false,
      message: `ログイン処理中にエラーが発生しました: ${error.message}`,
    };
  }
}

/**
 * ログインページのレンダリング（内部用）
 * @param {Object} params リクエストパラメータ
 * @returns {HtmlService.HtmlOutput} HTML出力
 */
function renderLoginPage(params = {}) {
  try {
    const template = HtmlService.createTemplateFromFile('LoginPage');
    template.params = params;
    return template
      .evaluate()
      .setTitle('StudyQuest - ログイン')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('renderLoginPage error:', error);
    return HtmlService.createHtmlOutput(
      `<h3>Error</h3><p>ログインページの読み込みに失敗しました: ${error.message}</p>`
    );
  }
}

/**
 * 回答ボードのレンダリング
 * @param {Object} userInfo ユーザー情報
 * @param {Object} params リクエストパラメータ
 * @returns {HtmlService.HtmlOutput} HTML出力
 */
function renderAnswerBoard(userInfo, params) {
  // ConfigManagerから直接最新の設定を取得
  let config;
  try {
    config = ConfigManager.getUserConfig(userInfo.userId);
    if (!config || !config.spreadsheetId) {
      console.warn('⚠️ ConfigManagerからの取得失敗、parsedConfigを使用');
      config = JSON.parse(userInfo.configJson || '{}') || {};
    }
  } catch (error) {
    console.error('❌ Config取得エラー:', error);
    config = JSON.parse(userInfo.configJson || '{}') || {};
  }

  // デバッグ: 実際に取得されたconfigの中身を確認
  // config取得状況デバッグ（必要時のみ有効化）
  // console.log('config状況:', {
  //   hasSpreadsheetId: !!config.spreadsheetId,
  //   hasSheetName: !!config.sheetName,
  //   spreadsheetIdValue: config.spreadsheetId?.substring(0, 20) + '...',
  //   sheetNameValue: config.sheetName,
  //   configKeys: Object.keys(config)
  // });

  console.log('renderAnswerBoard - userId:', userInfo.userId);
  console.log('renderAnswerBoard - mode:', params.mode);

  try {
    // 既存のPage.htmlテンプレートを使用
    const template = HtmlService.createTemplateFromFile('Page');

    // 基本情報設定
    template.userInfo = userInfo;
    template.mode = 'view';
    template.isAdminPanel = false;

    // ✅ 設定取得統一化：より確実な設定参照方法
    const userSpreadsheetId = config.spreadsheetId || config.spreadsheetId || null;
    const userSheetName = config.sheetName || config.sheetName || null;

    // 📊 ユーザーIDをテンプレートに適切に設定
    template.USER_ID = userInfo.userId || null;
    template.SHEET_NAME = userSheetName || '';

    // ✅ configJSON中心型: sheetConfig廃止、直接config使用

    // シンプルな判定: ユーザーがスプレッドシートを設定済みかどうか
    const hasUserConfig = !!(userSpreadsheetId && userSheetName);
    const finalSpreadsheetId = hasUserConfig ? userSpreadsheetId : params.spreadsheetId;
    const finalSheetName = hasUserConfig ? userSheetName : params.sheetName;

    console.log('renderAnswerBoard - hasUserConfig:', hasUserConfig);
    console.log('renderAnswerBoard - finalSpreadsheetId:', finalSpreadsheetId);
    console.log('renderAnswerBoard - finalSheetName:', finalSheetName);

    // ✅ テンプレート変数設定: sheetConfig廃止
    template.config = config;
    template.spreadsheetId = finalSpreadsheetId;
    template.sheetName = finalSheetName;
    template.isDirectPageAccess = params.isDirectPageAccess;
    template.isPublished = hasUserConfig;
    template.appPublished = config.appPublished || false;

    // 🦾 スマート診断情報をテンプレートに提供
    template.DIAGNOSTIC_INFO = {
      hasSpreadsheetId: !!(finalSpreadsheetId && finalSpreadsheetId !== 'null'),
      hasSheetName: !!(finalSheetName && finalSheetName !== 'null'),
      setupStatus: config.setupStatus || 'pending',
      appPublished: config.appPublished || false,
      suggestedAction: getSuggestedAction({
        hasSpreadsheetId: !!(finalSpreadsheetId && finalSpreadsheetId !== 'null'),
        hasSheetName: !!(finalSheetName && finalSheetName !== 'null'),
        setupStatus: config.setupStatus || 'pending',
        appPublished: config.appPublished || false,
      }),
      systemHealthy: hasUserConfig && config.appPublished,
    };

    // __OPINION_HEADER__テンプレート変数を設定（高精度・確実取得版）
    let opinionHeader = 'お題'; // デフォルト値
    let opinionHeaderSource = 'default';

    try {
      // ✅ Step 1: configJsonからopinionHeaderを優先取得（「お題」以外の場合）
      if (config?.opinionHeader && config.opinionHeader !== 'お題') {
        opinionHeader = config.opinionHeader;
        opinionHeaderSource = 'configJson';
        console.log('✅ renderAnswerBoard: configJsonからopinionHeader取得:', {
          value: opinionHeader.substring(0, 50) + (opinionHeader.length > 50 ? '...' : ''),
          length: opinionHeader.length,
          source: 'configJson',
        });
      }
      // ✅ Step 2: configJsonが「お題」の場合、または未設定の場合は高精度検出実行
      else if (finalSpreadsheetId && finalSheetName) {
        // 2-1: ヘッダー行から直接検出
        try {
          const spreadsheet = new ConfigurationManager().getSpreadsheet(finalSpreadsheetId);
          const sheet = spreadsheet.getSheetByName(finalSheetName);
          const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

          // 質問らしいヘッダーを検索
          const questionHeader = headerRow.find(
            (header) =>
              header &&
              typeof header === 'string' &&
              (header.includes('どうして') ||
                header.includes('なぜ') ||
                header.includes('思いますか'))
          );

          if (questionHeader && questionHeader !== 'お題') {
            opinionHeader = questionHeader;
            opinionHeaderSource = 'header_detection';
          }
        } catch (error) {
          console.warn('ヘッダー検出エラー:', error.message);
        }

        if (opinionHeader && opinionHeader !== 'お題') {
          console.log('✅ renderAnswerBoard: 高精度検出システムによるopinionHeader取得:', {
            value: opinionHeader.substring(0, 50) + (opinionHeader.length > 50 ? '...' : ''),
            length: opinionHeader.length,
            source: 'Core.gs高精度検出システム',
          });

          // 3. 取得したopinionHeaderをconfigJsonに保存（永続化・最適化）
          if (userInfo?.userId) {
            try {
              const updatedConfig = { ...config, opinionHeader };
              ConfigManager.saveConfig(userInfo.userId, updatedConfig);
              console.log(
                '💾 renderAnswerBoard: opinionHeader永続化完了 - 次回はconfigJsonから直接取得'
              );
            } catch (saveError) {
              console.warn('⚠️ renderAnswerBoard: configJson保存エラー:', saveError.message);
            }
          }
        } else {
          console.warn('⚠️ renderAnswerBoard: 高精度検出でもopinionHeaderが「お題」:', {
            opinionHeaderValue: opinionHeader,
            opinionHeaderSource,
            detectionStatus: '高精度検出実行済み',
          });
        }
      } else {
        console.warn('⚠️ renderAnswerBoard: スプレッドシート情報不足でopinionHeader検出不可:', {
          finalSpreadsheetId: !!finalSpreadsheetId,
          finalSheetName: !!finalSheetName,
        });
      }
    } catch (headerError) {
      console.error('❌ renderAnswerBoard: opinionHeader取得エラー:', {
        error: headerError.message,
        stack: headerError.stack,
        finalSpreadsheetId: !!finalSpreadsheetId,
        finalSheetName: !!finalSheetName,
      });
      opinionHeader = 'お題'; // エラー時フォールバック
      opinionHeaderSource = 'error_fallback';
    }

    // 最終確認ログ
    console.log('📋 renderAnswerBoard: opinionHeader最終設定完了:', {
      finalValue: opinionHeader.substring(0, 50) + (opinionHeader.length > 50 ? '...' : ''),
      source: opinionHeaderSource,
      isDefault: opinionHeader === 'お題',
      templateVariableSet: true,
    });
    template.__OPINION_HEADER__ = opinionHeader;

    // StudyQuestApp用のユーザーID設定（エラー対策）
    template.USER_ID = userInfo.userId || '';
    template.SHEET_NAME = finalSheetName || '';

    // データ取得とテンプレート設定の処理
    try {
      if (finalSpreadsheetId && finalSheetName) {
        // データ取得開始
        const dataResult = getData(userInfo.userId, finalSpreadsheetId, finalSheetName);
        template.data = dataResult.data;
        template.message = dataResult.message;
        template.hasData = !!(dataResult.data && dataResult.data.length > 0);
      } else {
        template.data = [];
        template.message = '表示するデータがありません';
        template.hasData = false;
      }

      // 現在の設定から表示設定を取得（シンプル版）
      const currentConfig = JSON.parse(userInfo.configJson || '{}') || {};
      const displaySettings = currentConfig.displaySettings || {};

      // 表示設定を適用
      template.displayMode = displaySettings.showNames ? 'named' : 'anonymous';
      template.showCounts = displaySettings.showReactions !== false;

      // 表示設定完了
    } catch (dataError) {
      console.error('renderAnswerBoard - データ取得エラー:', dataError);
      template.data = [];
      template.message = `データ取得中にエラーが発生しました: ${dataError.message}`;
      template.hasData = false;

      // エラー時も表示設定を適用（シンプル版）
      const currentConfig = JSON.parse(userInfo.configJson || '{}') || {};
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

    // シンプルな所有者チェック
    try {
      const currentUserEmail = Session.getActiveUser().getEmail();
      const boardOwnerEmail = userInfo.userEmail;
      
      // 単純な所有者判定
      template.isOwner = (currentUserEmail === boardOwnerEmail);
      
      console.log('🔍 renderAnswerBoard - 所有者チェック:', {
        currentUserEmail: currentUserEmail ? `${currentUserEmail.substring(0, 10)}...` : 'null', 
        boardOwnerEmail: boardOwnerEmail ? `${boardOwnerEmail.substring(0, 10)}...` : 'null',
        isOwner: template.isOwner
      });
    } catch (error) {
      console.error('renderAnswerBoard - 所有者チェックエラー:', error);
      template.isOwner = false;
    }

    return template
      .evaluate()
      .setTitle("Everyone's Answer Board")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('renderAnswerBoard - 全般エラー:', error);
    return HtmlService.createHtmlOutput(
      `<h3>エラー</h3><p>ページの表示中にエラーが発生しました: ${error.message}</p>`
    );
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

    // 🔥 設定情報を効率的に取得（parsedConfig優先）
    const config = JSON.parse(userInfo.configJson || '{}') || {};

    // configJSON中心型：公開状態判定
    const isPublished = config.appPublished === true;

    console.log('checkCurrentPublicationStatus - result:', {
      userId,
      isPublished,
      hasSpreadsheetId: !!config.spreadsheetId,
      hasSheetName: !!config.sheetName,
    });

    return {
      userId,
      isPublished,
      spreadsheetId: config.spreadsheetId || null,
      sheetName: config.sheetName || null,
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
 * テスト用シンプルAPI（テスト専用）
 */
function getTestMockResponse() {
  return createResponse(true, 'Test mock response', { reason: 'C' });
}

/**
 * アプリケーション設定保存
 * @param {Object} config 設定オブジェクト
 * @returns {Object} 保存結果
 */
function saveApplicationConfig(config) {
  console.log('アプリケーション設定を保存:', config);
  return createResponse(true, '保存しました', { config });
}

// publishApplication: AdminPanelBackend.gsに統合済み

/**
 * システムセットアップ実行関数
 * セットアップページから呼び出される
 * @param {string} serviceAccountJson サービスアカウントJSON文字列
 * @param {string} databaseSpreadsheetId データベーススプレッドシートID
 * @param {string} adminEmail 管理者メールアドレス（省略時は現在のユーザー）
 * @param {string} googleClientId Google Client ID（オプション）
 */
function setupApplication(
  serviceAccountJson,
  databaseSpreadsheetId,
  adminEmail = null,
  googleClientId = null
) {
  try {
    console.log('setupApplication - セットアップ開始');

    // 現在のユーザーのメールアドレスを取得
    const currentUserEmail = UserManager.getCurrentUserEmail();
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
      throw new Error(
        'スプレッドシートIDの形式が無効です（44文字の英数字、アンダースコア、ハイフンのみ）'
      );
    }

    // Script Propertiesに設定を保存
    const properties = PropertiesService.getScriptProperties();

    properties.setProperties({
      [PROPS_KEYS.SERVICE_ACCOUNT_CREDS]: serviceAccountJson,
      [PROPS_KEYS.DATABASE_SPREADSHEET_ID]: databaseSpreadsheetId,
      [PROPS_KEYS.ADMIN_EMAIL]: finalAdminEmail,
    });

    // Google Client IDが提供されている場合は設定
    if (googleClientId && googleClientId.trim()) {
      properties.setProperty(PROPS_KEYS.GOOGLE_CLIENT_ID, googleClientId.trim());
    }

    console.log('setupApplication - プロパティ設定完了');

    // データベースシートの初期化とヘッダー設定
    try {
      console.log('setupApplication - データベースシート初期化開始');
      initializeDatabaseSheet(databaseSpreadsheetId);
      console.log('setupApplication - データベースシート初期化完了');
    } catch (dbError) {
      console.error('データベースシート初期化エラー:', dbError);
      throw new Error(`データベースシートの初期化に失敗しました: ${dbError.message}`);
    }

    // セットアップ完了検証
    if (!isSystemSetup()) {
      throw new Error('セットアップ後の検証に失敗しました');
    }

    console.log('setupApplication - セットアップ完了');

    return createResponse(true, 'システムセットアップが正常に完了しました', {
      adminEmail: finalAdminEmail,
    });
  } catch (error) {
    console.error('setupApplication エラー:', error);
    throw new Error(`セットアップに失敗しました: ${error.message}`);
  }
}

/**
 * ユーザー情報取得のメインAPI
 * @param {string} format - 'object'|'string'|'email' 戻り値形式
 * @returns {Object|string} formatに応じたユーザー情報
 */
function getUser(format = 'object') {
  try {
    // UnifiedManagerを使用（重複関数統合）
    const userInfo = UnifiedManager.user.getCurrentInfo();
    const email = userInfo?.userEmail || null;
    const userId = userInfo?.userId || null;

    // セッション情報を保存（1時間有効）
    if (userId) {
      const sessionData = {
        userId,
        email,
        timestamp: Date.now(),
      };
      PropertiesService.getScriptProperties().setProperty(
        'LAST_USER_SESSION',
        JSON.stringify(sessionData)
      );
    }

    // シンプルな文字列形式
    if (format === 'string' || format === 'email') {
      return email || '';
    }

    // オブジェクト形式（デフォルト）
    return createResponse(true, email ? 'ユーザー取得成功' : 'ユーザー未認証', {
      email,
      userId,
      isAuthenticated: !!email,
    });
  } catch (error) {
    console.error('getUser エラー:', error);

    if (format === 'string' || format === 'email') {
      return '';
    }

    return {
      success: false,
      email: null,
      isAuthenticated: false,
      message: `取得失敗: ${error.message}`,
    };
  }
}

/**
 * URL システムリセット関数
 * ログインページで使用
 */
function forceUrlSystemReset() {
  try {
    console.log('URL system reset requested');
    return createResponse(true, 'URL system reset completed');
  } catch (error) {
    console.error('forceUrlSystemReset error:', error);
    return createResponse(false, error.message, null, error);
  }
}

/**
 * クライアントエラー報告関数
 * ErrorBoundary.htmlで使用
 * @param {Object} errorInfo エラー情報
 */
function reportClientError(errorInfo) {
  try {
    console.error('🚨 CLIENT ERROR:', errorInfo);
    return createResponse(true, 'Error reported successfully');
  } catch (error) {
    console.error('reportClientError failed:', error);
    return createResponse(false, error.message, null, error);
  }
}

/**
 * API統一レスポンス生成
 * 全API関数で一貫したレスポンス形式を提供
 * @param {boolean} success - 成功/失敗フラグ
 * @param {string|null} message - メッセージ
 * @param {Object|null} data - データ
 * @param {Error|null} error - エラーオブジェクト
 * @returns {Object} 統一レスポンス形式
 */
function createResponse(success, message = null, data = null, error = null) {
  const response = { success };

  if (message) response.message = message;
  if (data) response.data = data;
  if (error) {
    response.error = error.message;
    response.stack = error.stack;
  }
  response.timestamp = new Date().toISOString();

  return response;
}

/**
 * スプレッドシートURL追加関数
 * Unpublished.htmlで使用
 * @param {string} url スプレッドシートURL
 */
function addSpreadsheetUrl(url) {
  try {
    console.log('addSpreadsheetUrl called with:', url);
    return createResponse(true, 'URL added successfully');
  } catch (error) {
    console.error('addSpreadsheetUrl error:', error);
    return createResponse(false, 'URL追加に失敗しました', null, error);
  }
}

/**
 * ユーザー認証リセット関数
 * SharedUtilities.html、login.js.htmlで使用
 */
function resetAuth() {
  try {
    const result = new ConfigurationManager().getCurrentUserInfoSafely();
    const userEmail = result?.currentUserEmail;
    if (userEmail) {
      console.log('ユーザー認証をリセットしました');
    }
    return createResponse(true, '認証リセット完了');
  } catch (error) {
    console.error('resetAuth エラー:', error);
    return createResponse(false, error.message, null, error);
  }
}

/**
 * 強制ログアウト・リダイレクトテスト関数
 * ErrorBoundary.htmlで使用
 */
function testForceLogoutRedirect() {
  try {
    console.log('強制ログアウト・リダイレクトテストを実行');
    return createResponse(true, 'テスト実行完了');
  } catch (error) {
    console.error('testForceLogoutRedirect エラー:', error);
    return createResponse(false, error.message, null, error);
  }
}

/**
 * ユーザー認証検証関数
 * SharedUtilities.htmlで使用
 */
function verifyUserAuthentication() {
  try {
    const result = new ConfigurationManager().getCurrentUserInfoSafely();
    const email = result?.currentUserEmail;
    const isAuthenticated = !!email;
    return createResponse(true, isAuthenticated ? '認証済み' : '未認証', {
      authenticated: isAuthenticated,
      email: email || null,
    });
  } catch (error) {
    console.error('verifyUserAuthentication エラー:', error);
    return createResponse(false, error.message, { authenticated: false }, error);
  }
}

/**
 * 強制ログアウト・ログインページリダイレクト
 * ErrorBoundary.htmlで使用
 */
function forceLogoutAndRedirectToLogin() {
  try {
    console.log('forceLogoutAndRedirectToLogin called');
    return createResponse(true, 'Logout successful', {
      redirectUrl: `${getWebAppUrl()}?mode=login`,
    });
  } catch (error) {
    console.error('forceLogoutAndRedirectToLogin error:', error);
    return createResponse(false, error.message, null, error);
  }
}

// getSpreadsheetList: AdminPanelBackend.gsに統合済み

/**
 * 公開シートデータ取得
 * page.js.htmlで使用
 * @param {Object} params 取得パラメータ
 */
function getData(userId, classFilter, sortOrder, adminMode, bypassCache) {
  try {
    // CLAUDE.md準拠: 個別引数とオブジェクト引数の両方に対応
    let params;
    if (arguments.length === 1 && typeof userId === 'object' && userId !== null) {
      // オブジェクト引数の場合
      params = userId;
    } else {
      // 個別引数の場合
      params = { userId, classFilter, sortOrder, adminMode, bypassCache };
    }

    console.log('getData: Core.gs実装呼び出し開始', {
      argumentsLength: arguments.length,
      firstArgType: typeof userId,
      params,
    });

    // CLAUDE.md準拠: 統一データソース原則でCore.gs実装を呼び出し
    // 多段階認証フロー: ユーザーIDの確実な取得
    let targetUserId = params.userId;
    let currentUserEmail = null;

    // Step 1: 直接指定されたuserIdを確認
    if (targetUserId && typeof targetUserId === 'string' && targetUserId.trim()) {
      console.log('getData: 指定userIdを使用', targetUserId);
    } else {
      // Step 2: User.email()からの取得を試行
      try {
        currentUserEmail = UserManager.getCurrentEmail();
        console.log('getData: UserManager.getCurrentEmail()結果', currentUserEmail);
      } catch (emailError) {
        console.warn('getData: UserManager.getCurrentEmail()取得失敗', emailError.message);
      }

      // Step 3: メールアドレスからユーザー検索
      if (currentUserEmail && typeof currentUserEmail === 'string' && currentUserEmail.trim()) {
        try {
          const user = DB.findUserByEmail(currentUserEmail);
          targetUserId = user ? user.userId : null;
          console.log('getData: DB検索結果', {
            email: currentUserEmail,
            userId: targetUserId,
          });
        } catch (dbError) {
          console.warn('getData: DB検索失敗', dbError.message);
        }
      }

      // Step 4: セッション情報からの復元試行
      if (!targetUserId) {
        try {
          const sessionData =
            PropertiesService.getScriptProperties().getProperty('LAST_USER_SESSION');
          if (sessionData) {
            const session = JSON.parse(sessionData);
            if (session.userId && Date.now() - session.timestamp < 3600000) {
              // 1時間以内
              targetUserId = session.userId;
              console.log('getData: セッション復元成功', targetUserId);
            }
          }
        } catch (sessionError) {
          console.warn('getData: セッション復元失敗', sessionError.message);
        }
      }
    }

    if (!targetUserId) {
      const errorDetails = {
        providedUserId: params.userId,
        currentUserEmail,
        hasValidEmail: !!currentUserEmail,
        timestamp: new Date().toISOString(),
      };
      console.error('getData: 全認証方法失敗', errorDetails);
      throw new Error(`ユーザー認証情報が取得できません。詳細: ${JSON.stringify(errorDetails)}`);
    }

    return getPublishedSheetData(
      targetUserId,
      params.classFilter,
      params.sortOrder,
      params.adminMode
    );
  } catch (error) {
    console.error('getData エラー:', error);
    return createResponse(false, error.message, { data: [] }, error);
  }
}

/**
 * セットアップテスト実行
 */
// 注意: testSetup関数とtestDatabaseMigration関数は削除されました
// 必要な場合は、AdminPanelBackend.gsのexecuteConfigCleanup()を使用してください

/**
 * 🔧 統一エラーハンドリング関数
 */
function handleSystemError(context, error, userId = null, additionalData = {}) {
  const errorInfo = {
    context,
    message: error.message,
    stack: error.stack,
    userId,
    timestamp: new Date().toISOString(),
    ...additionalData,
  };

  console.error(`❌ [${context}] システムエラー:`, errorInfo);

  return {
    success: false,
    error: error.message,
    context,
    timestamp: errorInfo.timestamp,
  };
}

/**
 * 🔧 グローバル関数: getActiveUserInfo（Core.gs互換性用）
 * UnifiedManagerの軽量ラッパー - 完全統合版
 */
function getActiveUserInfo() {
  try {
    // UnifiedManagerを直接使用（重複削除）
    const userInfo = UnifiedManager.user.getCurrentInfo();
    if (!userInfo) return null;

    const config = userInfo.parsedConfig || {};
    return {
      email: userInfo.userEmail,
      userId: userInfo.userId,
      userEmail: userInfo.userEmail,
      spreadsheetId: config.spreadsheetId,
      configJson: userInfo.configJson,
      parsedConfig: config,
    };
  } catch (error) {
    console.error('getActiveUserInfo エラー:', error.message);
    return null;
  }
}

/**
 * 🔧 データ状態一元検証関数
 */
function validateUserDataState(userInfo) {
  const issues = [];
  const fixes = [];

  // 基本フィールド検証
  if (!userInfo.userId) {
    issues.push('userIdが設定されていません');
    fixes.push('ユーザーIDを設定');
  }

  if (!userInfo.userEmail) {
    issues.push('userEmailが設定されていません');
    fixes.push('メールアドレスを設定');
  }

  if (userInfo.isActive !== true && userInfo.isActive !== false) {
    issues.push('isActiveフラグが不正です');
    fixes.push('isActiveをbooleanに設定');
  }

  // 設定データ検証
  const config = JSON.parse(userInfo.configJson || '{}') || {};
  if (config.appPublished && !config.spreadsheetId) {
    issues.push('公開状態だがスプレッドシートIDがありません');
    fixes.push('スプレッドシート接続を確認');
  }

  if (config.isDraft === true && config.appPublished === true) {
    issues.push('ドラフト状態と公開状態が同時にtrueです');
    fixes.push('状態フラグを整理');
  }

  return {
    isValid: issues.length === 0,
    issues,
    fixes,
    summary: issues.length > 0 ? `${issues.length}個の問題を検出` : '状態正常',
  };
}

/**
 * 🛠️ GAS実行用診断・復旧関数群
 * エラー時にGASエディタから直接実行可能
 */

/**
 * システム診断関数（GAS実行用）
 */
function diagnoseSystem() {
  try {
    const result = new ConfigurationManager().getCurrentUserInfoSafely();
    const currentUser = result?.currentUserEmail;
    const userInfo = result?.userInfo;

    const diagnosis = {
      timestamp: new Date().toISOString(),
      currentUser,
      userExists: !!userInfo,
      userData: userInfo
        ? {
            userId: userInfo.userId,
            isActive: userInfo.isActive,
            hasConfig: !!userInfo.configJson,
            setupStatus: JSON.parse(userInfo.configJson || '{}').setupStatus,
          }
        : null,
      systemSetup: isSystemSetup(),
      recommendations: [],
    };

    if (!currentUser) {
      diagnosis.recommendations.push('ユーザーが認証されていません');
    }
    if (!userInfo) {
      diagnosis.recommendations.push('ユーザーをDBに登録してください');
    }
    if (userInfo && !userInfo.isActive) {
      diagnosis.recommendations.push('ユーザーを有効化してください');
    }

    console.log('✅ システム診断結果:', diagnosis);
    return diagnosis;
  } catch (error) {
    console.error('❌ システム診断エラー:', error.message);
    return { error: error.message };
  }
}

/**
 * 緊急キャッシュクリア関数（GAS実行用）
 */
function emergencyClearCache() {
  try {
    console.log('🚨 緊急キャッシュクリア実行...');
    SimpleCacheManager.clearAll();
    console.log('✅ 緊急キャッシュクリア完了');
    return { success: true, message: 'キャッシュクリア完了' };
  } catch (error) {
    console.error('❌ 緊急キャッシュクリアエラー:', error.message);
    return { success: false, error: error.message };
  }
}

/**
 * 現在のユーザー修復関数（GAS実行用）
 */
function repairCurrentUser() {
  try {
    console.log('🔧 現在のユーザー修復開始...');

    const result = new ConfigurationManager().getCurrentUserInfoSafely();
    const currentUser = result?.currentUserEmail;
    if (!currentUser) {
      throw new Error('認証されたユーザーが見つかりません');
    }

    let userInfo = DB.findUserByEmail(currentUser);
    if (!userInfo) {
      // ユーザーが存在しない場合は新規作成
      console.log('新規ユーザー作成中...');
      const newUserData = createCompleteUser(currentUser);
      DB.createUser(newUserData);
      userInfo = DB.findUserByEmail(currentUser);
    }

    // データ整合性チェック・修復
    if (userInfo.isActive !== true) {
      console.log('isActiveフラグ修復中...');
      DB.updateUser(userInfo.userId, { isActive: true });
    }

    console.log('✅ ユーザー修復完了:', {
      userId: userInfo.userId,
      userEmail: userInfo.userEmail,
      isActive: userInfo.isActive,
    });

    return { success: true, userInfo };
  } catch (error) {
    console.error('❌ ユーザー修復エラー:', error.message);
    return { success: false, error: error.message };
  }
}

/**
 * doPost - HTTP POST request handler
 * セキュアなPOST処理とAPI操作
 */
function doPost(e) {
  try {
    // 認証確認（必須）
    const userEmail = UserService.getCurrentEmail();
    if (!userEmail) {
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          message: '認証が必要です',
          code: 'AUTHENTICATION_REQUIRED'
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // アクセス権限検証
    const accessResult = SecurityService.validateWebAppAccess(userEmail, {
      method: 'POST',
      action: e.parameter?.action || 'unknown'
    });
    
    if (!accessResult.hasAccess) {
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          message: 'アクセスが拒否されました',
          reason: accessResult.reason,
          code: 'ACCESS_DENIED'
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // リクエストデータ検証とサニタイゼーション
    const validatedData = SecurityService.validateUserData(e.parameter || {});
    if (!validatedData.isValid) {
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          message: '入力データが無効です',
          errors: validatedData.errors,
          code: 'VALIDATION_ERROR'
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const params = validatedData.sanitizedData;
    
    // アクション別処理
    switch (params.action) {
      case 'add_reaction':
        return handleAddReactionAPI(params, userEmail);
      
      case 'toggle_highlight':
        return handleToggleHighlightAPI(params, userEmail);
      
      case 'update_config':
        return handleUpdateConfigAPI(params, userEmail);
      
      case 'refresh_data':
        return handleRefreshDataAPI(params, userEmail);
      
      default:
        return ContentService
          .createTextOutput(JSON.stringify({
            success: false,
            message: `無効なアクション: ${params.action}`,
            code: 'INVALID_ACTION'
          }))
          .setMimeType(ContentService.MimeType.JSON);
    }

  } catch (error) {
    // 統一エラーハンドリング
    const errorResponse = ErrorHandler.handle(error, {
      service: 'main',
      method: 'doPost',
      parameters: e?.parameter
    });
    
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: errorResponse.message,
        errorId: errorResponse.errorId,
        code: 'INTERNAL_ERROR'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * リアクション追加API（POST）
 * @param {Object} params - パラメータ
 * @param {string} userEmail - ユーザーメール
 * @returns {ContentService} JSON レスポンス
 */
function handleAddReactionAPI(params, userEmail) {
  const { userId, rowId, reactionType } = params;
  
  if (!userId || !rowId || !reactionType) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: '必須パラメータが不足しています',
        code: 'MISSING_PARAMETERS'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ConfigServiceからユーザー設定取得
  const config = ConfigService.getUserConfig(userId);
  if (!config) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: 'ユーザー設定が見つかりません',
        code: 'CONFIG_NOT_FOUND'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 行IDから数値に変換 (row_3 → 3)
  const rowNumber = parseInt(rowId.replace('row_', ''));
  if (isNaN(rowNumber) || rowNumber < 2) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: '無効な行IDです',
        code: 'INVALID_ROW_ID'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // DataServiceの新しいprocessReaction関数を使用
  const result = DataService.processReaction(
    config.spreadsheetId,
    config.sheetName,
    rowNumber,
    reactionType,
    userEmail
  );
  
  return ContentService
    .createTextOutput(JSON.stringify({
      success: result.status === 'success',
      message: result.message,
      newValue: result.newValue
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * ハイライト切り替えAPI（POST）
 * @param {Object} params - パラメータ
 * @param {string} userEmail - ユーザーメール
 * @returns {ContentService} JSON レスポンス
 */
function handleToggleHighlightAPI(params, userEmail) {
  const { userId, rowId } = params;
  
  if (!userId || !rowId) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: '必須パラメータが不足しています',
        code: 'MISSING_PARAMETERS'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ConfigServiceからユーザー設定取得
  const config = ConfigService.getUserConfig(userId);
  if (!config) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: 'ユーザー設定が見つかりません',
        code: 'CONFIG_NOT_FOUND'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 行IDから数値に変換 (row_3 → 3)
  const rowNumber = parseInt(rowId.replace('row_', ''));
  if (isNaN(rowNumber) || rowNumber < 2) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: '無効な行IDです',
        code: 'INVALID_ROW_ID'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // DataServiceの新しいprocessHighlightToggle関数を使用
  const result = DataService.processHighlightToggle(
    config.spreadsheetId,
    config.sheetName,
    rowNumber
  );
  
  return ContentService
    .createTextOutput(JSON.stringify({
      success: result.status === 'success',
      message: result.message,
      highlighted: result.highlighted
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 設定更新API（POST）
 * @param {Object} params - パラメータ
 * @param {string} userEmail - ユーザーメール
 * @returns {ContentService} JSON レスポンス
 */
function handleUpdateConfigAPI(params, userEmail) {
  const { userId, configUpdates } = params;
  
  if (!userId || !configUpdates) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: '必須パラメータが不足しています',
        code: 'MISSING_PARAMETERS'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 権限確認（設定変更は所有者のみ）
  const permission = SecurityService.checkUserPermission(userId, 'owner');
  if (!permission.hasPermission) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: '設定変更権限がありません',
        code: 'INSUFFICIENT_PERMISSION'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    const parsedUpdates = JSON.parse(configUpdates);
    const result = ConfigService.updateUserConfig(userId, parsedUpdates);
    
    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        data: result,
        message: '設定を更新しました'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: '設定の更新に失敗しました',
        error: error.message,
        code: 'UPDATE_ERROR'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * データ更新API（POST）
 * @param {Object} params - パラメータ
 * @param {string} userEmail - ユーザーメール
 * @returns {ContentService} JSON レスポンス
 */
function handleRefreshDataAPI(params, userEmail) {
  const { userId } = params;
  
  if (!userId) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: 'ユーザーIDが必要です',
        code: 'MISSING_USER_ID'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const result = DataService.refreshBoardData(userId);
  
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
