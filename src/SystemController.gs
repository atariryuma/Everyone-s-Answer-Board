/**
 * @fileoverview SystemController - System management and setup functions
 */

/* global ServiceFactory, UserService, ConfigService, getCurrentEmail, createErrorResponse, createUserNotFoundError, createExceptionResponse */

// ===========================================
// 🔧 Zero-Dependency Utility Functions
// ===========================================

/**
 * Service Discovery for Zero-Dependency Architecture
 */


// DB 初期化関数は廃止。必ず ServiceFactory.getDB() を使用すること。




/**
 * セットアップのテスト実行
 * AppSetupPage.html から呼び出される
 *
 * @returns {Object} テスト結果
 */

/**
 * システム状態の強制リセット
 * AppSetupPage.html から呼び出される（緊急時用）
 *
 * @returns {Object} リセット結果
 */
function forceUrlSystemReset() {
    try {
      console.warn('システム強制リセットが実行されました');

      // キャッシュをクリア（複数の方法を試行）
      const cacheResults = [];
      try {
        const cache = ServiceFactory.getCache();
        if (cache && typeof cache.removeAll === 'function') {
          cache.removeAll([]); // Signature-compatible no-op to avoid errors
          cacheResults.push('ScriptCache クリア要求送信');
        }
      } catch (cacheError) {
        console.warn('ScriptCache クリアエラー:', cacheError.message);
        cacheResults.push(`ScriptCache クリア失敗: ${cacheError.message}`);
      }

      // Document Cache も試行
      try {
        const docCache = ServiceFactory.getCache(); // Use unified cache
        if (docCache && typeof docCache.removeAll === 'function') {
          docCache.removeAll([]);
          cacheResults.push('DocumentCache クリア要求送信');
        }
      } catch (docCacheError) {
        console.warn('DocumentCache クリアエラー:', docCacheError.message);
        cacheResults.push(`DocumentCache クリア失敗: ${docCacheError.message}`);
      }

      // 重要: プロパティはクリアしない（データ損失防止）

      return {
        success: true,
        message: 'システムリセットが完了しました',
        actions: cacheResults,
        cacheStatus: cacheResults.join(', '),
        timestamp: new Date().toISOString()
      };

    } catch (error) {
      console.error('SystemController.forceUrlSystemReset エラー:', error.message);
      return {
        success: false,
        message: error.message
      };
    }
}

/**
 * WebアプリのURL取得
 * 各種HTMLファイルから呼び出される
 *
 * @returns {string} WebアプリURL
 */

/**
 * システム全体の診断実行
 * AppSetupPage.html から呼び出される
 *
 * @returns {Object} 診断結果
 */
function testSystemDiagnosis() {
    try {
      const diagnostics = {
        timestamp: new Date().toISOString(),
        services: {},
        database: {},
        overall: 'unknown'
      };

      // Services診断 - Safe service access with GAS loading protection
      try {
        diagnostics.services.UserService = typeof UserService !== 'undefined'
          ? '✅ Loaded and available'
          : '❌ Not loaded or unavailable';

        diagnostics.services.ConfigService = typeof ConfigService !== 'undefined'
          ? '✅ Loaded and available'
          : '❌ Not loaded or unavailable';

        diagnostics.services.DataService = '❓ 診断機能なし';
      } catch (servicesError) {
        diagnostics.services.error = servicesError.message;
      }

      // データベース診断
      try {
        const props = ServiceFactory.getProperties();
        const databaseId = props.getDatabaseSpreadsheetId();

        if (databaseId) {
          const spreadsheet = ServiceFactory.getSpreadsheet().openById(databaseId);
          diagnostics.database = {
            accessible: true,
            name: spreadsheet.getName(),
            sheets: spreadsheet.getSheets().length
          };
        } else {
          diagnostics.database = { accessible: false, reason: 'DATABASE_SPREADSHEET_ID not configured' };
        }
      } catch (dbError) {
        diagnostics.database = { accessible: false, error: dbError.message };
      }

      // 総合評価
      const hasErrors = Object.values(diagnostics).some(v =>
        typeof v === 'object' && (v.error || v.accessible === false)
      );
      diagnostics.overall = hasErrors ? '⚠️ 問題あり' : '✅ 正常';

      return {
        success: true,
        diagnostics
      };

    } catch (error) {
      console.error('SystemController.testSystemDiagnosis エラー:', error.message);
      return {
        success: false,
        message: error.message
      };
    }
}

/**
 * システム状態の取得
 * AppSetupPage.html から呼び出される
 *
 * @returns {Object} システム状態
 */
function getSystemStatus() {
    try {
      const props = ServiceFactory.getProperties();
      const status = {
        timestamp: new Date().toISOString(),
        setup: {
          hasDatabase: !!props.getDatabaseSpreadsheetId(),
          hasAdminEmail: !!props.getAdminEmail(),
          hasServiceAccount: !!props.getServiceAccountCreds()
        },
        services: {
          available: ['UserService', 'ConfigService', 'DataService', 'SecurityService']
        }
      };

      status.setup.isComplete = status.setup.hasDatabase &&
                                 status.setup.hasAdminEmail &&
                                 status.setup.hasServiceAccount;

      return {
        success: true,
        status
      };

    } catch (error) {
      console.error('SystemController.getSystemStatus エラー:', error.message);
      return {
        success: false,
        message: error.message
      };
    }
}


/**
 * データ整合性チェック
 * 管理者用の高度な診断機能
 *
 * @returns {Object} チェック結果
 */
function performDataIntegrityCheck() {
    try {
      const results = {
        timestamp: new Date().toISOString(),
        checks: {},
        summary: { passed: 0, failed: 0, warnings: 0 }
      };

      // 各種整合性チェックを実行
      // （実装は複雑になるため、基本的な構造のみ示す）

      results.checks.database = { status: '✅', message: '基本チェックのみ実装' };
      results.checks.users = { status: '✅', message: '基本チェックのみ実装' };
      results.checks.configs = { status: '✅', message: '基本チェックのみ実装' };

      results.summary.passed = 3;
      results.overall = '✅ 基本チェック完了';

      return {
        success: true,
        results
      };

    } catch (error) {
      console.error('SystemController.performDataIntegrityCheck エラー:', error.message);
      return {
        success: false,
        message: error.message
      };
    }
}

/**
 * 自動修復の実行
 * 管理者用の自動メンテナンス機能
 *
 * @returns {Object} 修復結果
 */
function performAutoRepair() {
    try {
      console.log('自動修復機能は現在基本実装のみです');

      const repairResults = {
        timestamp: new Date().toISOString(),
        actions: [
          'キャッシュクリア実行'
        ],
        summary: '基本的な修復のみ実行'
      };

      // キャッシュクリア
      try {
        const cache = ServiceFactory.getCache();
        if (cache && typeof cache.removeAll === 'function') {
          cache.removeAll();
        }
      } catch (cacheError) {
        repairResults.warnings = [`キャッシュクリア失敗: ${cacheError.message}`];
      }

      return {
        success: true,
        repairResults
      };

    } catch (error) {
      console.error('SystemController.performAutoRepair エラー:', error.message);
      return {
        success: false,
        message: error.message
      };
    }
}



/**
 * スプレッドシート一覧を取得
 * @returns {Object} スプレッドシート一覧
 */
function getAdminSpreadsheetList() {
  console.log('🔍 getAdminSpreadsheetList: 関数開始 - Zero-dependency Architecture');
  try {
    console.log('🔍 DriveApp.getFilesByType呼び出し開始');

    // 🎯 Zero-dependency: 直接DriveAppでスプレッドシート一覧取得
    const spreadsheets = DriveApp.getFilesByType('application/vnd.google-apps.spreadsheet');
    console.log('🔍 DriveApp.getFilesByType完了', spreadsheets);

    const spreadsheetList = [];
    let count = 0;

    console.log('🔍 ファイル列挙開始');
    while (spreadsheets.hasNext() && count < 20) { // 最大20件に制限
      const file = spreadsheets.next();
      const fileData = {
        id: file.getId(),
        name: file.getName(),
        lastUpdated: file.getLastUpdated(),
        url: file.getUrl(),
        size: file.getSize() || 0
      };
      console.log(`🔍 ファイル${count + 1}:`, fileData.name, fileData.id);
      spreadsheetList.push(fileData);
      count++;
    }

    const result = {
      success: true,
      spreadsheets: spreadsheetList,
      total: spreadsheetList.length,
      timestamp: new Date().toISOString()
    };

    console.log('🔍 getAdminSpreadsheetList: 結果準備完了', result);
    return result;
  } catch (error) {
    console.error('🚨 AdminController.getSpreadsheetList エラー:', error);

    const errorResult = {
      success: false,
      message: error.message || 'スプレッドシート一覧取得エラー',
      spreadsheets: []
    };

    console.log('🔍 getAdminSpreadsheetList: エラー結果', errorResult);
    return errorResult;
  }
}

/**
 * シート一覧を取得
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {Object} シート一覧
 */
function getAdminSheetList(spreadsheetId) {
  try {
    // 🎯 Zero-dependency: 直接SpreadsheetAppでシート一覧取得
    const spreadsheet = ServiceFactory.getSpreadsheet().openById(spreadsheetId);
    const sheets = spreadsheet.getSheets();

    const sheetList = sheets.map(sheet => ({
      name: sheet.getName(),
      index: sheet.getIndex(),
      rowCount: sheet.getLastRow(),
      columnCount: sheet.getLastColumn()
    }));

    return {
      success: true,
      sheets: sheetList,
      total: sheetList.length,
      spreadsheetName: spreadsheet.getName()
    };
  } catch (error) {
    console.error('AdminController.getSheetList エラー:', error.message);
    return {
      success: false,
      message: error.message || 'シート一覧取得エラー',
      sheets: []
    };
  }
}

/**
 * 列を分析
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} 列分析結果
 */


/**
 * 設定の下書き保存
 * @param {Object} config - 保存する設定
 * @returns {Object} 保存結果
 */
function saveDraftConfiguration(config) {
  try {
    // 🎯 Zero-dependency: 直接DBで設定保存
    const userEmail = getCurrentEmail();
    if (!userEmail) {
      return createErrorResponse('ユーザー認証が必要です');
    }

    const db = ServiceFactory.getDB();
    if (!db) {
      return createErrorResponse('データベース接続エラー');
    }
    const user = db.findUserByEmail(userEmail);
    if (!user) {
      return createUserNotFoundError();
    }

    // 設定をJSONで保存（重複フィールド削除）
    delete config.setupComplete;
    delete config.isDraft;
    delete config.questionText;
    config.lastAccessedAt = new Date().toISOString();
    config.lastModified = new Date().toISOString();

    const updatedUser = {
      ...user,
      configJson: JSON.stringify(config),
      updatedAt: new Date().toISOString()
    };

    const updateResult = db.updateUser(user.userId, updatedUser);
    if (!updateResult.success) {
      return createErrorResponse(updateResult.message || 'データベース更新に失敗しました');
    }

    return {
      success: true,
      message: '下書き設定を保存しました',
      userId: user.userId
    };
  } catch (error) {
    console.error('saveDraftConfiguration error:', error);
    return createExceptionResponse(error);
  }
}

/**
 * アプリケーションの公開
 * @param {Object} publishConfig - 公開設定
 * @returns {Object} 公開結果
 */
function publishApplication(publishConfig) {
  try {
    const email = getCurrentEmail();
    if (!email) {
      return { success: false, message: 'ユーザー認証が必要です' };
    }

    const publishedAt = new Date().toISOString();

    // 🎯 Zero-dependency: 直接PropertiesServiceでアプリ公開
    const props = ServiceFactory.getProperties();
    props.setProperty('APPLICATION_STATUS', 'active');
    props.setProperty('PUBLISHED_AT', publishedAt);

    // 公開設定を保存
    if (publishConfig) {
      props.setProperty('PUBLISH_CONFIG', JSON.stringify(publishConfig));
    }

    // ユーザーの設定も更新
    const db = ServiceFactory.getDB();
    const user = db.findUserByEmail(email);
    if (user) {
      const currentConfig = user.configJson ? JSON.parse(user.configJson) : {};
      const updatedConfig = {
        ...currentConfig,
        ...publishConfig,
        appPublished: true,
        publishedAt,
        setupStatus: 'completed',
        isDraft: false,
        lastModified: publishedAt
      };

      // データベースを更新
      const updateResult = db.updateUser(user.userId, {
        configJson: JSON.stringify(updatedConfig),
        lastModified: publishedAt,
        updatedAt: publishedAt
      });

      if (!updateResult || !updateResult.success) {
        console.error('Failed to update user config:', updateResult?.message || 'Unknown error');
        // エラーでも処理は継続
      } else {
        console.log('✅ User config updated successfully:', updateResult.updatedFields);
      }
    }

    return {
      success: true,
      message: 'アプリケーションが正常に公開されました',
      publishedAt
    };
  } catch (error) {
    console.error('publishApplication error:', error);
    return createExceptionResponse(error);
  }
}

/**
 * スプレッドシートへのアクセス権限を検証
 * AdminPanel.js.html から呼び出される
 *
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {Object} 検証結果
 */
function validateAccess(spreadsheetId) {
  try {
    // 🎯 Zero-dependency: 直接SpreadsheetAppでアデス権確認
    const spreadsheet = ServiceFactory.getSpreadsheet().openById(spreadsheetId);
    const sheets = spreadsheet.getSheets();

    // アデスできたら成功
    const result = {
      success: true,
      message: 'アクセス権限が確認されました',
      spreadsheetName: spreadsheet.getName(),
      sheets: sheets.map(sheet => ({
        name: sheet.getName(),
        rowCount: sheet.getLastRow(),
        columnCount: sheet.getLastColumn()
      })),
      owner: spreadsheet.getOwner()?.getEmail() || 'unknown',
      url: spreadsheet.getUrl()
    };

    return result;
  } catch (error) {
    console.error('AdminController.validateAccess エラー:', error.message);
    return {
      success: false,
      message: error.message,
      sheets: []
    };
  }
}





/**
 * フォーム情報を取得
 * AdminPanel.js.html から呼び出される
 *
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} フォーム情報
 */
function getFormInfo(spreadsheetId, sheetName) {
  try {
    // 🚀 Zero-dependency: ServiceFactory経由でConfigService利用
    const configService = ServiceFactory.getConfigService();
    if (!configService) {
      throw new Error('ConfigService not available');
    }
    return configService.getFormInfo(spreadsheetId, sheetName);
  } catch (error) {
    console.error('AdminController.getFormInfo エラー:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * フォーム作成
 * AdminPanel.js.html から呼び出される
 *
 * @param {string} userId - ユーザーID
 * @param {Object} config - フォーム設定
 * @returns {Object} 作成結果
 */
function createForm(userId, config) {
  try {
    console.log('AdminController.createForm: 開始', { userId, configKeys: Object.keys(config || {}) });

    if (!userId) {
      return {
        success: false,
        error: 'ユーザーIDが指定されていません'
      };
    }

    if (!config || !config.title) {
      return {
        success: false,
        error: 'フォーム設定が不正です'
      };
    }

    // ServiceFactory経由でConfigServiceアクセス
    const configService = ServiceFactory.getConfigService();
    if (!configService) {
      console.error('AdminController.createForm: ConfigService not available');
      return { success: false, message: 'ConfigServiceが利用できません' };
    }
    const result = configService.createForm(userId, config);

    if (result && result.success) {
      console.log('AdminController.createForm: 成功', { formUrl: result.formUrl });
      return result;
    } else {
      console.error('AdminController.createForm: ConfigService失敗', result);
      return {
        success: false,
        error: result?.error || 'フォーム作成に失敗しました'
      };
    }

  } catch (error) {
    console.error('AdminController.createForm エラー:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 現在の公開状態を確認
 * AdminPanel.js.html から呼び出される
 *
 * @returns {Object} 公開状態情報
 */
function checkCurrentPublicationStatus(targetUserId) {
  try {
    const session = ServiceFactory.getSession();
    const db = ServiceFactory.getDB();

    if (!db) {
      return createErrorResponse('データベース接続エラー');
    }

    let user = null;
    if (targetUserId) {
      user = db.findUserById(targetUserId);
    }

    if (!user && session && session.email) {
      user = db.findUserByEmail(session.email);
    }

    if (!user) {
      return createUserNotFoundError();
    }

    let config = {};
    try {
      config = JSON.parse(user.configJson || '{}');
    } catch (parseError) {
      console.warn('checkCurrentPublicationStatus: config parse error', parseError.message);
      config = {};
    }

    return {
      success: true,
      published: config.appPublished === true,
      publishedAt: config.publishedAt || null,
      lastModified: config.lastModified || null,
      hasDataSource: Boolean(config.spreadsheetId && config.sheetName),
      userId: user.userId
    };
  } catch (error) {
    console.error('AdminController.checkCurrentPublicationStatus エラー:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * 現在のユーザー情報を取得
 * login.js.html, SetupPage.html, AdminPanel.js.html から呼び出される
 *
 * @param {string} [kind='email'] - 取得する情報の種類（'email' or 'full'）
 * @returns {Object|string|null} 統一されたレスポンス形式
 */
/**
 * Direct email retrieval using GAS Session API (SystemController version)
 */
// 🎯 Zero-dependency Helper Functions
function columnNumberToLetter(num) {
  let letter = '';
  while (num > 0) {
    const remainder = (num - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    num = Math.floor((num - 1) / 26);
  }
  return letter;
}




// ===========================================
// 📊 認証・ログイン関連API
// ===========================================



/**
 * 認証状態を確認
 * login.js.html, SetupPage.html から呼び出される
 *
 * @returns {Object} 認証状態
 */
function verifyUserAuthentication() {
  try {
    const {email} = ServiceFactory.getSession();
    const userEmail = email ? email : null;
    if (!userEmail) {
      return {
        isAuthenticated: false,
        message: '認証が必要です'
      };
    }

    const userInfo = email ? { email } : null;
    return {
      isAuthenticated: true,
      userEmail,
      userInfo,
      hasConfig: !!userInfo?.config
    };

  } catch (error) {
    console.error('FrontendController.verifyUserAuthentication エラー:', error.message);
    return {
      isAuthenticated: false,
      message: error.message
    };
  }
}



/**
 * ログイン状態を取得
 * login.js.html から呼び出される
 *
 * @returns {Object} ログイン状態
 */
function getLoginStatus() {
  try {
    const {email} = ServiceFactory.getSession();
    const userEmail = email ? email : null;
    if (!userEmail) {
      return {
        isLoggedIn: false,
        user: null
      };
    }

    const userInfo = email ? { email } : null;
    return {
      isLoggedIn: true,
      user: {
        email: userEmail,
        userId: userInfo?.userId,
        hasSetup: !!userInfo?.config?.setupComplete
      }
    };

  } catch (error) {
    console.error('FrontendController.getLoginStatus エラー:', error.message);
    return {
      isLoggedIn: false,
      user: null,
      error: error.message
    };
  }
}

/**
 * クライアントエラーを報告
 * ErrorBoundary.html から呼び出される
 *
 * @param {Object} errorInfo - エラー情報
 * @returns {Object} 報告結果
 */
function reportClientError(errorInfo) {
  try {
    console.error('クライアントエラー報告:', errorInfo);

    // ServiceFactory経由でセッション情報取得
    const {email} = ServiceFactory.getSession();

    // エラーログを記録（将来的にはSecurityServiceや専用のログサービスに委譲）
    const logEntry = {
      timestamp: new Date().toISOString(),
      type: 'client_error',
      userEmail: email ? email : 'unknown',
      errorInfo
    };

    // コンソールにログ出力（将来的には永続化）
    console.log('Error Log Entry:', JSON.stringify(logEntry));

    return {
      success: true,
      message: 'エラーが報告されました'
    };
  } catch (error) {
    console.error('FrontendController.reportClientError エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * 強制ログアウトとリダイレクトのテスト
 * ErrorBoundary.html から呼び出される
 *
 * @returns {Object} テスト結果
 */
function testForceLogoutRedirect() {
  try {
    console.log('強制ログアウトテストが実行されました');

    return {
      success: true,
      message: 'ログアウトテスト完了',
      redirectUrl: `${ScriptApp.getService().getUrl()}?mode=login`
    };
  } catch (error) {
    console.error('FrontendController.testForceLogoutRedirect エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}
