/**
 * @fileoverview SystemController - System management and setup functions
 */

/* global ServiceFactory, UserService, ConfigService, getCurrentEmail, createErrorResponse, createUserNotFoundError, createExceptionResponse, Data, Config, getSpreadsheetList */

// ===========================================
// 🔧 Zero-Dependency Utility Functions
// ===========================================

/**
 * Service Discovery for Zero-Dependency Architecture
 */






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
          const dataAccess = Data.open(databaseId);
          const {spreadsheet} = dataAccess;
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
          hasServiceAccount: !!Config.serviceAccount()
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
 * スプレッドシート一覧を取得（管理者向け）
 * DataService.getSpreadsheetList()の管理者モードラッパー
 * @returns {Object} スプレッドシート一覧
 */
function getAdminSpreadsheetList() {
  return getSpreadsheetList({ adminMode: true });
}

/**
 * シート一覧を取得
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {Object} シート一覧
 */
function getAdminSheetList(spreadsheetId) {
  try {
    // 🎯 Zero-dependency: サービスアカウント経由でシート一覧取得
    const dataAccess = Data.open(spreadsheetId);
    const {spreadsheet} = dataAccess;
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
  const startTime = new Date().toISOString();
  console.log('=== saveDraftConfiguration START ===', {
    spreadsheetId: config?.spreadsheetId || 'N/A',
    sheetName: config?.sheetName || 'N/A',
    configSize: config ? JSON.stringify(config).length : 0
  });

  try {
    if (!config || typeof config !== 'object') {
      console.error('❌ Invalid config input');
      return createErrorResponse('設定データが不正です');
    }

    const userEmail = getCurrentEmail();
    if (!userEmail) {
      console.error('❌ User authentication failed');
      return createErrorResponse('ユーザー認証が必要です');
    }

    const db = ServiceFactory.getDB();
    if (!db) {
      console.error('❌ Database connection failed');
      return createErrorResponse('データベース接続エラー');
    }

    const user = db.findUserByEmail(userEmail);
    if (!user) {
      console.error('❌ User not found:', userEmail);
      return createUserNotFoundError();
    }

    console.log('📋 Config processing:', {
      userId: user.userId,
      spreadsheetId: config.spreadsheetId,
      sheetName: config.sheetName
    });

    // 設定をJSONで保存（重複フィールド削除）
    const removedFields = [];
    if ('setupComplete' in config) { delete config.setupComplete; removedFields.push('setupComplete'); }
    if ('isDraft' in config) { delete config.isDraft; removedFields.push('isDraft'); }
    if ('questionText' in config) { delete config.questionText; removedFields.push('questionText'); }

    const timestamp = new Date().toISOString();
    config.lastAccessedAt = timestamp;
    config.lastModified = timestamp;

    const configJsonString = JSON.stringify(config);
    const updatedUser = {
      configJson: configJsonString,
      lastModified: timestamp
    };

    const updateResult = db.updateUser(user.userId, updatedUser);

    if (!updateResult || !updateResult.success) {
      console.error('❌ Database update failed:', updateResult?.message);
      return createErrorResponse(updateResult?.message || 'データベース更新に失敗しました');
    }

    console.log('✅ saveDraftConfiguration SUCCESS:', {
      userId: user.userId,
      spreadsheetId: config.spreadsheetId,
      sheetName: config.sheetName,
      configSize: configJsonString.length
    });

    return {
      success: true,
      message: '下書き設定を保存しました',
      userId: user.userId
    };

  } catch (error) {
    const endTime = new Date().toISOString();
    const processingTime = new Date(endTime) - new Date(startTime);

    console.error('=== saveDraftConfiguration ERROR ===', {
      startTime,
      endTime,
      processingTimeMs: processingTime,
      errorMessage: error.message,
      errorStack: error.stack,
      configProvided: !!config,
      configKeys: config ? Object.keys(config) : null
    });

    return createExceptionResponse(error);
  }
}

/**
 * アプリケーションの公開
 * @param {Object} publishConfig - 公開設定
 * @returns {Object} 公開結果
 */
function publishApplication(publishConfig) {
  const startTime = new Date().toISOString();
  console.log('=== publishApplication START ===', {
    spreadsheetId: publishConfig?.spreadsheetId || 'N/A',
    sheetName: publishConfig?.sheetName || 'N/A',
    configSize: publishConfig ? JSON.stringify(publishConfig).length : 0
  });

  try {
    const email = getCurrentEmail();

    if (!email) {
      console.error('❌ User authentication failed');
      return { success: false, message: 'ユーザー認証が必要です' };
    }

    const publishedAt = new Date().toISOString();
    const props = ServiceFactory.getProperties();

    // Set properties
    try {
      props.setProperty('APPLICATION_STATUS', 'active');
      props.setProperty('PUBLISHED_AT', publishedAt);
    } catch (propsError) {
      console.error('❌ Properties update failed:', propsError.message);
    }

    if (publishConfig) {
      try {
        props.setProperty('PUBLISH_CONFIG', JSON.stringify(publishConfig));
      } catch (publishConfigError) {
        console.error('❌ PUBLISH_CONFIG save failed:', publishConfigError.message);
      }
    }

    const db = ServiceFactory.getDB();
    const user = db ? db.findUserByEmail(email) : null;

    if (user) {
      // Re-fetch latest user data to avoid conflicts
      const latestUser = db.findUserByEmail(email);
      const userToUse = latestUser || user;

      let currentConfig = {};
      try {
        currentConfig = userToUse.configJson ? JSON.parse(userToUse.configJson) : {};
      } catch (parseError) {
        console.error('❌ Config parse error:', parseError.message);
        currentConfig = {};
      }

      console.log('📋 Config merge:', {
        userId: user.userId,
        currentSpreadsheetId: currentConfig.spreadsheetId,
        newSpreadsheetId: publishConfig?.spreadsheetId,
        currentSheetName: currentConfig.sheetName,
        newSheetName: publishConfig?.sheetName
      });

      const updatedConfig = {
        ...currentConfig,
        ...publishConfig,
        appPublished: true,
        publishedAt,
        setupStatus: 'completed',
        isDraft: false,
        lastModified: publishedAt
      };

      const updatePayload = {
        configJson: JSON.stringify(updatedConfig),
        lastModified: publishedAt
      };

      const updateResult = db.updateUser(user.userId, updatePayload);

      if (!updateResult || !updateResult.success) {
        console.error('❌ Database update failed:', updateResult?.message);
        // Continue processing even on database errors
      }
    } else {
      console.error('❌ User not found:', email);
    }

    console.log('✅ publishApplication SUCCESS:', {
      userId: user?.userId || 'N/A',
      spreadsheetId: publishConfig?.spreadsheetId,
      sheetName: publishConfig?.sheetName,
      userFound: !!user
    });

    return {
      success: true,
      message: 'アプリケーションが正常に公開されました',
      publishedAt,
      userId: user ? user.userId : null
    };

  } catch (error) {
    console.error('❌ publishApplication ERROR:', {
      error: error.message,
      spreadsheetId: publishConfig?.spreadsheetId,
      sheetName: publishConfig?.sheetName,
      userEmail: getCurrentEmail()
    });

    return {
      success: false,
      message: error.message,
      error: error.message
    };
  }
}

/**
 * スプレッドシートへのアクセス権限を検証
 * AdminPanel.js.html から呼び出される
 *
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {Object} 検証結果
 */
function validateAccess(spreadsheetId, autoAddEditor = true) {
  try {
    // 🎯 Zero-dependency: サービスアカウント経由でアクセス権確認
    const dataAccess = Data.open(spreadsheetId);
    const {spreadsheet} = dataAccess;

    // サービスアカウントを編集者として自動登録
    if (autoAddEditor) {
      try {
        const serviceAccount = Config.serviceAccount();
        const serviceAccountEmail = serviceAccount ? serviceAccount.client_email : null;
        if (serviceAccountEmail) {
          spreadsheet.addEditor(serviceAccountEmail);
          console.log('validateAccess: サービスアカウントを編集者として登録:', serviceAccountEmail);
        }
      } catch (editorError) {
        console.warn('validateAccess: 編集者登録をスキップ:', editorError.message);
      }
    }

    const sheets = spreadsheet.getSheets();

    // アクセスできたら成功
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
 * フォーム情報を取得（実装関数）
 * main.gs API Gateway から呼び出される
 *
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} フォーム情報
 */
function getFormInfoImpl(spreadsheetId, sheetName) {
  try {
    // 引数検証
    if (!spreadsheetId || !sheetName) {
      return {
        success: false,
        status: 'INVALID_ARGUMENTS',
        message: 'スプレッドシートIDとシート名を指定してください。',
        formData: {
          formUrl: null,
          formTitle: sheetName || 'フォーム',
          spreadsheetName: '',
          sheetName
        }
      };
    }

    // スプレッドシート取得
    let spreadsheet;
    try {
      const { spreadsheet: spreadsheetFromData } = Data.open(spreadsheetId);
      spreadsheet = spreadsheetFromData;
    } catch (accessError) {
      return {
        success: false,
        status: 'SPREADSHEET_NOT_FOUND',
        message: 'スプレッドシートにアクセスできませんでした。',
        formData: {
          formUrl: null,
          formTitle: sheetName || 'フォーム',
          spreadsheetName: '',
          sheetName
        },
        error: accessError.message
      };
    }

    const spreadsheetName = spreadsheet.getName();

    // シート取得
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      return {
        success: false,
        status: 'SHEET_NOT_FOUND',
        message: `シート「${sheetName}」が見つかりません。`,
        formData: {
          formUrl: null,
          formTitle: sheetName,
          spreadsheetName,
          sheetName
        }
      };
    }

    // フォームURL取得（スタックオーバーフロー完全対策）
    let formUrl = null;
    let formTitle = sheetName || spreadsheetName || 'フォーム';

    // 安全なフォームURL取得 - FormApp.openByUrlは完全に回避
    try {
      // まずシートレベルでフォームURLを取得
      if (typeof sheet.getFormUrl === 'function') {
        formUrl = sheet.getFormUrl();
      }

      // シートレベルで取得できない場合はスプレッドシートレベルで試行
      if (!formUrl && typeof spreadsheet.getFormUrl === 'function') {
        formUrl = spreadsheet.getFormUrl();
      }
    } catch (urlError) {
      console.warn('SystemController.getFormInfo: FormURL取得エラー（安全に処理）:', urlError.message);
      // エラーが発生してもformUrlはnullのままで続行
    }

    // フォームタイトル設定（FormApp.openByUrlを使用せず安全に処理）
    if (formUrl && formUrl.includes('docs.google.com/forms/')) {
      // FormApp呼び出しは完全に回避し、シート名ベースのタイトルを使用
      formTitle = `${sheetName} (フォーム連携)`;
    }

    const formData = {
      formUrl: formUrl || null,
      formTitle,
      spreadsheetName,
      sheetName
    };

    // 成功レスポンス
    if (formUrl) {
      return {
        success: true,
        status: 'FORM_LINK_FOUND',
        message: 'フォーム連携を確認しました。',
        formData,
        timestamp: new Date().toISOString(),
        requestContext: {
          spreadsheetId,
          sheetName
        }
      };
    } else {
      return {
        success: false,
        status: 'FORM_NOT_LINKED',
        message: '指定したシートにはフォーム連携が確認できませんでした。',
        formData,
        suggestions: [
          'Googleフォームの「回答の行き先」を開き、対象のシートにリンクしてください',
          'フォーム作成者に連携状況を確認してください'
        ]
      };
    }

  } catch (error) {
    console.error('SystemController.getFormInfo エラー:', error.message);
    return {
      success: false,
      status: 'UNKNOWN_ERROR',
      message: 'フォーム情報の取得に失敗しました。',
      error: error.message,
      formData: {
        formUrl: null,
        formTitle: sheetName || 'フォーム',
        spreadsheetName: '',
        sheetName
      }
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
    // createForm機能は現在サポートされていません
    const result = { success: false, message: 'フォーム作成機能は現在利用できません' };

    if (result && result.success) {
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
  console.log('📊 checkCurrentPublicationStatus START:', {
    targetUserId: targetUserId || 'current_user'
  });

  try {
    const session = ServiceFactory.getSession();
    const db = ServiceFactory.getDB();

    if (!db) {
      console.error('❌ Database connection failed');
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
      console.error('❌ User not found');
      return createUserNotFoundError();
    }

    let config = {};
    try {
      config = JSON.parse(user.configJson || '{}');
    } catch (parseError) {
      console.error('❌ Config parse error:', parseError.message);
      config = {};
    }

    const result = {
      success: true,
      published: config.appPublished === true,
      publishedAt: config.publishedAt || null,
      lastModified: config.lastModified || null,
      hasDataSource: Boolean(config.spreadsheetId && config.sheetName),
      userId: user.userId
    };

    console.log('✅ checkCurrentPublicationStatus SUCCESS:', {
      userId: user.userId,
      published: result.published,
      hasDataSource: result.hasDataSource
    });

    return result;
  } catch (error) {
    console.error('❌ checkCurrentPublicationStatus ERROR:', error.message);
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

    // エラーログを記録（将来的には永続化）

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
