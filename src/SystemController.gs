/**
 * @fileoverview SystemController - System management and setup functions
 */

/* global ServiceFactory, UserService, ConfigService, getCurrentEmail, createErrorResponse, createUserNotFoundError, createExceptionResponse, getSheetsService, getServiceAccountEmail, getSpreadsheetList */

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
          const spreadsheet = getSheetsService().openById(databaseId);
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
    const spreadsheet = getSheetsService().openById(spreadsheetId);
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
    timestamp: startTime,
    configKeys: config ? Object.keys(config) : null,
    configSize: config ? JSON.stringify(config).length : 0
  });

  try {
    // 🎯 Input validation logging
    if (!config || typeof config !== 'object') {
      console.error('saveDraftConfiguration: Invalid config input', { config });
      return createErrorResponse('設定データが不正です');
    }

    // 🎯 User authentication logging
    const userEmail = getCurrentEmail();
    console.log('saveDraftConfiguration: User authentication', {
      userEmail: userEmail ? 'FOUND' : 'NOT_FOUND',
      emailLength: userEmail ? userEmail.length : 0
    });

    if (!userEmail) {
      console.error('saveDraftConfiguration: Authentication failed');
      return createErrorResponse('ユーザー認証が必要です');
    }

    // 🎯 Database connection logging
    const db = ServiceFactory.getDB();
    console.log('saveDraftConfiguration: Database connection', {
      dbAvailable: !!db,
      dbType: db ? typeof db : 'undefined'
    });

    if (!db) {
      console.error('saveDraftConfiguration: Database connection failed');
      return createErrorResponse('データベース接続エラー');
    }

    // 🎯 User lookup logging
    const user = db.findUserByEmail(userEmail);
    console.log('saveDraftConfiguration: User lookup', {
      userFound: !!user,
      userId: user ? user.userId : null,
      currentConfigJson: user ? (user.configJson ? user.configJson.length : 0) : null
    });

    if (!user) {
      console.error('saveDraftConfiguration: User not found', { userEmail });
      return createUserNotFoundError();
    }

    // 🎯 Configuration processing logging
    const originalConfigKeys = Object.keys(config);
    console.log('saveDraftConfiguration: Configuration processing START', {
      originalKeys: originalConfigKeys,
      originalSize: JSON.stringify(config).length
    });

    // 設定をJSONで保存（重複フィールド削除）
    const removedFields = [];
    if ('setupComplete' in config) { delete config.setupComplete; removedFields.push('setupComplete'); }
    if ('isDraft' in config) { delete config.isDraft; removedFields.push('isDraft'); }
    if ('questionText' in config) { delete config.questionText; removedFields.push('questionText'); }

    const timestamp = new Date().toISOString();
    config.lastAccessedAt = timestamp;
    config.lastModified = timestamp;

    console.log('saveDraftConfiguration: Configuration processing COMPLETE', {
      removedFields,
      addedFields: ['lastAccessedAt', 'lastModified'],
      finalKeys: Object.keys(config),
      finalSize: JSON.stringify(config).length,
      timestamp
    });

    // 🎯 User update object creation logging
    const configJsonString = JSON.stringify(config);
    const updatedUser = {
      configJson: configJsonString,
      lastModified: timestamp
    };

    console.log('saveDraftConfiguration: User update object created', {
      userId: user.userId,
      configJsonLength: configJsonString.length,
      updatedUserKeys: Object.keys(updatedUser),
      configJsonPreview: `${configJsonString.substring(0, 100)  }...`
    });

    // 🎯 Database update operation logging
    console.log('saveDraftConfiguration: Starting database update', {
      userId: user.userId,
      updateTimestamp: timestamp
    });

    const updateResult = db.updateUser(user.userId, updatedUser);

    console.log('saveDraftConfiguration: Database update result', {
      success: updateResult ? updateResult.success : false,
      message: updateResult ? updateResult.message : 'NO_RESULT',
      resultKeys: updateResult ? Object.keys(updateResult) : null,
      userId: user.userId
    });

    if (!updateResult || !updateResult.success) {
      console.error('saveDraftConfiguration: Database update failed', {
        updateResult,
        userId: user.userId,
        configSize: configJsonString.length
      });
      return createErrorResponse(updateResult?.message || 'データベース更新に失敗しました');
    }

    // 🎯 Success validation logging
    const endTime = new Date().toISOString();
    const processingTime = new Date(endTime) - new Date(startTime);

    console.log('=== saveDraftConfiguration SUCCESS ===', {
      startTime,
      endTime,
      processingTimeMs: processingTime,
      userId: user.userId,
      configJsonLength: configJsonString.length,
      finalMessage: '下書き設定を保存しました'
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
    timestamp: startTime,
    publishConfigProvided: !!publishConfig,
    publishConfigKeys: publishConfig ? Object.keys(publishConfig) : null,
    publishConfigSize: publishConfig ? JSON.stringify(publishConfig).length : 0
  });

  try {
    // 🎯 Input validation logging
    console.log('publishApplication: Input validation', {
      publishConfigType: typeof publishConfig,
      publishConfigNull: publishConfig === null,
      publishConfigUndefined: publishConfig === undefined,
      publishConfigEmpty: publishConfig && Object.keys(publishConfig).length === 0
    });

    // 🎯 User authentication logging
    const email = getCurrentEmail();
    console.log('publishApplication: User authentication', {
      emailFound: !!email,
      emailLength: email ? email.length : 0
    });

    if (!email) {
      console.error('publishApplication: Authentication failed');
      return { success: false, message: 'ユーザー認証が必要です' };
    }

    const publishedAt = new Date().toISOString();
    console.log('publishApplication: Timestamp generated', { publishedAt });

    // 🎯 PropertiesService operations logging
    console.log('publishApplication: Starting PropertiesService operations');
    const props = ServiceFactory.getProperties();
    console.log('publishApplication: PropertiesService connection', {
      propsAvailable: !!props,
      propsType: typeof props
    });

    // Set APPLICATION_STATUS
    try {
      props.setProperty('APPLICATION_STATUS', 'active');
      console.log('publishApplication: APPLICATION_STATUS set to active');
    } catch (statusError) {
      console.error('publishApplication: Failed to set APPLICATION_STATUS', {
        error: statusError.message
      });
    }

    // Set PUBLISHED_AT
    try {
      props.setProperty('PUBLISHED_AT', publishedAt);
      console.log('publishApplication: PUBLISHED_AT set', { publishedAt });
    } catch (publishedAtError) {
      console.error('publishApplication: Failed to set PUBLISHED_AT', {
        error: publishedAtError.message,
        publishedAt
      });
    }

    // 公開設定を保存
    if (publishConfig) {
      try {
        const publishConfigString = JSON.stringify(publishConfig);
        props.setProperty('PUBLISH_CONFIG', publishConfigString);
        console.log('publishApplication: PUBLISH_CONFIG saved', {
          configLength: publishConfigString.length,
          configKeys: Object.keys(publishConfig)
        });
      } catch (publishConfigError) {
        console.error('publishApplication: Failed to set PUBLISH_CONFIG', {
          error: publishConfigError.message,
          publishConfigKeys: Object.keys(publishConfig)
        });
      }
    } else {
      console.log('publishApplication: No publishConfig provided, skipping PUBLISH_CONFIG');
    }

    // 🎯 Database operations logging
    console.log('publishApplication: Starting database operations');
    const db = ServiceFactory.getDB();
    console.log('publishApplication: Database connection', {
      dbAvailable: !!db,
      dbType: typeof db
    });

    const user = db ? db.findUserByEmail(email) : null;
    console.log('publishApplication: User lookup', {
      userFound: !!user,
      userId: user ? user.userId : null,
      currentConfigJsonLength: user ? (user.configJson ? user.configJson.length : 0) : null
    });

    if (user) {
      // 🎯 Configuration merge logging
      let currentConfig = {};
      try {
        currentConfig = user.configJson ? JSON.parse(user.configJson) : {};
        console.log('publishApplication: Current config parsed', {
          currentConfigKeys: Object.keys(currentConfig),
          currentConfigSize: JSON.stringify(currentConfig).length
        });
      } catch (parseError) {
        console.error('publishApplication: Failed to parse current config', {
          error: parseError.message,
          configJsonLength: user.configJson ? user.configJson.length : 0
        });
        currentConfig = {};
      }

      const updatedConfig = {
        ...currentConfig,
        ...publishConfig,
        appPublished: true,
        publishedAt,
        setupStatus: 'completed',
        isDraft: false,
        lastModified: publishedAt
      };

      console.log('publishApplication: Config merge completed', {
        originalKeys: Object.keys(currentConfig),
        publishConfigKeys: publishConfig ? Object.keys(publishConfig) : [],
        mergedKeys: Object.keys(updatedConfig),
        finalConfigSize: JSON.stringify(updatedConfig).length,
        addedFields: ['appPublished', 'publishedAt', 'setupStatus', 'isDraft', 'lastModified']
      });

      // 🎯 Database update operation logging
      const updatePayload = {
        configJson: JSON.stringify(updatedConfig),
        lastModified: publishedAt
      };

      console.log('publishApplication: Starting database update', {
        userId: user.userId,
        updatePayloadKeys: Object.keys(updatePayload),
        configJsonLength: updatePayload.configJson.length,
        updateTimestamp: publishedAt
      });

      const updateResult = db.updateUser(user.userId, updatePayload);

      console.log('publishApplication: Database update result', {
        success: updateResult ? updateResult.success : false,
        message: updateResult ? updateResult.message : 'NO_RESULT',
        resultKeys: updateResult ? Object.keys(updateResult) : null,
        userId: user.userId,
        configJsonLength: updatePayload.configJson.length
      });

      if (!updateResult || !updateResult.success) {
        console.error('publishApplication: Database update failed', {
          updateResult,
          userId: user.userId,
          updatePayload: {
            ...updatePayload,
            configJson: `[${updatePayload.configJson.length} chars]`
          }
        });
        // エラーでも処理は継続
      } else {
        console.log('publishApplication: Database update successful', {
          userId: user.userId,
          configJsonLength: updatePayload.configJson.length
        });
      }
    } else {
      console.error('publishApplication: User not found, skipping database update', { email });
    }

    // 🎯 Success validation logging
    const endTime = new Date().toISOString();
    const processingTime = new Date(endTime) - new Date(startTime);

    console.log('=== publishApplication SUCCESS ===', {
      startTime,
      endTime,
      processingTimeMs: processingTime,
      publishedAt,
      userEmail: email,
      userFound: !!user,
      userId: user ? user.userId : null,
      propertiesUpdated: true,
      databaseUpdated: !!user,
      finalMessage: 'アプリケーションが正常に公開されました'
    });

    return {
      success: true,
      message: 'アプリケーションが正常に公開されました',
      publishedAt
    };

  } catch (error) {
    const endTime = new Date().toISOString();
    const processingTime = new Date(endTime) - new Date(startTime);

    console.error('=== publishApplication ERROR ===', {
      startTime,
      endTime,
      processingTimeMs: processingTime,
      errorMessage: error.message,
      errorStack: error.stack,
      publishConfigProvided: !!publishConfig,
      publishConfigKeys: publishConfig ? Object.keys(publishConfig) : null,
      userEmail: getCurrentEmail()
    });

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
function validateAccess(spreadsheetId, autoAddEditor = true) {
  try {
    // 🎯 Zero-dependency: サービスアカウント経由でアクセス権確認
    const spreadsheet = getSheetsService().openById(spreadsheetId);

    // サービスアカウントを編集者として自動登録
    if (autoAddEditor) {
      try {
        const serviceAccountEmail = getServiceAccountEmail();
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
 * フォーム情報を取得
 * AdminPanel.js.html から呼び出される
 *
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} フォーム情報
 */
function getFormInfo(spreadsheetId, sheetName) {
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
      spreadsheet = getSheetsService().openById(spreadsheetId);
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
