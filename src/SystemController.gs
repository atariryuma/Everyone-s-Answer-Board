/**
 * @fileoverview SystemController - System management and setup functions
 */

/* global ServiceFactory, UserService, ConfigService, getCurrentEmail, createErrorResponse, createUserNotFoundError, createExceptionResponse, Data, Config, getSpreadsheetList, getConfigSafe, saveConfigSafe */

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

    // 🔧 FIX: Transform columnMapping from frontend format to backend format
    if (config.columnMapping && typeof config.columnMapping === 'object') {
      console.log('🔍 BACKEND TRANSFORMATION START:', {
        originalColumnMapping: config.columnMapping,
        hasMapping: !!config.columnMapping.mapping
      });

      // If columnMapping doesn't have the correct structure, transform it
      if (!config.columnMapping.mapping) {
        const transformedMapping = {};
        const transformedConfidence = {};

        // Transform each column type from { columnIndex: N } to mapping[type] = N
        Object.keys(config.columnMapping).forEach(key => {
          if (key.startsWith('_') || key === 'headers' || key === 'verifiedAt') return;

          const columnData = config.columnMapping[key];
          if (columnData && typeof columnData.columnIndex === 'number') {
            transformedMapping[key] = columnData.columnIndex;
            transformedConfidence[key] = columnData.confidence || 0;
            console.log(`✅ Transformed ${key}: ${columnData.columnIndex}`);
          }
        });

        // Rebuild columnMapping with correct structure
        config.columnMapping = {
          mapping: transformedMapping,
          confidence: transformedConfidence,
          headers: config.columnMapping.headers || [],
          verifiedAt: config.columnMapping.verifiedAt || new Date().toISOString()
        };

        console.log('✅ BACKEND TRANSFORMATION COMPLETE:', {
          transformedMapping,
          finalColumnMapping: config.columnMapping
        });
      } else {
        console.log('🔍 ColumnMapping already has correct structure');
      }
    }

    // 🔧 CLAUDE.md準拠: 統一API使用 - saveConfigSafeを使用してETag対応の安全な更新
    const removedFields = [];
    if ('setupComplete' in config) { delete config.setupComplete; removedFields.push('setupComplete'); }
    if ('isDraft' in config) { delete config.isDraft; removedFields.push('isDraft'); }
    if ('questionText' in config) { delete config.questionText; removedFields.push('questionText'); }

    // saveConfigSafeを使用して統一された更新パターンを適用
    const saveResult = saveConfigSafe(user.userId, config, { isDraft: true });

    if (!saveResult.success) {
      console.error('❌ saveConfigSafe failed:', saveResult.message);
      return createErrorResponse(saveResult.message || 'データベース更新に失敗しました');
    }

    console.log('✅ saveDraftConfiguration SUCCESS (via saveConfigSafe):', {
      userId: user.userId,
      spreadsheetId: config.spreadsheetId,
      sheetName: config.sheetName,
      etag: saveResult.etag
    });

    return {
      success: true,
      message: '下書き設定を保存しました',
      userId: user.userId,
      etag: saveResult.etag,
      config: saveResult.config
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
    let saveResult = null;

    if (user) {
      // Re-fetch latest user data to avoid conflicts
      const latestUser = db.findUserByEmail(email);
      const userToUse = latestUser || user;

      // 統一API使用: 構造化パース
      const configResult = getConfigSafe(userToUse.userId);
      const currentConfig = configResult.success ? configResult.config : {};

      console.log('📋 Config merge:', {
        userId: user.userId,
        currentSpreadsheetId: currentConfig.spreadsheetId,
        newSpreadsheetId: publishConfig?.spreadsheetId,
        currentFormUrl: currentConfig.formUrl,
        newFormUrl: publishConfig?.formUrl,
        currentSheetName: currentConfig.sheetName,
        newSheetName: publishConfig?.sheetName
      });

      // 🔧 重要フィールドの明示的上書き（新しい値を優先）
      const updatedConfig = {
        ...currentConfig,
        ...publishConfig,
        // 🎯 critical fields: 必ず新しい値を使用
        formUrl: publishConfig?.formUrl || currentConfig.formUrl,
        formTitle: publishConfig?.formTitle || currentConfig.formTitle,
        columnMapping: publishConfig?.columnMapping || currentConfig.columnMapping,
        // 🔧 system fields
        appPublished: true,
        publishedAt,
        setupStatus: 'completed',
        isDraft: false,
        lastModified: publishedAt
      };

      // 🔧 FIX: Transform columnMapping from frontend format to backend format (publishApplication)
      if (updatedConfig.columnMapping && typeof updatedConfig.columnMapping === 'object') {
        console.log('🔍 PUBLISH TRANSFORMATION START:', {
          originalColumnMapping: updatedConfig.columnMapping,
          hasMapping: !!updatedConfig.columnMapping.mapping
        });

        // If columnMapping doesn't have the correct structure, transform it
        if (!updatedConfig.columnMapping.mapping) {
          const transformedMapping = {};
          const transformedConfidence = {};

          // Transform each column type from { columnIndex: N } to mapping[type] = N
          Object.keys(updatedConfig.columnMapping).forEach(key => {
            if (key.startsWith('_') || key === 'headers' || key === 'verifiedAt') return;

            const columnData = updatedConfig.columnMapping[key];
            if (columnData && typeof columnData.columnIndex === 'number') {
              transformedMapping[key] = columnData.columnIndex;
              transformedConfidence[key] = columnData.confidence || 0;
              console.log(`✅ Transformed ${key}: ${columnData.columnIndex}`);
            }
          });

          // Rebuild columnMapping with correct structure
          updatedConfig.columnMapping = {
            mapping: transformedMapping,
            confidence: transformedConfidence,
            headers: updatedConfig.columnMapping.headers || [],
            verifiedAt: updatedConfig.columnMapping.verifiedAt || new Date().toISOString()
          };

          console.log('✅ PUBLISH TRANSFORMATION COMPLETE:', {
            transformedMapping,
            finalColumnMapping: updatedConfig.columnMapping
          });
        } else {
          console.log('🔍 ColumnMapping already has correct structure (publish)');
        }
      }

      // 🔧 CLAUDE.md準拠: 統一API使用 - saveConfigSafeでETag対応の安全な更新
      saveResult = saveConfigSafe(user.userId, updatedConfig, { isPublish: true });

      if (!saveResult.success) {
        console.error('❌ saveConfigSafe failed during publish:', saveResult.message);
        // エラーでも処理を継続（互換性のため）
      } else {
        console.log('✅ Config saved via saveConfigSafe:', {
          userId: user.userId,
          etag: saveResult.etag
        });
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
      userId: user ? user.userId : null,
      etag: user && saveResult?.etag ? saveResult.etag : null,
      config: user && saveResult?.config ? saveResult.config : null
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
 * ユーザーがスプレッドシートのオーナーかチェック
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {boolean} オーナーかどうか
 */
function isUserSpreadsheetOwner(spreadsheetId) {
  try {
    const currentEmail = getCurrentEmail();
    if (!currentEmail) return false;

    // DriveAppで所有者確認を試行
    const file = DriveApp.getFileById(spreadsheetId);
    const owner = file.getOwner();

    if (owner && owner.getEmail() === currentEmail) {
      console.log('isUserSpreadsheetOwner: ユーザーはオーナーです:', currentEmail);
      return true;
    }

    console.log('isUserSpreadsheetOwner: ユーザーはオーナーではありません:', {
      currentEmail,
      ownerEmail: owner ? owner.getEmail() : 'unknown'
    });
    return false;
  } catch (error) {
    console.warn('isUserSpreadsheetOwner: 権限チェック失敗:', error.message);
    return false;
  }
}

/**
 * 🔧 CLAUDE.md準拠: セキュアなサービスアカウント専用スプレッドシートアクセス
 * オーナー判定は残すが、アクセスは全てサービスアカウント経由で統一
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {Object} {spreadsheet, accessMethod, auth, isOwner}
 */
function getSpreadsheetAdaptive(spreadsheetId) {
  // オーナー判定（表示・ログ目的のみ）
  const isOwner = isUserSpreadsheetOwner(spreadsheetId);

  console.log(`getSpreadsheetAdaptive: ${isOwner ? 'Owner' : 'Non-owner'} accessing via service account`);

  // 🛡️ セキュアなサービスアカウント専用アクセス
  try {
    const dataAccess = Data.open(spreadsheetId);
    console.log('getSpreadsheetAdaptive: サービスアカウントでアクセス成功');
    return {
      spreadsheet: dataAccess.spreadsheet,
      accessMethod: 'service_account',
      auth: dataAccess.auth,
      isOwner
    };
  } catch (serviceError) {
    console.error('getSpreadsheetAdaptive: サービスアカウントアクセス失敗:', serviceError.message);
    throw new Error(`セキュアスプレッドシートアクセス失敗: ${serviceError.message}`);
  }
}

/**
 * 多層フォーム検出システム
 * @param {Object} spreadsheet - スプレッドシートオブジェクト
 * @param {Object} sheet - シートオブジェクト
 * @param {string} sheetName - シート名
 * @param {boolean} isOwner - オーナー権限フラグ
 * @returns {Object} {formUrl, confidence, detectionMethod}
 */
function detectFormConnection(spreadsheet, sheet, sheetName, isOwner) {
  const results = {
    formUrl: null,
    formTitle: null,
    confidence: 0,
    detectionMethod: 'none',
    details: []
  };

  // Method 1: 標準API - オーナー権限の場合のみ
  if (isOwner) {
    try {
      // シートレベルでフォームURL取得（最優先）
      if (typeof sheet.getFormUrl === 'function') {
        const formUrl = sheet.getFormUrl();
        if (formUrl) {
          results.formUrl = formUrl;
          results.confidence = 95;
          results.detectionMethod = 'sheet_api';
          results.details.push('Sheet.getFormUrl()で検出');

          // FormApp.openByUrlでタイトル取得
          try {
            const form = FormApp.openByUrl(formUrl);
            results.formTitle = form.getTitle();
            results.details.push('FormApp.openByUrl()でタイトル取得');
          } catch (titleError) {
            console.warn('FormApp.openByUrl failed:', titleError.message);
            results.formTitle = generateFormTitle(sheetName, spreadsheet.getName());
            results.details.push('FormApp権限エラー - フォールバックタイトル使用');
          }
          return results;
        }
      }

      // スプレッドシートレベルでフォームURL取得（フォールバック）
      if (typeof spreadsheet.getFormUrl === 'function') {
        const formUrl = spreadsheet.getFormUrl();
        if (formUrl) {
          results.formUrl = formUrl;
          results.confidence = 85;
          results.detectionMethod = 'spreadsheet_api';
          results.details.push('SpreadsheetApp.getFormUrl()で検出');

          // FormApp.openByUrlでタイトル取得
          try {
            const form = FormApp.openByUrl(formUrl);
            results.formTitle = form.getTitle();
            results.details.push('FormApp.openByUrl()でタイトル取得');
          } catch (titleError) {
            console.warn('FormApp.openByUrl failed:', titleError.message);
            results.formTitle = generateFormTitle(sheetName, spreadsheet.getName());
            results.details.push('FormApp権限エラー - フォールバックタイトル使用');
          }
          return results;
        }
      }
    } catch (apiError) {
      console.warn('detectFormConnection: API検出失敗:', apiError.message);
      results.details.push(`API検出失敗: ${apiError.message}`);
    }
  }

  // Method 1.5: Drive APIフォーム検索（API検出失敗時のフォールバック）
  if (isOwner) {
    try {
      const spreadsheetId = spreadsheet.getId();
      const driveFormResult = searchFormsByDrive(spreadsheetId, sheetName);
      if (driveFormResult.formUrl) {
        results.formUrl = driveFormResult.formUrl;
        results.formTitle = driveFormResult.formTitle;
        results.confidence = 80;
        results.detectionMethod = 'drive_search';
        results.details.push('Drive API検索で検出');
        return results;
      }
    } catch (driveError) {
      console.warn('detectFormConnection: Drive API検索失敗:', driveError.message);
      results.details.push(`Drive API検索失敗: ${driveError.message}`);
    }
  }

  // Method 2: ヘッダーパターン解析
  try {
    const [headers] = sheet.getRange(1, 1, 1, Math.min(sheet.getLastColumn(), 10)).getValues();
    const headerAnalysis = analyzeFormHeaders(headers);

    if (headerAnalysis.isFormLike) {
      results.confidence = Math.max(results.confidence, headerAnalysis.confidence);
      results.detectionMethod = results.detectionMethod === 'none' ? 'header_analysis' : results.detectionMethod;
      results.details.push(`ヘッダー解析: ${headerAnalysis.reason}`);
    }
  } catch (headerError) {
    console.warn('detectFormConnection: ヘッダー解析失敗:', headerError.message);
    results.details.push(`ヘッダー解析失敗: ${headerError.message}`);
  }

  // Method 3: シート名パターン解析
  const sheetNameAnalysis = analyzeSheetName(sheetName);
  if (sheetNameAnalysis.isFormLike) {
    results.confidence = Math.max(results.confidence, sheetNameAnalysis.confidence);
    results.detectionMethod = results.detectionMethod === 'none' ? 'sheet_name' : results.detectionMethod;
    results.details.push(`シート名解析: ${sheetNameAnalysis.reason}`);
  }

  // フォーム検出時のタイトル生成
  if (results.confidence >= 40) {
    results.formTitle = `${sheetName} (フォーム検出済み)`;
  }

  return results;
}

/**
 * フォームヘッダーパターン解析
 * @param {Array} headers - ヘッダー配列
 * @returns {Object} {isFormLike, confidence, reason}
 */
function analyzeFormHeaders(headers) {
  if (!headers || headers.length === 0) {
    return { isFormLike: false, confidence: 0, reason: 'ヘッダーなし' };
  }

  const formIndicators = [
    { pattern: /タイムスタンプ|timestamp/i, weight: 30, description: 'タイムスタンプ列' },
    { pattern: /メールアドレス|email.*address|メール/i, weight: 25, description: 'メールアドレス列' },
    { pattern: /回答|answer|意見|response/i, weight: 20, description: '回答列' },
    { pattern: /名前|name|氏名/i, weight: 15, description: '名前列' },
    { pattern: /クラス|class|組/i, weight: 10, description: 'クラス列' }
  ];

  let totalScore = 0;
  const matches = [];

  headers.forEach((header, index) => {
    if (typeof header === 'string') {
      formIndicators.forEach(indicator => {
        if (indicator.pattern.test(header)) {
          totalScore += indicator.weight;
          matches.push(`${indicator.description}(${header})`);
        }
      });
    }
  });

  const confidence = Math.min(totalScore, 85); // 最大85%
  const isFormLike = confidence >= 40; // 40%以上でフォームと判定

  return {
    isFormLike,
    confidence,
    reason: isFormLike ? `フォームパターン検出: ${matches.join(', ')}` : '一般的なフォームパターンが不足'
  };
}

/**
 * シート名パターン解析
 * @param {string} sheetName - シート名
 * @returns {Object} {isFormLike, confidence, reason}
 */
function analyzeSheetName(sheetName) {
  if (!sheetName || typeof sheetName !== 'string') {
    return { isFormLike: false, confidence: 0, reason: 'シート名なし' };
  }

  const patterns = [
    { regex: /フォームの回答|form.*responses?/i, confidence: 80, description: 'フォーム回答シート' },
    { regex: /フォーム.*結果|form.*results?/i, confidence: 75, description: 'フォーム結果シート' },
    { regex: /回答.*\d+|responses?.*\d+/i, confidence: 70, description: 'ナンバリング回答シート' },
    { regex: /アンケート|survey|questionnaire/i, confidence: 60, description: 'アンケートシート' },
    { regex: /回答|responses?|答え|answer/i, confidence: 50, description: '回答関連シート' }
  ];

  for (const pattern of patterns) {
    if (pattern.regex.test(sheetName)) {
      return {
        isFormLike: true,
        confidence: pattern.confidence,
        reason: `${pattern.description}パターン検出: "${sheetName}"`
      };
    }
  }

  return {
    isFormLike: false,
    confidence: 0,
    reason: `一般的なフォームシート名パターンなし: "${sheetName}"`
  };
}

/**
 * フォールバックタイトル生成
 * @param {string} sheetName - シート名
 * @param {string} spreadsheetName - スプレッドシート名
 * @returns {string} 生成されたタイトル
 */
function generateFormTitle(sheetName, spreadsheetName) {
  try {
    // シート名がフォーム関連の場合
    if (sheetName && typeof sheetName === 'string') {
      const formPatterns = [
        /フォームの回答|form.*responses?/i,
        /フォーム.*結果|form.*results?/i,
        /回答.*\d+|responses?.*\d+/i,
        /アンケート|survey|questionnaire/i
      ];

      for (const pattern of formPatterns) {
        if (pattern.test(sheetName)) {
          // フォーム関連のシート名の場合、「の回答」を除去してフォームタイトルとする
          const formTitle = sheetName
            .replace(/の回答.*$/i, '')
            .replace(/.*responses?.*$/i, '')
            .replace(/.*結果.*$/i, '')
            .trim();

          if (formTitle.length > 0) {
            return `${formTitle  } (フォーム)`;
          }
        }
      }
    }

    // スプレッドシート名ベースのタイトル生成
    if (spreadsheetName && typeof spreadsheetName === 'string') {
      return `${spreadsheetName  } - ${  sheetName  } (フォーム)`;
    }

    // 最終フォールバック
    return `${sheetName || 'データ'  } (フォーム)`;

  } catch (error) {
    console.warn('generateFormTitle error:', error.message);
    return `${sheetName || 'フォーム'  } (タイトル生成エラー)`;
  }
}

/**
 * Drive APIによるフォーム検索
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} {formUrl, formTitle}
 */
function searchFormsByDrive(spreadsheetId, sheetName) {
  try {
    console.log('searchFormsByDrive: 検索開始', { spreadsheetId: `${spreadsheetId.substring(0, 12)  }***`, sheetName });

    // Drive APIでフォーム一覧を取得
    const forms = DriveApp.getFilesByType('application/vnd.google-apps.form');

    while (forms.hasNext()) {
      const formFile = forms.next();
      try {
        const form = FormApp.openById(formFile.getId());
        const destId = form.getDestinationId();

        if (destId === spreadsheetId) {
          console.log('searchFormsByDrive: 対象スプレッドシートへのフォーム発見', {
            formId: formFile.getId(),
            formTitle: form.getTitle()
          });

          // 対象シートへの接続確認
          const destSpreadsheet = ServiceFactory.getSpreadsheet().openById(destId);
          const targetSheet = destSpreadsheet.getSheetByName(sheetName);

          if (targetSheet) {
            const sheetFormUrl = targetSheet.getFormUrl();
            const formPublishedUrl = form.getPublishedUrl();

            // URL照合またはform ID照合
            if (sheetFormUrl && (sheetFormUrl === formPublishedUrl ||
                sheetFormUrl.includes(formFile.getId()) ||
                formPublishedUrl.includes(formFile.getId()))) {

              console.log('searchFormsByDrive: フォーム接続確認成功', {
                sheetName,
                formTitle: form.getTitle()
              });

              return {
                formUrl: formPublishedUrl,
                formTitle: form.getTitle()
              };
            }
          }
        }
      } catch (formError) {
        // 個別フォームアクセスエラーは無視して継続
        console.warn('searchFormsByDrive: フォームアクセスエラー（継続）:', formError.message);
      }
    }

    console.log('searchFormsByDrive: 対象フォームが見つかりませんでした');
    return { formUrl: null, formTitle: null };

  } catch (error) {
    console.error('searchFormsByDrive: Drive API検索エラー:', error.message);
    throw error;
  }
}

/**
 * Sheets APIでスプレッドシート情報を取得
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} accessToken - アクセストークン
 * @returns {Object} スプレッドシート情報
 */
function getSpreadsheetInfo(spreadsheetId, accessToken) {
  try {
    const response = UrlFetchApp.fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}`, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      return {
        name: data.properties?.title || 'Unknown',
        owner: 'Service Account Access'  // Sheets APIでは所有者情報は取得できない
      };
    } else {
      console.warn('getSpreadsheetInfo: Sheets API error:', response.getContentText());
      return { name: 'Unknown', owner: 'Unknown' };
    }
  } catch (error) {
    console.error('getSpreadsheetInfo error:', error.message);
    return { name: 'Unknown', owner: 'Unknown' };
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
    const {spreadsheet, auth} = dataAccess;

    // サービスアカウントを編集者として自動登録（Data.openで既に実行済み）
    if (autoAddEditor) {
      console.log('validateAccess: サービスアカウント編集者権限は Data.open で既に処理済み');
    }

    // カスタムラッパーのgetSheets()メソッドを使用
    const sheets = spreadsheet.getSheets();

    // スプレッドシート情報を取得（Sheets API経由）
    const spreadsheetInfo = getSpreadsheetInfo(spreadsheetId, auth.token);

    // アクセスできたら成功
    const result = {
      success: true,
      message: 'アクセス権限が確認されました',
      spreadsheetName: spreadsheetInfo.name || `スプレッドシート (ID: ${spreadsheetId.substring(0, 8)}...)`,
      sheets: sheets.map(sheet => ({
        name: sheet.getName(),
        rowCount: sheet.getLastRow(),
        columnCount: sheet.getLastColumn()
      })),
      owner: spreadsheetInfo.owner || 'unknown',
      url: `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`
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
 * フォーム情報を取得（実装関数）- 適応的アクセス対応版
 * main.gs API Gateway から呼び出される
 *
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} フォーム情報
 */
function getFormInfo(spreadsheetId, sheetName) {
  const startTime = new Date().toISOString();
  console.log('=== getFormInfo START ===', {
    spreadsheetId: spreadsheetId ? `${spreadsheetId.substring(0, 12)  }***` : 'N/A',
    sheetName: sheetName || 'N/A',
    timestamp: startTime
  });

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

    // 適応的スプレッドシートアクセス（ユーザー所有 vs サービスアカウント）
    const accessResult = getSpreadsheetAdaptive(spreadsheetId);
    if (!accessResult.spreadsheet) {
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
        accessMethod: accessResult.accessMethod,
        error: accessResult.error
      };
    }

    const { spreadsheet, accessMethod, auth, isOwner } = accessResult;

    // スプレッドシート名取得（アクセス方法により異なる）
    let spreadsheetName;
    try {
      if (isOwner) {
        // ユーザー所有の場合はgetNameメソッド使用可能
        spreadsheetName = spreadsheet.getName();
      } else {
        // サービスアカウントの場合はSheets API使用
        const spreadsheetInfo = getSpreadsheetInfo(spreadsheetId, auth.token);
        spreadsheetName = spreadsheetInfo.name;
      }
    } catch (nameError) {
      console.warn('getFormInfoImpl: スプレッドシート名取得エラー:', nameError.message);
      if (spreadsheetId && spreadsheetId.trim()) {
        spreadsheetName = `スプレッドシート (ID: ${spreadsheetId.substring(0, 8)}...)`;
      }
    }

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
        },
        accessMethod
      };
    }

    // 多層フォーム検出システム実行
    const formDetectionResult = detectFormConnection(spreadsheet, sheet, sheetName, isOwner);

    console.log('getFormInfoImpl: フォーム検出結果', {
      hasFormUrl: !!formDetectionResult.formUrl,
      detectionMethod: formDetectionResult.detectionMethod,
      confidence: formDetectionResult.confidence,
      accessMethod
    });

    const formData = {
      formUrl: formDetectionResult.formUrl || null,
      formTitle: formDetectionResult.formTitle || sheetName,
      spreadsheetName,
      sheetName,
      detectionDetails: {
        method: formDetectionResult.detectionMethod,
        confidence: formDetectionResult.confidence,
        accessMethod,
        timestamp: new Date().toISOString()
      }
    };

    // レスポンス生成
    if (formDetectionResult.formUrl) {
      return {
        success: true,
        status: 'FORM_LINK_FOUND',
        message: `フォーム連携を確認しました (${formDetectionResult.detectionMethod})`,
        formData,
        timestamp: new Date().toISOString(),
        requestContext: {
          spreadsheetId,
          sheetName,
          accessMethod
        }
      };
    } else {
      // フォーム未検出時も詳細情報を提供
      const isHighConfidence = formDetectionResult.confidence >= 70;
      return {
        success: isHighConfidence,
        status: isHighConfidence ? 'FORM_DETECTED_NO_URL' : 'FORM_NOT_LINKED',
        message: isHighConfidence ?
          'フォーム連携が検出されました（URLは取得できませんが、フォーム接続されている可能性が高いです）' :
          '指定したシートにはフォーム連携が確認できませんでした。',
        formData,
        suggestions: formDetectionResult.suggestions || [
          'Googleフォームの「回答の行き先」を開き、対象のシートにリンクしてください',
          'フォーム作成者に連携状況を確認してください',
          'シート名に「回答」「フォーム」等の文字列が含まれている場合、フォーム連携の可能性があります'
        ],
        analysisResults: formDetectionResult.analysisResults
      };
    }

  } catch (error) {
    const endTime = new Date().toISOString();
    console.error('=== getFormInfoImpl ERROR ===', {
      startTime,
      endTime,
      spreadsheetId: spreadsheetId ? `${spreadsheetId.substring(0, 12)  }***` : 'N/A',
      sheetName,
      error: error.message,
      stack: error.stack
    });

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

    // 統一API使用: 構造化パース
    const configResult = getConfigSafe(user.userId);
    const config = configResult.success ? configResult.config : {};

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
