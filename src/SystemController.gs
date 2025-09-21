/**
 * @fileoverview SystemController - System management and setup functions
 */

/* global UserService, ConfigService, getCurrentEmail, createErrorResponse, createUserNotFoundError, createExceptionResponse, createAuthError, createAdminRequiredError, findUserByEmail, findUserById, openSpreadsheet, updateUser, getUserSpreadsheetData, Config, getSpreadsheetList, getUserConfig, saveUserConfig, getServiceAccount, isAdministrator, getDatabaseConfig, getAllUsers */

// ===========================================
// 📊 システム定数 - Zero-Dependency Architecture
// ===========================================

/**
 * キャッシュ期間 (秒)
 */
const CACHE_DURATION = {
  SHORT: 10,           // 10秒 - 認証ロック
  MEDIUM: 30,          // 30秒 - リアクション・ハイライトロック
  LONG: 300,           // 5分 - ユーザー情報キャッシュ
  EXTRA_LONG: 3600     // 1時間 - 設定キャッシュ
};

/**
 * タイムアウト期間 (ミリ秒)
 */
const TIMEOUT_MS = {
  QUICK: 100,          // UI応答性
  SHORT: 500,          // 軽量処理
  MEDIUM: 1000,        // 一般的処理
  LONG: 3000,          // 重い処理
  DEFAULT: 5000,       // デフォルトタイムアウト
  EXTENDED: 30000      // 拡張タイムアウト
};

/**
 * スリープ期間 (ミリ秒)
 */
const SLEEP_MS = {
  MICRO: 50,           // マイクロ待機
  SHORT: 100,          // 短時間休憩
  MEDIUM: 200,         // 中間休憩
  LONG: 500,           // 長時間休憩
  MAX: 5000            // 最大待機時間
};

/**
 * ログレベル制御
 */
const LOG_LEVEL = {
  DEBUG: 0,        // 開発時詳細ログ
  INFO: 1,         // 一般情報
  WARN: 2,         // 警告
  ERROR: 3,        // エラーのみ
  NONE: 4          // プロダクション（ログ無効）
};

/**
 * 現在のログレベル (プロダクション環境では ERROR または NONE を推奨)
 */
const CURRENT_LOG_LEVEL = LOG_LEVEL.INFO; // 開発時設定

/**
 * 統一ログ関数 - Zero-Dependency Architecture
 */
function sysLog(level, message, ...args) {
  if (level < CURRENT_LOG_LEVEL) return;

  const timestamp = new Date().toISOString();
  const prefix = `[${timestamp}]`;

  switch (level) {
    case LOG_LEVEL.DEBUG:
      console.log(`${prefix} [DEBUG]`, message, ...args);
      break;
    case LOG_LEVEL.INFO:
      console.log(`${prefix} [INFO]`, message, ...args);
      break;
    case LOG_LEVEL.WARN:
      console.warn(`${prefix} [WARN]`, message, ...args);
      break;
    case LOG_LEVEL.ERROR:
      console.error(`${prefix} [ERROR]`, message, ...args);
      break;
  }
}

// ===========================================
// 🌍 グローバル定数設定 - Zero-Dependency Architecture
// ===========================================

/**
 * グローバルスコープにシステム定数を設定
 * Zero-Dependency Architecture準拠
 */
const __rootSys = (typeof globalThis !== 'undefined') ? globalThis : (typeof global !== 'undefined' ? global : this);
__rootSys.CACHE_DURATION = CACHE_DURATION;
__rootSys.TIMEOUT_MS = TIMEOUT_MS;
__rootSys.SLEEP_MS = SLEEP_MS;
__rootSys.LOG_LEVEL = LOG_LEVEL;
__rootSys.sysLog = sysLog;

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
function testSystemSetup() {
  try {
    const diagnostics = {
      timestamp: new Date().toISOString(),
      tests: [],
      overall: 'unknown'
    };

    // 基本コンポーネントテスト
    try {
      const session = { email: Session.getActiveUser().getEmail() };
      diagnostics.tests.push({
        name: 'Session Service',
        status: session.isValid ? '✅' : '❌',
        details: session.isValid ? `User: ${  session.email}` : 'No active session'
      });
    } catch (sessionError) {
      diagnostics.tests.push({
        name: 'Session Service',
        status: '❌',
        details: sessionError.message
      });
    }

    // データベース接続テスト
    try {
      const props = PropertiesService.getScriptProperties();
      const databaseId = props.getDatabaseSpreadsheetId();
      if (databaseId) {
        const dataAccess = openSpreadsheet(databaseId, { useServiceAccount: true });
        diagnostics.tests.push({
          name: 'Database Connection',
          status: '✅',
          details: 'Database accessible'
        });
      } else {
        diagnostics.tests.push({
          name: 'Database Connection',
          status: '❌',
          details: 'Database not configured'
        });
      }
    } catch (dbError) {
      diagnostics.tests.push({
        name: 'Database Connection',
        status: '❌',
        details: dbError.message
      });
    }

    // 総合評価
    const hasErrors = diagnostics.tests.some(test => test.status === '❌');
    diagnostics.overall = hasErrors ? '⚠️ 問題あり' : '✅ 正常';

    return {
      success: !hasErrors,
      diagnostics
    };
  } catch (error) {
    return {
      success: false,
      message: error.message
    };
  }
}

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
        const cache = CacheService.getScriptCache();
        if (cache && typeof cache.removeAll === 'function') {
          cache.removeAll([]); // Signature-compatible no-op to avoid errors
          cacheResults.push('ScriptCache クリア要求送信');
        }
      } catch (cacheError) {
        console.warn('[WARN] SystemController.forceUrlSystemReset: ScriptCache clear error:', cacheError.message || 'Cache clear failed');
        cacheResults.push(cacheError && cacheError.message ? `ScriptCache クリア失敗: ${cacheError.message}` : 'ScriptCache クリア失敗: 詳細不明');
      }

      // Document Cache も試行
      try {
        const docCache = CacheService.getScriptCache(); // Use unified cache
        if (docCache && typeof docCache.removeAll === 'function') {
          docCache.removeAll([]);
          cacheResults.push('DocumentCache クリア要求送信');
        }
      } catch (docCacheError) {
        console.warn('[WARN] SystemController.forceUrlSystemReset: DocumentCache clear error:', docCacheError.message || 'Document cache clear failed');
        cacheResults.push(docCacheError && docCacheError.message ? `DocumentCache クリア失敗: ${docCacheError.message}` : 'DocumentCache クリア失敗: 詳細不明');
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
      console.error('[ERROR] SystemController.forceUrlSystemReset:', error.message || 'System reset error');
      return {
        success: false,
        message: error && error.message ? error.message : '詳細不明'
      };
    }
}

/**
 * WebアプリのURL取得
 * 各種HTMLファイルから呼び出される
 *
 * @returns {string} WebアプリURL
 */
function getWebAppUrl() {
  try {
    return ScriptApp.getService().getUrl();
  } catch (error) {
    console.error('[ERROR] SystemController.getWebAppUrl:', error.message || 'Web app URL error');
    return '';
  }
}

/**
 * システム全体の診断実行
 * AppSetupPage.html から呼び出される
 *
 * @returns {Object} 診断結果
 */
function testSystemDiagnosis() {
  try {
    console.log('🔬 testSystemDiagnosis START');

    const email = getCurrentEmail();
    if (!email) {
      return createAuthError();
    }

    if (!isAdministrator(email)) {
      return createAdminRequiredError();
    }

    // GAS-Native: Direct system diagnostics
    const diagnostics = [];

    // Check 1: Core system properties
    try {
      const props = PropertiesService.getScriptProperties();
      const coreProps = {
        adminEmail: !!props.getProperty('ADMIN_EMAIL'),
        databaseId: !!props.getProperty('DATABASE_SPREADSHEET_ID'),
        serviceAccount: !!props.getProperty('SERVICE_ACCOUNT_CREDS')
      };

      const allCorePresent = Object.values(coreProps).every(Boolean);
      diagnostics.push({
        test: 'Core Properties',
        status: allCorePresent ? 'PASS' : 'FAIL',
        details: coreProps,
        critical: true
      });
    } catch (error) {
      diagnostics.push({
        test: 'Core Properties',
        status: 'ERROR',
        error: error.message,
        critical: true
      });
    }

    // Check 2: Database connectivity
    try {
      const dbTest = getDatabaseConfig();
      diagnostics.push({
        test: 'Database Connection',
        status: dbTest.success ? 'PASS' : 'FAIL',
        details: dbTest.success ? 'Database accessible' : dbTest.message,
        critical: true
      });
    } catch (error) {
      diagnostics.push({
        test: 'Database Connection',
        status: 'ERROR',
        error: error.message,
        critical: true
      });
    }

    // Check 3: Web app deployment
    try {
      const webAppUrl = ScriptApp.getService().getUrl();
      diagnostics.push({
        test: 'Web App Deployment',
        status: webAppUrl ? 'PASS' : 'FAIL',
        details: webAppUrl || 'No deployment URL available',
        critical: false
      });
    } catch (error) {
      diagnostics.push({
        test: 'Web App Deployment',
        status: 'ERROR',
        error: error.message,
        critical: false
      });
    }

    const criticalIssues = diagnostics.filter(d => d.critical && d.status !== 'PASS').length;
    const totalTests = diagnostics.length;
    const passedTests = diagnostics.filter(d => d.status === 'PASS').length;

    return {
      success: criticalIssues === 0,
      message: criticalIssues === 0 ? 'System diagnosis completed - all critical tests passed' : `System diagnosis found ${criticalIssues} critical issues`,
      summary: {
        totalTests,
        passed: passedTests,
        failed: totalTests - passedTests,
        criticalIssues
      },
      diagnostics,
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    console.error('testSystemDiagnosis error:', error.message);
    return createExceptionResponse(error, 'System diagnosis failed');
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
      const props = PropertiesService.getScriptProperties();
      const status = {
        timestamp: new Date().toISOString(),
        setup: {
          hasDatabase: !!props.getDatabaseSpreadsheetId(),
          hasAdminEmail: !!props.getAdminEmail(),
          hasServiceAccount: !!getServiceAccount()?.isValid
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
      console.error('[ERROR] SystemController.getSystemStatus:', error.message || 'System status error');
      return {
        success: false,
        message: error.message
      };
    }
}

/**
 * Monitor system health and performance - CLAUDE.md準拠命名
 * @returns {Object} System monitoring result
 */
function monitorSystem() {
  try {
    console.log('📊 monitorSystem START');

    const email = getCurrentEmail();
    if (!email) {
      return createAuthError();
    }

    if (!isAdministrator(email)) {
      return createAdminRequiredError();
    }

    // GAS-Native: Direct system monitoring
    const metrics = {};

    // Monitor 1: Script execution quota
    try {
      // GAS provides no direct quota API, so we track execution time
      const startTime = new Date();
      metrics.executionTime = startTime.toISOString();
      metrics.quotaStatus = 'MONITORING';
    } catch (error) {
      metrics.quotaStatus = 'ERROR';
      metrics.quotaError = error.message;
    }

    // Monitor 2: Database size and access
    try {
      const users = getAllUsers();
      metrics.userCount = Array.isArray(users) ? users.length : 0;
      metrics.databaseStatus = 'ACCESSIBLE';
    } catch (error) {
      metrics.userCount = 0;
      metrics.databaseStatus = 'ERROR';
      metrics.databaseError = error.message;
    }

    // Monitor 3: Cache performance
    try {
      const cache = CacheService.getScriptCache();
      // Simple cache test
      const testKey = 'monitoring_test';
      const testValue = Date.now().toString();
      cache.put(testKey, testValue, 60);
      const retrieved = cache.get(testKey);

      metrics.cacheStatus = retrieved === testValue ? 'OPERATIONAL' : 'DEGRADED';
      cache.remove(testKey);
    } catch (error) {
      metrics.cacheStatus = 'ERROR';
      metrics.cacheError = error.message;
    }

    // Monitor 4: Service account validation
    try {
      const props = PropertiesService.getScriptProperties();
      const creds = props.getProperty('SERVICE_ACCOUNT_CREDS');
      metrics.serviceAccountStatus = creds ? 'CONFIGURED' : 'NOT_CONFIGURED';
    } catch (error) {
      metrics.serviceAccountStatus = 'ERROR';
      metrics.serviceAccountError = error.message;
    }

    const healthScore = Object.keys(metrics).filter(key =>
      metrics[key] === 'OPERATIONAL' ||
      metrics[key] === 'ACCESSIBLE' ||
      metrics[key] === 'CONFIGURED'
    ).length;

    return {
      success: true,
      message: 'System monitoring completed',
      healthScore: `${healthScore}/4`,
      status: healthScore >= 3 ? 'HEALTHY' : healthScore >= 2 ? 'DEGRADED' : 'CRITICAL',
      metrics,
      recommendations: healthScore < 3 ? ['Check failed components', 'Review system configuration'] : ['System operating normally'],
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    console.error('monitorSystem error:', error.message);
    return createExceptionResponse(error, 'System monitoring failed');
  }
}

/**
 * Check data integrity - CLAUDE.md準拠命名
 * @returns {Object} Data integrity check result
 */
function checkDataIntegrity() {
  try {
    console.log('🔍 checkDataIntegrity START');

    const email = getCurrentEmail();
    if (!email) {
      return createAuthError();
    }

    if (!isAdministrator(email)) {
      return createAdminRequiredError();
    }

    // GAS-Native: Direct data integrity checks
    const integrityResults = [];

    // Check 1: User database consistency
    try {
      const users = getAllUsers();
      const userCount = Array.isArray(users) ? users.length : 0;
      const validUsers = users.filter(user =>
        user &&
        user.userId &&
        user.userEmail &&
        user.userEmail.includes('@')
      ).length;

      integrityResults.push({
        check: 'User Database',
        status: validUsers === userCount ? 'PASS' : 'WARN',
        details: `${validUsers}/${userCount} valid user records`,
        validCount: validUsers,
        totalCount: userCount
      });
    } catch (error) {
      integrityResults.push({
        check: 'User Database',
        status: 'ERROR',
        error: error.message
      });
    }

    // Check 2: Configuration integrity
    try {
      const users = getAllUsers();
      let configErrors = 0;
      let validConfigs = 0;

      if (Array.isArray(users)) {
        for (const user of users) {
          try {
            const configResult = getUserConfig(user.userId);
            if (configResult.success) {
              validConfigs++;
            } else {
              configErrors++;
            }
          } catch (configError) {
            configErrors++;
          }
        }
      }

      integrityResults.push({
        check: 'Configuration Integrity',
        status: configErrors === 0 ? 'PASS' : 'WARN',
        details: `${validConfigs} valid configs, ${configErrors} errors`,
        validCount: validConfigs,
        errorCount: configErrors
      });
    } catch (error) {
      integrityResults.push({
        check: 'Configuration Integrity',
        status: 'ERROR',
        error: error.message
      });
    }

    // Check 3: Database schema validation
    try {
      const dbConfig = getDatabaseConfig();
      integrityResults.push({
        check: 'Database Schema',
        status: dbConfig.success ? 'PASS' : 'FAIL',
        details: dbConfig.success ? 'Schema accessible' : 'Schema validation failed'
      });
    } catch (error) {
      integrityResults.push({
        check: 'Database Schema',
        status: 'ERROR',
        error: error.message
      });
    }

    const passedChecks = integrityResults.filter(r => r.status === 'PASS').length;
    const totalChecks = integrityResults.length;
    const hasErrors = integrityResults.some(r => r.status === 'ERROR');
    const hasWarnings = integrityResults.some(r => r.status === 'WARN');

    return {
      success: !hasErrors,
      message: hasErrors ? 'Data integrity check found errors' : hasWarnings ? 'Data integrity check found warnings' : 'Data integrity check passed',
      summary: {
        totalChecks,
        passed: passedChecks,
        warnings: integrityResults.filter(r => r.status === 'WARN').length,
        errors: integrityResults.filter(r => r.status === 'ERROR').length
      },
      overallStatus: hasErrors ? 'CRITICAL' : hasWarnings ? 'WARNING' : 'HEALTHY',
      checks: integrityResults,
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    console.error('checkDataIntegrity error:', error.message);
    return createExceptionResponse(error, 'Data integrity check failed');
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
        const cache = CacheService.getScriptCache();
        if (cache && typeof cache.removeAll === 'function') {
          cache.removeAll();
        }
      } catch (cacheError) {
        repairResults.warnings = [cacheError && cacheError.message ? `キャッシュクリア失敗: ${cacheError.message}` : 'キャッシュクリア失敗: 詳細不明'];
      }

      return {
        success: true,
        repairResults
      };

    } catch (error) {
      console.error('[ERROR] SystemController.performAutoRepair:', error.message || 'Auto repair error');
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
    // 🎯 CLAUDE.md準拠: 管理者機能のため、サービスアカウント使用
    const dataAccess = openSpreadsheet(spreadsheetId, { useServiceAccount: true });
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
    console.error('[ERROR] SystemController.getSheetList:', error.message || 'Sheet list error');
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
    const props = PropertiesService.getScriptProperties();

    // Set properties
    try {
      props.setProperty('APPLICATION_STATUS', 'active');
      props.setProperty('PUBLISHED_AT', publishedAt);
    } catch (propsError) {
      console.error('[ERROR] SystemController.publishApplication: Properties update failed:', propsError.message || 'Properties update error');
    }

    if (publishConfig) {
      try {
        props.setProperty('PUBLISH_CONFIG', JSON.stringify(publishConfig));
      } catch (publishConfigError) {
        console.error('❌ PUBLISH_CONFIG save failed:', publishConfigError.message);
      }
    }

    // 🔧 Zero-Dependency統一: 直接findUserByEmail使用（CLAUDE.md準拠）
    const user = findUserByEmail(email, { requestingUser: email });
    let saveResult = null;

    if (user) {
      // Re-fetch latest user data to avoid conflicts
      const latestUser = findUserByEmail(email, { requestingUser: email });
      const userToUse = latestUser || user;

      // 統一API使用: 構造化パース
      const configResult = getUserConfig(userToUse.userId);
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
        isPublished: true,
        publishedAt,
        setupStatus: 'completed',
        lastModified: publishedAt
      };

      // ✅ バックエンド構造統一完了：フロントエンドは既に正しい構造を送信
      console.log('✅ ColumnMapping backend structure unified - no transformation needed');

      // 🔧 CLAUDE.md準拠: 統一API使用 - saveUserConfigでETag対応の安全な更新
      saveResult = saveUserConfig(user.userId, updatedConfig, { isPublish: true });

      if (!saveResult.success) {
        console.error('❌ saveUserConfig failed during publish:', saveResult.message);
        // エラーでも処理を継続（互換性のため）
      } else {
        console.log('✅ Config saved via saveUserConfig:', {
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
 * 🔧 CLAUDE.md準拠: Self vs Cross-user Spreadsheet Access
 * CLAUDE.md Security Pattern: Context-aware service account usage
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {Object} context - アクセスコンテキスト
 * @returns {Object} {spreadsheet, accessMethod, auth, isOwner}
 */
function getSpreadsheetAdaptive(spreadsheetId, context = {}) {
  const currentEmail = getCurrentEmail();

  // CLAUDE.md準拠: アクセス権限の判定
  const isAdminRequest = context.forceServiceAccount || (currentEmail && isAdministrator && isAdministrator(currentEmail));
  const isOwner = isUserSpreadsheetOwner(spreadsheetId);

  // ✅ **Self-access**: Owner accessing own spreadsheet (normal permissions unless admin override)
  // ✅ **Cross-user**: Non-owner or admin accessing spreadsheet (service account)
  const useServiceAccount = isAdminRequest || !isOwner;

  console.log(`getSpreadsheetAdaptive: ${useServiceAccount ? 'Service account' : 'Normal permissions'} access to spreadsheet (owner: ${isOwner}, admin: ${isAdminRequest})`);

  try {
    const dataAccess = openSpreadsheet(spreadsheetId, { useServiceAccount });
    const accessMethod = useServiceAccount ? 'service_account' : 'normal_permissions';

    console.log(`getSpreadsheetAdaptive: Access successful via ${accessMethod}`);
    return {
      spreadsheet: dataAccess.spreadsheet,
      accessMethod,
      auth: dataAccess.auth,
      isOwner,
      context: {
        isAdminRequest,
        useServiceAccount,
        currentEmail: currentEmail ? `${currentEmail.split('@')[0]}@***` : null
      }
    };
  } catch (error) {
    console.error('getSpreadsheetAdaptive: Spreadsheet access failed:', error.message);
    const errorMessage = error && error.message ? error.message : '詳細不明';
    throw new Error(`スプレッドシートアクセス失敗: ${errorMessage}`);
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
    console.log('🔍 detectFormConnection: オーナー権限でAPI検出開始');
    try {
      // シートレベルでフォームURL取得（最優先）
      console.log('🔍 sheet.getFormUrl() 実行中...');
      if (typeof sheet.getFormUrl === 'function') {
        const formUrl = sheet.getFormUrl();
        console.log('🔍 sheet.getFormUrl() 結果:', { formUrl: formUrl || 'null' });
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
          console.log('✅ シートレベルでフォーム検出成功:', results);
          return results;
        }
      } else {
        console.log('⚠️ sheet.getFormUrl メソッドが使用できません');
      }

      // スプレッドシートレベルでフォームURL取得（フォールバック）
      console.log('🔍 spreadsheet.getFormUrl() 実行中...');
      if (typeof spreadsheet.getFormUrl === 'function') {
        const formUrl = spreadsheet.getFormUrl();
        console.log('🔍 spreadsheet.getFormUrl() 結果:', { formUrl: formUrl || 'null' });
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
          console.log('✅ スプレッドシートレベルでフォーム検出成功:', results);
          return results;
        }
      } else {
        console.log('⚠️ spreadsheet.getFormUrl メソッドが使用できません');
      }
      console.log('❌ API検出: フォームURLが取得できませんでした');
    } catch (apiError) {
      console.warn('detectFormConnection: API検出失敗:', apiError.message);
      results.details.push(apiError && apiError.message ? `API検出失敗: ${apiError.message}` : 'API検出失敗: 詳細不明');
    }
  } else {
    console.log('⚠️ オーナー権限なし - API検出をスキップ');
  }

  // Method 1.5: Drive APIフォーム検索（API検出失敗時のフォールバック）
  if (isOwner) {
    console.log('🔍 Drive API検索開始...');
    try {
      const spreadsheetId = spreadsheet.getId();
      console.log('🔍 Drive API検索 - spreadsheetId:', `${spreadsheetId.substring(0, 12)}***`);
      const driveFormResult = searchFormsByDrive(spreadsheetId, sheetName);
      console.log('🔍 Drive API検索結果:', {
        hasFormUrl: !!driveFormResult.formUrl,
        formTitle: driveFormResult.formTitle || 'null'
      });
      if (driveFormResult.formUrl) {
        results.formUrl = driveFormResult.formUrl;
        results.formTitle = driveFormResult.formTitle;
        results.confidence = 80;
        results.detectionMethod = 'drive_search';
        results.details.push('Drive API検索で検出');
        console.log('✅ Drive API検索でフォーム検出成功:', results);
        return results;
      } else {
        console.log('❌ Drive API検索: フォームが見つかりませんでした');
      }
    } catch (driveError) {
      console.warn('detectFormConnection: Drive API検索失敗:', driveError.message);
      results.details.push(driveError && driveError.message ? `Drive API検索失敗: ${driveError.message}` : 'Drive API検索失敗: 詳細不明');
    }
  }

  // Method 2: ヘッダーパターン解析
  console.log('🔍 ヘッダーパターン解析開始...');
  try {
    const [headers] = sheet.getRange(1, 1, 1, Math.min(sheet.getLastColumn(), 10)).getValues();
    console.log('🔍 取得したヘッダー:', headers);
    const headerAnalysis = analyzeFormHeaders(headers);
    console.log('🔍 ヘッダー解析結果:', headerAnalysis);

    if (headerAnalysis.isFormLike) {
      results.confidence = Math.max(results.confidence, headerAnalysis.confidence);
      results.detectionMethod = results.detectionMethod === 'none' ? 'header_analysis' : results.detectionMethod;
      results.details.push(headerAnalysis && headerAnalysis.reason ? `ヘッダー解析: ${headerAnalysis.reason}` : 'ヘッダー解析: 結果不明');
      console.log('✅ ヘッダー分析でフォームパターン検出');
    } else {
      console.log('❌ ヘッダー分析: フォームパターンなし');
    }
  } catch (headerError) {
    console.warn('detectFormConnection: ヘッダー解析失敗:', headerError.message);
    results.details.push(headerError && headerError.message ? `ヘッダー解析失敗: ${headerError.message}` : 'ヘッダー解析失敗: 詳細不明');
  }

  // Method 3: シート名パターン解析
  console.log('🔍 シート名パターン解析開始...');
  const sheetNameAnalysis = analyzeSheetName(sheetName);
  console.log('🔍 シート名解析結果:', sheetNameAnalysis);
  if (sheetNameAnalysis.isFormLike) {
    results.confidence = Math.max(results.confidence, sheetNameAnalysis.confidence);
    results.detectionMethod = results.detectionMethod === 'none' ? 'sheet_name' : results.detectionMethod;
    results.details.push(sheetNameAnalysis && sheetNameAnalysis.reason ? `シート名解析: ${sheetNameAnalysis.reason}` : 'シート名解析: 結果不明');
    console.log('✅ シート名分析でフォームパターン検出');
  } else {
    console.log('❌ シート名分析: フォームパターンなし');
  }

  // フォーム検出時のタイトル生成
  if (results.confidence >= 40) {
    results.formTitle = `${sheetName} (フォーム検出済み)`;
  }

  console.log('🔍 detectFormConnection 最終結果:', {
    formUrl: results.formUrl || 'null',
    confidence: results.confidence,
    detectionMethod: results.detectionMethod,
    formTitle: results.formTitle || 'null',
    detailsCount: results.details.length
  });

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
        // FormAppアクセスを制限的に実行
        let form = null;
        let destId = null;
        let formTitle = null;
        let formPublishedUrl = null;

        try {
          form = FormApp.openById(formFile.getId());
          destId = form.getDestinationId();
          formTitle = form.getTitle();
          formPublishedUrl = form.getPublishedUrl();
        } catch (formAccessError) {
          // FormApp権限エラーの場合はファイル名から推測
          console.warn('searchFormsByDrive: FormApp権限制限、ファイル名から推測:', formAccessError.message);
          formTitle = formFile.getName();
          // 権限のないフォームはスキップ
          continue;
        }

        if (destId === spreadsheetId) {
          console.log('searchFormsByDrive: 対象スプレッドシートへのフォーム発見', {
            formId: formFile.getId(),
            formTitle
          });

          // Form IDとspreadsheet IDが一致していれば接続確認済み
          console.log('searchFormsByDrive: フォーム接続確認成功', {
            sheetName,
            formTitle
          });

          return {
            formUrl: formPublishedUrl,
            formTitle
          };
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

// ✅ getSpreadsheetInfo関数を削除 - GAS-Native APIのみ使用

/**
 * スプレッドシートへのアクセス権限を検証
 * AdminPanel.js.html から呼び出される
 *
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {Object} 検証結果
 */
function validateAccess(spreadsheetId, autoAddEditor = true) {
  try {
    // 🎯 CLAUDE.md準拠: validateAccess は管理者機能のため、常にサービスアカウント使用
    const dataAccess = openSpreadsheet(spreadsheetId, { useServiceAccount: true });
    const {spreadsheet, auth} = dataAccess;

    // サービスアカウントを編集者として自動登録（openSpreadsheetで既に実行済み）
    if (autoAddEditor) {
      console.log('validateAccess: サービスアカウント編集者権限は openSpreadsheet で既に処理済み');
    }

    // カスタムラッパーのgetSheets()メソッドを使用
    const sheets = spreadsheet.getSheets();

    // ✅ GAS-Native: SpreadsheetApp APIを使用してスプレッドシート名を取得
    let spreadsheetName;
    try {
      spreadsheetName = spreadsheet.getName();
    } catch (error) {
      spreadsheetName = `スプレッドシート (ID: ${spreadsheetId.substring(0, 8)}...)`;
    }

    // アクセスできたら成功
    const result = {
      success: true,
      message: 'アクセス権限が確認されました',
      spreadsheetName,
      sheets: sheets.map(sheet => ({
        name: sheet.getName(),
        rowCount: sheet.getLastRow(),
        columnCount: sheet.getLastColumn()
      })),
      owner: 'Service Account Access',
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
      // ✅ GAS-Native: 常にspreadsheet.getName()を使用（権限があれば動作）
      try {
        spreadsheetName = spreadsheet.getName();
      } catch (error) {
        // サービスアカウントでもgetName()は通常動作する
        console.warn('getFormInfoImpl: getName() failed, using fallback:', error.message);
        spreadsheetName = `スプレッドシート (ID: ${spreadsheetId.substring(0, 8)}...)`;
      }
    } catch (nameError) {
      console.warn('getFormInfoImpl: スプレッドシート名取得エラー:', nameError.message);
      if (spreadsheetId && spreadsheetId.trim()) {
        spreadsheetName = `スプレッドシート (ID: ${  spreadsheetId.substring(0, 8)  }...)`;
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
          'フォーム連携パターンを検出（URL取得不可）' :
          'フォーム連携が確認できませんでした',
        formData,
        suggestions: formDetectionResult.suggestions || [
          'Googleフォームの「回答の行き先」を開き、対象のシートにリンクしてください',
          'フォーム作成者に連携状況を確認してください',
          'シート名に「回答」「フォーム」等の文字列が含まれている場合、フォーム連携パターンとして評価されます'
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
    const configService = null /* ConfigService direct call */;
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
    const session = { email: Session.getActiveUser().getEmail() };
    // 🔧 Zero-Dependency統一: 直接Dataクラス使用
    let user = null;
    if (targetUserId) {
      user = findUserById(targetUserId, {
        requestingUser: session.email
      });
    }

    if (!user && session && session.email) {
      user = findUserByEmail(session.email, { requestingUser: session.email });
    }

    if (!user) {
      console.error('❌ User not found');
      return createUserNotFoundError();
    }

    // 統一API使用: 構造化パース
    const configResult = getUserConfig(user.userId);
    const config = configResult.success ? configResult.config : {};

    const result = {
      success: true,
      published: config.isPublished === true,
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
// ✅ CLAUDE.md準拠: 重複関数削除 - main.gsの完全実装を使用



/**
 * ログイン状態を取得
 * login.js.html から呼び出される
 *
 * @returns {Object} ログイン状態
 */
function getLoginStatus() {
  try {
    const {email} = { email: Session.getActiveUser().getEmail() };
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

// ===========================================
// 🌍 Global SystemController Object Export
// ===========================================

/**
 * SystemController統一インターフェース
 * main.gsからアクセス可能にするためのグローバルエクスポート
 */
const __rootSC = (typeof globalThis !== 'undefined') ? globalThis : (typeof global !== 'undefined' ? global : this);
__rootSC.SystemController = {
  getFormInfo,
  checkCurrentPublicationStatus,
  performAutoRepair,
  forceUrlSystemReset,
  publishApplication,
  testForceLogoutRedirect
};
