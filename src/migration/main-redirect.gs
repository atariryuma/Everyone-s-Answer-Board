/**
 * @fileoverview メインリダイレクト - 既存のmain.gsから新実装への段階的移行
 * 
 * このファイルは既存のmain.gsのdoGet/doPost関数を
 * 新しい実装に段階的にリダイレクトするためのものです。
 */

/**
 * 移行フラグ - trueにすると新実装を使用
 */
const USE_NEW_ARCHITECTURE = true;

/**
 * 既存のdoGet関数の保存（バックアップ）
 */
const doGetLegacy = typeof doGet === 'function' ? doGet : null;

/**
 * 既存のdoPost関数の保存（バックアップ）
 */
const doPostLegacy = typeof doPost === 'function' ? doPost : null;

/**
 * 新しいdoGet実装（main-new.gsから）
 */
function doGetNew(e) {
  try {
    // システム初期化チェック
    if (!isSystemInitialized()) {
      return showSetupPage();
    }
    
    // ユーザー認証チェック
    const user = authenticateUser();
    if (!user) {
      return showLoginPage();
    }
    
    // リクエストパラメータ解析
    const params = parseRequestParams(e);
    
    // ルーティング
    switch (params.mode) {
      case 'setup':
        return showAppSetupPage(params.userId);
      case 'admin':
        return showAdminPanel(user, params);
      case 'unpublished':
        return showUnpublishedPage(params);
      default:
        return showMainPage(user, params);
    }
    
  } catch (error) {
    logError(error, 'doGetNew', ERROR_SEVERITY.CRITICAL, ERROR_CATEGORIES.SYSTEM);
    return showErrorPage(error);
  }
}

/**
 * 新しいdoPost実装
 */
function doPostNew(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    
    // APIルーティング
    const result = handleCoreApiRequest(action, params);
    
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    logError(error, 'doPostNew', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM);
    return ContentService
      .createTextOutput(JSON.stringify({ 
        success: false, 
        error: error.message 
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * メインエントリーポイント - 条件付きリダイレクト
 */
function doGet(e) {
  try {
    if (USE_NEW_ARCHITECTURE) {
      debugLog('🚀 Using NEW architecture for doGet');
      return doGetNew(e);
    } else {
      debugLog('📦 Using LEGACY architecture for doGet');
      return doGetLegacy ? doGetLegacy(e) : showErrorPage(new Error('Legacy doGet not found'));
    }
  } catch (error) {
    errorLog('Fatal error in doGet:', error);
    return showErrorPage(error);
  }
}

/**
 * POSTエントリーポイント - 条件付きリダイレクト
 */
function doPost(e) {
  try {
    if (USE_NEW_ARCHITECTURE) {
      debugLog('🚀 Using NEW architecture for doPost');
      return doPostNew(e);
    } else {
      debugLog('📦 Using LEGACY architecture for doPost');
      return doPostLegacy ? doPostLegacy(e) : ContentService
        .createTextOutput(JSON.stringify({ 
          success: false, 
          error: 'Legacy doPost not found' 
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    errorLog('Fatal error in doPost:', error);
    return ContentService
      .createTextOutput(JSON.stringify({ 
        success: false, 
        error: error.message 
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 段階的移行のためのヘルパー関数
 */
function isNewArchitectureEnabled() {
  // PropertiesServiceから設定を読み取ることも可能
  try {
    const props = PropertiesService.getScriptProperties();
    const useNew = props.getProperty('USE_NEW_ARCHITECTURE');
    return useNew === 'true' || USE_NEW_ARCHITECTURE;
  } catch (error) {
    return USE_NEW_ARCHITECTURE;
  }
}

/**
 * 移行状態をログに記録
 */
function logMigrationStatus() {
  const status = {
    newArchitectureEnabled: isNewArchitectureEnabled(),
    timestamp: new Date().toISOString(),
    hasLegacyDoGet: !!doGetLegacy,
    hasLegacyDoPost: !!doPostLegacy,
    services: {
      errorService: typeof logError === 'function',
      cacheService: typeof getCacheValue === 'function',
      databaseService: typeof getUserInfo === 'function',
      apiService: typeof handleCoreApiRequest === 'function'
    }
  };
  
  console.log('Migration Status:', JSON.stringify(status, null, 2));
  return status;
}

/**
 * 安全な移行テスト
 * 新旧両方のアーキテクチャをテストして結果を比較
 */
function testMigration(testParams) {
  const results = {
    legacy: null,
    new: null,
    match: false
  };
  
  try {
    // レガシーアーキテクチャのテスト
    if (doGetLegacy) {
      USE_NEW_ARCHITECTURE = false;
      results.legacy = doGet(testParams);
    }
    
    // 新アーキテクチャのテスト
    USE_NEW_ARCHITECTURE = true;
    results.new = doGetNew(testParams);
    
    // 結果の比較（簡易的）
    results.match = JSON.stringify(results.legacy) === JSON.stringify(results.new);
    
  } catch (error) {
    results.error = error.message;
  } finally {
    // デフォルトに戻す
    USE_NEW_ARCHITECTURE = true;
  }
  
  return results;
}