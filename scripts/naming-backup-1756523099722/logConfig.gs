/**
 * ログシステム初期化設定
 * アプリケーション起動時にULogシステムを自動設定
 */

/**
 * ログシステムの自動初期化
 * 環境に応じて適切なログレベルを設定
 */
function initializeLogging() {
  try {
    // 環境判定
    const isProduction = determineEnvironment();
    
    if (isProduction) {
      ULog.setProductionMode();
    } else {
      ULog.setDevelopmentMode();
    }
    
    ULog.info('ログシステム初期化完了', { 
      environment: isProduction ? 'production' : 'development',
      level: isProduction ? 'ERROR' : 'DEBUG'
    }, ULog.CATEGORIES.SYSTEM);
    
  } catch (error) {
    // フォールバック：基本的なコンソール出力
    console.error('ログシステム初期化エラー:', error.message);
    // 安全のため本番モードに設定
    ULog.setProductionMode();
  }
}

/**
 * 環境判定（本番 vs 開発）
 * @returns {boolean} 本番環境の場合true
 */
function determineEnvironment() {
  try {
    // 1. 開発者メールチェック
    const userEmail = Session.getActiveUser().getEmail();
    const isDeveloper = userEmail && (
      userEmail.includes('ryuma') || 
      userEmail.includes('dev') ||
      userEmail.includes('test')
    );
    
    // 2. PropertiesServiceから環境設定チェック
    const envSetting = PropertiesService.getScriptProperties().getProperty('ENVIRONMENT');
    if (envSetting === 'development') {
      return false;
    }
    
    // 3. デフォルトは本番環境（安全重視）
    return !isDeveloper;
    
  } catch (error) {
    // エラー時は本番環境扱い（安全重視）
    return true;
  }
}

/**
 * 既存のconsole.log呼び出しを段階的にULogに移行するヘルパー
 * @deprecated 新しいコードではULogを直接使用してください
 */
function migrateConsoleToULog() {
  // この関数は移行期間中のみ使用
  // 段階的移行が完了したら削除予定
  ULog.warn('console.logからULogへの移行を検討してください', {}, ULog.CATEGORIES.SYSTEM);
}

// アプリケーション起動時の自動初期化
// GAS環境での実行を考慮した条件付き初期化
if (typeof global === 'undefined' || !global.logInitialized) {
  try {
    initializeLogging();
    if (typeof global !== 'undefined') {
      global.logInitialized = true;
    }
  } catch (e) {
    // 初期化エラー時も処理を継続
    console.warn('ログ自動初期化をスキップ:', e.message);
  }
}