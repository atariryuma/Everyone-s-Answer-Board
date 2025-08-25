/**
 * @fileoverview エラーハンドリングサービス - シンプルで統一されたエラー管理
 */

/**
 * エラーの重要度レベル
 */
const ERROR_SEVERITY = {
  LOW: 'low',
  MEDIUM: 'medium',
  HIGH: 'high',
  CRITICAL: 'critical'
};

/**
 * エラーカテゴリ
 */
const ERROR_CATEGORIES = {
  AUTHENTICATION: 'authentication',
  AUTHORIZATION: 'authorization',
  DATABASE: 'database',
  CACHE: 'cache',
  NETWORK: 'network',
  VALIDATION: 'validation',
  SYSTEM: 'system',
  USER_INPUT: 'user_input'
};

/**
 * シンプルな統一エラーログ関数
 * @param {Error|string} error - エラーオブジェクトまたはメッセージ
 * @param {string} context - エラー発生箇所
 * @param {string} [severity='medium'] - エラーの重要度
 * @param {string} [category='system'] - エラーカテゴリ
 * @param {Object} [metadata={}] - 追加メタデータ
 */
function logError(error, context, severity = ERROR_SEVERITY.MEDIUM, category = ERROR_CATEGORIES.SYSTEM, metadata = {}) {
  const errorMessage = error instanceof Error ? error.message : String(error);
  const errorStack = error instanceof Error ? error.stack : undefined;
  
  const errorInfo = {
    timestamp: new Date().toISOString(),
    context: context,
    severity: severity,
    category: category,
    message: errorMessage,
    stack: errorStack,
    metadata: metadata
  };
  
  // 重要度に応じた出力
  if (severity === ERROR_SEVERITY.CRITICAL || severity === ERROR_SEVERITY.HIGH) {
    console.error(`[${severity.toUpperCase()}] ${context}:`, errorMessage, metadata);
  } else if (severity === ERROR_SEVERITY.MEDIUM) {
    console.warn(`[${severity.toUpperCase()}] ${context}:`, errorMessage);
  } else {
    console.log(`[${severity.toUpperCase()}] ${context}:`, errorMessage);
  }
  
  // クリティカルエラーの場合は管理者に通知（必要に応じて）
  if (severity === ERROR_SEVERITY.CRITICAL) {
    try {
      logCriticalErrorToSheet(errorInfo);
    } catch (logError) {
      console.error('Failed to log critical error to sheet:', logError);
    }
  }
  
  return errorInfo;
}

/**
 * クリティカルエラーをスプレッドシートに記録
 * @private
 */
function logCriticalErrorToSheet(errorInfo) {
  try {
    const dbSpreadsheetId = PropertiesService.getScriptProperties().getProperty('DATABASE_SPREADSHEET_ID');
    if (!dbSpreadsheetId) return;
    
    const spreadsheet = SpreadsheetApp.openById(dbSpreadsheetId);
    let errorSheet = spreadsheet.getSheetByName('CriticalErrors');
    
    if (!errorSheet) {
      errorSheet = spreadsheet.insertSheet('CriticalErrors');
      errorSheet.getRange(1, 1, 1, 5).setValues([['Timestamp', 'Context', 'Message', 'Category', 'Metadata']]);
    }
    
    errorSheet.appendRow([
      errorInfo.timestamp,
      errorInfo.context,
      errorInfo.message,
      errorInfo.category,
      JSON.stringify(errorInfo.metadata)
    ]);
  } catch (e) {
    // ログ記録の失敗は無視
  }
}

/**
 * データベースエラーのログ
 */
function logDatabaseError(error, operation, details = {}) {
  return logError(error, `database.${operation}`, ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.DATABASE, details);
}

/**
 * 認証エラーのログ
 */
function logAuthError(error, operation, details = {}) {
  return logError(error, `auth.${operation}`, ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.AUTHENTICATION, details);
}

/**
 * ネットワークエラーのログ
 */
function logNetworkError(error, operation, details = {}) {
  return logError(error, `network.${operation}`, ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.NETWORK, details);
}

/**
 * デバッグログ（開発環境のみ）
 */
function debugLog(...args) {
  if (typeof shouldEnableDebugMode === 'function' && shouldEnableDebugMode()) {
    console.log('[DEBUG]', ...args);
  }
}

/**
 * 情報ログ
 */
function infoLog(...args) {
  console.log('[INFO]', ...args);
}

/**
 * 警告ログ
 */
function warnLog(...args) {
  console.warn('[WARN]', ...args);
}

/**
 * エラーログ（簡易版）
 */
function errorLog(...args) {
  console.error('[ERROR]', ...args);
}