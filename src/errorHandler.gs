/**
 * @fileoverview 統一エラーハンドリングシステム
 * アプリケーション全体の一貫したエラー処理とログ出力を提供
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
 * 統一エラーハンドラークラス
 */
class UnifiedErrorHandler {
  constructor() {
    this.sessionId = Utilities.getUuid();
    this.errorCount = 0;
    this.startTime = Date.now();
  }

  /**
   * 構造化エラーログを生成・出力
   * @param {Error|string} error - エラーオブジェクトまたはメッセージ
   * @param {string} context - エラー発生箇所/関数名
   * @param {string} severity - エラーの重要度
   * @param {string} category - エラーカテゴリ
   * @param {object} metadata - 追加メタデータ
   * @returns {object} 構造化エラー情報
   */
  logError(error, context, severity = ERROR_SEVERITY.MEDIUM, category = ERROR_CATEGORIES.SYSTEM, metadata = {}) {
    this.errorCount++;

    const errorInfo = this._buildErrorInfo(error, context, severity, category, metadata);

    // 重要度に応じた出力方法
    switch (severity) {
      case ERROR_SEVERITY.CRITICAL:
        console.error(`🚨 CRITICAL ERROR [${context}]:`, JSON.stringify(errorInfo, null, 2));
        break;
      case ERROR_SEVERITY.HIGH:
        console.error(`❌ HIGH SEVERITY [${context}]:`, JSON.stringify(errorInfo, null, 2));
        break;
      case ERROR_SEVERITY.MEDIUM:
        console.warn(`⚠️ MEDIUM SEVERITY [${context}]:`, errorInfo.message, errorInfo.metadata);
        break;
      case ERROR_SEVERITY.LOW:
        console.log(`ℹ️ LOW SEVERITY [${context}]:`, errorInfo.message);
        break;
    }

    return errorInfo;
  }

  /**
   * データベース操作エラーの処理
   * @param {Error|string} error - エラーオブジェクトまたはメッセージ
   * @param {string} operation - データベース操作の種類
   * @param {object} operationDetails - 操作詳細
   */
  logDatabaseError(error, operation, operationDetails = {}) {
    const dbMetadata = {
      operation: operation,
      ...operationDetails,
      retryable: this._isRetryableError(error),
      timestamp: new Date().toISOString()
    };

    return this.logError(
      error,
      `database.${operation}`,
      ERROR_SEVERITY.MEDIUM,
      ERROR_CATEGORIES.DATABASE,
      dbMetadata
    );
  }

  /**
   * バリデーションエラーの処理
   * @param {string} field - バリデーション対象フィールド
   * @param {string} value - 入力値
   * @param {string} rule - バリデーションルール
   * @param {string} message - エラーメッセージ
   */
  logValidationError(field, value, rule, message) {
    const validationMetadata = {
      field: field,
      value: typeof value === 'string' ? value.substring(0, 100) : String(value).substring(0, 100), // 値は100文字まで
      rule: rule,
      timestamp: new Date().toISOString()
    };

    return this.logError(
      message,
      `validation.${field}`,
      ERROR_SEVERITY.LOW,
      ERROR_CATEGORIES.VALIDATION,
      validationMetadata
    );
  }

  /**
   * エラー情報オブジェクトを構築
   * @private
   */
  _buildErrorInfo(error, context, severity, category, metadata) {
    const baseInfo = {
      timestamp: new Date().toISOString(),
      sessionId: this.sessionId,
      errorId: Utilities.getUuid(),
      context: context,
      severity: severity,
      category: category,
      errorNumber: this.errorCount,
      uptime: Date.now() - this.startTime
    };

    // エラーオブジェクトの場合
    if (error instanceof Error) {
      return {
        ...baseInfo,
        message: error.message,
        name: error.name,
        stack: error.stack,
        metadata: metadata
      };
    }

    // 文字列の場合
    return {
      ...baseInfo,
      message: String(error),
      metadata: metadata
    };
  }

  /**
   * エラーがリトライ可能かどうかを判定
   * @private
   */
  _isRetryableError(error) {
    if (!error) return false;

    const retryablePatterns = [
      /timeout/i,
      /rate limit/i,
      /service unavailable/i,
      /temporary/i,
      /quota/i
    ];

    const errorMessage = error.message || String(error);
    return retryablePatterns.some(pattern => pattern.test(errorMessage));
  }
}

// グローバルインスタンス
const globalErrorHandler = new UnifiedErrorHandler();

/**
 * 便利な関数群 - グローバルエラーハンドラーのラッパー
 */

/**
 * 一般的なエラーをログに記録
 */
function logError(error, context, severity, category, metadata) {
  return globalErrorHandler.logError(error, context, severity, category, metadata);
}

/**
 * データベースエラーをログに記録
 */
function logDatabaseError(error, operation, operationDetails) {
  return globalErrorHandler.logDatabaseError(error, operation, operationDetails);
}

/**
 * バリデーションエラーをログに記録
 */
function logValidationError(field, value, rule, message) {
  return globalErrorHandler.logValidationError(field, value, rule, message);
}