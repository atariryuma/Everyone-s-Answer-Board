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
class ErrorHandler {
  constructor() {
    this.sessionId = Utilities.getUuid();
    this.errorCount = 0;
    this.startTime = Date.now();
  }

  /**
   * 構造化エラーログを生成・出力（ULog統合版）
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
    
    // ULogを使用した統一ログ出力
    switch (severity) {
      case ERROR_SEVERITY.CRITICAL:
        console.error(errorInfo.message, errorInfo.metadata, category);
        break;
      case ERROR_SEVERITY.HIGH:
        console.error(errorInfo.message, errorInfo.metadata, category);
        break;
      case ERROR_SEVERITY.MEDIUM:
        console.warn(errorInfo.message, errorInfo.metadata, category);
        break;
      case ERROR_SEVERITY.LOW:
        console.log(errorInfo.message, errorInfo.metadata, category);
        break;
    }
    
    return errorInfo;
  }

  /**
   * セキュリティ関連エラーの特別処理
   * @param {Error|string} error - エラーオブジェクトまたはメッセージ
   * @param {string} context - エラー発生箇所
   * @param {object} securityDetails - セキュリティ関連詳細情報
   */
  logSecurityError(error, context, securityDetails = {}) {
    const securityMetadata = {
      ...securityDetails,
      userAgent: typeof Session !== 'undefined' ? Session.getActiveUser()?.getEmail() : 'unknown',
      ipAddress: 'not_available_in_gas', // GASでは取得不可
      timestamp: new Date().toISOString(),
      securityAlert: true
    };

    return this.logError(
      error, 
      context, 
      ERROR_SEVERITY.HIGH, 
      ERROR_CATEGORIES.AUTHORIZATION, 
      securityMetadata
    );
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
   * パフォーマンス関連警告の記録
   * @param {string} operation - 操作名
   * @param {number} duration - 実行時間（ミリ秒）
   * @param {number} threshold - 閾値（ミリ秒）
   * @param {object} details - 詳細情報
   */
  logPerformanceWarning(operation, duration, threshold, details = {}) {
    if (duration <= threshold) {
      return; // 閾値以下の場合は記録しない
    }

    const perfMetadata = {
      operation: operation,
      duration: duration,
      threshold: threshold,
      overThresholdBy: duration - threshold,
      ...details,
      timestamp: new Date().toISOString()
    };

    return this.logError(
      `Performance threshold exceeded: ${operation} took ${duration}ms (threshold: ${threshold}ms)`,
      `performance.${operation}`,
      ERROR_SEVERITY.MEDIUM,
      ERROR_CATEGORIES.SYSTEM,
      perfMetadata
    );
  }

  /**
   * エラー統計を取得
   * @returns {object} エラー統計情報
   */
  getErrorStats() {
    return {
      sessionId: this.sessionId,
      totalErrors: this.errorCount,
      uptime: Date.now() - this.startTime,
      errorRate: this.errorCount / ((Date.now() - this.startTime) / 1000 / 60), // エラー/分
      timestamp: new Date().toISOString()
    };
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
const globalErrorHandler = new ErrorHandler();

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
 * セキュリティエラーをログに記録
 */
function logSecurityError(error, context, securityDetails) {
  return globalErrorHandler.logSecurityError(error, context, securityDetails);
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

/**
 * パフォーマンス警告をログに記録
 */
function logPerformanceWarning(operation, duration, threshold, details) {
  return globalErrorHandler.logPerformanceWarning(operation, duration, threshold, details);
}

/**
 * エラー統計を取得
 */
function getErrorStats() {
  return globalErrorHandler.getErrorStats();
}