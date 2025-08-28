/**
 * @fileoverview 回復力のある実行機構 - 指数バックオフとエラーハンドリング
 * Google Apps Script用の耐障害性実行ユーティリティ
 */

/**
 * 回復力のある実行機構クラス
 * 指数バックオフ、回路ブレーカー、グレースフルデグラデーションを提供
 */
class ResilientExecutor {
  constructor(options = {}) {
    this.config = {
      maxRetries: options.maxRetries || 3,
      baseDelay: options.baseDelay || 1000, // 1秒
      maxDelay: options.maxDelay || 30000, // 30秒
      backoffFactor: options.backoffFactor || 2,
      jitterEnabled: options.jitterEnabled !== false, // デフォルトtrue
      timeoutMs: options.timeoutMs || 120000, // 2分
    };

    // 回路ブレーカー状態管理
    this.circuitBreaker = {
      state: 'CLOSED', // CLOSED, OPEN, HALF_OPEN
      failures: 0,
      lastFailureTime: 0,
      successThreshold: 3, // HALF_OPENでの成功回数しきい値
      failureThreshold: 5, // OPENになる連続失敗回数
      openTimeout: 60000, // OPEN状態の継続時間（1分）
    };

    // 実行統計
    this.stats = {
      totalExecutions: 0,
      successfulExecutions: 0,
      failedExecutions: 0,
      retriedExecutions: 0,
      circuitBreakerActivations: 0,
    };
  }

  /**
   * 回復力のある実行（指数バックオフ付き）
   * @param {function} operation - 実行する操作
   * @param {object} options - オプション
   * @returns {Promise<any>} 実行結果
   */
  async execute(operation, options = {}) {
    const operationName = options.name || 'UnnamedOperation';
    const isIdempotent = options.idempotent !== false; // デフォルトtrue
    const fallback = options.fallback || null;

    this.stats.totalExecutions++;

    // 回路ブレーカーチェック
    if (this.circuitBreaker.state === 'OPEN') {
      const timeSinceLastFailure = Date.now() - this.circuitBreaker.lastFailureTime;
      if (timeSinceLastFailure < this.circuitBreaker.openTimeout) {
        throw new Error(
          `Circuit breaker is OPEN for ${operationName}. Remaining: ${this.circuitBreaker.openTimeout - timeSinceLastFailure}ms`
        );
      } else {
        this.circuitBreaker.state = 'HALF_OPEN';
        this.circuitBreaker.failures = 0;
        debugLog(`[ResilientExecutor] Circuit breaker for ${operationName} moved to HALF_OPEN`);
      }
    }

    let lastError = null;

    for (let attempt = 0; attempt <= this.config.maxRetries; attempt++) {
      try {
        debugLog(
          `[ResilientExecutor] Executing ${operationName} (attempt ${attempt + 1}/${this.config.maxRetries + 1})`
        );

        const result = this.executeWithTimeout(operation, this.config.timeoutMs);

        // 成功時の回路ブレーカー処理
        this.handleSuccess(operationName);

        if (attempt > 0) {
          this.stats.retriedExecutions++;
        }
        this.stats.successfulExecutions++;

        return result;
      } catch (error) {
        lastError = error;

        // リトライ不可能なエラーかチェック
        if (!this.isRetryableError(error) || !isIdempotent) {
          debugLog(
            `[ResilientExecutor] Non-retryable error for ${operationName}: ${error.message}`
          );
          break;
        }

        // 最後の試行でない場合は待機
        if (attempt < this.config.maxRetries) {
          const delay = this.calculateBackoffDelay(attempt);
          debugLog(
            `[ResilientExecutor] ${operationName} failed (attempt ${attempt + 1}). Retrying in ${delay}ms: ${error.message}`
          );
          this.sleep(delay);
        }
      }
    }

    // 全試行失敗時の処理
    this.handleFailure(operationName, lastError);
    this.stats.failedExecutions++;

    // フォールバック実行
    if (fallback && typeof fallback === 'function') {
      try {
        debugLog(`[ResilientExecutor] Executing fallback for ${operationName}`);
        return fallback(lastError);
      } catch (fallbackError) {
        warnLog(
          `[ResilientExecutor] Fallback failed for ${operationName}: ${fallbackError.message}`
        );
      }
    }

    throw lastError;
  }

  /**
   * タイムアウト付き実行
   * @param {function} operation - 実行する操作
   * @param {number} timeoutMs - タイムアウト（ミリ秒）
   * @returns {Promise<any>} 実行結果
   */
  async executeWithTimeout(operation, timeoutMs) {
    return new Promise((resolve, reject) => {
      const timeoutId = setTimeout(() => {
        reject(new Error(`Operation timed out after ${timeoutMs}ms`));
      }, timeoutMs);

      Promise.resolve()
        .then(() => operation())
        .then((result) => {
          clearTimeout(timeoutId);
          resolve(result);
        })
        .catch((error) => {
          clearTimeout(timeoutId);
          reject(error);
        });
    });
  }

  /**
   * 指数バックオフの遅延時間を計算
   * @param {number} attempt - 試行回数（0から開始）
   * @returns {number} 遅延時間（ミリ秒）
   */
  calculateBackoffDelay(attempt) {
    let delay = this.config.baseDelay * Math.pow(this.config.backoffFactor, attempt);
    delay = Math.min(delay, this.config.maxDelay);

    // ジッター追加（雷鳴群集効果を回避）
    if (this.config.jitterEnabled) {
      delay = delay * (0.5 + Math.random() * 0.5);
    }

    return Math.floor(delay);
  }

  /**
   * リトライ可能エラーかどうかを判定
   * @param {Error} error - エラーオブジェクト
   * @returns {boolean} リトライ可能かどうか
   */
  isRetryableError(error) {
    const retryableMessages = [
      'Service invoked too many times',
      'Rate Limit Exceeded',
      'User Rate Limit Exceeded',
      'Quota Error: User Rate Limit Exceeded',
      'Service error: Spreadsheets',
      'Internal error. Please try again',
      'timeout',
      'temporarily unavailable',
      'Service Unavailable',
      'Bad Gateway',
      'Gateway Timeout',
    ];

    const errorMessage = error.message || error.toString();
    return retryableMessages.some((msg) => errorMessage.toLowerCase().includes(msg.toLowerCase()));
  }

  /**
   * 成功時の回路ブレーカー処理
   * @param {string} operationName - 操作名
   */
  handleSuccess(operationName) {
    if (this.circuitBreaker.state === 'HALF_OPEN') {
      this.circuitBreaker.failures = 0;
      this.circuitBreaker.state = 'CLOSED';
      debugLog(`[ResilientExecutor] Circuit breaker for ${operationName} moved to CLOSED`);
    }
  }

  /**
   * 失敗時の回路ブレーカー処理
   * @param {string} operationName - 操作名
   * @param {Error} error - エラー
   */
  handleFailure(operationName, error) {
    this.circuitBreaker.failures++;
    this.circuitBreaker.lastFailureTime = Date.now();

    if (this.circuitBreaker.failures >= this.circuitBreaker.failureThreshold) {
      this.circuitBreaker.state = 'OPEN';
      this.stats.circuitBreakerActivations++;
      warnLog(
        `[ResilientExecutor] Circuit breaker for ${operationName} moved to OPEN after ${this.circuitBreaker.failures} failures`
      );
    }
  }

  /**
   * 非同期sleep
   * @param {number} ms - 待機時間（ミリ秒）
   * @returns {Promise<void>}
   */
  sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  /**
   * 実行統計を取得
   * @returns {Object} 統計情報
   */
  getStats() {
    return {
      ...this.stats,
      circuitBreakerState: this.circuitBreaker.state,
      successRate:
        this.stats.totalExecutions > 0
          ? ((this.stats.successfulExecutions / this.stats.totalExecutions) * 100).toFixed(2) + '%'
          : '0%',
    };
  }

  /**
   * 統計をリセット
   */
  resetStats() {
    this.stats = {
      totalExecutions: 0,
      successfulExecutions: 0,
      failedExecutions: 0,
      retriedExecutions: 0,
      circuitBreakerActivations: 0,
    };
    this.circuitBreaker.failures = 0;
    this.circuitBreaker.state = 'CLOSED';
  }
}

/**
 * グローバルな回復力のある実行機構インスタンス
 */
const resilientExecutor = new ResilientExecutor({
  maxRetries: 3,
  baseDelay: 1000,
  maxDelay: 30000,
  jitterEnabled: true,
});

/**
 * 便利なヘルパー関数群
 */

/**
 * 指数バックオフ付きでURLフェッチを実行
 * @param {string} url - URL
 * @param {object} options - フェッチオプション
 * @returns {Promise<HTTPResponse>} レスポンス
 */
function resilientUrlFetch(url, options = {}) {
  // 同期実行でリトライ機能付き
  const maxRetries = 3;
  const baseDelay = 1000;

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      debugLog(`resilientUrlFetch: ${url} (試行 ${attempt + 1}/${maxRetries + 1})`);

      const response = UrlFetchApp.fetch(url, {
        ...options,
        muteHttpExceptions: true,
      });

      // レスポンス検証
      if (!response || typeof response.getResponseCode !== 'function') {
        throw new Error('無効なレスポンスオブジェクトが返されました');
      }

      const responseCode = response.getResponseCode();

      // 成功時またはリトライ不要なエラー
      if (responseCode >= 200 && responseCode < 300) {
        if (attempt > 0) {
          infoLog(`resilientUrlFetch: リトライで成功 ${url}`);
        }
        return response;
      }

      // 4xx エラーはリトライしない
      if (responseCode >= 400 && responseCode < 500) {
        warnLog(`resilientUrlFetch: クライアントエラー ${responseCode} - ${url}`);
        return response;
      }

      // 最後の試行でない場合は待機してリトライ
      if (attempt < maxRetries) {
        const delay = baseDelay * Math.pow(2, attempt);
        warnLog(`resilientUrlFetch: ${responseCode}エラー、${delay}ms後にリトライ - ${url}`);
        Utilities.sleep(delay);
      } else {
        warnLog(`resilientUrlFetch: 最大リトライ回数に達しました - ${url}`);
        return response;
      }
    } catch (error) {
      // 最後の試行でない場合はリトライ
      if (attempt < maxRetries) {
        const delay = baseDelay * Math.pow(2, attempt);
        warnLog(`resilientUrlFetch: エラー、${delay}ms後にリトライ - ${url}: ${error.message}`);
        Utilities.sleep(delay);
      } else {
        console.error('[ERROR]', `resilientUrlFetch: 最終的に失敗 - ${url}: ${error.message}`);
        throw error;
      }
    }
  }
}

/**
 * 回復力のあるSpreadsheetサービス呼び出し
 * @param {function} operation - Spreadsheet操作
 * @param {string} operationName - 操作名
 * @returns {Promise<any>} 操作結果
 */
function resilientSpreadsheetOperation(operation, operationName = 'SpreadsheetOperation') {
  return resilientExecutor.execute(operation, {
    name: operationName,
    idempotent: true, // ほとんどのSpreadsheet操作は冪等
  });
}

/**
 * 回復力のあるCacheService操作
 * @param {function} operation - Cache操作
 * @param {string} operationName - 操作名
 * @param {function} fallback - フォールバック関数
 * @returns {Promise<any>} 操作結果
 */
function resilientCacheOperation(operation, operationName = 'CacheOperation', fallback = null) {
  return resilientExecutor.execute(operation, {
    name: operationName,
    idempotent: true,
    fallback: fallback,
  });
}
