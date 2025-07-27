/**
 * StudyQuest GAS Error Handler
 * GAS固有エラーの検出、処理、回復メカニズム
 */

/**
 * GAS固有エラーの定義
 */
const GAS_ERROR_TYPES = {
  EXECUTION_TIMEOUT: 'execution_timeout',
  PERMISSION_DENIED: 'permission_denied',
  QUOTA_EXCEEDED: 'quota_exceeded',
  SCRIPT_NOT_FOUND: 'script_not_found',
  LOCK_FAILED: 'lock_failed',
  CONCURRENT_ACCESS: 'concurrent_access',
  SERVICE_UNAVAILABLE: 'service_unavailable',
  NETWORK_ERROR: 'network_error',
  AUTHENTICATION_FAILED: 'authentication_failed',
  RATE_LIMITED: 'rate_limited'
};

/**
 * GASエラーハンドラー
 */
class GASErrorHandler {
  constructor(coreState) {
    this.coreState = coreState;
    
    // エラーパターンの定義
    this.errorPatterns = this.initializeErrorPatterns();
    
    // 回復戦略の定義
    this.recoveryStrategies = this.initializeRecoveryStrategies();
    
    // エラー統計
    this.errorStats = {
      totalErrors: 0,
      errorsByType: new Map(),
      recoveryAttempts: 0,
      successfulRecoveries: 0,
      lastErrorTime: 0
    };
    
    // 回復制御
    this.recoveryControl = {
      maxRetryAttempts: 3,
      baseRetryDelay: 2000,
      maxRetryDelay: 30000,
      backoffMultiplier: 2,
      circuitBreakerThreshold: 5,
      circuitBreakerResetTime: 300000, // 5分
      isCircuitOpen: false,
      circuitOpenTime: 0
    };
    
    // 実行時間監視
    this.executionMonitor = {
      startTime: Date.now(),
      warningThreshold: STUDY_QUEST_CONSTANTS.GAS_EXECUTION_TIMEOUT_MS * 0.8,
      criticalThreshold: STUDY_QUEST_CONSTANTS.GAS_EXECUTION_TIMEOUT_MS * 0.95,
      warningIssued: false
    };
    
    this.setupGlobalErrorHandling();
  }

  /**
   * エラーパターンの初期化
   */
  initializeErrorPatterns() {
    return new Map([
      [GAS_ERROR_TYPES.EXECUTION_TIMEOUT, [
        /exceeded maximum execution time/i,
        /execution timeout/i,
        /script execution timeout/i,
        /maximum execution time exceeded/i
      ]],
      [GAS_ERROR_TYPES.PERMISSION_DENIED, [
        /permission denied/i,
        /access denied/i,
        /insufficient permissions/i,
        /authorization required/i
      ]],
      [GAS_ERROR_TYPES.QUOTA_EXCEEDED, [
        /quota exceeded/i,
        /service invoked too many times/i,
        /rate limit exceeded/i,
        /daily limit exceeded/i
      ]],
      [GAS_ERROR_TYPES.SCRIPT_NOT_FOUND, [
        /script function not found/i,
        /function .+ not found/i,
        /no function .+ found/i
      ]],
      [GAS_ERROR_TYPES.LOCK_FAILED, [
        /lock could not be obtained/i,
        /failed to acquire lock/i,
        /lock timeout/i,
        /concurrent access detected/i
      ]],
      [GAS_ERROR_TYPES.SERVICE_UNAVAILABLE, [
        /service unavailable/i,
        /service temporarily unavailable/i,
        /server error/i,
        /internal error/i
      ]],
      [GAS_ERROR_TYPES.NETWORK_ERROR, [
        /network error/i,
        /connection failed/i,
        /timeout error/i,
        /no internet connection/i
      ]],
      [GAS_ERROR_TYPES.AUTHENTICATION_FAILED, [
        /authentication failed/i,
        /invalid credentials/i,
        /session expired/i,
        /login required/i
      ]],
      [GAS_ERROR_TYPES.RATE_LIMITED, [
        /rate limited/i,
        /too many requests/i,
        /request throttled/i
      ]]
    ]);
  }

  /**
   * 回復戦略の初期化
   */
  initializeRecoveryStrategies() {
    return new Map([
      [GAS_ERROR_TYPES.EXECUTION_TIMEOUT, {
        retryable: false,
        strategy: 'fragmentExecution',
        description: '実行を小さな単位に分割',
        handler: this.handleExecutionTimeout.bind(this)
      }],
      [GAS_ERROR_TYPES.PERMISSION_DENIED, {
        retryable: false,
        strategy: 'degradedMode',
        description: '機能制限モードで継続',
        handler: this.handlePermissionDenied.bind(this)
      }],
      [GAS_ERROR_TYPES.QUOTA_EXCEEDED, {
        retryable: true,
        maxRetries: 2,
        retryDelay: 60000, // 1分待機
        strategy: 'exponentialBackoff',
        description: 'クォータ回復待ち',
        handler: this.handleQuotaExceeded.bind(this)
      }],
      [GAS_ERROR_TYPES.SCRIPT_NOT_FOUND, {
        retryable: false,
        strategy: 'fallbackFunction',
        description: '代替機能の使用',
        handler: this.handleScriptNotFound.bind(this)
      }],
      [GAS_ERROR_TYPES.LOCK_FAILED, {
        retryable: true,
        maxRetries: 3,
        retryDelay: 1000,
        strategy: 'randomizedBackoff',
        description: 'ランダム待機後リトライ',
        handler: this.handleLockFailed.bind(this)
      }],
      [GAS_ERROR_TYPES.SERVICE_UNAVAILABLE, {
        retryable: true,
        maxRetries: 3,
        retryDelay: 5000,
        strategy: 'exponentialBackoff',
        description: 'サービス復旧待ち',
        handler: this.handleServiceUnavailable.bind(this)
      }],
      [GAS_ERROR_TYPES.NETWORK_ERROR, {
        retryable: true,
        maxRetries: 2,
        retryDelay: 3000,
        strategy: 'networkReconnect',
        description: 'ネットワーク再接続',
        handler: this.handleNetworkError.bind(this)
      }],
      [GAS_ERROR_TYPES.AUTHENTICATION_FAILED, {
        retryable: false,
        strategy: 'reauthenticate',
        description: '再認証が必要',
        handler: this.handleAuthenticationFailed.bind(this)
      }],
      [GAS_ERROR_TYPES.RATE_LIMITED, {
        retryable: true,
        maxRetries: 2,
        retryDelay: 15000,
        strategy: 'rateLimitBackoff',
        description: 'レート制限回復待ち',
        handler: this.handleRateLimited.bind(this)
      }]
    ]);
  }

  /**
   * グローバルエラーハンドリングの設定
   */
  setupGlobalErrorHandling() {
    // 未処理のPromise拒否をキャッチ
    window.addEventListener('unhandledrejection', (event) => {
      console.error('🚨 Unhandled promise rejection:', event.reason);
      this.handleError(event.reason, 'unhandledRejection');
    });
    
    // グローバルエラーハンドラー
    window.addEventListener('error', (event) => {
      console.error('🚨 Global error:', event.error);
      this.handleError(event.error, 'globalError');
    });
  }

  /**
   * メインエラーハンドリング
   */
  async handleError(error, context = 'unknown') {
    const errorInfo = this.analyzeError(error);
    
    console.error('🔍 Error analysis:', {
      type: errorInfo.type,
      message: errorInfo.message,
      context,
      retryable: errorInfo.retryable
    });
    
    // 統計更新
    this.updateErrorStats(errorInfo);
    
    // サーキットブレーカーチェック
    if (this.isCircuitOpen()) {
      console.warn('⚡ Circuit breaker is open, failing fast');
      throw new Error('Circuit breaker is open - too many recent failures');
    }
    
    // 回復試行
    if (errorInfo.retryable && !this.isCircuitOpen()) {
      return await this.attemptRecovery(errorInfo, context);
    } else {
      return await this.handleNonRetryableError(errorInfo, context);
    }
  }

  /**
   * エラー分析
   */
  analyzeError(error) {
    const errorMessage = error?.message || error?.toString() || 'Unknown error';
    let errorType = 'unknown';
    let retryable = false;
    
    // エラーパターンマッチング
    for (const [type, patterns] of this.errorPatterns.entries()) {
      if (patterns.some(pattern => pattern.test(errorMessage))) {
        errorType = type;
        break;
      }
    }
    
    // 回復戦略の取得
    const strategy = this.recoveryStrategies.get(errorType);
    if (strategy) {
      retryable = strategy.retryable;
    }
    
    // GAS実行時間チェック
    const executionTime = Date.now() - this.executionMonitor.startTime;
    if (executionTime > this.executionMonitor.warningThreshold) {
      console.warn(`⚠️ Long execution time detected: ${executionTime}ms`);
      
      if (executionTime > this.executionMonitor.criticalThreshold) {
        errorType = GAS_ERROR_TYPES.EXECUTION_TIMEOUT;
        retryable = false;
      }
    }
    
    return {
      type: errorType,
      message: errorMessage,
      originalError: error,
      retryable,
      strategy: strategy?.strategy || 'none',
      executionTime
    };
  }

  /**
   * 回復試行
   */
  async attemptRecovery(errorInfo, context) {
    const strategy = this.recoveryStrategies.get(errorInfo.type);
    if (!strategy) {
      throw errorInfo.originalError;
    }
    
    this.errorStats.recoveryAttempts++;
    
    console.log(`🔄 Attempting recovery with strategy: ${strategy.strategy}`);
    
    try {
      const result = await strategy.handler(errorInfo, context);
      
      this.errorStats.successfulRecoveries++;
      console.log('✅ Recovery successful');
      
      return result;
      
    } catch (recoveryError) {
      console.error('❌ Recovery failed:', recoveryError);
      
      // サーキットブレーカーのトリガー
      this.triggerCircuitBreaker();
      
      throw recoveryError;
    }
  }

  /**
   * リトライ不可能エラーの処理
   */
  async handleNonRetryableError(errorInfo, context) {
    console.warn(`⚠️ Non-retryable error detected: ${errorInfo.type}`);
    
    // エラー状態を Core に記録
    this.coreState.setError(errorInfo.originalError, `${context} - ${errorInfo.type}`);
    
    // 代替処理の試行
    const fallbackResult = await this.attemptFallback(errorInfo, context);
    
    if (fallbackResult) {
      console.log('✅ Fallback succeeded');
      return fallbackResult;
    }
    
    // 最終的な失敗
    throw errorInfo.originalError;
  }

  /**
   * 代替処理の試行
   */
  async attemptFallback(errorInfo, context) {
    const fallbackStrategies = {
      [GAS_ERROR_TYPES.PERMISSION_DENIED]: () => this.enterDegradedMode(),
      [GAS_ERROR_TYPES.SCRIPT_NOT_FOUND]: () => this.useFallbackFunction(context),
      [GAS_ERROR_TYPES.AUTHENTICATION_FAILED]: () => this.promptReauthentication()
    };
    
    const fallbackFn = fallbackStrategies[errorInfo.type];
    if (fallbackFn) {
      try {
        return await fallbackFn();
      } catch (fallbackError) {
        console.error('❌ Fallback failed:', fallbackError);
      }
    }
    
    return null;
  }

  /**
   * 個別エラーハンドラー
   */
  async handleExecutionTimeout(errorInfo, context) {
    console.log('⏱️ Handling execution timeout');
    
    // 実行を小さな単位に分割
    if (context.includes('rendering') || context.includes('batch')) {
      return await this.fragmentExecution(context);
    }
    
    // タイムアウト情報をUIに表示
    this.showTimeoutWarning();
    
    // 軽量版の処理に切り替え
    return await this.switchToLightweightMode();
  }

  async handlePermissionDenied(errorInfo, context) {
    console.log('🔒 Handling permission denied');
    
    // 機能制限モードに移行
    return await this.enterDegradedMode();
  }

  async handleQuotaExceeded(errorInfo, context) {
    console.log('📊 Handling quota exceeded');
    
    // API呼び出し頻度を削減
    this.reduceAPIFrequency();
    
    // 指定時間待機後にリトライ
    const strategy = this.recoveryStrategies.get(errorInfo.type);
    await this.delay(strategy.retryDelay);
    
    return { recovered: true, strategy: 'quotaBackoff' };
  }

  async handleLockFailed(errorInfo, context) {
    console.log('🔐 Handling lock failed');
    
    // ランダムな待機時間でリトライ
    const randomDelay = 1000 + Math.random() * 3000; // 1-4秒
    await this.delay(randomDelay);
    
    return { recovered: true, strategy: 'randomBackoff' };
  }

  async handleServiceUnavailable(errorInfo, context) {
    console.log('🚫 Handling service unavailable');
    
    // サービス状態をチェック
    const isServiceUp = await this.checkServiceHealth();
    
    if (isServiceUp) {
      return { recovered: true, strategy: 'serviceReconnect' };
    }
    
    // サービス復旧待ち
    await this.delay(5000);
    throw new Error('Service still unavailable');
  }

  async handleNetworkError(errorInfo, context) {
    console.log('🌐 Handling network error');
    
    // ネットワーク接続をチェック
    if (navigator.onLine === false) {
      return await this.handleOfflineMode();
    }
    
    // 短い待機後にリトライ
    await this.delay(3000);
    return { recovered: true, strategy: 'networkRetry' };
  }

  async handleAuthenticationFailed(errorInfo, context) {
    console.log('🔑 Handling authentication failed');
    
    // 再認証プロンプトを表示
    return await this.promptReauthentication();
  }

  async handleRateLimited(errorInfo, context) {
    console.log('🚦 Handling rate limited');
    
    // レート制限情報を表示
    this.showRateLimitWarning();
    
    // 長めの待機
    await this.delay(15000);
    
    return { recovered: true, strategy: 'rateLimitBackoff' };
  }

  /**
   * ユーティリティメソッド
   */
  async delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  isCircuitOpen() {
    if (!this.recoveryControl.isCircuitOpen) return false;
    
    // リセット時間をチェック
    const elapsed = Date.now() - this.recoveryControl.circuitOpenTime;
    if (elapsed > this.recoveryControl.circuitBreakerResetTime) {
      this.resetCircuitBreaker();
      return false;
    }
    
    return true;
  }

  triggerCircuitBreaker() {
    this.recoveryControl.isCircuitOpen = true;
    this.recoveryControl.circuitOpenTime = Date.now();
    
    console.warn('⚡ Circuit breaker triggered - failing fast for next requests');
  }

  resetCircuitBreaker() {
    this.recoveryControl.isCircuitOpen = false;
    this.recoveryControl.circuitOpenTime = 0;
    
    console.log('✅ Circuit breaker reset');
  }

  updateErrorStats(errorInfo) {
    this.errorStats.totalErrors++;
    this.errorStats.lastErrorTime = Date.now();
    
    const count = this.errorStats.errorsByType.get(errorInfo.type) || 0;
    this.errorStats.errorsByType.set(errorInfo.type, count + 1);
    
    // サーキットブレーカーのチェック
    if (count + 1 >= this.recoveryControl.circuitBreakerThreshold) {
      this.triggerCircuitBreaker();
    }
  }

  // 回復処理の具体的実装
  async fragmentExecution(context) {
    console.log('🔨 Fragmenting execution for timeout recovery');
    return { recovered: true, strategy: 'fragmented', reducedLoad: true };
  }

  async switchToLightweightMode() {
    console.log('⚡ Switching to lightweight mode');
    this.coreState.updateState({ isLowPerformanceMode: true });
    return { recovered: true, strategy: 'lightweight' };
  }

  async enterDegradedMode() {
    console.log('🔧 Entering degraded mode');
    this.coreState.updateState({ isDegradedMode: true });
    return { recovered: true, strategy: 'degraded', limitedFunctionality: true };
  }

  async useFallbackFunction(context) {
    console.log('🔄 Using fallback function');
    return { recovered: true, strategy: 'fallback', usedAlternative: true };
  }

  async promptReauthentication() {
    console.log('🔑 Prompting reauthentication');
    // 実際の実装では認証ダイアログを表示
    return { recovered: false, strategy: 'reauthentication', userActionRequired: true };
  }

  reduceAPIFrequency() {
    console.log('📉 Reducing API call frequency');
    // ポーリング間隔を延長
    if (window.studyQuestApp?.pollingManager) {
      window.studyQuestApp.pollingManager.extendInterval();
    }
  }

  async checkServiceHealth() {
    try {
      // 軽量なヘルスチェック
      const response = await fetch('/health', { method: 'HEAD' });
      return response.ok;
    } catch {
      return false;
    }
  }

  async handleOfflineMode() {
    console.log('📱 Entering offline mode');
    this.coreState.updateState({ isOffline: true });
    return { recovered: true, strategy: 'offline', cachedDataOnly: true };
  }

  showTimeoutWarning() {
    console.warn('⏱️ Execution timeout warning displayed');
    // 実際の実装ではUIに警告を表示
  }

  showRateLimitWarning() {
    console.warn('🚦 Rate limit warning displayed');
    // 実際の実装ではUIに警告を表示
  }

  /**
   * 統計情報の取得
   */
  getErrorStats() {
    return {
      ...this.errorStats,
      errorsByType: Object.fromEntries(this.errorStats.errorsByType),
      circuitBreakerStatus: {
        isOpen: this.recoveryControl.isCircuitOpen,
        openSince: this.recoveryControl.circuitOpenTime,
        resetIn: this.recoveryControl.isCircuitOpen ? 
          Math.max(0, this.recoveryControl.circuitBreakerResetTime - (Date.now() - this.recoveryControl.circuitOpenTime)) : 0
      },
      recoveryRate: this.errorStats.recoveryAttempts > 0 ? 
        (this.errorStats.successfulRecoveries / this.errorStats.recoveryAttempts * 100).toFixed(1) + '%' : '0%'
    };
  }

  /**
   * リソースのクリーンアップ
   */
  destroy() {
    // 統計の最終出力
    console.log('📊 Final error statistics:', this.getErrorStats());
    
    // リソースクリア
    this.errorPatterns.clear();
    this.recoveryStrategies.clear();
    this.errorStats.errorsByType.clear();
    
    console.log('🧹 GAS Error Handler destroyed');
  }
}

// グローバルエクスポート
window.GASErrorHandler = GASErrorHandler;
window.GAS_ERROR_TYPES = GAS_ERROR_TYPES;

