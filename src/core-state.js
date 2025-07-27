/**
 * StudyQuest Core State Management
 * GAS最適化済み - 状態管理とアプリケーション初期化
 */

// アプリケーション定数
const STUDY_QUEST_CONSTANTS = {
  REACTION_RATE_LIMIT_MS: 500,
  HIGHLIGHT_RATE_LIMIT_MS: 500,
  CACHE_TTL_MS: 1000,
  RETRY_DELAY_MS: 2000,
  POLLING_INTERVAL_MS: 15000, // GAS最適化: 5秒→15秒
  INIT_TIMEOUT_MS: 30000,
  
  // GAS固有の制約
  GAS_EXECUTION_TIMEOUT_MS: 360000, // 6分
  GAS_MAX_MEMORY_MB: 200,
  GAS_BATCH_SIZE_LIMIT: 10
};

/**
 * Core State Manager
 * アプリケーションの中心的な状態管理を担当
 */
class StudyQuestCoreState {
  constructor() {
    this.initializeState();
    this.initializeGASConstraints();
  }

  /**
   * 基本状態の初期化
   */
  initializeState() {
    this.state = {
      currentAnswers: [],
      isLoading: false,
      hasNewContent: false,
      newContentCount: 0,
      lastSeenCount: 0,
      userId: '',
      sheetName: '',
      currentActiveSheet: '',
      showAdminFeatures: false,
      isPollingActive: false,
      error: null,
      initialized: false
    };

    this.cache = new Map();
    this.executionStartTime = Date.now();
    this.performanceMetrics = {
      initTime: 0,
      renderTime: 0,
      apiCallCount: 0,
      memoryUsage: 0
    };
  }

  /**
   * GAS制約を考慮した初期化
   */
  initializeGASConstraints() {
    // GAS実行時間監視
    this.gasExecutionMonitor = {
      startTime: Date.now(),
      warningThreshold: STUDY_QUEST_CONSTANTS.GAS_EXECUTION_TIMEOUT_MS * 0.8, // 80%で警告
      criticalThreshold: STUDY_QUEST_CONSTANTS.GAS_EXECUTION_TIMEOUT_MS * 0.95 // 95%で強制停止
    };

    // メモリ使用量監視（可能な場合）
    if (performance.memory) {
      this.memoryMonitor = {
        initial: performance.memory.usedJSHeapSize,
        threshold: STUDY_QUEST_CONSTANTS.GAS_MAX_MEMORY_MB * 1024 * 1024 * 0.8
      };
    }
  }

  /**
   * ユーザーIDの検証とフォールバック
   */
  validateUserId(userId) {
    if (!userId || typeof userId !== 'string' || userId.trim() === '') {
      console.warn('⚠️ StudyQuestApp: Invalid or empty userId provided:', userId);
      
      try {
        const urlParams = new URLSearchParams(window.location.search);
        const fallbackUserId = urlParams.get('userId');
        
        if (fallbackUserId) {
          console.log('✅ StudyQuestApp: Using fallback userId from URL parameters:', fallbackUserId);
          return fallbackUserId;
        }
      } catch (error) {
        console.warn('⚠️ StudyQuestApp: Failed to get userId from URL parameters:', error.message);
      }
      
      console.error('❌ StudyQuestApp: No valid userId available. Data loading may fail.');
      return '';
    }
    
    if (userId.length === 36 && userId.includes('-')) {
      console.log('✅ StudyQuestApp: Valid userId format detected:', userId.substring(0, 8) + '...');
      return userId;
    }
    
    console.warn('⚠️ StudyQuestApp: Unusual userId format, but proceeding:', userId.substring(0, 8) + '...');
    return userId;
  }

  /**
   * 状態の更新（不変性を保持）
   */
  updateState(updates) {
    const oldState = { ...this.state };
    this.state = { ...this.state, ...updates };
    
    // 状態変更の監視（デバッグ用）
    if (this.isDevelopment()) {
      console.log('State updated:', {
        changes: this.getStateChanges(oldState, this.state),
        newState: this.state
      });
    }
    
    return this.state;
  }

  /**
   * 状態変更の差分を取得
   */
  getStateChanges(oldState, newState) {
    const changes = {};
    for (const key in newState) {
      if (newState[key] !== oldState[key]) {
        changes[key] = { from: oldState[key], to: newState[key] };
      }
    }
    return changes;
  }

  /**
   * GAS実行時間のチェック
   */
  checkGASExecutionTime() {
    const elapsed = Date.now() - this.gasExecutionMonitor.startTime;
    
    if (elapsed > this.gasExecutionMonitor.criticalThreshold) {
      throw new Error('GAS実行時間制限に到達しました。処理を中断します。');
    }
    
    if (elapsed > this.gasExecutionMonitor.warningThreshold) {
      console.warn('⚠️ GAS実行時間が警告レベルに到達しました:', elapsed + 'ms');
      return { warning: true, elapsed };
    }
    
    return { warning: false, elapsed };
  }

  /**
   * メモリ使用量のチェック
   */
  checkMemoryUsage() {
    if (!performance.memory || !this.memoryMonitor) return { warning: false };
    
    const current = performance.memory.usedJSHeapSize;
    const increase = current - this.memoryMonitor.initial;
    
    if (increase > this.memoryMonitor.threshold) {
      console.warn('⚠️ メモリ使用量が警告レベルに到達しました:', 
        Math.round(increase / 1024 / 1024) + 'MB');
      return { warning: true, usage: increase };
    }
    
    return { warning: false, usage: increase };
  }

  /**
   * 開発環境かどうかの判定
   */
  isDevelopment() {
    return window.location.hostname === 'localhost' || 
           window.location.hostname.includes('script.google.com');
  }

  /**
   * エラー状態の設定
   */
  setError(error, context = '') {
    const errorInfo = {
      message: error.message || error,
      context,
      timestamp: new Date().toISOString(),
      stack: error.stack
    };
    
    this.updateState({ error: errorInfo });
    console.error('❌ StudyQuest Error:', errorInfo);
    
    return errorInfo;
  }

  /**
   * エラー状態のクリア
   */
  clearError() {
    this.updateState({ error: null });
  }

  /**
   * パフォーマンスメトリクスの更新
   */
  updatePerformanceMetrics(metrics) {
    this.performanceMetrics = { ...this.performanceMetrics, ...metrics };
    
    // GAS環境での監視
    if (this.isDevelopment()) {
      console.log('Performance metrics updated:', this.performanceMetrics);
    }
  }

  /**
   * 状態のリセット（エラー回復用）
   */
  resetState() {
    const preservedState = {
      userId: this.state.userId,
      sheetName: this.state.sheetName
    };
    
    this.initializeState();
    this.updateState(preservedState);
    
    console.log('🔄 State reset completed');
  }

  /**
   * 状態のシリアライズ（デバッグ用）
   */
  serializeState() {
    return {
      state: this.state,
      performanceMetrics: this.performanceMetrics,
      gasExecutionTime: Date.now() - this.gasExecutionMonitor.startTime,
      memoryUsage: performance.memory ? 
        Math.round(performance.memory.usedJSHeapSize / 1024 / 1024) + 'MB' : 'unknown'
    };
  }

  /**
   * リソースのクリーンアップ
   */
  destroy() {
    this.cache.clear();
    this.state = null;
    this.performanceMetrics = null;
    this.gasExecutionMonitor = null;
    this.memoryMonitor = null;
    
    console.log('🧹 Core state destroyed');
  }
}

// グローバルな状態管理インスタンス
window.StudyQuestCoreState = StudyQuestCoreState;
window.STUDY_QUEST_CONSTANTS = STUDY_QUEST_CONSTANTS;

