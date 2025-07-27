/**
 * StudyQuest API Optimizer
 * GAS最適化済み - API呼び出しパターンの最適化、ポーリング、バッチ処理
 */

/**
 * 最適化されたポーリングマネージャー
 */
class OptimizedPollingManager {
  constructor(coreState, dataManager) {
    this.coreState = coreState;
    this.dataManager = dataManager;
    
    // GAS最適化された設定
    this.pollingSettings = {
      enabled: false,
      baseInterval: 15000, // 15秒（GAS最適化: 5秒→15秒）
      maxInterval: 60000,  // 最大1分
      backoffMultiplier: 1.5,
      currentInterval: 15000,
      consecutiveFailures: 0,
      maxConsecutiveFailures: 3
    };
    
    // 適応的ポーリング
    this.adaptiveSettings = {
      activityThreshold: 5, // 5分間の非アクティブでスロー化
      lastActivity: Date.now(),
      isSlowMode: false,
      fastModeInterval: 15000, // 高速モード
      slowModeInterval: 45000  // 低速モード
    };
    
    this.pollingTimer = null;
    this.isPolling = false;
    this.pollCount = 0;
    
    // パフォーマンス監視
    this.performanceMetrics = {
      totalPolls: 0,
      successfulPolls: 0,
      failedPolls: 0,
      avgResponseTime: 0,
      responseTimes: []
    };
    
    this.setupActivityTracking();
  }

  /**
   * アクティビティトラッキングの設定
   */
  setupActivityTracking() {
    // ユーザーアクティビティを監視
    const activityEvents = ['click', 'scroll', 'keydown', 'mousemove'];
    
    const updateActivity = this.debounce(() => {
      this.adaptiveSettings.lastActivity = Date.now();
      
      // スローモードから復帰
      if (this.adaptiveSettings.isSlowMode) {
        this.adaptiveSettings.isSlowMode = false;
        this.adjustPollingInterval();
        console.log('📈 Polling: Switched to fast mode due to user activity');
      }
    }, 1000);
    
    activityEvents.forEach(event => {
      document.addEventListener(event, updateActivity, { passive: true });
    });
    
    // 定期的なアクティビティチェック
    setInterval(() => {
      this.checkActivityLevel();
    }, 60000); // 1分ごと
  }

  /**
   * アクティビティレベルのチェック
   */
  checkActivityLevel() {
    const inactiveTime = Date.now() - this.adaptiveSettings.lastActivity;
    const thresholdMs = this.adaptiveSettings.activityThreshold * 60 * 1000;
    
    if (inactiveTime > thresholdMs && !this.adaptiveSettings.isSlowMode) {
      this.adaptiveSettings.isSlowMode = true;
      this.adjustPollingInterval();
      console.log('📉 Polling: Switched to slow mode due to inactivity');
    }
  }

  /**
   * ポーリング間隔の動的調整
   */
  adjustPollingInterval() {
    const targetInterval = this.adaptiveSettings.isSlowMode 
      ? this.adaptiveSettings.slowModeInterval 
      : this.adaptiveSettings.fastModeInterval;
    
    if (this.pollingSettings.currentInterval !== targetInterval) {
      this.pollingSettings.currentInterval = targetInterval;
      
      // 現在ポーリング中の場合は再起動
      if (this.isPolling) {
        this.restartPolling();
      }
    }
  }

  /**
   * ポーリング開始
   */
  startPolling() {
    if (this.isPolling) {
      console.log('⚠️ Polling already active');
      return;
    }
    
    this.pollingSettings.enabled = true;
    this.isPolling = true;
    this.consecutiveFailures = 0;
    
    console.log('🚀 Starting optimized polling:', {
      interval: this.pollingSettings.currentInterval,
      adaptiveMode: this.adaptiveSettings.isSlowMode ? 'slow' : 'fast'
    });
    
    this.scheduleNextPoll();
  }

  /**
   * ポーリング停止
   */
  stopPolling() {
    if (!this.isPolling) return;
    
    this.pollingSettings.enabled = false;
    this.isPolling = false;
    
    if (this.pollingTimer) {
      clearTimeout(this.pollingTimer);
      this.pollingTimer = null;
    }
    
    console.log('⏹️ Polling stopped');
  }

  /**
   * ポーリング再起動
   */
  restartPolling() {
    this.stopPolling();
    setTimeout(() => this.startPolling(), 100);
  }

  /**
   * 次のポーリングをスケジュール
   */
  scheduleNextPoll() {
    if (!this.pollingSettings.enabled) return;
    
    this.pollingTimer = setTimeout(async () => {
      await this.executePoll();
      if (this.pollingSettings.enabled) {
        this.scheduleNextPoll();
      }
    }, this.pollingSettings.currentInterval);
  }

  /**
   * ポーリング実行
   */
  async executePoll() {
    const startTime = performance.now();
    this.performanceMetrics.totalPolls++;
    
    try {
      // GAS実行時間チェック
      this.coreState.checkGASExecutionTime();
      
      console.log(`🔄 Polling check #${this.pollCount + 1}`);
      
      const result = await this.performPollingCheck();
      
      if (result && result.hasNewContent) {
        console.log('📢 New content detected:', result);
        this.handleNewContentDetected(result);
      }
      
      // 成功時の処理
      this.handlePollingSuccess();
      const responseTime = performance.now() - startTime;
      this.updatePerformanceMetrics(responseTime, true);
      
      this.pollCount++;
      
    } catch (error) {
      console.error('❌ Polling error:', error);
      this.handlePollingError(error);
      
      const responseTime = performance.now() - startTime;
      this.updatePerformanceMetrics(responseTime, false);
    }
  }

  /**
   * ポーリングチェックの実行
   */
  async performPollingCheck() {
    const currentAnswersCount = this.coreState.state.currentAnswers?.length || 0;
    const classFilter = this.getCurrentClassFilter();
    const sortOrder = this.getCurrentSortOrder();
    const adminMode = this.coreState.state.showAdminFeatures;
    
    // 軽量なカウントチェックから開始
    const countData = await this.dataManager.getDataCount(classFilter, sortOrder, adminMode);
    
    if (!countData || typeof countData.count !== 'number') {
      console.warn('⚠️ Invalid count data received:', countData);
      return null;
    }
    
    const serverCount = countData.count;
    const newContentCount = Math.max(0, serverCount - currentAnswersCount);
    
    console.log('📊 Polling check result:', {
      server: serverCount,
      client: currentAnswersCount,
      difference: newContentCount
    });
    
    return {
      hasNewContent: newContentCount > 0,
      newContentCount,
      serverCount,
      clientCount: currentAnswersCount
    };
  }

  /**
   * 新コンテンツ検出時の処理
   */
  handleNewContentDetected(result) {
    // 状態を更新
    this.coreState.updateState({
      hasNewContent: true,
      newContentCount: result.newContentCount
    });
    
    // UI通知をトリガー
    if (window.studyQuestApp && typeof window.studyQuestApp.showNewContentBanner === 'function') {
      window.studyQuestApp.showNewContentBanner([]);
    }
    
    // 統計更新
    this.coreState.updatePerformanceMetrics({
      lastPollingResult: result
    });
  }

  /**
   * ポーリング成功時の処理
   */
  handlePollingSuccess() {
    this.pollingSettings.consecutiveFailures = 0;
    this.performanceMetrics.successfulPolls++;
    
    // 間隔をリセット（バックオフから回復）
    if (this.pollingSettings.currentInterval > this.getBaseInterval()) {
      this.pollingSettings.currentInterval = this.getBaseInterval();
      console.log('✅ Polling interval reset to base value');
    }
  }

  /**
   * ポーリングエラー時の処理
   */
  handlePollingError(error) {
    this.pollingSettings.consecutiveFailures++;
    this.performanceMetrics.failedPolls++;
    
    // 指数バックオフ
    if (this.pollingSettings.consecutiveFailures <= this.pollingSettings.maxConsecutiveFailures) {
      this.applyExponentialBackoff();
    } else {
      console.error('❌ Too many consecutive polling failures, stopping polling');
      this.stopPolling();
      
      // エラー状態を記録
      this.coreState.setError(error, 'Polling Manager');
    }
  }

  /**
   * 指数バックオフの適用
   */
  applyExponentialBackoff() {
    const newInterval = Math.min(
      this.pollingSettings.currentInterval * this.pollingSettings.backoffMultiplier,
      this.pollingSettings.maxInterval
    );
    
    this.pollingSettings.currentInterval = newInterval;
    
    console.log(`⏰ Applied exponential backoff: ${newInterval}ms (failures: ${this.pollingSettings.consecutiveFailures})`);
  }

  /**
   * パフォーマンスメトリクスの更新
   */
  updatePerformanceMetrics(responseTime, success) {
    this.performanceMetrics.responseTimes.push(responseTime);
    
    // 最新20回の平均を計算
    if (this.performanceMetrics.responseTimes.length > 20) {
      this.performanceMetrics.responseTimes.shift();
    }
    
    this.performanceMetrics.avgResponseTime = 
      this.performanceMetrics.responseTimes.reduce((a, b) => a + b, 0) / 
      this.performanceMetrics.responseTimes.length;
    
    // Core状態に報告
    this.coreState.updatePerformanceMetrics({
      pollingMetrics: {
        avgResponseTime: this.performanceMetrics.avgResponseTime,
        successRate: (this.performanceMetrics.successfulPolls / this.performanceMetrics.totalPolls) * 100,
        totalPolls: this.performanceMetrics.totalPolls
      }
    });
  }

  /**
   * ユーティリティメソッド
   */
  getBaseInterval() {
    return this.adaptiveSettings.isSlowMode 
      ? this.adaptiveSettings.slowModeInterval 
      : this.adaptiveSettings.fastModeInterval;
  }

  getCurrentClassFilter() {
    const element = document.getElementById('classFilter');
    return element ? element.value : 'すべて';
  }

  getCurrentSortOrder() {
    const element = document.getElementById('sortOrder');
    return element ? element.value : 'newest';
  }

  debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
      const later = () => {
        clearTimeout(timeout);
        func(...args);
      };
      clearTimeout(timeout);
      timeout = setTimeout(later, wait);
    };
  }

  /**
   * 統計情報の取得
   */
  getStats() {
    return {
      isPolling: this.isPolling,
      currentInterval: this.pollingSettings.currentInterval,
      consecutiveFailures: this.pollingSettings.consecutiveFailures,
      adaptiveMode: this.adaptiveSettings.isSlowMode ? 'slow' : 'fast',
      performance: this.performanceMetrics,
      pollCount: this.pollCount
    };
  }

  /**
   * リソースのクリーンアップ
   */
  destroy() {
    this.stopPolling();
    console.log('🧹 Polling manager destroyed');
  }
}

/**
 * バッチ処理最適化マネージャー
 */
class BatchProcessingManager {
  constructor(coreState, dataManager) {
    this.coreState = coreState;
    this.dataManager = dataManager;
    
    // バッチ設定
    this.batchSettings = {
      maxBatchSize: STUDY_QUEST_CONSTANTS.GAS_BATCH_SIZE_LIMIT,
      batchTimeout: 2000, // 2秒でタイムアウト
      processingDelay: 500, // 処理間隔
      maxRetries: 3
    };
    
    // キューシステム
    this.reactionQueue = new Map();
    this.highlightQueue = new Map();
    this.processingTimer = null;
    this.isProcessing = false;
    
    // パフォーマンス監視
    this.batchMetrics = {
      totalBatches: 0,
      successfulBatches: 0,
      failedBatches: 0,
      averageBatchSize: 0,
      totalItemsProcessed: 0
    };
  }

  /**
   * リアクションをバッチに追加
   */
  addReactionToBatch(rowIndex, reaction, options = {}) {
    const key = `${rowIndex}-${reaction}`;
    
    if (this.reactionQueue.has(key)) {
      // 既存エントリのタイムスタンプを更新
      const existing = this.reactionQueue.get(key);
      existing.timestamp = Date.now();
      existing.retryCount = 0; // リトライカウントリセット
    } else {
      // 新しいエントリを追加
      this.reactionQueue.set(key, {
        rowIndex: parseInt(rowIndex),
        reaction,
        timestamp: Date.now(),
        retryCount: 0,
        ...options
      });
    }
    
    this.scheduleProcessing();
    
    console.log(`📦 Added reaction to batch queue: ${key} (queue size: ${this.reactionQueue.size})`);
  }

  /**
   * ハイライトをバッチに追加
   */
  addHighlightToBatch(rowIndex, options = {}) {
    const key = `highlight-${rowIndex}`;
    
    this.highlightQueue.set(key, {
      rowIndex: parseInt(rowIndex),
      timestamp: Date.now(),
      retryCount: 0,
      ...options
    });
    
    this.scheduleProcessing();
    
    console.log(`📦 Added highlight to batch queue: ${key} (queue size: ${this.highlightQueue.size})`);
  }

  /**
   * バッチ処理のスケジューリング
   */
  scheduleProcessing() {
    if (this.processingTimer) {
      clearTimeout(this.processingTimer);
    }
    
    // 即座処理が必要な条件
    const totalQueueSize = this.reactionQueue.size + this.highlightQueue.size;
    const shouldProcessImmediately = 
      totalQueueSize >= this.batchSettings.maxBatchSize || 
      this.hasUrgentItems();
    
    const delay = shouldProcessImmediately ? 0 : this.batchSettings.processingDelay;
    
    this.processingTimer = setTimeout(() => {
      this.processBatchQueues();
    }, delay);
  }

  /**
   * 緊急処理が必要なアイテムの確認
   */
  hasUrgentItems() {
    const urgentThreshold = Date.now() - (this.batchSettings.batchTimeout * 2);
    
    for (const item of this.reactionQueue.values()) {
      if (item.timestamp < urgentThreshold) return true;
    }
    
    for (const item of this.highlightQueue.values()) {
      if (item.timestamp < urgentThreshold) return true;
    }
    
    return false;
  }

  /**
   * バッチキューの処理
   */
  async processBatchQueues() {
    if (this.isProcessing) {
      console.log('⚠️ Batch processing already in progress');
      return;
    }
    
    if (this.reactionQueue.size === 0 && this.highlightQueue.size === 0) {
      return;
    }
    
    this.isProcessing = true;
    
    try {
      console.log(`🔄 Processing batch queues: ${this.reactionQueue.size} reactions, ${this.highlightQueue.size} highlights`);
      
      // リアクションバッチの処理
      if (this.reactionQueue.size > 0) {
        await this.processReactionBatch();
      }
      
      // ハイライトバッチの処理
      if (this.highlightQueue.size > 0) {
        await this.processHighlightBatch();
      }
      
    } catch (error) {
      console.error('❌ Batch processing error:', error);
      this.coreState.setError(error, 'Batch Processing');
    } finally {
      this.isProcessing = false;
    }
  }

  /**
   * リアクションバッチの処理
   */
  async processReactionBatch() {
    const batchOperations = Array.from(this.reactionQueue.values())
      .slice(0, this.batchSettings.maxBatchSize);
    
    if (batchOperations.length === 0) return;
    
    try {
      const result = await this.dataManager.addReactionBatch(batchOperations);
      
      if (result && result.success) {
        console.log(`✅ Reaction batch processed successfully: ${batchOperations.length} items`);
        
        // 成功したアイテムをキューから削除
        batchOperations.forEach(op => {
          const key = `${op.rowIndex}-${op.reaction}`;
          this.reactionQueue.delete(key);
        });
        
        this.updateBatchMetrics(batchOperations.length, true);
        
        // UIの更新通知
        this.notifyBatchCompletion('reactions', batchOperations, result);
        
      } else {
        throw new Error('Batch processing failed on server');
      }
      
    } catch (error) {
      console.warn('⚠️ Reaction batch failed, falling back to individual processing:', error);
      await this.fallbackToIndividualProcessing(batchOperations, 'reaction');
    }
  }

  /**
   * ハイライトバッチの処理
   */
  async processHighlightBatch() {
    const batchOperations = Array.from(this.highlightQueue.values())
      .slice(0, this.batchSettings.maxBatchSize);
    
    if (batchOperations.length === 0) return;
    
    try {
      // ハイライト用のバッチAPI呼び出し（実装が必要）
      const results = await Promise.all(
        batchOperations.map(op => 
          this.dataManager.runGas('toggleHighlight', op.rowIndex, this.coreState.state.sheetName)
        )
      );
      
      console.log(`✅ Highlight batch processed: ${batchOperations.length} items`);
      
      // 成功したアイテムをキューから削除
      batchOperations.forEach(op => {
        const key = `highlight-${op.rowIndex}`;
        this.highlightQueue.delete(key);
      });
      
      this.updateBatchMetrics(batchOperations.length, true);
      
    } catch (error) {
      console.warn('⚠️ Highlight batch failed:', error);
      await this.fallbackToIndividualProcessing(batchOperations, 'highlight');
    }
  }

  /**
   * 個別処理へのフォールバック
   */
  async fallbackToIndividualProcessing(operations, type) {
    console.log(`🔄 Falling back to individual processing for ${operations.length} ${type} operations`);
    
    for (const op of operations) {
      try {
        if (type === 'reaction') {
          await this.dataManager.addReaction(op.rowIndex, op.reaction, this.coreState.state.sheetName);
          const key = `${op.rowIndex}-${op.reaction}`;
          this.reactionQueue.delete(key);
        } else if (type === 'highlight') {
          await this.dataManager.runGas('toggleHighlight', op.rowIndex, this.coreState.state.sheetName);
          const key = `highlight-${op.rowIndex}`;
          this.highlightQueue.delete(key);
        }
        
        // 処理間隔を設ける（GAS負荷軽減）
        await new Promise(resolve => setTimeout(resolve, 200));
        
      } catch (error) {
        console.error(`❌ Individual ${type} processing failed:`, error);
        
        // リトライロジック
        op.retryCount = (op.retryCount || 0) + 1;
        if (op.retryCount < this.batchSettings.maxRetries) {
          console.log(`🔄 Retrying ${type} operation (attempt ${op.retryCount})`);
          // キューに戻す
          if (type === 'reaction') {
            const key = `${op.rowIndex}-${op.reaction}`;
            this.reactionQueue.set(key, op);
          } else {
            const key = `highlight-${op.rowIndex}`;
            this.highlightQueue.set(key, op);
          }
        } else {
          console.error(`❌ ${type} operation failed permanently after ${op.retryCount} retries`);
        }
      }
    }
    
    this.updateBatchMetrics(operations.length, false);
  }

  /**
   * バッチ完了通知
   */
  notifyBatchCompletion(type, operations, result) {
    // カスタムイベントを発火
    const event = new CustomEvent('batchProcessingComplete', {
      detail: { type, operations, result }
    });
    document.dispatchEvent(event);
  }

  /**
   * バッチメトリクスの更新
   */
  updateBatchMetrics(batchSize, success) {
    this.batchMetrics.totalBatches++;
    this.batchMetrics.totalItemsProcessed += batchSize;
    
    if (success) {
      this.batchMetrics.successfulBatches++;
    } else {
      this.batchMetrics.failedBatches++;
    }
    
    this.batchMetrics.averageBatchSize = 
      this.batchMetrics.totalItemsProcessed / this.batchMetrics.totalBatches;
    
    // Core状態に報告
    this.coreState.updatePerformanceMetrics({
      batchMetrics: this.batchMetrics
    });
  }

  /**
   * 統計情報の取得
   */
  getStats() {
    return {
      queues: {
        reactions: this.reactionQueue.size,
        highlights: this.highlightQueue.size,
        total: this.reactionQueue.size + this.highlightQueue.size
      },
      processing: {
        isProcessing: this.isProcessing,
        hasScheduledProcessing: !!this.processingTimer
      },
      metrics: this.batchMetrics
    };
  }

  /**
   * キューのクリア
   */
  clearQueues() {
    this.reactionQueue.clear();
    this.highlightQueue.clear();
    
    if (this.processingTimer) {
      clearTimeout(this.processingTimer);
      this.processingTimer = null;
    }
    
    console.log('🧹 Batch queues cleared');
  }

  /**
   * リソースのクリーンアップ
   */
  destroy() {
    this.clearQueues();
    this.isProcessing = false;
    console.log('🧹 Batch processing manager destroyed');
  }
}

// グローバルエクスポート
window.OptimizedPollingManager = OptimizedPollingManager;
window.BatchProcessingManager = BatchProcessingManager;

