/**
 * @fileoverview 超高速パフォーマンス最適化システム
 * 2024年最新のGAS技術とV8ランタイム機能を活用
 */

/**
 * 非同期バッチ処理マネージャー
 * 6分実行制限内での最大効率化を実現
 */
class PerformanceOptimizer {
  
  /**
   * 実行時間制限を考慮したバッチ処理
   * @param {Array} items - 処理対象アイテム
   * @param {Function} processor - 処理関数
   * @param {Object} options - オプション設定
   */
  static timeBoundedBatch(items, processor, options = {}) {
    const {
      maxExecutionTime = 330000, // 5.5分（バッファ付き）
      batchSize = 50,
      progressCallback = null
    } = options;
    
    const startTime = Date.now();
    const results = [];
    let processedCount = 0;
    
    // バッチごとに処理
    for (let i = 0; i < items.length; i += batchSize) {
      // 実行時間チェック
      if (Date.now() - startTime > maxExecutionTime) {
        console.warn(`実行時間制限により処理を中断: ${processedCount}/${items.length}`);
        break;
      }
      
      const batch = items.slice(i, i + batchSize);
      
      try {
        // バッチ処理実行
        const batchResults = this._processBatch(batch, processor);
        results.push(...batchResults);
        processedCount += batch.length;
        
        // 進捗報告
        if (progressCallback) {
          progressCallback(processedCount, items.length);
        }
        
        // 小休止（API制限対策）
        if (i + batchSize < items.length) {
          Utilities.sleep(50); // 50ms休止
        }
        
      } catch (error) {
        console.error(`バッチ処理エラー (${i}-${i + batchSize}):`, error.message);
        // エラー時は指数バックオフで再試行
        const retryResults = this._retryWithBackoff(batch, processor, 3);
        results.push(...retryResults);
      }
    }
    
    return {
      results: results,
      processed: processedCount,
      total: items.length,
      executionTime: Date.now() - startTime
    };
  }
  
  /**
   * 指数バックオフ付き再試行
   */
  static _retryWithBackoff(batch, processor, maxRetries = 3) {
    let delay = 1000; // 初期1秒
    
    for (let attempt = 0; attempt < maxRetries; attempt++) {
      try {
        if (attempt > 0) {
          console.log(`再試行 ${attempt}/${maxRetries}, 待機: ${delay}ms`);
          Utilities.sleep(delay);
          delay *= 2; // 指数的に増加
        }
        
        return this._processBatch(batch, processor);
      } catch (error) {
        if (attempt === maxRetries - 1) {
          console.error('最大再試行回数に達しました:', error.message);
          return []; // 失敗時は空配列
        }
      }
    }
    return [];
  }
  
  /**
   * バッチ処理実行
   */
  static _processBatch(batch, processor) {
    return batch.map(item => {
      try {
        return processor(item);
      } catch (error) {
        console.error('アイテム処理エラー:', error.message);
        return null;
      }
    }).filter(result => result !== null);
  }
}

/**
 * 超高速Sheets API最適化クラス
 */
class OptimizedSheetsAPI {
  
  constructor(accessToken) {
    this.accessToken = accessToken;
    this.baseUrl = 'https://sheets.googleapis.com/v4/spreadsheets';
    this.requestCount = 0;
    this.lastRequestTime = 0;
  }
  
  /**
   * レート制限対応付きAPI呼び出し
   */
  async makeRequest(url, options = {}) {
    // レート制限対策（100リクエスト/100秒）
    const now = Date.now();
    const timeSinceLastRequest = now - this.lastRequestTime;
    
    if (this.requestCount >= 90 && timeSinceLastRequest < 100000) {
      const waitTime = 100000 - timeSinceLastRequest;
      console.log(`レート制限対策で待機: ${waitTime}ms`);
      Utilities.sleep(waitTime);
      this.requestCount = 0;
    }
    
    try {
      const response = UrlFetchApp.fetch(url, {
        ...options,
        headers: {
          'Authorization': `Bearer ${this.accessToken}`,
          ...options.headers
        }
      });
      
      this.requestCount++;
      this.lastRequestTime = Date.now();
      
      if (response.getResponseCode() >= 400) {
        throw new Error(`API Error: ${response.getResponseCode()} - ${response.getContentText()}`);
      }
      
      return JSON.parse(response.getContentText());
    } catch (error) {
      console.error('API リクエストエラー:', error.message);
      throw error;
    }
  }
  
  /**
   * 超高速バッチ取得（並列処理対応）
   */
  async batchGetOptimized(spreadsheetId, ranges) {
    // 100範囲制限を考慮してチャンク分割
    const chunks = this._chunkArray(ranges, 100);
    const allResults = [];
    
    for (const chunk of chunks) {
      const url = `${this.baseUrl}/${spreadsheetId}/values:batchGet?` +
        chunk.map(range => `ranges=${encodeURIComponent(range)}`).join('&');
      
      const result = await this.makeRequest(url);
      allResults.push(...(result.valueRanges || []));
    }
    
    return { valueRanges: allResults };
  }
  
  /**
   * 超高速バッチ更新
   */
  async batchUpdateOptimized(spreadsheetId, requests) {
    // リクエストサイズ制限を考慮
    const chunks = this._chunkArray(requests, 1000);
    const results = [];
    
    for (const chunk of chunks) {
      const url = `${this.baseUrl}/${spreadsheetId}/values:batchUpdate`;
      
      const result = await this.makeRequest(url, {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify({
          data: chunk,
          valueInputOption: 'RAW'
        })
      });
      
      results.push(result);
    }
    
    return results;
  }
  
  /**
   * 配列チャンク分割ユーティリティ
   */
  _chunkArray(array, chunkSize) {
    const chunks = [];
    for (let i = 0; i < array.length; i += chunkSize) {
      chunks.push(array.slice(i, i + chunkSize));
    }
    return chunks;
  }
}

/**
 * メモリ効率最適化
 */
class MemoryOptimizer {
  
  /**
   * 大量データの効率的処理
   */
  static processLargeDataset(data, processor, options = {}) {
    const {
      chunkSize = 1000,
      clearMemoryInterval = 5000
    } = options;
    
    let processedCount = 0;
    const results = [];
    
    // ガベージコレクション対策
    for (let i = 0; i < data.length; i += chunkSize) {
      const chunk = data.slice(i, i + chunkSize);
      const chunkResults = chunk.map(processor);
      results.push(...chunkResults);
      
      processedCount += chunk.length;
      
      // 定期的にメモリクリア
      if (processedCount % clearMemoryInterval === 0) {
        // V8のガベージコレクションを促進
        chunk.length = 0;
        Utilities.sleep(1);
      }
    }
    
    return results;
  }
  
  /**
   * オブジェクトプール（再利用パターン）
   */
  static createObjectPool(factory, resetFn, initialSize = 10) {
    const pool = [];
    
    // 初期プール作成
    for (let i = 0; i < initialSize; i++) {
      pool.push(factory());
    }
    
    return {
      acquire: () => {
        return pool.length > 0 ? pool.pop() : factory();
      },
      
      release: (obj) => {
        resetFn(obj);
        pool.push(obj);
      },
      
      size: () => pool.length
    };
  }
}

/**
 * 並列処理シミュレーター（GAS制限内）
 */
class ParallelProcessor {
  
  /**
   * 疑似並列処理（時分割実行）
   */
  static timeSlicedParallel(tasks, options = {}) {
    const {
      sliceTimeMs = 100,
      maxConcurrent = 5
    } = options;
    
    const results = new Array(tasks.length);
    const activeTasks = [];
    let taskIndex = 0;
    
    const processSlice = () => {
      const sliceStart = Date.now();
      
      // 新しいタスクを開始
      while (activeTasks.length < maxConcurrent && taskIndex < tasks.length) {
        const currentIndex = taskIndex++;
        activeTasks.push({
          index: currentIndex,
          task: tasks[currentIndex],
          iterator: null
        });
      }
      
      // アクティブタスクを実行
      for (let i = activeTasks.length - 1; i >= 0; i--) {
        const activeTask = activeTasks[i];
        
        try {
          // タスクが完了したか確認
          if (typeof activeTask.task === 'function') {
            results[activeTask.index] = activeTask.task();
            activeTasks.splice(i, 1);
          }
        } catch (error) {
          console.error(`タスク ${activeTask.index} エラー:`, error.message);
          results[activeTask.index] = { error: error.message };
          activeTasks.splice(i, 1);
        }
        
        // タイムスライス制限チェック
        if (Date.now() - sliceStart > sliceTimeMs) {
          break;
        }
      }
      
      return activeTasks.length > 0 || taskIndex < tasks.length;
    };
    
    // すべてのタスクが完了するまで実行
    while (processSlice()) {
      Utilities.sleep(1); // 1ms休止
    }
    
    return results;
  }
}

/**
 * パフォーマンス監視とプロファイリング
 */
class PerformanceProfiler {
  
  constructor() {
    this.metrics = new Map();
    this.startTimes = new Map();
  }
  
  /**
   * 処理時間測定開始
   */
  start(label) {
    this.startTimes.set(label, Date.now());
  }
  
  /**
   * 処理時間測定終了
   */
  end(label) {
    const startTime = this.startTimes.get(label);
    if (startTime) {
      const duration = Date.now() - startTime;
      
      if (!this.metrics.has(label)) {
        this.metrics.set(label, []);
      }
      
      this.metrics.get(label).push(duration);
      this.startTimes.delete(label);
      
      return duration;
    }
    return 0;
  }
  
  /**
   * 統計レポート生成
   */
  getReport() {
    const report = {};
    
    this.metrics.forEach((times, label) => {
      const sum = times.reduce((a, b) => a + b, 0);
      const avg = sum / times.length;
      const min = Math.min(...times);
      const max = Math.max(...times);
      
      report[label] = {
        count: times.length,
        total: sum,
        average: avg,
        min: min,
        max: max
      };
    });
    
    return report;
  }
  
  /**
   * メトリクスクリア
   */
  clear() {
    this.metrics.clear();
    this.startTimes.clear();
  }
}

// グローバルプロファイラーインスタンス
const globalProfiler = new PerformanceProfiler();