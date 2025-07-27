/**
 * StudyQuest Data Layer
 * GAS最適化済み - API通信、キャッシュ管理、データ処理
 */

/**
 * 統一キャッシュシステム（TTL付き）
 */
class UnifiedCache {
  constructor() {
    this.data = new Map();
    this.timestamps = new Map();
    this.hitCount = 0;
    this.missCount = 0;
  }
  
  set(key, value, ttl = 300000) { // 5分のデフォルトTTL
    this.data.set(key, value);
    this.timestamps.set(key, Date.now() + ttl);
    
    // GAS環境でのメモリ管理
    if (this.data.size > 100) { // 制限を設けてメモリ使用量を抑制
      this.cleanup();
    }
  }
  
  get(key) {
    if (!this.data.has(key)) {
      this.missCount++;
      return undefined;
    }
    
    const expiry = this.timestamps.get(key);
    if (expiry && Date.now() > expiry) {
      this.delete(key);
      this.missCount++;
      return undefined;
    }
    
    this.hitCount++;
    return this.data.get(key);
  }
  
  has(key) {
    return this.get(key) !== undefined;
  }
  
  delete(key) {
    this.data.delete(key);
    this.timestamps.delete(key);
  }
  
  clear() {
    this.data.clear();
    this.timestamps.clear();
    this.hitCount = 0;
    this.missCount = 0;
  }
  
  cleanup() {
    const now = Date.now();
    let cleanedCount = 0;
    
    for (const [key, expiry] of this.timestamps.entries()) {
      if (now > expiry) {
        this.delete(key);
        cleanedCount++;
      }
    }
    
    if (cleanedCount > 0) {
      console.log(`🧹 Cache cleanup: ${cleanedCount} items removed`);
    }
  }
  
  getStats() {
    return {
      size: this.data.size,
      hitRate: this.hitCount / (this.hitCount + this.missCount) || 0,
      hitCount: this.hitCount,
      missCount: this.missCount
    };
  }
}

/**
 * GAS API通信マネージャー
 */
class GASApiManager {
  constructor(coreState) {
    this.coreState = coreState;
    this.cache = new UnifiedCache();
    this.requestQueue = new Map();
    this.batchQueue = new Map();
    this.apiCallCount = 0;
    
    // GAS固有の設定
    this.gasSettings = {
      timeout: 30000, // 30秒タイムアウト
      retryAttempts: 3,
      retryDelay: 2000,
      batchSize: STUDY_QUEST_CONSTANTS.GAS_BATCH_SIZE_LIMIT
    };
  }

  /**
   * GAS関数の実行（キャッシュ付き）
   */
  async runGas(funcName, ...args) {
    // GAS実行時間チェック
    this.coreState.checkGASExecutionTime();
    
    const cacheKey = this.generateCacheKey(funcName, args);
    const isStateChanging = ['toggleHighlight', 'addReaction', 'endPublication'].includes(funcName);
    
    // キャッシュチェック（状態変更系でない場合）
    if (!isStateChanging) {
      const cached = this.cache.get(cacheKey);
      if (cached !== undefined) {
        console.log(`📦 Cache hit for ${funcName}`);
        return cached;
      }
    }
    
    try {
      this.apiCallCount++;
      this.coreState.updatePerformanceMetrics({ apiCallCount: this.apiCallCount });
      
      const startTime = Date.now();
      console.log(`🚀 GAS API call: ${funcName}`, args);
      
      const result = await this.executeGASFunction(funcName, args);
      const duration = Date.now() - startTime;
      
      console.log(`✅ GAS API success: ${funcName} (${duration}ms)`, result);
      
      // 成功時はキャッシュに保存（状態変更系でない場合）
      if (!isStateChanging && result !== undefined) {
        const ttl = this.getTTLForFunction(funcName);
        this.cache.set(cacheKey, result, ttl);
      }
      
      return result;
      
    } catch (error) {
      console.error(`❌ GAS API error: ${funcName}`, error);
      this.coreState.setError(error, `GAS API call: ${funcName}`);
      
      // GAS固有エラーの処理
      if (this.isGASSpecificError(error)) {
        return this.handleGASError(error, funcName, args);
      }
      
      throw error;
    }
  }

  /**
   * GAS関数の実際の実行
   */
  async executeGASFunction(funcName, args) {
    return new Promise((resolve, reject) => {
      const timeoutId = setTimeout(() => {
        reject(new Error(`GAS function ${funcName} timed out after ${this.gasSettings.timeout}ms`));
      }, this.gasSettings.timeout);
      
      try {
        google.script.run
          .withSuccessHandler((result) => {
            clearTimeout(timeoutId);
            resolve(result);
          })
          .withFailureHandler((error) => {
            clearTimeout(timeoutId);
            reject(new Error(error.message || 'GAS execution failed'));
          })[funcName](...args);
      } catch (error) {
        clearTimeout(timeoutId);
        reject(error);
      }
    });
  }

  /**
   * バッチ処理の実行
   */
  async executeBatch(operations) {
    if (operations.length === 0) return [];
    
    // GAS制約に合わせてバッチサイズを制限
    const batches = this.chunkArray(operations, this.gasSettings.batchSize);
    const results = [];
    
    for (const batch of batches) {
      try {
        const batchResult = await this.runGas('executeBatch', batch);
        results.push(...(batchResult || []));
      } catch (error) {
        console.warn('Batch execution failed, falling back to individual calls:', error);
        
        // バッチ失敗時は個別実行にフォールバック
        for (const operation of batch) {
          try {
            const result = await this.runGas(operation.funcName, ...operation.args);
            results.push(result);
          } catch (individualError) {
            console.error('Individual operation failed:', individualError);
            results.push({ error: individualError.message });
          }
        }
      }
    }
    
    return results;
  }

  /**
   * データ取得の最適化されたメソッド
   */
  async getPublishedSheetData(classFilter, sortOrder, adminMode, bypassCache = false) {
    const userId = this.coreState.state.userId;
    
    if (bypassCache) {
      this.cache.delete(this.generateCacheKey('getPublishedSheetData', [userId, classFilter, sortOrder, adminMode]));
    }
    
    return await this.runGas('getPublishedSheetData', userId, classFilter, sortOrder, adminMode, bypassCache);
  }

  /**
   * 増分データ取得
   */
  async getIncrementalSheetData(classFilter, sortOrder, adminMode, sinceRowCount) {
    const userId = this.coreState.state.userId;
    return await this.runGas('getIncrementalSheetData', userId, classFilter, sortOrder, adminMode, sinceRowCount);
  }

  /**
   * データ数の取得
   */
  async getDataCount(classFilter, sortOrder, adminMode) {
    const userId = this.coreState.state.userId;
    return await this.runGas('getDataCount', userId, classFilter, sortOrder, adminMode);
  }

  /**
   * リアクションの追加
   */
  async addReaction(rowIndex, reaction, sheetName) {
    const userId = this.coreState.state.userId;
    return await this.runGas('addReaction', userId, rowIndex, reaction, sheetName);
  }

  /**
   * バッチリアクション処理
   */
  async addReactionBatch(batchOperations) {
    const userId = this.coreState.state.userId;
    return await this.runGas('addReactionBatch', userId, batchOperations);
  }

  /**
   * 管理者権限チェック
   */
  async checkAdmin() {
    return await this.runGas('checkAdmin');
  }

  /**
   * キャッシュクリア
   */
  async clearCache() {
    this.cache.clear();
    return await this.runGas('clearCache');
  }

  /**
   * ユーティリティメソッド
   */
  generateCacheKey(funcName, args) {
    return `${funcName}:${JSON.stringify(args)}`;
  }

  getTTLForFunction(funcName) {
    const ttlMap = {
      'getPublishedSheetData': 30000, // 30秒
      'getDataCount': 15000, // 15秒
      'getIncrementalSheetData': 10000, // 10秒
      'checkAdmin': 300000, // 5分
      'getAvailableSheets': 300000 // 5分
    };
    
    return ttlMap[funcName] || STUDY_QUEST_CONSTANTS.CACHE_TTL_MS;
  }

  isGASSpecificError(error) {
    const gasErrorPatterns = [
      'Script function not found',
      'Permission denied',
      'Service invoked too many times',
      'Exceeded maximum execution time',
      'Lock could not be obtained'
    ];
    
    return gasErrorPatterns.some(pattern => 
      error.message.includes(pattern)
    );
  }

  handleGASError(error, funcName, args) {
    console.warn(`🔧 Handling GAS-specific error for ${funcName}:`, error.message);
    
    // エラー種別に応じた処理
    if (error.message.includes('Exceeded maximum execution time')) {
      // 実行時間超過の場合は軽量な代替処理
      return this.getFailsafeResponse(funcName);
    }
    
    if (error.message.includes('Service invoked too many times')) {
      // API制限の場合は遅延後にリトライ
      return this.retryWithDelay(funcName, args, 5000);
    }
    
    // その他のエラーは上位に委譲
    throw error;
  }

  getFailsafeResponse(funcName) {
    const failsafeResponses = {
      'getPublishedSheetData': { data: [], status: 'limited' },
      'getDataCount': { count: 0, status: 'limited' },
      'getIncrementalSheetData': { data: [], newCount: 0, status: 'limited' }
    };
    
    return failsafeResponses[funcName] || { status: 'error' };
  }

  async retryWithDelay(funcName, args, delay) {
    await new Promise(resolve => setTimeout(resolve, delay));
    return await this.runGas(funcName, ...args);
  }

  chunkArray(array, chunkSize) {
    const chunks = [];
    for (let i = 0; i < array.length; i += chunkSize) {
      chunks.push(array.slice(i, i + chunkSize));
    }
    return chunks;
  }

  /**
   * 統計情報の取得
   */
  getStats() {
    return {
      cache: this.cache.getStats(),
      apiCallCount: this.apiCallCount,
      queueSize: this.requestQueue.size,
      batchQueueSize: this.batchQueue.size
    };
  }

  /**
   * リソースのクリーンアップ
   */
  destroy() {
    this.cache.clear();
    this.requestQueue.clear();
    this.batchQueue.clear();
    console.log('🧹 Data layer destroyed');
  }
}

// グローバルエクスポート
window.UnifiedCache = UnifiedCache;
window.GASApiManager = GASApiManager;

