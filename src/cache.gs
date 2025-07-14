/**
 * @fileoverview 統合キャッシュマネージャー
 * アプリケーション全体のキャッシュ戦略を管理します。
 * シングルトンパターンで実装され、常に単一のインスタンス(cacheManager)を使用します。
 */

/**
 * キャッシュマネージャークラス
 */
class CacheManager {
  constructor() {
    this.scriptCache = CacheService.getScriptCache();
    this.memoCache = new Map(); // メモ化用の高速キャッシュ
    this.dependencyMap = new Map(); // キャッシュ依存関係マップ
    this.defaultTTL = 21600; // デフォルトTTL（6時間）
    
    // パフォーマンス監視用の統計情報
    this.stats = {
      hits: 0,
      misses: 0,
      errors: 0,
      totalOps: 0,
      lastReset: Date.now(),
      cascadeInvalidations: 0
    };
    
    // キャッシュ依存関係の初期化
    this._initializeDependencies();
  }

  /**
   * キャッシュから値を取得、なければ指定された関数で生成して保存します。
   * 最適化版：優先度ベースの2層キャッシュ
   * @param {string} key - キャッシュキー
   * @param {function} valueFn - 値を生成する関数
   * @param {object} [options] - オプション { ttl: number, enableMemoization: boolean, priority: string }
   * @returns {*} キャッシュされた値
   */
  get(key, valueFn, options = {}) {
    const { 
      ttl = this.defaultTTL, 
      enableMemoization = false,
      priority = 'normal' // 'high', 'normal', 'low'
    } = options;
    
    // パフォーマンス監視
    this.stats.totalOps++;

    // Input validation
    if (!key || typeof key !== 'string') {
      console.error(`[Cache] Invalid key: ${key}`);
      this.stats.errors++;
      return valueFn();
    }

    // 優先度に基づくキャッシュ戦略決定
    const useHighSpeedPath = priority === 'high' || enableMemoization;
    
    // 1. 高速パス：メモ化キャッシュ優先
    if (useHighSpeedPath && this.memoCache.has(key)) {
      try {
        const memoEntry = this.memoCache.get(key);
        if (!memoEntry.ttl || (memoEntry.createdAt + memoEntry.ttl * 1000 > Date.now())) {
          debugLog(`[Cache] High-speed hit for key: ${key}`);
          this.stats.hits++;
          return memoEntry.value;
        } else {
          // 期限切れエントリを削除
          this.memoCache.delete(key);
        }
      } catch (e) {
        console.warn(`[Cache] Memo cache access failed for key: ${key}`, e.message);
        this.stats.errors++;
        this.memoCache.delete(key);
      }
    }

    // 2. 標準パス：ScriptCache
    let cachedValue = null;
    let useScriptCache = true;
    
    try {
      cachedValue = this.scriptCache.get(key);
      if (cachedValue !== null) {
        debugLog(`[Cache] ScriptCache hit for key: ${key}`);
        this.stats.hits++;
        const parsedValue = JSON.parse(cachedValue);
        
        // 高優先度データは次回のためにメモ化
        if (useHighSpeedPath && this.memoCache.size < 100) { // メモリ制限
          this.memoCache.set(key, { value: parsedValue, createdAt: Date.now(), ttl });
        }
        return parsedValue;
      }
    } catch (e) {
      console.warn(`[Cache] ScriptCache access failed for key: ${key}`, e.message);
      this.stats.errors++;
      useScriptCache = false;
      
      // 破損エントリを安全に削除
      try {
        this.scriptCache.remove(key);
      } catch (removeError) {
        console.warn(`[Cache] Failed to remove corrupted entry: ${key}`, removeError.message);
      }
    }

    // 3. キャッシュミス：値生成
    debugLog(`[Cache] Miss for key: ${key}. Generating new value.`);
    this.stats.misses++;
    
    let newValue;
    try {
      newValue = valueFn();
    } catch (e) {
      console.error(`[Cache] Value generation failed for key: ${key}`, e.message);
      this.stats.errors++;
      throw e;
    }
    
    // 4. 効率的なキャッシュ保存
    this._efficientCacheStore(key, newValue, ttl, useHighSpeedPath, useScriptCache);

    // パフォーマンス監視（最適化：低頻度）
    if (this.stats.totalOps % 200 === 0) {
      const hitRate = (this.stats.hits / this.stats.totalOps * 100).toFixed(1);
      debugLog(`[Cache] Performance: ${hitRate}% hit rate, ${this.memoCache.size} memo entries`);
    }

    return newValue;
  }

  /**
   * 効率的なキャッシュ保存
   * @private
   */
  _efficientCacheStore(key, value, ttl, useHighSpeedPath, useScriptCache) {
    // 並列でキャッシュ保存を試行
    const promises = [];
    
    // メモ化キャッシュ保存（高速）
    if (useHighSpeedPath) {
      try {
        // メモリ制限チェック
        if (this.memoCache.size >= 150) {
          this._cleanupMemoCache();
        }
        this.memoCache.set(key, { value, createdAt: Date.now(), ttl });
      } catch (e) {
        console.warn(`[Cache] Memo cache store failed: ${e.message}`);
        this.stats.errors++;
      }
    }
    
    // ScriptCache保存（永続化）
    if (useScriptCache) {
      try {
        const stringValue = JSON.stringify(value);
        this.scriptCache.put(key, stringValue, ttl);
      } catch (e) {
        console.warn(`[Cache] ScriptCache store failed: ${e.message}`);
        this.stats.errors++;
      }
    }
  }

  /**
   * メモ化キャッシュのクリーンアップ
   * @private
   */
  _cleanupMemoCache() {
    const now = Date.now();
    let cleanedCount = 0;
    
    for (const [key, entry] of this.memoCache.entries()) {
      // 期限切れまたは古いエントリを削除
      if (entry.ttl && (entry.createdAt + entry.ttl * 1000 < now)) {
        this.memoCache.delete(key);
        cleanedCount++;
      }
    }
    
    // 期限切れがない場合、LRUに基づいて削除
    if (cleanedCount === 0 && this.memoCache.size > 100) {
      const entries = Array.from(this.memoCache.entries());
      entries.sort((a, b) => a[1].createdAt - b[1].createdAt);
      
      // 古い25%を削除
      const toDelete = Math.floor(entries.length * 0.25);
      for (let i = 0; i < toDelete; i++) {
        this.memoCache.delete(entries[i][0]);
        cleanedCount++;
      }
    }
    
    debugLog(`[Cache] Memo cache cleanup: removed ${cleanedCount} entries`);
  }

  /**
   * 複数のキーを一括で取得します。
   * @param {string[]} keys - キーの配列
   * @param {function} valuesFn - 見つからなかったキーの値を取得する関数
   * @param {object} [options] - オプション { ttl: number }
   * @returns {Object.<string, *>} キーと値のオブジェクト
   */
  batchGet(keys, valuesFn, options = {}) {
    const { ttl = this.defaultTTL } = options;
    const results = {};
    let missingKeys = [];

    try {
      const cachedValues = this.scriptCache.getAll(keys);
      for (const key in cachedValues) {
        results[key] = JSON.parse(cachedValues[key]);
      }
      missingKeys = keys.filter(k => !results.hasOwnProperty(k));
    } catch (e) {
      console.warn('[Cache] batchGet failed', e.message);
      missingKeys = keys;
    }

    if (missingKeys.length > 0) {
      const newValues = valuesFn(missingKeys);
      const newCacheValues = {};
      for (const key in newValues) {
        results[key] = newValues[key];
        newCacheValues[key] = JSON.stringify(newValues[key]);
      }
      this.scriptCache.putAll(newCacheValues, ttl);
    }

    return results;
  }

  /**
   * 指定されたキーのキャッシュを削除します。
   * @param {string} key - 削除するキャッシュキー
   */
  remove(key) {
    if (!key || typeof key !== 'string') {
      console.warn(`[Cache] Invalid key for removal: ${key}`);
      return;
    }
    
    try {
      this.scriptCache.remove(key);
    } catch (e) {
      console.warn(`[Cache] Failed to remove scriptCache for key: ${key}`, e.message);
    }
    
    try {
      this.memoCache.delete(key);
    } catch (e) {
      console.warn(`[Cache] Failed to remove memoCache for key: ${key}`, e.message);
    }
    
    debugLog(`[Cache] Removed cache for key: ${key}`);
  }

  /**
   * パターンに一致するキーのキャッシュを削除します。
   * @param {string} pattern - 削除するキーのパターン（部分一致）
   */
  clearByPattern(pattern) {
    if (!pattern || typeof pattern !== 'string') {
      console.warn(`[Cache] Invalid pattern for clearByPattern: ${pattern}`);
      return;
    }
    
    // メモ化キャッシュから一致するキーを削除
    const keysToRemove = [];
    let failedRemovals = 0;
    
    try {
      for (const key of this.memoCache.keys()) {
        if (key.includes(pattern)) {
          keysToRemove.push(key);
        }
      }
    } catch (e) {
      console.warn(`[Cache] Failed to iterate memoCache keys for pattern: ${pattern}`, e.message);
    }
    
    keysToRemove.forEach(key => {
      try {
        this.memoCache.delete(key);
        this.scriptCache.remove(key);
      } catch (e) {
        console.warn(`[Cache] Failed to remove key during pattern clear: ${key}`, e.message);
        failedRemovals++;
      }
    });
    
    debugLog(`[Cache] Cleared ${keysToRemove.length - failedRemovals} cache entries matching pattern: ${pattern} (${failedRemovals} failed)`);
  }

  /**
   * キャッシュ依存関係を初期化します。
   * @private
   */
  _initializeDependencies() {
    // ユーザー関連キャッシュの依存関係を定義
    const userDependencies = [
      'user_*',
      'status_*',
      'sheets_*',
      'config_*',
      'form_*'
    ];
    
    // スプレッドシート関連キャッシュの依存関係
    const spreadsheetDependencies = [
      'hdr_*',
      'data_*',
      'sheets_*'
    ];
    
    // 依存関係をマップに登録
    this.dependencyMap.set('user_change', userDependencies);
    this.dependencyMap.set('spreadsheet_change', spreadsheetDependencies);
    this.dependencyMap.set('form_change', ['form_*', 'status_*', 'user_*']);
    
    console.log('⚙️ [Cache] Dependency map initialized:', this.dependencyMap.size, 'relationships');
  }
  
  /**
   * 依存関係に基づいてカスケード無効化を実行します。
   * @param {string} changeType - 変更タイプ ('user_change', 'spreadsheet_change', 'form_change')
   * @param {string} [specificKey] - 特定のキーを指定した場合
   */
  invalidateDependents(changeType, specificKey = null) {
    const dependencies = this.dependencyMap.get(changeType);
    if (!dependencies) {
      console.warn(`[Cache] Unknown change type for dependency invalidation: ${changeType}`);
      return;
    }
    
    let invalidatedCount = 0;
    
    if (specificKey) {
      // 特定キーの無効化
      try {
        this.memoCache.delete(specificKey);
        this.scriptCache.remove(specificKey);
        invalidatedCount++;
      } catch (e) {
        console.warn(`[Cache] Failed to invalidate specific key: ${specificKey}`, e.message);
      }
    }
    
    // パターンベースの無効化
    dependencies.forEach(pattern => {
      try {
        this.clearByPattern(pattern.replace('*', ''));
        invalidatedCount++;
      } catch (e) {
        console.warn(`[Cache] Failed to clear pattern during cascade: ${pattern}`, e.message);
      }
    });
    
    this.stats.cascadeInvalidations++;
    console.log(`🔄 [Cache] Cascade invalidation completed: ${changeType}, ${invalidatedCount} patterns cleared`);
  }
  
  /**
   * 期限切れのキャッシュをクリアします（この機能はGASでは自動です）。
   * メモ化キャッシュをクリアする目的で実装します。
   */
  clearExpired() {
    this.memoCache.clear();
    debugLog('[Cache] Cleared memoization cache.');
  }
  
  /**
   * キャッシュキーをサニタイズします。
   * @param {string} key - 元のキー
   * @param {string} [namespace='default'] - ネームスペース
   * @returns {string} サニタイズされたキー
   */
  sanitizeKey(key, namespace = 'default') {
    if (typeof key !== 'string') {
      console.warn('[Cache] Key must be a string, got:', typeof key);
      key = String(key);
    }
    
    // 危険な文字を除去し、長さを制限
    const sanitized = key
      .replace(/[^a-zA-Z0-9\-_]/g, '_') // 英数字、ハイフン、アンダースコア以外をアンダースコアに変換
      .substring(0, 200); // 最大200文字に制限
    
    // ネームスペースを付与
    const namespacedKey = `${namespace}:${sanitized}`;
    
    // キーが変更された場合は警告
    if (namespacedKey !== `${namespace}:${key}`) {
      console.warn(`[Cache] Key sanitized: '${key}' -> '${namespacedKey}'`);
    }
    
    return namespacedKey;
  }
  
  /**
   * ネームスペース全体をクリアします。
   * @param {string} namespace - クリアするネームスペース
   */
  clearNamespace(namespace) {
    const pattern = `${namespace}:`;
    this.clearByPattern(pattern);
    console.log(`🖾️ [Cache] Cleared namespace: ${namespace}`);
  }

  /**
   * 全てのキャッシュを強制的にクリアします。
   * メモ化キャッシュとスクリプトキャッシュの両方をクリアします。
   */
  clearAll() {
    let memoCacheCleared = false;
    let scriptCacheCleared = false;
    
    try {
      // メモ化キャッシュをクリア
      this.memoCache.clear();
      memoCacheCleared = true;
      debugLog('[Cache] Cleared memoization cache.');
    } catch (e) {
      console.warn('[Cache] Failed to clear memoization cache:', e.message);
    }
    
    try {
      // スクリプトキャッシュを完全にクリア
      this.scriptCache.removeAll([]);
      scriptCacheCleared = true;
      debugLog('[Cache] Cleared script cache.');
    } catch (e) {
      console.warn('[Cache] Failed to clear script cache:', e.message);
    }
    
    // 統計をリセット
    try {
      this.resetStats();
    } catch (e) {
      console.warn('[Cache] Failed to reset stats:', e.message);
    }
    
    console.log(`[Cache] clearAll() completed - MemoCache: ${memoCacheCleared ? 'OK' : 'FAILED'}, ScriptCache: ${scriptCacheCleared ? 'OK' : 'FAILED'}`);
    
    return {
      memoCacheCleared,
      scriptCacheCleared,
      success: memoCacheCleared && scriptCacheCleared
    };
  }

  /**
   * キャッシュの健全性情報を取得します。
   * @returns {object} 健全性情報
   */
  getHealth() {
    const uptime = Date.now() - this.stats.lastReset;
    const hitRate = this.stats.totalOps > 0 ? (this.stats.hits / this.stats.totalOps * 100).toFixed(1) : 0;
    const errorRate = this.stats.totalOps > 0 ? (this.stats.errors / this.stats.totalOps * 100).toFixed(1) : 0;
    
    return {
      memoCacheSize: this.memoCache.size,
      status: this.stats.errors / this.stats.totalOps < 0.1 ? 'ok' : 'degraded',
      stats: {
        totalOperations: this.stats.totalOps,
        hits: this.stats.hits,
        misses: this.stats.misses,
        errors: this.stats.errors,
        hitRate: hitRate + '%',
        errorRate: errorRate + '%',
        uptimeMs: uptime
      }
    };
  }

  /**
   * キャッシュ統計をリセットします。
   */
  resetStats() {
    this.stats = {
      hits: 0,
      misses: 0,
      errors: 0,
      totalOps: 0,
      lastReset: Date.now()
    };
    debugLog('[Cache] Statistics reset');
  }
}

// --- シングルトンインスタンスの作成 ---
const cacheManager = new CacheManager();

// Google Apps Script環境でグローバルにアクセス可能にする
if (typeof global !== 'undefined') {
  global.cacheManager = cacheManager;
}

// ==============================================
// 後方互換性のためのラッパー関数（移行期間中）
// ==============================================

/**
 * @deprecated cacheManager.get() を使用してください
 */

/**
 * ヘッダーインデックスをキャッシュ付きで取得
 */
function getHeadersCached(spreadsheetId, sheetName) {
  const key = `hdr_${spreadsheetId}_${sheetName}`;
  const indices = cacheManager.get(key, () => {
    try {
      console.log(`[getHeadersCached] Starting for spreadsheetId: ${spreadsheetId}, sheetName: ${sheetName}`);
      
      var service = getSheetsService();
      if (!service) {
        throw new Error('Sheets service is not available');
      }
      
      var range = sheetName + '!1:1';
      console.log(`[getHeadersCached] Fetching range: ${range}`);
      
      // Use the updated API pattern consistent with other functions
      var url = service.baseUrl + '/' + spreadsheetId + '/values/' + encodeURIComponent(range);
      var response = UrlFetchApp.fetch(url, {
        headers: { 'Authorization': 'Bearer ' + service.accessToken }
      });
      
      var responseData = JSON.parse(response.getContentText());
      console.log(`[getHeadersCached] API response:`, JSON.stringify(responseData, null, 2));
      
      if (!responseData) {
        throw new Error('API response is null or undefined');
      }
      
      if (!responseData.values) {
        console.warn(`[getHeadersCached] No values in response for ${range}`);
        return {};
      }
      
      if (!responseData.values[0] || responseData.values[0].length === 0) {
        console.warn(`[getHeadersCached] Empty header row for ${range}`);
        return {};
      }
      
      var headers = responseData.values[0];
      console.log(`[getHeadersCached] Headers found:`, headers);
      
      var indices = {};
      
      // すべてのヘッダーのインデックスを生成（動的ヘッダーに対応）
      headers.forEach(function(headerName, index) {
        if (headerName && headerName.trim() !== '') {
          indices[headerName] = index;
          console.log(`[getHeadersCached] Mapped ${headerName} -> ${index}`);
        }
      });
      
      console.log(`[getHeadersCached] Final indices:`, indices);
      return indices;
    } catch (error) {
      console.error('getHeadersCached error:', error.toString());
      console.error('Error stack:', error.stack);
      return {};
    }
  }, { ttl: 1800, enableMemoization: true }); // 30分間キャッシュ + メモ化

  if (!indices || Object.keys(indices).length === 0) {
    cacheManager.remove(key);
    return {};
  }

  return indices;
}

/**
 * ユーザー関連キャッシュを無効化します。
 * @param {string} userId - ユーザーID
 * @param {string} email - 管理者メールアドレス
 * @param {string} [spreadsheetId] - 関連スプレッドシートID
 * @param {boolean} [clearPattern=false] - パターンベースのクリアを行うか
 */
function invalidateUserCache(userId, email, spreadsheetId, clearPattern) {
  const keysToRemove = [];
  
  if (userId) {
    keysToRemove.push('user_' + userId);
  }
  if (email) {
    keysToRemove.push('email_' + email);
  }
  if (spreadsheetId) {
    // スプレッドシートIDを含むキーを削除
    keysToRemove.push('hdr_' + spreadsheetId);
    keysToRemove.push('data_' + spreadsheetId);
    keysToRemove.push('sheets_' + spreadsheetId);
  }
  
  keysToRemove.forEach(key => {
    cacheManager.remove(key);
  });
  
  // さらに包括的なパターンマッチングが必要な場合
  if (clearPattern && spreadsheetId) {
    cacheManager.clearByPattern(spreadsheetId);
  }
  
  debugLog(`[Cache] Invalidated user cache for userId: ${userId}, email: ${email}, spreadsheetId: ${spreadsheetId}`);
}

/**
 * データベース関連キャッシュをクリアします。
 * エラー時やデータ整合性が必要な場合に使用します。
 */
function clearDatabaseCache() {
  try {
    // データベース関連のキャッシュパターンをクリア
    cacheManager.clearByPattern('user_');
    cacheManager.clearByPattern('email_');
    cacheManager.clearByPattern('hdr_');
    cacheManager.clearByPattern('data_');
    cacheManager.clearByPattern('sheets_');
    
    // サービスアカウントトークンは保持（パフォーマンス向上のため）
    
    debugLog('[Cache] Database cache cleared successfully');
  } catch (error) {
    console.error('clearDatabaseCache error:', error.message);
    // エラーが発生しても処理を継続
  }
}

