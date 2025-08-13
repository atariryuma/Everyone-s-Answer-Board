/**
 * @fileoverview 統合キャッシュマネージャー - 全キャッシュシステムを統合
 * cache.gs、unifiedExecutionCache.gs、spreadsheetCache.gs の機能を統合
 * 既存の全ての関数をそのまま保持し、互換性を完全維持
 */

// =============================================================================
// SECTION 1: 統合キャッシュマネージャークラス（元cache.gs）
// =============================================================================

/**
 * キャッシュマネージャークラス
 */
class CacheManager {
  constructor() {
    this.scriptCache = CacheService.getScriptCache();
    this.memoCache = new Map(); // メモ化用の高速キャッシュ
    this.defaultTTL = 21600; // デフォルトTTL（6時間）

    // パフォーマンス監視用の統計情報
    this.stats = {
      hits: 0,
      misses: 0,
      errors: 0,
      totalOps: 0,
      lastReset: Date.now()
    };
  }

  /**
   * キャッシュから値を取得、なければ指定された関数で生成して保存します（最適化版）
   * @param {string} key - キャッシュキー
   * @param {function} valueFn - 値を生成する関数
   * @param {object} [options] - オプション { ttl: number, enableMemoization: boolean, usePropertiesFallback: boolean }
   * @returns {*} キャッシュされた値
   */
  get(key, valueFn, options = {}) {
    const { 
      ttl = this.defaultTTL, 
      enableMemoization = false, 
      usePropertiesFallback = false 
    } = options;

    // パフォーマンス監視とバリデーション
    this.stats.totalOps++;
    if (!this.validateKey(key)) {
      this.stats.errors++;
      return valueFn ? valueFn() : null;
    }

    // 階層化キャッシュアクセス（最適化）
    const cacheResult = this.getFromCacheHierarchy(key, enableMemoization, ttl, usePropertiesFallback);
    if (cacheResult.found) {
      this.stats.hits++;
      return cacheResult.value;
    }

    // キャッシュミス: 新しい値を生成
    debugLog(`[Cache] Miss for key: ${key}. Generating new value.`);
    this.stats.misses++;

    let newValue;

    try {
      newValue = valueFn();
    } catch (e) {
      errorLog(`[Cache] Value generation failed for key: ${key}`, e.message);
      this.stats.errors++;
      throw e;
    }

    // キャッシュ保存（エラーが発生しても値は返す）
    try {
      const stringValue = JSON.stringify(newValue);
      this.scriptCache.put(key, stringValue, ttl);
      if (enableMemoization) {
        this.memoCache.set(key, { value: newValue, createdAt: Date.now(), ttl });
      }
    } catch (e) {
      errorLog(`[Cache] Failed to cache value for key: ${key}`, e.message);
      this.stats.errors++;
      // キャッシュ保存に失敗しても値は返す
    }

    // パフォーマンス監視ログ（低頻度）
    if (this.stats.totalOps % 100 === 0) {
      const hitRate = (this.stats.hits / this.stats.totalOps * 100).toFixed(1);
      debugLog(`[Cache] Performance: ${hitRate}% hit rate (${this.stats.hits}/${this.stats.totalOps}), ${this.stats.errors} errors`);
    }

    return newValue;
  }

  // =============================================================================
  // 最適化されたヘルパー関数群
  // =============================================================================

  /**
   * キーの妥当性検証
   * @param {string} key - 検証するキー
   * @returns {boolean} 妥当性
   */
  validateKey(key) {
    if (!key || typeof key !== 'string') {
      errorLog(`[Cache] Invalid key: ${key}`);
      return false;
    }
    if (key.length > 100) {
      warnLog(`[Cache] Key too long: ${key}`);
      return false;
    }
    return true;
  }

  /**
   * 階層化キャッシュから値を取得（PropertiesService統合版）
   * @param {string} key - キー
   * @param {boolean} enableMemoization - メモ化キャッシュ有効
   * @param {number} ttl - TTL
   * @param {boolean} usePropertiesFallback - PropertiesServiceフォールバック有効
   * @returns {Object} { found: boolean, value: any }
   */
  getFromCacheHierarchy(key, enableMemoization, ttl, usePropertiesFallback = false) {
    // Level 1: メモ化キャッシュ（最高速）
    if (enableMemoization && this.memoCache.has(key)) {
      try {
        const memoEntry = this.memoCache.get(key);
        if (!memoEntry.ttl || (memoEntry.createdAt + memoEntry.ttl * 1000 > Date.now())) {
          debugLog(`[Cache] L1(Memo) hit: ${key}`);
          return { found: true, value: memoEntry.value };
        } else {
          this.memoCache.delete(key); // 期限切れを削除
        }
      } catch (e) {
        warnLog(`[Cache] L1(Memo) error: ${key}`, e.message);
        this.memoCache.delete(key);
      }
    }

    // Level 2: Apps Scriptキャッシュ
    try {
      const cachedValue = this.scriptCache.get(key);
      if (cachedValue !== null) {
        debugLog(`[Cache] L2(Script) hit: ${key}`);
        const parsedValue = this.parseScriptCacheValue(key, cachedValue);

        // メモ化キャッシュに昇格
        if (enableMemoization && parsedValue !== null) {
          this.memoCache.set(key, { value: parsedValue, createdAt: Date.now(), ttl });
        }

        return { found: true, value: parsedValue };
      }
    } catch (e) {
      warnLog(`[Cache] L2(Script) error: ${key}`, e.message);
      this.stats.errors++;
      this.handleCorruptedCacheEntry(key);
    }

    // Level 3: PropertiesService フォールバック（設定値用）
    if (usePropertiesFallback) {
      try {
        const propsValue = PropertiesService.getScriptProperties().getProperty(key);
        if (propsValue !== null) {
          debugLog(`[Cache] L3(Properties) hit: ${key}`);
          const parsedValue = this.parsePropertiesValue(propsValue);
          
          // 上位キャッシュに昇格（長期TTL設定）
          this.setToCacheHierarchy(key, parsedValue, Math.max(ttl, 3600), enableMemoization);
          
          return { found: true, value: parsedValue };
        }
      } catch (e) {
        warnLog(`[Cache] L3(Properties) error: ${key}`, e.message);
        this.stats.errors++;
      }
    }

    return { found: false, value: null };
  }

  /**
   * ScriptCacheの値を安全にパース
   * @param {string} key - キー
   * @param {string} cachedValue - キャッシュされた値
   * @returns {any} パースされた値
   */
  parseScriptCacheValue(key, cachedValue) {
    // 特定のキーは生文字列として扱う
    const rawStringKeys = ['WEB_APP_URL'];
    if (rawStringKeys.includes(key)) {
      return cachedValue;
    }

    try {
      return JSON.parse(cachedValue);
    } catch (e) {
      warnLog(`[Cache] Parse failed for ${key}, returning raw string`);
      return cachedValue;
    }
  }

  /**
   * 破損したキャッシュエントリを処理
   * @param {string} key - キー
   */
  handleCorruptedCacheEntry(key) {
    try {
      this.scriptCache.remove(key);
      this.memoCache.delete(key);
      debugLog(`[Cache] Cleaned corrupted entry: ${key}`);
    } catch (removeError) {
      warnLog(`[Cache] Failed to clean corrupted entry: ${key}`, removeError.message);
    }
  }

  /**
   * 複数のキーを一括で取得します。
   * @param {string[]} keys - キーの配列
   * @param {function} valuesFn - 見つからなかったキーの値を取得する関数
   * @param {object} [options] - オプション { ttl: number }
   * @returns {Object.<string, *>} キーと値のオブジェクト
   */
  async batchGet(keys, valuesFn, options = {}) {
    const { ttl = this.defaultTTL } = options;
    
    try {
      // 統一バッチ処理システムでの処理を試行
      if (typeof unifiedBatchProcessor !== 'undefined') {
        const currentUserId = Session.getActiveUser().getEmail();
        const cacheOperations = keys.map(key => ({ key: key }));
        
        const cachedResults = unifiedBatchProcessor.batchCacheOperation(
          'get', 
          cacheOperations, 
          currentUserId,
          { concurrency: 5 }
        );
        
        const results = {};
        const missingKeys = [];
        
        for (let i = 0; i < keys.length; i++) {
          const key = keys[i];
          const cached = cachedResults[i];
          
          if (cached && !cached.error && cached !== null) {
            try {
              results[key] = JSON.parse(cached);
            } catch (parseError) {
              results[key] = cached;
            }
          } else {
            missingKeys.push(key);
          }
        }
        
        if (missingKeys.length > 0) {
          const newValues = Promise.resolve(valuesFn(missingKeys));
          const setCacheOps = Object.keys(newValues).map(key => ({
            key: key,
            value: JSON.stringify(newValues[key]),
            ttl: ttl
          }));
          
          for (let key in newValues) {
            results[key] = newValues[key];
          }
          
          if (setCacheOps.length > 0) {
            unifiedBatchProcessor.batchCacheOperation('set', setCacheOps, currentUserId, { concurrency: 5 });
          }
        }
        
        return results;
      }
    } catch (error) {
      warnLog('[Cache] 統一バッチ処理エラー:', error.message);
    }
    
    // フォールバック: 従来実装
    const results = {};
    let missingKeys = [];

    try {
      const cachedValues = this.scriptCache.getAll(keys);
      for (let key in cachedValues) {
        results[key] = JSON.parse(cachedValues[key]);
      }
      missingKeys = keys.filter(k => !results.hasOwnProperty(k));
    } catch (e) {
      warnLog('[Cache] batchGet failed', e.message);
      missingKeys = keys;
    }

    if (missingKeys.length > 0) {
      const newValues = valuesFn(missingKeys);
      const newCacheValues = {};
      for (let key in newValues) {
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
      warnLog(`[Cache] Invalid key for removal: ${key}`);
      return;
    }

    try {
      this.scriptCache.remove(key);
    } catch (e) {
      warnLog(`[Cache] Failed to remove scriptCache for key: ${key}`, e.message);
    }

    try {
      this.memoCache.delete(key);
    } catch (e) {
      warnLog(`[Cache] Failed to remove memoCache for key: ${key}`, e.message);
    }

    debugLog(`[Cache] Removed cache for key: ${key}`);
  }

  /**
   * スプレッドシート関連のキャッシュを無効化します（リアクション・ハイライト操作後）
   * @param {string} spreadsheetId - スプレッドシートID
   * @param {string} sheetName - シート名（オプション）
   */
  invalidateSheetData(spreadsheetId, sheetName = null) {
    try {
      // シート全体のデータキャッシュを無効化
      const patterns = [
        `batchGet_${spreadsheetId}_`,
        `sheet_data_${spreadsheetId}_`,
        `published_data_${spreadsheetId}_`
      ];

      if (sheetName) {
        patterns.push(`batchGet_${spreadsheetId}_["'${sheetName}'!A:Z"]`);
      }

      let removedCount = 0;
      for (const pattern of patterns) {
        // メモキャッシュから削除
        for (const key of this.memoCache.keys()) {
          if (key.includes(pattern)) {
            this.memoCache.delete(key);
            removedCount++;
          }
        }

        // スクリプトキャッシュからも削除を試行
        if (this.scriptCache) {
          try {
            this.scriptCache.remove(pattern);
          } catch (e) {
            // スクリプトキャッシュでは正確なキー一致のみ削除可能
          }
        }
      }

      debugLog(`[Cache] Invalidated sheet data cache for ${spreadsheetId}, removed ${removedCount} entries`);
      return removedCount;
    } catch (e) {
      warnLog(`[Cache] Failed to invalidate sheet data cache: ${e.message}`);
      this.stats.errors++;
      return 0;
    }
  }

  /**
   * パターンに一致するキーのキャッシュを安全に削除します。
   * @param {string} pattern - 削除するキーのパターン（部分一致）
   * @param {object} [options] - オプション { strict: boolean, maxKeys: number }
   */
  clearByPattern(pattern, options = {}) {
    const { strict = false, maxKeys = 1000 } = options;

    if (!pattern || typeof pattern !== 'string') {
      warnLog(`[Cache] Invalid pattern for clearByPattern: ${pattern}`);
      return 0;
    }

    // セキュリティチェック: 過度に広範囲な削除を防ぐ
    if (!strict && (pattern.length < 3 || pattern === '*' || pattern === '.*')) {
      warnLog(`[Cache] Pattern too broad for safe removal: ${pattern}. Use strict=true to override.`);
      return 0;
    }

    // メモ化キャッシュから一致するキーを安全に削除
    const keysToRemove = [];
    let failedRemovals = 0;
    let skippedCount = 0;

    try {
      for (const key of this.memoCache.keys()) {
        // 安全性チェック: 重要なシステムキーは保護
        if (this._isProtectedKey(key)) {
          skippedCount++;
          continue;
        }

        if (key.includes(pattern)) {
          keysToRemove.push(key);

          // 大量削除の防止
          if (keysToRemove.length >= maxKeys) {
            warnLog(`[Cache] Reached maxKeys limit (${maxKeys}) for pattern: ${pattern}`);
            break;
          }
        }
      }
    } catch (e) {
      warnLog(`[Cache] Failed to iterate memoCache keys for pattern: ${pattern}`, e.message);
      this.stats.errors++;
    }

    // 削除前の確認ログ
    if (keysToRemove.length > 100) {
      warnLog(`[Cache] Large pattern deletion detected: ${keysToRemove.length} keys for pattern: ${pattern}`);
    }

    keysToRemove.forEach(key => {
      try {
        this.memoCache.delete(key);
        this.scriptCache.remove(key);
      } catch (e) {
        warnLog(`[Cache] Failed to remove key during pattern clear: ${key}`, e.message);
        failedRemovals++;
        this.stats.errors++;
      }
    });

    const successCount = keysToRemove.length - failedRemovals;

    debugLog(`[Cache] Pattern clear completed: ${successCount} removed, ${failedRemovals} failed, ${skippedCount} protected (pattern: ${pattern})`);

    return successCount;
  }

  /**
   * 保護すべきキーかどうかを判定
   * @private
   */
  _isProtectedKey(key) {
    const protectedPatterns = [
      'SA_TOKEN_CACHE',           // サービスアカウントトークン
      'WEB_APP_URL',             // WebアプリURL
      'SYSTEM_CONFIG',           // システム設定
      'DOMAIN_INFO'              // ドメイン情報
    ];

    return protectedPatterns.some(pattern => key.includes(pattern));
  }

  /**
   * インテリジェント無効化：関連するキャッシュエントリを安全に連鎖的に無効化
   * @param {string} entityType - エンティティタイプ (user, spreadsheet, form等)
   * @param {string} entityId - エンティティID
   * @param {Array<string>} relatedIds - 関連するエンティティID群
   * @param {object} [options] - オプション { dryRun: boolean, maxRelated: number }
   */
  invalidateRelated(entityType, entityId, relatedIds = [], options = {}) {
    const { dryRun = false, maxRelated = 50 } = options;
    const invalidationLog = {
      entityType,
      entityId,
      totalRemoved: 0,
      errors: [],
      patterns: [],
      startTime: Date.now()
    };

    try {
      // 入力検証
      if (!entityType || !entityId) {
        throw new Error('entityType and entityId are required');
      }

      // 関連IDの数制限
      if (relatedIds.length > maxRelated) {
        warnLog(`[Cache] Too many related IDs (${relatedIds.length}), limiting to ${maxRelated}`);
        relatedIds = relatedIds.slice(0, maxRelated);
      }

      debugLog(`🔗 関連キャッシュ無効化開始: ${entityType}/${entityId} (${dryRun ? 'DRY RUN' : 'LIVE'})`);

      // 1. メインエンティティのキャッシュ無効化
      const patterns = this._getInvalidationPatterns(entityType, entityId);
      invalidationLog.patterns.push(...patterns);

      patterns.forEach(pattern => {
        try {
          if (!dryRun) {
            const removed = this.clearByPattern(pattern, { strict: false, maxKeys: 100 });
            invalidationLog.totalRemoved += removed;
          } else {
            debugLog(`[Cache] DRY RUN: Would clear pattern: ${pattern}`);
          }
        } catch (error) {
          invalidationLog.errors.push(`Pattern ${pattern}: ${error.message}`);
        }
      });

      // 2. 関連エンティティのキャッシュ無効化（検証付き）
      relatedIds.forEach(relatedId => {
        if (!relatedId || relatedId === entityId) {
          return; // 無効な関連IDまたは自分自身はスキップ
        }

        try {
          const relatedPatterns = this._getInvalidationPatterns(entityType, relatedId);
          relatedPatterns.forEach(pattern => {
            if (!dryRun) {
              const removed = this.clearByPattern(pattern, { strict: false, maxKeys: 50 });
              invalidationLog.totalRemoved += removed;
            } else {
              debugLog(`[Cache] DRY RUN: Would clear related pattern: ${pattern}`);
            }
          });
        } catch (error) {
          invalidationLog.errors.push(`Related ID ${relatedId}: ${error.message}`);
        }
      });

      // 3. クロスエンティティ関連キャッシュの無効化
      if (!dryRun) {
        this._invalidateCrossEntityCache(entityType, entityId, relatedIds, invalidationLog);
      }

      invalidationLog.duration = Date.now() - invalidationLog.startTime;

      if (invalidationLog.errors.length > 0) {
        warnLog(`⚠️ 関連キャッシュ無効化で一部エラー: ${entityType}/${entityId}`, invalidationLog.errors);
      } else {
        debugLog(`✅ 関連キャッシュ無効化完了: ${entityType}/${entityId} (${invalidationLog.totalRemoved} entries, ${invalidationLog.duration}ms)`);
      }

      return invalidationLog;

    } catch (error) {
      invalidationLog.errors.push(`Fatal: ${error.message}`);
      invalidationLog.duration = Date.now() - invalidationLog.startTime;
      errorLog(`❌ 関連キャッシュ無効化致命的エラー: ${entityType}/${entityId}`, error);
      this.stats.errors++;
      return invalidationLog;
    }
  }

  /**
   * エンティティタイプに基づく無効化パターンを取得
   * @private
   */
  _getInvalidationPatterns(entityType, entityId) {
    const patterns = [];

    switch (entityType) {
      case 'user':
        patterns.push(
          `user_${entityId}*`,
          `userInfo_${entityId}*`,
          `config_${entityId}*`,
          `appUrls_${entityId}*`,
          `status_${entityId}*`
        );
        break;

      case 'spreadsheet':
        patterns.push(
          `sheets_${entityId}*`,
          `batchGet_${entityId}*`,
          `headers_${entityId}*`,
          `spreadsheet_${entityId}*`
        );
        break;

      case 'form':
        patterns.push(
          `form_${entityId}*`,
          `formUrl_${entityId}*`,
          `formResponse_${entityId}*`
        );
        break;

      default:
        patterns.push(`${entityType}_${entityId}*`);
    }

    return patterns;
  }

  /**
   * クロスエンティティキャッシュの安全な無効化
   * @private
   */
  _invalidateCrossEntityCache(entityType, entityId, relatedIds, invalidationLog) {
    try {
      // ユーザー変更時は関連するスプレッドシート・フォームキャッシュも無効化
      if (entityType === 'user' && relatedIds.length > 0) {
        relatedIds.forEach(relatedId => {
          if (relatedId && typeof relatedId === 'string' && relatedId.length > 0) {
            try {
              const removed1 = this.clearByPattern(`user_${entityId}_spreadsheet_${relatedId}`, { maxKeys: 20 });
              const removed2 = this.clearByPattern(`user_${entityId}_form_${relatedId}`, { maxKeys: 20 });
              invalidationLog.totalRemoved += (removed1 + removed2);
            } catch (error) {
              invalidationLog.errors.push(`Cross-entity user-${relatedId}: ${error.message}`);
            }
          }
        });
      }

      // スプレッドシート変更時は関連するユーザーキャッシュも無効化
      if (entityType === 'spreadsheet' && relatedIds.length > 0) {
        relatedIds.forEach(userId => {
          if (userId && typeof userId === 'string' && userId.length > 0) {
            try {
              const removed1 = this.clearByPattern(`user_${userId}_spreadsheet_${entityId}`, { maxKeys: 20 });
              const removed2 = this.clearByPattern(`config_${userId}`, { maxKeys: 10 }); // ユーザー設定も無効化
              invalidationLog.totalRemoved += (removed1 + removed2);
            } catch (error) {
              invalidationLog.errors.push(`Cross-entity spreadsheet-${userId}: ${error.message}`);
            }
          }
        });
      }

      // フォーム変更時の追加ロジック
      if (entityType === 'form' && relatedIds.length > 0) {
        relatedIds.forEach(userId => {
          if (userId && typeof userId === 'string' && userId.length > 0) {
            try {
              const removed = this.clearByPattern(`user_${userId}_form_${entityId}`, { maxKeys: 10 });
              invalidationLog.totalRemoved += removed;
            } catch (error) {
              invalidationLog.errors.push(`Cross-entity form-${userId}: ${error.message}`);
            }
          }
        });
      }

    } catch (error) {
      invalidationLog.errors.push(`Cross-entity fatal: ${error.message}`);
      warnLog('クロスエンティティキャッシュ無効化エラー:', error.message);
      this.stats.errors++;
    }
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
      warnLog('[Cache] Failed to clear memoization cache:', e.message);
    }

    try {
      // スクリプトキャッシュを完全にクリア
      this.scriptCache.removeAll();
      scriptCacheCleared = true;
      debugLog('[Cache] Cleared script cache.');
    } catch (e) {
      warnLog('[Cache] Failed to clear script cache:', e.message);
    }

    // 統計をリセット
    try {
      this.resetStats();
    } catch (e) {
      warnLog('[Cache] Failed to reset stats:', e.message);
    }

    debugLog(`[Cache] clearAll() completed - MemoCache: ${memoCacheCleared ? 'OK' : 'FAILED'}, ScriptCache: ${scriptCacheCleared ? 'OK' : 'FAILED'}`);

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

  /**
   * PropertiesServiceの値をパース（JSONまたはプリミティブ）
   * @param {string} value - 値
   * @returns {*} パースされた値
   */
  parsePropertiesValue(value) {
    if (!value) return value;
    
    try {
      // JSON形式の場合はパース
      return JSON.parse(value);
    } catch (e) {
      // プリミティブ値の場合はそのまま返す
      return value;
    }
  }

  /**
   * 階層化キャッシュに値を設定
   * @param {string} key - キー
   * @param {*} value - 値
   * @param {number} ttl - TTL
   * @param {boolean} enableMemoization - メモ化有効
   */
  setToCacheHierarchy(key, value, ttl, enableMemoization) {
    try {
      // Level 2: Apps Scriptキャッシュ
      this.scriptCache.put(key, JSON.stringify(value), ttl);
      
      // Level 1: メモ化キャッシュ
      if (enableMemoization) {
        this.memoCache.set(key, { 
          value: value, 
          createdAt: Date.now(), 
          ttl: ttl 
        });
      }
    } catch (e) {
      warnLog(`[Cache] setToCacheHierarchy error: ${key}`, e.message);
    }
  }

  /**
   * PropertiesService統合キャッシュヘルパー
   * 設定値などの永続的なデータに特化
   * @param {string} key - キー
   * @param {function} valueFn - 値を生成する関数（オプション）
   * @param {object} options - オプション
   * @returns {*} 値
   */
  getConfig(key, valueFn, options = {}) {
    const { 
      ttl = 3600,  // 設定値は1時間キャッシュ
      enableMemoization = true,
      usePropertiesFallback = true 
    } = options;
    
    return this.get(key, valueFn, { ttl, enableMemoization, usePropertiesFallback });
  }
}

// フロントエンド統一キャッシュ制御機能を追加
CacheManager.prototype.clearInProgress = false;
CacheManager.prototype.pendingClears = [];

/**
 * 統一キャッシュクリア - 全キャッシュシステムを順次クリア
 * @param {Object} options - クリアオプション
 * @returns {Promise} クリア完了Promise
 */
CacheManager.prototype.clearAllFrontendCaches = function(options = {}) {
  const { force = false, timeout = 10000 } = options;

  // 既にクリア中の場合は待機
  if (this.clearInProgress && !force) {
    if (this.debugMode) {
      debugLog('🔄 Cache clear already in progress, waiting...');
    }
    return new Promise((resolve, reject) => {
      this.pendingClears.push({ resolve, reject });
    });
  }

  this.clearInProgress = true;
  const startTime = Date.now();

  return new Promise(async (resolve, reject) => {
    try {
      if (this.debugMode) {
        debugLog('🗑️ Starting unified cache clear process');
      }

      // キャッシュクリア操作のリスト（優先順位順）
      const clearOperations = [
        {
          name: 'UnifiedCache',
          operation: () => {
            if (typeof window !== 'undefined' && window.unifiedCache && typeof window.unifiedCache.clear === 'function') {
              window.unifiedCache.clear();
              return true;
            }
            return false;
          }
        },
        {
          name: 'GasOptimizerCache',
          operation: () => {
            if (typeof window !== 'undefined' && window.gasOptimizer && typeof window.gasOptimizer.clearCache === 'function') {
              window.gasOptimizer.clearCache();
              return true;
            }
            return false;
          }
        },
        {
          name: 'SharedUtilitiesCache',
          operation: () => {
            if (typeof window !== 'undefined' && window.sharedUtilities && window.sharedUtilities.cache && typeof window.sharedUtilities.cache.clear === 'function') {
              window.sharedUtilities.cache.clear();
              return true;
            }
            return false;
          }
        },
        {
          name: 'DOMElementCache',
          operation: () => {
            if (typeof window !== 'undefined' && window.sharedUtilities && window.sharedUtilities.dom && typeof window.sharedUtilities.dom.clearElementCache === 'function') {
              window.sharedUtilities.dom.clearElementCache();
              return true;
            }
            return false;
          }
        },
        {
          name: 'ThrottleDebounceCache',
          operation: () => {
            if (typeof window !== 'undefined' && window.sharedUtilities && window.sharedUtilities.throttle && typeof window.sharedUtilities.throttle.clearAll === 'function') {
              window.sharedUtilities.throttle.clearAll();
              return true;
            }
            return false;
          }
        },
        {
          name: 'ScriptCache',
          operation: () => {
            try {
              this.clearAll();
              return true;
            } catch (error) {
              return false;
            }
          }
        }
      ];

      const results = [];
      
      // 順次実行でキャッシュクリア（競合回避）
      for (const clearOp of clearOperations) {
        try {
          const success = clearOp.operation();
          results.push({ name: clearOp.name, success });
          
          if (this.debugMode && success) {
            debugLog(`✅ ${clearOp.name} cleared successfully`);
          }
          
          // 各操作間に短い間隔を設ける
          if (typeof Utilities !== 'undefined') {
            Utilities.sleep(50);
          }
          
        } catch (error) {
          warnLog(`⚠️ Failed to clear ${clearOp.name}:`, error);
          results.push({ name: clearOp.name, success: false, error: error.message });
        }
      }

      const successCount = results.filter(r => r.success).length;
      const totalTime = Date.now() - startTime;

      if (this.debugMode) {
        debugLog(`🎉 Cache clear completed: ${successCount}/${results.length} caches cleared in ${totalTime}ms`);
      }

      // 待機中のクリア要求を解決
      this.resolvePendingClears(results);

      resolve({
        success: true,
        results,
        successCount,
        totalCount: results.length,
        duration: totalTime
      });

    } catch (error) {
      errorLog('❌ Unified cache clear failed:', error);
      this.rejectPendingClears(error);
      reject(error);
    } finally {
      this.clearInProgress = false;
    }
  });
};

/**
 * 特定タイプのキャッシュのみクリア
 * @param {string} cacheType - キャッシュタイプ
 * @returns {Promise} クリア結果
 */
CacheManager.prototype.clearSpecificCache = function(cacheType) {
  return new Promise((resolve, reject) => {
    const cacheOperations = {
      unified: () => typeof window !== 'undefined' && window.unifiedCache?.clear(),
      gasOptimizer: () => typeof window !== 'undefined' && window.gasOptimizer?.clearCache(),
      sharedUtilities: () => typeof window !== 'undefined' && window.sharedUtilities?.cache?.clear(),
      domElements: () => typeof window !== 'undefined' && window.sharedUtilities?.dom?.clearElementCache(),
      throttleDebounce: () => typeof window !== 'undefined' && window.sharedUtilities?.throttle?.clearAll(),
      script: () => this.clearAll()
    };

    const operation = cacheOperations[cacheType];
    if (!operation) {
      reject(new Error(`Unknown cache type: ${cacheType}`));
      return;
    }

    try {
      operation();
      if (this.debugMode) {
        debugLog(`✅ ${cacheType} cache cleared`);
      }
      resolve({ success: true, cacheType });
    } catch (error) {
      warnLog(`⚠️ Failed to clear ${cacheType} cache:`, error);
      resolve({ success: false, cacheType, error: error.message });
    }
  });
};

/**
 * キャッシュ状態の診断
 * @returns {Object} 診断結果
 */
CacheManager.prototype.diagnoseFrontendCache = function() {
  const caches = {
    unifiedCache: {
      available: typeof window !== 'undefined' && !!(window.unifiedCache && window.unifiedCache.clear),
      size: typeof window !== 'undefined' ? (window.unifiedCache?.size || 'unknown') : 'unavailable'
    },
    gasOptimizer: {
      available: typeof window !== 'undefined' && !!(window.gasOptimizer && window.gasOptimizer.clearCache),
      size: typeof window !== 'undefined' ? (window.gasOptimizer?.cache?.size || 'unknown') : 'unavailable'
    },
    sharedUtilities: {
      available: typeof window !== 'undefined' && !!(window.sharedUtilities && window.sharedUtilities.cache),
      size: typeof window !== 'undefined' ? (window.sharedUtilities?.cache?.size || 'unknown') : 'unavailable'
    },
    domElements: {
      available: typeof window !== 'undefined' && !!(window.sharedUtilities && window.sharedUtilities.dom),
      size: 'unknown'
    },
    scriptCache: {
      available: true,
      size: this.memoCache.size
    }
  };

  return {
    clearInProgress: this.clearInProgress,
    pendingClears: this.pendingClears.length,
    caches,
    health: this.getHealth()
  };
};

/**
 * 待機中のクリア要求を解決
 */
CacheManager.prototype.resolvePendingClears = function(results) {
  const pending = this.pendingClears.splice(0);
  pending.forEach(({ resolve }) => resolve(results));
};

/**
 * 待機中のクリア要求を拒否
 */
CacheManager.prototype.rejectPendingClears = function(error) {
  const pending = this.pendingClears.splice(0);
  pending.forEach(({ reject }) => reject(error));
};

// =============================================================================
// SECTION 2: 統一実行レベルキャッシュ管理システム（元unifiedExecutionCache.gs）
// =============================================================================

/**
 * 統一実行レベルキャッシュ管理システム
 * 全モジュール間でのキャッシュ共有と効率的な管理
 */
class UnifiedExecutionCache {
  constructor() {
    this.userInfoCache = null;
    this.sheetsServiceCache = null;
    this.lastUserIdKey = null;
    this.executionStartTime = Date.now();
    this.maxLifetime = 300000; // 5分間の最大ライフタイム
  }

  /**
   * ユーザー情報キャッシュを取得
   * @param {string} userId - ユーザーID
   * @returns {object|null} キャッシュされたユーザー情報
   */
  getUserInfo(userId) {
    if (this.isExpired()) {
      this.clearAll();
      return null;
    }

    if (this.userInfoCache && this.lastUserIdKey === userId) {
      debugLog(`✅ 統一キャッシュヒット: ユーザー情報 (${userId})`);
      return this.userInfoCache;
    }

    return null;
  }

  /**
   * ユーザー情報をキャッシュに保存
   * @param {string} userId - ユーザーID
   * @param {object} userInfo - ユーザー情報
   */
  setUserInfo(userId, userInfo) {
    this.userInfoCache = userInfo;
    this.lastUserIdKey = userId;
    debugLog(`💾 統一キャッシュ保存: ユーザー情報 (${userId})`);
  }

  /**
   * SheetsServiceキャッシュを取得
   * @returns {object|null} キャッシュされたSheetsService
   */
  getSheetsService() {
    if (this.isExpired()) {
      this.clearAll();
      return null;
    }

    if (this.sheetsServiceCache) {
      debugLog(`✅ 統一キャッシュヒット: SheetsService`);
      return this.sheetsServiceCache;
    }

    return null;
  }

  /**
   * SheetsServiceをキャッシュに保存
   * @param {object} service - SheetsService
   */
  setSheetsService(service) {
    this.sheetsServiceCache = service;
    debugLog(`💾 統一キャッシュ保存: SheetsService`);
  }

  /**
   * ユーザー情報キャッシュをクリア
   */
  clearUserInfo() {
    this.userInfoCache = null;
    this.lastUserIdKey = null;
    debugLog(`🗑️ 統一キャッシュクリア: ユーザー情報`);
  }

  /**
   * SheetsServiceキャッシュをクリア
   */
  clearSheetsService() {
    this.sheetsServiceCache = null;
    debugLog(`🗑️ 統一キャッシュクリア: SheetsService`);
  }

  /**
   * 全キャッシュをクリア
   */
  clearAll() {
    this.clearUserInfo();
    this.clearSheetsService();
    debugLog(`🗑️ 統一キャッシュ全クリア`);
  }

  /**
   * キャッシュが期限切れかチェック
   * @returns {boolean} 期限切れの場合true
   */
  isExpired() {
    return (Date.now() - this.executionStartTime) > this.maxLifetime;
  }

  /**
   * キャッシュ統計情報を取得
   * @returns {object} キャッシュ統計
   */
  getStats() {
    return {
      hasUserInfo: !!this.userInfoCache,
      hasSheetsService: !!this.sheetsServiceCache,
      lastUserIdKey: this.lastUserIdKey,
      executionTime: Date.now() - this.executionStartTime,
      isExpired: this.isExpired()
    };
  }

  /**
   * 統一キャッシュマネージャーとの同期
   * @param {string} operation - 操作タイプ ('userDataChange', 'configChange', 'systemChange')
   */
  syncWithUnifiedCache(operation) {
    if (typeof cacheManager !== 'undefined' && cacheManager) {
      try {
        switch (operation) {
          case 'userDataChange':
            if (this.lastUserIdKey) {
              cacheManager.remove(`user_${this.lastUserIdKey}`);
              cacheManager.remove(`userinfo_${this.lastUserIdKey}`);
            }
            break;
          case 'configChange':
            cacheManager.remove('system_config');
            break;
          case 'systemChange':
            // システム全体のキャッシュクリア
            break;
        }
        debugLog(`🔄 統一キャッシュマネージャーと同期: ${operation}`);
      } catch (error) {
        debugLog(`⚠️ 統一キャッシュマネージャー同期エラー: ${error.message}`);
      }
    }
  }
}

// =============================================================================
// SECTION 3: SpreadsheetApp最適化システム（元spreadsheetCache.gs）
// =============================================================================

// メモリキャッシュ - 実行セッション内で有効
// ユーザー分離メモリキャッシュ - マルチテナント対応
let userScopedMemoryCache = {}; // { userId: { spreadsheetId: data } }

// キャッシュ設定
const SPREADSHEET_CACHE_CONFIG = {
  MEMORY_CACHE_TTL: 300000,    // 5分間（メモリキャッシュ）
  SESSION_CACHE_TTL: 1800000,  // 30分間（PropertiesServiceキャッシュ）
  MAX_CACHE_SIZE: 50,          // 最大キャッシュエントリ数
  CACHE_KEY_PREFIX: 'ss_cache_'
};

/**
 * マルチテナント対応キャッシュされたSpreadsheetオブジェクトを取得
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} userId - ユーザーID（テナント分離用・必須）
 * @param {Object} options - オプション設定
 * @param {boolean} options.forceRefresh - 強制リフレッシュフラグ
 * @param {boolean} options.isViewMode - 生徒閲覧モードフラグ
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} Spreadsheetオブジェクト
 */
function getCachedSpreadsheet(spreadsheetId, userId, options = {}) {
  const { forceRefresh = false, isViewMode = false } = options;
  
  // セキュリティ検証: userIdは必須（テナント分離のため）
  if (!userId || typeof userId !== 'string') {
    throw new Error('SECURITY_ERROR: userId is required for tenant isolation');
  }
  
  if (!spreadsheetId || typeof spreadsheetId !== 'string') {
    throw new Error('有効なスプレッドシートIDが必要です');
  }

  // テナント境界検証
  const currentUserId = Session.getActiveUser().getEmail();
  if (typeof multiTenantSecurity !== 'undefined') {
    const operation = isViewMode ? 'view_mode_access' : 'spreadsheet_cache_access';
    if (!multiTenantSecurity.validateTenantBoundary(currentUserId, userId, operation)) {
      throw new Error('TENANT_BOUNDARY_VIOLATION: Unauthorized cross-tenant access');
    }
  }

  // テナント分離キャッシュキー生成
  const cacheKey = buildUserScopedKey('ss_cache', userId, spreadsheetId);
  const now = Date.now();

  // ユーザー分離メモリキャッシュの初期化
  if (!userScopedMemoryCache[userId]) {
    userScopedMemoryCache[userId] = {};
  }

  // 強制リフレッシュの場合はキャッシュをクリア
  if (forceRefresh) {
    delete userScopedMemoryCache[userId][spreadsheetId];
    try {
      resilientCacheOperation(
        () => PropertiesService.getScriptProperties().deleteProperty(cacheKey),
        'PropertiesService.deleteProperty'
      );
    } catch (error) {
      debugLog('キャッシュクリアエラー:', error.message);
    }
  }

  // Phase 1: ユーザー分離メモリキャッシュをチェック
  const memoryEntry = userScopedMemoryCache[userId][spreadsheetId];
  if (memoryEntry && (now - memoryEntry.timestamp) < SPREADSHEET_CACHE_CONFIG.MEMORY_CACHE_TTL) {
    debugLog('✅ SpreadsheetApp.openById メモリキャッシュヒット:', `User:${userId.substring(0, 10)}, SS:${spreadsheetId.substring(0, 10)}`);
    return memoryEntry.spreadsheet;
  }

  // Phase 2: セッションキャッシュをチェック
  try {
    const sessionCacheData = resilientCacheOperation(
      () => PropertiesService.getScriptProperties().getProperty(cacheKey),
      'PropertiesService.getProperty'
    );
    if (sessionCacheData) {
      const sessionEntry = JSON.parse(sessionCacheData);
      if ((now - sessionEntry.timestamp) < SPREADSHEET_CACHE_CONFIG.SESSION_CACHE_TTL) {
        // セッションキャッシュからSpreadsheetオブジェクトを再構築
        const spreadsheet = resilientSpreadsheetOperation(
          () => SpreadsheetApp.openById(spreadsheetId),
          'SpreadsheetApp.openById'
        );
        
        // ユーザー分離メモリキャッシュに保存
        userScopedMemoryCache[userId][spreadsheetId] = {
          spreadsheet: spreadsheet,
          timestamp: now
        };
        
        debugLog('✅ SpreadsheetApp.openById セッションキャッシュヒット:', `User:${userId.substring(0, 10)}, SS:${spreadsheetId.substring(0, 10)}`);
        return spreadsheet;
      }
    }
  } catch (error) {
    debugLog('セッションキャッシュ読み込みエラー:', error.message);
  }

  // Phase 3: 新規取得とキャッシュ保存
  debugLog('🔄 SpreadsheetApp.openById 新規取得:', `User:${userId.substring(0, 10)}, SS:${spreadsheetId.substring(0, 10)}`);
  
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    
    // ユーザー分離メモリキャッシュに保存
    userScopedMemoryCache[userId][spreadsheetId] = {
      spreadsheet: spreadsheet,
      timestamp: now
    };
    
    // セッションキャッシュに保存（軽量データのみ）
    try {
      const sessionData = {
        spreadsheetId: spreadsheetId,
        timestamp: now,
        url: spreadsheet.getUrl()
      };
      
      PropertiesService.getScriptProperties().setProperty(
        cacheKey, 
        JSON.stringify(sessionData)
      );
    } catch (sessionError) {
      debugLog('セッションキャッシュ保存エラー:', sessionError.message);
    }
    
    // キャッシュサイズ管理
    cleanupOldCacheEntries();
    
    return spreadsheet;
    
  } catch (error) {
    logError(error, 'SpreadsheetApp.openById', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.DATABASE, { spreadsheetId });
    throw new Error(`スプレッドシートの取得に失敗しました: ${error.message}`);
  }
}

/**
 * ユーザー分離対応の古いキャッシュエントリをクリーンアップ
 * @param {string} [targetUserId] - 特定ユーザーのみクリーンアップする場合のユーザーID
 */
function cleanupOldCacheEntries(targetUserId = null) {
  const now = Date.now();
  let totalCleaned = 0;
  
  try {
    // ユーザー分離メモリキャッシュのクリーンアップ
    const userIds = targetUserId ? [targetUserId] : Object.keys(userScopedMemoryCache);
    
    userIds.forEach(userId => {
      if (!userScopedMemoryCache[userId]) return;
      
      const userCache = userScopedMemoryCache[userId];
      const spreadsheetIds = Object.keys(userCache);
      
      // ユーザー別キャッシュサイズ制限チェック
      if (spreadsheetIds.length > SPREADSHEET_CACHE_CONFIG.MAX_CACHE_SIZE) {
        const sortedEntries = spreadsheetIds.map(spreadsheetId => ({
          spreadsheetId: spreadsheetId,
          timestamp: userCache[spreadsheetId].timestamp
        })).sort((a, b) => b.timestamp - a.timestamp);
        
        // 古いエントリを削除
        const entriesToDelete = sortedEntries.slice(SPREADSHEET_CACHE_CONFIG.MAX_CACHE_SIZE);
        entriesToDelete.forEach(entry => {
          delete userCache[entry.spreadsheetId];
          totalCleaned++;
        });
        
        debugLog(`🧹 ユーザー${userId.substring(0, 10)}のメモリキャッシュクリーンアップ: ${entriesToDelete.length}件削除`);
      }
      
      // 期限切れエントリの削除
      spreadsheetIds.forEach(spreadsheetId => {
        const entry = userCache[spreadsheetId];
        if (entry && (now - entry.timestamp) > SPREADSHEET_CACHE_CONFIG.MEMORY_CACHE_TTL) {
          delete userCache[spreadsheetId];
          totalCleaned++;
        }
      });
      
      // 空のユーザーキャッシュを削除
      if (Object.keys(userCache).length === 0) {
        delete userScopedMemoryCache[userId];
        debugLog(`🗑️ 空のユーザーキャッシュを削除: ${userId.substring(0, 10)}`);
      }
    });
    
    if (totalCleaned > 0) {
      debugLog(`✅ キャッシュクリーンアップ完了: ${totalCleaned}件のエントリを削除`);
    }
    
  } catch (error) {
    debugLog('⚠️ キャッシュクリーンアップエラー:', error.message);
  }
}

/**
 * マルチテナント対応特定スプレッドシートのキャッシュを無効化
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} [userId] - ユーザーID（指定時はそのユーザーのみ無効化）
 */
function invalidateSpreadsheetCache(spreadsheetId, userId = null) {
  if (!spreadsheetId) return;
  
  let invalidatedCount = 0;
  
  try {
    // ユーザー分離メモリキャッシュから削除
    if (userId) {
      // 特定ユーザーのキャッシュのみ削除
      if (userScopedMemoryCache[userId] && userScopedMemoryCache[userId][spreadsheetId]) {
        delete userScopedMemoryCache[userId][spreadsheetId];
        invalidatedCount++;
        debugLog(`🗑️ ユーザー${userId.substring(0, 10)}のSpreadsheetCache無効化: ${spreadsheetId.substring(0, 10)}`);
      }
    } else {
      // 全ユーザーのキャッシュから削除
      Object.keys(userScopedMemoryCache).forEach(uid => {
        if (userScopedMemoryCache[uid] && userScopedMemoryCache[uid][spreadsheetId]) {
          delete userScopedMemoryCache[uid][spreadsheetId];
          invalidatedCount++;
        }
      });
      debugLog(`🗑️ 全ユーザーのSpreadsheetCache無効化: ${spreadsheetId.substring(0, 10)} (${invalidatedCount}件)`);
    }
    
    // セッションキャッシュから削除（ユーザー分離対応）
    try {
      const props = PropertiesService.getScriptProperties();
      const allProps = props.getProperties();
      
      Object.keys(allProps).forEach(key => {
        if (key.includes(spreadsheetId) && key.startsWith('MT_ss_cache_')) {
          props.deleteProperty(key);
          invalidatedCount++;
        }
      });
    } catch (sessionError) {
      debugLog('セッションキャッシュ無効化エラー:', sessionError.message);
    }
    
  } catch (error) {
    debugLog('スプレッドシートキャッシュ無効化エラー:', error.message);
  }
  
  return invalidatedCount;
}

/**
 * マルチテナント対応全スプレッドシートキャッシュをクリア
 * @param {string} [userId] - 特定ユーザーのみクリアする場合のユーザーID
 */
function clearAllSpreadsheetCache(userId = null) {
  let clearedCount = 0;
  
  try {
    // ユーザー分離メモリキャッシュクリア
    if (userId) {
      // 特定ユーザーのみクリア
      if (userScopedMemoryCache[userId]) {
        clearedCount = Object.keys(userScopedMemoryCache[userId]).length;
        delete userScopedMemoryCache[userId];
        debugLog(`🧹 ユーザー${userId.substring(0, 10)}のSpreadsheetCacheクリア: ${clearedCount}件`);
      }
    } else {
      // 全ユーザーのキャッシュクリア
      Object.keys(userScopedMemoryCache).forEach(uid => {
        clearedCount += Object.keys(userScopedMemoryCache[uid] || {}).length;
      });
      userScopedMemoryCache = {};
      debugLog(`🧹 全ユーザーのメモリキャッシュクリア: ${clearedCount}件`);
    }
    
    // セッションキャッシュクリア（マルチテナント対応）
    const props = PropertiesService.getScriptProperties();
    const allProps = props.getProperties();
    let sessionClearedCount = 0;
    
    Object.keys(allProps).forEach(key => {
      // マルチテナント対応キー（MT_ss_cache_）とレガシーキー（ss_cache_）両方をクリア
      if (key.startsWith('MT_ss_cache_') || key.startsWith(SPREADSHEET_CACHE_CONFIG.CACHE_KEY_PREFIX)) {
        if (userId) {
          // 特定ユーザーのキーのみ削除
          if (key.includes(userId.substring(0, 16))) { // ハッシュ化されたuserIdの一部で判定
            props.deleteProperty(key);
            sessionClearedCount++;
          }
        } else {
          // 全てのキーを削除
          props.deleteProperty(key);
          sessionClearedCount++;
        }
      }
    });
    
    debugLog(`🧹 セッションキャッシュクリア: ${sessionClearedCount}件`);
    debugLog(`✅ SpreadsheetCacheクリア完了: メモリ${clearedCount}件 + セッション${sessionClearedCount}件`);
    
  } catch (error) {
    debugLog('SpreadsheetCacheクリアエラー:', error.message);
  }
  
  return clearedCount;
}

/**
 * マルチテナント対応キャッシュ統計情報を取得
 * @param {string} [userId] - 特定ユーザーの統計のみ取得する場合のユーザーID
 * @returns {Object} キャッシュ統計
 */
function getSpreadsheetCacheStats(userId = null) {
  const now = Date.now();
  let memoryEntries = 0;
  let userCount = 0;
  let expiredEntries = 0;
  
  try {
    if (userId) {
      // 特定ユーザーの統計
      if (userScopedMemoryCache[userId]) {
        memoryEntries = Object.keys(userScopedMemoryCache[userId]).length;
        userCount = 1;
        
        // 期限切れエントリのカウント
        Object.values(userScopedMemoryCache[userId]).forEach(entry => {
          if ((now - entry.timestamp) > SPREADSHEET_CACHE_CONFIG.MEMORY_CACHE_TTL) {
            expiredEntries++;
          }
        });
      }
    } else {
      // 全ユーザーの統計
      userCount = Object.keys(userScopedMemoryCache).length;
      Object.values(userScopedMemoryCache).forEach(userCache => {
        memoryEntries += Object.keys(userCache).length;
        
        // 期限切れエントリのカウント
        Object.values(userCache).forEach(entry => {
          if ((now - entry.timestamp) > SPREADSHEET_CACHE_CONFIG.MEMORY_CACHE_TTL) {
            expiredEntries++;
          }
        });
      });
    }
    
    // セッションキャッシュ統計
    let sessionEntries = 0;
    const allProps = PropertiesService.getScriptProperties().getProperties();
    sessionEntries = Object.keys(allProps).filter(key => 
      key.startsWith('MT_ss_cache_') || key.startsWith(SPREADSHEET_CACHE_CONFIG.CACHE_KEY_PREFIX)
    ).length;
    
    return {
      memoryEntries: memoryEntries,
      sessionEntries: sessionEntries,
      userCount: userCount,
      expiredEntries: expiredEntries,
      maxCacheSize: SPREADSHEET_CACHE_CONFIG.MAX_CACHE_SIZE,
      memoryTTL: SPREADSHEET_CACHE_CONFIG.MEMORY_CACHE_TTL,
      sessionTTL: SPREADSHEET_CACHE_CONFIG.SESSION_CACHE_TTL,
      isMultiTenant: true
    };
    
  } catch (error) {
    debugLog('キャッシュ統計取得エラー:', error.message);
    return {
      error: error.message,
      memoryEntries: 0,
      sessionEntries: 0,
      userCount: 0,
      expiredEntries: 0,
      maxCacheSize: SPREADSHEET_CACHE_CONFIG.MAX_CACHE_SIZE,
      memoryTTL: SPREADSHEET_CACHE_CONFIG.MEMORY_CACHE_TTL,
      sessionTTL: SPREADSHEET_CACHE_CONFIG.SESSION_CACHE_TTL,
      isMultiTenant: true
    };
  }
}

/**
 * SpreadsheetApp.openById()の最適化されたラッパー関数（マルチテナント対応）
 * 既存コードの置き換え用
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} [userId] - ユーザーID（省略時は現在のユーザー）
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} Spreadsheetオブジェクト
 */
function openSpreadsheetOptimized(spreadsheetId, userId = null) {
  const effectiveUserId = userId || Session.getActiveUser().getEmail();
  return getCachedSpreadsheet(spreadsheetId, effectiveUserId);
}

// =============================================================================
// SECTION 4: 統合インスタンス作成と後方互換性関数
// =============================================================================

// --- シングルトンインスタンスの作成 ---
const cacheManager = new CacheManager();

// Google Apps Script環境でグローバルにアクセス可能にする
if (typeof global !== 'undefined') {
  global.cacheManager = cacheManager;
}

// グローバル統一キャッシュインスタンス
let globalUnifiedCache = null;

/**
 * 統一実行キャッシュのインスタンスを取得
 * @returns {UnifiedExecutionCache} 統一キャッシュインスタンス
 */
function getUnifiedExecutionCache() {
  if (!globalUnifiedCache) {
    globalUnifiedCache = new UnifiedExecutionCache();
    debugLog(`🏗️ 統一実行キャッシュ初期化`);
  }
  return globalUnifiedCache;
}

/**
 * 統一実行キャッシュをリセット
 */
function resetUnifiedExecutionCache() {
  globalUnifiedCache = null;
  debugLog(`🔄 統一実行キャッシュリセット`);
}

// =============================================================================
// SECTION 5: 後方互換性のための関数群（元cache.gs）
// =============================================================================

/**
 * ヘッダーインデックスをキャッシュ付きで安定取得（リトライ機能付き）
 */
function getHeadersCached(spreadsheetId, sheetName) {
  const key = `hdr_${spreadsheetId}_${sheetName}`;
  const validationKey = `hdr_validation_${spreadsheetId}_${sheetName}`;

  const indices = cacheManager.get(key, () => {
    return getHeadersWithRetry(spreadsheetId, sheetName, 3);
  }, { ttl: 1800, enableMemoization: true }); // 30分間キャッシュ + メモ化

  // 必須列の存在チェック（改善された復旧メカニズム + 重複防止）
  if (indices && Object.keys(indices).length > 0) {
    // バリデーション結果もキャッシュして重複処理を防止
    const hasRequiredColumns = cacheManager.get(validationKey, () => {
      return validateRequiredHeaders(indices);
    }, { ttl: 1800, enableMemoization: true });
    if (!hasRequiredColumns.isValid) {
      // 完全に必須列が不足している場合のみ復旧を試行
      // 片方でも存在する場合は設定によるものとして受け入れる
      if (!hasRequiredColumns.hasReasonColumn && !hasRequiredColumns.hasOpinionColumn) {
        warnLog(`[getHeadersCached] Critical headers missing: ${hasRequiredColumns.missing.join(', ')}, attempting recovery`);
        cacheManager.remove(key);

        // 即座に再取得を試行（リトライ回数を削減）
        const recoveredIndices = getHeadersWithRetry(spreadsheetId, sheetName, 1);
        if (recoveredIndices && Object.keys(recoveredIndices).length > 0) {
          const recoveredValidation = validateRequiredHeaders(recoveredIndices);
          if (recoveredValidation.hasReasonColumn || recoveredValidation.hasOpinionColumn) {
            debugLog(`[getHeadersCached] Successfully recovered headers with basic columns`);
            return recoveredIndices;
          }
        }
      } else {
        // 一部の列が存在する場合は警告のみで継続
        debugLog(`[getHeadersCached] Partial header validation - continuing with available columns: reason=${hasRequiredColumns.hasReasonColumn}, opinion=${hasRequiredColumns.hasOpinionColumn}`);
      }
    }
  }

  if (!indices || Object.keys(indices).length === 0) {
    cacheManager.remove(key);
    return {};
  }

  return indices;
}

/**
 * リトライ機能付きヘッダー取得
 * @private
 */
function getHeadersWithRetry(spreadsheetId, sheetName, maxRetries = 3) {
  let lastError = null;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      debugLog(`[getHeadersWithRetry] Attempt ${attempt}/${maxRetries} for spreadsheetId: ${spreadsheetId}, sheetName: ${sheetName}`);

      const service = getSheetsServiceCached();
      if (!service) {
        throw new Error('Sheets service is not available');
      }

      const range = sheetName + '!1:1';
      debugLog(`[getHeadersWithRetry] Fetching range: ${range}`);

      // Use the updated API pattern consistent with other functions
      const url = service.baseUrl + '/' + spreadsheetId + '/values/' + encodeURIComponent(range);
      const response = UrlFetchApp.fetch(url, {
        headers: { 'Authorization': 'Bearer ' + service.accessToken },
        muteHttpExceptions: true
      });

      // レスポンスオブジェクト検証とHTTPステータスチェック
      if (!response || typeof response.getResponseCode !== 'function') {
        throw new Error('Cache API: 無効なレスポンスオブジェクトが返されました');
      }
      
      if (response.getResponseCode() !== 200) {
        throw new Error(`HTTP ${response.getResponseCode()}: ${response.getContentText()}`);
      }

      const responseData = JSON.parse(response.getContentText());
      debugLog(`[getHeadersWithRetry] API response (attempt ${attempt}):`, JSON.stringify(responseData, null, 2));

      if (!responseData) {
        throw new Error('API response is null or undefined');
      }

      if (!responseData.values) {
        warnLog(`[getHeadersWithRetry] No values in response for ${range} (attempt ${attempt})`);
        throw new Error('No values in response');
      }

      if (!responseData.values[0] || responseData.values[0].length === 0) {
        warnLog(`[getHeadersWithRetry] Empty header row for ${range} (attempt ${attempt})`);
        throw new Error('Empty header row');
      }

      const headers = responseData.values[0];
      debugLog(`[getHeadersWithRetry] Headers found (attempt ${attempt}):`, headers);

      const indices = {};

      // タイムスタンプ列を除外してヘッダーのインデックスを生成
      headers.forEach(function(headerName, index) {
        if (headerName && headerName.trim() !== '' && headerName !== 'タイムスタンプ') {
          indices[headerName] = index;
          debugLog(`[getHeadersWithRetry] Mapped ${headerName} -> ${index}`);
        }
      });

      // 最低限のヘッダーが存在するかチェック
      if (Object.keys(indices).length === 0) {
        throw new Error('No valid headers found');
      }

      debugLog(`[getHeadersWithRetry] Final indices (attempt ${attempt}):`, indices);
      return indices;

    } catch (error) {
      lastError = error;
      errorLog(`[getHeadersWithRetry] Attempt ${attempt}/${maxRetries} failed:`, error.toString());

      // 最後の試行でない場合は待機してリトライ
      if (attempt < maxRetries) {
        const waitTime = attempt * 1000; // 1秒、2秒、3秒...
        debugLog(`[getHeadersWithRetry] Waiting ${waitTime}ms before retry...`);
        Utilities.sleep(waitTime);
      }
    }
  }

  // 全てのリトライが失敗した場合
  errorLog(`[getHeadersWithRetry] All ${maxRetries} attempts failed. Last error:`, lastError.toString());
  if (lastError.stack) {
    errorLog('Error stack:', lastError.stack);
  }

  return {};
}

/**
 * 必須ヘッダーの存在を検証（設定ベース対応）
 * @private
 */
function validateRequiredHeaders(indices) {
  // 最低限必要なヘッダーパターンをチェック
  // 理由列・回答列は設定によって名前が変わる可能性があるため、
  // より柔軟な検証を実装
  const reasonPatterns = [
    COLUMN_HEADERS.REASON,     // 理由
    '理由',
    'そう思う理由',
    'reason'
  ];
  
  const opinionPatterns = [
    COLUMN_HEADERS.OPINION,    // 回答
    '回答',
    '意見',
    '考え',
    'opinion',
    'answer'
  ];

  const missing = [];
  let hasReasonColumn = false;
  let hasOpinionColumn = false;

  // 理由列の存在チェック（いずれかのパターンが存在すればOK）
  for (const pattern of reasonPatterns) {
    if (indices[pattern] !== undefined) {
      hasReasonColumn = true;
      break;
    }
  }

  // 回答列の存在チェック（いずれかのパターンが存在すればOK）
  for (const pattern of opinionPatterns) {
    if (indices[pattern] !== undefined) {
      hasOpinionColumn = true;
      break;
    }
  }

  if (!hasReasonColumn) {
    missing.push(COLUMN_HEADERS.REASON);
  }
  if (!hasOpinionColumn) {
    missing.push(COLUMN_HEADERS.OPINION);
  }

  // 基本的なヘッダーが1つも存在しない場合は無効
  const hasBasicHeaders = Object.keys(indices).length > 0;

  return {
    isValid: hasBasicHeaders && (hasReasonColumn || hasOpinionColumn), // 最低1つは必要
    missing: missing,
    hasReasonColumn: hasReasonColumn,
    hasOpinionColumn: hasOpinionColumn
  };
}

/**
 * ユーザー関連キャッシュを無効化します。
 * @param {string} userId - ユーザーID
 * @param {string} email - 管理者メールアドレス
 * @param {string} [spreadsheetId] - 関連スプレッドシートID
 * @param {boolean} [clearPattern=false] - パターンベースのクリアを行うか
 */
function invalidateUserCache(userId, email, spreadsheetId, clearPattern, dbSpreadsheetId) {
  const keysToRemove = [];

  if (userId) {
    keysToRemove.push('user_' + userId);
    keysToRemove.push('unified_user_info_' + userId);
    // 回答ボードの表示データ・シートデータのユーザー別キャッシュを確実に無効化
    try {
      cacheManager.clearByPattern(`publishedData_${userId}_`, { strict: false, maxKeys: 200 });
      cacheManager.clearByPattern(`sheetData_${userId}_`, { strict: false, maxKeys: 200 });
      // 設定キャッシュ（シート名と紐づく）もユーザー単位でクリア
      cacheManager.clearByPattern(`config_v3_${userId}_`, { strict: false, maxKeys: 200 });
    } catch (e) {
      warnLog('invalidateUserCache: user-scoped pattern clear failed:', e.message);
    }
  }
  if (email) {
    keysToRemove.push('email_' + email);
    keysToRemove.push('unified_user_info_' + email);
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

  if (dbSpreadsheetId) {
    cacheManager.invalidateRelated('spreadsheet', dbSpreadsheetId);
  }

  debugLog(`[Cache] Invalidated user cache for userId: ${userId}, email: ${email}, spreadsheetId: ${spreadsheetId}`);
}

/**
 * クリティカル更新時の包括的キャッシュ同期化
 * データベース更新直後に使用し、すべての関連キャッシュを確実にクリア
 * @param {string} userId - ユーザーID
 * @param {string} email - メールアドレス
 * @param {string} oldSpreadsheetId - 古いスプレッドシートID
 * @param {string} newSpreadsheetId - 新しいスプレッドシートID
 */
function synchronizeCacheAfterCriticalUpdate(userId, email, oldSpreadsheetId, newSpreadsheetId) {
  try {
    infoLog('🔄 クリティカル更新後のキャッシュ同期開始...');

    // 段階1: 基本ユーザーキャッシュクリア
    invalidateUserCache(userId, email, oldSpreadsheetId, true);
    if (newSpreadsheetId && newSpreadsheetId !== oldSpreadsheetId) {
      invalidateUserCache(userId, email, newSpreadsheetId, true);
    }

    // 段階2: 実行レベルキャッシュクリア
    clearExecutionUserInfoCache();

    // 段階3: 関連データベースキャッシュクリア
    clearDatabaseCache();

    // 段階4: メモ化キャッシュの強制リセット（利用可能な場合）
    try {
      if (typeof resetMemoizationCache === 'function') {
        resetMemoizationCache();
      }
    } catch (memoError) {
      warnLog('メモ化キャッシュリセットをスキップ:', memoError.message);
    }

    infoLog('✅ クリティカル更新後のキャッシュ同期完了');

    // 少し待ってから検証用に短い待機
    Utilities.sleep(100);

  } catch (error) {
    errorLog('❌ キャッシュ同期エラー:', error);
    throw new Error('キャッシュ同期に失敗しました: ' + error.message);
  }
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
    errorLog('clearDatabaseCache error:', error.message);
    // エラーが発生しても処理を継続
  }
}

/**
 * 頻繁にアクセスされるデータを事前にキャッシュに読み込みます（プリウォーミング）
 * @param {string} activeUserEmail - 現在のユーザーのメールアドレス
 * @returns {object} プリウォーミング結果
 */
function preWarmCache(activeUserEmail) {
  const startTime = Date.now();
  const results = {
    timestamp: new Date().toISOString(),
    preWarmedItems: [],
    errors: [],
    duration: 0,
    success: true
  };

  try {
    infoLog('🔥 キャッシュプリウォーミング開始:', activeUserEmail);

    // 1. サービスアカウントトークンの事前取得
    try {
      getServiceAccountTokenCached();
      results.preWarmedItems.push('service_account_token');
      debugLog('[Cache] Pre-warmed service account token');
    } catch (error) {
      results.errors.push('service_account_token: ' + error.message);
    }

    // 2. ユーザー情報の事前取得（メール/ID両方）
    if (activeUserEmail) {
      try {
        const userInfo = findUserByEmail(activeUserEmail);
        if (userInfo) {
          results.preWarmedItems.push('user_by_email');

          // ユーザーIDベースのキャッシュも事前取得
          if (userInfo.userId) {
            findUserById(userInfo.userId);
            results.preWarmedItems.push('user_by_id');
          }

          // スプレッドシート情報の事前取得
          if (userInfo.spreadsheetId) {
            try {
              const config = JSON.parse(userInfo.configJson || '{}');
              if (config.publishedSheetName) {
                getHeadersCached(userInfo.spreadsheetId, config.publishedSheetName);
                results.preWarmedItems.push('sheet_headers');
              }
            } catch (configError) {
              results.errors.push('sheet_headers: ' + configError.message);
            }
          }
        }
        debugLog('[Cache] Pre-warmed user data for:', activeUserEmail);
      } catch (error) {
        results.errors.push('user_data: ' + error.message);
      }
    }

    // 3. システム設定の事前取得
    try {
      getWebAppUrl();
      results.preWarmedItems.push('webapp_url');
      debugLog('[Cache] Pre-warmed webapp URL');
    } catch (error) {
      results.errors.push('webapp_url: ' + error.message);
    }

    // 4. ドメイン情報の事前取得
    try {
      getDeployUserDomainInfo();
      results.preWarmedItems.push('domain_info');
      debugLog('[Cache] Pre-warmed domain info');
    } catch (error) {
      results.errors.push('domain_info: ' + error.message);
    }

    results.duration = Date.now() - startTime;
    results.success = results.errors.length === 0;

    infoLog('✅ キャッシュプリウォーミング完了:', results.preWarmedItems.length, 'items,', results.duration + 'ms');

    if (results.errors.length > 0) {
      warnLog('⚠️ プリウォーミング中のエラー:', results.errors);
    }

    return results;

  } catch (error) {
    results.duration = Date.now() - startTime;
    results.success = false;
    results.errors.push('fatal_error: ' + error.message);
    errorLog('❌ キャッシュプリウォーミングエラー:', error);
    return results;
  }
}

/**
 * キャッシュ効率化のための統計情報とチューニング推奨事項を提供
 * @returns {object} キャッシュ分析結果
 */
function analyzeCacheEfficiency() {
  try {
    const health = cacheManager.getHealth();
    const analysis = {
      timestamp: new Date().toISOString(),
      currentHealth: health,
      efficiency: 'unknown',
      recommendations: [],
      optimizationOpportunities: []
    };

    const hitRate = parseFloat(health.stats.hitRate);
    const errorRate = parseFloat(health.stats.errorRate);
    const totalOps = health.stats.totalOperations;

    // 効率レベルの判定
    if (hitRate >= 85 && errorRate < 5 && totalOps > 50) {
      analysis.efficiency = 'excellent';
    } else if (hitRate >= 70 && errorRate < 10) {
      analysis.efficiency = 'good';
    } else if (hitRate >= 50 && errorRate < 15) {
      analysis.efficiency = 'acceptable';
    } else {
      analysis.efficiency = 'poor';
    }

    // 推奨事項の生成
    if (hitRate < 70) {
      analysis.recommendations.push({
        priority: 'high',
        action: 'キャッシュヒット率向上',
        details: 'TTL設定の見直し、メモ化の活用、キャッシュキー設計の最適化を検討してください。'
      });
    }

    if (errorRate > 10) {
      analysis.recommendations.push({
        priority: 'medium',
        action: 'エラー率削減',
        details: 'キャッシュアクセスエラーの原因調査とエラーハンドリングの改善が必要です。'
      });
    }

    if (health.memoCacheSize > 1000) {
      analysis.recommendations.push({
        priority: 'low',
        action: 'メモリ使用量最適化',
        details: 'メモ化キャッシュサイズが大きくなっています。定期的なクリアを検討してください。'
      });
    }

    // 最適化機会の特定
    if (totalOps > 100 && hitRate < 80) {
      analysis.optimizationOpportunities.push('プリウォーミング戦略の導入');
    }

    if (errorRate < 5 && hitRate > 60) {
      analysis.optimizationOpportunities.push('TTL延長によるさらなる高速化');
    }

    debugLog('[Cache] Efficiency analysis completed:', analysis.efficiency);
    return analysis;

  } catch (error) {
    errorLog('analyzeCacheEfficiency error:', error);
    return {
      timestamp: new Date().toISOString(),
      error: error.message,
      efficiency: 'error',
      recommendations: [{
        priority: 'high',
        action: '分析エラー対応',
        details: 'キャッシュ分析中にエラーが発生しました。システムの状態を確認してください。'
      }]
    };
  }
}

// =============================================================================
// SECTION 6: 後方互換性のための関数群（元unifiedExecutionCache.gs）
// =============================================================================

// 後方互換性のための関数
function clearExecutionUserInfoCache() {
  const cache = getUnifiedExecutionCache();
  cache.clearUserInfo();
  cache.syncWithUnifiedCache('userDataChange');
}

function clearExecutionSheetsServiceCache() {
  const cache = getUnifiedExecutionCache();
  cache.clearSheetsService();
}

function clearAllExecutionCache() {
  const cache = getUnifiedExecutionCache();
  cache.clearAll();
  cache.syncWithUnifiedCache('systemChange');
}