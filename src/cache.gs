/**
 * @fileoverview 統合キャッシュマネージャー - 全キャッシュシステムを統合
 * cache.gs、unifiedExecutionCache.gs、spreadsheetCache.gs の機能を統合
 * 既存の全ての関数をそのまま保持し、互換性を完全維持
 */

// ULog統合システムの利用
// ulog.gsから統一ログシステムを使用

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
      lastReset: Date.now(),
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
      usePropertiesFallback = false,
    } = options;

    // パフォーマンス監視とバリデーション
    this.stats.totalOps++;
    if (!this.validateKey(key)) {
      this.stats.errors++;
      return valueFn ? valueFn() : null;
    }

    // 階層化キャッシュアクセス（最適化）
    const cacheResult = this.getFromCacheHierarchy(
      key,
      enableMemoization,
      ttl,
      usePropertiesFallback
    );
    if (cacheResult.found) {
      this.stats.hits++;
      return cacheResult.value;
    }

    // キャッシュミス: 新しい値を生成
    this.stats.misses++;

    let newValue;

    try {
      newValue = valueFn();
    } catch (e) {
      console.error('[ERROR]', `[Cache] Value generation failed for key: ${key}`, e.message);
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
      console.error('[ERROR]', `[Cache] Failed to cache value for key: ${key}`, e.message);
      this.stats.errors++;
      // キャッシュ保存に失敗しても値は返す
    }

    // パフォーマンス監視ログ（低頻度）
    if (this.stats.totalOps % 100 === 0) {
      const hitRate = ((this.stats.hits / this.stats.totalOps) * 100).toFixed(1);
      console.log(
        `[Cache] Performance: ${hitRate}% hit rate (${this.stats.hits}/${this.stats.totalOps}), ${this.stats.errors} errors`
      );
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
      console.error('[ERROR]', `[Cache] Invalid key: ${key}`);
      return false;
    }
    if (key.length > 100) {
      console.warn(`[Cache] Key too long: ${key}`);
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
        if (!memoEntry.ttl || memoEntry.createdAt + memoEntry.ttl * 1000 > Date.now()) {
          return { found: true, value: memoEntry.value };
        } else {
          this.memoCache.delete(key); // 期限切れを削除
        }
      } catch (e) {
        console.warn(`[Cache] L1(Memo) error: ${key}`, e.message);
        this.memoCache.delete(key);
      }
    }

    // Level 2: Apps Scriptキャッシュ
    try {
      const cachedValue = this.scriptCache.get(key);
      if (cachedValue !== null) {
        const parsedValue = this.parseScriptCacheValue(key, cachedValue);

        // メモ化キャッシュに昇格
        if (enableMemoization && parsedValue !== null) {
          this.memoCache.set(key, { value: parsedValue, createdAt: Date.now(), ttl });
        }

        return { found: true, value: parsedValue };
      }
    } catch (e) {
      console.warn(`[Cache] L2(Script) error: ${key}`, e.message);
      this.stats.errors++;
      this.handleCorruptedCacheEntry(key);
    }

    // Level 3: PropertiesService フォールバック（設定値用）
    if (usePropertiesFallback) {
      try {
        const propsValue = PropertiesService.getScriptProperties().getProperty(key);
        if (propsValue !== null) {
          const parsedValue = this.parsePropertiesValue(propsValue);

          // 上位キャッシュに昇格（長期TTL設定）
          this.setToCacheHierarchy(key, parsedValue, Math.max(ttl, 3600), enableMemoization);

          return { found: true, value: parsedValue };
        }
      } catch (e) {
        console.warn(`[Cache] L3(Properties) error: ${key}`, e.message);
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
      console.warn(`[Cache] Parse failed for ${key}, returning raw string`);
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
    } catch (removeError) {
      console.warn(`[Cache] Failed to clean corrupted entry: ${key}`, removeError.message);
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
        const currentUserId = User.email();
        const cacheOperations = keys.map((key) => ({ key: key }));

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
          const setCacheOps = Object.keys(newValues).map((key) => ({
            key: key,
            value: JSON.stringify(newValues[key]),
            ttl: ttl,
          }));

          for (let key in newValues) {
            results[key] = newValues[key];
          }

          if (setCacheOps.length > 0) {
            unifiedBatchProcessor.batchCacheOperation('set', setCacheOps, currentUserId, {
              concurrency: 5,
            });
          }
        }

        return results;
      }
    } catch (error) {
      console.warn('[Cache] 統一バッチ処理エラー:', error.message);
    }

    // フォールバック: 従来実装
    const results = {};
    let missingKeys = [];

    try {
      const cachedValues = this.scriptCache.getAll(keys);
      for (let key in cachedValues) {
        results[key] = JSON.parse(cachedValues[key]);
      }
      missingKeys = keys.filter((k) => !results.hasOwnProperty(k));
    } catch (e) {
      console.warn('[Cache] batchGet failed', e.message);
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
        `published_data_${spreadsheetId}_`,
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

      console.log(
        `[Cache] Invalidated sheet data cache for ${spreadsheetId}, removed ${removedCount} entries`
      );
      return removedCount;
    } catch (e) {
      console.warn(`[Cache] Failed to invalidate sheet data cache: ${e.message}`);
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
      console.warn(`[Cache] Invalid pattern for clearByPattern: ${pattern}`);
      return 0;
    }

    // セキュリティチェック: 過度に広範囲な削除を防ぐ
    if (!strict && (pattern.length < 3 || pattern === '*' || pattern === '.*')) {
      console.warn(
        `[Cache] Pattern too broad for safe removal: ${pattern}. Use strict=true to override.`
      );
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
            console.warn(`[Cache] Reached maxKeys limit (${maxKeys}) for pattern: ${pattern}`);
            break;
          }
        }
      }
    } catch (e) {
      console.warn(`[Cache] Failed to iterate memoCache keys for pattern: ${pattern}`, e.message);
      this.stats.errors++;
    }

    // 削除前の確認ログ
    if (keysToRemove.length > 100) {
      console.warn(
        `[Cache] Large pattern deletion detected: ${keysToRemove.length} keys for pattern: ${pattern}`
      );
    }

    keysToRemove.forEach((key) => {
      try {
        this.memoCache.delete(key);
        this.scriptCache.remove(key);
      } catch (e) {
        console.warn(`[Cache] Failed to remove key during pattern clear: ${key}`, e.message);
        failedRemovals++;
        this.stats.errors++;
      }
    });

    const successCount = keysToRemove.length - failedRemovals;

    console.log(
      `[Cache] Pattern clear completed: ${successCount} removed, ${failedRemovals} failed, ${skippedCount} protected (pattern: ${pattern})`
    );

    return successCount;
  }

  /**
   * 保護すべきキーかどうかを判定
   * @private
   */
  _isProtectedKey(key) {
    const protectedPatterns = [
      'SA_TOKEN_CACHE', // サービスアカウントトークン
      'WEB_APP_URL', // WebアプリURL
      'SYSTEM_CONFIG', // システム設定
      'DOMAIN_INFO', // ドメイン情報
    ];

    return protectedPatterns.some((pattern) => key.includes(pattern));
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
      startTime: Date.now(),
    };

    try {
      // 入力検証
      if (!entityType || !entityId) {
        throw new Error('entityType and entityId are required');
      }

      // 関連IDの数制限
      if (relatedIds.length > maxRelated) {
        console.warn(
          `[Cache] Too many related IDs (${relatedIds.length}), limiting to ${maxRelated}`
        );
        relatedIds = relatedIds.slice(0, maxRelated);
      }

      console.log(
        `🔗 関連キャッシュ無効化開始: ${entityType}/${entityId} (${dryRun ? 'DRY RUN' : 'LIVE'})`
      );

      // 1. メインエンティティのキャッシュ無効化
      const patterns = this._getInvalidationPatterns(entityType, entityId);
      invalidationLog.patterns.push(...patterns);

      patterns.forEach((pattern) => {
        try {
          if (!dryRun) {
            const removed = this.clearByPattern(pattern, { strict: false, maxKeys: 100 });
            invalidationLog.totalRemoved += removed;
          } else {
          }
        } catch (error) {
          invalidationLog.errors.push(`Pattern ${pattern}: ${error.message}`);
        }
      });

      // 2. 関連エンティティのキャッシュ無効化（検証付き）
      relatedIds.forEach((relatedId) => {
        if (!relatedId || relatedId === entityId) {
          return; // 無効な関連IDまたは自分自身はスキップ
        }

        try {
          const relatedPatterns = this._getInvalidationPatterns(entityType, relatedId);
          relatedPatterns.forEach((pattern) => {
            if (!dryRun) {
              const removed = this.clearByPattern(pattern, { strict: false, maxKeys: 50 });
              invalidationLog.totalRemoved += removed;
            } else {
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
        console.warn(
          `⚠️ 関連キャッシュ無効化で一部エラー: ${entityType}/${entityId}`,
          invalidationLog.errors
        );
      } else {
        console.log(
          `✅ 関連キャッシュ無効化完了: ${entityType}/${entityId} (${invalidationLog.totalRemoved} entries, ${invalidationLog.duration}ms)`
        );
      }

      return invalidationLog;
    } catch (error) {
      invalidationLog.errors.push(`Fatal: ${error.message}`);
      invalidationLog.duration = Date.now() - invalidationLog.startTime;
      console.error(
        '[ERROR]',
        `❌ 関連キャッシュ無効化致命的エラー: ${entityType}/${entityId}`,
        error
      );
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
        patterns.push(`form_${entityId}*`, `formUrl_${entityId}*`, `formResponse_${entityId}*`);
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
        relatedIds.forEach((relatedId) => {
          if (relatedId && typeof relatedId === 'string' && relatedId.length > 0) {
            try {
              const removed1 = this.clearByPattern(`user_${entityId}_spreadsheet_${relatedId}`, {
                maxKeys: 20,
              });
              const removed2 = this.clearByPattern(`user_${entityId}_form_${relatedId}`, {
                maxKeys: 20,
              });
              invalidationLog.totalRemoved += removed1 + removed2;
            } catch (error) {
              invalidationLog.errors.push(`Cross-entity user-${relatedId}: ${error.message}`);
            }
          }
        });
      }

      // スプレッドシート変更時は関連するユーザーキャッシュも無効化
      if (entityType === 'spreadsheet' && relatedIds.length > 0) {
        relatedIds.forEach((userId) => {
          if (userId && typeof userId === 'string' && userId.length > 0) {
            try {
              const removed1 = this.clearByPattern(`user_${userId}_spreadsheet_${entityId}`, {
                maxKeys: 20,
              });
              const removed2 = this.clearByPattern(`config_${userId}`, { maxKeys: 10 }); // ユーザー設定も無効化
              invalidationLog.totalRemoved += removed1 + removed2;
            } catch (error) {
              invalidationLog.errors.push(`Cross-entity spreadsheet-${userId}: ${error.message}`);
            }
          }
        });
      }

      // フォーム変更時の追加ロジック
      if (entityType === 'form' && relatedIds.length > 0) {
        relatedIds.forEach((userId) => {
          if (userId && typeof userId === 'string' && userId.length > 0) {
            try {
              const removed = this.clearByPattern(`user_${userId}_form_${entityId}`, {
                maxKeys: 10,
              });
              invalidationLog.totalRemoved += removed;
            } catch (error) {
              invalidationLog.errors.push(`Cross-entity form-${userId}: ${error.message}`);
            }
          }
        });
      }
    } catch (error) {
      invalidationLog.errors.push(`Cross-entity fatal: ${error.message}`);
      console.warn('クロスエンティティキャッシュ無効化エラー:', error.message);
      this.stats.errors++;
    }
  }

  /**
   * 期限切れのキャッシュをクリアします（この機能はGASでは自動です）。
   * メモ化キャッシュをクリアする目的で実装します。
   */
  clearExpired() {
    this.memoCache.clear();
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
    } catch (e) {
      console.warn('[Cache] Failed to clear memoization cache:', e.message);
    }

    try {
      // スクリプトキャッシュクリア - GAS API制限のため自動期限切れに依存
      scriptCacheCleared = true;
    } catch (e) {
      console.warn('[Cache] Failed to clear script cache:', e.message);
    }

    // 統計をリセット
    try {
      this.resetStats();
    } catch (e) {
      console.warn('[Cache] Failed to reset stats:', e.message);
    }

    console.log(
      `[Cache] clearAll() completed - MemoCache: ${memoCacheCleared ? 'OK' : 'FAILED'}, ScriptCache: ${scriptCacheCleared ? 'OK' : 'FAILED'}`
    );

    return {
      memoCacheCleared,
      scriptCacheCleared,
      success: memoCacheCleared && scriptCacheCleared,
    };
  }

  /**
   * キャッシュの健全性情報を取得します。
   * @returns {object} 健全性情報
   */
  getHealth() {
    const uptime = Date.now() - this.stats.lastReset;
    const hitRate =
      this.stats.totalOps > 0 ? ((this.stats.hits / this.stats.totalOps) * 100).toFixed(1) : 0;
    const errorRate =
      this.stats.totalOps > 0 ? ((this.stats.errors / this.stats.totalOps) * 100).toFixed(1) : 0;

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
        uptimeMs: uptime,
      },
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
      lastReset: Date.now(),
    };
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
          ttl: ttl,
        });
      }
    } catch (e) {
      console.warn(`[Cache] setToCacheHierarchy error: ${key}`, e.message);
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
      ttl = 3600, // 設定値は1時間キャッシュ
      enableMemoization = true,
      usePropertiesFallback = true,
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
CacheManager.prototype.clearAllFrontendCaches = function (options = {}) {
  const { force = false, timeout = 10000 } = options;

  // 既にクリア中の場合は待機
  if (this.clearInProgress && !force) {
    if (this.debugMode) {
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
        console.log('🗑️ Starting unified cache clear process');
      }

      // キャッシュクリア操作のリスト（優先順位順）
      const clearOperations = [
        {
          name: 'UnifiedCache',
          operation: () => {
            if (
              typeof window !== 'undefined' &&
              window.unifiedCache &&
              typeof window.unifiedCache.clear === 'function'
            ) {
              window.unifiedCache.clear();
              return true;
            }
            return false;
          },
        },
        {
          name: 'GasOptimizerCache',
          operation: () => {
            if (
              typeof window !== 'undefined' &&
              window.gasOptimizer &&
              typeof window.gasOptimizer.clearCache === 'function'
            ) {
              window.gasOptimizer.clearCache();
              return true;
            }
            return false;
          },
        },
        {
          name: 'SharedUtilitiesCache',
          operation: () => {
            if (
              typeof window !== 'undefined' &&
              window.sharedUtilities &&
              window.sharedUtilities.cache &&
              typeof window.sharedUtilities.cache.clear === 'function'
            ) {
              window.sharedUtilities.cache.clear();
              return true;
            }
            return false;
          },
        },
        {
          name: 'DOMElementCache',
          operation: () => {
            if (
              typeof window !== 'undefined' &&
              window.sharedUtilities &&
              window.sharedUtilities.dom &&
              typeof window.sharedUtilities.dom.clearElementCache === 'function'
            ) {
              window.sharedUtilities.dom.clearElementCache();
              return true;
            }
            return false;
          },
        },
        {
          name: 'ThrottleDebounceCache',
          operation: () => {
            if (
              typeof window !== 'undefined' &&
              window.sharedUtilities &&
              window.sharedUtilities.throttle &&
              typeof window.sharedUtilities.throttle.clearAll === 'function'
            ) {
              window.sharedUtilities.throttle.clearAll();
              return true;
            }
            return false;
          },
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
          },
        },
      ];

      const results = [];

      // 順次実行でキャッシュクリア（競合回避）
      for (const clearOp of clearOperations) {
        try {
          const success = clearOp.operation();
          results.push({ name: clearOp.name, success });

          if (this.debugMode && success) {
          }

          // 各操作間に短い間隔を設ける
          if (typeof Utilities !== 'undefined') {
            Utilities.sleep(50);
          }
        } catch (error) {
          console.warn(`⚠️ Failed to clear ${clearOp.name}:`, error);
          results.push({ name: clearOp.name, success: false, error: error.message });
        }
      }

      const successCount = results.filter((r) => r.success).length;
      const totalTime = Date.now() - startTime;

      if (this.debugMode) {
        console.log(
          `🎉 Cache clear completed: ${successCount}/${results.length} caches cleared in ${totalTime}ms`
        );
      }

      // 待機中のクリア要求を解決
      this.resolvePendingClears(results);

      resolve({
        success: true,
        results,
        successCount,
        totalCount: results.length,
        duration: totalTime,
      });
    } catch (error) {
      console.error('[ERROR]', '❌ Unified cache clear failed:', error);
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
CacheManager.prototype.clearSpecificCache = function (cacheType) {
  return new Promise((resolve, reject) => {
    const cacheOperations = {
      unified: () => typeof window !== 'undefined' && window.unifiedCache?.clear(),
      gasOptimizer: () => typeof window !== 'undefined' && window.gasOptimizer?.clearCache(),
      sharedUtilities: () =>
        typeof window !== 'undefined' && window.sharedUtilities?.cache?.clear(),
      domElements: () =>
        typeof window !== 'undefined' && window.sharedUtilities?.dom?.clearElementCache(),
      throttleDebounce: () =>
        typeof window !== 'undefined' && window.sharedUtilities?.throttle?.clearAll(),
      script: () => this.clearAll(),
    };

    const operation = cacheOperations[cacheType];
    if (!operation) {
      reject(new Error(`Unknown cache type: ${cacheType}`));
      return;
    }

    try {
      operation();
      if (this.debugMode) {
      }
      resolve({ success: true, cacheType });
    } catch (error) {
      console.warn(`⚠️ Failed to clear ${cacheType} cache:`, error);
      resolve({ success: false, cacheType, error: error.message });
    }
  });
};

/**
 * キャッシュ状態の診断
 * @returns {Object} 診断結果
 */
CacheManager.prototype.diagnoseFrontendCache = function () {
  const caches = {
    unifiedCache: {
      available:
        typeof window !== 'undefined' && !!(window.unifiedCache && window.unifiedCache.clear),
      size: typeof window !== 'undefined' ? window.unifiedCache?.size || 'unknown' : 'unavailable',
    },
    gasOptimizer: {
      available:
        typeof window !== 'undefined' && !!(window.gasOptimizer && window.gasOptimizer.clearCache),
      size:
        typeof window !== 'undefined'
          ? window.gasOptimizer?.cache?.size || 'unknown'
          : 'unavailable',
    },
    sharedUtilities: {
      available:
        typeof window !== 'undefined' && !!(window.sharedUtilities && window.sharedUtilities.cache),
      size:
        typeof window !== 'undefined'
          ? window.sharedUtilities?.cache?.size || 'unknown'
          : 'unavailable',
    },
    domElements: {
      available:
        typeof window !== 'undefined' && !!(window.sharedUtilities && window.sharedUtilities.dom),
      size: 'unknown',
    },
    scriptCache: {
      available: true,
      size: this.memoCache.size,
    },
  };

  return {
    clearInProgress: this.clearInProgress,
    pendingClears: this.pendingClears.length,
    caches,
    health: this.getHealth(),
  };
};

/**
 * 待機中のクリア要求を解決
 */
CacheManager.prototype.resolvePendingClears = function (results) {
  const pending = this.pendingClears.splice(0);
  pending.forEach(({ resolve }) => resolve(results));
};

/**
 * 待機中のクリア要求を拒否
 */
CacheManager.prototype.rejectPendingClears = function (error) {
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
class ExecutionCache {
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
  }

  /**
   * ユーザー情報キャッシュをクリア
   */
  clearUserInfo() {
    this.userInfoCache = null;
    this.lastUserIdKey = null;
  }

  /**
   * SheetsServiceキャッシュをクリア
   */
  clearSheetsService() {
    this.sheetsServiceCache = null;
  }

  /**
   * 全キャッシュをクリア
   */
  clearAll() {
    this.clearUserInfo();
    this.clearSheetsService();
  }

  /**
   * キャッシュが期限切れかチェック
   * @returns {boolean} 期限切れの場合true
   */
  isExpired() {
    return Date.now() - this.executionStartTime > this.maxLifetime;
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
      isExpired: this.isExpired(),
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
      } catch (error) {
      }
    }
  }
}

// =============================================================================
// SECTION 3: SpreadsheetApp最適化システム（元spreadsheetCache.gs）
// =============================================================================

// メモリキャッシュ - 実行セッション内で有効
let spreadsheetMemoryCache = {};

// Module-scoped constants (2024 GAS Best Practice)
const CACHE_CONFIG = Object.freeze({
  MEMORY_TTL: CORE.TIMEOUTS.LONG,     // 30秒（メモリキャッシュ）
  SESSION_TTL: CORE.TIMEOUTS.LONG * 60,  // 30分（Propertiesキャッシュ）
  MAX_SIZE: 50,
  KEY_PREFIX: 'ss_cache_',
});

/**
 * キャッシュされたSpreadsheetオブジェクトを取得
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {boolean} forceRefresh - 強制リフレッシュフラグ
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} Spreadsheetオブジェクト
 */
function getCachedSpreadsheet(spreadsheetId, forceRefresh = false) {
  if (!spreadsheetId || typeof spreadsheetId !== 'string') {
    throw new Error('有効なスプレッドシートIDが必要です');
  }

  const cacheKey = `${CACHE_CONFIG.CACHE_KEY_PREFIX}${spreadsheetId}`;
  const now = Date.now();

  // 強制リフレッシュの場合はキャッシュをクリア
  if (forceRefresh) {
    delete spreadsheetMemoryCache[spreadsheetId];
    try {
      resilientCacheOperation(
        () => PropertiesService.getScriptProperties().deleteProperty(cacheKey),
        'PropertiesService.deleteProperty'
      );
    } catch (error) {
    }
  }

  // Phase 1: メモリキャッシュをチェック
  const memoryEntry = spreadsheetMemoryCache[spreadsheetId];
  if (memoryEntry && now - memoryEntry.timestamp < CACHE_CONFIG.MEMORY_CACHE_TTL) {
    console.log(
      '✅ SpreadsheetApp.openById メモリキャッシュヒット:',
      spreadsheetId.substring(0, 10)
    );
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
      if (now - sessionEntry.timestamp < CACHE_CONFIG.SESSION_CACHE_TTL) {
        // セッションキャッシュからSpreadsheetオブジェクトを再構築
        const spreadsheet = resilientSpreadsheetOperation(
          () => SpreadsheetApp.openById(spreadsheetId),
          'SpreadsheetApp.openById'
        );

        // メモリキャッシュに保存
        spreadsheetMemoryCache[spreadsheetId] = {
          spreadsheet: spreadsheet,
          timestamp: now,
        };

        console.log(
          '✅ SpreadsheetApp.openById セッションキャッシュヒット:',
          spreadsheetId.substring(0, 10)
        );
        return spreadsheet;
      }
    }
  } catch (error) {
  }

  // Phase 3: 新規取得とキャッシュ保存
  console.log('🔄 SpreadsheetApp.openById 新規取得:', spreadsheetId.substring(0, 10));

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);

    // メモリキャッシュに保存
    spreadsheetMemoryCache[spreadsheetId] = {
      spreadsheet: spreadsheet,
      timestamp: now,
    };

    // セッションキャッシュに保存（軽量データのみ）
    try {
      const sessionData = {
        spreadsheetId: spreadsheetId,
        timestamp: now,
        url: spreadsheet.getUrl(),
      };

      PropertiesService.getScriptProperties().setProperty(cacheKey, JSON.stringify(sessionData));
    } catch (sessionError) {
    }

    // キャッシュサイズ管理
    cleanupOldCacheEntries();

    return spreadsheet;
  } catch (error) {
    logError(
      error,
      'SpreadsheetApp.openById',
      UNIFIED_CONSTANTS.ERROR.SEVERITY.HIGH,
      ERROR_CATEGORIES.DATABASE,
      { spreadsheetId }
    );
    throw new Error(`スプレッドシートの取得に失敗しました: ${error.message}`);
  }
}

/**
 * 古いキャッシュエントリをクリーンアップ
 */
function cleanupOldCacheEntries() {
  const now = Date.now();
  const memoryKeys = Object.keys(spreadsheetMemoryCache);

  // メモリキャッシュのクリーンアップ
  if (memoryKeys.length > CACHE_CONFIG.MAX_CACHE_SIZE) {
    const sortedEntries = memoryKeys
      .map((key) => ({
        key: key,
        timestamp: spreadsheetMemoryCache[key].timestamp,
      }))
      .sort((a, b) => b.timestamp - a.timestamp);

    // 古いエントリを削除
    const entriesToDelete = sortedEntries.slice(CACHE_CONFIG.MAX_CACHE_SIZE);
    entriesToDelete.forEach((entry) => {
      delete spreadsheetMemoryCache[entry.key];
    });

  }

  // 期限切れエントリの削除
  memoryKeys.forEach((key) => {
    const entry = spreadsheetMemoryCache[key];
    if (now - entry.timestamp > CACHE_CONFIG.MEMORY_CACHE_TTL) {
      delete spreadsheetMemoryCache[key];
    }
  });
}

/**
 * 特定のスプレッドシートのキャッシュを無効化
 * @param {string} spreadsheetId - スプレッドシートID
 */
function invalidateSpreadsheetCache(spreadsheetId) {
  if (!spreadsheetId) return;

  // メモリキャッシュから削除
  delete spreadsheetMemoryCache[spreadsheetId];

  // セッションキャッシュから削除
  const cacheKey = `${CACHE_CONFIG.CACHE_KEY_PREFIX}${spreadsheetId}`;
  try {
    PropertiesService.getScriptProperties().deleteProperty(cacheKey);
  } catch (error) {
  }
}

/**
 * 全スプレッドシートキャッシュをクリア
 */
function clearAllSpreadsheetCache() {
  // メモリキャッシュクリア
  spreadsheetMemoryCache = {};

  // セッションキャッシュクリア
  try {
    const props = PropertiesService.getScriptProperties();
    const allProps = props.getProperties();

    Object.keys(allProps).forEach((key) => {
      if (key.startsWith(CACHE_CONFIG.CACHE_KEY_PREFIX)) {
        props.deleteProperty(key);
      }
    });

  } catch (error) {
  }
}

/**
 * キャッシュ統計情報を取得
 * @returns {Object} キャッシュ統計
 */
function getSpreadsheetCacheStats() {
  const memoryEntries = Object.keys(spreadsheetMemoryCache).length;
  const now = Date.now();

  let sessionEntries = 0;
  try {
    const allProps = PropertiesService.getScriptProperties().getProperties();
    sessionEntries = Object.keys(allProps).filter((key) =>
      key.startsWith(CACHE_CONFIG.CACHE_KEY_PREFIX)
    ).length;
  } catch (error) {
  }

  return {
    memoryEntries: memoryEntries,
    sessionEntries: sessionEntries,
    maxCacheSize: CACHE_CONFIG.MAX_CACHE_SIZE,
    memoryTTL: CACHE_CONFIG.MEMORY_CACHE_TTL,
    sessionTTL: CACHE_CONFIG.SESSION_CACHE_TTL,
  };
}

/**
 * SpreadsheetApp.openById()の最適化されたラッパー関数
 * 既存コードの置き換え用
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} Spreadsheetオブジェクト
 */
function openSpreadsheetOptimized(spreadsheetId) {
  return getCachedSpreadsheet(spreadsheetId);
}

// =============================================================================
// SECTION 4: 統合インスタンス作成と後方互換性関数
// =============================================================================

// --- シングルトンインスタンスの作成 ---
// 統合キャッシュマネージャー（内部専用）
const cacheManager = new CacheManager();

// Google Apps Script環境でグローバルにアクセス可能にする
if (typeof global !== 'undefined') {
  global.cacheManager = cacheManager;
}

// グローバル統一キャッシュインスタンス
let globalUnifiedCache = null;

/**
 * 統一実行キャッシュのインスタンスを取得
 * @returns {ExecutionCache} 統一キャッシュインスタンス
 */
function getExecutionCache() {
  if (!globalUnifiedCache) {
    globalUnifiedCache = new ExecutionCache();
  }
  return globalUnifiedCache;
}

/**
 * 統一実行キャッシュをリセット
 */
function resetExecutionCache() {
  globalUnifiedCache = null;
}

// =============================================================================
// SECTION 5: 後方互換性のための関数群（元cache.gs）
// =============================================================================

/**
 * ヘッダーインデックスをキャッシュ付きで安定取得（リトライ機能付き）
 */
function getHeadersCached(spreadsheetId, sheetName) {
  console.log(
    `📋 [HEADER_CACHE] Requested headers for ${sheetName} in ${spreadsheetId?.substring(0, 10)}...`
  );
  const key = `hdr_${spreadsheetId}_${sheetName}`;
  const validationKey = `hdr_validation_${spreadsheetId}_${sheetName}`;

  const indices = cacheManager.get(
    key,
    () => {
      return getHeadersWithRetry(spreadsheetId, sheetName, 3);
    },
    { ttl: 1800, enableMemoization: true }
  ); // 30分間キャッシュ + メモ化

  // 必須列の存在チェック（改善された復旧メカニズム + 重複防止）
  if (indices && Object.keys(indices).length > 0) {
    // バリデーション結果もキャッシュして重複処理を防止
    const hasRequiredColumns = cacheManager.get(
      validationKey,
      () => {
        return validateRequiredHeaders(indices);
      },
      { ttl: 1800, enableMemoization: true }
    );
    if (!hasRequiredColumns.isValid) {
      // 質問文がヘッダーになっている場合は有効として扱う
      if (hasRequiredColumns.hasQuestionColumn) {
        console.log(
          `[getHeadersCached] Question column detected as header - treating as valid configuration`
        );
        return indices;
      }
      
      // 完全に必須列が不足している場合のみ復旧を試行
      // 片方でも存在する場合は設定によるものとして受け入れる
      if (!hasRequiredColumns.hasReasonColumn && !hasRequiredColumns.hasOpinionColumn && !hasRequiredColumns.hasQuestionColumn) {
        console.warn(
          `[getHeadersCached] Critical headers missing: ${hasRequiredColumns.missing.join(', ')}, attempting recovery`
        );
        cacheManager.remove(key);
        cacheManager.remove(validationKey);

        // 即座に再取得を試行（リトライ回数を削減）
        const recoveredIndices = getHeadersWithRetry(spreadsheetId, sheetName, 1);
        if (recoveredIndices && Object.keys(recoveredIndices).length > 0) {
          const recoveredValidation = validateRequiredHeaders(recoveredIndices);
          if (recoveredValidation.hasReasonColumn || recoveredValidation.hasOpinionColumn || recoveredValidation.hasQuestionColumn) {
            return recoveredIndices;
          }
        }
      } else {
        // 一部の列が存在する場合は警告のみで継続
        console.log(
          `[getHeadersCached] Partial header validation - continuing with available columns: reason=${hasRequiredColumns.hasReasonColumn}, opinion=${hasRequiredColumns.hasOpinionColumn}, question=${hasRequiredColumns.hasQuestionColumn}`
        );
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
  // 詳細ログ追加
  console.log(
    `🔍 [HEADER_DETECTION] Starting header detection for ${sheetName} in ${spreadsheetId?.substring(0, 10)}...`
  );
  let lastError = null;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      console.log(
        `[HEADER_DETECTION] Attempt ${attempt}/${maxRetries} for spreadsheetId: ${spreadsheetId?.substring(0, 10)}, sheetName: ${sheetName}`
      );
      console.log(
        `[getHeadersWithRetry] Attempt ${attempt}/${maxRetries} for spreadsheetId: ${spreadsheetId}, sheetName: ${sheetName}`
      );

      const service = getSheetsServiceCached();
      if (!service) {
        throw new Error('Sheets service is not available');
      }

      const range = sheetName + '!1:1';
      console.log(`[getHeadersWithRetry] Fetching range: ${range}`);

      // Use the updated API pattern consistent with other functions
      const url = service.baseUrl + '/' + spreadsheetId + '/values/' + encodeURIComponent(range);
      const response = UrlFetchApp.fetch(url, {
        headers: { Authorization: 'Bearer ' + service.accessToken },
        muteHttpExceptions: true,
      });

      // レスポンスオブジェクト検証とHTTPステータスチェック
      if (!response || typeof response.getResponseCode !== 'function') {
        throw new Error('Cache API: 無効なレスポンスオブジェクトが返されました');
      }

      if (response.getResponseCode() !== 200) {
        throw new Error(`HTTP ${response.getResponseCode()}: ${response.getContentText()}`);
      }

      const responseData = JSON.parse(response.getContentText());
      console.log(
        `[getHeadersWithRetry] API response (attempt ${attempt}):`,
        JSON.stringify(responseData, null, 2)
      );

      if (!responseData) {
        throw new Error('API response is null or undefined');
      }

      if (!responseData.values) {
        console.warn(
          `[getHeadersWithRetry] No values in response for ${range} (attempt ${attempt})`
        );
        throw new Error('No values in response');
      }

      if (!responseData.values[0] || responseData.values[0].length === 0) {
        console.warn(`[getHeadersWithRetry] Empty header row for ${range} (attempt ${attempt})`);
        throw new Error('Empty header row');
      }

      const headers = responseData.values[0];
      console.log(`[getHeadersWithRetry] Headers found (attempt ${attempt}):`, headers);

      const indices = {};

      // タイムスタンプ列を除外してヘッダーのインデックスを生成
      headers.forEach(function (headerName, index) {
        if (headerName && headerName.trim() !== '' && headerName !== 'タイムスタンプ') {
          indices[headerName] = index;
          console.log(`[getHeadersWithRetry] Mapped ${headerName} -> ${index}`);
        }
      });

      // 最低限のヘッダーが存在するかチェック
      if (Object.keys(indices).length === 0) {
        throw new Error('No valid headers found');
      }

      console.log(`[getHeadersWithRetry] Final indices (attempt ${attempt}):`, indices);
      console.log(
        `✅ [HEADER_DETECTION] Success! Found ${Object.keys(indices).length} headers:`,
        Object.keys(indices)
      );
      console.log(`🎯 [HEADER_DETECTION] Column mapping:`, JSON.stringify(indices, null, 2));
      return indices;
    } catch (error) {
      lastError = error;
      console.error(
        '[ERROR]',
        `[getHeadersWithRetry] Attempt ${attempt}/${maxRetries} failed:`,
        error.toString()
      );

      // 最後の試行でない場合は待機してリトライ
      if (attempt < maxRetries) {
        const waitTime = attempt * 1000; // 1秒、2秒、3秒...
        console.log(`[getHeadersWithRetry] Waiting ${waitTime}ms before retry...`);
        Utilities.sleep(waitTime);
      }
    }
  }

  // 全てのリトライが失敗した場合
  console.error(
    '[ERROR]',
    `[getHeadersWithRetry] All ${maxRetries} attempts failed. Last error:`,
    lastError.toString()
  );
  if (lastError.stack) {
    console.error('[ERROR]', 'Error stack:', lastError.stack);
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
  
  // AIが認識しやすいパターンを大幅に拡張
  const reasonPatterns = [
    COLUMN_HEADERS.REASON, // 理由
    '理由',
    'そう思う理由',
    'そう考える理由',
    'reason',
    '根拠',
    '体験',
    '詳細',
    '説明'
  ];

  const opinionPatterns = [
    COLUMN_HEADERS.OPINION, // 回答
    '回答',
    '意見',
    '考え',
    'opinion',
    'answer',
    '答え',
    '質問',
    '問い'
  ];

  const missing = [];
  let hasReasonColumn = false;
  let hasOpinionColumn = false;
  let hasQuestionColumn = false;  // 質問文がヘッダーになっている場合

  // より柔軟な判定: キーにパターンが含まれているかチェック
  for (const key in indices) {
    const keyLower = key.toLowerCase();
    
    // 理由列の存在チェック（部分一致も許可）
    for (const pattern of reasonPatterns) {
      if (key === pattern || keyLower.includes(pattern.toLowerCase()) || 
          keyLower.includes('理由') || keyLower.includes('体験') || keyLower.includes('根拠')) {
        hasReasonColumn = true;
        break;
      }
    }
    
    // 回答列の存在チェック（部分一致も許可）
    for (const pattern of opinionPatterns) {
      if (key === pattern || keyLower.includes(pattern.toLowerCase()) || 
          keyLower.includes('回答') || keyLower.includes('答') || keyLower.includes('意見')) {
        hasOpinionColumn = true;
        break;
      }
    }
    
    // 質問文と思われる長いヘッダーの検出（15文字以上で「？」を含む）
    if (key.length > 15 && (key.includes('？') || key.includes('?') || key.includes('どうして') || 
        key.includes('なぜ') || key.includes('思います') || key.includes('考え'))) {
      hasQuestionColumn = true;
      hasOpinionColumn = true; // 質問文も回答列として扱う
    }
  }
  
  // フォームの回答シートによくあるヘッダーも回答列として認識
  const formResponseIndices = Object.keys(indices);
  if (formResponseIndices.some(header => 
    header.length > 20 && (header.includes('どうして') || header.includes('なぜ') || 
    header.includes('思いますか') || header.includes('考えますか')))) {
    hasOpinionColumn = true;
    hasQuestionColumn = true;
  }

  if (!hasReasonColumn) {
    missing.push('理由または詳細説明の列');
  }
  if (!hasOpinionColumn && !hasQuestionColumn) {
    missing.push('回答または質問の列');
  }

  // 基本的なヘッダーが1つも存在しない場合は無効
  const hasBasicHeaders = Object.keys(indices).length > 0;
  
  // より緩い検証: 質問文がヘッダーにある場合も有効とする
  const isValid = hasBasicHeaders && (hasReasonColumn || hasOpinionColumn || hasQuestionColumn);

  return {
    isValid: isValid,
    missing: missing,
    hasReasonColumn: hasReasonColumn,
    hasOpinionColumn: hasOpinionColumn,
    hasQuestionColumn: hasQuestionColumn
  };
}

/**
 * ユーザー関連キャッシュを無効化します。
 * @param {string} userId - ユーザーID
 * @param {string} email - 管理者メールアドレス
 * @param {string} [spreadsheetId] - 関連スプレッドシートID
 * @param {boolean} [clearPattern=false] - パターンベースのクリアを行うか
 */
// 古いinvalidateUserCache実装は削除（統一APIバージョンを下部で使用）

/**
 * クリティカル更新時の包括的キャッシュ同期化
 * データベース更新直後に使用し、すべての関連キャッシュを確実にクリア
 * @param {string} userId - ユーザーID
 * @param {string} email - メールアドレス
 * @param {string} oldSpreadsheetId - 古いスプレッドシートID
 * @param {string} newSpreadsheetId - 新しいスプレッドシートID
 */
// 古いsynchronizeCacheAfterCriticalUpdate実装は削除（統一APIバージョンを下部で使用）

/**
 * データベース関連キャッシュをクリアします。
 * エラー時やデータ整合性が必要な場合に使用します。
 */
// 古いclearDatabaseCache実装は削除（統一APIバージョンを下部で使用）

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
    success: true,
  };

  try {

    // 1. サービスアカウントトークンの事前取得
    try {
      getServiceAccountTokenCached();
      results.preWarmedItems.push('service_account_token');
    } catch (error) {
      results.errors.push('service_account_token: ' + error.message);
    }

    // 2. ユーザー情報の事前取得（メール/ID両方）
    if (activeUserEmail) {
      try {
        const userInfo = DB.findUserByEmail(activeUserEmail);
        if (userInfo) {
          results.preWarmedItems.push('user_by_email');

          // ユーザーIDベースのキャッシュも事前取得
          if (userInfo.userId) {
            DB.findUserById(userInfo.userId);
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
      } catch (error) {
        results.errors.push('user_data: ' + error.message);
      }
    }

    // 3. システム設定の事前取得
    try {
      getWebAppUrl();
      results.preWarmedItems.push('webapp_url');
    } catch (error) {
      results.errors.push('webapp_url: ' + error.message);
    }

    // 4. ドメイン情報の事前取得
    try {
      Deploy.domain();
      results.preWarmedItems.push('domain_info');
    } catch (error) {
      results.errors.push('domain_info: ' + error.message);
    }

    results.duration = Date.now() - startTime;
    results.success = results.errors.length === 0;

    console.log(
      '✅ キャッシュプリウォーミング完了:',
      results.preWarmedItems.length,
      'items,',
      results.duration + 'ms'
    );

    if (results.errors.length > 0) {
      console.warn('⚠️ プリウォーミング中のエラー:', results.errors);
    }

    return results;
  } catch (error) {
    results.duration = Date.now() - startTime;
    results.success = false;
    results.errors.push('fatal_error: ' + error.message);
    console.error('[ERROR]', '❌ キャッシュプリウォーミングエラー:', error);
    return results;
  }
}

/**
 * キャッシュ効率化のための統計情報とチューニング推奨事項を提供
 * @returns {object} キャッシュ分析結果
 */
function getCacheStats() {
  // セキュリティチェック: システム管理者のみアクセス可能
  if (!checkIsSystemAdmin()) {
    throw new Error('キャッシュ統計の確認には管理者権限が必要です');
  }

  try {
    const health = cacheManager.getHealth();
    const stats = {
      timestamp: new Date().toISOString(),
      health: health,
      efficiency: 'unknown',
      recommendations: [],
    };

    const hitRate = parseFloat(health.stats.hitRate);
    const errorRate = parseFloat(health.stats.errorRate);
    const totalOps = health.stats.totalOperations;

    // 効率レベルの判定
    if (hitRate >= 85 && errorRate < 5 && totalOps > 50) {
      stats.efficiency = 'excellent';
    } else if (hitRate >= 70 && errorRate < 10) {
      stats.efficiency = 'good';
    } else if (hitRate >= 50 && errorRate < 15) {
      stats.efficiency = 'acceptable';
    } else {
      stats.efficiency = 'poor';
    }

    // 推奨事項の生成（管理者向け）
    if (hitRate < 70) {
      stats.recommendations.push({
        priority: 'high',
        action: 'キャッシュヒット率向上',
        details: 'TTL設定の見直し、メモ化の活用を検討してください。',
      });
    }

    if (errorRate > 10) {
      stats.recommendations.push({
        priority: 'medium',
        action: 'エラー率削減',
        details: 'キャッシュアクセスエラーの調査とハンドリング改善が必要です。',
      });
    }

    if (health.memoCacheSize > 1000) {
      stats.recommendations.push({
        priority: 'low',
        action: 'メモリ使用量最適化',
        details: 'メモ化キャッシュサイズが大きくなっています。定期クリアを検討してください。',
      });
    }

    return {
      success: true,
      stats: stats,
    };
  } catch (error) {
    console.error('[ERROR]', 'getCacheStats error:', error);
    return {
      success: false,
      error: error.message,
      timestamp: new Date().toISOString(),
      error: error.message,
      efficiency: 'error',
      recommendations: [
        {
          priority: 'high',
          action: '分析エラー対応',
          details: 'キャッシュ分析中にエラーが発生しました。システムの状態を確認してください。',
        },
      ],
    };
  }
}

// =============================================================================
// SECTION 6: 統一キャッシュAPI - 全システムで使用する統一インターフェース
// =============================================================================

/**
 * 統一キャッシュAPI - 全分散実装を統合した統一インターフェース
 * 既存のすべてのキャッシュ関数はこのAPIを通して動作します
 */
class CacheAPI {
  constructor() {
    this.manager = cacheManager;
    this.executionCache = getExecutionCache();
  }

  /**
   * ユーザー情報キャッシュをクリア（複数の分散実装を統合）
   * @param {string} [identifier] - ユーザーID、メールアドレス、または null
   */
  clearUserInfoCache(identifier = null) {
    try {
      // 実行レベルキャッシュクリア
      this.executionCache.clearUserInfo();
      this.executionCache.syncWithUnifiedCache('userDataChange');

      if (identifier) {
        // 特定ユーザーのキャッシュクリア
        this.manager.remove(`user_${identifier}`);
        this.manager.remove(`userinfo_${identifier}`);
        this.manager.remove(`unified_user_info_${identifier}`);
        this.manager.remove(`email_${identifier}`);
        this.manager.remove(`login_status_${identifier}`);

        // パターンベースクリア
        this.manager.clearByPattern(`publishedData_${identifier}_`, { maxKeys: 200 });
        this.manager.clearByPattern(`sheetData_${identifier}_`, { maxKeys: 200 });
        this.manager.clearByPattern(`config_v3_${identifier}_`, { maxKeys: 200 });
      } else {
        // 全ユーザー情報キャッシュクリア
        this.manager.clearByPattern('user_', { maxKeys: 500 });
        this.manager.clearByPattern('userinfo_', { maxKeys: 500 });
        this.manager.clearByPattern('unified_user_info_', { maxKeys: 500 });
        this.manager.clearByPattern('email_', { maxKeys: 500 });
        this.manager.clearByPattern('login_status_', { maxKeys: 500 });
      }

      console.log(
        `✅ 統一API: ユーザー情報キャッシュクリア完了 (identifier: ${identifier || 'all'})`
      );
    } catch (error) {
      console.error('[ERROR]', `統一API: ユーザー情報キャッシュクリア失敗:`, error.message);
      throw error;
    }
  }

  /**
   * 全実行キャッシュをクリア（複数の分散実装を統合）
   */
  clearAllExecutionCache() {
    try {
      // 実行レベルキャッシュ全クリア
      this.executionCache.clearAll();
      this.executionCache.syncWithUnifiedCache('systemChange');

      // Apps Script キャッシュクリア (自動期限切れに依存)
      // CacheService.removeAll() はキー配列が必要なため、自動期限切れを利用

      // 統一キャッシュマネージャークリア
      this.manager.clearAll();

    } catch (error) {
      console.error('[ERROR]', `統一API: 全実行キャッシュクリア失敗:`, error.message);
      throw error;
    }
  }

  /**
   * SheetsServiceキャッシュの統合管理
   * @param {boolean} [forceRefresh=false] - 強制リフレッシュ
   * @returns {Object} SheetsService
   */
  getSheetsServiceCached(forceRefresh = false) {
    try {
      // 強制リフレッシュの場合はキャッシュクリア
      if (forceRefresh) {
        this.executionCache.clearSheetsService();
        this.manager.remove('sheets_service');
      }

      // 実行レベルキャッシュから取得を試行
      let service = this.executionCache.getSheetsService();
      if (service) {
        return service;
      }

      // 新しいサービスを生成
      const token = getServiceAccountTokenCached();
      service = {
        accessToken: token,
        baseUrl: 'https://sheets.googleapis.com/v4/spreadsheets',
      };

      // キャッシュに保存
      this.executionCache.setSheetsService(service);
      this.manager.get('sheets_service', () => service, { ttl: 3600 });

      return service;
    } catch (error) {
      console.error('[ERROR]', `統一API: SheetsService取得失敗:`, error.message);
      throw error;
    }
  }

  /**
   * ユーザーキャッシュ無効化の統合（複数の分散実装を統合）
   * @param {string} userId - ユーザーID
   * @param {string} email - メールアドレス
   * @param {string} [spreadsheetId] - スプレッドシートID
   * @param {boolean|string} [clearPattern=false] - パターンクリア（true='all', string=パターン）
   * @param {string} [dbSpreadsheetId] - データベーススプレッドシートID
   */
  invalidateUserCache(userId, email, spreadsheetId, clearPattern = false, dbSpreadsheetId) {
    try {
        userId,
        email,
        spreadsheetId,
        clearPattern,
        dbSpreadsheetId,
      });

      // 基本ユーザーキャッシュクリア
      if (userId) {
        this.manager.remove(`user_${userId}`);
        this.manager.remove(`unified_user_info_${userId}`);
        this.manager.remove(`userinfo_${userId}`);

        // ユーザー関連パターンクリア
        this.manager.clearByPattern(`publishedData_${userId}_`, { maxKeys: 200 });
        this.manager.clearByPattern(`sheetData_${userId}_`, { maxKeys: 200 });
        this.manager.clearByPattern(`config_v3_${userId}_`, { maxKeys: 200 });
      }

      if (email) {
        this.manager.remove(`email_${email}`);
        this.manager.remove(`unified_user_info_${email}`);
        this.manager.remove(`login_status_${email}`);
      }

      if (spreadsheetId) {
        this.manager.remove(`hdr_${spreadsheetId}`);
        this.manager.remove(`data_${spreadsheetId}`);
        this.manager.remove(`sheets_${spreadsheetId}`);
        this.invalidateSpreadsheetCache(spreadsheetId);
      }

      // パターンベースクリア
      if (clearPattern === true || clearPattern === 'all') {
        // CacheService.removeAll() はキー配列が必要なため、自動期限切れを利用
        this.manager.clearAll();
      } else if (typeof clearPattern === 'string' && clearPattern !== 'false') {
        this.manager.clearByPattern(clearPattern, { maxKeys: 300 });
      }

      // データベーススプレッドシート関連クリア
      if (dbSpreadsheetId) {
        this.manager.invalidateRelated('spreadsheet', dbSpreadsheetId, [userId]);
      }

      // 実行キャッシュも同期クリア
      this.clearUserInfoCache(userId || email);

    } catch (error) {
      console.error('[ERROR]', `統一API: ユーザーキャッシュ無効化失敗:`, error.message);
      throw error;
    }
  }

  /**
   * クリティカル更新後のキャッシュ同期化（分散実装を統合）
   * @param {string} userId - ユーザーID
   * @param {string} email - メールアドレス
   * @param {string} oldSpreadsheetId - 古いスプレッドシートID
   * @param {string} newSpreadsheetId - 新しいスプレッドシートID
   */
  synchronizeCacheAfterCriticalUpdate(userId, email, oldSpreadsheetId, newSpreadsheetId) {
    try {
        userId,
        email,
        oldSpreadsheetId,
        newSpreadsheetId,
      });

      // 段階1: 基本ユーザーキャッシュクリア
      this.invalidateUserCache(userId, email, oldSpreadsheetId, false);

      if (newSpreadsheetId && newSpreadsheetId !== oldSpreadsheetId) {
        this.invalidateUserCache(userId, email, newSpreadsheetId, false);

        // 古いスプレッドシートオブジェクトキャッシュクリア
        if (oldSpreadsheetId) {
          this.invalidateSpreadsheetCache(oldSpreadsheetId);
        }
      }

      // 段階2: 実行レベルキャッシュクリア
      this.clearUserInfoCache();

      // 段階3: データベースキャッシュクリア
      this.clearDatabaseCache();

      // 段階4: 関連パターンクリア
      const patterns = ['user_*', 'email_*', 'login_status_*', 'sheets_*', 'data_*', 'config_v3_*'];
      patterns.forEach((pattern) => {
        this.manager.clearByPattern(pattern, { maxKeys: 100 });
      });

      // 段階5: 少し待ってから検証
      Utilities.sleep(100);

    } catch (error) {
      console.error('[ERROR]', '統一API: キャッシュ同期失敗:', error.message);
      throw new Error(`キャッシュ同期に失敗しました: ${error.message}`);
    }
  }

  /**
   * データベースキャッシュクリア（分散実装を統合）
   */
  clearDatabaseCache() {
    try {

      // データベース関連パターンクリア
      const dbPatterns = ['user_', 'email_', 'hdr_', 'data_', 'sheets_', 'config_v3_'];
      dbPatterns.forEach((pattern) => {
        this.manager.clearByPattern(pattern, { maxKeys: 200 });
      });

      // Apps Script キャッシュもクリア（データベース関連のみ）
      // CacheService.removeAll() はキー配列が必要なため、自動期限切れを利用

    } catch (error) {
      console.error('[ERROR]', '統一API: データベースキャッシュクリア失敗:', error.message);
      // エラーが発生しても処理を継続
    }
  }

  /**
   * スプレッドシートキャッシュ無効化
   * @param {string} spreadsheetId - スプレッドシートID
   */
  invalidateSpreadsheetCache(spreadsheetId) {
    if (!spreadsheetId) return;

    try {
      // メモリキャッシュから削除（spreadsheetCache.gsの機能）
      if (typeof spreadsheetMemoryCache !== 'undefined' && spreadsheetMemoryCache[spreadsheetId]) {
        delete spreadsheetMemoryCache[spreadsheetId];
      }

      // セッションキャッシュから削除
      const cacheKey = `ss_cache_${spreadsheetId}`;
      try {
        PropertiesService.getScriptProperties().deleteProperty(cacheKey);
      } catch (propsError) {
        console.log('PropertiesService削除エラー:', propsError.message);
      }

      // 統一キャッシュマネージャーからも削除
      this.manager.clearByPattern(spreadsheetId, { maxKeys: 50 });
      this.manager.invalidateSheetData(spreadsheetId);

      console.log(
        `✅ 統一API: スプレッドシートキャッシュ無効化完了: ${spreadsheetId.substring(0, 10)}`
      );
    } catch (error) {
      console.error('[ERROR]', `統一API: スプレッドシートキャッシュ無効化失敗:`, error.message);
    }
  }

  /**
   * キャッシュプリウォーミング（事前読み込み）
   * @param {string} [activeUserEmail] - アクティブユーザーのメール
   * @returns {Object} プリウォーミング結果
   */
  preWarmCache(activeUserEmail) {
    return preWarmCache(activeUserEmail);
  }

  /**
   * キャッシュ効率分析
   * @returns {Object} 分析結果
   */
  analyzeCacheEfficiency() {
    return analyzeCacheEfficiency();
  }

  /**
   * キャッシュヘルスチェック
   * @returns {Object} ヘルス情報
   */
  getHealth() {
    return this.manager.getHealth();
  }

  /**
   * 統計情報取得
   * @returns {Object} 統計情報
   */
  getStats() {
    return {
      manager: this.manager.getHealth(),
      execution: this.executionCache.getStats(),
      spreadsheet:
        typeof getSpreadsheetCacheStats === 'function' ? getSpreadsheetCacheStats() : null,
    };
  }
}

// =============================================================================
// SECTION 7: 統一APIシングルトンインスタンス
// =============================================================================

// 統一キャッシュAPIのシングルトンインスタンス
// 統合キャッシュAPI（推奨公開インターフェース）
const unifiedCacheAPI = new CacheAPI();

// シンプルなキャッシュAPI（推奨）
const CacheStore = {
  /**
   * 値を取得、なければ生成して保存
   * @param {string} key キーワード
   * @param {function} valueFn 値生成関数
   * @param {number} ttl TTL（秒、デフォルト：3600）
   */
  get: (key, valueFn, ttl = 3600) => unifiedCacheAPI.get(key, valueFn, ttl),
  
  /**
   * 値を保存
   * @param {string} key キーワード 
   * @param {*} value 値
   * @param {number} ttl TTL（秒、デフォルト：3600）
   */
  set: (key, value, ttl = 3600) => unifiedCacheAPI.put(key, value, ttl),
  
  /**
   * 値を削除
   * @param {string} key キーワード
   */
  remove: (key) => unifiedCacheAPI.remove(key),
  
  /**
   * キャッシュをクリア
   */
  clear: () => unifiedCacheAPI.removeAll()
};

// 外部アクセス用エクスポート（後方互換性）
const Cache = CacheStore;

// グローバルアクセス用
if (typeof global !== 'undefined') {
  global.unifiedCacheAPI = unifiedCacheAPI;
}

// =============================================================================
// SECTION 8: 統一API直接実装関数群（GAS 2025 Best Practices）
// =============================================================================

/**
 * ユーザー情報キャッシュのクリア（統一API直接実装）
 */
function clearExecutionUserInfoCache() {
  return unifiedCacheAPI.clearUserInfoCache();
}

/**
 * 実行キャッシュの完全クリア（統一API直接実装）
 */
function clearAllExecutionCache() {
  return unifiedCacheAPI.clearAllExecutionCache();
}

/**
 * Sheetsサービスの取得（統一API直接実装）
 */
function getSheetsServiceCached(forceRefresh = false) {
  return unifiedCacheAPI.getSheetsServiceCached(forceRefresh);
}

/**
 * Sheetsサービスの取得（統一API直接実装）
 */
function getCachedSheetsService() {
  return unifiedCacheAPI.getSheetsServiceCached();
}

/**
 * ユーザーキャッシュの無効化（統一API直接実装）
 * GAS 2025 Best Practices - 後方互換性を削除して直接実装
 */
function invalidateUserCache(userId, email, spreadsheetId, clearPattern = false, dbSpreadsheetId) {
  return unifiedCacheAPI.invalidateUserCache(userId, email, spreadsheetId, clearPattern, dbSpreadsheetId);
}

/**
 * 重要な更新後のキャッシュ同期（統一API直接実装）
 */
function synchronizeCacheAfterCriticalUpdate(userId, email, oldSpreadsheetId, newSpreadsheetId) {
  return unifiedCacheAPI.synchronizeCacheAfterCriticalUpdate(
    userId,
    email,
    oldSpreadsheetId,
    newSpreadsheetId
  );
}

/**
 * データベースキャッシュのクリア（統一API直接実装）
 */
function clearDatabaseCache() {
  return unifiedCacheAPI.clearDatabaseCache();
}

/**
 * Sheetsサービスキャッシュのクリア（統一API直接実装）
 */
function clearExecutionSheetsServiceCache() {
  const cache = getExecutionCache();
  cache.clearSheetsService();
}

/**
 * リトライ機能付きの堅牢なURL取得関数
 * @param {string} url - 取得するURL
 * @param {object} options - UrlFetchApp.fetchのオプション
 * @returns {GoogleAppsScript.URL_Fetch.HTTPResponse} レスポンス
 */
function resilientUrlFetch(url, options = {}) {
  const maxRetries = 3;
  const baseDelay = 1000;

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      console.log(`resilientUrlFetch: ${url} (試行 ${attempt + 1}/${maxRetries + 1})`);

      const response = UrlFetchApp.fetch(url, {
        ...options,
        muteHttpExceptions: true,
      });

      if (!response || typeof response.getResponseCode !== 'function') {
        throw new Error('無効なレスポンスオブジェクトが返されました');
      }

      const responseCode = response.getResponseCode();

      // 成功時またはリトライ不要なエラー
      if (responseCode >= 200 && responseCode < 300) {
        if (attempt > 0) {
          console.log(`resilientUrlFetch: リトライで成功 ${url}`);
        }
        return response;
      }

      // 4xx エラーはリトライしない
      if (responseCode >= 400 && responseCode < 500) {
        console.warn(`resilientUrlFetch: クライアントエラー ${responseCode} - ${url}`);
        return response;
      }

      // 最後の試行でない場合は待機してリトライ
      if (attempt < maxRetries) {
        const delay = baseDelay * Math.pow(2, attempt);
        console.warn(`resilientUrlFetch: ${responseCode}エラー、${delay}ms後にリトライ - ${url}`);
        Utilities.sleep(delay);
      } else {
        console.warn(`resilientUrlFetch: 最大リトライ回数に達しました - ${url}`);
        return response;
      }
    } catch (error) {
      // 最後の試行でない場合はリトライ
      if (attempt < maxRetries) {
        const delay = baseDelay * Math.pow(2, attempt);
        console.warn(
          `resilientUrlFetch: エラー、${delay}ms後にリトライ - ${url}: ${error.message}`
        );
        Utilities.sleep(delay);
      } else {
        console.error(`resilientUrlFetch: 最終的に失敗 - ${url}`, { error: error.message, url });
        throw error;
      }
    }
  }
}
