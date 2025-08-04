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
   * @param {object} [options] - オプション { ttl: number, enableMemoization: boolean }
   * @returns {*} キャッシュされた値
   */
  get(key, valueFn, options = {}) {
    const { ttl = this.defaultTTL, enableMemoization = false } = options;

    // パフォーマンス監視とバリデーション
    this.stats.totalOps++;
    if (!this.validateKey(key)) {
      this.stats.errors++;
      return valueFn ? valueFn() : null;
    }

    // 階層化キャッシュアクセス（最適化）
    const cacheResult = this.getFromCacheHierarchy(key, enableMemoization, ttl);
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
   * 階層化キャッシュから値を取得
   * @param {string} key - キー
   * @param {boolean} enableMemoization - メモ化キャッシュ有効
   * @param {number} ttl - TTL
   * @returns {Object} { found: boolean, value: any }
   */
  getFromCacheHierarchy(key, enableMemoization, ttl) {
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
      warnLog('[Cache] batchGet failed', e.message);
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

      // HTTPステータスチェック
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
function invalidateUserCache(userId, email, spreadsheetId, clearPattern) {
  const keysToRemove = [];

  if (userId) {
    keysToRemove.push('user_' + userId);
    keysToRemove.push('unified_user_info_' + userId);
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
      getWebAppUrlCached();
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

