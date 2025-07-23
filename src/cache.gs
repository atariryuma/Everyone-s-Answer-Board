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
   * キャッシュから値を取得、なければ指定された関数で生成して保存します。
   * @param {string} key - キャッシュキー
   * @param {function} valueFn - 値を生成する関数
   * @param {object} [options] - オプション { ttl: number, enableMemoization: boolean }
   * @returns {*} キャッシュされた値
   */
  get(key, valueFn, options = {}) {
    const { ttl = this.defaultTTL, enableMemoization = false } = options;
    
    // パフォーマンス監視
    this.stats.totalOps++;
    const startTime = Date.now();

    // Input validation
    if (!key || typeof key !== 'string') {
      console.error(`[Cache] Invalid key: ${key}`);
      this.stats.errors++;
      return valueFn();
    }

    // 1. メモ化キャッシュのチェック
    if (enableMemoization && this.memoCache.has(key)) {
      try {
        const memoEntry = this.memoCache.get(key);
        // メモ化キャッシュの有効期限チェック（オプション）
        if (!memoEntry.ttl || (memoEntry.createdAt + memoEntry.ttl * 1000 > Date.now())) {
          debugLog(`[Cache] Memo hit for key: ${key}`);
          this.stats.hits++;
          return memoEntry.value;
        }
      } catch (e) {
        console.warn(`[Cache] Memo cache access failed for key: ${key}`, e.message);
        this.stats.errors++;
        this.memoCache.delete(key);
      }
    }

    // 2. Apps Scriptキャッシュのチェック
    try {
      const cachedValue = this.scriptCache.get(key);
      if (cachedValue !== null) {
        debugLog(`[Cache] ScriptCache hit for key: ${key}`);
        this.stats.hits++;
        const parsedValue = JSON.parse(cachedValue);
        if (enableMemoization) {
          this.memoCache.set(key, { value: parsedValue, createdAt: Date.now(), ttl });
        }
        return parsedValue;
      }
    } catch (e) {
      console.warn(`[Cache] Failed to parse cache for key: ${key}`, e.message);
      this.stats.errors++;
      // 破損したキャッシュエントリを削除
      try {
        this.scriptCache.remove(key);
      } catch (removeError) {
        console.warn(`[Cache] Failed to remove corrupted cache entry: ${key}`, removeError.message);
      }
    }

    // 3. 値の生成とキャッシュ保存
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
    
    // キャッシュ保存（エラーが発生しても値は返す）
    try {
      const stringValue = JSON.stringify(newValue);
      this.scriptCache.put(key, stringValue, ttl);
      if (enableMemoization) {
        this.memoCache.set(key, { value: newValue, createdAt: Date.now(), ttl });
      }
    } catch (e) {
      console.error(`[Cache] Failed to cache value for key: ${key}`, e.message);
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
   * インテリジェント無効化：関連するキャッシュエントリを連鎖的に無効化
   * @param {string} entityType - エンティティタイプ (user, spreadsheet, form等)
   * @param {string} entityId - エンティティID
   * @param {Array<string>} relatedIds - 関連するエンティティID群
   */
  invalidateRelated(entityType, entityId, relatedIds = []) {
    try {
      console.log(`🔗 関連キャッシュ無効化開始: ${entityType}/${entityId}`);
      
      // 1. メインエンティティのキャッシュ無効化
      const patterns = this._getInvalidationPatterns(entityType, entityId);
      patterns.forEach(pattern => {
        this.clearByPattern(pattern);
      });
      
      // 2. 関連エンティティのキャッシュ無効化
      relatedIds.forEach(relatedId => {
        if (relatedId && relatedId !== entityId) {
          const relatedPatterns = this._getInvalidationPatterns(entityType, relatedId);
          relatedPatterns.forEach(pattern => {
            this.clearByPattern(pattern);
          });
        }
      });
      
      // 3. クロスエンティティ関連キャッシュの無効化
      this._invalidateCrossEntityCache(entityType, entityId, relatedIds);
      
      console.log(`✅ 関連キャッシュ無効化完了: ${entityType}/${entityId}`);
      
    } catch (error) {
      console.warn(`⚠️ 関連キャッシュ無効化エラー: ${entityType}/${entityId}`, error.message);
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
   * クロスエンティティキャッシュの無効化
   * @private
   */
  _invalidateCrossEntityCache(entityType, entityId, relatedIds) {
    try {
      // ユーザー変更時は関連するスプレッドシート・フォームキャッシュも無効化
      if (entityType === 'user' && relatedIds.length > 0) {
        relatedIds.forEach(relatedId => {
          this.clearByPattern(`user_${entityId}_spreadsheet_${relatedId}*`);
          this.clearByPattern(`user_${entityId}_form_${relatedId}*`);
        });
      }
      
      // スプレッドシート変更時は関連するユーザーキャッシュも無効化
      if (entityType === 'spreadsheet' && relatedIds.length > 0) {
        relatedIds.forEach(userId => {
          this.clearByPattern(`user_${userId}_spreadsheet_${entityId}*`);
          this.clearByPattern(`config_${userId}*`); // ユーザー設定も無効化
        });
      }
      
    } catch (error) {
      console.warn('クロスエンティティキャッシュ無効化エラー:', error.message);
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
      console.warn('[Cache] Failed to clear memoization cache:', e.message);
    }
    
    try {
      // スクリプトキャッシュを完全にクリア
      this.scriptCache.removeAll();
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
      
      var service = getSheetsServiceCached();
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
      
      // タイムスタンプ列を除外してヘッダーのインデックスを生成
      headers.forEach(function(headerName, index) {
        if (headerName && headerName.trim() !== '' && headerName !== 'タイムスタンプ') {
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
    console.log('🔥 キャッシュプリウォーミング開始:', activeUserEmail);

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

    console.log('✅ キャッシュプリウォーミング完了:', results.preWarmedItems.length, 'items,', results.duration + 'ms');
    
    if (results.errors.length > 0) {
      console.warn('⚠️ プリウォーミング中のエラー:', results.errors);
    }

    return results;

  } catch (error) {
    results.duration = Date.now() - startTime;
    results.success = false;
    results.errors.push('fatal_error: ' + error.message);
    console.error('❌ キャッシュプリウォーミングエラー:', error);
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
    console.error('analyzeCacheEfficiency error:', error);
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

