/**
 * @fileoverview キャッシュサービス - シンプルで統一されたキャッシュ管理
 */

// グローバル名前空間の作成
if (typeof Services === 'undefined') {
  var Services = {};
}

Services.Cache = (function() {
  /**
   * キャッシュ設定
   */
  const CONFIG = {
    DEFAULT_TTL: 300, // 5分（秒単位）
    USER_DATA_TTL: 600, // 10分
    SPREADSHEET_TTL: 1800, // 30分
    MAX_KEY_LENGTH: 250 // GAS CacheServiceの制限
  };

  /**
   * メモリキャッシュ（実行中のみ有効）
   */
  const memoryCache = new Map();
  const memoryCacheExpiry = new Map();

  /**
   * キャッシュから値を取得（メモリ優先、次にCacheService）
   * @param {string} key - キャッシュキー
   * @returns {*} キャッシュされた値またはnull
   */
  function getValue(key) {
    // メモリキャッシュをチェック
    if (memoryCache.has(key)) {
      const expiry = memoryCacheExpiry.get(key);
      if (expiry && expiry > Date.now()) {
        return memoryCache.get(key);
      }
      // 期限切れの場合は削除
      memoryCache.delete(key);
      memoryCacheExpiry.delete(key);
    }
    
    // CacheServiceをチェック
    try {
      const cache = CacheService.getScriptCache();
      const value = cache.get(key);
      if (value) {
        try {
          return JSON.parse(value);
        } catch (e) {
          return value; // JSON以外の場合はそのまま返す
        }
      }
    } catch (error) {
      if (typeof Services !== 'undefined' && Services.Error) {
        Services.Error.logError(error, 'Cache.getValue', Services.Error.SEVERITY.LOW, Services.Error.CATEGORIES.CACHE, { key });
      }
    }
    
    return null;
  }

  /**
   * キャッシュに値を設定
   * @param {string} key - キャッシュキー
   * @param {*} value - 保存する値
   * @param {number} [ttl] - TTL（秒）
   */
  function setValue(key, value, ttl = CONFIG.DEFAULT_TTL) {
    // メモリキャッシュに保存
    memoryCache.set(key, value);
    memoryCacheExpiry.set(key, Date.now() + (ttl * 1000));
    
    // CacheServiceに保存
    try {
      const cache = CacheService.getScriptCache();
      const serialized = typeof value === 'string' ? value : JSON.stringify(value);
      
      // キーの長さをチェック
      if (key.length > CONFIG.MAX_KEY_LENGTH) {
        if (typeof warnLog === 'function') {
          warnLog(`Cache key too long: ${key.substring(0, 50)}...`);
        }
        return;
      }
      
      cache.put(key, serialized, ttl);
    } catch (error) {
      if (typeof Services !== 'undefined' && Services.Error) {
        Services.Error.logError(error, 'Cache.setValue', Services.Error.SEVERITY.LOW, Services.Error.CATEGORIES.CACHE, { key });
      }
    }
  }

  /**
   * キャッシュから値を削除
   * @param {string} key - キャッシュキー
   */
  function removeValue(key) {
    // メモリキャッシュから削除
    memoryCache.delete(key);
    memoryCacheExpiry.delete(key);
    
    // CacheServiceから削除
    try {
      const cache = CacheService.getScriptCache();
      cache.remove(key);
    } catch (error) {
      if (typeof Services !== 'undefined' && Services.Error) {
        Services.Error.logError(error, 'Cache.removeValue', Services.Error.SEVERITY.LOW, Services.Error.CATEGORIES.CACHE, { key });
      }
    }
  }

  /**
   * パターンに一致するキャッシュをクリア
   * @param {string} pattern - キーパターン（正規表現）
   */
  function clearByPattern(pattern) {
    const regex = new RegExp(pattern);
    
    // メモリキャッシュをクリア
    for (const key of memoryCache.keys()) {
      if (regex.test(key)) {
        memoryCache.delete(key);
        memoryCacheExpiry.delete(key);
      }
    }
    
    // CacheServiceは個別キーの列挙ができないため、既知のキーパターンで削除
    const knownPatterns = [
      'user_',
      'spreadsheet_',
      'header_',
      'data_',
      'config_'
    ];
    
    try {
      const cache = CacheService.getScriptCache();
      for (const prefix of knownPatterns) {
        if (pattern.includes(prefix)) {
          // 一般的なキーパターンで削除を試みる
          for (let i = 0; i < 10; i++) {
            cache.remove(prefix + i);
          }
        }
      }
    } catch (error) {
      if (typeof Services !== 'undefined' && Services.Error) {
        Services.Error.logError(error, 'Cache.clearByPattern', Services.Error.SEVERITY.LOW, Services.Error.CATEGORIES.CACHE, { pattern });
      }
    }
  }

  /**
   * 全キャッシュをクリア
   */
  function clearAll() {
    // メモリキャッシュをクリア
    memoryCache.clear();
    memoryCacheExpiry.clear();
    
    // CacheServiceは全削除メソッドがないため、警告のみ
    if (typeof warnLog === 'function') {
      warnLog('Note: CacheService cache cannot be fully cleared programmatically');
    }
  }

  /**
   * ユーザーデータのキャッシュを取得/設定
   * @param {string} userId - ユーザーID
   * @param {Function} [fetchFunction] - データ取得関数
   * @returns {Object} ユーザーデータ
   */
  function getUserDataCached(userId, fetchFunction) {
    const key = `user_${userId}`;
    let userData = getValue(key);
    
    if (!userData && fetchFunction) {
      userData = fetchFunction(userId);
      if (userData) {
        setValue(key, userData, CONFIG.USER_DATA_TTL);
      }
    }
    
    return userData;
  }

  /**
   * スプレッドシートのキャッシュを取得/設定
   * @param {string} spreadsheetId - スプレッドシートID
   * @param {Function} [fetchFunction] - スプレッドシート取得関数
   * @returns {Object} スプレッドシートオブジェクト
   */
  function getSpreadsheetCached(spreadsheetId, fetchFunction) {
    const key = `spreadsheet_${spreadsheetId}`;
    let spreadsheet = getValue(key);
    
    if (!spreadsheet && fetchFunction) {
      spreadsheet = fetchFunction(spreadsheetId);
      if (spreadsheet) {
        setValue(key, spreadsheet, CONFIG.SPREADSHEET_TTL);
      }
    }
    
    return spreadsheet;
  }

  /**
   * ユーザー関連のキャッシュをクリア
   * @param {string} userId - ユーザーID
   */
  function clearUserCache(userId) {
    clearByPattern(`user_${userId}`);
    clearByPattern(`data_${userId}`);
  }

  /**
   * スプレッドシート関連のキャッシュをクリア
   * @param {string} spreadsheetId - スプレッドシートID
   */
  function clearSpreadsheetCache(spreadsheetId) {
    clearByPattern(`spreadsheet_${spreadsheetId}`);
    clearByPattern(`header_${spreadsheetId}`);
  }

  /**
   * キャッシュ統計情報を取得
   * @returns {Object} キャッシュ統計
   */
  function getStats() {
    return {
      memoryEntries: memoryCache.size,
      memorySize: Array.from(memoryCache.values()).reduce((size, value) => {
        return size + JSON.stringify(value).length;
      }, 0),
      timestamp: new Date().toISOString()
    };
  }

  // Public API
  return {
    CONFIG,
    getValue,
    setValue,
    removeValue,
    clearByPattern,
    clearAll,
    getUserDataCached,
    getSpreadsheetCached,
    clearUserCache,
    clearSpreadsheetCache,
    getStats
  };
})();

// グローバル関数の互換性レイヤー（既存コードとの互換性のため）
function getCacheValue(key) {
  return Services.Cache.getValue(key);
}

function setCacheValue(key, value, ttl) {
  return Services.Cache.setValue(key, value, ttl);
}

function removeCacheValue(key) {
  return Services.Cache.removeValue(key);
}

function clearAllCache() {
  return Services.Cache.clearAll();
}

function clearUserCache(userId) {
  return Services.Cache.clearUserCache(userId);
}

function clearSpreadsheetCache(spreadsheetId) {
  return Services.Cache.clearSpreadsheetCache(spreadsheetId);
}

function getSpreadsheetCached(spreadsheetId, fetchFunction) {
  return Services.Cache.getSpreadsheetCached(spreadsheetId, fetchFunction);
}