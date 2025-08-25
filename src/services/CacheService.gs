/**
 * @fileoverview キャッシュサービス - シンプルで統一されたキャッシュ管理
 */

/**
 * キャッシュ設定
 */
const CACHE_CONFIG = {
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
function getCacheValue(key) {
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
    logError(error, 'getCacheValue', ERROR_SEVERITY.LOW, ERROR_CATEGORIES.CACHE, { key });
  }
  
  return null;
}

/**
 * キャッシュに値を設定
 * @param {string} key - キャッシュキー
 * @param {*} value - 保存する値
 * @param {number} [ttl] - TTL（秒）
 */
function setCacheValue(key, value, ttl = CACHE_CONFIG.DEFAULT_TTL) {
  // メモリキャッシュに保存
  memoryCache.set(key, value);
  memoryCacheExpiry.set(key, Date.now() + (ttl * 1000));
  
  // CacheServiceに保存
  try {
    const cache = CacheService.getScriptCache();
    const serialized = typeof value === 'string' ? value : JSON.stringify(value);
    
    // キーの長さをチェック
    if (key.length > CACHE_CONFIG.MAX_KEY_LENGTH) {
      warnLog(`Cache key too long: ${key.substring(0, 50)}...`);
      return;
    }
    
    cache.put(key, serialized, ttl);
  } catch (error) {
    logError(error, 'setCacheValue', ERROR_SEVERITY.LOW, ERROR_CATEGORIES.CACHE, { key });
  }
}

/**
 * キャッシュから値を削除
 * @param {string} key - キャッシュキー
 */
function removeCacheValue(key) {
  // メモリキャッシュから削除
  memoryCache.delete(key);
  memoryCacheExpiry.delete(key);
  
  // CacheServiceから削除
  try {
    const cache = CacheService.getScriptCache();
    cache.remove(key);
  } catch (error) {
    logError(error, 'removeCacheValue', ERROR_SEVERITY.LOW, ERROR_CATEGORIES.CACHE, { key });
  }
}

/**
 * パターンに一致するキャッシュをクリア
 * @param {string} pattern - キーパターン（正規表現）
 */
function clearCacheByPattern(pattern) {
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
    logError(error, 'clearCacheByPattern', ERROR_SEVERITY.LOW, ERROR_CATEGORIES.CACHE, { pattern });
  }
}

/**
 * 全キャッシュをクリア
 */
function clearAllCache() {
  // メモリキャッシュをクリア
  memoryCache.clear();
  memoryCacheExpiry.clear();
  
  // CacheServiceは全削除メソッドがないため、警告のみ
  warnLog('Note: CacheService cache cannot be fully cleared programmatically');
}

/**
 * ユーザーデータのキャッシュを取得/設定
 * @param {string} userId - ユーザーID
 * @param {Function} [fetchFunction] - データ取得関数
 * @returns {Object} ユーザーデータ
 */
function getUserDataCached(userId, fetchFunction) {
  const key = `user_${userId}`;
  let userData = getCacheValue(key);
  
  if (!userData && fetchFunction) {
    userData = fetchFunction(userId);
    if (userData) {
      setCacheValue(key, userData, CACHE_CONFIG.USER_DATA_TTL);
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
  let spreadsheet = getCacheValue(key);
  
  if (!spreadsheet && fetchFunction) {
    spreadsheet = fetchFunction(spreadsheetId);
    if (spreadsheet) {
      setCacheValue(key, spreadsheet, CACHE_CONFIG.SPREADSHEET_TTL);
    }
  }
  
  return spreadsheet;
}

/**
 * ユーザー関連のキャッシュをクリア
 * @param {string} userId - ユーザーID
 */
function clearUserCache(userId) {
  clearCacheByPattern(`user_${userId}`);
  clearCacheByPattern(`data_${userId}`);
}

/**
 * スプレッドシート関連のキャッシュをクリア
 * @param {string} spreadsheetId - スプレッドシートID
 */
function clearSpreadsheetCache(spreadsheetId) {
  clearCacheByPattern(`spreadsheet_${spreadsheetId}`);
  clearCacheByPattern(`header_${spreadsheetId}`);
}

/**
 * キャッシュ統計情報を取得
 * @returns {Object} キャッシュ統計
 */
function getCacheStats() {
  return {
    memoryEntries: memoryCache.size,
    memorySize: Array.from(memoryCache.values()).reduce((size, value) => {
      return size + JSON.stringify(value).length;
    }, 0),
    timestamp: new Date().toISOString()
  };
}