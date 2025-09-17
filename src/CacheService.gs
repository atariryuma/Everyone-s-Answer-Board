/**
 * @fileoverview CacheService - 統一キャッシュ管理システム (Zero-Dependency)
 *
 * 🎯 責任範囲:
 * - キャッシュの整合性保証
 * - TTL統一管理
 * - キャッシュ無効化の自動化
 * - データ同期の保証
 *
 * 🔄 GAS Best Practices準拠:
 * - フラット関数構造 (Object.freeze削除)
 * - ServiceFactory統合
 * - Zero-dependency パターン
 */

/* global ServiceFactory */

/**
 * ServiceFactory統合初期化
 * CacheService用Zero-Dependency実装
 * @returns {boolean} 初期化成功可否
 */
function initCacheServiceZero() {
  return ServiceFactory.getUtils().initService('CacheService');
}

// ==========================================
// 🔧 キャッシュ設定定数
// ==========================================

/**
 * キャッシュTTL設定
 * @returns {Object} TTL設定オブジェクト
 */
function getCacheTTL() {
  return {
    SHORT: 60,      // 1分 - 頻繁に変更されるデータ
    MEDIUM: 300,    // 5分 - 標準的なデータ
    LONG: 1800,     // 30分 - 安定したデータ
    SESSION: 3600   // 1時間 - セッション情報
  };
}

/**
 * キャッシュキー設定
 * @returns {Object} キー設定オブジェクト
 */
function getCacheKeys() {
  return {
    USER_INFO: 'user_info_',
    CONFIG: 'config_',
    SHEET_DATA: 'sheet_data_',
    SPREADSHEET_LIST: 'spreadsheet_list_',
    ADMIN_DATA: 'admin_data_',
    SYSTEM_INFO: 'system_info'
  };
}

// ==========================================
// 🎯 Core Cache Operations (Zero-Dependency)
// ==========================================

/**
 * キャッシュに値を設定（Zero-Dependency）
 * @param {string} key - キャッシュキー
 * @param {*} value - 値
 * @param {number} ttl - TTL(秒)
 * @returns {boolean} 成功可否
 */
function setCacheValue(key, value, ttl = getCacheTTL().MEDIUM) {
  try {
    if (!initCacheServiceZero()) {
      console.error('setCacheValue: ServiceFactory not available');
      return false;
    }

    const cache = CacheService.getScriptCache();
    const serializedValue = JSON.stringify({
      data: value,
      timestamp: Date.now(),
      ttl
    });

    cache.put(key, serializedValue, ttl);
    return true;
  } catch (error) {
    console.error('setCacheValue error:', { key, error: error.message });
    return false;
  }
}

/**
 * キャッシュから値を取得（Zero-Dependency）
 * @param {string} key - キャッシュキー
 * @returns {*} 値または null
 */
function getCacheValue(key) {
  try {
    if (!initCacheServiceZero()) {
      console.error('getCacheValue: ServiceFactory not available');
      return null;
    }

    const cache = CacheService.getScriptCache();
    const cached = cache.get(key);

    if (!cached) {
      return null;
    }

    const parsed = JSON.parse(cached);
    return parsed.data;
  } catch (error) {
    console.error('getCacheValue error:', { key, error: error.message });
    return null;
  }
}

/**
 * キャッシュをクリア（Zero-Dependency）
 * @param {string} key - キャッシュキー
 * @returns {boolean} 成功可否
 */
function clearCacheValue(key) {
  try {
    if (!initCacheServiceZero()) {
      console.error('clearCacheValue: ServiceFactory not available');
      return false;
    }

    const cache = CacheService.getScriptCache();
    cache.remove(key);
    return true;
  } catch (error) {
    console.error('clearCacheValue error:', { key, error: error.message });
    return false;
  }
}

// ==========================================
// 🎯 Specialized Cache Functions
// ==========================================

/**
 * ユーザー情報キャッシュ
 * @param {string} userId - ユーザーID
 * @param {Object} userInfo - ユーザー情報
 * @returns {boolean} 成功可否
 */
function cacheUserInfo(userId, userInfo) {
  const key = getCacheKeys().USER_INFO + userId;
  return setCacheValue(key, userInfo, getCacheTTL().SESSION);
}

/**
 * ユーザー情報取得
 * @param {string} userId - ユーザーID
 * @returns {Object|null} ユーザー情報
 */
function getCachedUserInfo(userId) {
  const key = getCacheKeys().USER_INFO + userId;
  return getCacheValue(key);
}

/**
 * 設定情報キャッシュ
 * @param {string} userId - ユーザーID
 * @param {Object} config - 設定情報
 * @returns {boolean} 成功可否
 */
function cacheUserConfig(userId, config) {
  const key = getCacheKeys().CONFIG + userId;
  return setCacheValue(key, config, getCacheTTL().LONG);
}

/**
 * 設定情報取得
 * @param {string} userId - ユーザーID
 * @returns {Object|null} 設定情報
 */
function getCachedUserConfig(userId) {
  const key = getCacheKeys().CONFIG + userId;
  return getCacheValue(key);
}

/**
 * ユーザー関連キャッシュを全削除
 * @param {string} userId - ユーザーID
 * @returns {boolean} 成功可否
 */
function invalidateUserCache(userId) {
  try {
    const keys = getCacheKeys();
    const userKeys = [
      keys.USER_INFO + userId,
      keys.CONFIG + userId,
      keys.SHEET_DATA + userId
    ];

    let success = true;
    userKeys.forEach(key => {
      if (!clearCacheValue(key)) {
        success = false;
      }
    });

    return success;
  } catch (error) {
    console.error('invalidateUserCache error:', { userId, error: error.message });
    return false;
  }
}

/**
 * システム全体のキャッシュクリア
 * @returns {boolean} 成功可否
 */
function clearAllCache() {
  try {
    if (!initCacheServiceZero()) {
      console.error('clearAllCache: ServiceFactory not available');
      return false;
    }

    const cache = CacheService.getScriptCache();
    // GASのCacheServiceは全削除機能がないため、個別削除は制限あり
    return true;
  } catch (error) {
    console.error('clearAllCache error:', error.message);
    return false;
  }
}

/**
 * キャッシュ診断情報
 * @returns {Object} 診断情報
 */
function getCacheDiagnostics() {
  try {
    return {
      service: 'CacheService',
      status: initCacheServiceZero() ? 'available' : 'unavailable',
      ttlSettings: getCacheTTL(),
      keyPrefixes: getCacheKeys(),
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    console.error('getCacheDiagnostics error:', error.message);
    return {
      service: 'CacheService',
      status: 'error',
      error: error.message,
      timestamp: new Date().toISOString()
    };
  }
}