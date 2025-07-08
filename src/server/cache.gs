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

    // 1. メモ化キャッシュのチェック
    if (enableMemoization && this.memoCache.has(key)) {
      const memoEntry = this.memoCache.get(key);
      // メモ化キャッシュの有効期限チェック（オプション）
      if (!memoEntry.ttl || (memoEntry.createdAt + memoEntry.ttl * 1000 > Date.now())) {
        debugLog(`[Cache] Memo hit for key: ${key}`);
        return memoEntry.value;
      }
    }

    // 2. Apps Scriptキャッシュのチェック
    try {
      const cachedValue = this.scriptCache.get(key);
      if (cachedValue !== null) {
        debugLog(`[Cache] ScriptCache hit for key: ${key}`);
        const parsedValue = JSON.parse(cachedValue);
        if (enableMemoization) {
          this.memoCache.set(key, { value: parsedValue, createdAt: Date.now(), ttl });
        }
        return parsedValue;
      }
    } catch (e) {
      console.warn(`[Cache] Failed to parse cache for key: ${key}`, e.message);
    }

    // 3. 値の生成とキャッシュ保存
    debugLog(`[Cache] Miss for key: ${key}. Generating new value.`);
    const newValue = valueFn();
    
    try {
      const stringValue = JSON.stringify(newValue);
      this.scriptCache.put(key, stringValue, ttl);
      if (enableMemoization) {
        this.memoCache.set(key, { value: newValue, createdAt: Date.now(), ttl });
      }
    } catch (e) {
      console.error(`[Cache] Failed to cache value for key: ${key}`, e.message);
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
    this.scriptCache.remove(key);
    this.memoCache.delete(key);
    debugLog(`[Cache] Removed cache for key: ${key}`);
  }

  /**
   * パターンに一致するキーのキャッシュを削除します。
   * @param {string} pattern - 削除するキーのパターン（部分一致）
   */
  clearByPattern(pattern) {
    // この機能はCacheServiceにネイティブサポートされていないため、
    // 全てのキーを取得してフィルタリングする必要があります。大規模なキャッシュではパフォーマンスに影響する可能性があります。
    // 現在の実装では、特定のキーを直接削除するアプローチを推奨します。
    // 今回は前方互換性のために、特定のキーを削除するラッパーとして実装します。
    this.remove(pattern);
    console.warn(`[Cache] clearByPattern was called. For better performance, use direct key removal. Key removed: ${pattern}`);
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
   * キャッシュの健全性情報を取得します。
   * @returns {object} 健全性情報
   */
  getHealth() {
    // CacheServiceにはサイズを取得するAPIがないため、メモ化キャッシュのサイズのみを返します。
    return {
      memoCacheSize: this.memoCache.size,
      status: 'ok'
    };
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
function getUserCached(userId) {
  return cacheManager.get(`user_${userId}`, () => findUserById(userId), { enableMemoization: true });
}

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
  });

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
 */
function invalidateUserCache(userId, email, spreadsheetId) {
  if (userId) {
    cacheManager.remove('user_' + userId);
  }
  if (email) {
    cacheManager.remove('email_' + email);
  }
  if (spreadsheetId) {
    cacheManager.clearByPattern(spreadsheetId);
  }
}

/**
 * @deprecated cacheManager.get() を使用してください
 */
function getWebAppUrlCached() {
  return cacheManager.get('WEB_APP_URL', () => {
    try {
      const url = ScriptApp.getService().getUrl();
      return url ? (url.endsWith('/') ? url.slice(0, -1) : url) : '';
    } catch (e) {
      return '';
    }
  });
}
