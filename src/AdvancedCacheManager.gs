/**
 * @fileoverview 超高速キャッシュ管理システム - 2024年最新GAS技術
 * V8ランタイムの最新機能とメモ化パターンを活用
 */

// V8ランタイムを活用したES6+の定数定義
const CACHE_CONFIG = {
  // 最大6時間制限を活用
  MAX_TTL: 21600, // 6時間
  DEFAULT_TTL: 3600, // 1時間
  SHORT_TTL: 300, // 5分
  
  // マルチスコープ戦略
  SCOPES: {
    DOCUMENT: 'document',
    SCRIPT: 'script', 
    USER: 'user'
  },
  
  // キーパターン
  PATTERNS: {
    USER: 'u_',
    AUTH: 'auth_',
    HEADERS: 'hdr_',
    ROSTER: 'rstr_',
    TEMP: 'tmp_'
  }
};

/**
 * 超高速マルチレイヤーキャッシュシステム
 * 2024年最新のGASパフォーマンスパターンを実装
 */
class AdvancedCacheManager {
  
  /**
   * インテリジェントキャッシュ取得
   * @param {string} key - キャッシュキー
   * @param {Function} fetchFunction - データ取得関数
   * @param {Object} options - オプション設定
   */
  static smartGet(key, fetchFunction, options = {}) {
    const {
      ttl = CACHE_CONFIG.DEFAULT_TTL,
      scope = CACHE_CONFIG.SCOPES.SCRIPT,
      fallbackToProperties = true,
      enableMemoization = true
    } = options;
    
    // メモ化チェック（V8ランタイムの高速オブジェクト操作を活用）
    if (enableMemoization && this._memoCache.has(key)) {
      const memoEntry = this._memoCache.get(key);
      if (Date.now() - memoEntry.timestamp < CACHE_CONFIG.SHORT_TTL * 1000) {
        return memoEntry.value;
      }
    }
    
    // マルチスコープキャッシュチェック
    const cache = this._getCache(scope);
    let cachedValue = cache.get(key);
    
    if (cachedValue !== null) {
      try {
        const parsed = JSON.parse(cachedValue);
        // メモ化更新
        if (enableMemoization) {
          this._memoCache.set(key, { value: parsed, timestamp: Date.now() });
        }
        return parsed;
      } catch (e) {
        console.warn(`キャッシュデータ解析エラー (${key}):`, e.message);
      }
    }
    
    // PropertiesServiceフォールバック（長期キャッシュ）
    if (fallbackToProperties) {
      const propValue = this._getFromProperties(key);
      if (propValue !== null) {
        // CacheServiceにも保存してアクセス高速化
        cache.put(key, JSON.stringify(propValue), Math.min(ttl, CACHE_CONFIG.MAX_TTL));
        return propValue;
      }
    }
    
    // データ新規取得とキャッシュ
    try {
      const value = fetchFunction();
      if (value !== null && value !== undefined) {
        this._smartPut(key, value, { ttl, scope, fallbackToProperties });
        
        // メモ化
        if (enableMemoization) {
          this._memoCache.set(key, { value, timestamp: Date.now() });
        }
      }
      return value;
    } catch (error) {
      console.error(`データ取得エラー (${key}):`, error.message);
      return null;
    }
  }
  
  /**
   * インテリジェントキャッシュ保存
   * @param {string} key - キーー
   * @param {any} value - 値
   * @param {Object} options - オプション
   */
  static _smartPut(key, value, options = {}) {
    const {
      ttl = CACHE_CONFIG.DEFAULT_TTL,
      scope = CACHE_CONFIG.SCOPES.SCRIPT,
      fallbackToProperties = true
    } = options;
    
    const serialized = JSON.stringify(value);
    const cache = this._getCache(scope);
    
    // CacheServiceに保存
    try {
      cache.put(key, serialized, Math.min(ttl, CACHE_CONFIG.MAX_TTL));
    } catch (e) {
      console.warn(`CacheService保存警告 (${key}):`, e.message);
    }
    
    // 長期保存用PropertiesService
    if (fallbackToProperties && ttl > CACHE_CONFIG.DEFAULT_TTL) {
      try {
        const propData = {
          value: value,
          expiresAt: Date.now() + (ttl * 1000),
          created: Date.now()
        };
        PropertiesService.getScriptProperties().setProperty(
          `CACHE_${key}`, 
          JSON.stringify(propData)
        );
      } catch (e) {
        console.warn(`PropertiesService保存警告 (${key}):`, e.message);
      }
    }
  }
  
  /**
   * スコープ別キャッシュインスタンス取得
   */
  static _getCache(scope) {
    switch (scope) {
      case CACHE_CONFIG.SCOPES.DOCUMENT:
        return CacheService.getDocumentCache();
      case CACHE_CONFIG.SCOPES.USER:
        return CacheService.getUserCache();
      default:
        return CacheService.getScriptCache();
    }
  }
  
  /**
   * PropertiesServiceからデータ取得
   */
  static _getFromProperties(key) {
    try {
      const props = PropertiesService.getScriptProperties();
      const propValue = props.getProperty(`CACHE_${key}`);
      
      if (propValue) {
        const data = JSON.parse(propValue);
        if (data.expiresAt > Date.now()) {
          return data.value;
        } else {
          // 期限切れデータ削除
          props.deleteProperty(`CACHE_${key}`);
        }
      }
    } catch (e) {
      console.warn(`PropertiesService取得警告 (${key}):`, e.message);
    }
    return null;
  }
  
  /**
   * バッチキャッシュ操作（高速化）
   */
  static batchGet(keys, fetchFunction, options = {}) {
    const results = {};
    const missingKeys = [];
    
    // 既存キャッシュ一括チェック
    const cache = this._getCache(options.scope || CACHE_CONFIG.SCOPES.SCRIPT);
    keys.forEach(key => {
      const cached = cache.get(key);
      if (cached !== null) {
        try {
          results[key] = JSON.parse(cached);
        } catch (e) {
          missingKeys.push(key);
        }
      } else {
        missingKeys.push(key);
      }
    });
    
    // 不足データの一括取得
    if (missingKeys.length > 0) {
      try {
        const newData = fetchFunction(missingKeys);
        Object.keys(newData).forEach(key => {
          results[key] = newData[key];
          this._smartPut(key, newData[key], options);
        });
      } catch (e) {
        console.error('バッチデータ取得エラー:', e.message);
      }
    }
    
    return results;
  }
  
  /**
   * 条件付きキャッシュクリア
   */
  static conditionalClear(pattern, condition = null) {
    // CacheServiceはパターンベース削除をサポートしないため、
    // PropertiesServiceで管理
    try {
      const props = PropertiesService.getScriptProperties();
      const allProperties = props.getProperties();
      
      Object.keys(allProperties).forEach(key => {
        if (key.startsWith('CACHE_') && key.includes(pattern)) {
          if (!condition || condition(key, allProperties[key])) {
            props.deleteProperty(key);
          }
        }
      });
    } catch (e) {
      console.error('条件付きクリアエラー:', e.message);
    }
  }
  
  /**
   * キャッシュ統計と健康状態チェック
   */
  static getHealth() {
    try {
      const props = PropertiesService.getScriptProperties();
      const allProperties = props.getProperties();
      
      const cacheProps = Object.keys(allProperties)
        .filter(key => key.startsWith('CACHE_'));
      
      let totalSize = 0;
      let expiredCount = 0;
      const now = Date.now();
      
      cacheProps.forEach(key => {
        const value = allProperties[key];
        totalSize += value.length;
        
        try {
          const data = JSON.parse(value);
          if (data.expiresAt && data.expiresAt < now) {
            expiredCount++;
          }
        } catch (e) {
          // 無効なデータ
          expiredCount++;
        }
      });
      
      return {
        totalItems: cacheProps.length,
        estimatedSize: totalSize,
        expiredItems: expiredCount,
        healthScore: Math.max(0, 100 - (expiredCount / cacheProps.length * 100)),
        memoCacheSize: this._memoCache.size
      };
    } catch (e) {
      console.error('健康状態チェックエラー:', e.message);
      return { error: e.message };
    }
  }
}

// V8ランタイム高速Mapインスタンス（メモ化用）
AdvancedCacheManager._memoCache = new Map();

/**
 * 特化型キャッシュヘルパー関数群
 */

/**
 * ユーザー情報専用高速キャッシュ
 */
function getUserCached(userId) {
  return AdvancedCacheManager.smartGet(
    `${CACHE_CONFIG.PATTERNS.USER}${userId}`,
    () => fetchUserFromDatabase('userId', userId),
    {
      ttl: CACHE_CONFIG.DEFAULT_TTL,
      scope: CACHE_CONFIG.SCOPES.SCRIPT,
      enableMemoization: true
    }
  );
}

/**
 * 認証トークン専用キャッシュ
 */
function getAuthTokenCached() {
  return AdvancedCacheManager.smartGet(
    `${CACHE_CONFIG.PATTERNS.AUTH}token`,
    () => generateNewServiceAccountToken(),
    {
      ttl: 3300, // 55分（1時間-5分バッファ）
      scope: CACHE_CONFIG.SCOPES.SCRIPT,
      fallbackToProperties: false // セキュリティ上の理由
    }
  );
}

/**
 * ヘッダー情報専用キャッシュ
 */
function getHeadersCached(spreadsheetId, sheetName) {
  return AdvancedCacheManager.smartGet(
    `${CACHE_CONFIG.PATTERNS.HEADERS}${spreadsheetId}_${sheetName}`,
    () => {
      const service = getSheetsService();
      const response = batchGetSheetsData(service, spreadsheetId, [`${sheetName}!1:1`]);
      const headers = response.valueRanges[0].values?.[0] || [];
      
      // ヘッダーインデックスマップを生成
      const indices = {};
      Object.values(COLUMN_HEADERS).forEach(header => {
        const index = headers.indexOf(header);
        if (index !== -1) indices[header] = index;
      });
      
      return indices;
    },
    {
      ttl: CACHE_CONFIG.MAX_TTL, // ヘッダーは変更頻度が低い
      scope: CACHE_CONFIG.SCOPES.SCRIPT
    }
  );
}

/**
 * 高速キャッシュクリーンアップ（期限切れアイテム削除）
 */
function performCacheCleanup() {
  AdvancedCacheManager.conditionalClear('', (key, value) => {
    try {
      const data = JSON.parse(value);
      return data.expiresAt && data.expiresAt < Date.now();
    } catch (e) {
      return true; // 無効なデータは削除
    }
  });
  
  // メモ化キャッシュもクリア
  AdvancedCacheManager._memoCache.clear();
}