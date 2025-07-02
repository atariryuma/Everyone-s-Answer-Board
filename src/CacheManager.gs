/**
 * @fileoverview 統合キャッシュ管理システム
 */

class CacheManager {
  static SHORT_TTL = 300;    // 5分 - CacheService
  static MEDIUM_TTL = 1800;  // 30分 - CacheService
  static LONG_TTL = 86400;   // 24時間 - PropertiesService
  
  /**
   * 階層キャッシュから値を取得
   * @param {string} key - キャッシュキー
   * @param {function} fetchFunction - データ取得関数
   * @param {string} level - キャッシュレベル ('short', 'medium', 'long')
   * @returns {any} キャッシュされた値またはfetchFunctionの結果
   */
  static async get(key, fetchFunction, level = 'short') {
    // 1. CacheServiceから取得を試行
    const cache = CacheService.getScriptCache();
    let cachedValue = cache.get(key);
    
    if (cachedValue !== null) {
      debugLog(`CacheServiceヒット: ${key}`);
      try {
        return JSON.parse(cachedValue);
      } catch (e) {
        return cachedValue;
      }
    }
    
    // 2. 長期キャッシュの場合はPropertiesServiceを確認
    if (level === 'long') {
      const props = PropertiesService.getScriptProperties();
      const propKey = `CACHE_${key}`;
      const propValue = props.getProperty(propKey);
      
      if (propValue) {
        try {
          const data = JSON.parse(propValue);
          const now = Date.now();
          
          if (data.expiresAt > now) {
            debugLog(`PropertiesServiceヒット: ${key}`);
            // CacheServiceにも保存して次回高速化
            cache.put(key, JSON.stringify(data.value), Math.min(this.SHORT_TTL, Math.floor((data.expiresAt - now) / 1000)));
            return data.value;
          } else {
            // 期限切れのデータを削除
            props.deleteProperty(propKey);
          }
        } catch (e) {
          props.deleteProperty(propKey);
        }
      }
    }
    
    // 3. データを新規取得
    const value = await fetchFunction();
    
    if (value !== null && value !== undefined) {
      this.set(key, value, level);
    }
    
    return value;
  }
  
  /**
   * 階層キャッシュに値を保存
   * @param {string} key - キャッシュキー
   * @param {any} value - 保存する値
   * @param {string} level - キャッシュレベル
   */
  static set(key, value, level = 'short') {
    try {
      const cache = CacheService.getScriptCache();
      let ttl;
      
      switch (level) {
        case 'short':
          ttl = this.SHORT_TTL;
          break;
        case 'medium':
          ttl = this.MEDIUM_TTL;
          break;
        case 'long':
          ttl = this.LONG_TTL;
          break;
        default:
          ttl = this.SHORT_TTL;
      }
      
      const serializedValue = typeof value === 'string' ? value : JSON.stringify(value);
      
      // CacheServiceに保存
      cache.put(key, serializedValue, Math.min(ttl, 21600)); // CacheServiceの最大6時間制限
      
      // 長期キャッシュの場合はPropertiesServiceにも保存
      if (level === 'long') {
        const props = PropertiesService.getScriptProperties();
        const expiresAt = Date.now() + (ttl * 1000);
        const data = {
          value: value,
          expiresAt: expiresAt
        };
        
        const propKey = `CACHE_${key}`;
        try {
          props.setProperty(propKey, JSON.stringify(data));
        } catch (e) {
          // PropertiesServiceの容量制限に達した場合の処理
          console.warn(`PropertiesService保存失敗: ${e.message}`);
        }
      }
      
      debugLog(`キャッシュ保存 (${level}): ${key}`);
    } catch (error) {
      console.error(`キャッシュ保存エラー: ${error.message}`);
    }
  }
  
  /**
   * キャッシュを削除
   * @param {string} key - キャッシュキー
   */
  static remove(key) {
    const cache = CacheService.getScriptCache();
    cache.remove(key);
    
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty(`CACHE_${key}`);
    
    debugLog(`キャッシュ削除: ${key}`);
  }
  
  /**
   * パターンマッチでキャッシュを削除
   * @param {string} pattern - 削除パターン
   */
  static removePattern(pattern) {
    const cache = CacheService.getScriptCache();
    const props = PropertiesService.getScriptProperties();
    
    // PropertiesServiceからパターンマッチするキーを削除
    const properties = props.getProperties();
    Object.keys(properties).forEach(key => {
      if (key.startsWith('CACHE_') && key.includes(pattern)) {
        props.deleteProperty(key);
        const cacheKey = key.replace('CACHE_', '');
        cache.remove(cacheKey);
        debugLog(`パターン削除: ${cacheKey}`);
      }
    });
  }
  
  /**
   * ヘッダー情報をキャッシュ
   * @param {string} spreadsheetId - スプレッドシートID
   * @param {string} sheetName - シート名
   * @returns {object} ヘッダーインデックス
   */
  static async getHeaderIndices(spreadsheetId, sheetName) {
    const key = `headers_${spreadsheetId}_${sheetName}`;
    
    return await this.get(key, async () => {
      const service = DatabaseManager.getSheetsService();
      const response = await service.batchGet(spreadsheetId, [`${sheetName}!1:1`]);
      const headers = response.valueRanges[0].values ? response.valueRanges[0].values[0] : [];
      
      const indices = {};
      Object.values(COLUMN_HEADERS).forEach(header => {
        const index = headers.indexOf(header);
        if (index !== -1) {
          indices[header] = index;
        }
      });
      
      return indices;
    }, 'medium');
  }
  
  /**
   * 名簿データをキャッシュ
   * @param {string} spreadsheetId - スプレッドシートID
   * @returns {object} 名簿マップ
   */
  static async getRosterMap(spreadsheetId) {
    const key = `roster_${spreadsheetId}`;
    
    return await this.get(key, async () => {
      try {
        const service = DatabaseManager.getSheetsService();
        const response = await service.batchGet(spreadsheetId, [`${ROSTER_CONFIG.SHEET_NAME}!A:Z`]);
        const values = response.valueRanges[0].values || [];
        
        if (values.length === 0) return {};
        
        const headers = values[0];
        const emailIndex = headers.indexOf(ROSTER_CONFIG.EMAIL_COLUMN);
        const nameIndex = headers.indexOf(ROSTER_CONFIG.NAME_COLUMN);
        const classIndex = headers.indexOf(ROSTER_CONFIG.CLASS_COLUMN);
        
        const rosterMap = {};
        
        for (let i = 1; i < values.length; i++) {
          const row = values[i];
          if (emailIndex !== -1 && row[emailIndex]) {
            rosterMap[row[emailIndex]] = {
              name: nameIndex !== -1 ? row[nameIndex] : '',
              class: classIndex !== -1 ? row[classIndex] : ''
            };
          }
        }
        
        return rosterMap;
      } catch (e) {
        console.error('名簿データ取得エラー:', e);
        return {};
      }
    }, 'medium');
  }
  
  /**
   * 全キャッシュクリア
   */
  static clearAll() {
    // CacheServiceクリア
    const cache = CacheService.getScriptCache();
    cache.removeAll(Object.keys(cache.getAll || {}));
    
    // PropertiesServiceのキャッシュデータクリア
    const props = PropertiesService.getScriptProperties();
    const properties = props.getProperties();
    Object.keys(properties).forEach(key => {
      if (key.startsWith('CACHE_')) {
        props.deleteProperty(key);
      }
    });
    
    debugLog('全キャッシュをクリアしました');
  }
  
  /**
   * キャッシュ統計を取得
   * @returns {object} キャッシュ統計
   */
  static getStats() {
    const props = PropertiesService.getScriptProperties();
    const properties = props.getProperties();
    
    let longTermCount = 0;
    let totalSize = 0;
    
    Object.keys(properties).forEach(key => {
      if (key.startsWith('CACHE_')) {
        longTermCount++;
        totalSize += properties[key].length;
      }
    });
    
    return {
      longTermCacheCount: longTermCount,
      estimatedSize: totalSize,
      lastCleared: properties.CACHE_LAST_CLEARED || 'Never'
    };
  }
}