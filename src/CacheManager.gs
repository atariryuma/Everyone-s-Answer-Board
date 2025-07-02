/**
 * @fileoverview 統合キャッシュ管理システム - GAS互換版
 */

// キャッシュ管理の定数
var CACHE_SHORT_TTL = 300;    // 5分 - CacheService
var CACHE_MEDIUM_TTL = 1800;  // 30分 - CacheService
var CACHE_LONG_TTL = 86400;   // 24時間 - PropertiesService

/**
 * 階層キャッシュから値を取得
 * @param {string} key - キャッシュキー
 * @param {function} fetchFunction - データ取得関数
 * @param {string} level - キャッシュレベル ('short', 'medium', 'long')
 * @returns {any} キャッシュされた値またはfetchFunctionの結果
 */
function getCachedValue(key, fetchFunction, level) {
  level = level || 'short';
  
  // 1. CacheServiceから取得を試行
  var cache = CacheService.getScriptCache();
  var cachedValue = cache.get(key);
  
  if (cachedValue !== null) {
    debugLog('CacheServiceヒット: ' + key);
    try {
      return JSON.parse(cachedValue);
    } catch (e) {
      return cachedValue;
    }
  }
  
  // 2. 長期キャッシュの場合はPropertiesServiceを確認
  if (level === 'long') {
    var props = PropertiesService.getScriptProperties();
    var propKey = 'CACHE_' + key;
    var propValue = props.getProperty(propKey);
    
    if (propValue) {
      try {
        var data = JSON.parse(propValue);
        var now = Date.now();
        
        if (data.expiresAt > now) {
          debugLog('PropertiesServiceヒット: ' + key);
          // CacheServiceにも保存して次回高速化
          var remainingTtl = Math.min(CACHE_SHORT_TTL, Math.floor((data.expiresAt - now) / 1000));
          cache.put(key, JSON.stringify(data.value), remainingTtl);
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
  var value = fetchFunction();
  
  if (value !== null && value !== undefined) {
    setCachedValue(key, value, level);
  }
  
  return value;
}

/**
 * 階層キャッシュに値を保存
 * @param {string} key - キャッシュキー
 * @param {any} value - 保存する値
 * @param {string} level - キャッシュレベル
 */
function setCachedValue(key, value, level) {
  level = level || 'short';
  
  try {
    var cache = CacheService.getScriptCache();
    var ttl;
    
    switch (level) {
      case 'short':
        ttl = CACHE_SHORT_TTL;
        break;
      case 'medium':
        ttl = CACHE_MEDIUM_TTL;
        break;
      case 'long':
        ttl = CACHE_LONG_TTL;
        break;
      default:
        ttl = CACHE_SHORT_TTL;
    }
    
    var serializedValue = typeof value === 'string' ? value : JSON.stringify(value);
    
    // CacheServiceに保存
    cache.put(key, serializedValue, Math.min(ttl, 21600)); // CacheServiceの最大6時間制限
    
    // 長期キャッシュの場合はPropertiesServiceにも保存
    if (level === 'long') {
      var props = PropertiesService.getScriptProperties();
      var expiresAt = Date.now() + (ttl * 1000);
      var data = {
        value: value,
        expiresAt: expiresAt
      };
      
      var propKey = 'CACHE_' + key;
      try {
        props.setProperty(propKey, JSON.stringify(data));
      } catch (e) {
        // PropertiesServiceの容量制限に達した場合の処理
        console.warn('PropertiesService保存失敗: ' + e.message);
      }
    }
    
    debugLog('キャッシュ保存 (' + level + '): ' + key);
  } catch (error) {
    console.error('キャッシュ保存エラー: ' + error.message);
  }
}

/**
 * キャッシュを削除
 * @param {string} key - キャッシュキー
 */
function removeCachedValue(key) {
  var cache = CacheService.getScriptCache();
  cache.remove(key);
  
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty('CACHE_' + key);
  
  debugLog('キャッシュ削除: ' + key);
}

/**
 * パターンマッチでキャッシュを削除
 * @param {string} pattern - 削除パターン
 */
function removeCachedPattern(pattern) {
  var props = PropertiesService.getScriptProperties();
  
  // PropertiesServiceからパターンマッチするキーを削除
  var properties = props.getProperties();
  Object.keys(properties).forEach(function(key) {
    if (key.indexOf('CACHE_') === 0 && key.indexOf(pattern) !== -1) {
      props.deleteProperty(key);
      var cacheKey = key.replace('CACHE_', '');
      var cache = CacheService.getScriptCache();
      cache.remove(cacheKey);
      debugLog('パターン削除: ' + cacheKey);
    }
  });
}

/**
 * ヘッダー情報をキャッシュ
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {object} ヘッダーインデックス
 */
function getHeaderIndicesCached(spreadsheetId, sheetName) {
  var key = 'headers_' + spreadsheetId + '_' + sheetName;
  
  return getCachedValue(key, function() {
    var service = getOptimizedSheetsService();
    var response = batchGetSheetsData(service, spreadsheetId, [sheetName + '!1:1']);
    var headers = response.valueRanges[0].values ? response.valueRanges[0].values[0] : [];
    
    var indices = {};
    Object.keys(COLUMN_HEADERS).forEach(function(columnKey) {
      var header = COLUMN_HEADERS[columnKey];
      var index = headers.indexOf(header);
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
function getRosterMapCached(spreadsheetId) {
  var key = 'roster_' + spreadsheetId;
  
  return getCachedValue(key, function() {
    try {
      var service = getOptimizedSheetsService();
      var response = batchGetSheetsData(service, spreadsheetId, [ROSTER_CONFIG.SHEET_NAME + '!A:Z']);
      var values = response.valueRanges[0].values || [];
      
      if (values.length === 0) return {};
      
      var headers = values[0];
      var emailIndex = headers.indexOf(ROSTER_CONFIG.EMAIL_COLUMN);
      var nameIndex = headers.indexOf(ROSTER_CONFIG.NAME_COLUMN);
      var classIndex = headers.indexOf(ROSTER_CONFIG.CLASS_COLUMN);
      
      var rosterMap = {};
      
      for (var i = 1; i < values.length; i++) {
        var row = values[i];
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
function clearAllCache() {
  // PropertiesServiceのキャッシュデータクリア
  var props = PropertiesService.getScriptProperties();
  var properties = props.getProperties();
  Object.keys(properties).forEach(function(key) {
    if (key.indexOf('CACHE_') === 0) {
      props.deleteProperty(key);
    }
  });
  
  // CacheServiceのキャッシュデータクリア
  CacheService.getScriptCache().removeAll();

  debugLog('全キャッシュをクリアしました');
}

/**
 * キャッシュ統計を取得
 * @returns {object} キャッシュ統計
 */
function getCacheStats() {
  var props = PropertiesService.getScriptProperties();
  var properties = props.getProperties();
  
  var longTermCount = 0;
  var totalSize = 0;
  
  Object.keys(properties).forEach(function(key) {
    if (key.indexOf('CACHE_') === 0) {
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