/**
 * SpreadsheetApp.openById()呼び出し最適化システム
 * メモリキャッシュとセッションキャッシュを組み合わせた高速化
 */

// メモリキャッシュ - 実行セッション内で有効
let spreadsheetMemoryCache = {};

// キャッシュ設定
const SPREADSHEET_CACHE_CONFIG = {
  MEMORY_CACHE_TTL: 300000,    // 5分間（メモリキャッシュ）
  SESSION_CACHE_TTL: 1800000,  // 30分間（PropertiesServiceキャッシュ）
  MAX_CACHE_SIZE: 50,          // 最大キャッシュエントリ数
  CACHE_KEY_PREFIX: 'ss_cache_'
};

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

  const cacheKey = `${SPREADSHEET_CACHE_CONFIG.CACHE_KEY_PREFIX}${spreadsheetId}`;
  const now = Date.now();

  // 強制リフレッシュの場合はキャッシュをクリア
  if (forceRefresh) {
    delete spreadsheetMemoryCache[spreadsheetId];
    try {
      PropertiesService.getScriptProperties().deleteProperty(cacheKey);
    } catch (error) {
      debugLog('キャッシュクリアエラー:', error.message);
    }
  }

  // Phase 1: メモリキャッシュをチェック
  const memoryEntry = spreadsheetMemoryCache[spreadsheetId];
  if (memoryEntry && (now - memoryEntry.timestamp) < SPREADSHEET_CACHE_CONFIG.MEMORY_CACHE_TTL) {
    debugLog('✅ SpreadsheetApp.openById メモリキャッシュヒット:', spreadsheetId.substring(0, 10));
    return memoryEntry.spreadsheet;
  }

  // Phase 2: セッションキャッシュをチェック
  try {
    const sessionCacheData = PropertiesService.getScriptProperties().getProperty(cacheKey);
    if (sessionCacheData) {
      const sessionEntry = JSON.parse(sessionCacheData);
      if ((now - sessionEntry.timestamp) < SPREADSHEET_CACHE_CONFIG.SESSION_CACHE_TTL) {
        // セッションキャッシュからSpreadsheetオブジェクトを再構築
        const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
        
        // メモリキャッシュに保存
        spreadsheetMemoryCache[spreadsheetId] = {
          spreadsheet: spreadsheet,
          timestamp: now
        };
        
        debugLog('✅ SpreadsheetApp.openById セッションキャッシュヒット:', spreadsheetId.substring(0, 10));
        return spreadsheet;
      }
    }
  } catch (error) {
    debugLog('セッションキャッシュ読み込みエラー:', error.message);
  }

  // Phase 3: 新規取得とキャッシュ保存
  debugLog('🔄 SpreadsheetApp.openById 新規取得:', spreadsheetId.substring(0, 10));
  
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    
    // メモリキャッシュに保存
    spreadsheetMemoryCache[spreadsheetId] = {
      spreadsheet: spreadsheet,
      timestamp: now
    };
    
    // セッションキャッシュに保存（軽量データのみ）
    try {
      const sessionData = {
        spreadsheetId: spreadsheetId,
        timestamp: now,
        url: spreadsheet.getUrl()
      };
      
      PropertiesService.getScriptProperties().setProperty(
        cacheKey, 
        JSON.stringify(sessionData)
      );
    } catch (sessionError) {
      debugLog('セッションキャッシュ保存エラー:', sessionError.message);
    }
    
    // キャッシュサイズ管理
    cleanupOldCacheEntries();
    
    return spreadsheet;
    
  } catch (error) {
    logError(error, 'SpreadsheetApp.openById', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.DATABASE, { spreadsheetId });
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
  if (memoryKeys.length > SPREADSHEET_CACHE_CONFIG.MAX_CACHE_SIZE) {
    const sortedEntries = memoryKeys.map(key => ({
      key: key,
      timestamp: spreadsheetMemoryCache[key].timestamp
    })).sort((a, b) => b.timestamp - a.timestamp);
    
    // 古いエントリを削除
    const entriesToDelete = sortedEntries.slice(SPREADSHEET_CACHE_CONFIG.MAX_CACHE_SIZE);
    entriesToDelete.forEach(entry => {
      delete spreadsheetMemoryCache[entry.key];
    });
    
    debugLog(`🧹 メモリキャッシュクリーンアップ: ${entriesToDelete.length}件削除`);
  }
  
  // 期限切れエントリの削除
  memoryKeys.forEach(key => {
    const entry = spreadsheetMemoryCache[key];
    if ((now - entry.timestamp) > SPREADSHEET_CACHE_CONFIG.MEMORY_CACHE_TTL) {
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
  const cacheKey = `${SPREADSHEET_CACHE_CONFIG.CACHE_KEY_PREFIX}${spreadsheetId}`;
  try {
    PropertiesService.getScriptProperties().deleteProperty(cacheKey);
    debugLog('🗑️ SpreadsheetCache無効化:', spreadsheetId.substring(0, 10));
  } catch (error) {
    debugLog('キャッシュ無効化エラー:', error.message);
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
    
    Object.keys(allProps).forEach(key => {
      if (key.startsWith(SPREADSHEET_CACHE_CONFIG.CACHE_KEY_PREFIX)) {
        props.deleteProperty(key);
      }
    });
    
    debugLog('🧹 全SpreadsheetCacheクリア完了');
  } catch (error) {
    debugLog('全キャッシュクリアエラー:', error.message);
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
    sessionEntries = Object.keys(allProps).filter(key => 
      key.startsWith(SPREADSHEET_CACHE_CONFIG.CACHE_KEY_PREFIX)
    ).length;
  } catch (error) {
    debugLog('キャッシュ統計取得エラー:', error.message);
  }
  
  return {
    memoryEntries: memoryEntries,
    sessionEntries: sessionEntries,
    maxCacheSize: SPREADSHEET_CACHE_CONFIG.MAX_CACHE_SIZE,
    memoryTTL: SPREADSHEET_CACHE_CONFIG.MEMORY_CACHE_TTL,
    sessionTTL: SPREADSHEET_CACHE_CONFIG.SESSION_CACHE_TTL
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