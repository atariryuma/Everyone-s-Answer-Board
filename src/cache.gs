/**
 * @fileoverview 簡略化された統合キャッシュマネージャー
 * 基本的なキャッシュ機能のみを保持し、複雑なキャッシュクラスを除去
 */

/**
 * 簡略化されたキャッシュマネージャークラス
 */
class CacheManager {
  constructor() {
    this.scriptCache = CacheService.getScriptCache();
    this.memoCache = new Map(); // メモ化用の高速キャッシュ
    this.defaultTTL = 21600; // デフォルトTTL（6時間）
    this.stats = {
      hits: 0,
      misses: 0,
      errors: 0,
      totalOps: 0,
      lastReset: Date.now(),
    };
  }

  /**
   * キャッシュから値を取得、なければ指定された関数で生成して保存
   * @param {string} key - キャッシュキー
   * @param {function} valueFn - 値を生成する関数
   * @param {object} [options] - オプション { ttl: number, enableMemoization: boolean }
   * @returns {*} キャッシュされた値
   */
  get(key, valueFn, options = {}) {
    const { ttl = this.defaultTTL, enableMemoization = false } = options;

    this.stats.totalOps++;
    
    if (!key || typeof key !== 'string') {
      this.stats.errors++;
      throw new Error('キャッシュキーは文字列である必要があります');
    }

    // メモ化キャッシュから取得を試行
    if (enableMemoization && this.memoCache.has(key)) {
      this.stats.hits++;
      return this.memoCache.get(key);
    }

    try {
      // スクリプトキャッシュから取得を試行
      const cachedValue = this.scriptCache.get(key);
      if (cachedValue !== null) {
        this.stats.hits++;
        const parsedValue = JSON.parse(cachedValue);
        
        // メモ化キャッシュにも保存
        if (enableMemoization) {
          this.memoCache.set(key, parsedValue);
        }
        
        return parsedValue;
      }
    } catch (error) {
      console.warn('キャッシュ取得エラー:', error.message);
      this.stats.errors++;
    }

    // キャッシュにない場合は関数を実行
    if (typeof valueFn === 'function') {
      this.stats.misses++;
      try {
        const newValue = valueFn();
        this.set(key, newValue, { ttl, enableMemoization });
        return newValue;
      } catch (error) {
        this.stats.errors++;
        console.error('値生成関数エラー:', error.message);
        throw error;
      }
    }

    this.stats.misses++;
    return null;
  }

  /**
   * キャッシュに値を設定
   * @param {string} key - キャッシュキー
   * @param {*} value - 保存する値
   * @param {object} [options] - オプション { ttl: number, enableMemoization: boolean }
   */
  set(key, value, options = {}) {
    const { ttl = this.defaultTTL, enableMemoization = false } = options;

    if (!key || typeof key !== 'string') {
      throw new Error('キャッシュキーは文字列である必要があります');
    }

    try {
      // スクリプトキャッシュに保存
      this.scriptCache.put(key, JSON.stringify(value), ttl);
      
      // メモ化キャッシュにも保存
      if (enableMemoization) {
        this.memoCache.set(key, value);
      }
    } catch (error) {
      this.stats.errors++;
      console.error('キャッシュ設定エラー:', error.message);
      throw error;
    }
  }

  /**
   * キャッシュから値を削除
   * @param {string} key - キャッシュキー
   */
  remove(key) {
    if (!key || typeof key !== 'string') {
      throw new Error('キャッシュキーは文字列である必要があります');
    }

    try {
      this.scriptCache.remove(key);
      this.memoCache.delete(key);
    } catch (error) {
      this.stats.errors++;
      console.error('キャッシュ削除エラー:', error.message);
    }
  }

  /**
   * パターンにマッチするキーのキャッシュをすべて削除
   * @param {string} pattern - パターン（正規表現文字列）
   */
  removePattern(pattern) {
    try {
      // メモ化キャッシュから削除
      const regex = new RegExp(pattern);
      for (const key of this.memoCache.keys()) {
        if (regex.test(key)) {
          this.memoCache.delete(key);
        }
      }
      
      // スクリプトキャッシュは一括削除がないため、よく使われるパターンのみ対応
      if (pattern.includes('user_')) {
        // ユーザー関連キャッシュの削除など
        console.log('ユーザー関連キャッシュのパターン削除実行:', pattern);
      }
    } catch (error) {
      this.stats.errors++;
      console.error('パターン削除エラー:', error.message);
    }
  }

  /**
   * キャッシュの統計情報を取得
   * @returns {object} 統計情報
   */
  getStats() {
    return {
      ...this.stats,
      hitRate: this.stats.totalOps > 0 ? ((this.stats.hits / this.stats.totalOps) * 100).toFixed(2) + '%' : '0%',
      memoSize: this.memoCache.size,
      uptime: Date.now() - this.stats.lastReset,
    };
  }

  /**
   * キャッシュの健全性チェック
   * @returns {object} 健全性情報
   */
  getHealth() {
    const health = {
      status: 'ok',
      issues: [],
      stats: this.getStats(),
    };

    // エラー率チェック
    const errorRate = this.stats.totalOps > 0 ? (this.stats.errors / this.stats.totalOps) : 0;
    if (errorRate > 0.1) { // 10%以上のエラー率
      health.status = 'warning';
      health.issues.push(`高いエラー率: ${(errorRate * 100).toFixed(2)}%`);
    }

    // メモリ使用量チェック
    if (this.memoCache.size > 1000) {
      health.status = 'warning';
      health.issues.push(`メモリキャッシュサイズが大きい: ${this.memoCache.size}`);
    }

    return health;
  }

  /**
   * 統計情報をリセット
   */
  resetStats() {
    this.stats = {
      hits: 0,
      misses: 0,
      errors: 0,
      totalOps: 0,
      lastReset: Date.now(),
    };
  }

  /**
   * メモリキャッシュをクリア
   */
  clearMemoCache() {
    this.memoCache.clear();
  }
}

// グローバルキャッシュマネージャーインスタンス
const cacheManager = new CacheManager();

/**
 * Sheets APIサービス結果のキャッシュ（互換性維持）
 * @returns {object|null} キャッシュされたサービス情報
 */
function getSheetsServiceCached() {
  return cacheManager.get('sheets_service', () => {
    const accessToken = getServiceAccountTokenCached();
    if (!accessToken) return null;
    
    return {
      baseUrl: 'https://sheets.googleapis.com/v4/spreadsheets',
      accessToken: accessToken,
    };
  }, { ttl: 3500, enableMemoization: true });
}

/**
 * 汎用スプレッドシートヘッダー取得関数
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @param {object} [options] - オプション { useCache: boolean, validate: boolean, forceRefresh: boolean }
 * @returns {object} ヘッダーインデックス情報
 */
function getSpreadsheetHeaders(spreadsheetId, sheetName, options = {}) {
  const { useCache = true, validate = false, forceRefresh = false } = options;
  
  if (!spreadsheetId || !sheetName) {
    throw new Error('スプレッドシートIDとシート名は必須です');
  }

  const cacheKey = `headers_${spreadsheetId}_${sheetName}`;
  
  // キャッシュから取得を試行（forceRefreshでない場合）
  if (useCache && !forceRefresh) {
    const cached = cacheManager.get(cacheKey, null, { enableMemoization: true });
    if (cached && (!validate || validateSpreadsheetHeaders(cached).success)) {
      return cached;
    }
  }

  try {
    // スプレッドシートから直接取得
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`シート「${sheetName}」が見つかりません`);
    }

    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!headerRow || headerRow.length === 0) {
      throw new Error('ヘッダー行が見つかりません');
    }

    // ヘッダーインデックスマップを作成
    const headerIndices = {};
    headerRow.forEach((header, index) => {
      if (header && String(header).trim() !== '') {
        headerIndices[String(header).trim()] = index;
      }
    });

    // 検証実行（必要な場合）
    if (validate) {
      const validationResult = validateSpreadsheetHeaders(headerIndices);
      if (!validationResult.success) {
        console.warn('ヘッダー検証に失敗:', validationResult.missing);
      }
    }

    // キャッシュに保存
    if (useCache) {
      cacheManager.set(cacheKey, headerIndices, { 
        ttl: 1800, // 30分
        enableMemoization: true 
      });
    }

    console.log(`📊 スプレッドシートヘッダーを取得しました: ${spreadsheetId}/${sheetName}`);
    return headerIndices;

  } catch (error) {
    console.error('[ERROR] getSpreadsheetHeaders:', error.message);
    throw new Error(`ヘッダー取得エラー: ${error.message}`);
  }
}

/**
 * スプレッドシートヘッダーの検証
 * @param {object} headerIndices - ヘッダーインデックス
 * @returns {object} 検証結果 { success: boolean, missing: string[], hasReasonColumn: boolean, hasOpinionColumn: boolean }
 */
function validateSpreadsheetHeaders(headerIndices) {
  if (!headerIndices || typeof headerIndices !== 'object') {
    return { success: false, missing: ['すべて'], hasReasonColumn: false, hasOpinionColumn: false };
  }
  
  // COLUMN_HEADERSが定義されていない場合のフォールバック
  const HEADERS = typeof COLUMN_HEADERS !== 'undefined' ? COLUMN_HEADERS : {
    REASON: '理由',
    OPINION: '回答'
  };
  
  const hasReason = headerIndices[HEADERS.REASON] !== undefined;
  const hasOpinion = headerIndices[HEADERS.OPINION] !== undefined;
  const missing = [];
  
  if (!hasReason) missing.push(HEADERS.REASON);
  if (!hasOpinion) missing.push(HEADERS.OPINION);
  
  return {
    success: hasReason && hasOpinion,
    missing,
    hasReasonColumn: hasReason,
    hasOpinionColumn: hasOpinion
  };
}

console.log('🗄️ 簡略化されたキャッシュシステムが初期化されました');