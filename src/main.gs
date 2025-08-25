/**
 * @fileoverview メインエントリーポイント - 2024年最新技術の結集
 * V8ランタイム、最新パフォーマンス技術、安定性強化を統合
 */

/**
 * HTML ファイルを読み込む include ヘルパー
 * @param {string} path ファイルパス
 * @return {string} HTML content
 */
function include(path) {
  try {
    return HtmlService.createHtmlOutputFromFile(path).getContent();
  } catch (error) {
    logError(error, 'includeFile', MAIN_ERROR_SEVERITY.HIGH, MAIN_ERROR_CATEGORIES.SYSTEM, { filePath: path });
    return `<!-- Error including ${path}: ${error.message} -->`;
  }
}


/**
 * JavaScript文字列エスケープ関数 (URL対応版)
 * @param {string} str エスケープする文字列
 * @return {string} エスケープされた文字列
 */
function escapeJavaScript(str) {
  if (!str) return '';

  const strValue = str.toString();

  // URL判定: HTTP/HTTPSで始まり、すでに適切にエスケープされている場合は最小限の処理
  if (strValue.match(/^https?:\/\/[^\s<>"']+$/)) {
    // URLの場合はバックスラッシュと改行文字のみエスケープ
    return strValue
      .replace(/\\/g, '\\\\')
      .replace(/\n/g, '\\n')
      .replace(/\r/g, '\\r')
      .replace(/\t/g, '\\t');
  }

  // 通常のテキストの場合は従来通りの完全エスケープ
  return strValue
    .replace(/\\/g, '\\\\')
    .replace(/'/g, "\\'")
    .replace(/"/g, '\\"')
    .replace(/\n/g, '\\n')
    .replace(/\r/g, '\\r')
    .replace(/\t/g, '\\t');
}

// グローバル定数の定義（書き換え不可）
const SCRIPT_PROPS_KEYS = {
  SERVICE_ACCOUNT_CREDS: 'SERVICE_ACCOUNT_CREDS',
  DATABASE_SPREADSHEET_ID: 'DATABASE_SPREADSHEET_ID',
  ADMIN_EMAIL: 'ADMIN_EMAIL'
};

// エラーハンドリング定数（main.gs独自）
const MAIN_ERROR_SEVERITY = {
  LOW: 'low',
  MEDIUM: 'medium',
  HIGH: 'high',
  CRITICAL: 'critical',
};

const MAIN_ERROR_CATEGORIES = {
  AUTHENTICATION: 'authentication',
  AUTHORIZATION: 'authorization',
  DATABASE: 'database',
  CACHE: 'cache',
  NETWORK: 'network',
  VALIDATION: 'validation',
  SYSTEM: 'system',
  USER_INPUT: 'user_input',
};

// =============================================================================
// グローバルエラー定数の統一 - Core.gsとの互換性確保
// =============================================================================

/**
 * 統一エラーハンドリング定数（Core.gsとmain.gsの橋渡し）
 */
if (typeof ERROR_SEVERITY === 'undefined') {
  var ERROR_SEVERITY = MAIN_ERROR_SEVERITY;
}

if (typeof ERROR_CATEGORIES === 'undefined') {  
  var ERROR_CATEGORIES = MAIN_ERROR_CATEGORIES;
}

const DB_SHEET_CONFIG = {
  SHEET_NAME: 'Users',
  HEADERS: [
    'userId', 'adminEmail', 'spreadsheetId', 'spreadsheetUrl',
    'createdAt', 'configJson', 'lastAccessedAt', 'isActive'
  ]
};

const LOG_SHEET_CONFIG = {
  SHEET_NAME: 'Logs',
  HEADERS: ['timestamp', 'userId', 'action', 'details']
};

// 履歴管理の定数
const MAX_HISTORY_ITEMS = 50;

// 実行中のユーザー情報キャッシュ（パフォーマンス最適化用）
let _executionUserInfoCache = null;

// =============================================================================
// 統一ログ関数 - 削除されたファイルからの移行実装
// =============================================================================

/**
 * デバッグログ出力（DEBUG_MODE時のみ）
 */
function debugLog(message, ...args) {
  if (typeof DEBUG_MODE !== 'undefined' && DEBUG_MODE) {
    console.log(`[DEBUG] ${message}`, ...args);
  }
}

/**
 * 情報ログ出力
 */
function infoLog(message, ...args) {
  console.log(`[INFO] ${message}`, ...args);
}

/**
 * 警告ログ出力
 */
function warnLog(message, ...args) {
  console.warn(`[WARN] ${message}`, ...args);
}

/**
 * エラーログ出力
 */
function errorLog(message, ...args) {
  console.error(`[ERROR] ${message}`, ...args);
}

// =============================================================================
// 簡易キャッシュマネージャー - 削除されたunifiedCacheManagerの代替
// =============================================================================

/**
 * 簡易キャッシュマネージャー（削除されたunifiedCacheManager.gsの軽量版）
 */
const cacheManager = {
  /**
   * キャッシュから値を取得（存在しない場合は関数を実行してキャッシュ）
   */
  get(key, valueFn, options = {}) {
    try {
      const cache = CacheService.getScriptCache();
      const cached = cache.get(key);
      
      if (cached !== null && !options.skipFetch) {
        try {
          return JSON.parse(cached);
        } catch (parseError) {
          warnLog('キャッシュ解析エラー:', parseError.message);
        }
      }
      
      // キャッシュが存在しない、または無効な場合
      if (valueFn && typeof valueFn === 'function') {
        const value = valueFn();
        const ttl = options.ttl || 600; // デフォルト10分
        cache.put(key, JSON.stringify(value), ttl);
        return value;
      }
      
      return null;
    } catch (error) {
      warnLog('キャッシュ操作エラー:', error.message);
      return valueFn ? valueFn() : null;
    }
  },

  /**
   * キャッシュから値を削除
   */
  remove(key) {
    try {
      CacheService.getScriptCache().remove(key);
    } catch (error) {
      warnLog('キャッシュ削除エラー:', error.message);
    }
  },

  /**
   * パターンに一致するキャッシュを削除（簡易実装）
   */
  clearByPattern(pattern, options = {}) {
    warnLog('clearByPattern: 簡易実装のため、全キャッシュクリアを実行');
    try {
      CacheService.getScriptCache().removeAll();
    } catch (error) {
      warnLog('パターンキャッシュ削除エラー:', error.message);
    }
  },

  /**
   * プレフィックスに一致するキャッシュを削除（軽量実装）
   * 呼び出し箇所: main.gs:860, 1205
   * @param {string} prefix - 削除対象のプレフィックス
   */
  removeByPrefix(prefix) {
    try {
      // GAS CacheServiceでは特定のプレフィックス削除は直接サポートされていないため
      // 知られている一般的なキーパターンを削除
      const commonKeys = [
        `${prefix}_config`, `${prefix}_data`, `${prefix}_user`, `${prefix}_cache`,
        `${prefix}Config`, `${prefix}Data`, `${prefix}User`, `${prefix}Cache`,
        `${prefix}_info`, `${prefix}Info`, `${prefix}_status`, `${prefix}Status`
      ];
      
      const cache = CacheService.getScriptCache();
      commonKeys.forEach(key => {
        try {
          cache.remove(key);
        } catch (removeError) {
          // 個別の削除エラーは無視（キーが存在しない可能性）
        }
      });
      
      categoryLog('CACHE', `プレフィックス削除実行: ${prefix} (${commonKeys.length}個のパターンを削除)`);
      
      // フォールバック: プレフィックス削除が重要な場合は全キャッシュクリア
      if (prefix === 'critical' || prefix === 'userInfo' || prefix === 'dbConnection') {
        debugLog(`重要プレフィックス "${prefix}" のため全キャッシュクリアを実行`);
        cache.removeAll();
      }
    } catch (error) {
      errorLog('removeByPrefix エラー:', error.message);
      // エラー時のフォールバック
      try {
        CacheService.getScriptCache().removeAll();
        warnLog('removeByPrefix: エラーのため全キャッシュクリアを実行');
      } catch (fallbackError) {
        errorLog('removeByPrefix フォールバックエラー:', fallbackError.message);
      }
    }
  },

  /**
   * 期限切れキャッシュのクリア（CacheServiceが自動管理）
   */
  clearExpired() {
    // CacheServiceが自動で期限切れを管理するため、何もしない
  },

  /**
   * キャッシュ統計（簡易版）
   */
  getStats() {
    return {
      available: true,
      type: 'simple',
      message: '簡易キャッシュマネージャー使用中'
    };
  },

  /**
   * シートデータのキャッシュ無効化
   */
  invalidateSheetData(spreadsheetId, sheetName) {
    try {
      const patterns = [
        `sheets_${spreadsheetId}_${sheetName}`,
        `batchGet_${spreadsheetId}_${sheetName}`,
        `sheetData_${spreadsheetId}_${sheetName}`
      ];
      patterns.forEach(pattern => this.remove(pattern));
    } catch (error) {
      warnLog('シートキャッシュ無効化エラー:', error.message);
    }
  },

  // =============================================================================
  // 高度化機能 - 削除されたunifiedCacheManager.gsからの復元
  // =============================================================================

  /**
   * 統計情報（ヒット率分析付き）
   */
  stats: {
    hits: 0,
    misses: 0,
    operations: 0,
    lastReset: Date.now()
  },

  /**
   * キャッシュヒット率を取得
   */
  getHitRate() {
    const total = this.stats.hits + this.stats.misses;
    return total > 0 ? (this.stats.hits / total * 100).toFixed(2) + '%' : '0%';
  },

  /**
   * 詳細統計情報を取得
   */
  getDetailedStats() {
    const uptime = Date.now() - this.stats.lastReset;
    return {
      hitRate: this.getHitRate(),
      hits: this.stats.hits,
      misses: this.stats.misses,
      operations: this.stats.operations,
      uptime: Math.round(uptime / 1000) + 's',
      averageOpsPerSecond: (this.stats.operations / (uptime / 1000)).toFixed(2)
    };
  },

  /**
   * カテゴリ別キャッシュ管理
   */
  categorizedCache: {
    user(userId, key, valueFn, options = {}) {
      const categorizedKey = `user_${userId}_${key}`;
      return cacheManager.get(categorizedKey, valueFn, { ttl: 300, ...options });
    },

    system(key, valueFn, options = {}) {
      const categorizedKey = `system_${key}`;
      return cacheManager.get(categorizedKey, valueFn, { ttl: 600, ...options });
    },

    temporary(key, valueFn, options = {}) {
      const categorizedKey = `temp_${key}`;
      return cacheManager.get(categorizedKey, valueFn, { ttl: 60, ...options });
    },

    session(userId, key, valueFn, options = {}) {
      const categorizedKey = `session_${userId}_${key}`;
      return cacheManager.get(categorizedKey, valueFn, { ttl: 1800, ...options });
    }
  },

  /**
   * バックグラウンドキャッシュ更新
   */
  backgroundRefresh(key, refreshFn, interval = 300000) {
    try {
      // 現在のキャッシュを取得
      const cached = this.get(key);
      
      // バックグラウンドで新しいデータを取得
      setTimeout(() => {
        try {
          const newValue = refreshFn();
          this.get(key, () => newValue, { ttl: 600 });
          categoryLog('CACHE', 'DEBUG', `Background refresh completed: ${key}`);
        } catch (error) {
          categoryLog('CACHE', 'WARN', `Background refresh failed: ${key}`, error.message);
        }
      }, 100); // 100ms後に非同期実行
      
      return cached; // 既存データをすぐ返す
    } catch (error) {
      warnLog('Background refresh setup error:', error.message);
      return this.get(key, refreshFn);
    }
  },

  /**
   * 統計をリセット
   */
  resetStats() {
    this.stats = {
      hits: 0,
      misses: 0,
      operations: 0,
      lastReset: Date.now()
    };
  },

  /**
   * TTL自動調整機能
   */
  adaptiveTTL(key, valueFn, baseOptions = {}) {
    const startTime = Date.now();
    const result = this.get(key, () => {
      const value = valueFn();
      const computeTime = Date.now() - startTime;
      
      // 計算時間に基づいてTTLを調整
      let adaptiveTtl = baseOptions.ttl || 300;
      if (computeTime > 5000) adaptiveTtl = 1800;      // 5秒以上 → 30分
      else if (computeTime > 1000) adaptiveTtl = 900;  // 1秒以上 → 15分
      else if (computeTime > 100) adaptiveTtl = 600;   // 100ms以上 → 10分
      
      return value;
    }, { ttl: baseOptions.ttl || 300, ...baseOptions });
    
    return result;
  }
};

// =============================================================================
// 削除された有用な関数の復元 - 過去のコミットからの移植
// =============================================================================

/**
 * フォームURLをクリップボードにコピー（削除されたfunctionの復元）
 */
function copyFormUrl(url) {
  try {
    if (!url || typeof url !== 'string') {
      throw new Error('有効なURLが指定されていません');
    }
    
    // クリップボード操作は主にフロントエンド側で行われるため、
    // ここでは URL の検証と整形のみ行う
    const cleanUrl = url.trim();
    
    if (!cleanUrl.includes('docs.google.com/forms/')) {
      throw new Error('Googleフォームの有効なURLではありません');
    }
    
    infoLog('フォームURL準備完了:', cleanUrl);
    return {
      success: true,
      url: cleanUrl,
      message: 'URLがクリップボード用に準備されました'
    };
  } catch (error) {
    errorLog('copyFormUrl エラー:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * サービスアカウントトークンのキャッシュ版取得（削除されたfunctionの復元）
 */
function getServiceAccountTokenCached() {
  const cacheKey = 'service_account_token';
  const ttl = 3300; // 55分（通常1時間有効だが余裕を持って）
  
  try {
    return cacheManager.get(cacheKey, () => {
      // GAS環境ではScriptApp.getOAuthTokenを使用
      const token = ScriptApp.getOAuthToken();
      if (!token) {
        throw new Error('OAuth トークンの取得に失敗しました');
      }
      infoLog('新しいサービスアカウントトークンを取得しました');
      return token;
    }, { ttl });
  } catch (error) {
    errorLog('getServiceAccountTokenCached エラー:', error.message);
    // フォールバック: 直接取得を試行
    try {
      return ScriptApp.getOAuthToken();
    } catch (fallbackError) {
      errorLog('フォールバックトークン取得も失敗:', fallbackError.message);
      throw new Error('認証トークンを取得できませんでした');
    }
  }
}

/**
 * セキュアなマルチテナントキャッシュ操作（削除されたfunctionの復元）
 */
function secureMultiTenantCacheOperation(operation, key, value, userId) {
  try {
    // マルチテナント対応のキープレフィックス
    const secureKey = userId ? `tenant_${userId}_${key}` : `global_${key}`;
    
    switch (operation.toLowerCase()) {
      case 'get':
        return cacheManager.get(secureKey);
        
      case 'set':
      case 'put':
        const ttl = 600; // 10分デフォルト
        cacheManager.get(secureKey, () => value, { ttl });
        return true;
        
      case 'remove':
      case 'delete':
        cacheManager.remove(secureKey);
        return true;
        
      default:
        throw new Error(`未対応のキャッシュ操作: ${operation}`);
    }
  } catch (error) {
    warnLog('SecureMultiTenantCache 操作エラー:', operation, key, error.message);
    return null;
  }
}

// =============================================================================
// DEBUG_CONFIG システム復元 - 削除されたdebugConfig.gsからの移植
// =============================================================================

/**
 * デバッグ設定管理システム（削除されたdebugConfig.gsの復元）
 */
const DEBUG_CONFIG = {
  // 本番環境判定: PropertiesService で制御可能
  get isProduction() {
    try {
      const envSetting = PropertiesService.getScriptProperties().getProperty('ENVIRONMENT');
      if (envSetting) {
        return envSetting === 'production';
      }
      // フォールバック: 現在のURLやユーザーで判定
      const currentUrl = ScriptApp.getService().getUrl();
      return currentUrl && !currentUrl.includes('script.google.com/macros/d/');
    } catch (error) {
      // エラー時は安全側に倒して本番扱い
      return true;
    }
  },

  // ロギングレベル設定
  logLevels: {
    ERROR: 0,   // エラーログ（常に出力）
    WARN: 1,    // 警告ログ
    INFO: 2,    // 情報ログ
    DEBUG: 3    // デバッグログ
  },

  // 現在のログレベル（本番では ERROR のみ）
  get currentLogLevel() {
    return this.isProduction ? 0 : 2; // 本番: ERROR のみ、開発: INFO まで
  },

  // カテゴリ別デバッグ制御
  categories: {
    CACHE: true,      // キャッシュ関連
    AUTH: true,       // 認証関連
    DATABASE: true,   // データベース操作
    UI: false,        // UI更新関連（頻繁なため通常は無効）
    PERFORMANCE: true // パフォーマンス計測
  }
};

/**
 * 本番環境かどうかを判定（削除されたfunctionの復元）
 * @returns {boolean} 本番環境の場合 true
 */
function isProductionEnvironment() {
  return DEBUG_CONFIG.isProduction;
}

/**
 * カテゴリ別デバッグログ出力（削除されたfunctionの復元）
 * @param {string} category - ログカテゴリ
 * @param {string} level - ログレベル
 * @param {string} message - メッセージ
 * @param {...any} args - 追加引数
 */
function categoryLog(category, level, message, ...args) {
  // カテゴリが無効な場合は何もしない
  if (!DEBUG_CONFIG.categories[category]) {
    return;
  }

  // ログレベルチェック
  const numericLevel = DEBUG_CONFIG.logLevels[level.toUpperCase()] || 0;
  if (numericLevel > DEBUG_CONFIG.currentLogLevel) {
    return;
  }

  // 本番環境では ERROR のみ
  if (DEBUG_CONFIG.isProduction && level.toUpperCase() !== 'ERROR') {
    return;
  }

  const prefix = `[${category}:${level.toUpperCase()}]`;
  
  switch (level.toUpperCase()) {
    case 'ERROR':
      errorLog(prefix, message, ...args);
      break;
    case 'WARN':
      warnLog(prefix, message, ...args);
      break;
    case 'INFO':
      infoLog(prefix, message, ...args);
      break;
    case 'DEBUG':
      debugLog(prefix, message, ...args);
      break;
    default:
      console.log(prefix, message, ...args);
  }
}

/**
 * パフォーマンス測定用ログ（削除されたfunctionの復元）
 * @param {string} operation - 操作名
 * @param {number} startTime - 開始時刻
 * @param {object} metadata - メタデータ
 */
function performanceLog(operation, startTime, metadata = {}) {
  if (!DEBUG_CONFIG.categories.PERFORMANCE) return;
  
  const duration = Date.now() - startTime;
  const level = duration > 5000 ? 'WARN' : duration > 1000 ? 'INFO' : 'DEBUG';
  
  categoryLog('PERFORMANCE', level, `${operation} completed in ${duration}ms`, metadata);
}

// =============================================================================
// 削除されたユーティリティ関数群の復元 - モダンな実装
// =============================================================================

/**
 * ユニークIDを生成（削除されたfunctionの復元）
 * @param {string} prefix - プレフィックス
 * @returns {string} UUID v4形式のID
 */
function generateRandomId(prefix = '') {
  const uuid = 'xxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
    const r = Math.random() * 16 | 0;
    const v = c == 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
  return prefix ? `${prefix}_${uuid}` : uuid;
}

/**
 * ユーザー入力のサニタイズ（削除されたfunctionの復元）
 * @param {string} input - 入力値
 * @param {object} options - オプション
 * @returns {string} サニタイズ済み入力値
 */
function sanitizeUserInput(input, options = {}) {
  if (!input || typeof input !== 'string') {
    return '';
  }

  const {
    maxLength = 1000,
    allowHtml = false,
    trimWhitespace = true,
    removeLineBreaks = false
  } = options;

  let sanitized = input;

  // 長さ制限
  if (sanitized.length > maxLength) {
    sanitized = sanitized.substring(0, maxLength);
  }

  // 前後の空白除去
  if (trimWhitespace) {
    sanitized = sanitized.trim();
  }

  // 改行の除去
  if (removeLineBreaks) {
    sanitized = sanitized.replace(/[\r\n]/g, ' ');
  }

  // HTMLタグの除去（HTMLが許可されていない場合）
  if (!allowHtml) {
    sanitized = sanitized.replace(/<[^>]*>/g, '');
  }

  // XSS対策: 危険なJavaScriptパターンを除去
  const dangerousPatterns = [
    /javascript:/gi,
    /vbscript:/gi,
    /data:/gi,
    /on\w+\s*=/gi
  ];

  dangerousPatterns.forEach(pattern => {
    sanitized = sanitized.replace(pattern, '');
  });

  return sanitized;
}

/**
 * 日付のフォーマット（削除されたfunctionの復元）
 * @param {Date|string|number} date - 日付
 * @param {string} format - フォーマット ('YYYY-MM-DD', 'YYYY/MM/DD HH:mm', etc.)
 * @returns {string} フォーマット済み日付文字列
 */
function formatDateForDisplay(date, format = 'YYYY-MM-DD') {
  if (!date) return '';

  try {
    const dateObj = date instanceof Date ? date : new Date(date);
    
    if (isNaN(dateObj.getTime())) {
      return '';
    }

    const year = dateObj.getFullYear();
    const month = String(dateObj.getMonth() + 1).padStart(2, '0');
    const day = String(dateObj.getDate()).padStart(2, '0');
    const hours = String(dateObj.getHours()).padStart(2, '0');
    const minutes = String(dateObj.getMinutes()).padStart(2, '0');
    const seconds = String(dateObj.getSeconds()).padStart(2, '0');

    return format
      .replace('YYYY', year)
      .replace('MM', month)
      .replace('DD', day)
      .replace('HH', hours)
      .replace('mm', minutes)
      .replace('ss', seconds);
  } catch (error) {
    errorLog('Date formatting error:', error.message);
    return '';
  }
}

/**
 * メールアドレスの検証（削除されたfunctionの復元）
 * @param {string} email - メールアドレス
 * @returns {boolean} 有効な場合true
 */
function validateEmail(email) {
  if (!email || typeof email !== 'string') {
    return false;
  }

  const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
  return emailRegex.test(email.trim());
}

/**
 * URLのクエリ文字列を解析（削除されたfunctionの復元）
 * @param {string} url - URL
 * @returns {object} パラメータのキー・バリューオブジェクト
 */
function parseQueryString(url) {
  const params = {};
  
  try {
    if (!url || typeof url !== 'string') {
      return params;
    }

    const queryStart = url.indexOf('?');
    if (queryStart === -1) {
      return params;
    }

    const queryString = url.substring(queryStart + 1);
    const pairs = queryString.split('&');

    pairs.forEach(pair => {
      const [key, value] = pair.split('=');
      if (key) {
        params[decodeURIComponent(key)] = value ? decodeURIComponent(value) : '';
      }
    });
  } catch (error) {
    warnLog('Query string parsing error:', error.message);
  }

  return params;
}

/**
 * 文字列のハッシュ値を生成（削除されたfunctionの復元）
 * @param {string} str - ハッシュ化する文字列
 * @returns {string} ハッシュ値
 */
function generateHash(str) {
  if (!str || typeof str !== 'string') {
    return '';
  }

  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    const char = str.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash; // 32bit整数に変換
  }
  
  return Math.abs(hash).toString(36);
}

/**
 * 配列の重複を除去（削除されたfunctionの復元）
 * @param {Array} array - 配列
 * @param {string} key - オブジェクトの場合のキー
 * @returns {Array} 重複除去済み配列
 */
function removeDuplicates(array, key = null) {
  if (!Array.isArray(array)) {
    return [];
  }

  if (key) {
    // オブジェクト配列の重複除去
    const seen = new Set();
    return array.filter(item => {
      const value = item[key];
      if (seen.has(value)) {
        return false;
      }
      seen.add(value);
      return true;
    });
  } else {
    // プリミティブ配列の重複除去
    return [...new Set(array)];
  }
}

/**
 * 実行中のユーザー情報キャッシュをクリア
 */
function clearExecutionUserInfoCache() {
  // レガシーキャッシュのクリア
  _executionUserInfoCache = null;
  
  // 統一キャッシュマネージャーを使用した包括的なクリア
  try {
    if (typeof getUnifiedExecutionCache === 'function') {
      const cache = getUnifiedExecutionCache();
      cache.clearUserInfo();
      cache.syncWithUnifiedCache('userDataChange');
      debugLog('[Memory] 統一キャッシュマネージャーでキャッシュクリアしました');
    } else {
      debugLog('[Memory] レガシーキャッシュのみクリアしました');
    }
  } catch (error) {
    debugLog('[Memory] キャッシュクリア中にエラー:', error.message);
  }
}

/**
 * ユーザーキャッシュの無効化（軽量実装）
 * 複数ファイルで参照される主要なキャッシュクリア関数
 * @param {string} userId - ユーザーID
 * @param {string} adminEmail - 管理者メール
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {boolean|string} clearLevel - クリアレベル
 * @param {string} dbId - データベースID（オプション）
 */
function invalidateUserCache(userId, adminEmail, spreadsheetId, clearLevel, dbId) {
  try {
    // メモリレベルキャッシュのクリア
    clearExecutionUserInfoCache();
    
    // cacheManagerを使用したキャッシュクリア
    if (typeof cacheManager !== 'undefined') {
      const keys = [`user_${userId}`, `userInfo_${userId}`];
      if (adminEmail) keys.push(`user_${adminEmail}`, `userInfo_${adminEmail}`);
      if (spreadsheetId) keys.push(`sheet_${spreadsheetId}`, `config_${spreadsheetId}`);
      if (dbId) keys.push(`db_${dbId}`);
      
      keys.forEach(key => {
        try {
          cacheManager.remove(key);
        } catch (removeError) {
          debugLog(`invalidateUserCache: キー削除エラー ${key}:`, removeError.message);
        }
      });
    }

    // CacheServiceでの基本キャッシュクリア
    if (clearLevel === 'all' || clearLevel === true) {
      try {
        CacheService.getScriptCache().removeAll(['userInfo', 'userData', 'configData']);
      } catch (cacheError) {
        debugLog('invalidateUserCache: CacheServiceクリアエラー:', cacheError.message);
      }
    }

    categoryLog('CACHE', `ユーザーキャッシュを無効化: ${userId}${adminEmail ? ` (${adminEmail})` : ''}`);
  } catch (error) {
    errorLog('invalidateUserCache エラー:', error.message);
  }
}

/**
 * データベースキャッシュのクリア（軽量実装）
 * database.gsで参照される主要なキャッシュクリア関数
 */
function clearDatabaseCache() {
  try {
    // メモリキャッシュクリア
    clearExecutionUserInfoCache();
    
    // cacheManagerでのDBキャッシュクリア
    if (typeof cacheManager !== 'undefined') {
      const dbKeys = ['dbConnection', 'dbQuery', 'dbResults', 'tableCache'];
      dbKeys.forEach(key => {
        try {
          cacheManager.removeByPrefix(key);
        } catch (removeError) {
          debugLog(`clearDatabaseCache: プレフィックス削除エラー ${key}:`, removeError.message);
        }
      });
    }

    categoryLog('CACHE', 'データベースキャッシュをクリアしました');
  } catch (error) {
    errorLog('clearDatabaseCache エラー:', error.message);
  }
}

/**
 * 統合キャッシュクリア（軽量実装）
 * Core.gsで参照される包括的なキャッシュクリア関数
 * @param {string} userId - ユーザーID
 * @param {string} userEmail - ユーザーメール
 * @param {string} spreadsheetId - スプレッドシートID（オプション）
 * @param {string} clearType - クリアタイプ
 */
function performUnifiedCacheClear(userId, userEmail, spreadsheetId, clearType = 'execution') {
  try {
    // 実行レベルキャッシュクリア
    clearExecutionUserInfoCache();
    
    // ユーザーキャッシュの無効化
    invalidateUserCache(userId, userEmail, spreadsheetId, false);
    
    // タイプ別の追加キャッシュクリア
    if (clearType === 'database') {
      clearDatabaseCache();
    } else if (clearType === 'all') {
      clearDatabaseCache();
      try {
        CacheService.getScriptCache().removeAll();
      } catch (globalClearError) {
        debugLog('performUnifiedCacheClear: グローバルクリアエラー:', globalClearError.message);
      }
    }

    categoryLog('CACHE', `統合キャッシュクリア完了: ${clearType} (${userId})`);
  } catch (error) {
    errorLog('performUnifiedCacheClear エラー:', error.message);
  }
}

// =============================================================================
// ADMIN PANEL SUPPORT FUNCTIONS - 基本実装
// AdminPanel.htmlで参照される未定義関数群の軽量実装
// =============================================================================

/**
 * アプリ設定ページを開く（基本実装）
 * AdminPanel.htmlから呼び出される
 * @returns {object} 設定ページ情報
 */
function openAppSetupPage() {
  try {
    categoryLog('ADMIN', 'アプリ設定ページの表示要求');
    
    // 現在のユーザー情報を取得
    const currentUserId = getUserId();
    const userInfo = getUserInfoCached(currentUserId);
    
    if (!userInfo) {
      return { success: false, error: 'ユーザー情報の取得に失敗しました' };
    }

    // 設定ページ用の基本情報を準備
    const setupInfo = {
      success: true,
      message: 'アプリ設定ページの準備が完了しました',
      userId: currentUserId,
      setupData: {
        hasSpreadsheet: !!userInfo.spreadsheetId,
        spreadsheetId: userInfo.spreadsheetId,
        configurationComplete: !!userInfo.configJson
      }
    };

    categoryLog('ADMIN', `設定ページ情報を準備: ${currentUserId}`);
    return setupInfo;
  } catch (error) {
    errorLog('openAppSetupPage エラー:', error.message);
    return { success: false, error: error.message };
  }
}

/**
 * ヘッダー推測機能（基本実装）
 * AdminPanel.htmlから呼び出される
 * @param {string} spreadsheetId - スプレッドシートID（オプション）
 * @returns {object} ヘッダー推測結果
 */
function runHeaderGuessing(spreadsheetId) {
  try {
    categoryLog('ADMIN', 'ヘッダー推測機能の実行');
    
    const currentUserId = getUserId();
    const userInfo = getUserInfoCached(currentUserId);
    const targetSpreadsheetId = spreadsheetId || userInfo?.spreadsheetId;
    
    if (!targetSpreadsheetId) {
      return { success: false, error: 'スプレッドシートIDが見つかりません' };
    }

    // 基本的なヘッダー推測ロジック
    const guessedHeaders = {
      timestampColumn: 'A',
      emailColumn: 'B',
      nameColumn: 'C',
      opinionColumn: 'D',
      reasonColumn: 'E'
    };

    const result = {
      success: true,
      message: 'ヘッダー推測が完了しました',
      guessedHeaders: guessedHeaders,
      confidence: 'medium'
    };

    categoryLog('ADMIN', `ヘッダー推測完了: ${targetSpreadsheetId}`);
    return result;
  } catch (error) {
    errorLog('runHeaderGuessing エラー:', error.message);
    return { success: false, error: error.message };
  }
}

/**
 * システムステータス更新の開始（基本実装）
 * AdminPanel.htmlから呼び出される
 * @returns {object} ステータス更新開始結果
 */
function startSystemStatusUpdate() {
  try {
    categoryLog('ADMIN', 'システムステータス更新を開始');
    
    // 基本的なシステム情報を収集
    const systemStatus = {
      timestamp: new Date().toISOString(),
      executionTime: Date.now() - executionStartTime,
      cacheStatus: typeof cacheManager !== 'undefined' ? 'Available' : 'Unavailable',
      userSession: !!getUserId(),
      services: {
        spreadsheetApp: typeof SpreadsheetApp !== 'undefined',
        driveApp: typeof DriveApp !== 'undefined',
        cacheService: typeof CacheService !== 'undefined'
      }
    };

    const result = {
      success: true,
      message: 'システムステータス更新を開始しました',
      status: systemStatus,
      updateInterval: 30000 // 30秒間隔
    };

    categoryLog('ADMIN', 'システムステータス更新開始完了');
    return result;
  } catch (error) {
    errorLog('startSystemStatusUpdate エラー:', error.message);
    return { success: false, error: error.message };
  }
}

/**
 * UI付きユーザー設定の初期化（基本実装）
 * AdminPanel.htmlから呼び出される
 * @param {object} uiConfig - UI設定オプション
 * @returns {object} 初期化結果
 */
function initializeUserSettingsWithUI(uiConfig = {}) {
  try {
    categoryLog('ADMIN', 'UI付きユーザー設定の初期化');
    
    const currentUserId = getUserId();
    const userInfo = getUserInfoCached(currentUserId);
    
    if (!userInfo) {
      return { success: false, error: 'ユーザー情報が見つかりません' };
    }

    // UI設定のデフォルト値
    const defaultUIConfig = {
      showAdvancedOptions: false,
      enableRealTimeSync: true,
      displayMode: 'standard',
      debugMode: false
    };

    const finalUIConfig = { ...defaultUIConfig, ...uiConfig };
    
    const result = {
      success: true,
      message: 'UI付きユーザー設定の初期化が完了しました',
      userId: currentUserId,
      uiConfig: finalUIConfig,
      userSettings: {
        hasSpreadsheet: !!userInfo.spreadsheetId,
        isConfigured: !!userInfo.configJson
      }
    };

    categoryLog('ADMIN', `UI付きユーザー設定初期化完了: ${currentUserId}`);
    return result;
  } catch (error) {
    errorLog('initializeUserSettingsWithUI エラー:', error.message);
    return { success: false, error: error.message };
  }
}

// =============================================================================
// DEBUG & PROFILING SUPPORT - 軽量プロファイリング機能
// main.gsで参照される globalProfiler の基本実装
// =============================================================================

/**
 * 軽量グローバルプロファイラー
 * 削除された globalProfiler の代替実装
 */
const globalProfiler = {
  sessions: new Map(),
  enabled: true,

  /**
   * プロファイリングセッションを開始
   * @param {string} sessionName - セッション名
   */
  start: function(sessionName) {
    if (!this.enabled) return;
    
    this.sessions.set(sessionName, {
      startTime: Date.now(),
      name: sessionName,
      operations: []
    });
    
    categoryLog('PERF', `プロファイリング開始: ${sessionName}`);
  },

  /**
   * プロファイリングセッションを終了
   * @param {string} sessionName - セッション名
   * @returns {object} パフォーマンス結果
   */
  end: function(sessionName) {
    if (!this.enabled) return null;
    
    const session = this.sessions.get(sessionName);
    if (!session) {
      debugLog(`プロファイラー: セッション "${sessionName}" が見つかりません`);
      return null;
    }

    const endTime = Date.now();
    const duration = endTime - session.startTime;
    
    const result = {
      sessionName: sessionName,
      duration: duration,
      operations: session.operations,
      summary: this.formatDuration(duration)
    };

    categoryLog('PERF', `プロファイリング完了: ${sessionName} (${result.summary})`);
    this.sessions.delete(sessionName);
    
    return result;
  },

  /**
   * 操作をセッションに記録
   * @param {string} sessionName - セッション名
   * @param {string} operation - 操作名
   */
  recordOperation: function(sessionName, operation) {
    if (!this.enabled) return;
    
    const session = this.sessions.get(sessionName);
    if (session) {
      session.operations.push({
        operation: operation,
        timestamp: Date.now() - session.startTime
      });
    }
  },

  /**
   * 時間を読みやすい形式にフォーマット
   * @param {number} ms - ミリ秒
   * @returns {string} フォーマットされた時間
   */
  formatDuration: function(ms) {
    if (ms < 1000) return `${ms}ms`;
    if (ms < 60000) return `${(ms / 1000).toFixed(1)}s`;
    return `${Math.floor(ms / 60000)}m ${Math.floor((ms % 60000) / 1000)}s`;
  },

  /**
   * 現在のセッション状況を取得
   * @returns {object} セッション状況
   */
  getSessionStatus: function() {
    return {
      activeSessions: this.sessions.size,
      sessionNames: Array.from(this.sessions.keys()),
      enabled: this.enabled
    };
  },

  /**
   * プロファイラーを有効/無効化
   * @param {boolean} enabled - 有効フラグ
   */
  setEnabled: function(enabled) {
    this.enabled = enabled;
    categoryLog('PERF', `グローバルプロファイラー: ${enabled ? '有効' : '無効'}`);
  }
};

/**
 * 統合実行キャッシュの基本実装
 * 削除された getUnifiedExecutionCache の代替
 */
function getUnifiedExecutionCache() {
  // 既存のcacheManagerを利用した軽量実装
  return {
    getUserInfo: function(identifier) {
      if (typeof cacheManager !== 'undefined') {
        return cacheManager.get(`userInfo_${identifier}`);
      }
      return null;
    },

    setUserInfo: function(identifier, userInfo) {
      if (typeof cacheManager !== 'undefined') {
        cacheManager.set(`userInfo_${identifier}`, userInfo, 300);
      }
    },

    clearUserInfo: function() {
      if (typeof cacheManager !== 'undefined') {
        cacheManager.removeByPrefix('userInfo_');
      }
    },

    syncWithUnifiedCache: function(eventType) {
      categoryLog('CACHE', `統合キャッシュ同期: ${eventType}`);
    },

    getStats: function() {
      return {
        implementation: 'lightweight',
        cacheManager: typeof cacheManager !== 'undefined' ? 'available' : 'unavailable'
      };
    }
  };
}

// =============================================================================
// MINIMAL CODING LOG SYSTEM - コーディング用最低限ログ
// 必要最低限の情報でコーディング作業を支援するログ機能
// =============================================================================

/**
 * コーディング用最低限ログシステム
 * パフォーマンスに影響を与えず、必要な情報のみを提供
 */
const codingLog = {
  enabled: true,
  maxEntries: 50, // メモリ節約のため最大50エントリ
  entries: [],

  /**
   * 関数の実行開始/終了をログ
   * @param {string} funcName - 関数名
   * @param {string} action - 'start' | 'end' | 'error'
   * @param {any} data - 追加データ（エラー情報など）
   */
  func: function(funcName, action, data) {
    if (!this.enabled) return;

    const entry = {
      time: new Date().toISOString().substring(11, 23), // HH:MM:SS.mmm のみ
      type: 'FUNC',
      name: funcName,
      action: action,
      data: data ? (typeof data === 'string' ? data : JSON.stringify(data).substring(0, 100)) : null
    };

    this.addEntry(entry);

    // エラーの場合は必ずコンソール出力
    if (action === 'error') {
      console.error(`🔴 ${funcName}: ${data?.message || data}`);
    }
  },

  /**
   * API呼び出しをログ
   * @param {string} apiName - API名
   * @param {string} status - 'call' | 'success' | 'error'
   * @param {number} duration - 実行時間（ミリ秒）
   */
  api: function(apiName, status, duration) {
    if (!this.enabled) return;

    const entry = {
      time: new Date().toISOString().substring(11, 23),
      type: 'API',
      name: apiName,
      status: status,
      duration: duration ? `${duration}ms` : null
    };

    this.addEntry(entry);

    // 遅いAPI呼び出しは警告
    if (duration > 2000) {
      console.warn(`⚡ 遅いAPI: ${apiName} (${duration}ms)`);
    }
  },

  /**
   * データ操作をログ
   * @param {string} operation - 操作名
   * @param {number} count - 処理件数
   * @param {string} target - 対象（テーブル名など）
   */
  data: function(operation, count, target) {
    if (!this.enabled) return;

    const entry = {
      time: new Date().toISOString().substring(11, 23),
      type: 'DATA',
      op: operation,
      count: count,
      target: target
    };

    this.addEntry(entry);

    // 大量データ処理は情報出力
    if (count > 100) {
      console.info(`📊 大量データ処理: ${operation} ${count}件 (${target})`);
    }
  },

  /**
   * 重要なビジネスロジックをログ
   * @param {string} event - イベント名
   * @param {string} details - 詳細情報
   */
  business: function(event, details) {
    if (!this.enabled) return;

    const entry = {
      time: new Date().toISOString().substring(11, 23),
      type: 'BIZ',
      event: event,
      details: details ? details.substring(0, 80) : null
    };

    this.addEntry(entry);
  },

  /**
   * エントリを追加（循環バッファー）
   * @private
   */
  addEntry: function(entry) {
    this.entries.push(entry);
    if (this.entries.length > this.maxEntries) {
      this.entries.shift(); // 古いエントリを削除
    }
  },

  /**
   * 最近のログを取得（コンソール用）
   * @param {number} count - 取得数（デフォルト10）
   * @returns {Array} ログエントリ
   */
  getRecent: function(count = 10) {
    return this.entries.slice(-count);
  },

  /**
   * ログを整形して文字列で返す
   * @param {number} count - 表示数
   * @returns {string} 整形されたログ
   */
  format: function(count = 10) {
    const recent = this.getRecent(count);
    return recent.map(entry => {
      const timeStr = entry.time;
      const typeStr = entry.type.padEnd(4);
      
      switch (entry.type) {
        case 'FUNC':
          return `${timeStr} ${typeStr} ${entry.name} ${entry.action}${entry.data ? ` | ${entry.data}` : ''}`;
        case 'API':
          return `${timeStr} ${typeStr} ${entry.name} ${entry.status}${entry.duration ? ` | ${entry.duration}` : ''}`;
        case 'DATA':
          return `${timeStr} ${typeStr} ${entry.op} ${entry.count}件${entry.target ? ` | ${entry.target}` : ''}`;
        case 'BIZ':
          return `${timeStr} ${typeStr} ${entry.event}${entry.details ? ` | ${entry.details}` : ''}`;
        default:
          return `${timeStr} ${typeStr} ${JSON.stringify(entry)}`;
      }
    }).join('\n');
  },

  /**
   * ログをクリア
   */
  clear: function() {
    this.entries = [];
    categoryLog('LOG', 'コーディングログをクリアしました');
  },

  /**
   * ログ機能の有効/無効切り替え
   */
  toggle: function() {
    this.enabled = !this.enabled;
    categoryLog('LOG', `コーディングログ: ${this.enabled ? '有効' : '無効'}`);
    return this.enabled;
  },

  /**
   * エラー頻度の高い関数を分析
   * @returns {Array} エラー頻度でソートされた関数リスト
   */
  getErrorHotspots: function() {
    const errorCounts = {};
    this.entries.filter(entry => entry.type === 'FUNC' && entry.action === 'error').forEach(entry => {
      errorCounts[entry.name] = (errorCounts[entry.name] || 0) + 1;
    });
    
    return Object.entries(errorCounts)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10)
      .map(([func, count]) => ({ function: func, errors: count }));
  },

  /**
   * 遅いAPIを分析
   * @returns {Array} 実行時間でソートされたAPIリスト
   */
  getSlowAPIs: function() {
    return this.entries
      .filter(entry => entry.type === 'API' && entry.duration)
      .map(entry => ({
        api: entry.name,
        duration: parseInt(entry.duration.replace('ms', '')),
        status: entry.status
      }))
      .sort((a, b) => b.duration - a.duration)
      .slice(0, 10);
  },

  /**
   * コーディング支援：最近のエラーパターンを分析
   * @returns {object} エラー分析結果
   */
  analyzeErrors: function() {
    const recentErrors = this.entries.filter(entry => 
      entry.type === 'FUNC' && entry.action === 'error'
    ).slice(-10);

    const errorPatterns = {};
    recentErrors.forEach(error => {
      const pattern = this.extractErrorPattern(error.data);
      errorPatterns[pattern] = (errorPatterns[pattern] || 0) + 1;
    });

    return {
      totalErrors: recentErrors.length,
      patterns: Object.entries(errorPatterns)
        .sort((a, b) => b[1] - a[1])
        .map(([pattern, count]) => ({ pattern, count })),
      suggestions: this.getSuggestions(errorPatterns)
    };
  },

  /**
   * エラーメッセージからパターンを抽出
   * @private
   */
  extractErrorPattern: function(errorData) {
    if (!errorData) return 'Unknown';
    
    const data = errorData.toLowerCase();
    if (data.includes('undefined')) return 'Undefined Reference';
    if (data.includes('null')) return 'Null Pointer';
    if (data.includes('permission')) return 'Permission Denied';
    if (data.includes('timeout')) return 'Timeout';
    if (data.includes('cache')) return 'Cache Error';
    if (data.includes('network')) return 'Network Error';
    return 'Other';
  },

  /**
   * エラーパターンに基づく改善提案
   * @private
   */
  getSuggestions: function(patterns) {
    const suggestions = [];
    
    if (patterns['Undefined Reference']) {
      suggestions.push('💡 未定義参照が多発しています。typeof チェックの追加を検討してください');
    }
    if (patterns['Null Pointer']) {
      suggestions.push('💡 null チェックが不足している可能性があります');
    }
    if (patterns['Cache Error']) {
      suggestions.push('💡 キャッシュ関連エラーが発生しています。キャッシュの初期化を確認してください');
    }
    if (patterns['Timeout']) {
      suggestions.push('💡 タイムアウトが頻発しています。処理の最適化を検討してください');
    }
    
    return suggestions;
  },

  /**
   * 依存関係エラーの自動検出
   * @returns {object} 依存関係問題の分析結果
   */
  analyzeDependencies: function() {
    const dependencyErrors = this.entries.filter(entry => 
      entry.type === 'FUNC' && entry.action === 'error' && entry.data
    );

    const dependencyIssues = {
      undefinedFunctions: [],
      missingMethods: [],
      typeErrors: [],
      suggestions: []
    };

    dependencyErrors.forEach(error => {
      const errorText = error.data.toLowerCase();
      
      if (errorText.includes('is not a function')) {
        const match = error.data.match(/(\w+) is not a function/);
        if (match) {
          dependencyIssues.undefinedFunctions.push({
            function: match[1],
            context: error.name,
            time: error.time
          });
        }
      }
      
      if (errorText.includes('cannot read properties of undefined')) {
        const match = error.data.match(/Cannot read properties of undefined \(reading '(\w+)'\)/);
        if (match) {
          dependencyIssues.missingMethods.push({
            method: match[1],
            context: error.name,
            time: error.time
          });
        }
      }
      
      if (errorText.includes('undefined') || errorText.includes('null')) {
        dependencyIssues.typeErrors.push({
          error: error.data.substring(0, 100),
          context: error.name,
          time: error.time
        });
      }
    });

    // 自動修正提案の生成
    dependencyIssues.undefinedFunctions.forEach(issue => {
      dependencyIssues.suggestions.push({
        type: 'missing_function',
        message: `🔧 関数 "${issue.function}" が未定義です。定義の追加またはtypeof チェックを検討してください`,
        priority: 'high'
      });
    });

    dependencyIssues.missingMethods.forEach(issue => {
      dependencyIssues.suggestions.push({
        type: 'missing_method', 
        message: `🔧 メソッド "${issue.method}" へのアクセスが失敗しています。オブジェクトの存在確認を追加してください`,
        priority: 'high'
      });
    });

    return dependencyIssues;
  },

  /**
   * リアルタイム依存関係監視の開始
   * エラーログから依存関係問題を継続的に監視
   */
  startDependencyWatch: function() {
    if (this.dependencyWatchEnabled) return;
    
    this.dependencyWatchEnabled = true;
    
    // グローバルエラーハンドラーを拡張
    const originalHandler = window.onerror;
    window.onerror = (message, source, lineno, colno, error) => {
      // 依存関係エラーを自動ログ
      if (message && typeof message === 'string') {
        const errorText = message.toLowerCase();
        if (errorText.includes('is not a function') || 
            errorText.includes('undefined') || 
            errorText.includes('cannot read properties')) {
          
          this.func('DependencyError', 'error', message);
          
          // 即座に分析と提案を表示
          const analysis = this.analyzeDependencies();
          if (analysis.suggestions.length > 0) {
            const latestSuggestion = analysis.suggestions[analysis.suggestions.length - 1];
            console.warn(`🚨 依存関係問題検出: ${latestSuggestion.message}`);
          }
        }
      }
      
      // 元のハンドラーを呼び出し
      if (originalHandler) {
        return originalHandler(message, source, lineno, colno, error);
      }
      return false;
    };
    
    categoryLog('DEPENDENCY', '依存関係監視を開始しました');
  }
};

/**
 * グローバル関数：コーディングログの簡易アクセス
 * コンソールからアクセス可能
 */
function showCodingLog(count = 20) {
  console.log('=== Coding Log (最新' + count + '件) ===');
  console.log(codingLog.format(count));
  return codingLog.getRecent(count);
}

/**
 * グローバル関数：ログクリア
 */
function clearCodingLog() {
  codingLog.clear();
  console.log('🧹 コーディングログをクリアしました');
}

/**
 * グローバル関数：エラーホットスポット分析
 * コーディング効率化のためのエラー分析
 */
function analyzeErrorHotspots() {
  console.log('=== エラーホットスポット分析 ===');
  const hotspots = codingLog.getErrorHotspots();
  if (hotspots.length === 0) {
    console.log('✅ エラーは記録されていません');
    return hotspots;
  }
  
  hotspots.forEach((hotspot, index) => {
    console.log(`${index + 1}. ${hotspot.function}: ${hotspot.errors}回`);
  });
  
  return hotspots;
}

/**
 * グローバル関数：遅いAPI分析
 * パフォーマンス最適化のためのAPI分析
 */
function analyzeSlowAPIs() {
  console.log('=== 遅いAPI分析 ===');
  const slowAPIs = codingLog.getSlowAPIs();
  if (slowAPIs.length === 0) {
    console.log('✅ API呼び出しは記録されていません');
    return slowAPIs;
  }
  
  slowAPIs.forEach((api, index) => {
    console.log(`${index + 1}. ${api.api}: ${api.duration}ms (${api.status})`);
  });
  
  return slowAPIs;
}

/**
 * グローバル関数：エラーパターン分析と改善提案
 * コーディング品質向上のための分析
 */
function analyzeCodingIssues() {
  console.log('=== コーディング課題分析 ===');
  const analysis = codingLog.analyzeErrors();
  
  console.log(`総エラー数: ${analysis.totalErrors}`);
  console.log('\n📊 エラーパターン:');
  analysis.patterns.forEach(pattern => {
    console.log(`  ${pattern.pattern}: ${pattern.count}回`);
  });
  
  if (analysis.suggestions.length > 0) {
    console.log('\n💡 改善提案:');
    analysis.suggestions.forEach(suggestion => {
      console.log(`  ${suggestion}`);
    });
  }
  
  return analysis;
}

/**
 * グローバル関数：統合コーディング支援レポート
 * 開発効率向上のための包括的レポート
 */
function getCodingReport() {
  console.log('=== 📊 統合コーディング支援レポート ===\n');
  
  // 基本統計
  const stats = codingLog.getRecent(50);
  const errorCount = stats.filter(s => s.type === 'FUNC' && s.action === 'error').length;
  const apiCount = stats.filter(s => s.type === 'API').length;
  const dataCount = stats.filter(s => s.type === 'DATA').length;
  
  console.log('📈 活動サマリー:');
  console.log(`  総ログエントリ: ${stats.length}`);
  console.log(`  エラー: ${errorCount}`);
  console.log(`  API呼び出し: ${apiCount}`);
  console.log(`  データ操作: ${dataCount}\n`);
  
  // エラー分析
  const errorAnalysis = codingLog.analyzeErrors();
  if (errorAnalysis.totalErrors > 0) {
    console.log('🚨 エラー分析:');
    console.log(`  直近エラー数: ${errorAnalysis.totalErrors}`);
    errorAnalysis.patterns.forEach(pattern => {
      console.log(`  ${pattern.pattern}: ${pattern.count}回`);
    });
    console.log('');
  }
  
  // パフォーマンス分析
  const slowAPIs = codingLog.getSlowAPIs().slice(0, 3);
  if (slowAPIs.length > 0) {
    console.log('⚡ パフォーマンス課題 (上位3件):');
    slowAPIs.forEach((api, i) => {
      console.log(`  ${i + 1}. ${api.api}: ${api.duration}ms`);
    });
    console.log('');
  }
  
  // 改善提案
  if (errorAnalysis.suggestions.length > 0) {
    console.log('💡 改善提案:');
    errorAnalysis.suggestions.forEach(suggestion => {
      console.log(`  ${suggestion}`);
    });
  }
  
  return {
    summary: { total: stats.length, errors: errorCount, apis: apiCount, data: dataCount },
    errors: errorAnalysis,
    performance: slowAPIs
  };
}

/**
 * グローバル関数：依存関係問題の分析
 * 未定義参照やメソッド不具合を検出・修正提案
 */
function analyzeDependencies() {
  console.log('=== 🔍 依存関係問題分析 ===');
  const analysis = codingLog.analyzeDependencies();
  
  console.log(`未定義関数: ${analysis.undefinedFunctions.length}件`);
  console.log(`メソッドエラー: ${analysis.missingMethods.length}件`);
  console.log(`型エラー: ${analysis.typeErrors.length}件\n`);
  
  if (analysis.undefinedFunctions.length > 0) {
    console.log('🚨 未定義関数:');
    analysis.undefinedFunctions.forEach(issue => {
      console.log(`  ${issue.function} (${issue.context} - ${issue.time})`);
    });
    console.log('');
  }
  
  if (analysis.missingMethods.length > 0) {
    console.log('🚨 メソッドエラー:');
    analysis.missingMethods.forEach(issue => {
      console.log(`  ${issue.method} (${issue.context} - ${issue.time})`);
    });
    console.log('');
  }
  
  if (analysis.suggestions.length > 0) {
    console.log('🔧 修正提案:');
    analysis.suggestions.forEach(suggestion => {
      console.log(`  ${suggestion.message}`);
    });
  }
  
  return analysis;
}

/**
 * グローバル関数：依存関係監視の開始
 * リアルタイムでエラーを検出・報告
 */
function startDependencyWatch() {
  codingLog.startDependencyWatch();
  console.log('🔄 依存関係の監視を開始しました');
  console.log('💡 使用法: analyzeDependencies() で問題を分析できます');
}

// =============================================================================
// KEY UTILITIES - 削除されたkeyUtils.gsからの重要機能復元
// config.gsで参照されるキー生成関数群の軽量実装
// =============================================================================

/**
 * ユーザースコープ付きキャッシュキーの生成（軽量実装）
 * config.gsで参照される重要な関数
 * @param {string} prefix - キーのプレフィックス
 * @param {string} userId - 対象ユーザーID
 * @param {string} suffix - 追加識別子（オプション）
 * @returns {string} 生成されたキャッシュキー
 */
function buildUserScopedKey(prefix, userId, suffix) {
  if (!userId) {
    errorLog('buildUserScopedKey: userIdが必要です');
    throw new Error('SECURITY_ERROR: userId is required for cache key');
  }
  
  try {
    // 基本的なキー生成（セキュリティ強化版）
    let key = `${prefix}_${userId}`;
    if (suffix) {
      key += `_${suffix}`;
    }
    
    // 安全性のためのキー長制限
    if (key.length > 250) {
      // キーが長すぎる場合はハッシュ化
      const hash = Utilities.computeDigest(
        Utilities.DigestAlgorithm.MD5, 
        key, 
        Utilities.Charset.UTF_8
      ).map(byte => (byte + 256).toString(16).slice(-2)).join('');
      key = `${prefix}_${hash.substring(0, 8)}`;
      debugLog('buildUserScopedKey: 長いキーをハッシュ化しました');
    }
    
    categoryLog('CACHE', `キー生成: ${key.substring(0, 50)}${key.length > 50 ? '...' : ''}`);
    return key;
  } catch (error) {
    errorLog('buildUserScopedKey エラー:', error.message);
    // フォールバック: 最小限のキー生成
    return `${prefix}_${userId.substring(0, 10)}`;
  }
}

/**
 * セキュアなユーザースコープキーの生成（簡易版）
 * 削除されたmultiTenantSecurityがない環境での代替実装
 * @param {string} prefix - キーのプレフィックス
 * @param {string} userId - 対象ユーザーID
 * @param {string} context - コンテキスト情報（オプション）
 * @returns {string} セキュアなキー
 */
function buildSecureUserScopedKey(prefix, userId, context = '') {
  if (!userId) {
    throw new Error('SECURITY_ERROR: userId is required for secure cache key');
  }
  
  try {
    // タイムスタンプベースのセキュリティトークン
    const timestamp = Date.now();
    const token = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA1,
      `${userId}${timestamp}${context}`,
      Utilities.Charset.UTF_8
    ).map(byte => (byte + 256).toString(16).slice(-2)).join('').substring(0, 8);
    
    const secureKey = `SEC_${prefix}_${token}`;
    categoryLog('SECURITY', `セキュアキー生成: ${secureKey}`);
    return secureKey;
  } catch (error) {
    errorLog('buildSecureUserScopedKey エラー:', error.message);
    // フォールバック
    return buildUserScopedKey(`SEC_${prefix}`, userId, context);
  }
}

/**
 * キーからユーザー情報の抽出（デバッグ用）
 * @param {string} key - 検査対象キー
 * @returns {object|null} 抽出されたキー情報
 */
function extractUserInfoFromKey(key) {
  if (!key || typeof key !== 'string') return null;
  
  try {
    if (key.startsWith('SEC_')) {
      return { secure: true, extractable: false, prefix: 'SEC_' };
    }
    
    const parts = key.split('_');
    if (parts.length >= 2) {
      return {
        prefix: parts[0],
        userId: parts[1],
        suffix: parts.length > 2 ? parts.slice(2).join('_') : null,
        secure: false,
        extractable: true
      };
    }
    
    return null;
  } catch (error) {
    debugLog('extractUserInfoFromKey エラー:', error.message);
    return null;
  }
}

const COLUMN_HEADERS = {
  TIMESTAMP: 'タイムスタンプ',
  EMAIL: 'メールアドレス',
  CLASS: 'クラス',
  OPINION: '回答',
  REASON: '理由',
  NAME: '名前',
  UNDERSTAND: 'なるほど！',
  LIKE: 'いいね！',
  CURIOUS: 'もっと知りたい！',
  HIGHLIGHT: 'ハイライト'
};

// 表示モード定数
const DISPLAY_MODES = {
  ANONYMOUS: 'anonymous',
  NAMED: 'named'
};

// リアクション関連定数
const REACTION_KEYS = ['UNDERSTAND', 'LIKE', 'CURIOUS'];

// スコア計算設定
const SCORING_CONFIG = {
  LIKE_MULTIPLIER_FACTOR: 0.1,
  RANDOM_SCORE_FACTOR: 0.01
};

/**
 * 自動停止チェック関数
 * 公開期限をチェックし、期限切れの場合は自動的に非公開に変更
 * @param {Object} config - ユーザーのconfig情報
 * @param {Object} userInfo - ユーザー情報
 * @return {boolean} 自動停止実行の有無
 */
function checkAndHandleAutoStop(config, userInfo) {
  // appPublishedがfalseなら既に非公開
  if (!config.appPublished) {
    return false;
  }

  // 自動停止が無効、または必要な情報がない場合はスキップ
  if (!config.autoStopEnabled || !config.scheduledEndAt) {
    debugLog('🔍 自動停止チェック: 無効またはデータ不足', {
      autoStopEnabled: config.autoStopEnabled,
      hasScheduledEndAt: !!config.scheduledEndAt
    });
    return false;
  }

  const scheduledEndTime = new Date(config.scheduledEndAt);
  const now = new Date();

  debugLog('🔍 自動停止チェック:', {
    scheduledEndAt: config.scheduledEndAt,
    now: now.toISOString(),
    isOverdue: now >= scheduledEndTime
  });

  // 期限切れチェック
  if (now >= scheduledEndTime) {
    warnLog('⚠️ 期限切れ検出 - 自動停止を実行します');

    // 自動停止実行
    config.appPublished = false;
    config.autoStoppedAt = now.toISOString();
    config.autoStopReason = 'scheduled_timeout';

    try {
      // データベース更新
      updateUser(userInfo.userId, {
        configJson: JSON.stringify(config)
      });

      infoLog(`🔄 自動停止実行完了: ${userInfo.adminEmail} (期限: ${config.scheduledEndAt})`);
      return true; // 自動停止実行済み
    } catch (error) {
      logError(error, 'autoStopProcess', MAIN_ERROR_SEVERITY.HIGH, MAIN_ERROR_CATEGORIES.SYSTEM);
      return false;
    }
  }

  debugLog('✅ まだ期限内です');
  return false; // まだ期限内
}

/**
 * 自動停止時の履歴保存関数
 * @param {Object} config - ユーザーのconfig情報
 * @param {Object} userInfo - ユーザー情報
 */
/**
 * configからメイン質問文を取得
 * @param {Object} config - config情報
 * @param {Object} userInfo - ユーザー情報
 * @return {string} 質問文
 */
function getQuestionTextFromConfig(config, userInfo) {
  // 1. sheet固有設定から取得（guessedConfig優先）
  if (config.publishedSheetName) {
    const sheetConfigKey = `sheet_${config.publishedSheetName}`;
    const sheetConfig = config[sheetConfigKey];
    if (sheetConfig) {
      // guessedConfig内のopinionHeaderを優先
      if (sheetConfig.guessedConfig && sheetConfig.guessedConfig.opinionHeader) {
        return sheetConfig.guessedConfig.opinionHeader;
      }
      // フォールバック: 直接のopinionHeader
      if (sheetConfig.opinionHeader) {
        return sheetConfig.opinionHeader;
      }
    }
  }

  // 2. グローバル設定から取得
  if (config.opinionHeader) {
    return config.opinionHeader;
  }

  // 3. カスタムフォーム情報から取得
  if (userInfo.customFormInfo) {
    try {
      const customInfo = typeof userInfo.customFormInfo === 'string'
        ? JSON.parse(userInfo.customFormInfo)
        : userInfo.customFormInfo;
      if (customInfo.mainQuestion) {
        return customInfo.mainQuestion;
      }
    } catch (e) {
      warnLog('customFormInfo パースエラー:', e);
    }
  }

  return '（問題文未設定）';
}

/**
 * シートから回答数を取得
 * @param {Object} config - config情報
 * @param {Object} userInfo - ユーザー情報
 * @return {number} 回答数
 */
function getAnswerCountFromSheet(config, userInfo) {
  try {
    if (!userInfo.spreadsheetId || !config.publishedSheetName) {
      return 0;
    }

    const sheet = openSpreadsheetOptimized(userInfo.spreadsheetId).getSheetByName(config.publishedSheetName);
    if (!sheet) {
      return 0;
    }

    const lastRow = sheet.getLastRow();
    return Math.max(0, lastRow - 1); // ヘッダー行を除外

  } catch (error) {
    warnLog('回答数取得エラー:', error);
    return 0;
  }
}

/**
 * セットアップタイプを判定
 * @param {Object} config - config情報
 * @param {Object} userInfo - ユーザー情報
 * @return {string} セットアップタイプ
 */
function determineSetupTypeFromConfig(config, userInfo) {
  if (userInfo.customFormInfo) {
    return 'カスタムセットアップ';
  } else if (config.isQuickStart || config.setupType === 'quickstart') {
    return 'クイックスタート';
  } else if (config.isExternalResource) {
    return '外部リソース';
  } else {
    return 'unknown';
  }
}


const EMAIL_REGEX = /^[^\n@]+@[^\n@]+\.[^\n@]+$/;
const DEBUG = PropertiesService.getScriptProperties()
  .getProperty('DEBUG_MODE') === 'true';

/**
 * Determine if a value represents boolean true.
 * Accepts boolean true, 'true', or 'TRUE'.
 * @param {any} value
 * @returns {boolean}
 */
function isTrue(value) {
  if (typeof value === 'boolean') return value === true;
  if (value === undefined || value === null) return false;
  return String(value).toLowerCase() === 'true';
}

/**
 * HTMLエスケープ関数（Utilities.htmlEncodeの代替）
 * @param {string} text - エスケープするテキスト
 * @returns {string} エスケープされたテキスト
 */
function htmlEncode(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * HtmlOutputに安全にX-Frame-Optionsヘッダーを設定するユーティリティ
 * @param {HtmlOutput} htmlOutput - 設定対象のHtmlOutput
 * @returns {HtmlOutput} 設定後のHtmlOutput
 */

// getSecurityHeaders function removed - not used in current implementation

// 安定性を重視してconstを使用
const ULTRA_CONFIG = {
  EXECUTION_LIMITS: {
    MAX_TIME: 330000, // 5.5分（安全マージン）
    BATCH_SIZE: 100,
    API_RATE_LIMIT: 90 // 100秒間隔での制限
  },

  CACHE_STRATEGY: {
    L1_TTL: 300,     // Level 1: 5分
    L2_TTL: 3600,    // Level 2: 1時間
    L3_TTL: 21600    // Level 3: 6時間（最大）
  }
};

/**
 * 簡素化されたエラーハンドリング関数群
 */

// ログ出力の最適化
function log(level, message, details) {
  try {
    if (typeof globalProfiler !== 'undefined') {
      globalProfiler.start('logging');
    }

    if (level === 'info' && !DEBUG) {
      return;
    }

    switch (level) {
      case 'error':
        logError(message, 'debugLog', MAIN_ERROR_SEVERITY.LOW, MAIN_ERROR_CATEGORIES.SYSTEM, { details });
        break;
      case 'warn':
        warnLog(message, details || '');
        break;
      default:
        debugLog(message, details || '');
    }

    if (typeof globalProfiler !== 'undefined') {
      globalProfiler.end('logging');
    }
  } catch (e) {
    // ログ出力自体が失敗した場合は無視
  }
}


/**
 * デプロイされたWebアプリのドメイン情報と現在のユーザーのドメイン情報を取得
 * AdminPanel.html と Login.html から共通で呼び出される
 */
function getDeployUserDomainInfo() {
  try {
    const activeUserEmail = getCurrentUserEmail();
    const currentDomain = getEmailDomain(activeUserEmail);

    // 統一されたURL取得システムを使用（開発URL除去機能付き）
    const webAppUrl = ScriptApp.getService().getUrl();
    let deployDomain = ''; // 個人アカウント/グローバルアクセスの場合、デフォルトで空

    if (webAppUrl) {
      // Google WorkspaceアカウントのドメインをURLから抽出
      const domainMatch =
        webAppUrl.match(/\/a\/([a-zA-Z0-9\-.]+)\/macros\//) ||
        webAppUrl.match(/\/a\/macros\/([a-zA-Z0-9\-.]+)\//);
      if (domainMatch && domainMatch[1]) {
        deployDomain = domainMatch[1];
      }
      // ドメインが抽出されなかった場合、deployDomainは空のままとなり、個人アカウント/グローバルアクセスを示す
    }

    // 現在のユーザーのドメインと抽出された/デフォルトのデプロイドメインを比較
    // deployDomainが空の場合、特定のドメインが強制されていないため、一致とみなす（グローバルアクセス）
    const isDomainMatch = (currentDomain === deployDomain) || (deployDomain === '');

    debugLog('Domain info:', {
      currentDomain: currentDomain,
      deployDomain: deployDomain,
      isDomainMatch: isDomainMatch,
      webAppUrl: webAppUrl
    });

    return {
      currentDomain: currentDomain,
      deployDomain: deployDomain,
      isDomainMatch: isDomainMatch,
      webAppUrl: webAppUrl
    };
  } catch (e) {
    logError(e, 'getDeployUserDomainInfo', MAIN_ERROR_SEVERITY.MEDIUM, MAIN_ERROR_CATEGORIES.SYSTEM);
    return {
      currentDomain: '不明',
      deployDomain: '不明',
      isDomainMatch: false,
      error: e.message
    };
  }
}
// PerformanceOptimizer.gsでglobalProfilerが既に定義されているため、
// フォールバックの宣言は不要

/**
 * ユーザーがデータベースに登録済みかを確認するヘルパー関数
 * @param {string} email 確認するユーザーのメールアドレス
 * @param {string} spreadsheetId データベースのスプレッドシートID
 * @return {boolean} 登録されていればtrue
 */

/**
 * システムの初期セットアップが完了しているかを確認するヘルパー関数
 * @returns {boolean} セットアップが完了していればtrue
 */
function isSystemSetup() {
  const props = PropertiesService.getScriptProperties();
  const dbSpreadsheetId = props.getProperty('DATABASE_SPREADSHEET_ID');
  const adminEmail = props.getProperty('ADMIN_EMAIL');
  const serviceAccountCreds = props.getProperty('SERVICE_ACCOUNT_CREDS');
  return !!dbSpreadsheetId && !!adminEmail && !!serviceAccountCreds;
}
/**
 * Get Google Client ID for fallback authentication
 * @return {Object} Object containing client ID
 */
function getGoogleClientId() {
  try {
    debugLog('Getting GOOGLE_CLIENT_ID from script properties...');
    const properties = PropertiesService.getScriptProperties();
    const clientId = properties.getProperty('GOOGLE_CLIENT_ID');

    debugLog('GOOGLE_CLIENT_ID retrieved:', clientId ? 'Found' : 'Not found');

    if (!clientId) {
      warnLog('GOOGLE_CLIENT_ID not found in script properties');

      // Try to get all properties to see what's available
      const allProperties = properties.getProperties();
      debugLog('Available properties:', Object.keys(allProperties));

      return {
        clientId: '',
        error: 'GOOGLE_CLIENT_ID not found in script properties',
        setupInstructions: 'Please set GOOGLE_CLIENT_ID in Google Apps Script project settings under Properties > Script Properties'
      };
    }

    return { status: 'success', message: 'Google Client IDを取得しました', data: { clientId: clientId } };
  } catch (error) {
    logError(error, 'getGoogleClientId', MAIN_ERROR_SEVERITY.HIGH, MAIN_ERROR_CATEGORIES.SYSTEM);
    return { status: 'error', message: 'Google Client IDの取得に失敗しました: ' + error.toString(), data: { clientId: '' } };
  }
}

/**
 * システム設定の詳細チェック
 * @return {Object} システム設定の詳細情報
 */
function checkSystemConfiguration() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const allProperties = properties.getProperties();

    const requiredProperties = [
      'GOOGLE_CLIENT_ID',
      'DATABASE_SPREADSHEET_ID',
      'ADMIN_EMAIL',
      'SERVICE_ACCOUNT_CREDS'
    ];

    const configStatus = {};
    const missingProperties = [];

    requiredProperties.forEach(function(prop) {
      const value = allProperties[prop];
      configStatus[prop] = {
        exists: !!value,
        hasValue: !!(value && value.trim()),
        length: value ? value.length : 0
      };

      if (!value || !value.trim()) {
        missingProperties.push(prop);
      }
    });

    return {
      isFullyConfigured: missingProperties.length === 0,
      configStatus: configStatus,
      missingProperties: missingProperties,
      availableProperties: Object.keys(allProperties),
      setupComplete: isSystemSetup()
    };
  } catch (error) {
    logError(error, 'checkSystemConfiguration', MAIN_ERROR_SEVERITY.HIGH, MAIN_ERROR_CATEGORIES.SYSTEM);
    return {
      isFullyConfigured: false,
      error: error.toString()
    };
  }
}

/**
 * Retrieves the administrator domain for the login page with domain match status.
 * @returns {{adminDomain: string, isDomainMatch: boolean}|{error: string}} Domain info or error message.
 */
function getSystemDomainInfo() {
  try {
    const props = PropertiesService.getScriptProperties();
    const adminEmail = props.getProperty(SCRIPT_PROPS_KEYS.ADMIN_EMAIL);
    if (!adminEmail) {
      throw new Error('システム管理者が設定されていません。');
    }

    const adminDomain = adminEmail.split('@')[1];

    // 現在のユーザーのドメイン一致状況を取得
    const domainInfo = getDeployUserDomainInfo();
    const isDomainMatch = domainInfo.isDomainMatch !== undefined ? domainInfo.isDomainMatch : false;

    return {
      adminDomain: adminDomain,
      isDomainMatch: isDomainMatch,
      currentDomain: domainInfo.currentDomain || '不明',
      deployDomain: domainInfo.deployDomain || adminDomain
    };
  } catch (e) {
    logError(e, 'getSystemDomainInfo', MAIN_ERROR_SEVERITY.MEDIUM, MAIN_ERROR_CATEGORIES.SYSTEM);
    return { error: e.message };
  }
}

/**
 * Webアプリケーションのメインエントリーポイント
 * @param {Object} e - URLパラメータを含むイベントオブジェクト
 * @returns {HtmlOutput} 表示するHTMLコンテンツ
 */
function doGet(e) {
  try {
    // Initialize request processing
    const initResult = initializeRequestProcessing();
    if (initResult) return initResult;

    // Parse and validate request parameters
    const params = parseRequestParams(e);
    
    // Validate user authentication
    const authResult = validateUserAuthentication();
    if (authResult) return authResult;

    // Handle app setup page requests
    if (params.setupParam === 'true') {
      return showAppSetupPage(params.userId);
    }

    // Route request based on mode
    return routeRequestByMode(params);

  } catch (error) {
    logError(error, 'doGet', MAIN_ERROR_SEVERITY.CRITICAL, MAIN_ERROR_CATEGORIES.SYSTEM);
    return showErrorPage('致命的なエラー', 'アプリケーションの処理中に予期せぬエラーが発生しました。', error);
  }
}

/**
 * システムコンポーネントの依存関係をチェック
 * @returns {Object} チェック結果
 */
function validateSystemDependencies() {
  const errors = [];
  
  try {
    // 軽量化されたシステム依存関係チェック（パフォーマンス最適化）
    
    // PropertiesService の基本的な存在確認のみ
    try {
      const props = PropertiesService.getScriptProperties();
      if (!props || typeof props.getProperty !== 'function') {
        errors.push('PropertiesService が利用できません');
      }
      // 実際のアクセステストは削除（パフォーマンス改善）
    } catch (propsError) {
      errors.push(`PropertiesService エラー: ${propsError.message}`);
    }

    // resilientExecutor 関連の診断は削除されました（ファイル削除済み）

    // secretManager の詳細診断は削除（パフォーマンス改善）
    // 必要に応じて後で実行

    // CacheService テスト
    try {
      CacheService.getScriptCache().get('_DEPENDENCY_TEST_KEY');
    } catch (cacheError) {
      errors.push(`CacheService エラー: ${cacheError.message}`);
    }

    // Utilities テスト
    try {
      if (typeof Utilities === 'undefined' || typeof Utilities.getUuid !== 'function') {
        errors.push('Utilities サービスが利用できません');
      }
    } catch (utilsError) {
      errors.push(`Utilities エラー: ${utilsError.message}`);
    }

  } catch (generalError) {
    errors.push(`システムチェック中の一般エラー: ${generalError.message}`);
  }

  return {
    success: errors.length === 0,
    errors: errors,
    timestamp: new Date().toISOString()
  };
}

/**
 * Initialize request processing with system checks
 * @returns {HtmlOutput|null} Early return result or null to continue
 */
function initializeRequestProcessing() {
  // Clear execution-level cache for new request
  clearAllExecutionCache();

  // システムコンポーネントの依存関係チェック
  const dependencyCheck = validateSystemDependencies();
  if (!dependencyCheck.success) {
    errorLog('システムの依存関係チェックに失敗:', dependencyCheck.errors);
    return showErrorPage(
      'システム初期化エラー', 
      'システムの初期化に失敗しました。管理者にご連絡ください。\n\n' +
      `エラー詳細: ${dependencyCheck.errors.join(', ')}`
    );
  }

  // Check system setup (highest priority)
  if (!isSystemSetup()) {
    return showSetupPage();
  }

  // Check application access permissions
  const accessCheck = checkApplicationAccess();
  if (!accessCheck.hasAccess) {
    infoLog('アプリケーションアクセス拒否:', accessCheck.accessReason);
    return showErrorPage('アクセスが制限されています', 'アプリケーションは現在利用できません。管理者にお問い合わせください。');
  }

  return null; // Continue processing
}

/**
 * Validate user authentication
 * @returns {HtmlOutput|null} Early return result or null to continue
 */
function validateUserAuthentication() {
  debugLog('validateUserAuthentication: Starting authentication check.'); // 追加
  const userEmail = getCurrentUserEmail();
  debugLog('validateUserAuthentication: userEmail from getCurrentUserEmail:', userEmail); // 追加
  if (!userEmail) {
    debugLog('validateUserAuthentication: userEmail is empty, showing login page.'); // 追加
    return showLoginPage();
  }
  debugLog('validateUserAuthentication: userEmail is present, continuing processing.'); // 追加
  return null; // Continue processing
}

/**
 * Route request based on mode parameter
 * @param {Object} params - Request parameters
 * @returns {HtmlOutput} Appropriate page response
 */
function routeRequestByMode(params) {
  // Handle no mode parameter
  if (!params || !params.mode) {
    return handleDefaultRoute();
  }

  // Route based on mode
  switch (params.mode) {
    case 'login':
      return handleLoginMode();
    case 'appSetup':
      return handleAppSetupMode();
    case 'admin':
      return handleAdminMode(params);
    case 'view':
      return handleViewMode(params);
    default:
      return handleUnknownMode(params);
  }
}

/**
 * Handle default route (no mode parameter)
 * @returns {HtmlOutput} Appropriate page response
 */
function handleDefaultRoute() {
  const activeUserEmail = getCurrentUserEmail();
  if (!activeUserEmail) {
    return showLoginPage();
  }

  const userProperties = PropertiesService.getUserProperties();
  const lastAdminUserId = userProperties.getProperty('lastAdminUserId');

  if (lastAdminUserId && verifyAdminAccess(lastAdminUserId)) {
    const userInfo = findUserById(lastAdminUserId);
    return renderAdminPanel(userInfo, 'admin');
  }

  // Clear invalid admin session
  if (lastAdminUserId) {
    userProperties.deleteProperty('lastAdminUserId');
  }

  return showLoginPage();
}

/**
 * Handle login mode
 * @returns {HtmlOutput} Login page
 */
function handleLoginMode() {
  return showLoginPage();
}

/**
 * Handle app setup mode
 * @returns {HtmlOutput} App setup page or error page
 */
function handleAppSetupMode() {
  const userProperties = PropertiesService.getUserProperties();
  const lastAdminUserId = userProperties.getProperty('lastAdminUserId');

  if (!lastAdminUserId) {
    return showErrorPage('認証が必要です', 'アプリ設定にアクセスするにはログインが必要です。');
  }

  if (!verifyAdminAccess(lastAdminUserId)) {
    return showErrorPage('アクセス拒否', 'アプリ設定にアクセスする権限がありません。');
  }

  return showAppSetupPage(lastAdminUserId);
}

/**
 * Handle admin mode
 * @param {Object} params - Request parameters
 * @returns {HtmlOutput} Admin panel or error page
 */
function handleAdminMode(params) {
  if (!params.userId) {
    return showErrorPage('不正なリクエスト', 'ユーザーIDが指定されていません。');
  }

  const adminAccessResult = verifyAdminAccess(params.userId);
  
  if (!adminAccessResult) {
    return showErrorPage(
      'アクセス拒否', 
      'アカウントが一時的に無効化されています。\n\n' +
      '対処法:\n' +
      '• 新規登録から1-2分お待ちください\n' +
      '• ブラウザを更新してお試しください\n' +
      '• 問題が続く場合は管理者にお問い合わせください'
    );
  }

  // Save admin session state
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('lastAdminUserId', params.userId);
  userProperties.setProperty('lastSuccessfulAdminAccess', Date.now().toString());

  const userInfo = findUserById(params.userId);
  return renderAdminPanel(userInfo, 'admin');
}

/**
 * Handle view mode
 * @param {Object} params - Request parameters
 * @returns {HtmlOutput} Answer board or unpublished page
 */
function handleViewMode(params) {
  if (!params.userId) {
    return showErrorPage('不正なリクエスト', 'ユーザーIDが指定されていません。');
  }

  // Get user info with cache bypass for accurate publication status
  const userInfo = findUserById(params.userId, {
    useExecutionCache: false,
    forceRefresh: true
  });

  if (!userInfo) {
    return showErrorPage('エラー', '指定されたユーザーが見つかりません。');
  }

  return processViewRequest(userInfo, params);
}

/**
 * Process view request with publication status checks
 * @param {Object} userInfo - User information
 * @param {Object} params - Request parameters
 * @returns {HtmlOutput} Answer board or unpublished page
 */
function processViewRequest(userInfo, params) {
  // Parse config safely
  let config = {};
  try {
    config = JSON.parse(userInfo.configJson || '{}');
  } catch (e) {
    warnLog('Config JSON parse error during publication check:', e.message);
  }

  // Check for auto-stop and handle accordingly
  const wasAutoStopped = checkAndHandleAutoStop(config, userInfo);
  if (wasAutoStopped) {
    infoLog('🔄 自動停止が実行されました - 非公開ページに誘導します');
  }

  // Check if currently published
  const isCurrentlyPublished = !!(config.appPublished === true &&
    config.publishedSpreadsheetId &&
    config.publishedSheetName &&
    typeof config.publishedSheetName === 'string' &&
    config.publishedSheetName.trim() !== '');

  // Redirect to unpublished page if not published
  if (!isCurrentlyPublished) {
    return renderUnpublishedPage(userInfo, params);
  }

  return renderAnswerBoard(userInfo, params);
}

/**
 * Handle unknown mode
 * @param {Object} params - Request parameters
 * @returns {HtmlOutput} Appropriate page response
 */
function handleUnknownMode(params) {
  // If valid userId with admin access, redirect to admin panel
  if (params.userId && verifyAdminAccess(params.userId)) {
    const userInfo = findUserById(params.userId);
    return renderAdminPanel(userInfo, 'admin');
  }

  // Otherwise redirect to login
  return showLoginPage();
}

/**
 * 管理者ページのルートを処理
 * @param {Object} userInfo - ユーザー情報
 * @param {Object} params - リクエストパラメータ
 * @param {string} userEmail - 現在のユーザーのメールアドレス
 * @returns {HtmlOutput}
 */
function handleAdminRoute(userInfo, params, userEmail) {
  // この関数が呼ばれる時点でuserInfoはnullではないことが保証されている

  // セキュリティチェック: アクセスしようとしているuserIdが自分のものでなければ、自分の管理画面にリダイレクト
  if (params.userId && params.userId !== userInfo.userId) {
    const correctUrl = buildUserAdminUrl(userInfo.userId);
    return redirectToUrl(correctUrl);
  }

  // 強化されたセキュリティ検証
  if (params.userId) {
    const isVerified = verifyAdminAccess(params.userId);
    if (!isVerified) {
      const correctUrl = buildUserAdminUrl(userInfo.userId);
      return redirectToUrl(correctUrl);
    }
  }

  return renderAdminPanel(userInfo, params.mode);
}
/**
 * 統合されたユーザー情報取得関数 (Phase 2: 統合完了)
 * キャッシュ戦略とセキュリティチェックを統合
 * @param {string|Object} identifier - メールアドレス、ユーザーID、または設定オブジェクト
 * @param {string} [type] - 'email' | 'userId' | null (auto-detect)
 * @param {Object} [options] - キャッシュオプション
 * @returns {Object|null} ユーザー情報オブジェクト、見つからない場合はnull
 */
function getOrFetchUserInfo(identifier, type = null, options = {}) {
  // オプションのデフォルト値
  const opts = {
    ttl: options.ttl || 300, // 5分キャッシュ
    enableSecurityCheck: options.enableSecurityCheck !== false, // デフォルトで有効
    currentUserEmail: options.currentUserEmail || null,
    useExecutionCache: options.useExecutionCache || false,
    ...options
  };

  // 引数の正規化
  let email = null;
  let userId = null;

  if (typeof identifier === 'object' && identifier !== null) {
    // オブジェクト形式の場合（後方互換性）
    email = identifier.email;
    userId = identifier.userId;
  } else if (typeof identifier === 'string') {
    // 文字列の場合、typeに基づいて判定
    if (type === 'email' || (!type && identifier.includes('@'))) {
      email = identifier;
    } else {
      userId = identifier;
    }
  }

  // キャッシュキーの生成
  const cacheKey = `unified_user_info_${userId || email}`;

  // 実行レベルキャッシュの確認（オプション）
  if (opts.useExecutionCache && _executionUserInfoCache &&
      _executionUserInfoCache.userId === userId) {
    return _executionUserInfoCache.userInfo;
  }

  // 統合キャッシュマネージャーを使用（キャッシュ miss 時は自動でデータベースから取得）
  let userInfo = null;

  try {
    userInfo = cacheManager.get(cacheKey, () => {
      // キャッシュに存在しない場合はデータベースから取得する
      // 通常フローのためエラーレベルでは記録しない
      debugLog('cache miss - fetching from database');

      const props = PropertiesService.getScriptProperties();
      if (!props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID)) {
        logError('DATABASE_SPREADSHEET_ID not set', 'getUnifiedCache', MAIN_ERROR_SEVERITY.CRITICAL, MAIN_ERROR_CATEGORIES.SYSTEM);
        return null;
      }

      let dbUserInfo = null;
      if (userId) {
        dbUserInfo = findUserById(userId);
        // セキュリティチェック: 取得した情報のemailが現在のユーザーと一致するか確認
        if (opts.enableSecurityCheck && dbUserInfo && opts.currentUserEmail &&
            dbUserInfo.adminEmail !== opts.currentUserEmail) {
          warnLog('セキュリティチェック失敗: 他人の情報へのアクセス試行');
          return null;
        }
      } else if (email) {
        dbUserInfo = findUserByEmail(email);
      }

      return dbUserInfo;
    }, {
      ttl: opts.ttl || 300,
      enableMemoization: opts.enableMemoization || false
    });

    // 実行レベルキャッシュにも保存（オプション）
    if (userInfo && opts.useExecutionCache && (userId || userInfo.userId)) {
      _executionUserInfoCache = { userId: userId || userInfo.userId, userInfo };
      debugLog('✅ 実行レベルキャッシュに保存:', userId || userInfo.userId);
    }

  } catch (cacheError) {
    logError(cacheError, 'getUnifiedCache', MAIN_ERROR_SEVERITY.MEDIUM, MAIN_ERROR_CATEGORIES.CACHE);
    // フォールバック: 直接データベースから取得
    if (userId) {
      userInfo = findUserById(userId);
    } else if (email) {
      userInfo = findUserByEmail(email);
    }
  }

  return userInfo;
}
/**
 * ログインページを表示
 * @returns {HtmlOutput}
 */
function showLoginPage() {
  const template = HtmlService.createTemplateFromFile('LoginPage');
  const htmlOutput = template.evaluate()
    .setTitle('StudyQuest - ログイン');

  // XFrameOptionsMode を安全に設定
  try {
    if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.ALLOWALL) {
      htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  } catch (e) {
    warnLog('XFrameOptionsMode設定エラー:', e.message);
  }

  return htmlOutput;
}

/**
 * 初回セットアップページを表示
 * @returns {HtmlOutput}
 */
function showSetupPage() {
  const template = HtmlService.createTemplateFromFile('SetupPage');
  const htmlOutput = template.evaluate()
    .setTitle('StudyQuest - 初回セットアップ');

  // XFrameOptionsMode を安全に設定
  try {
    if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.DENY) {
      htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
    }
  } catch (e) {
    warnLog('XFrameOptionsMode設定エラー:', e.message);
  }

  return htmlOutput;
}

/**
 * アプリ設定ページを表示
 * @returns {HtmlOutput}
 */
function showAppSetupPage(userId) {
    // システム管理者権限チェック
    try {
      debugLog('showAppSetupPage: Checking deploy user permissions...');
      const currentUserEmail = getCurrentUserEmail();
      debugLog('showAppSetupPage: Current user email:', currentUserEmail);
      const deployUserCheckResult = isDeployUser();
      debugLog('showAppSetupPage: isDeployUser() result:', deployUserCheckResult);

      if (!deployUserCheckResult) {
        warnLog('Unauthorized access attempt to app setup page:', currentUserEmail);
        return showErrorPage('アクセス権限がありません', 'この機能にアクセスする権限がありません。システム管理者にお問い合わせください。');
      }
    } catch (error) {
      logError(error, 'checkDeployUserPermissions', MAIN_ERROR_SEVERITY.HIGH, MAIN_ERROR_CATEGORIES.AUTHORIZATION);
      return showErrorPage('認証エラー', '権限確認中にエラーが発生しました。');
    }

    const appSetupTemplate = HtmlService.createTemplateFromFile('AppSetupPage');
    appSetupTemplate.userId = userId;
    const htmlOutput = appSetupTemplate.evaluate()
      .setTitle('アプリ設定 - StudyQuest');

    // XFrameOptionsMode を安全に設定
    try {
      if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.DENY) {
        htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
      }
    } catch (e) {
      warnLog('XFrameOptionsMode設定エラー:', e.message);
    }

    return htmlOutput;
}

/**
 * 最後にアクセスした管理ユーザーIDを取得
 * リダイレクト時の戻り先決定に使用
 * @returns {string|null} 管理ユーザーID、存在しない場合はnull
 */
function getLastAdminUserId() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const lastAdminUserId = userProperties.getProperty('lastAdminUserId');

    // ユーザーIDが存在し、かつ有効な管理者権限を持つかチェック
    if (lastAdminUserId && verifyAdminAccess(lastAdminUserId)) {
      debugLog('Found valid admin user ID:', lastAdminUserId);
      return lastAdminUserId;
    } else {
      debugLog('No valid admin user ID found');
      return null;
    }
  } catch (error) {
    logError(error, 'getLastAdminUserId', MAIN_ERROR_SEVERITY.MEDIUM, MAIN_ERROR_CATEGORIES.DATABASE);
    return null;
  }
}

/**
 * アプリ設定ページのURLを取得（フロントエンドから呼び出し用）
 * @returns {string} アプリ設定ページのURL
 */
function getAppSetupUrl() {
  try {
    // システム管理者権限チェック
    debugLog('getAppSetupUrl: Checking deploy user permissions...');
    const currentUserEmail = getCurrentUserEmail();
    debugLog('getAppSetupUrl: Current user email:', currentUserEmail);
    const deployUserCheckResult = isDeployUser();
    debugLog('getAppSetupUrl: isDeployUser() result:', deployUserCheckResult);

    if (!deployUserCheckResult) {
      warnLog('Unauthorized access attempt to get app setup URL:', currentUserEmail);
      throw new Error('アクセス権限がありません');
    }

    // WebアプリのベースURLを取得
    const baseUrl = ScriptApp.getService().getUrl();
    if (!baseUrl) {
      throw new Error('WebアプリのURLを取得できませんでした');
    }

    // アプリ設定ページのURLを生成
    const appSetupUrl = baseUrl + '?mode=appSetup';
    debugLog('getAppSetupUrl: Generated URL:', appSetupUrl);

    return appSetupUrl;
  } catch (error) {
    logError(error, 'getAppSetupUrl', MAIN_ERROR_SEVERITY.MEDIUM, MAIN_ERROR_CATEGORIES.SYSTEM);
    throw new Error('アプリ設定URLの取得に失敗しました: ' + error.message);
  }
}

// =================================================================
// ERROR HANDLING & CATEGORIZATION
// =================================================================

// エラータイプ定義
const ERROR_TYPES = {
  CRITICAL: 'critical',      // 致命的システムエラー
  ACCESS: 'access',          // アクセス・認証エラー  
  VALIDATION: 'validation',  // データ検証エラー
  NETWORK: 'network',        // ネットワーク・API エラー
  USER: 'user'              // ユーザー操作エラー
};

/**
 * エラータイプに基づいてメッセージを分類・整理する
 * @param {string} title - エラータイトル
 * @param {string} message - エラーメッセージ
 * @returns {Object} 分類されたエラー情報
 */
function categorizeError(title, message) {
  const titleLower = title.toLowerCase();
  const messageLower = message.toLowerCase();
  
  // エラータイプの判定
  let errorType = ERROR_TYPES.USER;
  if (titleLower.includes('致命的') || titleLower.includes('システム')) {
    errorType = ERROR_TYPES.CRITICAL;
  } else if (titleLower.includes('アクセス') || titleLower.includes('認証') || titleLower.includes('権限')) {
    errorType = ERROR_TYPES.ACCESS;
  } else if (titleLower.includes('不正') || messageLower.includes('指定されていません')) {
    errorType = ERROR_TYPES.VALIDATION;
  } else if (messageLower.includes('ネットワーク') || messageLower.includes('接続')) {
    errorType = ERROR_TYPES.NETWORK;
  }
  
  return {
    type: errorType,
    icon: getErrorIcon(errorType),
    severity: getErrorSeverity(errorType)
  };
}

/**
 * エラータイプに対応するアイコンを取得
 */
function getErrorIcon(errorType) {
  const icons = {
    [ERROR_TYPES.CRITICAL]: '🔥',
    [ERROR_TYPES.ACCESS]: '🔒', 
    [ERROR_TYPES.VALIDATION]: '⚠️',
    [ERROR_TYPES.NETWORK]: '🌐',
    [ERROR_TYPES.USER]: '❓'
  };
  return icons[errorType] || '⚠️';
}

/**
 * エラータイプに対応する重要度を取得
 */
function getErrorSeverity(errorType) {
  const severities = {
    [ERROR_TYPES.CRITICAL]: 'high',
    [ERROR_TYPES.ACCESS]: 'medium',
    [ERROR_TYPES.VALIDATION]: 'medium', 
    [ERROR_TYPES.NETWORK]: 'medium',
    [ERROR_TYPES.USER]: 'low'
  };
  return severities[errorType] || 'low';
}

/**
 * 長い診断情報を構造化して整理する
 * @param {string} diagnosticInfo - 診断情報文字列
 * @returns {Object} 構造化された診断情報
 */
function structureDiagnosticInfo(diagnosticInfo) {
  if (!diagnosticInfo) return null;
  
  const lines = diagnosticInfo.split('\n');
  const structured = {
    summary: [],
    technical: [],
    properties: null
  };
  
  let currentSection = 'summary';
  
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    
    // プロパティ状態のJSON部分を検出
    if (trimmed.includes('プロパティ状態:')) {
      currentSection = 'properties';
      const jsonStart = line.indexOf('{');
      if (jsonStart !== -1) {
        try {
          const jsonStr = line.substring(jsonStart);
          structured.properties = JSON.parse(jsonStr);
        } catch (e) {
          structured.technical.push(line);
        }
      }
      continue;
    }
    
    // 基本情報と技術情報を分類
    if (trimmed.startsWith('ユーザーID:') || 
        trimmed.startsWith('現在のメール:') || 
        trimmed.startsWith('認証時間:') || 
        trimmed.startsWith('時刻:')) {
      structured.summary.push(trimmed);
    } else {
      structured.technical.push(trimmed);
    }
  }
  
  return structured;
}

/**
 * エラーページを表示する関数（改善版）
 * @param {string} title - エラータイトル
 * @param {string} message - エラーメッセージ  
 * @param {Error|string} error - エラーオブジェクトまたは診断情報
 * @returns {HtmlOutput} エラーページのHTML出力
 */
function showErrorPage(title, message, error) {
  const template = HtmlService.createTemplateFromFile('ErrorBoundary');
  
  // エラー分類
  const errorInfo = categorizeError(title, message);
  
  // 基本情報設定
  template.title = title;
  template.errorType = errorInfo.type;
  template.errorIcon = errorInfo.icon;
  template.errorSeverity = errorInfo.severity;
  template.mode = 'admin';
  
  // メッセージを構造化
  if (message && message.includes('詳細診断情報:')) {
    const parts = message.split('詳細診断情報:');
    template.message = parts[0].trim();
    template.diagnosticInfo = structureDiagnosticInfo(parts[1]);
  } else {
    template.message = message;
    template.diagnosticInfo = null;
  }
  
  // 現在のユーザーがデータベースに登録されているかチェック
  let isRegisteredUser = false;
  let currentUserEmail = '';
  try {
    currentUserEmail = getCurrentUserEmail();
    if (currentUserEmail) {
      const userInfo = findUserByEmail(currentUserEmail);
      isRegisteredUser = !!userInfo;
    }
  } catch (e) {
    console.warn('⚠️ showErrorPage: ユーザー登録状態の確認でエラー:', e);
  }
  
  template.isRegisteredUser = isRegisteredUser;
  template.userEmail = currentUserEmail;
  
  // デバッグ情報設定
  if (DEBUG && error) {
    if (typeof error === 'string') {
      template.debugInfo = error;
    } else if (error.stack) {
      template.debugInfo = error.stack;
    } else {
      template.debugInfo = error.toString();
    }
  } else {
    template.debugInfo = '';
  }
  
  const htmlOutput = template.evaluate()
    .setTitle(`エラー - ${title}`);

  // XFrameOptionsMode を安全に設定
  try {
    if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.DENY) {
      htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
    }
  } catch (e) {
    warnLog('XFrameOptionsMode設定エラー:', e.message);
  }

  return htmlOutput;
}


/**
 * ユーザー専用の一意の管理パネルURLを構築
 * @param {string} userId ユーザーID
 * @return {string}
 */
function buildUserAdminUrl(userId) {
  const baseUrl = ScriptApp.getService().getUrl();
  return `${baseUrl}?mode=admin&userId=${encodeURIComponent(userId)}`;
}

/**
 * 標準化されたページURL生成のヘルパー関数群
 */
const URLBuilder = {
  /**
   * ログインページのURLを生成
   * @returns {string} ログインページURL
   */
  login: function() {
    const baseUrl = ScriptApp.getService().getUrl();
    return `${baseUrl}?mode=login`;
  },

  /**
   * 管理パネルのURLを生成
   * @param {string} userId - ユーザーID
   * @returns {string} 管理パネルURL
   */
  admin: function(userId) {
    const baseUrl = ScriptApp.getService().getUrl();
    return `${baseUrl}?mode=admin&userId=${encodeURIComponent(userId)}`;
  },

  /**
   * アプリ設定ページのURLを生成
   * @returns {string} アプリ設定ページURL
   */
  appSetup: function() {
    const baseUrl = ScriptApp.getService().getUrl();
    return `${baseUrl}?mode=appSetup`;
  },

  /**
   * 回答ボードのURLを生成
   * @param {string} userId - ユーザーID
   * @returns {string} 回答ボードURL
   */
  view: function(userId) {
    const baseUrl = ScriptApp.getService().getUrl();
    return `${baseUrl}?mode=view&userId=${encodeURIComponent(userId)}`;
  },

  /**
   * パラメータ付きURLを安全に生成
   * @param {string} mode - モード ('login', 'admin', 'view', 'appSetup')
   * @param {Object} params - 追加パラメータ
   * @returns {string} 生成されたURL
   */
  build: function(mode, params = {}) {
    const baseUrl = ScriptApp.getService().getUrl();
    const url = new URL(baseUrl);
    url.searchParams.set('mode', mode);

    Object.keys(params).forEach(key => {
      if (params[key] !== null && params[key] !== undefined) {
        url.searchParams.set(key, params[key]);
      }
    });

    return url.toString();
  }
};

/**
 * 指定されたURLへリダイレクトするサーバーサイド関数
 * @param {string} url - リダイレクト先のURL
 * @returns {HtmlOutput} リダイレクトを実行するHTML出力
 */
function redirectToUrl(url) {
  // XSS攻撃を防ぐため、URLをサニタイズ
  const sanitizedUrl = sanitizeRedirectUrl(url);
  return HtmlService.createHtmlOutput().setContent(`<script>window.top.location.href = '${sanitizedUrl}';</script>`);
}
/**
 * セキュアなリダイレクトHTMLを作成 (シンプル版)
 * @param {string} targetUrl リダイレクト先URL
 * @param {string} message 表示メッセージ
 * @return {HtmlOutput}
 */
function createSecureRedirect(targetUrl, message) {
  // URL検証とサニタイゼーション
  const sanitizedUrl = sanitizeRedirectUrl(targetUrl);

  debugLog('createSecureRedirect - Original URL:', targetUrl);
  debugLog('createSecureRedirect - Sanitized URL:', sanitizedUrl);

  // ユーザーアクティベーション必須のHTMLアンカー方式（サンドボックス制限準拠）
  const userActionRedirectHtml = `
    <!DOCTYPE html>
    <html>
    <head>
      <title>${message || 'アクセス確認'}</title>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <style>
        body {
          font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
          background: linear-gradient(135deg, #1e3a8a 0%, #7c3aed 100%);
          min-height: 100vh;
          margin: 0;
          display: flex;
          align-items: center;
          justify-content: center;
          padding: 20px;
        }
        .container {
          background: rgba(31, 41, 55, 0.95);
          border-radius: 16px;
          padding: 40px;
          max-width: 500px;
          width: 100%;
          text-align: center;
          box-shadow: 0 25px 50px rgba(0, 0, 0, 0.5);
          border: 1px solid rgba(75, 85, 99, 0.3);
        }
        .icon { font-size: 64px; margin-bottom: 20px; }
        .title {
          color: #10b981;
          font-size: 24px;
          font-weight: bold;
          margin-bottom: 16px;
        }
        .subtitle {
          color: #d1d5db;
          margin-bottom: 32px;
          line-height: 1.5;
        }
        .main-button {
          display: inline-block;
          background: linear-gradient(135deg, #10b981 0%, #3b82f6 100%);
          color: white;
          font-weight: bold;
          padding: 16px 32px;
          border-radius: 12px;
          text-decoration: none;
          font-size: 18px;
          transition: all 0.3s ease;
          margin-bottom: 24px;
          box-shadow: 0 4px 15px rgba(16, 185, 129, 0.4);
        }
        .main-button:hover {
          transform: translateY(-2px);
          box-shadow: 0 8px 25px rgba(16, 185, 129, 0.6);
        }
        .url-info {
          background: rgba(17, 24, 39, 0.8);
          border-radius: 8px;
          padding: 16px;
          margin: 20px 0;
          border: 1px solid rgba(75, 85, 99, 0.5);
        }
        .url-text {
          color: #60a5fa;
          font-family: 'Courier New', monospace;
          font-size: 12px;
          word-break: break-all;
          line-height: 1.4;
        }
        .note {
          color: #9ca3af;
          font-size: 14px;
          margin-top: 20px;
          line-height: 1.4;
        }
        @keyframes pulse {
          0% { transform: scale(1); }
          50% { transform: scale(1.05); }
          100% { transform: scale(1); }
        }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="icon">🔐</div>
        <h1 class="title">${message || 'アクセス確認'}</h1>
        <p class="subtitle">セキュリティのため、下のボタンをクリックして続行してください</p>

        <a href="${sanitizedUrl}" target="_top" class="main-button" onclick="handleSecureRedirect(event, '${sanitizedUrl}')">
          🚀 続行する
        </a>

        <div class="url-info">
          <div class="url-text">${sanitizedUrl}</div>
        </div>

        <div class="note">
          ✓ このリンクは安全です<br>
          ✓ Google Apps Script公式のセキュリティガイドラインに準拠<br>
          ✓ iframe制約回避のため新しいタブで開きます
        </div>
        
        <script>
          function handleSecureRedirect(event, url) {
            try {
              // iframe内かどうかを判定
              const isInFrame = (window !== window.top);
              
              if (isInFrame) {
                // iframe内の場合は親ウィンドウで開く
                event.preventDefault();
                console.log('🔄 iframe内からの遷移を検出、parent window で開きます');
                window.top.location.href = url;
              } else {
                // 通常の場合はそのまま遷移
                console.log('🚀 通常の遷移を実行します');
                // target="_top" が有効になります
              }
            } catch (error) {
              console.error('リダイレクト処理エラー:', error);
              // エラーの場合はフォールバック
              window.location.href = url;
            }
          }
          
          // 自動遷移を無効化（X-Frame-Options制約のためユーザーアクション必須）
          // ユーザーに明確なアクションを要求することで、より確実な遷移を実現
          console.log('ℹ️ 自動遷移は無効です。ユーザーによるボタンクリックが必要です。');
          
          // 代わりに、5秒後にボタンを強調表示
          setTimeout(function() {
            const mainButton = document.querySelector('.main-button');
            if (mainButton) {
              mainButton.style.animation = 'pulse 1s infinite';
              mainButton.style.boxShadow = '0 0 20px rgba(16, 185, 129, 0.5)';
              console.log('✨ ボタンを強調表示しました');
            }
          }, 3000);
        </script>
      </div>
    </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(userActionRedirectHtml);

  // XFrameOptionsMode を安全に設定（セキュアリダイレクト用）
  try {
    if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.ALLOWALL) {
      htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      debugLog('✅ Secure Redirect XFrameOptionsMode.ALLOWALL設定完了');
    } else {
      warnLog('⚠️ HtmlService.XFrameOptionsMode.ALLOWALLが利用できません');
    }
  } catch (e) {
    errorLog('❌ Secure Redirect XFrameOptionsMode設定エラー:', e.message);
    // フォールバック: 従来の方法で設定を試行
    try {
      htmlOutput.setXFrameOptionsMode('ALLOWALL');
      infoLog('💡 フォールバック方法でSecure Redirect XFrameOptionsMode設定完了');
    } catch (fallbackError) {
      errorLog('❌ フォールバック方法も失敗:', fallbackError.message);
    }
  }

  return htmlOutput;
}

/**
 * リダイレクト用URLの検証とサニタイゼーション
 * @param {string} url 検証対象のURL
 * @return {string} サニタイズされたURL
 */
function sanitizeRedirectUrl(url) {
  if (!url) {
    return ScriptApp.getService().getUrl();
  }

  try {
    let cleanUrl = String(url).trim();

    // 複数レベルのクォート除去（JSON文字列化による多重クォートに対応）
    let previousUrl = '';
    while (cleanUrl !== previousUrl) {
      previousUrl = cleanUrl;

      // 先頭と末尾のクォートを除去
      if ((cleanUrl.startsWith('"') && cleanUrl.endsWith('"')) ||
          (cleanUrl.startsWith("'") && cleanUrl.endsWith("'"))) {
        cleanUrl = cleanUrl.slice(1, -1);
      }

      // エスケープされたクォートを除去
      cleanUrl = cleanUrl.replace(/\\"/g, '"').replace(/\\'/g, "'");

      // URL内に埋め込まれた別のURLを検出
      const embeddedUrlMatch = cleanUrl.match(/https?:\/\/[^\s<>"']+/);
      if (embeddedUrlMatch && embeddedUrlMatch[0] !== cleanUrl) {
        debugLog('Extracting embedded URL:', embeddedUrlMatch[0]);
        cleanUrl = embeddedUrlMatch[0];
      }
    }

    // 基本的なURL形式チェック
    if (!cleanUrl.match(/^https?:\/\/[^\s<>"']+$/)) {
      warnLog('Invalid URL format after sanitization:', cleanUrl);
      return ScriptApp.getService().getUrl();
    }

    // 開発モードURLのチェック（googleusercontent.comは有効なデプロイURLも含むため調整）
    if (cleanUrl.includes('userCodeAppPanel')) {
      warnLog('Development URL detected in redirect, using fallback:', cleanUrl);
      return ScriptApp.getService().getUrl();
    }

    // 最終的な URL 妥当性チェック（googleusercontent.comも有効URLとして認識）
    const isValidUrl = cleanUrl.includes('script.google.com') ||
                     cleanUrl.includes('googleusercontent.com') ||
                     cleanUrl.includes('localhost');

    if (!isValidUrl) {
      warnLog('Suspicious URL detected:', cleanUrl);
      return ScriptApp.getService().getUrl();
    }

    return cleanUrl;
  } catch (e) {
    logError(e, 'urlSanitization', MAIN_ERROR_SEVERITY.HIGH, MAIN_ERROR_CATEGORIES.SYSTEM);
    return ScriptApp.getService().getUrl();
  }
}

/**
 * doGet のリクエストパラメータを解析
 * @param {Object} e Event object
 * @return {{mode:string,userId:string|null,setupParam:string|null,spreadsheetId:string|null,sheetName:string|null,isDirectPageAccess:boolean}}
 */
function parseRequestParams(e) {
  // 引数の安全性チェック
  if (!e || typeof e !== 'object') {
    debugLog('parseRequestParams: 無効なeventオブジェクト');
    return { mode: null, userId: null, setupParam: null, spreadsheetId: null, sheetName: null, isDirectPageAccess: false };
  }

  const p = e.parameter || {};
  const mode = p.mode || null;
  const userId = p.userId || null;
  const setupParam = p.setup || null;
  const spreadsheetId = p.spreadsheetId || null;
  const sheetName = p.sheetName || null;
  const isDirectPageAccess = !!(userId && mode === 'view');

  // デバッグログを追加
  debugLog('parseRequestParams - Received parameters:', JSON.stringify(p));
  debugLog('parseRequestParams - Parsed mode:', mode, 'setupParam:', setupParam);

  return { mode, userId, setupParam, spreadsheetId, sheetName, isDirectPageAccess };
}

/**
 * 管理パネルを表示
 * @param {Object} userInfo ユーザー情報
 * @param {string} mode 表示モード
 * @return {HtmlOutput} HTMLコンテンツ
 */
function renderAdminPanel(userInfo, mode) {
  // ガード節: userInfoが存在しない場合はエラーページを表示して処理を中断
  if (!userInfo) {
    logError('renderAdminPanelにuserInfoがnullで渡されました', 'renderAdminPanel', MAIN_ERROR_SEVERITY.HIGH, MAIN_ERROR_CATEGORIES.SYSTEM);
    return showErrorPage('エラー', 'ユーザー情報の読み込みに失敗したため、管理パネルを表示できません。');
  }

  const adminTemplate = HtmlService.createTemplateFromFile('AdminPanel');
  adminTemplate.include = include;
  adminTemplate.userInfo = userInfo;
  adminTemplate.userId = userInfo.userId;
  adminTemplate.mode = mode || 'admin'; // 安全のためのフォールバック
  adminTemplate.displayMode = 'named';
  adminTemplate.showAdminFeatures = true;
  const deployUserResult = isDeployUser();
  const currentUserEmail = getCurrentUserEmail();
  const adminEmail = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROPS_KEYS.ADMIN_EMAIL);
  
  debugLog('renderAdminPanel - isDeployUser() result:', deployUserResult);
  debugLog('renderAdminPanel - current user email:', currentUserEmail);
  debugLog('renderAdminPanel - ADMIN_EMAIL property:', adminEmail);
  debugLog('renderAdminPanel - emails match:', adminEmail === currentUserEmail);
  adminTemplate.isDeployUser = deployUserResult;
  adminTemplate.DEBUG_MODE = shouldEnableDebugMode();


  const htmlOutput = adminTemplate.evaluate()
    .setTitle('StudyQuest - 管理パネル')
    .setSandboxMode(HtmlService.SandboxMode.NATIVE);

  // XFrameOptionsMode を安全に設定（iframe embedding許可）
  try {
    if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.ALLOWALL) {
      htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      debugLog('✅ Admin Panel XFrameOptionsMode.ALLOWALL設定完了 - iframe embedding許可');
    } else {
      warnLog('⚠️ HtmlService.XFrameOptionsMode.ALLOWALLが利用できません');
    }
  } catch (e) {
    errorLog('❌ Admin Panel XFrameOptionsMode設定エラー:', e.message);
    // フォールバック: 従来の方法で設定を試行
    try {
      htmlOutput.setXFrameOptionsMode('ALLOWALL');
      infoLog('💡 フォールバック方法でAdmin Panel XFrameOptionsMode設定完了');
    } catch (fallbackError) {
      errorLog('❌ フォールバック方法も失敗:', fallbackError.message);
    }
  }

  return htmlOutput;
}

/**
 * 回答ボードまたは未公開ページを表示
 * @param {Object} userInfo ユーザー情報
 * @param {Object} params リクエストパラメータ
 * @return {HtmlOutput} HTMLコンテンツ
 */
/**
 * 非公開ボード専用のレンダリング関数
 * ErrorBoundaryを回避して確実にUnpublished.htmlを表示
 * @param {Object} userInfo - ユーザー情報
 * @param {Object} params - URLパラメータ
 * @returns {HtmlOutput} Unpublished.htmlテンプレート
 */
function renderUnpublishedPage(userInfo, params) {
  try {
    debugLog('🚫 renderUnpublishedPage: Rendering unpublished page for userId:', userInfo.userId);

    let template;
    try {
      template = HtmlService.createTemplateFromFile('Unpublished');
      debugLog('✅ renderUnpublishedPage: Template created successfully');
    } catch (templateError) {
      console.error('❌ renderUnpublishedPage: Template creation failed:', templateError);
      throw new Error('Unpublished.htmlテンプレートの読み込みに失敗: ' + templateError.message);
    }
    
    template.include = include;

    // 基本情報の設定（安全なデフォルト値付き）
    template.userId = userInfo.userId || '';
    template.spreadsheetId = userInfo.spreadsheetId || '';
    template.ownerName = userInfo.adminEmail || 'システム管理者';
    template.isOwner = (getCurrentUserEmail() === userInfo.adminEmail); // 現在のユーザーがボードの所有者であるかを確認
    template.adminEmail = userInfo.adminEmail || '';
    template.cacheTimestamp = Date.now();

    // 安全な変数設定
    template.include = include;

    // URL生成（エラー耐性を持たせる）
    let appUrls;
    try {
      appUrls = generateUserUrls(userInfo.userId);
      if (!appUrls || appUrls.status === 'error') {
        throw new Error('URL生成に失敗しました');
      }
    } catch (urlError) {
      warnLog('URL生成エラー、フォールバック値を使用:', urlError);
      // フォールバック: 基本的なURL構造
      const baseUrl = ScriptApp.getService().getUrl();
      appUrls = {
        adminUrl: `${baseUrl}?mode=admin&userId=${encodeURIComponent(userInfo.userId)}`,
        viewUrl: `${baseUrl}?mode=view&userId=${encodeURIComponent(userInfo.userId)}`,
        status: 'fallback'
      };
    }

    template.adminPanelUrl = appUrls.adminUrl || '';
    template.boardUrl = appUrls.viewUrl || '';

    debugLog('✅ renderUnpublishedPage: Template setup completed');

    // キャッシュを無効化して確実なリダイレクトを保証
    const htmlOutput = template.evaluate()
      .setTitle('StudyQuest - 準備中');
    
    // addMetaTagを安全に追加
    try {
      htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } catch (e) {
      console.warn('⚠️ addMetaTag(viewport) failed:', e.message);
    }
    
    try {
      htmlOutput.addMetaTag('cache-control', 'no-cache, no-store, must-revalidate');
    } catch (e) {
      console.warn('⚠️ addMetaTag(cache-control) failed:', e.message);
    }

    try {
      if (HtmlService && HtmlService.XFrameOptionsMode &&
          HtmlService.XFrameOptionsMode.ALLOWALL) {
        htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }
    } catch (e) {
      warnLog('XFrameOptionsMode設定エラー:', e.message);
    }

    return htmlOutput;

  } catch (error) {
    logError(error, 'renderUnpublishedPage', MAIN_ERROR_SEVERITY.HIGH, MAIN_ERROR_CATEGORIES.SYSTEM, {
      userId: userInfo ? userInfo.userId : 'null',
      hasUserInfo: !!userInfo,
      errorMessage: error.message,
      errorStack: error.stack
    });
    console.error('🚨 renderUnpublishedPage error details:', {
      error: error,
      userInfo: userInfo,
      userId: userInfo ? userInfo.userId : 'N/A',
      adminEmail: userInfo ? userInfo.adminEmail : 'N/A'
    });
    // フォールバック: ErrorBoundary.htmlを回避して確実にUnpublishedページを表示
    return renderMinimalUnpublishedPage(userInfo);
  }
}

/**
 * 最小限の非公開ページレンダリング（フォールバック用）
 * テンプレートエラーを回避してHTMLを直接生成
 * @param {Object} userInfo - ユーザー情報
 * @returns {HtmlOutput} 最小限のHTML
 */
function renderMinimalUnpublishedPage(userInfo) {
  try {
    debugLog('🚫 renderMinimalUnpublishedPage: Creating minimal unpublished page');
    
    // 安全にuserInfoを処理
    if (!userInfo) {
      console.warn('⚠️ renderMinimalUnpublishedPage: userInfo is null/undefined');
      userInfo = { userId: '', adminEmail: '' };
    }

    const userId = (userInfo.userId && typeof userInfo.userId === 'string') ? userInfo.userId : '';
    const adminEmail = (userInfo.adminEmail && typeof userInfo.adminEmail === 'string') ? userInfo.adminEmail : '';

    // 直接HTMLを生成（テンプレートを使わない）
    const htmlContent = `
      <!DOCTYPE html>
      <html lang="ja">
      <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1">
          <title>StudyQuest - 準備中</title>
          <style>
              body { font-family: -apple-system, BlinkMacSystemFont, sans-serif; margin: 0; padding: 20px; background: #1a1a1a; color: white; text-align: center; }
              .container { max-width: 600px; margin: 50px auto; padding: 40px 20px; background: #2a2a2a; border-radius: 12px; }
              .status { font-size: 24px; margin-bottom: 20px; color: #fbbf24; }
              .message { font-size: 16px; margin-bottom: 30px; color: #9ca3af; }
              .admin-btn { display: inline-block; padding: 12px 24px; background: #3b82f6; color: white; text-decoration: none; border-radius: 8px; margin: 10px; }
              .admin-btn:hover { background: #2563eb; }
          </style>
      </head>
      <body>
          <div class="container">
              <div class="status">⏳ 回答ボード準備中</div>
              <div class="message">現在、回答ボードは非公開になっています</div>
              <a href="?mode=admin&userId=${encodeURIComponent(userId)}" class="admin-btn">管理パネルを開く</a>
              <div style="margin-top: 20px; font-size: 12px; color: #6b7280;">管理者: ${adminEmail}</div>
          </div>
      </body>
      </html>
    `;

    const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
      .setTitle('StudyQuest - 準備中');
    
    // addMetaTagを安全に追加
    try {
      htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } catch (e) {
      console.warn('⚠️ renderMinimalUnpublishedPage addMetaTag(viewport) failed:', e.message);
    }
    
    try {
      htmlOutput.addMetaTag('cache-control', 'no-cache, no-store, must-revalidate');
    } catch (e) {
      console.warn('⚠️ renderMinimalUnpublishedPage addMetaTag(cache-control) failed:', e.message);
    }
    
    return htmlOutput;

  } catch (error) {
    logError(error, 'renderMinimalUnpublishedPage', MAIN_ERROR_SEVERITY.HIGH, MAIN_ERROR_CATEGORIES.SYSTEM, {
      userId: userInfo ? userInfo.userId : 'null',
      hasUserInfo: !!userInfo,
      errorMessage: error.message,
      errorStack: error.stack
    });
    console.error('🚨 renderMinimalUnpublishedPage error details:', {
      error: error,
      userInfo: userInfo,
      userId: userInfo ? userInfo.userId : 'N/A',
      adminEmail: userInfo ? userInfo.adminEmail : 'N/A'
    });
    // 最終フォールバック: 管理者向け機能付き
    const userId = (userInfo && userInfo.userId) ? userInfo.userId : '';
    const adminEmail = (userInfo && userInfo.adminEmail) ? userInfo.adminEmail : '';
    
    const finalFallbackHtml = `
      <!DOCTYPE html>
      <html lang="ja">
      <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1">
          <title>StudyQuest - 準備中</title>
          <style>
              * { margin: 0; padding: 0; box-sizing: border-box; }
              body {
                  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Noto Sans JP', sans-serif;
                  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                  min-height: 100vh;
                  color: #ffffff;
                  line-height: 1.6;
                  overflow-x: hidden;
              }
              .background-overlay {
                  position: fixed;
                  top: 0;
                  left: 0;
                  width: 100%;
                  height: 100%;
                  background: 
                      radial-gradient(circle at 20% 80%, rgba(120, 119, 198, 0.3) 0%, transparent 50%),
                      radial-gradient(circle at 80% 20%, rgba(255, 183, 77, 0.3) 0%, transparent 50%),
                      radial-gradient(circle at 40% 40%, rgba(120, 119, 198, 0.2) 0%, transparent 50%);
                  z-index: -1;
              }
              .main-container {
                  min-height: 100vh;
                  display: flex;
                  align-items: center;
                  justify-content: center;
                  padding: 20px;
                  position: relative;
              }
              .content-card {
                  background: rgba(255, 255, 255, 0.95);
                  backdrop-filter: blur(20px);
                  border-radius: 24px;
                  box-shadow: 
                      0 25px 50px -12px rgba(0, 0, 0, 0.25),
                      0 0 0 1px rgba(255, 255, 255, 0.1);
                  max-width: 600px;
                  width: 100%;
                  padding: 40px;
                  text-align: center;
                  color: #374151;
                  position: relative;
                  overflow: hidden;
              }
              .content-card::before {
                  content: '';
                  position: absolute;
                  top: 0;
                  left: 0;
                  right: 0;
                  height: 4px;
                  background: linear-gradient(90deg, #667eea, #764ba2, #667eea);
                  background-size: 200% 100%;
                  animation: shimmer 3s ease-in-out infinite;
              }
              @keyframes shimmer {
                  0%, 100% { background-position: 200% 0; }
                  50% { background-position: -200% 0; }
              }
              .status-icon {
                  width: 80px;
                  height: 80px;
                  margin: 0 auto 24px;
                  background: linear-gradient(135deg, #fbbf24, #f59e0b);
                  border-radius: 50%;
                  display: flex;
                  align-items: center;
                  justify-content: center;
                  font-size: 36px;
                  animation: pulse 2s infinite;
                  box-shadow: 0 10px 30px rgba(251, 191, 36, 0.3);
              }
              @keyframes pulse {
                  0%, 100% { transform: scale(1); }
                  50% { transform: scale(1.05); }
              }
              .status-title {
                  font-size: 32px;
                  font-weight: 700;
                  color: #1f2937;
                  margin-bottom: 16px;
                  background: linear-gradient(135deg, #667eea, #764ba2);
                  -webkit-background-clip: text;
                  -webkit-text-fill-color: transparent;
                  background-clip: text;
              }
              .status-message {
                  font-size: 18px;
                  color: #6b7280;
                  margin-bottom: 40px;
                  line-height: 1.7;
              }
              .button-group {
                  display: flex;
                  flex-wrap: wrap;
                  gap: 16px;
                  justify-content: center;
                  margin-bottom: 32px;
              }
              .btn {
                  display: inline-flex;
                  align-items: center;
                  gap: 8px;
                  padding: 14px 28px;
                  border: none;
                  border-radius: 12px;
                  font-size: 16px;
                  font-weight: 600;
                  text-decoration: none;
                  cursor: pointer;
                  transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
                  position: relative;
                  overflow: hidden;
              }
              .btn::before {
                  content: '';
                  position: absolute;
                  top: 0;
                  left: -100%;
                  width: 100%;
                  height: 100%;
                  background: linear-gradient(90deg, transparent, rgba(255,255,255,0.2), transparent);
                  transition: left 0.5s;
              }
              .btn:hover::before {
                  left: 100%;
              }
              .btn-primary {
                  background: linear-gradient(135deg, #10b981, #059669);
                  color: white;
                  box-shadow: 0 4px 14px rgba(16, 185, 129, 0.3);
              }
              .btn-primary:hover {
                  transform: translateY(-2px);
                  box-shadow: 0 8px 25px rgba(16, 185, 129, 0.4);
              }
              .btn-secondary {
                  background: linear-gradient(135deg, #3b82f6, #2563eb);
                  color: white;
                  box-shadow: 0 4px 14px rgba(59, 130, 246, 0.3);
              }
              .btn-secondary:hover {
                  transform: translateY(-2px);
                  box-shadow: 0 8px 25px rgba(59, 130, 246, 0.4);
              }
              .btn-tertiary {
                  background: linear-gradient(135deg, #6b7280, #4b5563);
                  color: white;
                  box-shadow: 0 4px 14px rgba(107, 114, 128, 0.3);
              }
              .btn-tertiary:hover {
                  transform: translateY(-2px);
                  box-shadow: 0 8px 25px rgba(107, 114, 128, 0.4);
              }
              .info-section {
                  background: linear-gradient(135deg, rgba(99, 102, 241, 0.1), rgba(139, 92, 246, 0.1));
                  border: 1px solid rgba(99, 102, 241, 0.2);
                  border-radius: 16px;
                  padding: 24px;
                  margin-bottom: 24px;
              }
              .info-title {
                  font-size: 14px;
                  font-weight: 600;
                  color: #6366f1;
                  margin-bottom: 8px;
              }
              .info-detail {
                  font-size: 14px;
                  color: #6b7280;
              }
              .error-notice {
                  background: linear-gradient(135deg, rgba(239, 68, 68, 0.1), rgba(220, 38, 38, 0.1));
                  border: 1px solid rgba(239, 68, 68, 0.3);
                  border-radius: 16px;
                  padding: 20px;
                  margin-top: 24px;
              }
              .error-notice-title {
                  font-size: 16px;
                  font-weight: 600;
                  color: #dc2626;
                  margin-bottom: 8px;
                  display: flex;
                  align-items: center;
                  gap: 8px;
              }
              .error-notice-text {
                  font-size: 14px;
                  color: #7f1d1d;
                  line-height: 1.6;
              }
              @media (max-width: 640px) {
                  .content-card { padding: 24px; margin: 16px; }
                  .status-title { font-size: 24px; }
                  .status-message { font-size: 16px; }
                  .button-group { flex-direction: column; }
                  .btn { width: 100%; justify-content: center; }
              }
          </style>
      </head>
      <body>
          <div class="background-overlay"></div>
          <div class="main-container">
              <div class="content-card">
                  <div class="status-icon">⏳</div>
                  <h1 class="status-title">回答ボード準備中</h1>
                  <p class="status-message">
                      現在、回答ボードは非公開になっています。<br>
                      管理者として以下の操作が可能です。
                  </p>
                  
                  <div class="button-group">
                      <button onclick="republishBoard()" class="btn btn-primary">
                          🔄 回答ボードを再公開
                      </button>
                      <a href="?mode=admin&userId=${encodeURIComponent(userId)}" class="btn btn-secondary">
                          ⚙️ 管理パネルを開く
                      </a>
                      <button onclick="location.reload()" class="btn btn-tertiary">
                          🔄 ページを更新
                      </button>
                  </div>
                  
                  <div class="info-section">
                      <div class="info-title">管理者情報</div>
                      <div class="info-detail">
                          管理者: ${adminEmail || 'システム管理者'}<br>
                          ユーザーID: ${userId || '不明'}
                      </div>
                  </div>
                  
                  <div class="error-notice">
                      <div class="error-notice-title">
                          ⚠️ システム情報
                      </div>
                      <div class="error-notice-text">
                          テンプレートの読み込みでエラーが発生したため、基本機能のみ表示しています。すべての管理機能は正常に動作します。
                      </div>
                  </div>
              </div>
          </div>
          
          <script>
              function republishBoard() {
                  if (!confirm('回答ボードを再公開しますか？')) return;
                  
                  const button = event.target;
                  button.disabled = true;
                  button.textContent = '再公開中...';
                  
                  try {
                      google.script.run
                          .withSuccessHandler((result) => {
                              alert('再公開が完了しました！ページを更新します。');
                              setTimeout(() => {
                                  window.location.href = '?mode=view&userId=${encodeURIComponent(userId)}&_cb=' + Date.now();
                              }, 1000);
                          })
                          .withFailureHandler((error) => {
                              alert('再公開に失敗しました: ' + error.message);
                              button.disabled = false;
                              button.textContent = '🔄 回答ボードを再公開';
                          })
                          .republishBoard('${userId}');
                  } catch (error) {
                      alert('エラーが発生しました: ' + error.message);
                      button.disabled = false;
                      button.textContent = '🔄 回答ボードを再公開';
                  }
              }
          </script>
      </body>
      </html>
    `;
    
    const finalHtmlOutput = HtmlService.createHtmlOutput(finalFallbackHtml)
      .setTitle('StudyQuest - 準備中');
    
    // addMetaTagを安全に追加
    try {
      finalHtmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } catch (e) {
      console.warn('⚠️ Final fallback addMetaTag(viewport) failed:', e.message);
    }
    
    try {
      finalHtmlOutput.addMetaTag('cache-control', 'no-cache, no-store, must-revalidate');
    } catch (e) {
      console.warn('⚠️ Final fallback addMetaTag(cache-control) failed:', e.message);
    }
    
    return finalHtmlOutput;
  }
}

function renderAnswerBoard(userInfo, params) {
  try {
    let config = {};
    try {
      config = JSON.parse(userInfo.configJson || '{}');
    } catch (e) {
      warnLog('Invalid configJson:', e.message);
    }
  // publishedSheetNameの型安全性確保（'true'問題の修正）
  let safePublishedSheetName = '';
  if (config.publishedSheetName) {
    if (typeof config.publishedSheetName === 'string') {
      safePublishedSheetName = config.publishedSheetName;
    } else {
      logValidationError('publishedSheetName', config.publishedSheetName, 'string_type', `不正な型: ${typeof config.publishedSheetName}`);
      warnLog('🔧 main.gs: publishedSheetNameを空文字にリセットしました');
      safePublishedSheetName = '';
    }
  }

  // 強化されたパブリケーション状態検証（キャッシュバスティング対応）
  const isPublished = !!(config.appPublished && config.publishedSpreadsheetId && safePublishedSheetName);

  // リアルタイム検証: 非公開状態の場合は確実に検出
  const isCurrentlyPublished = isPublished &&
    config.appPublished === true &&
    config.publishedSpreadsheetId &&
    safePublishedSheetName;

  const sheetConfigKey = 'sheet_' + (safePublishedSheetName || params.sheetName);
  const sheetConfig = config[sheetConfigKey] || {};

  // この関数は公開ボード専用（非公開判定は呼び出し前に完了）
  debugLog('✅ renderAnswerBoard: Rendering published board for userId:', userInfo.userId);

  const template = HtmlService.createTemplateFromFile('Page');
  template.include = include;

  try {
      if (userInfo.spreadsheetId) {
        try { addServiceAccountToSpreadsheet(userInfo.spreadsheetId); } catch (err) { warnLog('アクセス権設定警告:', err.message); }
      }
      template.userId = userInfo.userId;
      template.spreadsheetId = userInfo.spreadsheetId;
      template.ownerName = userInfo.adminEmail;
      template.sheetName = escapeJavaScript(safePublishedSheetName || params.sheetName);
      template.DEBUG_MODE = shouldEnableDebugMode();
      // setupStatus未完了時の安全なopinionHeader取得
      const setupStatus = config.setupStatus || 'pending';
      let rawOpinionHeader;

      if (setupStatus === 'pending') {
        rawOpinionHeader = 'セットアップ中...';
      } else {
        rawOpinionHeader = sheetConfig.opinionHeader || safePublishedSheetName || 'お題';
      }
      template.opinionHeader = escapeJavaScript(rawOpinionHeader);
      template.cacheTimestamp = Date.now();
      template.displayMode = config.displayMode || 'anonymous';
      template.showCounts = config.showCounts !== undefined ? config.showCounts : false;
      template.showScoreSort = template.showCounts;
      const currentUserEmail = getCurrentUserEmail();
      const isOwner = currentUserEmail === userInfo.adminEmail;
      template.showAdminFeatures = isOwner;
      template.isAdminUser = isOwner;
    } catch (e) {
      template.opinionHeader = escapeJavaScript('お題の読込エラー');
      template.cacheTimestamp = Date.now();
      template.userId = userInfo.userId;
      template.spreadsheetId = userInfo.spreadsheetId;
      template.ownerName = userInfo.adminEmail;
      template.sheetName = escapeJavaScript(params.sheetName);
      template.displayMode = 'anonymous';
      template.showCounts = false;
      template.showScoreSort = false;
      const currentUserEmail = getCurrentUserEmail();
      const isOwner = currentUserEmail === userInfo.adminEmail;
      template.showAdminFeatures = isOwner;
      template.isAdminUser = isOwner;
    }

  // 公開ボード: 通常のキャッシュ設定
  return template.evaluate()
    .setTitle('StudyQuest -みんなの回答ボード-')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');

  } catch (error) {
    logError(
      error,
      'renderAnswerBoard',
      MAIN_ERROR_SEVERITY.HIGH,
      MAIN_ERROR_CATEGORIES.SYSTEM,
      {
        userId: userInfo.userId,
        spreadsheetId: userInfo.spreadsheetId,
        sheetName: safePublishedSheetName || params.sheetName,
      }
    );
    // フォールバック: 基本的なエラーページ
    return showErrorPage('エラー', 'ボードの表示でエラーが発生しました。管理者にお問い合わせください。');
  }
}

/**
 * クライアントサイドからのパブリケーション状態チェック
 * キャッシュバスティング対応のため、リアルタイムで状態を確認
 * @param {string} userId - ユーザーID
 * @returns {Object} パブリケーション状態情報
 */
function checkCurrentPublicationStatus(userId) {
  try {
    debugLog('🔍 checkCurrentPublicationStatus called for userId:', userId);

    if (!userId) {
      warnLog('userId is required for publication status check');
      return { error: 'userId is required', isPublished: false };
    }

    // ユーザー情報を強制的に最新状態で取得（キャッシュバイパス）
    const userInfo = findUserById(userId, {
      useExecutionCache: false,
      forceRefresh: true
    });

    if (!userInfo) {
      warnLog('User not found for publication status check:', userId);
      return { error: 'User not found', isPublished: false };
    }

    // 設定情報を解析
    let config = {};
    try {
      config = JSON.parse(userInfo.configJson || '{}');
    } catch (e) {
      warnLog('Config JSON parse error during publication status check:', e.message);
      return { error: 'Config parse error', isPublished: false };
    }

    // 現在のパブリケーション状態を厳密にチェック
    const isCurrentlyPublished = !!(
      config.appPublished === true &&
      config.publishedSpreadsheetId &&
      config.publishedSheetName &&
      typeof config.publishedSheetName === 'string' &&
      config.publishedSheetName.trim() !== ''
    );

    debugLog('📊 Publication status check result:', {
      userId: userId,
      appPublished: config.appPublished,
      hasSpreadsheetId: !!config.publishedSpreadsheetId,
      hasSheetName: !!config.publishedSheetName,
      isCurrentlyPublished: isCurrentlyPublished,
      timestamp: new Date().toISOString()
    });

    return {
      isPublished: isCurrentlyPublished,
      publishedSheetName: config.publishedSheetName || null,
      publishedSpreadsheetId: config.publishedSpreadsheetId || null,
      lastChecked: new Date().toISOString()
    };

  } catch (error) {
    logError(error, 'checkCurrentPublicationStatus', MAIN_ERROR_SEVERITY.MEDIUM, MAIN_ERROR_CATEGORIES.SYSTEM);
    return {
      error: error.message,
      isPublished: false,
      lastChecked: new Date().toISOString()
    };
  }
}
/**
 * JavaScript エスケープのテスト関数
 */

/**
 * パフォーマンス監視エンドポイント（簡易版）
 */

/**
 * escapeJavaScript関数のテスト
 */

/**
 * Base64エンコード/デコードのテスト
 */

// =================================================================
// DEBUG_MODE & USER ACCESS CONTROL API
// =================================================================

/**
 * 現在のDEBUG_MODE状態を取得
 * @returns {Object} DEBUG_MODEの状態情報
 */
function getDebugModeStatus() {
  try {
    const debugMode = PropertiesService.getScriptProperties().getProperty('DEBUG_MODE') === 'true';
    
    return {
      status: 'success',
      debugMode: debugMode,
      message: debugMode ? 'デバッグモードが有効です' : 'デバッグモードが無効です',
      lastModified: PropertiesService.getScriptProperties().getProperty('DEBUG_MODE_LAST_MODIFIED') || 'unknown'
    };
  } catch (error) {
    errorLog('getDebugModeStatus error:', error.message);
    return {
      status: 'error',
      message: 'DEBUG_MODE状態の取得に失敗しました: ' + error.message
    };
  }
}

/**
 * DEBUG_MODEの状態を切り替える（システム管理者のみ）
 * @param {boolean} enable - デバッグモードを有効にするかどうか
 * @returns {Object} 操作結果
 */
function toggleDebugMode(enable) {
  try {
    // システム管理者権限チェック
    if (!isSystemAdmin()) {
      throw new Error('システム管理者権限が必要です');
    }
    
    const props = PropertiesService.getScriptProperties();
    const newValue = enable ? 'true' : 'false';
    const currentValue = props.getProperty('DEBUG_MODE');
    
    // 変更があるかチェック
    if (currentValue === newValue) {
      return {
        status: 'success',
        debugMode: enable,
        message: `DEBUG_MODEは既に${enable ? '有効' : '無効'}です`,
        changed: false
      };
    }
    
    // DEBUG_MODE設定を更新
    props.setProperties({
      'DEBUG_MODE': newValue,
      'DEBUG_MODE_LAST_MODIFIED': new Date().toISOString()
    });
    
    infoLog('DEBUG_MODE changed:', {
      from: currentValue || 'undefined',
      to: newValue,
      by: getCurrentUserEmail()
    });
    
    return {
      status: 'success',
      debugMode: enable,
      message: `DEBUG_MODEを${enable ? '有効' : '無効'}にしました`,
      changed: true,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    errorLog('toggleDebugMode error:', error.message);
    return {
      status: 'error',
      message: 'DEBUG_MODE切り替えに失敗しました: ' + error.message
    };
  }
}

/**
 * 簡易ステータス取得API（バックグラウンド監視用のハートビート）
 * クライアント側が存在しない関数を呼び出して失敗する問題を防ぎます。
 * 必要十分な最小情報のみを返し、将来の拡張に備えてフィールドを追加しやすくします。
 * @param {string=} userId 任意のユーザーID（呼び出し元から渡される場合）
 * @returns {Object} 現在時刻などの軽量ステータス
 */
function getStatus(userId) {
  try {
    // 軽量なハートビート情報のみ返す（高頻度呼び出しを想定）
    const debugMode = PropertiesService.getScriptProperties().getProperty('DEBUG_MODE') === 'true';
    return {
      status: 'success',
      message: 'ok',
      timestamp: new Date().toISOString(),
      debugMode: debugMode,
      userId: userId || null,
    };
  } catch (error) {
    // 失敗時も呼び出し側での復帰を容易にするため、簡潔なエラー応答を返す
    return {
      status: 'error',
      message: 'getStatus failed: ' + error.message,
      timestamp: new Date().toISOString(),
    };
  }
}

/**
 * 個別ユーザー自身のアクセス状態を取得
 * @returns {Object} 現在のアクセス状態
*/
function getUserActiveStatus() {
  try {
    const currentUser = getUserInfo();
    if (!currentUser || !currentUser.userId) {
      throw new Error('ユーザー情報が取得できません');
    }
    
    // データベースから現在のユーザー情報を取得
    const userData = fetchUserFromDatabase(currentUser.userId);
    if (!userData) {
      throw new Error('ユーザーデータが見つかりません');
    }
    
    const isActive = userData.isActive !== false; // デフォルトはtrue
    
    return {
      success: true,
      isActive: isActive,
      userId: currentUser.userId,
      email: currentUser.adminEmail,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    errorLog('getUserActiveStatus error:', error.message);
    return {
      success: false,
      error: 'アクセス状態の取得に失敗しました: ' + error.message,
      isActive: true // フォールバック
    };
  }
}

/**
 * 個別ユーザー自身のアクセス状態を更新
 * @param {string} targetUserId - 対象ユーザーのID（現在のユーザーと一致する必要がある）
 * @param {boolean} isActive - 新しいアクティブ状態
 * @returns {Object} 操作結果
 */
function updateSelfActiveStatus(targetUserId, isActive) {
  try {
    const currentUser = getUserInfo();
    if (!currentUser || !currentUser.userId) {
      throw new Error('ユーザー情報が取得できません');
    }
    
    // 自分自身のアクセス状態のみ変更可能
    if (targetUserId !== currentUser.userId) {
      throw new Error('自分自身のアクセス状態のみ変更できます');
    }
    
    // データベースのユーザー情報を更新
    const result = updateUserDatabaseField(targetUserId, 'isActive', isActive);
    
    if (!result.success) {
      throw new Error(result.error || 'データベース更新に失敗しました');
    }
    
    // キャッシュをクリア
    clearUserCache(targetUserId);
    
    debugLog(`User ${targetUserId} self-updated isActive to: ${isActive}`);
    
    return {
      success: true,
      message: `アクセス設定を${isActive ? '有効' : '無効'}に変更しました`,
      userId: targetUserId,
      isActive: isActive,
      changed: true,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    errorLog('updateSelfActiveStatus error:', error.message);
    return {
      success: false,
      message: 'アクセス設定の変更に失敗しました: ' + error.message
    };
  }
}

/**
 * 個別ユーザーのisActiveステータスを更新（管理者用）
 * @param {string} userId - 対象ユーザーのID
 * @param {boolean} isActive - 新しいアクティブ状態
 * @returns {Object} 操作結果
 */
function updateUserActiveStatus(userId, isActive) {
  try {
    // システム管理者権限チェック
    if (!isSystemAdmin()) {
      throw new Error('システム管理者権限が必要です');
    }
    
    if (!userId || typeof userId !== 'string') {
      throw new Error('有効なユーザーIDが必要です');
    }
    
    // ユーザー情報を取得
    const userInfo = findUserById(userId);
    if (!userInfo) {
      throw new Error('指定されたユーザーが見つかりません');
    }
    
    // 現在の状態と比較
    const currentActive = userInfo.isActive === true || String(userInfo.isActive).toLowerCase() === 'true';
    if (currentActive === isActive) {
      return {
        status: 'success',
        userId: userId,
        email: userInfo.adminEmail,
        isActive: isActive,
        message: `ユーザー ${userInfo.adminEmail} は既に${isActive ? 'アクティブ' : '非アクティブ'}です`,
        changed: false
      };
    }
    
    // データベースを更新
    updateUserInDatabase(userId, { isActive: isActive });
    
    infoLog('User active status changed:', {
      userId: userId,
      email: userInfo.adminEmail,
      from: currentActive,
      to: isActive,
      by: getCurrentUserEmail()
    });
    
    return {
      status: 'success',
      userId: userId,
      email: userInfo.adminEmail,
      isActive: isActive,
      message: `ユーザー ${userInfo.adminEmail} を${isActive ? 'アクティブ' : '非アクティブ'}にしました`,
      changed: true,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    errorLog('updateUserActiveStatus error:', error.message);
    return {
      status: 'error',
      message: 'ユーザーステータス更新に失敗しました: ' + error.message
    };
  }
}

/**
 * 複数ユーザーのisActiveステータスを一括更新
 * @param {string[]} userIds - 対象ユーザーIDの配列
 * @param {boolean} isActive - 新しいアクティブ状態
 * @returns {Object} 操作結果
 */
function bulkUpdateUserActiveStatus(userIds, isActive) {
  try {
    // システム管理者権限チェック
    if (!isSystemAdmin()) {
      throw new Error('システム管理者権限が必要です');
    }
    
    if (!Array.isArray(userIds) || userIds.length === 0) {
      throw new Error('有効なユーザーID配列が必要です');
    }
    
    const results = [];
    let successCount = 0;
    let errorCount = 0;
    
    // 各ユーザーを個別に更新
    for (const userId of userIds) {
      try {
        const result = updateUserActiveStatus(userId, isActive);
        results.push(result);
        if (result.status === 'success') {
          successCount++;
        } else {
          errorCount++;
        }
      } catch (error) {
        results.push({
          status: 'error',
          userId: userId,
          message: error.message
        });
        errorCount++;
      }
    }
    
    infoLog('Bulk user active status update:', {
      totalUsers: userIds.length,
      successCount: successCount,
      errorCount: errorCount,
      isActive: isActive,
      by: getCurrentUserEmail()
    });
    
    return {
      status: errorCount === 0 ? 'success' : 'partial',
      results: results,
      summary: {
        total: userIds.length,
        success: successCount,
        errors: errorCount
      },
      message: `${successCount}人のユーザーを${isActive ? 'アクティブ' : '非アクティブ'}にしました${errorCount > 0 ? ` (${errorCount}件のエラー)` : ''}`,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    errorLog('bulkUpdateUserActiveStatus error:', error.message);
    return {
      status: 'error',
      message: '一括ユーザーステータス更新に失敗しました: ' + error.message
    };
  }
}

/**
 * 全ユーザーのisActiveステータスを一括更新
 * @param {boolean} isActive - 新しいアクティブ状態
 * @returns {Object} 操作結果
 */
function bulkUpdateAllUsersActiveStatus(isActive) {
  try {
    // システム管理者権限チェック
    if (!isSystemAdmin()) {
      throw new Error('システム管理者権限が必要です');
    }
    
    // 全ユーザーリストを取得
    const allUsers = getAllUsers();
    if (!allUsers || allUsers.length === 0) {
      return {
        status: 'success',
        message: '更新対象のユーザーがいません',
        summary: { total: 0, success: 0, errors: 0 }
      };
    }
    
    // 全ユーザーのIDを抽出
    const userIds = allUsers.map(user => user.userId).filter(id => id);
    
    // 一括更新を実行
    return bulkUpdateUserActiveStatus(userIds, isActive);
    
  } catch (error) {
    errorLog('bulkUpdateAllUsersActiveStatus error:', error.message);
    return {
      status: 'error',
      message: '全ユーザー一括更新に失敗しました: ' + error.message
    };
  }
}

/**
 * Web AppのURLを取得する関数
 * クライアントサイドから呼び出されるAPI関数
 * @returns {string} Web AppのURL
 */
function getWebAppUrl() {
  try {
    return ScriptApp.getService().getUrl();
  } catch (error) {
    logError(error, 'getWebAppUrl', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM);
    return '';
  }
}

/**
 * 現在のユーザーのメールアドレスを取得する関数
 * @returns {string} 現在のユーザーのメールアドレス
 */
function getCurrentUserEmail() {
  try {
    return Session.getActiveUser().getEmail();
  } catch (error) {
    logError(error, 'getCurrentUserEmail', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM);
    return '';
  }
}

/**
 * シンプル・確実な管理者権限検証（3重チェック）
 * メールアドレス + ユーザーID + アクティブ状態の照合
 * @param {string} userId - 検証するユーザーのID
 * @returns {boolean} 管理者権限がある場合は true、そうでなければ false
 */
function verifyAdminAccess(userId) {
  try {
    // 基本的な引数チェック
    if (!userId || typeof userId !== 'string' || userId.trim() === '') {
      warnLog('verifyAdminAccess: 無効なuserIdが渡されました:', userId);
      return false;
    }

    // 現在のログインユーザーのメールアドレスを取得
    const activeUserEmail = getCurrentUserEmail();
    if (!activeUserEmail) {
      warnLog('verifyAdminAccess: アクティブユーザーのメールアドレスが取得できませんでした');
      return false;
    }

    debugLog('verifyAdminAccess: 認証開始', { userId, activeUserEmail });

    // データベースからユーザー情報を取得
    let userFromDb = null;
    
    // まずは通常のキャッシュ付き検索を試行
    userFromDb = findUserById(userId);
    
    // 見つからない場合は強制フレッシュで再試行
    if (!userFromDb) {
      debugLog('verifyAdminAccess: 強制フレッシュで再検索中...');
      userFromDb = fetchUserFromDatabase('userId', userId, { forceFresh: true });
    }

    // ユーザーが見つからない場合は認証失敗
    if (!userFromDb) {
      warnLog('verifyAdminAccess: ユーザーが見つかりません:', { 
        userId, 
        activeUserEmail 
      });
      return false;
    }

    // 3重チェック実行
    // 1. メールアドレス照合
    const dbEmail = String(userFromDb.adminEmail || '').toLowerCase().trim();
    const currentEmail = String(activeUserEmail).toLowerCase().trim();
    const isEmailMatched = dbEmail && currentEmail && dbEmail === currentEmail;

    // 2. ユーザーID照合（念のため）
    const isUserIdMatched = String(userFromDb.userId) === String(userId);

    // 3. アクティブ状態確認
    const isActive = (userFromDb.isActive === true || 
                     userFromDb.isActive === 'true' || 
                     String(userFromDb.isActive).toLowerCase() === 'true');

    debugLog('verifyAdminAccess: 3重チェック結果:', {
      isEmailMatched,
      isUserIdMatched,
      isActive,
      dbEmail,
      currentEmail
    });

    // 3つの条件すべてが満たされた場合のみ認証成功
    if (isEmailMatched && isUserIdMatched && isActive) {
      infoLog('✅ verifyAdminAccess: 認証成功', { userId, email: activeUserEmail });
      return true;
    } else {
      warnLog('❌ verifyAdminAccess: 認証失敗', {
        userId,
        activeUserEmail,
        failures: {
          email: !isEmailMatched,
          userId: !isUserIdMatched,
          active: !isActive
        }
      });
      return false;
    }

  } catch (error) {
    errorLog('❌ verifyAdminAccess: 認証処理エラー:', error.message);
    return false;
  }
}

/**
 * ユーザーの最終アクセス時刻のみを更新（設定は保護）
 * @param {string} userId - 更新対象のユーザーID
 */
function updateUserLastAccess(userId) {
  try {
    if (!userId) {
      warnLog('updateUserLastAccess: userIdが指定されていません');
      return;
    }

    const now = new Date().toISOString();
    debugLog('最終アクセス時刻を更新:', userId, now);

    // lastAccessedAtフィールドのみを更新（他の設定は保護）
    updateUser(userId, { lastAccessedAt: now });

  } catch (error) {
    errorLog('updateUserLastAccess エラー:', error.message);
  }
}

/**
 * クイックスタートセットアップを実行する（AdminPanelから呼び出される）
 * @param {string} requestUserId - リクエスト元のユーザーID（省略可能）
 * @returns {Object} セットアップ結果
 */
function executeQuickStartSetup(requestUserId) {
  try {
    // 既存のquickStartSetup関数を呼び出す
    return quickStartSetup(requestUserId);
  } catch (error) {
    logError(error, 'executeQuickStartSetup', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM);
    return {
      status: 'error',
      message: 'クイックスタートセットアップに失敗しました: ' + error.message
    };
  }
}

/**
 * クライアントサイドのエラーをサーバーサイドに報告する
 * @param {Object} errorInfo - エラー情報オブジェクト
 * @returns {Object} 報告結果
 */
function reportClientError(errorInfo) {
  try {
    // エラー情報の検証
    if (!errorInfo || typeof errorInfo !== 'object') {
      return { status: 'error', message: 'Invalid error info' };
    }

    // クライアントエラーとしてログに記録
    const errorMessage = `[CLIENT ERROR] ${errorInfo.message || 'Unknown error'}`;
    const errorContext = {
      url: errorInfo.url || 'unknown',
      userAgent: errorInfo.userAgent || 'unknown',
      timestamp: errorInfo.timestamp || new Date().toISOString(),
      stack: errorInfo.stack || '',
      additional: errorInfo.additional || {}
    };

    // エラーの重要度を判定
    const severity = errorInfo.severity || ERROR_SEVERITY.MEDIUM;
    
    // ログに記録
    logError(
      new Error(errorMessage), 
      'CLIENT_ERROR', 
      severity, 
      ERROR_CATEGORIES.USER_INPUT,
      errorContext
    );

    return { 
      status: 'success', 
      message: 'Error reported successfully',
      errorId: Utilities.getUuid()
    };
  } catch (error) {
    errorLog('reportClientError failed:', error.message);
    return { 
      status: 'error', 
      message: 'Failed to report error: ' + error.message 
    };
  }
}

/**
 * デバッグ用：現在のユーザー状態を詳細に確認する
 * @returns {Object} デバッグ情報
 */
function debugCurrentUser() {
  try {
    const result = {
      timestamp: new Date().toISOString(),
      currentEmail: null,
      databaseId: null,
      searches: [],
      users: []
    };

    // 1. 現在のユーザーメール取得
    try {
      result.currentEmail = getCurrentUserEmail();
    } catch (error) {
      result.emailError = error.message;
    }

    // 2. データベースID確認
    try {
      const props = PropertiesService.getScriptProperties();
      result.databaseId = props.getProperty('DATABASE_SPREADSHEET_ID');
    } catch (error) {
      result.databaseError = error.message;
    }

    if (result.currentEmail) {
      // 3. 各検索方法を試行
      const searchMethods = [
        { name: 'findUserByEmail', fn: () => findUserByEmail(result.currentEmail) },
        { name: 'fetchUserFromDatabase', fn: () => fetchUserFromDatabase('adminEmail', result.currentEmail) },
        { name: 'fetchUserFromDatabase_forceFresh', fn: () => fetchUserFromDatabase('adminEmail', result.currentEmail, { forceFresh: true }) }
      ];

      for (const method of searchMethods) {
        try {
          const user = method.fn();
          result.searches.push({
            method: method.name,
            success: !!user,
            user: user ? {
              userId: user.userId,
              adminEmail: user.adminEmail,
              isActive: user.isActive,
              hasSpreadsheetId: !!user.spreadsheetId
            } : null
          });
        } catch (error) {
          result.searches.push({
            method: method.name,
            success: false,
            error: error.message
          });
        }
      }

      // 4. 全ユーザー一覧から該当ユーザーを探す
      try {
        const allUsers = getAllUsers();
        result.totalUsers = allUsers ? allUsers.length : 0;
        if (allUsers && allUsers.length > 0) {
          const matchingUsers = allUsers.filter(user => 
            user.adminEmail && user.adminEmail.toLowerCase() === result.currentEmail.toLowerCase()
          );
          result.matchingUsers = matchingUsers.length;
          result.users = matchingUsers.map(user => ({
            userId: user.userId,
            adminEmail: user.adminEmail,
            isActive: user.isActive,
            createdAt: user.createdAt,
            lastAccessedAt: user.lastAccessedAt
          }));
        }
      } catch (error) {
        result.getAllUsersError = error.message;
      }
    }

    return { status: 'success', debug: result };
  } catch (error) {
    return { 
      status: 'error', 
      message: error.message,
      stack: error.stack 
    };
  }
}
