/**
 * @fileoverview 共通ユーティリティ関数群
 * 重複コードの排除と一貫性向上のための共通関数を提供
 */

/**
 * 現在のアクティブユーザーのメールアドレスを安全に取得
 * @returns {string} ユーザーメールアドレス（取得失敗時は空文字）
 */
function getCurrentUserEmail() {
  try {
    return Session.getActiveUser().getEmail() || '';
  } catch (error) {
    logError(error, 'getCurrentUserEmail', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.AUTHENTICATION);
    return '';
  }
}

/**
 * スクリプトプロパティを安全に取得
 * @param {string} key - プロパティキー
 * @param {string} defaultValue - デフォルト値
 * @returns {string} プロパティ値
 */
function getScriptProperty(key, defaultValue = '') {
  try {
    return PropertiesService.getScriptProperties().getProperty(key) || defaultValue;
  } catch (error) {
    logError(error, 'getScriptProperty', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM, { key });
    return defaultValue;
  }
}

/**
 * スクリプトプロパティを安全に設定
 * @param {string} key - プロパティキー
 * @param {string} value - 設定値
 * @returns {boolean} 設定成功フラグ
 */
function setScriptProperty(key, value) {
  try {
    PropertiesService.getScriptProperties().setProperty(key, value);
    return true;
  } catch (error) {
    logError(error, 'setScriptProperty', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM, { key });
    return false;
  }
}

/**
 * ユーザープロパティを安全に取得
 * @param {string} key - プロパティキー
 * @param {string} defaultValue - デフォルト値
 * @returns {string} プロパティ値
 */
function getUserProperty(key, defaultValue = '') {
  try {
    return PropertiesService.getUserProperties().getProperty(key) || defaultValue;
  } catch (error) {
    logError(error, 'getUserProperty', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM, { key });
    return defaultValue;
  }
}

/**
 * ユーザープロパティを安全に設定
 * @param {string} key - プロパティキー
 * @param {string} value - 設定値
 * @returns {boolean} 設定成功フラグ
 */
function setUserProperty(key, value) {
  try {
    PropertiesService.getUserProperties().setProperty(key, value);
    return true;
  } catch (error) {
    logError(error, 'setUserProperty', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM, { key });
    return false;
  }
}

/**
 * ユーザープロパティを安全に削除
 * @param {string} key - プロパティキー
 * @returns {boolean} 削除成功フラグ
 */
function deleteUserProperty(key) {
  try {
    PropertiesService.getUserProperties().deleteProperty(key);
    return true;
  } catch (error) {
    logError(error, 'deleteUserProperty', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM, { key });
    return false;
  }
}

/**
 * JSON文字列を安全にパース
 * @param {string} jsonString - JSON文字列
 * @param {object} defaultValue - パース失敗時のデフォルト値
 * @returns {object} パース結果
 */
function safeJsonParse(jsonString, defaultValue = {}) {
  try {
    return JSON.parse(jsonString || '{}');
  } catch (error) {
    logError(error, 'safeJsonParse', ERROR_SEVERITY.LOW, ERROR_CATEGORIES.VALIDATION, 
             { jsonStringLength: jsonString ? jsonString.length : 0 });
    return defaultValue;
  }
}

/**
 * オブジェクトを安全にJSON文字列化
 * @param {object} obj - オブジェクト
 * @param {string} defaultValue - 変換失敗時のデフォルト値
 * @returns {string} JSON文字列
 */
function safeJsonStringify(obj, defaultValue = '{}') {
  try {
    return JSON.stringify(obj);
  } catch (error) {
    logError(error, 'safeJsonStringify', ERROR_SEVERITY.LOW, ERROR_CATEGORIES.VALIDATION);
    return defaultValue;
  }
}

/**
 * 文字列の有効性チェック（null、undefined、空文字チェック）
 * @param {string} str - チェック対象文字列
 * @returns {boolean} 有効な文字列かどうか
 */
function isValidString(str) {
  return str !== null && str !== undefined && typeof str === 'string' && str.trim() !== '';
}

/**
 * 配列の有効性チェック（null、undefined、空配列チェック）
 * @param {Array} arr - チェック対象配列
 * @returns {boolean} 有効な配列かどうか
 */
function isValidArray(arr) {
  return Array.isArray(arr) && arr.length > 0;
}

/**
 * オブジェクトの有効性チェック（null、undefined、空オブジェクトチェック）
 * @param {object} obj - チェック対象オブジェクト
 * @returns {boolean} 有効なオブジェクトかどうか
 */
function isValidObject(obj) {
  return obj !== null && obj !== undefined && typeof obj === 'object' && !Array.isArray(obj) && Object.keys(obj).length > 0;
}

/**
 * 安全な文字列変換
 * @param {any} value - 変換対象値
 * @param {string} defaultValue - 変換失敗時のデフォルト値
 * @returns {string} 文字列
 */
function safeToString(value, defaultValue = '') {
  try {
    if (value === null || value === undefined) {
      return defaultValue;
    }
    return String(value);
  } catch (error) {
    logError(error, 'safeToString', ERROR_SEVERITY.LOW, ERROR_CATEGORIES.VALIDATION);
    return defaultValue;
  }
}

/**
 * 安全な数値変換
 * @param {any} value - 変換対象値
 * @param {number} defaultValue - 変換失敗時のデフォルト値
 * @returns {number} 数値
 */
function safeToNumber(value, defaultValue = 0) {
  try {
    const num = Number(value);
    return isNaN(num) ? defaultValue : num;
  } catch (error) {
    logError(error, 'safeToNumber', ERROR_SEVERITY.LOW, ERROR_CATEGORIES.VALIDATION);
    return defaultValue;
  }
}

/**
 * 配列から重複要素を除去
 * @param {Array} arr - 対象配列
 * @returns {Array} 重複除去後の配列
 */
function removeDuplicates(arr) {
  if (!Array.isArray(arr)) {
    return [];
  }
  return [...new Set(arr)];
}

/**
 * オブジェクトのディープクローン
 * @param {object} obj - クローン対象オブジェクト
 * @returns {object} クローンされたオブジェクト
 */
function deepClone(obj) {
  try {
    return JSON.parse(JSON.stringify(obj));
  } catch (error) {
    logError(error, 'deepClone', ERROR_SEVERITY.LOW, ERROR_CATEGORIES.SYSTEM);
    return obj;
  }
}

/**
 * 実行時間計測デコレータ
 * @param {Function} func - 計測対象関数
 * @param {string} functionName - 関数名
 * @param {number} warningThreshold - 警告閾値（ミリ秒）
 * @returns {Function} ラップされた関数
 */
function measureExecutionTime(func, functionName, warningThreshold = 1000) {
  return function(...args) {
    const startTime = Date.now();
    try {
      const result = func.apply(this, args);
      const duration = Date.now() - startTime;
      
      if (duration > warningThreshold) {
        logPerformanceWarning(functionName, duration, warningThreshold, { args: args.length });
      }
      
      return result;
    } catch (error) {
      const duration = Date.now() - startTime;
      logError(error, functionName, ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM, 
               { executionTime: duration, args: args.length });
      throw error;
    }
  };
}

/**
 * リトライ実行ユーティリティ
 * @param {Function} operation - 実行する操作
 * @param {number} maxRetries - 最大リトライ数
 * @param {number} delay - リトライ間隔（ミリ秒）
 * @param {string} operationName - 操作名
 * @returns {any} 操作の結果
 */
function retryOperation(operation, maxRetries = 3, delay = 1000, operationName = 'unknown') {
  let lastError = null;
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      return operation();
    } catch (error) {
      lastError = error;
      logError(error, `${operationName}_retry_${attempt}`, ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM, 
               { attempt, maxRetries });
      
      if (attempt < maxRetries) {
        Utilities.sleep(delay);
      }
    }
  }
  
  throw new Error(`${operationName} failed after ${maxRetries} attempts: ${lastError.message}`);
}