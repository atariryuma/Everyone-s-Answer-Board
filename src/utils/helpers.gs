/**
 * @fileoverview Helper Utilities
 *
 * 🎯 責任範囲:
 * - 列マッピング・インデックス操作
 * - データフォーマッティング
 * - 汎用ヘルパー関数
 * - 計算・変換ユーティリティ
 *
 * 🔄 GAS Best Practices準拠:
 * - フラット関数構造 (Object.freeze削除)
 * - 直接的な関数エクスポート
 * - 簡素なユーティリティ関数群
 */

/* global getCurrentUserEmail, getUserConfig, DB, PROPS_KEYS, CONSTANTS, formatDataDateTime, URL */

// ===========================================
// 📋 列操作ヘルパー関数群
// ===========================================

/**
 * 統一列インデックス取得
 * @param {Object} config - ユーザー設定
 * @param {string} columnType - 列タイプ (answer/reason/class/name)
 * @returns {number} 列インデックス（見つからない場合は-1）
 */
function getHelperColumnIndex(config, columnType) {
    const index = config?.columnMapping?.mapping?.[columnType];
    return typeof index === 'number' ? index : -1;
}

/**
 * 列インデックスから実際のヘッダー名を取得
 * @param {Array} headers - ヘッダー配列
 * @param {number} columnIndex - 列インデックス
 * @returns {string} ヘッダー名（見つからない場合は空文字）
 */
function getHelperColumnHeaderByIndex(headers, columnIndex) {
    if (!Array.isArray(headers) || columnIndex < 0 || columnIndex >= headers.length) {
      return '';
    }
    return headers[columnIndex] || '';
}

/**
 * 設定から全列インデックスを一括取得
 * @param {Object} config - ユーザー設定
 * @returns {Object} 列タイプ別インデックスマップ
 */
function getAllHelperColumnIndices(config) {
    return {
      answer: getHelperColumnIndex(config, 'answer'),
      reason: getHelperColumnIndex(config, 'reason'),
      class: getHelperColumnIndex(config, 'class'),
      name: getHelperColumnIndex(config, 'name'),
    };
}

/**
 * 列マッピングの有効性検証
 * @param {Object} columnMapping - 列マッピング
 * @returns {Object} 検証結果
 */
function validateHelperColumnMapping(columnMapping) {
    const result = {
      isValid: true,
      errors: [],
      warnings: []
    };

    if (!columnMapping || !columnMapping.mapping) {
      result.isValid = false;
      result.errors.push('列マッピングが見つかりません');
      return result;
    }

    const {mapping} = columnMapping;
    const requiredColumns = ['answer'];
    const optionalColumns = ['reason', 'class', 'name'];

    // 必須列チェック
    requiredColumns.forEach(col => {
      if (typeof mapping[col] !== 'number' || mapping[col] < 0) {
        result.isValid = false;
        result.errors.push(`必須列 '${col}' が正しくマッピングされていません`);
      }
    });

    // オプション列チェック
    optionalColumns.forEach(col => {
      if (mapping[col] !== undefined && (typeof mapping[col] !== 'number' || mapping[col] < 0)) {
        result.warnings.push(`オプション列 '${col}' のマッピングが無効です`);
      }
    });

    return result;
}

// ===========================================
// 📏 フォーマッティングヘルパー関数群
// ===========================================

// formatTimestamp - formatters.jsに統一 (重複削除完了)

/**
 * 完全な日時フォーマット
 * @param {string|Date} timestamp - タイムスタンプ
 * @returns {string} 完全フォーマット済み日時
 */
function formatHelperFullTimestamp(timestamp) {
    if (!timestamp) return '不明';
    
    try {
      const date = timestamp instanceof Date ? timestamp : new Date(timestamp);
      return date.toLocaleString('ja-JP', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit'
      });
    } catch (error) {
      console.warn('FormatHelpers.formatFullTimestamp: フォーマットエラー', error.message);
      return '不明';
    }
}

/**
 * テキストの安全な切り詰め
 * @param {string} text - 元テキスト
 * @param {number} maxLength - 最大長
 * @param {string} suffix - 切り詰め時の接尾辞
 * @returns {string} 切り詰められたテキスト
 */
function formatHelperTruncateText(text, maxLength = 100, suffix = '...') {
    if (!text || typeof text !== 'string') return '';
    if (text.length <= maxLength) return text;
    
    return text.substring(0, maxLength - suffix.length) + suffix;
}

/**
 * 数値を%表示にフォーマット
 * @param {number} value - 数値（0-1）
 * @param {number} decimals - 小数点以下桁数
 * @returns {string} パーセント表示
 */
function formatHelperPercentage(value, decimals = 1) {
    if (typeof value !== 'number' || isNaN(value)) return '0%';
    return `${(value * 100).toFixed(decimals)}%`;
}

/**
 * ファイルサイズをヒューマンリーダブルに
 * @param {number} bytes - バイト数
 * @returns {string} フォーマット済みサイズ
 */
function formatHelperFileSize(bytes) {
    if (typeof bytes !== 'number' || bytes < 0) return '0 B';
    
    const units = ['B', 'KB', 'MB', 'GB'];
    let size = bytes;
    let unitIndex = 0;
    
    while (size >= 1024 && unitIndex < units.length - 1) {
      size /= 1024;
      unitIndex++;
    }
    
    return `${size.toFixed(1)} ${units[unitIndex]}`;
}

// ===========================================
// 🔍 検証ヘルパー関数群
// ===========================================

/**
 * 空値・null・undefined チェック
 * @param {*} value - チェック対象
 * @returns {boolean} 空かどうか
 */
function isHelperEmpty(value) {
    return value === null || value === undefined || value === '' || 
           (Array.isArray(value) && value.length === 0) ||
           (typeof value === 'object' && Object.keys(value).length === 0);
}

/**
 * 文字列の有効性チェック
 * @param {*} value - チェック対象
 * @param {number} minLength - 最小長
 * @param {number} maxLength - 最大長
 * @returns {Object} 検証結果
 */
function validateHelperString(value, minLength = 0, maxLength = Infinity) {
    const result = { isValid: true, error: null };
    
    if (typeof value !== 'string') {
      result.isValid = false;
      result.error = '文字列である必要があります';
      return result;
    }
    
    if (value.length < minLength) {
      result.isValid = false;
      result.error = `${minLength}文字以上である必要があります`;
      return result;
    }
    
    if (value.length > maxLength) {
      result.isValid = false;
      result.error = `${maxLength}文字以下である必要があります`;
      return result;
    }
    
    return result;
}

/**
 * 配列の有効性チェック
 * @param {*} value - チェック対象
 * @param {number} minItems - 最小要素数
 * @param {number} maxItems - 最大要素数
 * @returns {Object} 検証結果
 */
function validateHelperArray(value, minItems = 0, maxItems = Infinity) {
    const result = { isValid: true, error: null };
    
    if (!Array.isArray(value)) {
      result.isValid = false;
      result.error = '配列である必要があります';
      return result;
    }
    
    if (value.length < minItems) {
      result.isValid = false;
      result.error = `${minItems}個以上の要素が必要です`;
      return result;
    }
    
    if (value.length > maxItems) {
      result.isValid = false;
      result.error = `${maxItems}個以下の要素である必要があります`;
      return result;
    }
    
    return result;
}

// ===========================================
// 🧮 計算ヘルパー関数群
// ===========================================

/**
 * 配列の合計値計算
 * @param {Array} numbers - 数値配列
 * @returns {number} 合計値
 */
function calculateHelperSum(numbers) {
    if (!Array.isArray(numbers)) return 0;
    return numbers.reduce((sum, num) => sum + (typeof num === 'number' ? num : 0), 0);
}

/**
 * 配列の平均値計算
 * @param {Array} numbers - 数値配列
 * @returns {number} 平均値
 */
function calculateHelperAverage(numbers) {
    if (!Array.isArray(numbers) || numbers.length === 0) return 0;
    return calculateHelperSum(numbers) / numbers.length;
}

/**
 * パーセンテージ計算
 * @param {number} part - 部分値
 * @param {number} total - 総数
 * @returns {number} パーセンテージ（0-1）
 */
function calculateHelperPercentage(part, total) {
    if (typeof part !== 'number' || typeof total !== 'number' || total === 0) return 0;
    return Math.max(0, Math.min(1, part / total));
}

/**
 * 範囲内に値を制限
 * @param {number} value - 対象値
 * @param {number} min - 最小値
 * @param {number} max - 最大値
 * @returns {number} 制限された値
 */
function calculateHelperClamp(value, min, max) {
    return Math.max(min, Math.min(max, value));
}

/**
 * 二つの値の差の割合
 * @param {number} oldValue - 古い値
 * @param {number} newValue - 新しい値
 * @returns {number} 変化率（-1 to Infinity）
 */
function calculateHelperChangeRate(oldValue, newValue) {
    if (typeof oldValue !== 'number' || typeof newValue !== 'number' || oldValue === 0) return 0;
    return (newValue - oldValue) / oldValue;
}

// ===========================================
// 🔄 レガシー互換関数 (GAS Best Practices準拠)
// ===========================================

/**
 * 列インデックス取得 (互換性維持)
 * @param {Object} config - 設定
 * @param {string} columnType - 列タイプ
 * @returns {number} インデックス
 */
function getColumnIndex(config, columnType) {
  return getHelperColumnIndex(config, columnType);
}

/**
 * ヘッダー名取得 (互換性維持)
 * @param {Array} headers - ヘッダー配列
 * @param {number} columnIndex - インデックス
 * @returns {string} ヘッダー名
 */
function getColumnHeaderByIndex(headers, columnIndex) {
  return getHelperColumnHeaderByIndex(headers, columnIndex);
}

/**
 * 全列インデックス取得 (互換性維持)
 * @param {Object} config - 設定
 * @returns {Object} インデックスマップ
 */
function getAllColumnIndices(config) {
  return getAllHelperColumnIndices(config);
}

// formatTimestamp関数を削除 - formatters.gsで統一

/**
 * ヘルパー関数診断
 * @returns {Object} 診断結果
 */
function diagnoseHelpers() {
  return {
    service: 'UtilityHelpers',
    timestamp: new Date().toISOString(),
    functions: {
      columnHelpers: typeof getHelperColumnIndex === 'function',
      formatHelpers: typeof formatHelperFullTimestamp === 'function',
      validationHelpers: typeof isHelperEmpty === 'function',
      calculationHelpers: typeof calculateHelperSum === 'function'
    },
    legacyCompatibility: {
      getColumnIndex: typeof getColumnIndex === 'function',
      formatTimestamp: false // 削除済み - formatters.gsを使用
    }
  };
}