/**
 * @fileoverview Helper Utilities
 *
 * 🎯 責任範囲:
 * - 列マッピング・インデックス操作
 * - データフォーマッティング
 * - 汎用ヘルパー関数
 * - 計算・変換ユーティリティ
 *
 * 🔄 移行元:
 * - Base.gs の列操作関数
 * - Core.gs の汎用ヘルパー
 * - 各所に分散している共通処理
 */

/* global UserService, ConfigService, DB, PROPS_KEYS, CONSTANTS, DataFormatter, URL */

/**
 * ColumnHelpers - 列操作ヘルパー
 * スプレッドシートの列マッピングに関する統一操作
 */
const ColumnHelpers = Object.freeze({

  /**
   * 統一列インデックス取得
   * @param {Object} config - ユーザー設定
   * @param {string} columnType - 列タイプ (answer/reason/class/name)
   * @returns {number} 列インデックス（見つからない場合は-1）
   */
  getColumnIndex(config, columnType) {
    const index = config?.columnMapping?.mapping?.[columnType];
    return typeof index === 'number' ? index : -1;
  },

  /**
   * 列インデックスから実際のヘッダー名を取得
   * @param {Array} headers - ヘッダー配列
   * @param {number} columnIndex - 列インデックス
   * @returns {string} ヘッダー名（見つからない場合は空文字）
   */
  getColumnHeaderByIndex(headers, columnIndex) {
    if (!Array.isArray(headers) || columnIndex < 0 || columnIndex >= headers.length) {
      return '';
    }
    return headers[columnIndex] || '';
  },

  /**
   * 設定から全列インデックスを一括取得
   * @param {Object} config - ユーザー設定
   * @returns {Object} 列タイプ別インデックスマップ
   */
  getAllColumnIndices(config) {
    return {
      answer: this.getColumnIndex(config, 'answer'),
      reason: this.getColumnIndex(config, 'reason'),
      class: this.getColumnIndex(config, 'class'),
      name: this.getColumnIndex(config, 'name'),
    };
  },

  /**
   * 列マッピングの有効性検証
   * @param {Object} columnMapping - 列マッピング
   * @returns {Object} 検証結果
   */
  validateColumnMapping(columnMapping) {
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

});

/**
 * FormatHelpers - フォーマッティングヘルパー
 * データ変換・表示用フォーマット処理
 */
const FormatHelpers = Object.freeze({

  /**
   * タイムスタンプを読みやすい形式にフォーマット
   * @param {string|Date} timestamp - タイムスタンプ
   * @returns {string} フォーマット済み日時
   */
  formatTimestamp(timestamp) {
    // 統一: DataFormatterを使用（要インポート確認）
    if (typeof DataFormatter !== 'undefined') {
      return DataFormatter.formatDateTime(timestamp, { style: 'short' });
    }
    // フォールバック処理
    if (!timestamp) return '不明';
    try {
      const date = timestamp instanceof Date ? timestamp : new Date(timestamp);
      return date.toLocaleString('ja-JP', {
        month: 'numeric',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
      });
    } catch (error) {
      return '不明';
    }
  },

  /**
   * 完全な日時フォーマット
   * @param {string|Date} timestamp - タイムスタンプ
   * @returns {string} 完全フォーマット済み日時
   */
  formatFullTimestamp(timestamp) {
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
  },

  /**
   * テキストの安全な切り詰め
   * @param {string} text - 元テキスト
   * @param {number} maxLength - 最大長
   * @param {string} suffix - 切り詰め時の接尾辞
   * @returns {string} 切り詰められたテキスト
   */
  truncateText(text, maxLength = 100, suffix = '...') {
    if (!text || typeof text !== 'string') return '';
    if (text.length <= maxLength) return text;
    
    return text.substring(0, maxLength - suffix.length) + suffix;
  },

  /**
   * 数値を%表示にフォーマット
   * @param {number} value - 数値（0-1）
   * @param {number} decimals - 小数点以下桁数
   * @returns {string} パーセント表示
   */
  formatPercentage(value, decimals = 1) {
    if (typeof value !== 'number' || isNaN(value)) return '0%';
    return `${(value * 100).toFixed(decimals)}%`;
  },

  /**
   * ファイルサイズをヒューマンリーダブルに
   * @param {number} bytes - バイト数
   * @returns {string} フォーマット済みサイズ
   */
  formatFileSize(bytes) {
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

});

/**
 * ValidationHelpers - 検証ヘルパー
 * データ検証・バリデーション処理
 */
const ValidationHelpers = Object.freeze({

  /**
   * 空値・null・undefined チェック
   * @param {*} value - チェック対象
   * @returns {boolean} 空かどうか
   */
  isEmpty(value) {
    return value === null || value === undefined || value === '' || 
           (Array.isArray(value) && value.length === 0) ||
           (typeof value === 'object' && Object.keys(value).length === 0);
  },

  /**
   * 文字列の有効性チェック
   * @param {*} value - チェック対象
   * @param {number} minLength - 最小長
   * @param {number} maxLength - 最大長
   * @returns {Object} 検証結果
   */
  validateString(value, minLength = 0, maxLength = Infinity) {
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
  },

  /**
   * 配列の有効性チェック
   * @param {*} value - チェック対象
   * @param {number} minItems - 最小要素数
   * @param {number} maxItems - 最大要素数
   * @returns {Object} 検証結果
   */
  validateArray(value, minItems = 0, maxItems = Infinity) {
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

});

/**
 * CalculationHelpers - 計算ヘルパー
 * 数値計算・統計処理
 */
const CalculationHelpers = Object.freeze({

  /**
   * 配列の合計値計算
   * @param {Array} numbers - 数値配列
   * @returns {number} 合計値
   */
  sum(numbers) {
    if (!Array.isArray(numbers)) return 0;
    return numbers.reduce((sum, num) => sum + (typeof num === 'number' ? num : 0), 0);
  },

  /**
   * 配列の平均値計算
   * @param {Array} numbers - 数値配列
   * @returns {number} 平均値
   */
  average(numbers) {
    if (!Array.isArray(numbers) || numbers.length === 0) return 0;
    return this.sum(numbers) / numbers.length;
  },

  /**
   * パーセンテージ計算
   * @param {number} part - 部分値
   * @param {number} total - 総数
   * @returns {number} パーセンテージ（0-1）
   */
  percentage(part, total) {
    if (typeof part !== 'number' || typeof total !== 'number' || total === 0) return 0;
    return Math.max(0, Math.min(1, part / total));
  },

  /**
   * 範囲内に値を制限
   * @param {number} value - 対象値
   * @param {number} min - 最小値
   * @param {number} max - 最大値
   * @returns {number} 制限された値
   */
  clamp(value, min, max) {
    return Math.max(min, Math.min(max, value));
  },

  /**
   * 二つの値の差の割合
   * @param {number} oldValue - 古い値
   * @param {number} newValue - 新しい値
   * @returns {number} 変化率（-1 to Infinity）
   */
  changeRate(oldValue, newValue) {
    if (typeof oldValue !== 'number' || typeof newValue !== 'number' || oldValue === 0) return 0;
    return (newValue - oldValue) / oldValue;
  }

});

/**
 * 汎用ヘルパー関数（グローバル）
 * レガシーコードとの互換性維持
 */

// Base.gsからの移行関数
function getColumnIndex(config, columnType) {
  return ColumnHelpers.getColumnIndex(config, columnType);
}

function getColumnHeaderByIndex(headers, columnIndex) {
  return ColumnHelpers.getColumnHeaderByIndex(headers, columnIndex);
}

function getAllColumnIndices(config) {
  return ColumnHelpers.getAllColumnIndices(config);
}

// Core.gsからの移行関数
function formatTimestamp(timestamp) {
  return FormatHelpers.formatTimestamp(timestamp);
}

// 診断関数
function diagnoseHelpers() {
  return {
    service: 'UtilityHelpers',
    timestamp: new Date().toISOString(),
    modules: {
      columnHelpers: typeof ColumnHelpers !== 'undefined',
      formatHelpers: typeof FormatHelpers !== 'undefined', 
      validationHelpers: typeof ValidationHelpers !== 'undefined',
      calculationHelpers: typeof CalculationHelpers !== 'undefined'
    },
    legacyCompatibility: {
      getColumnIndex: typeof getColumnIndex === 'function',
      formatTimestamp: typeof formatTimestamp === 'function'
    }
  };
}