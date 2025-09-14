/**
 * @fileoverview Data Formatting & Transformation
 * 
 * 🎯 責任範囲:
 * - データの表示用フォーマット
 * - 型変換・データ変換
 * - レスポンス形式の統一
 * - 出力データの正規化
 * 
 * 🔄 移行元:
 * - Core.gs のフォーマット関数
 * - DataService の変換処理
 * - 各所に分散している表示処理
 */

/* global CONSTANTS, AppLogger */

/**
 * ResponseFormatter - レスポンス形式統一
 * API応答とエラーレスポンスの標準化
 */
const ResponseFormatter = Object.freeze({

  /**
   * 成功レスポンス作成
   * @param {*} data - データ
   * @param {Object} metadata - メタデータ
   * @returns {Object} 標準成功レスポンス
   */
  createSuccessResponse(data, metadata = {}) {
    const response = {
      success: true,
      timestamp: new Date().toISOString(),
      data: data || null
    };

    // データが配列の場合はカウント追加
    if (Array.isArray(data)) {
      response.count = data.length;
    }

    // メタデータをマージ
    return { ...response, ...metadata };
  },

  /**
   * エラーレスポンス作成
   * @param {string} message - エラーメッセージ
   * @param {Object} details - 詳細情報
   * @returns {Object} 標準エラーレスポンス
   */
  createErrorResponse(message, details = {}) {
    return {
      success: false,
      message: message || 'エラーが発生しました',
      timestamp: new Date().toISOString(),
      data: null,
      count: 0,
      ...details
    };
  },

  /**
   * バルクデータレスポンス作成
   * @param {Object} bulkData - バルクデータ
   * @param {Object} options - オプション
   * @returns {Object} バルクレスポンス
   */
  createBulkResponse(bulkData, options = {}) {
    const response = this.createSuccessResponse(bulkData, {
      type: 'bulk_data',
      executionTime: options.executionTime,
      cached: options.cached || false
    });

    // 各データの統計情報追加
    if (bulkData) {
      response.statistics = {
        userInfo: !!bulkData.userInfo,
        config: !!bulkData.config,
        sheetData: Array.isArray(bulkData.sheetData) ? bulkData.sheetData.length : 0,
        systemInfo: !!bulkData.systemInfo
      };
    }

    return response;
  },

  /**
   * タイムスタンプフォーマット（短縮形式）
   * @param {string|Date} timestamp - タイムスタンプ
   * @returns {string} フォーマット済み日時
   */
  formatTimestamp(timestamp) {
    try {
      const date = typeof timestamp === 'string' ? new Date(timestamp) : timestamp;
      if (isNaN(date.getTime())) {
        return new Date().toISOString();
      }

      // 短縮形式: YYYY-MM-DD HH:mm
      return date.toLocaleString('ja-JP', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit'
      });
    } catch (error) {
      console.error('ResponseFormatter.formatTimestamp error:', error);
      return new Date().toISOString();
    }
  }

});

/**
 * DataFormatter - データ表示用フォーマット
 * ユーザー向け表示データの変換
 */
const DataFormatter = Object.freeze({

  /**
   * 日時フォーマット（日本語）
   * @param {string|Date} timestamp - タイムスタンプ
   * @param {Object} options - フォーマットオプション
   * @returns {string} フォーマット済み日時
   */
  formatDateTime(timestamp, options = {}) {
    const {
      style = 'short', // short, full, date, time
      locale = 'ja-JP'
    } = options;

    if (!timestamp) return '不明';

    try {
      const date = timestamp instanceof Date ? timestamp : new Date(timestamp);

      switch (style) {
        case 'short':
          return date.toLocaleString(locale, {
            month: 'numeric',
            day: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
          });
        
        case 'full':
          return date.toLocaleString(locale, {
            year: 'numeric',
            month: 'long',
            day: 'numeric',
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit'
          });
        
        case 'date':
          return date.toLocaleDateString(locale, {
            year: 'numeric',
            month: 'long',
            day: 'numeric'
          });
        
        case 'time':
          return date.toLocaleTimeString(locale, {
            hour: '2-digit',
            minute: '2-digit'
          });
        
        default:
          return date.toLocaleString(locale);
      }
    } catch (error) {
      console.warn('DataFormatter.formatDateTime: フォーマットエラー', error.message);
      return '不明';
    }
  },

  /**
   * 相対時間フォーマット（〜前）
   * @param {string|Date} timestamp - タイムスタンプ
   * @returns {string} 相対時間表示
   */
  formatRelativeTime(timestamp) {
    if (!timestamp) return '不明';

    try {
      const date = timestamp instanceof Date ? timestamp : new Date(timestamp);
      const now = new Date();
      const diffMs = now.getTime() - date.getTime();
      const diffSec = Math.floor(diffMs / 1000);
      const diffMin = Math.floor(diffSec / 60);
      const diffHour = Math.floor(diffMin / 60);
      const diffDay = Math.floor(diffHour / 24);

      if (diffSec < 60) return '今';
      if (diffMin < 60) return `${diffMin}分前`;
      if (diffHour < 24) return `${diffHour}時間前`;
      if (diffDay < 7) return `${diffDay}日前`;
      
      // 1週間以上は通常フォーマット
      return this.formatDateTime(date, { style: 'date' });
    } catch (error) {
      console.warn('DataFormatter.formatRelativeTime: フォーマットエラー', error.message);
      return '不明';
    }
  },

  /**
   * 数値フォーマット（日本語）
   * @param {number} value - 数値
   * @param {Object} options - オプション
   * @returns {string} フォーマット済み数値
   */
  formatNumber(value, options = {}) {
    const {
      style = 'decimal', // decimal, percent, currency
      minimumFractionDigits = 0,
      maximumFractionDigits = 2
    } = options;

    if (typeof value !== 'number' || isNaN(value)) return '0';

    try {
      return value.toLocaleString('ja-JP', {
        style,
        minimumFractionDigits,
        maximumFractionDigits
      });
    } catch (error) {
      return value.toString();
    }
  },

  /**
   * パーセンテージフォーマット
   * @param {number} ratio - 比率（0-1）
   * @param {number} decimals - 小数点以下桁数
   * @returns {string} パーセント表示
   */
  formatPercentage(ratio, decimals = 1) {
    if (typeof ratio !== 'number' || isNaN(ratio)) return '0%';
    return `${(ratio * 100).toFixed(decimals)}%`;
  },

  /**
   * ファイルサイズフォーマット
   * @param {number} bytes - バイト数
   * @returns {string} ヒューマンリーダブルサイズ
   */
  formatFileSize(bytes) {
    if (typeof bytes !== 'number' || bytes < 0) return '0 B';

    const units = ['B', 'KB', 'MB', 'GB', 'TB'];
    let size = bytes;
    let unitIndex = 0;

    while (size >= 1024 && unitIndex < units.length - 1) {
      size /= 1024;
      unitIndex++;
    }

    const decimals = unitIndex === 0 ? 0 : 1;
    return `${size.toFixed(decimals)} ${units[unitIndex]}`;
  },

  /**
   * テキスト切り詰めフォーマット
   * @param {string} text - 元テキスト
   * @param {number} maxLength - 最大長
   * @param {string} suffix - 切り詰め時の接尾辞
   * @returns {string} 切り詰められたテキスト
   */
  truncateText(text, maxLength = 100, suffix = '...') {
    if (!text || typeof text !== 'string') return '';
    if (text.length <= maxLength) return text;

    // 単語境界で切り詰め（日本語対応）
    const truncated = text.substring(0, maxLength - suffix.length);
    
    // 最後が英単語の途中でない場合はそのまま
    if (!/[a-zA-Z]$/.test(truncated)) {
      return truncated + suffix;
    }

    // 英単語の途中の場合は単語境界まで戻る
    const lastSpace = truncated.lastIndexOf(' ');
    if (lastSpace > maxLength * 0.8) { // 80%以上の場合のみ単語境界を考慮
      return truncated.substring(0, lastSpace) + suffix;
    }

    return truncated + suffix;
  }

});

/**
 * HTMLFormatter - HTML出力用フォーマット
 * セキュアなHTML生成とエスケープ処理
 */
const HTMLFormatter = Object.freeze({

  /**
   * HTMLエスケープ
   * @param {string} text - エスケープ対象テキスト
   * @returns {string} エスケープ済みテキスト
   */
  escapeHtml(text) {
    if (!text || typeof text !== 'string') return '';

    return text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#x27;')
      .replace(/\//g, '&#x2F;');
  },

  /**
   * 改行をBRタグに変換
   * @param {string} text - 元テキスト
   * @returns {string} BR変換済みテキスト
   */
  nl2br(text) {
    if (!text || typeof text !== 'string') return '';
    
    return this.escapeHtml(text).replace(/\n/g, '<br>');
  },

  /**
   * マークダウン風の簡易フォーマット
   * @param {string} text - 元テキスト
   * @returns {string} フォーマット済みHTML
   */
  formatSimpleMarkdown(text) {
    if (!text || typeof text !== 'string') return '';

    let formatted = this.escapeHtml(text);

    // **太字**
    formatted = formatted.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
    
    // *斜体*
    formatted = formatted.replace(/\*(.*?)\*/g, '<em>$1</em>');
    
    // 改行
    formatted = formatted.replace(/\n/g, '<br>');

    // URLの自動リンク（Google関連のみ）
    const urlPattern = /(https?:\/\/(?:docs\.google\.com|forms\.gle|script\.google\.com|drive\.google\.com)[^\s]+)/g;
    formatted = formatted.replace(urlPattern, '<a href="$1" target="_blank" rel="noopener">$1</a>');

    return formatted;
  },

  /**
   * CSSクラス名生成
   * @param {Object} conditions - 条件オブジェクト
   * @returns {string} CSS クラス名
   */
  generateCssClasses(conditions) {
    return Object.entries(conditions)
      .filter(([className, condition]) => condition)
      .map(([className]) => className)
      .join(' ');
  }

});

/**
 * ConfigFormatter - 設定表示用フォーマット
 * 設定データの人間可読フォーマット
 */
const ConfigFormatter = Object.freeze({

  /**
   * セットアップステップ表示名
   * @param {number} step - ステップ番号
   * @returns {string} ステップ表示名
   */
  formatSetupStep(step) {
    const steps = {
      1: '📝 基本設定',
      2: '🔗 データソース連携',
      3: '🚀 アプリケーション公開'
    };
    return steps[step] || `ステップ ${step}`;
  },

  /**
   * アクセスレベル表示名
   * @param {string} level - アクセスレベル
   * @returns {string} 表示名
   */
  formatAccessLevel(level) {
    const levels = {
      'owner': '👑 所有者',
      'system_admin': '🔧 システム管理者',
      'authenticated_user': '✅ 認証済みユーザー',
      'guest': '👤 ゲスト',
      'none': '❌ アクセス不可'
    };
    return levels[level] || level;
  },

  /**
   * 設定完了度表示
   * @param {number} score - 完了度スコア（0-100）
   * @returns {string} 完了度表示
   */
  formatCompletionScore(score) {
    if (typeof score !== 'number') return '不明';

    const percentage = Math.round(score);
    const emoji = score >= 90 ? '🎉' : score >= 70 ? '✅' : score >= 50 ? '⚠️' : '❌';
    
    return `${emoji} ${percentage}%`;
  },

  /**
   * 列マッピング表示
   * @param {Object} columnMapping - 列マッピング
   * @returns {string} マッピング表示
   */
  formatColumnMapping(columnMapping) {
    if (!columnMapping || !columnMapping.mapping) return 'なし';

    const {mapping} = columnMapping;
    const mappedColumns = Object.entries(mapping)
      .filter(([key, value]) => typeof value === 'number')
      .map(([key, value]) => `${key}:${value}`)
      .join(', ');

    return mappedColumns || 'なし';
  }

});

/**
 * レガシー互換関数
 * 既存コードとの互換性維持
 */

// Core.gsからの移行
function formatTimestamp(date) {
  try {
    return ResponseFormatter.formatTimestamp(date);
  } catch (error) {
    console.error('formatTimestamp error:', error);
    return new Date(date).toISOString();
  }
}

function createSuccessResponse(data, metadata) {
  try {
    return ResponseFormatter.createSuccessResponse(data, metadata);
  } catch (error) {
    console.error('createSuccessResponse error:', error);
    return { success: true, data };
  }
}

function createErrorResponse(message, details) {
  try {
    return ResponseFormatter.createErrorResponse(message, details);
  } catch (error) {
    console.error('createErrorResponse error:', error);
    return { success: false, error: message };
  }
}

// DataServiceからの移行
function formatFullTimestamp(timestamp) {
  return DataFormatter.formatDateTime(timestamp, { style: 'full' });
}

/**
 * フォーマッター診断
 * @returns {Object} 診断結果
 */
function diagnoseFormatters() {
  const tests = [
    {
      name: 'DateTime Formatting',
      test: () => DataFormatter.formatDateTime(new Date()).length > 0
    },
    {
      name: 'HTML Escaping',
      test: () => HTMLFormatter.escapeHtml('<script>') === '&lt;script&gt;'
    },
    {
      name: 'Number Formatting',
      test: () => DataFormatter.formatNumber(1234.56).includes('1,234')
    },
    {
      name: 'Response Creation',
      test: () => ResponseFormatter.createSuccessResponse(['test']).count === 1
    }
  ];

  const results = tests.map(({ name, test }) => {
    try {
      return { name, status: test() ? '✅' : '❌', error: null };
    } catch (error) {
      return { name, status: '❌', error: error.message };
    }
  });

  return {
    service: 'DataFormatters',
    timestamp: new Date().toISOString(),
    tests: results,
    modules: {
      responseFormatter: typeof ResponseFormatter !== 'undefined',
      dataFormatter: typeof DataFormatter !== 'undefined',
      htmlFormatter: typeof HTMLFormatter !== 'undefined',
      configFormatter: typeof ConfigFormatter !== 'undefined'
    },
    overall: results.every(r => r.status === '✅') ? '✅' : '⚠️'
  };
}

// ===========================================
// 🎯 GAS Best Practice: Simple Functions
// ===========================================

/**
 * Simple timestamp formatting function for GAS flat architecture
 * @param {string|Date} timestamp - Timestamp to format
 * @returns {string} Formatted timestamp
 */
function formatTimestampSimple(timestamp) {
  try {
    const date = typeof timestamp === 'string' ? new Date(timestamp) : timestamp;
    if (isNaN(date.getTime())) {
      return new Date().toISOString();
    }

    // Short format: YYYY-MM-DD HH:mm
    return date.toLocaleString('ja-JP', {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit'
    });
  } catch (error) {
    console.error('formatTimestampSimple error:', error);
    return new Date().toISOString();
  }
}