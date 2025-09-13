/**
 * @fileoverview Logger - 統一ロギングサービス
 *
 * 🎯 目的:
 * - デバッグログの統一フォーマット
 * - 絵文字・レベル別ログ管理
 * - パフォーマンス計測統合
 */

/**
 * UnifiedLogger - 統一ロギングサービス
 */
const UnifiedLogger = Object.freeze({

  // ==========================================
  // 🎯 ログレベル定義
  // ==========================================

  LEVELS: Object.freeze({
    DEBUG: { emoji: '🔍', level: 0, name: 'DEBUG' },
    INFO: { emoji: 'ℹ️', level: 1, name: 'INFO' },
    WARN: { emoji: '⚠️', level: 2, name: 'WARN' },
    ERROR: { emoji: '❌', level: 3, name: 'ERROR' },
    SUCCESS: { emoji: '✅', level: 1, name: 'SUCCESS' }
  }),

  // ==========================================
  // 🔧 核心ログ機能
  // ==========================================

  /**
   * デバッグログ（🔍統一）
   * @param {string} tag - ログタグ
   * @param {Object} data - データオブジェクト
   * @param {string} message - オプションメッセージ
   */
  debug(tag, data, message = '') {
    const level = this.LEVELS.DEBUG;
    const formatted = this.format(level, tag, data, message);
    console.log(formatted.text, formatted.data);
  },

  /**
   * 情報ログ
   * @param {string} tag - ログタグ
   * @param {Object} data - データオブジェクト
   * @param {string} message - オプションメッセージ
   */
  info(tag, data, message = '') {
    const level = this.LEVELS.INFO;
    const formatted = this.format(level, tag, data, message);
    console.info(formatted.text, formatted.data);
  },

  /**
   * 警告ログ
   * @param {string} tag - ログタグ
   * @param {Object} data - データオブジェクト
   * @param {string} message - オプションメッセージ
   */
  warn(tag, data, message = '') {
    const level = this.LEVELS.WARN;
    const formatted = this.format(level, tag, data, message);
    console.warn(formatted.text, formatted.data);
  },

  /**
   * エラーログ
   * @param {string} tag - ログタグ
   * @param {Object} data - データオブジェクト
   * @param {string} message - オプションメッセージ
   */
  error(tag, data, message = '') {
    const level = this.LEVELS.ERROR;
    const formatted = this.format(level, tag, data, message);
    console.error(formatted.text, formatted.data);
  },

  /**
   * 成功ログ
   * @param {string} tag - ログタグ
   * @param {Object} data - データオブジェクト
   * @param {string} message - オプションメッセージ
   */
  success(tag, data, message = '') {
    const level = this.LEVELS.SUCCESS;
    const formatted = this.format(level, tag, data, message);
    console.log(formatted.text, formatted.data);
  },

  // ==========================================
  // 🎛 フォーマッター
  // ==========================================

  /**
   * ログフォーマット
   * @param {Object} level - ログレベル
   * @param {string} tag - タグ
   * @param {Object} data - データ
   * @param {string} message - メッセージ
   * @returns {Object} フォーマット済みログ
   */
  format(level, tag, data, message) {
    const timestamp = new Date().toISOString().substr(11, 8); // HH:MM:SS
    const text = `${level.emoji} ${timestamp} [${tag}]${message ? `: ${  message}` : ''}`;

    return {
      text,
      data: data || {},
      level: level.name,
      tag,
      timestamp
    };
  },

  // ==========================================
  // 📊 パフォーマンス計測
  // ==========================================

  /**
   * パフォーマンス計測開始
   * @param {string} operation - 操作名
   * @returns {Object} タイマーオブジェクト
   */
  startTimer(operation) {
    const startTime = Date.now();
    return {
      operation,
      startTime,
      end: () => {
        const executionTime = Date.now() - startTime;
        this.info('Performance', {
          operation,
          executionTime: `${executionTime}ms`,
          completed: true
        });
        return executionTime;
      }
    };
  },

  /**
   * 関数実行時間計測
   * @param {string} name - 関数名
   * @param {Function} fn - 実行する関数
   * @returns {*} 関数の戻り値
   */
  async measureFunction(name, fn) {
    const timer = this.startTimer(name);
    try {
      const result = await fn();
      timer.end();
      return result;
    } catch (error) {
      const executionTime = Date.now() - timer.startTime;
      this.error('Performance', {
        operation: name,
        executionTime: `${executionTime}ms`,
        error: error.message
      });
      throw error;
    }
  },

  // ==========================================
  // 🔧 ユーティリティ
  // ==========================================

  /**
   * サービス診断
   * @returns {Object} 診断結果
   */
  diagnose() {
    return {
      service: 'UnifiedLogger',
      timestamp: new Date().toISOString(),
      levels: Object.keys(this.LEVELS),
      features: [
        'Unified emoji logging',
        'Performance measurement',
        'Structured data logging',
        'Multi-level support'
      ],
      status: '✅ Active'
    };
  }

});

// ==========================================
// 📊 グローバル互換関数
// ==========================================

/**
 * レガシー互換性のためのグローバル関数
 */
function logDebug(tag, data, message) {
  return UnifiedLogger.debug(tag, data, message);
}

function logInfo(tag, data, message) {
  return UnifiedLogger.info(tag, data, message);
}

function logWarn(tag, data, message) {
  return UnifiedLogger.warn(tag, data, message);
}

function logError(tag, data, message) {
  return UnifiedLogger.error(tag, data, message);
}

function logSuccess(tag, data, message) {
  return UnifiedLogger.success(tag, data, message);
}