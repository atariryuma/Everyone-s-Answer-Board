/**
 * @fileoverview CacheManager - キャッシュ統合管理
 *
 * 🎯 責任範囲:
 * - 統一キャッシュ管理
 * - TTL管理
 * - キャッシュ統計
 */

/* global CacheService */

/**
 * 統一キャッシュサービス
 * 全サービスが使用する共通キャッシュ管理
 */
// eslint-disable-next-line no-unused-vars
const AppCacheService = Object.freeze({

  /**
   * TTL定数
   */
  TTL: Object.freeze({
    SHORT: 60,    // 1分
    MEDIUM: 300,  // 5分
    LONG: 1800    // 30分
  }),

  /**
   * キャッシュ取得
   * @param {string} key - キャッシュキー
   * @returns {any} キャッシュされたデータまたはnull
   */
  get(key) {
    try {
      const cache = CacheService.getScriptCache();
      const cached = cache.get(key);
      if (cached) {
        return JSON.parse(cached);
      }
      return null;
    } catch (error) {
      console.error('AppCacheService.get: エラー', { key, error: error.message });
      return null;
    }
  },

  /**
   * キャッシュ保存
   * @param {string} key - キャッシュキー
   * @param {any} value - 保存するデータ
   * @param {number} ttl - TTL秒数（デフォルト: MEDIUM）
   */
  set(key, value, ttl = this.TTL.MEDIUM) {
    try {
      const cache = CacheService.getScriptCache();
      cache.put(key, JSON.stringify(value), ttl);
    } catch (error) {
      console.error('AppCacheService.set: エラー', { key, error: error.message });
    }
  },

  /**
   * キャッシュ削除
   * @param {string} key - キャッシュキー
   */
  remove(key) {
    try {
      const cache = CacheService.getScriptCache();
      cache.remove(key);
    } catch (error) {
      console.error('AppCacheService.remove: エラー', { key, error: error.message });
    }
  }
});

/**
 * 統一エラーハンドラー
 * 全サービスが使用するエラー処理
 */
// eslint-disable-next-line no-unused-vars
const ErrorHandler = Object.freeze({

  /**
   * エラー処理
   * @param {Error} error - エラーオブジェクト
   * @param {string} context - エラー発生コンテキスト
   * @returns {Object} 統一エラーレスポンス
   */
  handle(error, context = 'unknown') {
    const errorInfo = {
      success: false,
      message: error.message || 'システムエラーが発生しました',
      errorCode: this.generateErrorCode(error),
      context,
      timestamp: new Date().toISOString()
    };

    // ログ出力
    console.error(`ErrorHandler[${context}]:`, errorInfo);

    return errorInfo;
  },

  /**
   * エラーコード生成
   * @param {Error} error - エラーオブジェクト
   * @returns {string} エラーコード
   */
  generateErrorCode(error) {
    if (!error.message) return 'ERR_UNKNOWN';

    // メッセージベースの簡易コード生成
    const hash = error.message.split('').reduce((a, b) => {
      a = ((a << 5) - a) + b.charCodeAt(0);
      return a & a;
    }, 0);

    return `ERR_${Math.abs(hash).toString(16).substr(0, 6).toUpperCase()}`;
  }
});