/**
 * @fileoverview CacheManager - キャッシュ管理ユーティリティ
 *
 * 🎯 責任範囲:
 * - レガシー互換性の維持
 * - キャッシュ管理のヘルパー関数
 *
 * ⚠️ 重要: 実際のキャッシュサービスは以下に統合されました：
 * - AppCacheService → infrastructure/CacheService.gs
 * - ErrorHandler → core/errors.gs
 *
 * このファイルはレガシー互換性のためのヘルパーのみを提供します。
 */

/* global AppCacheService, ErrorHandler */

/**
 * キャッシュマネージャーのヘルパー関数
 * レガシーコードとの互換性を維持するためのユーティリティ
 */
// eslint-disable-next-line no-unused-vars
const CacheManagerHelpers = Object.freeze({

  /**
   * 統合キャッシュサービスへの直接アクセス
   * @returns {Object} AppCacheService
   */
  getCacheService() {
    return AppCacheService;
  },

  /**
   * 統合エラーハンドラーへの直接アクセス
   * @returns {Object} ErrorHandler
   */
  getErrorHandler() {
    return ErrorHandler;
  },

  /**
   * キャッシュ統計情報取得
   * @returns {Object} キャッシュ統計
   */
  getStats() {
    try {
      return {
        timestamp: new Date().toISOString(),
        cacheService: 'AppCacheService (infrastructure/CacheService.gs)',
        errorHandler: 'ErrorHandler (core/errors.gs)',
        status: 'integrated',
        message: 'Cache management is now centralized'
      };
    } catch (error) {
      return {
        error: error.message,
        status: 'error'
      };
    }
  },

  /**
   * 簡便キャッシュ操作（レガシー互換）
   * @param {string} key - キャッシュキー
   * @returns {any} キャッシュされたデータまたはnull
   */
  quickGet(key) {
    return AppCacheService.get(key);
  },

  /**
   * 簡便キャッシュ保存（レガシー互換）
   * @param {string} key - キャッシュキー
   * @param {any} value - 保存するデータ
   * @param {number} ttl - TTL秒数
   */
  quickSet(key, value, ttl) {
    return AppCacheService.set(key, value, ttl);
  },

  /**
   * 簡便キャッシュ削除（レガシー互換）
   * @param {string} key - キャッシュキー
   */
  quickRemove(key) {
    return AppCacheService.remove(key);
  }
});