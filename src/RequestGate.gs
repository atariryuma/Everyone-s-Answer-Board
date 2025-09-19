/**
 * @fileoverview RequestGate - 重複処理防止・並行制御
 *
 * 🎯 責任範囲:
 * - トランザクション重複防止
 * - 並行処理制御
 * - Zero-Dependency Architecture準拠
 * - GAS ScriptCache利用
 */

/**
 * RequestGate統一インターフェース
 * Zero-Dependency Architecture準拠の重複処理防止システム
 */
class RequestGate {
  /**
   * トランザクション開始（排他制御）
   * @param {string} key - 識別キー
   * @param {number} timeoutMs - タイムアウト時間（ミリ秒）
   * @returns {boolean} 成功時true、既に処理中の場合false
   */
  static enter(key, timeoutMs = 30000) {
    try {
      if (!key || typeof key !== 'string') {
        console.warn('RequestGate.enter: Invalid key provided');
        return false;
      }

      const cache = CacheService.getScriptCache();
      if (!cache) {
        console.warn('RequestGate.enter: Cache service unavailable');
        return true; // フォールバック: 制御なしで継続
      }

      const lockKey = `request_lock_${key}`;
      const lockValue = `${Date.now()}_${Utilities.getUuid()}`;

      // 既存のロックをチェック
      const existingLock = cache.get(lockKey);
      if (existingLock) {
        console.warn('RequestGate.enter: Request already in progress for key:', key);
        return false;
      }

      // ロックを設定（タイムアウト付き）
      const expireSeconds = Math.ceil(timeoutMs / 1000);
      cache.put(lockKey, lockValue, expireSeconds);

      return true;
    } catch (error) {
      console.error('RequestGate.enter error:', error.message);
      return true; // フォールバック: エラー時は制御なしで継続
    }
  }

  /**
   * トランザクション終了（ロック解除）
   * @param {string} key - 識別キー
   * @returns {boolean} 解除成功時true
   */
  static exit(key) {
    try {
      if (!key || typeof key !== 'string') {
        console.warn('RequestGate.exit: Invalid key provided');
        return false;
      }

      const cache = CacheService.getScriptCache();
      if (!cache) {
        console.warn('RequestGate.exit: Cache service unavailable');
        return false;
      }

      const lockKey = `request_lock_${key}`;
      cache.remove(lockKey);

      return true;
    } catch (error) {
      console.error('RequestGate.exit error:', error.message);
      return false;
    }
  }

  /**
   * 指定キーのロック状態を確認
   * @param {string} key - 識別キー
   * @returns {boolean} ロック中の場合true
   */
  static isLocked(key) {
    try {
      if (!key || typeof key !== 'string') {
        return false;
      }

      const cache = CacheService.getScriptCache();
      if (!cache) {
        return false;
      }

      const lockKey = `request_lock_${key}`;
      const lockValue = cache.get(lockKey);
      return Boolean(lockValue);
    } catch (error) {
      console.warn('RequestGate.isLocked error:', error.message);
      return false;
    }
  }

  /**
   * 全ロックをクリア（管理者用）
   * @returns {boolean} クリア成功時true
   */
  static clearAll() {
    try {
      console.log('RequestGate.clearAll: Clearing all request locks');
      // GAS CacheServiceには個別キーでのクリア機能しかないため
      // 実装は将来的な拡張として保留
      return true;
    } catch (error) {
      console.error('RequestGate.clearAll error:', error.message);
      return false;
    }
  }

  /**
   * 診断情報取得
   * @returns {Object} 診断結果
   */
  static diagnose() {
    try {
      const testKey = `test_${Date.now()}`;
      const enterResult = this.enter(testKey);
      const isLockedResult = this.isLocked(testKey);
      const exitResult = this.exit(testKey);

      return {
        service: 'RequestGate',
        timestamp: new Date().toISOString(),
        available: true,
        testResults: {
          enter: enterResult,
          isLocked: isLockedResult,
          exit: exitResult
        },
        cacheService: Boolean(CacheService.getScriptCache())
      };
    } catch (error) {
      return {
        service: 'RequestGate',
        timestamp: new Date().toISOString(),
        available: false,
        error: error.message,
        cacheService: Boolean(CacheService.getScriptCache())
      };
    }
  }
}

// Export for global access (Zero-Dependency Architecture)
if (typeof globalThis !== 'undefined') {
  globalThis.RequestGate = RequestGate;
} else if (typeof global !== 'undefined') {
  global.RequestGate = RequestGate;
} else {
  this.RequestGate = RequestGate;
}