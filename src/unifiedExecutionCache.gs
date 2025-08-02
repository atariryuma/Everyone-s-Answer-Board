/**
 * 統一実行レベルキャッシュ管理システム
 * 全モジュール間でのキャッシュ共有と効率的な管理
 */

class UnifiedExecutionCache {
  constructor() {
    this.userInfoCache = null;
    this.sheetsServiceCache = null;
    this.lastUserIdKey = null;
    this.executionStartTime = Date.now();
    this.maxLifetime = 300000; // 5分間の最大ライフタイム
  }

  /**
   * ユーザー情報キャッシュを取得
   * @param {string} userId - ユーザーID
   * @returns {object|null} キャッシュされたユーザー情報
   */
  getUserInfo(userId) {
    if (this.isExpired()) {
      this.clearAll();
      return null;
    }

    if (this.userInfoCache && this.lastUserIdKey === userId) {
      debugLog(`✅ 統一キャッシュヒット: ユーザー情報 (${userId ? '***' : 'none'})`);
      return this.userInfoCache;
    }

    return null;
  }

  /**
   * ユーザー情報をキャッシュに保存
   * @param {string} userId - ユーザーID
   * @param {object} userInfo - ユーザー情報
   */
  setUserInfo(userId, userInfo) {
    this.userInfoCache = userInfo;
    this.lastUserIdKey = userId;
    debugLog(`💾 統一キャッシュ保存: ユーザー情報 (${userId ? '***' : 'none'})`);
  }

  /**
   * SheetsServiceキャッシュを取得
   * @returns {object|null} キャッシュされたSheetsService
   */
  getSheetsService() {
    if (this.isExpired()) {
      this.clearAll();
      return null;
    }

    if (this.sheetsServiceCache) {
      debugLog(`✅ 統一キャッシュヒット: SheetsService`);
      return this.sheetsServiceCache;
    }

    return null;
  }

  /**
   * SheetsServiceをキャッシュに保存
   * @param {object} service - SheetsService
   */
  setSheetsService(service) {
    this.sheetsServiceCache = service;
    debugLog(`💾 統一キャッシュ保存: SheetsService`);
  }

  /**
   * ユーザー情報キャッシュをクリア
   */
  clearUserInfo() {
    this.userInfoCache = null;
    this.lastUserIdKey = null;
    debugLog(`🗑️ 統一キャッシュクリア: ユーザー情報`);
  }

  /**
   * SheetsServiceキャッシュをクリア
   */
  clearSheetsService() {
    this.sheetsServiceCache = null;
    debugLog(`🗑️ 統一キャッシュクリア: SheetsService`);
  }

  /**
   * 全キャッシュをクリア
   */
  clearAll() {
    this.clearUserInfo();
    this.clearSheetsService();
    debugLog(`🗑️ 統一キャッシュ全クリア`);
  }

  /**
   * キャッシュが期限切れかチェック
   * @returns {boolean} 期限切れの場合true
   */
  isExpired() {
    return (Date.now() - this.executionStartTime) > this.maxLifetime;
  }

  /**
   * キャッシュ統計情報を取得
   * @returns {object} キャッシュ統計
   */
  getStats() {
    return {
      hasUserInfo: !!this.userInfoCache,
      hasSheetsService: !!this.sheetsServiceCache,
      lastUserIdKey: this.lastUserIdKey,
      executionTime: Date.now() - this.executionStartTime,
      isExpired: this.isExpired()
    };
  }

  /**
   * 統一キャッシュマネージャーとの同期
   * @param {string} operation - 操作タイプ ('userDataChange', 'configChange', 'systemChange')
   */
  syncWithUnifiedCache(operation) {
    if (typeof cacheManager !== 'undefined' && cacheManager) {
      try {
        // パフォーマンス最適化: 必要最小限のキャッシュクリアのみ実行
        switch (operation) {
          case 'userDataChange':
            // ユーザー特定データのみクリア
            if (this.lastUserIdKey) {
              cacheManager.remove(`user_${this.lastUserIdKey}`);
              // userinfo_は重複なので除去
            }
            break;
          case 'configChange':
            // 設定変更時のみシステム設定クリア
            cacheManager.remove('system_config');
            break;
          case 'systemChange':
            // 重大なシステム変更時のみ全体クリア（慎重に使用）
            debugLog('🚨 システム全体キャッシュクリアは重大操作のため慎重に実行');
            break;
        }
        debugLog(`🔄 統一キャッシュマネージャーと同期: ${operation}`);
      } catch (error) {
        debugLog(`⚠️ 統一キャッシュマネージャー同期エラー: ${error.message}`);
      }
    }
  }
}

// グローバル統一キャッシュインスタンス
let globalUnifiedCache = null;

/**
 * 統一実行キャッシュのインスタンスを取得
 * @returns {UnifiedExecutionCache} 統一キャッシュインスタンス
 */
function getUnifiedExecutionCache() {
  if (!globalUnifiedCache) {
    globalUnifiedCache = new UnifiedExecutionCache();
    debugLog(`🏗️ 統一実行キャッシュ初期化`);
  }
  return globalUnifiedCache;
}

/**
 * 統一実行キャッシュをリセット
 */
function resetUnifiedExecutionCache() {
  globalUnifiedCache = null;
  debugLog(`🔄 統一実行キャッシュリセット`);
}

// 後方互換性のための関数
function clearExecutionUserInfoCache() {
  const cache = getUnifiedExecutionCache();
  cache.clearUserInfo();
  cache.syncWithUnifiedCache('userDataChange');
}

function clearExecutionSheetsServiceCache() {
  const cache = getUnifiedExecutionCache();
  cache.clearSheetsService();
}

function clearAllExecutionCache() {
  const cache = getUnifiedExecutionCache();
  cache.clearAll();
  cache.syncWithUnifiedCache('systemChange');
}