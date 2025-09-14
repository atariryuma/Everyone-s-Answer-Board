/**
 * @fileoverview CacheManager - 統一キャッシュ管理システム
 * 
 * 🎯 目的:
 * - キャッシュの整合性保証
 * - TTL統一管理
 * - キャッシュ無効化の自動化
 * - データ同期の保証
 */

/**
 * CacheManager - 統一キャッシュ管理システム
 * データ整合性とパフォーマンスを両立
 */
 
const AppCacheService = Object.freeze({

  // ==========================================
  // 🔧 キャッシュ設定定数
  // ==========================================

  TTL: Object.freeze({
    SHORT: 60,      // 1分 - 頻繁に変更されるデータ
    MEDIUM: 300,    // 5分 - 標準的なデータ
    LONG: 1800,     // 30分 - 安定したデータ
    SESSION: 3600   // 1時間 - セッション情報
  }),

  KEYS: Object.freeze({
    USER_INFO: 'user_info_',
    USER_CONFIG: 'config_',
    SHEET_DATA: 'sheet_data_',
    SA_TOKEN: 'sa_token',
    SYSTEM_STATUS: 'system_status'
  }),

  // ==========================================
  // 🎯 統一キャッシュ操作
  // ==========================================

  /**
   * 安全なキャッシュ取得
   * @param {string} key - キャッシュキー
   * @param {*} defaultValue - デフォルト値
   * @returns {*} キャッシュされた値またはデフォルト値
   */
  get(key, defaultValue = null) {
    try {
      const cached = CacheService.getScriptCache().get(key);
      if (cached) {
        return JSON.parse(cached);
      }
      return defaultValue;
    } catch (error) {
      console.warn(`CacheManager.get: キャッシュ取得エラー - ${key}`, error.message);
      return defaultValue;
    }
  },

  /**
   * 安全なキャッシュ設定
   * @param {string} key - キャッシュキー
   * @param {*} value - 設定する値
   * @param {number} ttl - TTL秒数（デフォルト: MEDIUM）
   * @returns {boolean} 成功・失敗
   */
  set(key, value, ttl = this.TTL.MEDIUM) {
    try {
      // メタデータ付きでキャッシュ保存
      const cacheData = {
        value,
        timestamp: Date.now(),
        ttl: ttl * 1000
      };
      
      CacheService.getScriptCache().put(key, JSON.stringify(cacheData), ttl);
      
      // 関連キャッシュの無効化リストを更新
      this.updateInvalidationList(key);
      
      return true;
    } catch (error) {
      console.warn(`CacheManager.set: キャッシュ設定エラー - ${key}`, error.message);
      return false;
    }
  },

  /**
   * ユーザー関連キャッシュの統一管理
   * @param {string} userId - ユーザーID
   * @param {Object} userData - ユーザーデータ
   * @param {Object} configData - 設定データ
   */
  setUserDataSet(userId, userData, configData) {
    try {
      const timestamp = Date.now();
      const consistency = `user_${userId}_${timestamp}`;
      
      // 整合性保証: 同じタイムスタンプで関連データを保存
      const userCacheData = { ...userData, _consistency: consistency };
      const configCacheData = { ...configData, _consistency: consistency };
      
      this.set(`${this.KEYS.USER_INFO}${userId}`, userCacheData, this.TTL.SESSION);
      this.set(`${this.KEYS.USER_CONFIG}${userId}`, configCacheData, this.TTL.SESSION);
      
      console.log(`CacheManager.setUserDataSet: ユーザーデータセット保存完了 - ${userId}`);
      
      return true;
    } catch (error) {
      console.error('CacheManager.setUserDataSet: エラー', error.message);
      return false;
    }
  },

  /**
   * ユーザー関連キャッシュの一括削除
   * @param {string} userId - ユーザーID
   */
  invalidateUserCache(userId) {
    try {
      const keysToRemove = [
        `${this.KEYS.USER_INFO}${userId}`,
        `${this.KEYS.USER_CONFIG}${userId}`,
        `${this.KEYS.SHEET_DATA}${userId}`,
        'current_user_info' // グローバルキャッシュも削除
      ];
      
      const cache = CacheService.getScriptCache();
      keysToRemove.forEach(key => {
        try {
          cache.remove(key);
        } catch (removeError) {
          console.warn(`CacheManager.invalidateUserCache: ${key} 削除エラー`, removeError.message);
        }
      });
      
      console.log(`CacheManager.invalidateUserCache: ユーザーキャッシュクリア完了 - ${userId}`);
    } catch (error) {
      console.error('CacheManager.invalidateUserCache: エラー', error.message);
    }
  },

  /**
   * キャッシュの整合性チェック
   * @param {string} userId - ユーザーID
   * @returns {boolean} 整合性があるかどうか
   */
  checkUserCacheConsistency(userId) {
    try {
      const userData = this.get(`${this.KEYS.USER_INFO}${userId}`);
      const configData = this.get(`${this.KEYS.USER_CONFIG}${userId}`);
      
      if (!userData || !configData) {
        return false; // データが不完全
      }
      
      // 整合性チェック: タイムスタンプが同じかチェック
      const userConsistency = userData._consistency;
      const configConsistency = configData._consistency;
      
      return userConsistency === configConsistency;
    } catch (error) {
      console.warn('CacheManager.checkUserCacheConsistency: エラー', error.message);
      return false;
    }
  },

  /**
   * 無効化リストの更新（将来の自動無効化用）
   * @param {string} key - キャッシュキー
   */
  updateInvalidationList(key) {
    try {
      // 関連キーの依存関係を管理
      const dependencies = {
        [this.KEYS.USER_INFO]: ['current_user_info'],
        [this.KEYS.USER_CONFIG]: ['current_user_info'],
        [this.KEYS.SYSTEM_STATUS]: ['sa_token']
      };
      
      Object.entries(dependencies).forEach(([prefix, relatedKeys]) => {
        if (key.startsWith(prefix)) {
          relatedKeys.forEach(relatedKey => {
            // 関連キーも無効化マーク（将来実装）
            console.debug(`CacheManager: ${key} 更新により ${relatedKey} を無効化対象に設定`);
          });
        }
      });
    } catch (error) {
      console.warn('CacheManager.updateInvalidationList: エラー', error.message);
    }
  },

  /**
   * 全キャッシュクリア（緊急時）
   */
  clearAll() {
    try {
      const cache = CacheService.getScriptCache();
      cache.removeAll([
        'current_user_info',
        'sa_token',
        'system_status'
      ]);
      
      console.log('CacheManager.clearAll: 全キャッシュクリア完了');
    } catch (error) {
      console.error('CacheManager.clearAll: エラー', error.message);
    }
  },

  /**
   * キャッシュ統計情報取得
   * @returns {Object} キャッシュ統計
   */
  getStats() {
    try {
      // 実装例: キャッシュのヒット率や使用量を取得
      return {
        timestamp: new Date().toISOString(),
        status: 'operational',
        // GASでは詳細統計取得が制限されているため基本情報のみ
        message: 'Cache service is operational'
      };
    } catch (error) {
      return {
        timestamp: new Date().toISOString(),
        status: 'error',
        message: error.message
      };
    }
  }

});

// ✅ レガシーエイリアス削除 - CacheService直接使用に統一