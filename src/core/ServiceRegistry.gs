/**
 * @fileoverview ServiceRegistry - サービス間依存関係管理
 *
 * 🎯 目的:
 * - 循環依存の解消
 * - 遅延初期化（Lazy Loading）による安全な参照
 * - サービス間の疎結合化
 *
 * 🔄 解決する循環依存:
 * - DataService ↔ ConfigService
 * - SecurityService ↔ UserService
 * - 各Service ↔ Database層
 */

/* global UserService, ConfigService, DataService, SecurityService, PROPS_KEYS, DB */

/**
 * サービス登録レジストリ
 * 遅延初期化で循環依存を回避
 */
// eslint-disable-next-line no-unused-vars
const ServiceRegistry = Object.freeze({
  
  // ==========================================
  // 🔄 遅延初期化 - 循環依存回避パターン
  // ==========================================
  
  /**
   * 安全なUserService参照取得
   * @returns {Object} UserService オブジェクト
   */
  getUserService() {
    return this._getObjectWithRetry('UserService', () => UserService);
  },

  /**
   * 安全なConfigService参照取得
   * @returns {Object} ConfigService オブジェクト
   */
  getConfigService() {
    return this._getObjectWithRetry('ConfigService', () => ConfigService);
  },

  /**
   * 安全なDataService参照取得
   * @returns {Object} DataService オブジェクト
   */
  getDataService() {
    return this._getObjectWithRetry('DataService', () => DataService);
  },

  /**
   * 安全なSecurityService参照取得
   * @returns {Object} SecurityService オブジェクト
   */
  getSecurityService() {
    return this._getObjectWithRetry('SecurityService', () => SecurityService);
  },

  // ✅ NOTE: Controller getters removed - Controllers now use flat functions architecture
  // ✅ NOTE: Service getters removed - Direct service calls recommended for better performance

  /**
   * 安全なDatabaseService参照取得（再試行機能付き）
   * @returns {Object} DatabaseService オブジェクト
   * @deprecated Use direct DatabaseService calls instead
   */
  getDatabaseService() {
    // For DatabaseService, we still need this as it's an Object.freeze structure
    try {
      if (typeof DB !== 'undefined') {
        return DB;
      }
      throw new Error('DatabaseService not available');
    } catch (error) {
      console.error('ServiceRegistry.getDatabaseService: Error accessing DatabaseService', error);
      return null;
    }
  },

  /**
   * 汎用オブジェクト取得メソッド（再試行機能付き）
   * @param {string} objectName - オブジェクト名
   * @param {Function} accessor - オブジェクトアクセサー関数
   * @returns {Object} オブジェクト
   * @private
   */
  _getObjectWithRetry(objectName, accessor) {
    // 最大3回まで再試行（指数バックオフ）
    for (let attempt = 0; attempt < 3; attempt++) {
      try {
        const obj = accessor();
        if (typeof obj !== 'undefined' && obj !== null) {
          return obj;
        }
      } catch (error) {
        // ReferenceError の場合は続行、その他は再スロー
        if (!(error instanceof ReferenceError)) {
          throw error;
        }
      }

      console.warn(`ServiceRegistry._getObjectWithRetry: ${objectName}未定義 - 試行 ${attempt + 1}/3`);
      if (attempt < 2) { // 最後の試行では待機しない
        Utilities.sleep(50 * Math.pow(2, attempt)); // 50ms, 100ms, 200ms
      }
    }

    throw new Error(`${objectName} not initialized after 3 attempts`);
  },
  
  // ==========================================
  // 🎯 共通機能メソッド - 循環依存なし
  // ==========================================
  
  /**
   * 現在のユーザーメール取得（共通処理）
   * UserService への依存を避けた直接実装
   * @returns {string|null} ユーザーメールアドレス
   */
  getCurrentUserEmail() {
    try {
      const email = Session.getActiveUser().getEmail();
      if (!email) {
        console.warn('ServiceRegistry.getCurrentUserEmail: セッションからメールアドレスを取得できません');
        return null;
      }
      return email;
    } catch (error) {
      console.error('ServiceRegistry.getCurrentUserEmail: エラー', error.message);
      return null;
    }
  },
  
  /**
   * システム管理者判定（共通処理）
   * UserService への依存を避けた直接実装
   * @returns {boolean} システム管理者かどうか
   */
  isSystemAdmin() {
    try {
      const currentEmail = this.getCurrentUserEmail();
      if (!currentEmail) {
        return false;
      }
      
      const props = PropertiesService.getScriptProperties();
      const adminEmail = props.getProperty(PROPS_KEYS.ADMIN_EMAIL);
      
      if (!adminEmail) {
        console.warn('ServiceRegistry.isSystemAdmin: ADMIN_EMAILが設定されていません');
        return false;
      }
      
      const isAdmin = currentEmail === adminEmail;
      if (isAdmin) {
        console.info('ServiceRegistry.isSystemAdmin: システム管理者を認証', { email: currentEmail });
      }
      
      return isAdmin;
    } catch (error) {
      console.error('ServiceRegistry.isSystemAdmin: エラー', {
        error: error.message
      });
      return false;
    }
  },
  
  /**
   * キャッシュサービスの安全なアクセス
   * @param {string} key - キャッシュキー
   * @param {*} defaultValue - デフォルト値
   * @returns {*} キャッシュされた値またはデフォルト値
   */
  getCacheValue(key, defaultValue = null) {
    try {
      const cached = CacheService.getScriptCache().get(key);
      return cached ? JSON.parse(cached) : defaultValue;
    } catch (error) {
      console.warn(`ServiceRegistry.getCacheValue: キャッシュ取得エラー - ${key}`, error.message);
      return defaultValue;
    }
  },
  
  /**
   * キャッシュサービスの安全な設定
   * @param {string} key - キャッシュキー
   * @param {*} value - 設定する値
   * @param {number} ttl - TTL秒数
   * @returns {boolean} 成功・失敗
   */
  setCacheValue(key, value, ttl = 300) {
    try {
      CacheService.getScriptCache().put(key, JSON.stringify(value), ttl);
      return true;
    } catch (error) {
      console.warn(`ServiceRegistry.setCacheValue: キャッシュ設定エラー - ${key}`, error.message);
      return false;
    }
  },
  
  /**
   * プロパティサービスの安全なアクセス
   * @param {string} key - プロパティキー
   * @param {string} defaultValue - デフォルト値
   * @returns {string|null} プロパティ値またはデフォルト値
   */
  getProperty(key, defaultValue = null) {
    try {
      const props = PropertiesService.getScriptProperties();
      return props.getProperty(key) || defaultValue;
    } catch (error) {
      console.warn(`ServiceRegistry.getProperty: プロパティ取得エラー - ${key}`, error.message);
      return defaultValue;
    }
  },
  
  /**
   * 安全なDB操作実行
   * infrastructure/DatabaseService.gs の関数への安全なアクセス
   * @param {Function} operation - 実行する操作
   * @param {string} operationName - 操作名（ログ用）
   * @returns {*} 操作結果
   */
  executeDbOperation(operation, operationName = 'unknown') {
    try {
      return operation();
    } catch (error) {
      console.error(`ServiceRegistry.executeDbOperation: DB操作エラー - ${operationName}`, error.message);
      throw error;
    }
  },
  
  // ==========================================
  // 🔧 診断・テスト機能
  // ==========================================
  
  /**
   * サービスレジストリ診断
   * @returns {Object} 診断結果
   */
  diagnose() {
    const results = {
      service: 'ServiceRegistry',
      timestamp: new Date().toISOString(),
      checks: []
    };
    
    // サービス・コントローラー・インフラ層可用性チェック
    const objects = [
      'UserService', 'ConfigService', 'DataService', 'SecurityService',
      'AdminController', 'DataController', 'FrontendController', 'SystemController',
      'CacheService', 'DatabaseService'
    ];
    objects.forEach(objectName => {
      try {
        const methodName = objectName.includes('Controller') ?
          `get${objectName}` : `get${objectName}`;
        const service = this[methodName]();
        results.checks.push({
          name: `${objectName} Availability`,
          status: service ? '✅' : '❌',
          details: `${objectName} service accessibility`
        });
      } catch (error) {
        results.checks.push({
          name: `${objectName} Availability`,
          status: '❌',
          details: error.message
        });
      }
    });
    
    // 基本機能チェック
    results.checks.push({
      name: 'Session Access',
      status: this.getCurrentUserEmail() ? '✅' : '❌',
      details: 'Session.getActiveUser() access'
    });
    
    results.checks.push({
      name: 'Properties Access',
      status: this.getProperty('test', 'default') !== null ? '✅' : '❌',
      details: 'PropertiesService access'
    });
    
    results.checks.push({
      name: 'Cache Access',
      status: this.setCacheValue('test', 'value', 1) ? '✅' : '❌',
      details: 'CacheService access'
    });
    
    results.overall = results.checks.every(check => check.status === '✅') ? '✅' : '⚠️';
    
    return results;
  }
  
});