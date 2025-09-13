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

/* global UserService, ConfigService, DataService, SecurityService, PROPS_KEYS */

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
    if (typeof UserService === 'undefined') {
      throw new Error('UserService not initialized');
    }
    return UserService;
  },
  
  /**
   * 安全なConfigService参照取得
   * @returns {Object} ConfigService オブジェクト
   */
  getConfigService() {
    if (typeof ConfigService === 'undefined') {
      throw new Error('ConfigService not initialized');
    }
    return ConfigService;
  },
  
  /**
   * 安全なDataService参照取得
   * @returns {Object} DataService オブジェクト
   */
  getDataService() {
    if (typeof DataService === 'undefined') {
      throw new Error('DataService not initialized');
    }
    return DataService;
  },
  
  /**
   * 安全なSecurityService参照取得
   * @returns {Object} SecurityService オブジェクト
   */
  getSecurityService() {
    if (typeof SecurityService === 'undefined') {
      throw new Error('SecurityService not initialized');
    }
    return SecurityService;
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
    
    // サービス可用性チェック
    const services = ['UserService', 'ConfigService', 'DataService', 'SecurityService'];
    services.forEach(serviceName => {
      try {
        const service = this[`get${serviceName}`]();
        results.checks.push({
          name: `${serviceName} Availability`,
          status: service ? '✅' : '❌',
          details: `${serviceName} service accessibility`
        });
      } catch (error) {
        results.checks.push({
          name: `${serviceName} Availability`,
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