/**
 * @fileoverview ServiceFactory - 統一サービスアクセス層
 *
 * 🎯 責任範囲:
 * - GAS Platform APIへの統一アクセス
 * - 依存関係の除去とゼロ依存アーキテクチャ
 * - サービス読み込み順序問題の解決
 * - パフォーマンス最適化
 *
 * 🔄 GAS Best Practices準拠:
 * - 直接Platform API利用
 * - 依存関係なしの自己完結型
 * - エラー処理とフォールバック機能
 */

/* global DB, DatabaseOperations */

/**
 * ServiceFactory - ゼロ依存サービスアクセス層
 * GASファイル読み込み順序に依存しない堅牢なアーキテクチャ
 */
const ServiceFactory = Object.freeze({

  // ===========================================
  // 🔧 Database Operations
  // ===========================================

  /**
   * データベース操作オブジェクト取得
   * @returns {Object|null} DatabaseOperations or fallback
   */
  getDB() {
    try {
      // Method 1: グローバルDB変数
      if (typeof DB !== 'undefined' && DB) {
        return DB;
      }

      // Method 2: DatabaseOperations直接参照
      if (typeof DatabaseOperations !== 'undefined') {
        return DatabaseOperations;
      }

      // Method 3: Fallback - null return
      console.warn('ServiceFactory.getDB: Database service not available');
      return null;
    } catch (error) {
      console.error('ServiceFactory.getDB error:', error.message);
      return null;
    }
  },

  // ===========================================
  // 🔐 Session Management
  // ===========================================

  /**
   * 現在のユーザーセッション情報取得
   * @returns {Object} Session information
   */
  getSession() {
    try {
      const session = {
        email: null,
        effectiveEmail: null,
        isValid: false
      };

      // Primary method: Session.getActiveUser()
      try {
        session.email = Session.getActiveUser().getEmail();
        if (session.email) {
          session.isValid = true;
        }
      } catch (activeError) {
        console.warn('ServiceFactory.getSession: getActiveUser failed:', activeError.message);
      }

      // Fallback method: Session.getEffectiveUser()
      if (!session.email) {
        try {
          session.effectiveEmail = Session.getEffectiveUser().getEmail();
          if (session.effectiveEmail) {
            session.email = session.effectiveEmail;
            session.isValid = true;
          }
        } catch (effectiveError) {
          console.warn('ServiceFactory.getSession: getEffectiveUser failed:', effectiveError.message);
        }
      }

      return session;
    } catch (error) {
      console.error('ServiceFactory.getSession error:', error.message);
      return { email: null, effectiveEmail: null, isValid: false };
    }
  },

  // ===========================================
  // ⚙️ Properties Management
  // ===========================================

  /**
   * システムプロパティ管理オブジェクト取得
   * @returns {Object} Properties service wrapper
   */
  getProperties() {
    try {
      const props = PropertiesService.getScriptProperties();

      return {
        get: (key) => {
          try {
            return props.getProperty(key);
          } catch (error) {
            console.error(`ServiceFactory.getProperties.get(${key}):`, error.message);
            return null;
          }
        },

        set: (key, value) => {
          try {
            props.setProperty(key, value);
            return true;
          } catch (error) {
            console.error(`ServiceFactory.getProperties.set(${key}):`, error.message);
            return false;
          }
        },

        getAll: () => {
          try {
            return props.getProperties();
          } catch (error) {
            console.error('ServiceFactory.getProperties.getAll:', error.message);
            return {};
          }
        },

        // 直接値取得（定数インライン化）
        getServiceAccountCreds: () => props.getProperty('SERVICE_ACCOUNT_CREDS'),
        getDatabaseSpreadsheetId: () => props.getProperty('DATABASE_SPREADSHEET_ID'),
        getAdminEmail: () => props.getProperty('ADMIN_EMAIL'),
        getGoogleClientId: () => props.getProperty('GOOGLE_CLIENT_ID')
      };
    } catch (error) {
      console.error('ServiceFactory.getProperties error:', error.message);
      return {
        get: () => null,
        set: () => false,
        getAll: () => ({}),
        getServiceAccountCreds: () => null,
        getDatabaseSpreadsheetId: () => null,
        getAdminEmail: () => null,
        getGoogleClientId: () => null
      };
    }
  },

  // ===========================================
  // 🗄️ Cache Management
  // ===========================================

  /**
   * キャッシュサービス管理オブジェクト取得
   * @returns {Object} Cache service wrapper
   */
  getCache() {
    try {
      const cache = CacheService.getScriptCache();

      return {
        get: (key, defaultValue = null) => {
          try {
            const cached = cache.get(key);
            return cached ? JSON.parse(cached) : defaultValue;
          } catch (error) {
            console.warn(`ServiceFactory.getCache.get(${key}):`, error.message);
            return defaultValue;
          }
        },

        put: (key, value, ttlSeconds = 300) => {
          try {
            cache.put(key, JSON.stringify(value), ttlSeconds);
            return true;
          } catch (error) {
            console.error(`ServiceFactory.getCache.put(${key}):`, error.message);
            return false;
          }
        },

        remove: (key) => {
          try {
            cache.remove(key);
            return true;
          } catch (error) {
            console.error(`ServiceFactory.getCache.remove(${key}):`, error.message);
            return false;
          }
        },

        removeAll: () => {
          try {
            cache.removeAll();
            return true;
          } catch (error) {
            console.error('ServiceFactory.getCache.removeAll:', error.message);
            return false;
          }
        }
      };
    } catch (error) {
      console.error('ServiceFactory.getCache error:', error.message);
      return {
        get: () => null,
        put: () => false,
        remove: () => false,
        removeAll: () => false
      };
    }
  },

  // ===========================================
  // 📊 Spreadsheet Operations
  // ===========================================

  /**
   * スプレッドシート操作ヘルパー
   * @returns {Object} Spreadsheet operations wrapper
   */
  getSpreadsheet() {
    return {
      openById: (spreadsheetId) => {
        try {
          return SpreadsheetApp.openById(spreadsheetId);
        } catch (error) {
          console.error(`ServiceFactory.getSpreadsheet.openById(${spreadsheetId}):`, error.message);
          return null;
        }
      },

      getActiveSpreadsheet: () => {
        try {
          return SpreadsheetApp.getActiveSpreadsheet();
        } catch (error) {
          console.warn('ServiceFactory.getSpreadsheet.getActiveSpreadsheet:', error.message);
          return null;
        }
      }
    };
  },

  // ===========================================
  // 🔧 Utility Functions
  // ===========================================

  /**
   * ユーティリティ関数群
   * @returns {Object} Utility functions
   */
  getUtils() {
    return {
      generateUserId: () => {
        return Utilities.getUuid();
      },

      getCurrentTimestamp: () => {
        return new Date().toISOString();
      },

      validateEmail: (email) => {
        const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
        return emailRegex.test(email);
      },

      sanitizeInput: (input) => {
        if (typeof input !== 'string') return '';
        return input.replace(/[<>&"']/g, '');
      }
    };
  },

  // ===========================================
  // 🏥 Health Check & Diagnostics
  // ===========================================

  /**
   * サービス診断情報
   * @returns {Object} Diagnostic information
   */
  diagnose() {
    const startTime = Date.now();

    const diagnostics = {
      timestamp: new Date().toISOString(),
      factory: 'ServiceFactory',
      version: '1.0.0',
      services: {},
      platform: {},
      executionTime: 0
    };

    try {
      // Database service check
      diagnostics.services.database = {
        available: this.getDB() !== null,
        type: typeof DB !== 'undefined' ? 'Global' :
              typeof DatabaseOperations !== 'undefined' ? 'Direct' : 'None'
      };

      // Session service check
      const session = this.getSession();
      diagnostics.services.session = {
        available: session.isValid,
        email: session.email ? 'Available' : 'None'
      };

      // Properties service check
      const props = this.getProperties();
      diagnostics.services.properties = {
        available: props.getAdminEmail() !== null
      };

      // Cache service check
      const cache = this.getCache();
      const testKey = `test_key_${  Date.now()}`;
      cache.put(testKey, 'test', 10);
      const testResult = cache.get(testKey);
      cache.remove(testKey);
      diagnostics.services.cache = {
        available: testResult === 'test'
      };

      // Platform APIs check
      diagnostics.platform = {
        session: typeof Session !== 'undefined',
        properties: typeof PropertiesService !== 'undefined',
        cache: typeof CacheService !== 'undefined',
        spreadsheet: typeof SpreadsheetApp !== 'undefined',
        utilities: typeof Utilities !== 'undefined'
      };

    } catch (error) {
      diagnostics.error = error.message;
    }

    diagnostics.executionTime = `${Date.now() - startTime}ms`;
    return diagnostics;
  }

});

// ===========================================
// 🌐 グローバル Export
// ===========================================

/**
 * レガシー互換性のためのグローバル関数
 * 既存コードの段階的移行を支援
 */

/**
 * 現在のユーザーメール取得（ServiceFactory版）
 * @returns {string|null} User email
 */
function getCurrentEmailFromFactory() {
  return ServiceFactory.getSession().email;
}

/**
 * システムプロパティ取得（ServiceFactory版）
 * @param {string} key - Property key
 * @returns {string|null} Property value
 */
function getSystemPropertyFromFactory(key) {
  return ServiceFactory.getProperties().get(key);
}

/**
 * データベース操作取得（ServiceFactory版）
 * @returns {Object|null} Database operations
 */
function getDatabaseFromFactory() {
  return ServiceFactory.getDB();
}