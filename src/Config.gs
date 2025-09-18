/**
 * @fileoverview Config - 統一設定管理システム
 *
 * 責任範囲:
 * - 設定の単一エントリーポイント (Single Source of Truth)
 * - サービスアカウント設定一元化
 * - Zero-Dependency Architecture準拠
 * - 設定値検証・暗号化
 */

/* global ServiceFactory */

/**
 * 統一設定管理クラス
 * CLAUDE.md準拠のZero-Dependency実装
 */
class Config {
  /**
   * サービスアカウント設定取得
   * 全設定の単一エントリーポイント
   * @returns {Object|null} Service account configuration
   */
  static serviceAccount() {
    try {
      const props = ServiceFactory.getProperties();
      const credsJson = props.getProperty('SERVICE_ACCOUNT_CREDS');

      if (!credsJson) {
        return null;
      }

      const config = JSON.parse(credsJson);

      // 必須フィールド検証
      const requiredFields = ['client_email', 'private_key', 'project_id', 'type'];
      for (const field of requiredFields) {
        if (!config[field]) {
          throw new Error(`Missing required field: ${field}`);
        }
      }

      // Service account type validation
      if (config.type !== 'service_account') {
        throw new Error('Invalid service account type');
      }

      return config;
    } catch (error) {
      console.error('Config.serviceAccount error:', error.message);
      return null;
    }
  }

  /**
   * データベース設定取得
   * @returns {Object|null} Database configuration
   */
  static database() {
    try {
      const props = ServiceFactory.getProperties();
      const databaseId = props.getProperty('DATABASE_SPREADSHEET_ID');

      if (!databaseId) {
        return null;
      }

      return {
        spreadsheetId: databaseId,
        isValid: this.validateSpreadsheetId(databaseId)
      };
    } catch (error) {
      console.error('Config.database error:', error.message);
      return null;
    }
  }

  /**
   * 管理者設定取得
   * @returns {Object|null} Admin configuration
   */
  static admin() {
    try {
      const props = ServiceFactory.getProperties();
      const adminEmail = props.getProperty('ADMIN_EMAIL');

      if (!adminEmail) {
        return null;
      }

      return {
        email: adminEmail,
        isValid: this.validateEmail(adminEmail)
      };
    } catch (error) {
      console.error('Config.admin error:', error.message);
      return null;
    }
  }

  /**
   * システム設定取得
   * @returns {Object} System configuration
   */
  static system() {
    try {
      const serviceAccount = this.serviceAccount();
      const database = this.database();
      const admin = this.admin();

      return {
        serviceAccount,
        database,
        admin,
        isComplete: !!(serviceAccount && database && admin),
        version: '2.0.0'
      };
    } catch (error) {
      console.error('Config.system error:', error.message);
      return {
        serviceAccount: null,
        database: null,
        admin: null,
        isComplete: false,
        error: error.message
      };
    }
  }

  /**
   * 設定値設定
   * @param {string} key - Configuration key
   * @param {string} value - Configuration value
   * @returns {boolean} Success status
   */
  static set(key, value) {
    try {
      const props = ServiceFactory.getProperties();

      // Validation based on key type
      if (key === 'SERVICE_ACCOUNT_CREDS') {
        const config = JSON.parse(value);
        if (!this.validateServiceAccountConfig(config)) {
          throw new Error('Invalid service account configuration');
        }
      } else if (key === 'DATABASE_SPREADSHEET_ID') {
        if (!this.validateSpreadsheetId(value)) {
          throw new Error('Invalid spreadsheet ID format');
        }
      } else if (key === 'ADMIN_EMAIL') {
        if (!this.validateEmail(value)) {
          throw new Error('Invalid email format');
        }
      }

      props.setProperty(key, value);
      return true;
    } catch (error) {
      console.error('Config.set error:', error.message);
      return false;
    }
  }

  /**
   * バッチ設定
   * @param {Object} configs - Configuration object
   * @returns {boolean} Success status
   */
  static setBatch(configs) {
    try {
      const props = ServiceFactory.getProperties();

      // Validate all configs first
      for (const [key, value] of Object.entries(configs)) {
        if (!this.validateConfigValue(key, value)) {
          throw new Error(`Invalid configuration for key: ${key}`);
        }
      }

      // Set all configs if validation passes
      props.setProperties(configs);
      return true;
    } catch (error) {
      console.error('Config.setBatch error:', error.message);
      return false;
    }
  }

  /**
   * スプレッドシートID検証
   * @param {string} id - Spreadsheet ID
   * @returns {boolean} Validation result
   */
  static validateSpreadsheetId(id) {
    if (!id || typeof id !== 'string') {
      return false;
    }

    // Google Sheets ID format validation
    const pattern = /^[a-zA-Z0-9-_]{44}$/;
    return pattern.test(id);
  }

  /**
   * メールアドレス検証
   * @param {string} email - Email address
   * @returns {boolean} Validation result
   */
  static validateEmail(email) {
    if (!email || typeof email !== 'string') {
      return false;
    }

    const pattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return pattern.test(email);
  }

  /**
   * サービスアカウント設定検証
   * @param {Object} config - Service account configuration
   * @returns {boolean} Validation result
   */
  static validateServiceAccountConfig(config) {
    if (!config || typeof config !== 'object') {
      return false;
    }

    const requiredFields = ['type', 'project_id', 'private_key_id', 'private_key', 'client_email', 'client_id'];
    return requiredFields.every(field => config[field]) && config.type === 'service_account';
  }

  /**
   * 設定値検証
   * @param {string} key - Configuration key
   * @param {string} value - Configuration value
   * @returns {boolean} Validation result
   */
  static validateConfigValue(key, value) {
    switch (key) {
      case 'SERVICE_ACCOUNT_CREDS':
        try {
          const config = JSON.parse(value);
          return this.validateServiceAccountConfig(config);
        } catch {
          return false;
        }
      case 'DATABASE_SPREADSHEET_ID':
        return this.validateSpreadsheetId(value);
      case 'ADMIN_EMAIL':
        return this.validateEmail(value);
      default:
        return typeof value === 'string' && value.length > 0;
    }
  }

  /**
   * 設定診断
   * @returns {Object} Diagnostic information
   */
  static diagnose() {
    const results = {
      service: 'Config',
      timestamp: new Date().toISOString(),
      checks: []
    };

    // Service account check
    try {
      const sa = this.serviceAccount();
      results.checks.push({
        name: 'Service Account Configuration',
        status: sa ? '✅' : '❌',
        details: sa ? 'Service account configuration valid' : 'Service account configuration missing or invalid'
      });
    } catch (error) {
      results.checks.push({
        name: 'Service Account Configuration',
        status: '❌',
        details: error.message
      });
    }

    // Database check
    try {
      const db = this.database();
      results.checks.push({
        name: 'Database Configuration',
        status: db && db.isValid ? '✅' : '❌',
        details: db && db.isValid ? 'Database configuration valid' : 'Database configuration missing or invalid'
      });
    } catch (error) {
      results.checks.push({
        name: 'Database Configuration',
        status: '❌',
        details: error.message
      });
    }

    // Admin check
    try {
      const admin = this.admin();
      results.checks.push({
        name: 'Admin Configuration',
        status: admin && admin.isValid ? '✅' : '❌',
        details: admin && admin.isValid ? 'Admin configuration valid' : 'Admin configuration missing or invalid'
      });
    } catch (error) {
      results.checks.push({
        name: 'Admin Configuration',
        status: '❌',
        details: error.message
      });
    }

    results.overall = results.checks.every(check => check.status === '✅') ? '✅' : '⚠️';
    return results;
  }
}

// Export for global access
if (typeof globalThis !== 'undefined') {
  globalThis.Config = Config;
} else if (typeof global !== 'undefined') {
  global.Config = Config;
} else {
  this.Config = Config;
}