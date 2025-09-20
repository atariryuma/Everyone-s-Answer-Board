/**
 * @fileoverview Config - 統一設定管理システム
 *
 * 責任範囲:
 * - 設定の単一エントリーポイント (Single Source of Truth)
 * - サービスアカウント設定一元化
 * - Zero-Dependency Architecture準拠（直接関数実装）
 * - 設定値検証・暗号化
 */

/* global validateEmail, validateSpreadsheetId, getServiceAccount */

// ===========================================
// ⚙️ 統一設定管理
// ===========================================

/**
 * データベース設定取得
 * @returns {Object|null} Database configuration
 */
function getDatabaseConfig() {
  try {
    const props = PropertiesService.getScriptProperties();
    const databaseId = props.getProperty('DATABASE_SPREADSHEET_ID');

    if (!databaseId) {
      return null;
    }

    return {
      spreadsheetId: databaseId,
      isValid: validateSpreadsheetId(databaseId).isValid
    };
  } catch (error) {
    console.error('getDatabaseConfig error:', error.message);
    return null;
  }
}

/**
 * 管理者設定取得
 * @returns {Object|null} Admin configuration
 */
function getAdminConfig() {
  try {
    const props = PropertiesService.getScriptProperties();
    const adminEmail = props.getProperty('ADMIN_EMAIL');

    if (!adminEmail) {
      return null;
    }

    return {
      email: adminEmail,
      isValid: validateEmail(adminEmail).isValid
    };
  } catch (error) {
    console.error('getAdminConfig error:', error.message);
    return null;
  }
}

/**
 * サービスアカウント設定存在確認
 * @returns {boolean} Service account configuration exists
 */
function hasServiceAccountConfig() {
  try {
    const props = PropertiesService.getScriptProperties();
    const credsJson = props.getProperty('SERVICE_ACCOUNT_CREDS');
    return !!credsJson;
  } catch (error) {
    console.error('hasServiceAccountConfig error:', error.message);
    return false;
  }
}

/**
 * Google OAuth クライアント設定取得
 * @returns {Object|null} OAuth client configuration
 */
function getOAuthConfig() {
  try {
    const props = PropertiesService.getScriptProperties();
    const clientId = props.getProperty('GOOGLE_CLIENT_ID');

    if (!clientId) {
      return null;
    }

    return {
      clientId,
      isValid: !!clientId && clientId.length > 20
    };
  } catch (error) {
    console.error('getOAuthConfig error:', error.message);
    return null;
  }
}

/**
 * システム設定の完全性確認
 * @returns {Object} System configuration status
 */
function getSystemConfigStatus() {
  const database = getDatabaseConfig();
  const admin = getAdminConfig();
  const serviceAccount = hasServiceAccountConfig();
  const oauth = getOAuthConfig();

  return {
    database: {
      configured: !!database,
      valid: database ? database.isValid : false
    },
    admin: {
      configured: !!admin,
      valid: admin ? admin.isValid : false
    },
    serviceAccount: {
      configured: serviceAccount,
      valid: serviceAccount
    },
    oauth: {
      configured: !!oauth,
      valid: oauth ? oauth.isValid : false
    },
    overall: !!(database && admin && serviceAccount && oauth) &&
             database.isValid && admin.isValid && oauth.isValid
  };
}

/**
 * 設定診断
 * @returns {Object} Configuration diagnostic information
 */
function diagnoseConfig() {
  const results = {
    service: 'Config',
    timestamp: new Date().toISOString(),
    checks: []
  };

  const status = getSystemConfigStatus();

  // Database configuration check
  results.checks.push({
    name: 'Database Configuration',
    status: status.database.valid ? '✅' : '❌',
    details: status.database.valid ? 'Database configuration valid' : 'Database not configured or invalid'
  });

  // Admin configuration check
  results.checks.push({
    name: 'Admin Configuration',
    status: status.admin.valid ? '✅' : '❌',
    details: status.admin.valid ? 'Admin configuration valid' : 'Admin not configured or invalid'
  });

  // Service account check
  results.checks.push({
    name: 'Service Account Configuration',
    status: status.serviceAccount.valid ? '✅' : '❌',
    details: status.serviceAccount.valid ? 'Service account configured' : 'Service account not configured'
  });

  // OAuth configuration check
  results.checks.push({
    name: 'OAuth Configuration',
    status: status.oauth.valid ? '✅' : '❌',
    details: status.oauth.valid ? 'OAuth configuration valid' : 'OAuth not configured or invalid'
  });

  results.overall = status.overall ? '✅' : '⚠️';
  return results;
}

// ===========================================
// 🔧 設定ヘルパー関数
// ===========================================

/**
 * 基本設定値を取得
 * @param {string} key - Property key
 * @returns {string|null} Property value
 */
function getProperty(key) {
  try {
    return PropertiesService.getScriptProperties().getProperty(key);
  } catch (error) {
    console.error('getProperty error:', error.message);
    return null;
  }
}

/**
 * 基本設定値を設定
 * @param {string} key - Property key
 * @param {string} value - Property value
 * @returns {boolean} Success status
 */
function setProperty(key, value) {
  try {
    PropertiesService.getScriptProperties().setProperty(key, value);
    return true;
  } catch (error) {
    console.error('setProperty error:', error.message);
    return false;
  }
}

/**
 * 複数の設定値を一括設定
 * @param {Object} properties - Properties object
 * @returns {boolean} Success status
 */
function setProperties(properties) {
  try {
    PropertiesService.getScriptProperties().setProperties(properties);
    return true;
  } catch (error) {
    console.error('setProperties error:', error.message);
    return false;
  }
}