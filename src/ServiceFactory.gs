/* global Auth, Data, RequestGate */
/**
 * @fileoverview ServiceFactory - 統一サービスアクセス層
 *
 * 🎯 責任範囲:
 * - GAS Platform APIs統一アクセス
 * - Zero-Dependency Architecture実装
 * - Service Layer抽象化
 * - Cross-Service統合
 *
 * 🔄 GAS Best Practices準拠:
 * - 直接的な関数エクスポート
 * - フラット関数構造
 * - プラットフォームAPI統合
 */

/* global getUserSheetData, dsAddReaction, dsToggleHighlight, getUserConfig, validateUserData, validateSession, connectToSheetInternal, getFormInfo */

// ===========================================
// 🔧 Session Management
// ===========================================

function getSession() {
  try {
    const email = Session.getActiveUser().getEmail();
    return {
      isValid: Boolean(email),
      email: email || null
    };
  } catch (error) {
    console.warn('ServiceFactory.getSession: Session access error:', error.message);
    return {
      isValid: false,
      email: null
    };
  }
}

// ===========================================
// 📊 Properties Management
// ===========================================

function getProperties() {
  try {
    const scriptProps = PropertiesService.getScriptProperties();

    return {
      getDatabaseSpreadsheetId() {
        return scriptProps.getProperty('DATABASE_SPREADSHEET_ID');
      },

      getAdminEmail() {
        return scriptProps.getProperty('ADMIN_EMAIL');
      },

      getProperty(key) {
        return scriptProps.getProperty(key);
      },

      setProperty(key, value) {
        return scriptProps.setProperty(key, value);
      },

      setProperties(properties) {
        return scriptProps.setProperties(properties);
      }
    };
  } catch (error) {
    console.error('ServiceFactory.getProperties: Properties access error:', error.message);
    return null;
  }
}

// ===========================================
// 💾 Cache Management
// ===========================================

function getCache() {
  try {
    const cache = CacheService.getScriptCache();

    return {
      get(key) {
        try {
          const value = cache.get(key);
          return value ? JSON.parse(value) : null;
        } catch (parseError) {
          console.warn('ServiceFactory.getCache.get: Parse error for key:', key);
          return null;
        }
      },

      put(key, value, expirationInSeconds = 3600) {
        try {
          return cache.put(key, JSON.stringify(value), expirationInSeconds);
        } catch (error) {
          console.warn('ServiceFactory.getCache.put: Cache put error:', error.message);
          return false;
        }
      },

      remove(key) {
        try {
          return cache.remove(key);
        } catch (error) {
          console.warn('ServiceFactory.getCache.remove: Cache remove error:', error.message);
          return false;
        }
      },

      removeAll(keys) {
        try {
          // GAS Cache.removeAll requires an array of keys; no API for clearing entire cache.
          if (Array.isArray(keys) && keys.length > 0) {
            cache.removeAll(keys);
            return true;
          }
          // No keys provided: perform a no-op to avoid API signature error.
          console.warn('ServiceFactory.getCache.removeAll: No keys provided; skipping clear');
          return false;
        } catch (error) {
          console.warn('ServiceFactory.getCache.removeAll: Cache clear error:', error.message);
          return false;
        }
      }
    };
  } catch (error) {
    console.error('ServiceFactory.getCache: Cache service error:', error.message);
    return null;
  }
}

// ===========================================
// 🗄️ Database Access
// ===========================================

function getDB() {
  try {
    const root = (typeof globalThis !== 'undefined') ? globalThis : (typeof global !== 'undefined' ? global : this);

    // Use Data.gs Zero-Dependency implementation
    if (root && root.Data) {
      return root.Data;
    }

    // Direct access fallback
    if (typeof globalThis !== 'undefined' && globalThis.Data) {
      try { root.Data = globalThis.Data; } catch (_) { void 0; }
      return globalThis.Data;
    }

    if (typeof global !== 'undefined' && global.Data) {
      try { root.Data = global.Data; } catch (_) { void 0; }
      return global.Data;
    }

    console.warn('ServiceFactory.getDB: Data class not available');
    return null;
  } catch (error) {
    console.error('ServiceFactory.getDB: Database access error:', error.message);
    return null;
  }
}

// ===========================================
// 📋 Spreadsheet Operations
// ===========================================

function getSpreadsheet() {
  return {
    openById(id) {
      const authKey = `spreadsheet_auth_${id}`;

      // 🛡️ RequestGate safe access (optional dependency)
      if (typeof RequestGate !== 'undefined' && !RequestGate.enter(authKey)) {
        console.warn('ServiceFactory.getSpreadsheet.openById: Authentication in progress, waiting...');
        Utilities.sleep(200);
      } else if (typeof RequestGate !== 'undefined') {
        RequestGate.exit(authKey);
      }

      try {
        // 🔧 CLAUDE.md準拠: サービスアカウント専用アクセス - セキュリティ強化
        const auth = typeof Auth !== 'undefined' ? Auth.serviceAccount() : null;

        // 🛡️ サービスアカウント認証の厳格な検証
        if (!auth) {
          throw new Error('ServiceFactory.getSpreadsheet: Auth service unavailable');
        }

        if (!auth.isValid || !auth.token) {
          throw new Error(`ServiceFactory.getSpreadsheet: Service account authentication failed - ${auth.error || 'Unknown error'}`);
        }

        // ✅ セキュアなサービスアカウント専用アクセス
        console.log('ServiceFactory.getSpreadsheet.openById: Using secure service account authentication');

        if (typeof Data !== 'undefined' && typeof Data.openSpreadsheetWithServiceAccount === 'function') {
          return Data.openSpreadsheetWithServiceAccount(id, auth.token);
        } else {
          throw new Error('ServiceFactory.getSpreadsheet: Data.openSpreadsheetWithServiceAccount not available');
        }

      } catch (error) {
        console.error('ServiceFactory.getSpreadsheet.openById: Error opening spreadsheet:', error.message);
        return null;
      }
    },

    create(name) {
      try {
        return SpreadsheetApp.create(name);
      } catch (error) {
        console.error('ServiceFactory.getSpreadsheet.create: Error creating spreadsheet:', error.message);
        return null;
      }
    }
  };
}

// ===========================================
// 🔧 Utility Functions
// ===========================================

function getUtils() {
  return {
    generateId() {
      return Utilities.getUuid();
    },

    initService(serviceName) {
      try {
        if (typeof ServiceFactory === 'undefined') {
          console.warn(`init${serviceName}: ServiceFactory not available`);
          return false;
        }
        console.log(`✅ ${serviceName} (Zero-Dependency) initialized successfully`);
        return true;
      } catch (error) {
        console.error(`init${serviceName} failed:`, error.message);
        return false;
      }
    },

    formatDate(date, format = 'yyyy-MM-dd HH:mm:ss') {
      try {
        return Utilities.formatDate(date, Session.getScriptTimeZone(), format);
      } catch (error) {
        console.warn('ServiceFactory.getUtils.formatDate: Format error:', error.message);
        return date.toISOString();
      }
    },

    getTimeZone() {
      try {
        return Session.getScriptTimeZone();
      } catch (error) {
        console.warn('ServiceFactory.getUtils.getTimeZone: Timezone error:', error.message);
        return 'UTC';
      }
    },

    getWebAppUrl() {
      try {
        return ScriptApp.getService().getUrl();
      } catch (error) {
        console.warn('ServiceFactory.getUtils.getWebAppUrl: WebApp URL error:', error.message);
        return '';
      }
    }
  };
}

// ===========================================
// 👤 User Service Accessor
// ===========================================

function getUserService() {
  try {
    const root = (typeof globalThis !== 'undefined') ? globalThis : this;
    if (root.UserService && typeof root.UserService.isSystemAdmin === 'function') {
      return root.UserService;
    }
    // Fallback: wrap global isSystemAdmin function if available
    if (typeof global.isSystemAdmin === 'function') {
      return { isSystemAdmin: global.isSystemAdmin };
    }
    // Safe default stub
    return {
      isSystemAdmin: () => false
    };
  } catch (error) {
    console.warn('ServiceFactory.getUserService: Access error:', error.message);
    return { isSystemAdmin: () => false };
  }
}

// ===========================================
// 📊 Data Service Accessor
// ===========================================

function getDataService() {
  try {
    const root = (typeof globalThis !== 'undefined') ? globalThis : this;
    if (root.DataService) return root.DataService;

    // Build a minimal shim from available globals (best-effort)
    const shim = {};
    if (typeof getUserSheetData === 'function') {
      shim.getUserSheetData = getUserSheetData;
    }
    if (typeof dsAddReaction === 'function') {
      shim.addReaction = dsAddReaction;
    }
    if (typeof dsToggleHighlight === 'function') {
      shim.toggleHighlight = dsToggleHighlight;
    }
    if (typeof connectToSheetInternal === 'function') {
      shim.connectToSheetInternal = connectToSheetInternal;
    }
    if (Object.keys(shim).length > 0) return shim;

    console.warn('ServiceFactory.getDataService: DataService not available');
    return null;
  } catch (error) {
    console.error('ServiceFactory.getDataService: Access error:', error.message);
    return null;
  }
}

// ===========================================
// ⚙️ Config Service Accessor
// ===========================================

function getConfigService() {
  try {
    const root = (typeof globalThis !== 'undefined') ? globalThis : this;
    if (root.ConfigService) return root.ConfigService;

    // Build minimal shim from available globals
    const shim = {};
    if (typeof getUserConfig === 'function') shim.getUserConfig = getUserConfig;
    if (typeof getFormInfo === 'function') shim.getFormInfo = getFormInfo;
    if (Object.keys(shim).length > 0) return shim;

    console.warn('ServiceFactory.getConfigService: ConfigService not available');
    return null;
  } catch (error) {
    console.error('ServiceFactory.getConfigService: Access error:', error.message);
    return null;
  }
}
// ===========================================
// 🔍 Diagnostics
// ===========================================

function diagnose() {
  const results = {
    service: 'ServiceFactory',
    timestamp: new Date().toISOString(),
    checks: []
  };

  // Session check
  try {
    const session = getSession();
    results.checks.push({
      name: 'Session Service',
      status: session.isValid ? '✅' : '⚠️',
      details: session.isValid ? `User: ${session.email}` : 'No active session'
    });
  } catch (error) {
    results.checks.push({
      name: 'Session Service',
      status: '❌',
      details: error.message
    });
  }

  // Properties check
  try {
    const props = getProperties();
    results.checks.push({
      name: 'Properties Service',
      status: props ? '✅' : '❌',
      details: props ? 'Properties service accessible' : 'Properties service failed'
    });
  } catch (error) {
    results.checks.push({
      name: 'Properties Service',
      status: '❌',
      details: error.message
    });
  }

  // Cache check
  try {
    const cache = getCache();
    results.checks.push({
      name: 'Cache Service',
      status: cache ? '✅' : '❌',
      details: cache ? 'Cache service accessible' : 'Cache service failed'
    });
  } catch (error) {
    results.checks.push({
      name: 'Cache Service',
      status: '❌',
      details: error.message
    });
  }

  // Database check
  try {
    const db = getDB();
    results.checks.push({
      name: 'Database Service',
      status: db ? '✅' : '⚠️',
      details: db ? 'Database operations available' : 'Data class not found'
    });
  } catch (error) {
    results.checks.push({
      name: 'Database Service',
      status: '❌',
      details: error.message
    });
  }

  results.overall = results.checks.every(check => check.status === '✅') ? '✅' : '⚠️';
  return results;
}

// ===========================================
// 🌍 Global ServiceFactory Object
// ===========================================

/**
 * ServiceFactory統一インターフェース
 * Zero-Dependency Architecture統合層
 */
const __rootSF = (typeof globalThis !== 'undefined') ? globalThis : (typeof global !== 'undefined' ? global : this);
__rootSF.ServiceFactory = {
  getSession,
  getProperties,
  getCache,
  getDB,
  getSpreadsheet,
  getUtils,
  getUserService,
  getDataService,
  getConfigService,
  // Lazy accessor for SecurityService (optional usage sites)
  getSecurityService: (function () {
    return function getSecurityService() {
      try {
        const root = (typeof globalThis !== 'undefined') ? globalThis : this;
        if (root.SecurityService) return root.SecurityService;
        const shim = {};
        if (typeof validateUserData === 'function') shim.validateUserData = validateUserData;
        if (typeof validateSession === 'function') shim.validateSession = validateSession;
        if (Object.keys(shim).length > 0) return shim;
        console.warn('ServiceFactory.getSecurityService: SecurityService not available');
        return null;
      } catch (e) {
        console.error('ServiceFactory.getSecurityService: Access error:', e.message);
        return null;
      }
    };
  })(),
  diagnose
};
