/**
 * Google Apps Script API モック
 * JestでGAS APIをテストするためのモック実装
 */

// SpreadsheetApp モック
global.SpreadsheetApp = {
  getActiveSpreadsheet: jest.fn(() => ({
    getSheetByName: jest.fn((_name) => ({
      getRange: jest.fn((row, col, numRows, numCols) => ({
        getValues: jest.fn(() => Array(numRows || 1).fill(Array(numCols || 1).fill(''))),
        setValues: jest.fn(),
        getValue: jest.fn(() => ''),
        setValue: jest.fn(),
        clear: jest.fn(),
      })),
      getLastRow: jest.fn(() => 10),
      getLastColumn: jest.fn(() => 5),
      appendRow: jest.fn(),
      deleteRow: jest.fn(),
      insertRowAfter: jest.fn(),
      getName: jest.fn(() => 'Sheet1'),
    })),
    getSheets: jest.fn(() => []),
    insertSheet: jest.fn(),
    deleteSheet: jest.fn(),
  })),
  getActiveSheet: jest.fn(() => ({
    getRange: jest.fn(() => ({
      getValues: jest.fn(() => [[]]),
      setValues: jest.fn(),
    })),
    getLastRow: jest.fn(() => 10),
    getLastColumn: jest.fn(() => 5),
  })),
  create: jest.fn((_name) => ({
    getName: jest.fn(() => _name),
  })),
  openById: jest.fn((_id) => ({
    getId: jest.fn(() => _id),
  })),
};

// 統一ログシステム モック
global.ULog = {
  // ログレベル定義（実際のULogと同じ）
  LEVELS: {
    ERROR: 'ERROR',
    WARN: 'WARN', 
    INFO: 'INFO',
    DEBUG: 'DEBUG'
  },

  // ログレベル優先度
  LEVEL_PRIORITY: {
    DEBUG: 1,
    INFO: 2,
    WARN: 3,
    ERROR: 4
  },

  // 現在のログレベル設定
  currentLogLevel: 'INFO',

  // カテゴリ定義
  CATEGORIES: {
    SYSTEM: 'SYSTEM',
    AUTH: 'AUTH',
    API: 'API',
    DATABASE: 'DATABASE',
    UI: 'UI',
    CACHE: 'CACHE',
    SECURITY: 'SECURITY',
    BATCH: 'BATCH',
    WORKFLOW: 'WORKFLOW'
  },

  // ログレベル設定
  setLogLevel: jest.fn(function(level) {
    if (this.LEVELS[level]) {
      this.currentLogLevel = level;
    }
  }),

  // ログレベルチェック
  shouldLog: jest.fn(function(level) {
    const currentPriority = this.LEVEL_PRIORITY[this.currentLogLevel] || 2;
    const checkPriority = this.LEVEL_PRIORITY[level] || 2;
    return checkPriority >= currentPriority;
  }),

  // 関数名取得（モック）
  _getFunctionName: jest.fn(() => 'testFunction'),

  // コアログ機能
  _logCore: jest.fn(),

  // ログメソッド
  debug: jest.fn(),
  info: jest.fn(),
  warn: jest.fn(),
  error: jest.fn()
};

// 統一レスポンス関数 モック
global.createSuccessResponse = jest.fn((data = null, message = null) => ({
  success: true,
  data: data,
  message: message,
  error: null
}));

global.createErrorResponse = jest.fn((error, message = null, data = null) => ({
  success: false,
  data: data,
  message: message,
  error: error
}));

global.createUnifiedResponse = jest.fn((success, data = null, message = null, error = null) => ({
  success: success,
  data: data,
  message: message,
  error: error
}));

// 統一キャッシュマネージャー モック
global.cacheManager = {
  get: jest.fn(),
  put: jest.fn(),
  remove: jest.fn(),
  clear: jest.fn(),
  getStats: jest.fn(() => ({ hits: 0, misses: 0 })),
};

// その他の統一システム関数 モック
global.findUserById = jest.fn();
global.findUserByEmail = jest.fn();
global.getOrFetchUserInfo = jest.fn();
global.getCurrentUserEmail = jest.fn(() => 'test@example.com');
global.coreGetCurrentUserEmail = jest.fn(() => 'test@example.com');

// 統合システム定数 モック
global.UNIFIED_CONSTANTS = {
  ERROR: {
    SEVERITY: {
      LOW: 'low',
      MEDIUM: 'medium', 
      HIGH: 'high',
      CRITICAL: 'critical'
    },
    CATEGORIES: {
      AUTHENTICATION: 'authentication',
      AUTHORIZATION: 'authorization',
      DATABASE: 'database',
      CACHE: 'cache',
      NETWORK: 'network',
      VALIDATION: 'validation',
      SYSTEM: 'system',
      USER_INPUT: 'user_input',
      EXTERNAL: 'external',
      CONFIG: 'config',
      SECURITY: 'security'
    }
  },
  CACHE: {
    TTL: {
      SHORT: 300,
      MEDIUM: 600,
      LONG: 3600,
      EXTENDED: 21600
    },
    BATCH_SIZE: {
      SMALL: 20,
      MEDIUM: 50,
      LARGE: 100,
      XLARGE: 200
    }
  },
  TIMEOUTS: {
    SHORT: 5000,
    MEDIUM: 15000,
    LONG: 30000,
    EXTENDED: 60000
  }
};

global.getSheetsServiceCached = jest.fn();
global.batchGetSheetsData = jest.fn();
global.openSpreadsheetOptimized = jest.fn();
global.isPublishedBoard = jest.fn(() => true);
global.formatSheetDataForFrontend = jest.fn();
global.getSpreadsheetsData = jest.fn();
global.clearExecutionUserInfoCache = jest.fn();
global.invalidateUserCache = jest.fn();
global.getUnifiedExecutionCache = jest.fn(() => ({
  clearUserInfo: jest.fn(),
  clearSheetsService: jest.fn(),
  clearAll: jest.fn()
}));

// セキュリティ・権限関数 モック
global.isDeployUser = jest.fn(() => true);
global.executeWithStandardizedLock = jest.fn((callback) => callback());
global.getSecureDatabaseId = jest.fn(() => 'mock-database-id');
global.getServiceAccountTokenCached = jest.fn(() => 'mock-token');
global.generateNewServiceAccountToken = jest.fn(() => 'new-mock-token');
global.getServiceAccountEmail = jest.fn(() => 'service@example.com');

// データベース操作関数 モック
global.DB_SHEET_CONFIG = {
  sheetName: 'UserDatabase',
  headers: ['userId', 'adminEmail', 'createdAt']
};
global.DELETE_LOG_SHEET_CONFIG = {
  sheetName: 'DeleteLog',
  headers: ['userId', 'deletedAt', 'reason']
};
global.DIAGNOSTIC_LOG_SHEET_CONFIG = {
  sheetName: 'DiagnosticLog',
  headers: ['timestamp', 'level', 'message']
};

// スクリプトプロパティキー モック
global.SCRIPT_PROPS_KEYS = {
  DATABASE_ID: 'DATABASE_ID',
  SERVICE_ACCOUNT_CREDS: 'SERVICE_ACCOUNT_CREDS',
  ADMIN_EMAIL: 'ADMIN_EMAIL'
};

// バッチプロセッサ モック
global.unifiedBatchProcessor = {
  batchUpdateSheet: jest.fn(),
  batchReadSheet: jest.fn()
};

// キャッシュ操作 モック
global.clearDatabaseCache = jest.fn();
global.synchronizeCacheAfterCriticalUpdate = jest.fn();

// ユーティリティ関数 モック
global.resilientUrlFetch = jest.fn(() => ({ getResponseCode: () => 200 }));
global.shareSpreadsheetWithServiceAccount = jest.fn();
global.createUserFolder = jest.fn(() => 'mock-folder-id');

// 正規表現 モック
global.EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

// unifiedCacheAPI モック
global.unifiedCacheAPI = {
  invalidateUserCache: jest.fn(),
  clearDatabaseCache: jest.fn(),
  synchronizeCacheAfterCriticalUpdate: jest.fn()
};

// 統一Sheet管理システム関数のモック
global.getPublishedSheetDataUnified = jest.fn();
global.getSheetDataUnified = jest.fn();
global.getIncrementalSheetDataUnified = jest.fn();
global.getSheetHeadersUnified = jest.fn();
global.getHeadersCachedUnified = jest.fn();
global.getHeaderIndicesUnified = jest.fn();
global.getCurrentSpreadsheetUnified = jest.fn();
global.getEffectiveSpreadsheetIdUnified = jest.fn();
global.getAvailableSheetsUnified = jest.fn();
global.getSheetDetailsUnified = jest.fn();

// PropertiesService モック
global.PropertiesService = {
  getScriptProperties: jest.fn(() => ({
    getProperty: jest.fn((_key) => null),
    setProperty: jest.fn(),
    deleteProperty: jest.fn(),
    getProperties: jest.fn(() => ({})),
    setProperties: jest.fn(),
    deleteAllProperties: jest.fn(),
  })),
  getUserProperties: jest.fn(() => ({
    getProperty: jest.fn((_key) => null),
    setProperty: jest.fn(),
    deleteProperty: jest.fn(),
    getProperties: jest.fn(() => ({})),
    setProperties: jest.fn(),
    deleteAllProperties: jest.fn(),
  })),
  getDocumentProperties: jest.fn(() => ({
    getProperty: jest.fn((_key) => null),
    setProperty: jest.fn(),
    deleteProperty: jest.fn(),
    getProperties: jest.fn(() => ({})),
    setProperties: jest.fn(),
    deleteAllProperties: jest.fn(),
  })),
};

// CacheService モック
global.CacheService = {
  getScriptCache: jest.fn(() => ({
    get: jest.fn((_key) => null),
    put: jest.fn(),
    remove: jest.fn(),
    getAll: jest.fn(() => ({})),
    putAll: jest.fn(),
    removeAll: jest.fn(),
  })),
  getUserCache: jest.fn(() => ({
    get: jest.fn((_key) => null),
    put: jest.fn(),
    remove: jest.fn(),
    getAll: jest.fn(() => ({})),
    putAll: jest.fn(),
    removeAll: jest.fn(),
  })),
  getDocumentCache: jest.fn(() => ({
    get: jest.fn((_key) => null),
    put: jest.fn(),
    remove: jest.fn(),
    getAll: jest.fn(() => ({})),
    putAll: jest.fn(),
    removeAll: jest.fn(),
  })),
};

// UrlFetchApp モック
global.UrlFetchApp = {
  fetch: jest.fn((_url, _params) => ({
    getResponseCode: jest.fn(() => 200),
    getContentText: jest.fn(() => JSON.stringify({ success: true })),
    getBlob: jest.fn(),
    getHeaders: jest.fn(() => ({})),
    getAllHeaders: jest.fn(() => ({})),
  })),
};

// HtmlService モック
global.HtmlService = {
  createTemplateFromFile: jest.fn((filename) => ({
    evaluate: jest.fn(() => ({
      getContent: jest.fn(() => `<html>${filename}</html>`),
      setTitle: jest.fn(),
      setSandboxMode: jest.fn(),
      setWidth: jest.fn(),
      setHeight: jest.fn(),
      addMetaTag: jest.fn(),
    })),
    filename,
  })),
  createHtmlOutput: jest.fn((html) => ({
    getContent: jest.fn(() => html),
    setTitle: jest.fn(),
    setSandboxMode: jest.fn(),
    setWidth: jest.fn(),
    setHeight: jest.fn(),
    addMetaTag: jest.fn(),
  })),
  SandboxMode: {
    IFRAME: 'IFRAME',
    NATIVE: 'NATIVE',
    EMULATED: 'EMULATED',
  },
};

// Utilities モック
global.Utilities = {
  formatDate: jest.fn((_date, _timeZone, _format) => 'formatted date'),
  formatString: jest.fn((template, ..._args) => template),
  getUuid: jest.fn(() => 'mock-uuid-1234'),
  base64Encode: jest.fn((data) => Buffer.from(data).toString('base64')),
  base64Decode: jest.fn((data) => Buffer.from(data, 'base64').toString()),
  newBlob: jest.fn((data) => ({ getBytes: () => data })),
  parseCsv: jest.fn((_csv) => [['col1', 'col2']]),
  sleep: jest.fn(),
  jsonParse: jest.fn((str) => JSON.parse(str)),
  jsonStringify: jest.fn((obj) => JSON.stringify(obj)),
};

// Logger モック
global.Logger = {
  log: jest.fn(console.log),
  clear: jest.fn(),
  getLog: jest.fn(() => ''),
};

// ScriptApp モック
global.ScriptApp = {
  newTrigger: jest.fn((_functionName) => ({
    timeBased: jest.fn(() => ({
      everyMinutes: jest.fn(() => ({ create: jest.fn() })),
      everyHours: jest.fn(() => ({ create: jest.fn() })),
      everyDays: jest.fn(() => ({ create: jest.fn() })),
      atHour: jest.fn(() => ({ create: jest.fn() })),
    })),
    forSpreadsheet: jest.fn(() => ({
      onEdit: jest.fn(() => ({ create: jest.fn() })),
      onChange: jest.fn(() => ({ create: jest.fn() })),
      onFormSubmit: jest.fn(() => ({ create: jest.fn() })),
    })),
  })),
  deleteTrigger: jest.fn(),
  getProjectTriggers: jest.fn(() => []),
  getAuthorizationInfo: jest.fn(() => ({
    getAuthorizationStatus: jest.fn(),
    getAuthorizationUrl: jest.fn(),
  })),
};

// DriveApp モック
global.DriveApp = {
  getFileById: jest.fn((_id) => ({
    getName: jest.fn(() => 'file.txt'),
    getBlob: jest.fn(),
    getId: jest.fn(() => _id),
    getUrl: jest.fn(() => `https://drive.google.com/file/d/${_id}`),
    setSharing: jest.fn(),
    makeCopy: jest.fn(),
  })),
  getFolderById: jest.fn((_id) => ({
    getName: jest.fn(() => 'folder'),
    createFile: jest.fn(),
    createFolder: jest.fn(),
    getFiles: jest.fn(() => ({
      hasNext: jest.fn(() => false),
      next: jest.fn(),
    })),
  })),
  createFile: jest.fn((_blob) => ({
    getId: jest.fn(() => 'new-file-id'),
  })),
  createFolder: jest.fn((name) => ({
    getId: jest.fn(() => 'new-folder-id'),
    getName: jest.fn(() => name),
  })),
};

// Session モック
global.Session = {
  getActiveUser: jest.fn(() => ({
    getEmail: jest.fn(() => 'user@example.com'),
  })),
  getEffectiveUser: jest.fn(() => ({
    getEmail: jest.fn(() => 'user@example.com'),
  })),
  getTemporaryActiveUserKey: jest.fn(() => 'temp-key'),
  getTimeZone: jest.fn(() => 'Asia/Tokyo'),
  getScriptTimeZone: jest.fn(() => 'Asia/Tokyo'),
};

// LockService モック
global.LockService = {
  getScriptLock: jest.fn(() => ({
    tryLock: jest.fn(() => true),
    releaseLock: jest.fn(),
    hasLock: jest.fn(() => true),
    waitLock: jest.fn(),
  })),
  getUserLock: jest.fn(() => ({
    tryLock: jest.fn(() => true),
    releaseLock: jest.fn(),
    hasLock: jest.fn(() => true),
    waitLock: jest.fn(),
  })),
  getDocumentLock: jest.fn(() => ({
    tryLock: jest.fn(() => true),
    releaseLock: jest.fn(),
    hasLock: jest.fn(() => true),
    waitLock: jest.fn(),
  })),
};

// console モック（GAS環境用）
if (!global.console) {
  global.console = {
    log: jest.fn(),
    error: jest.fn(),
    warn: jest.fn(),
    info: jest.fn(),
    debug: jest.fn(),
  };
}

module.exports = {
  SpreadsheetApp: global.SpreadsheetApp,
  PropertiesService: global.PropertiesService,
  CacheService: global.CacheService,
  UrlFetchApp: global.UrlFetchApp,
  HtmlService: global.HtmlService,
  Utilities: global.Utilities,
  Logger: global.Logger,
  ScriptApp: global.ScriptApp,
  DriveApp: global.DriveApp,
  Session: global.Session,
  LockService: global.LockService,
  console: global.console,
};
