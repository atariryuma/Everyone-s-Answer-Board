/**
 * Jest セットアップファイル - CLAUDE.md準拠版
 * 🚀 GAS V8ランタイム + configJSON中心型システム対応
 * 
 * 機能:
 * - GAS API の完全モック化
 * - 5フィールドデータベース環境
 * - configJSON中心設計対応
 * - SystemManager統合テスト環境
 */

// Jest Extended の拡張マッチャーをロード
require('jest-extended');

// グローバルタイマー関数の設定（clearTimeoutエラー対策）
global.clearTimeout = global.clearTimeout || (() => {});
global.setTimeout = global.setTimeout || (() => {});
global.clearInterval = global.clearInterval || (() => {});
global.setInterval = global.setInterval || (() => {});

// 🎯 CLAUDE.md準拠 - システム定数の設定
global.CORE = Object.freeze({
  TIMEOUTS: { SHORT: 1000, MEDIUM: 5000, LONG: 30000, FLOW: 300000 },
  STATUS: { ACTIVE: 'active', INACTIVE: 'inactive', PENDING: 'pending', ERROR: 'error' },
  HTTP_STATUS: { OK: 200, BAD_REQUEST: 400, UNAUTHORIZED: 401, FORBIDDEN: 403 }
});

global.PROPS_KEYS = Object.freeze({
  SERVICE_ACCOUNT_CREDS: 'SERVICE_ACCOUNT_CREDS',
  DATABASE_SPREADSHEET_ID: 'DATABASE_SPREADSHEET_ID',
  ADMIN_EMAIL: 'ADMIN_EMAIL'
});

// 🚀 configJSON中心型データベース定数（5フィールド構造）
global.DB_CONFIG = Object.freeze({
  SHEET_NAME: 'Users',
  HEADERS: Object.freeze([
    'userId',           // [0] UUID - 必須ID（検索用）
    'userEmail',        // [1] メールアドレス - 必須認証（検索用）
    'isActive',         // [2] アクティブ状態 - 必須フラグ（検索用）
    'configJson',       // [3] 全設定統合 - メインデータ（JSON）
    'lastModified',     // [4] 最終更新 - 監査用
  ])
});

global.SYSTEM_CONSTANTS = {
  DATABASE: global.DB_CONFIG,
  REACTIONS: Object.freeze({
    KEYS: ['UNDERSTAND', 'LIKE', 'CURIOUS'],
    LABELS: Object.freeze({
      UNDERSTAND: 'なるほど！',
      LIKE: 'いいね！', 
      CURIOUS: 'もっと知りたい！',
      HIGHLIGHT: 'ハイライト'
    })
  }),
  COLUMNS: Object.freeze({
    TIMESTAMP: 'タイムスタンプ',
    EMAIL: 'メールアドレス',
    CLASS: 'クラス',
    OPINION: '回答',
    REASON: '理由',
    NAME: '名前'
  })
};

// 🔧 GAS API モック設定
const createMockGASService = (serviceName) => ({
  [serviceName]: jest.fn(() => ({
    // 共通メソッドのモック
    getId: jest.fn(() => 'mock-id'),
    getName: jest.fn(() => 'mock-name'),
    getUrl: jest.fn(() => 'https://mock-url.com'),
    // 各サービス固有のメソッドは個別テストで設定
  }))
});

// SpreadsheetApp モック（configJSON対応）
global.SpreadsheetApp = {
  openById: jest.fn((id) => ({
    getId: () => id,
    getName: () => 'テストスプレッドシート',
    getSheetByName: jest.fn((name) => ({
      getName: () => name,
      getLastRow: jest.fn(() => 10),
      getLastColumn: jest.fn(() => 5),
      getRange: jest.fn(() => ({
        getValues: jest.fn(() => [
          ['userId', 'userEmail', 'isActive', 'configJson', 'lastModified'],
          ['test-user-1', 'test1@example.com', true, '{"spreadsheetId":"test-sheet-1"}', '2025-01-01T00:00:00Z']
        ]),
        setValues: jest.fn(),
        getValue: jest.fn(),
        setValue: jest.fn()
      })),
      appendRow: jest.fn(),
      insertRowAfter: jest.fn(),
      deleteRow: jest.fn(),
      getDataRange: jest.fn(() => ({
        getValues: jest.fn(() => [])
      }))
    })),
    getSheets: jest.fn(() => []),
    insertSheet: jest.fn()
  })),
  create: jest.fn(() => ({ getId: () => 'new-spreadsheet-id' })),
  getActiveSpreadsheet: jest.fn(),
  flush: jest.fn()
};

// PropertiesService モック（configJSON対応）
const mockProperties = new Map();
global.PropertiesService = {
  getScriptProperties: jest.fn(() => ({
    getProperty: jest.fn((key) => mockProperties.get(key) || null),
    setProperty: jest.fn((key, value) => mockProperties.set(key, value)),
    deleteProperty: jest.fn((key) => mockProperties.delete(key)),
    getProperties: jest.fn(() => Object.fromEntries(mockProperties))
  })),
  getUserProperties: jest.fn(() => ({
    getProperty: jest.fn(),
    setProperty: jest.fn(),
    deleteProperty: jest.fn()
  }))
};

// CacheService モック
const mockCache = new Map();
global.CacheService = {
  getScriptCache: jest.fn(() => ({
    get: jest.fn((key) => mockCache.get(key) || null),
    put: jest.fn((key, value, ttl) => mockCache.set(key, value)),
    remove: jest.fn((key) => mockCache.delete(key)),
    removeAll: jest.fn(() => mockCache.clear())
  }))
};

// Session モック
global.Session = {
  getActiveUser: jest.fn(() => ({
    getEmail: jest.fn(() => 'test@example.com'),
    getName: jest.fn(() => 'テストユーザー')
  })),
  getEffectiveUser: jest.fn(() => ({
    getEmail: jest.fn(() => 'test@example.com')
  })),
  getTemporaryActiveUserKey: jest.fn(() => 'temp-key')
};

// Utilities モック
global.Utilities = {
  getUuid: jest.fn(() => 'test-uuid-' + Date.now()),
  sleep: jest.fn(),
  formatDate: jest.fn((date) => date.toISOString()),
  jsonStringify: jest.fn((obj) => JSON.stringify(obj)),
  jsonParse: jest.fn((str) => JSON.parse(str)),
  base64Encode: jest.fn((input) => Buffer.from(input).toString('base64')),
  base64Decode: jest.fn((input) => Buffer.from(input, 'base64').toString()),
  computeHmacSha256Signature: jest.fn(() => 'mock-signature')
};

// UrlFetchApp モック
global.UrlFetchApp = {
  fetch: jest.fn((url, options = {}) => ({
    getResponseCode: jest.fn(() => 200),
    getContentText: jest.fn(() => JSON.stringify({ success: true })),
    getBlob: jest.fn(() => ({ getBytes: () => new Uint8Array() })),
    getHeaders: jest.fn(() => ({}))
  })),
  fetchAll: jest.fn(() => [])
};

// HtmlService モック
global.HtmlService = {
  createTemplateFromFile: jest.fn((filename) => ({
    evaluate: jest.fn(() => ({
      setXFrameOptionsMode: jest.fn(),
      getContent: jest.fn(() => '<html><body>Mock HTML</body></html>')
    })),
    getRawContent: jest.fn(() => 'Mock template content')
  })),
  createHtmlOutput: jest.fn((content) => ({
    setXFrameOptionsMode: jest.fn(),
    getContent: jest.fn(() => content || '<html><body>Mock HTML</body></html>')
  })),
  XFrameOptionsMode: {
    ALLOWALL: 'ALLOWALL',
    SAMEORIGIN: 'SAMEORIGIN',
    DENY: 'DENY'
  }
};

// ScriptApp モック
global.ScriptApp = {
  getProjectId: jest.fn(() => 'test-project-id'),
  newTrigger: jest.fn(() => ({
    timeBased: jest.fn(() => ({
      everyMinutes: jest.fn(() => ({ create: jest.fn() })),
      everyHours: jest.fn(() => ({ create: jest.fn() })),
      everyDays: jest.fn(() => ({ create: jest.fn() }))
    }))
  }))
};

// ContentService モック
global.ContentService = {
  createTextOutput: jest.fn((content) => ({
    setMimeType: jest.fn(),
    getContent: jest.fn(() => content)
  })),
  MimeType: {
    JSON: 'application/json',
    TEXT: 'text/plain',
    HTML: 'text/html'
  }
};

// Logger モック
global.Logger = {
  log: jest.fn()
};

// LockService モック
global.LockService = {
  getScriptLock: jest.fn(() => ({
    tryLock: jest.fn(() => true),
    waitLock: jest.fn(),
    releaseLock: jest.fn()
  }))
};

// 🆕 SystemManager モック（統合テスト対応）
global.SystemManager = {
  checkSetupStatus: jest.fn(() => ({
    isComplete: true,
    hasDatabase: true,
    hasServiceAccount: true,
    hasAdminEmail: true
  })),
  migrateToSimpleSchema: jest.fn(() => ({ success: true, migrated: 0 })),
  testSchemaOptimization: jest.fn(() => ({ 
    success: true, 
    performance: { improvement: '60%' } 
  }))
};

// 🔧 その他のプロジェクト固有モック
global.DB = {
  createUser: jest.fn((userData) => ({ 
    success: true, 
    userId: userData.userId || 'test-user-id' 
  })),
  findUserById: jest.fn((userId) => ({
    userId,
    userEmail: 'test@example.com',
    isActive: true,
    configJson: '{"spreadsheetId":"test-sheet","sheetName":"回答データ"}',
    lastModified: '2025-01-01T00:00:00Z'
  })),
  findUserByEmail: jest.fn((email) => ({
    userId: 'test-user-id',
    userEmail: email,
    isActive: true,
    configJson: '{"spreadsheetId":"test-sheet","sheetName":"回答データ"}',
    lastModified: '2025-01-01T00:00:00Z'
  })),
  updateUser: jest.fn(() => ({ success: true })),
  getAllUsers: jest.fn(() => [])
};

global.App = {
  getAccess: jest.fn(() => ({
    verifyAccess: jest.fn(() => ({ 
      hasAccess: true, 
      accessLevel: 'authenticated_user' 
    }))
  })),
  getConfig: jest.fn(() => ({
    getUserConfig: jest.fn(() => ({
      spreadsheetId: 'test-sheet',
      sheetName: '回答データ',
      setupStatus: 'completed'
    })),
    updateUserConfig: jest.fn(() => ({ success: true }))
  }))
};

global.SecurityValidator = {
  validateUserData: jest.fn((data) => ({
    isValid: true,
    errors: [],
    sanitizedData: data
  })),
  isValidUUID: jest.fn(() => true),
  isValidEmail: jest.fn(() => true)
};

// 🔧 グローバルヘルパー関数
global.console = {
  ...console,
  // GAS互換のログ関数
  info: jest.fn(),
  warn: jest.fn(),
  error: jest.fn()
};

// ⚡ configJSON中心型テストユーティリティ
global.createMockUserData = (overrides = {}) => ({
  userId: 'test-user-' + Date.now(),
  userEmail: 'test@example.com',
  isActive: true,
  configJson: JSON.stringify({
    spreadsheetId: 'test-spreadsheet-id',
    sheetName: '回答データ',
    createdAt: '2025-01-01T00:00:00Z',
    lastAccessedAt: '2025-01-01T12:00:00Z',
    setupStatus: 'completed',
    appPublished: true,
    ...overrides.configData
  }),
  lastModified: new Date().toISOString(),
  ...overrides
});

global.createMockConfig = (overrides = {}) => ({
  spreadsheetId: 'test-spreadsheet-id',
  sheetName: '回答データ',
  formUrl: 'https://forms.gle/test',
  setupStatus: 'completed',
  appPublished: true,
  createdAt: '2025-01-01T00:00:00Z',
  lastAccessedAt: '2025-01-01T12:00:00Z',
  ...overrides
});

// 🛡️ テスト前後のクリーンアップ
beforeEach(() => {
  // モックの状態をリセット
  jest.clearAllMocks();
  mockProperties.clear();
  mockCache.clear();
});

afterEach(() => {
  // 追加のクリーンアップが必要な場合はここに記述
});

console.info('🚀 Jest セットアップ完了 - CLAUDE.md準拠 configJSON中心型テスト環境');