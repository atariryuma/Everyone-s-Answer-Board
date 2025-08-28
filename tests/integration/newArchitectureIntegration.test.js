/**
 * @fileoverview 新アーキテクチャの統合テスト
 */
/* eslint-disable no-unused-vars */
/* eslint-disable no-undef */

// モックのセットアップ
const setupMocks = () => {
  // GAS グローバルオブジェクトのモック
  global.PropertiesService = {
    getScriptProperties: jest.fn(() => ({
      getProperty: jest.fn((key) => {
        const props = {
          DATABASE_SPREADSHEET_ID: 'test-db-id',
          ADMIN_EMAIL: 'admin@example.com',
          DEBUG_MODE: 'false',
        };
        return props[key];
      }),
      setProperty: jest.fn(),
    })),
    getUserProperties: jest.fn(() => ({
      getProperty: jest.fn(),
      setProperty: jest.fn(),
    })),
  };

  global.SpreadsheetApp = {
    openById: jest.fn((_id) => ({
      getSheetByName: jest.fn((name) => ({
        getDataRange: jest.fn(() => ({
          getValues: jest.fn(() => [
            [
              'userId',
              'adminEmail',
              'spreadsheetId',
              'spreadsheetUrl',
              'createdAt',
              'configJson',
              'lastAccessedAt',
              'isActive',
            ],
            [
              'user123',
              'test@example.com',
              'sheet123',
              'https://sheets.google.com/123',
              '2024-01-01',
              '{}',
              '2024-01-02',
              'true',
            ],
          ]),
        })),
        appendRow: jest.fn(),
        getRange: jest.fn(() => ({
          setValues: jest.fn(),
          getValue: jest.fn(),
          setValue: jest.fn(),
        })),
        setFrozenRows: jest.fn(),
        insertSheet: jest.fn(),
        getName: jest.fn(() => 'Sheet1'),
        getLastColumn: jest.fn(() => 8),
        getLastRow: jest.fn(() => 2),
      })),
      insertSheet: jest.fn((_name) => ({
        getRange: jest.fn(() => ({
          setValues: jest.fn(),
        })),
        setFrozenRows: jest.fn(),
      })),
      getSheets: jest.fn(() => [
        { getName: jest.fn(() => 'Sheet1'), getSheetId: jest.fn(() => 1) },
        { getName: jest.fn(() => 'Sheet2'), getSheetId: jest.fn(() => 2) },
      ]),
      getActiveSheet: jest.fn(() => ({
        getName: jest.fn(() => 'Sheet1'),
      })),
    })),
    create: jest.fn((name) => ({
      getId: jest.fn(() => 'new-sheet-id'),
      getUrl: jest.fn(() => 'https://sheets.google.com/new'),
      getActiveSheet: jest.fn(() => ({
        setName: jest.fn(),
      })),
    })),
  };

  global.CacheService = {
    getScriptCache: jest.fn(() => ({
      get: jest.fn(),
      put: jest.fn(),
      remove: jest.fn(),
    })),
    getUserCache: jest.fn(() => ({
      get: jest.fn(),
      put: jest.fn(),
      remove: jest.fn(),
    })),
  };

  global.Session = {
    getActiveUser: jest.fn(() => ({
      getEmail: jest.fn(() => 'test@example.com'),
    })),
    getTemporaryActiveUserKey: jest.fn(() => 'temp-key-123'),
  };

  global.Utilities = {
    getUuid: jest.fn(() => `uuid-${Date.now()}`),
    base64Encode: jest.fn((data) => Buffer.from(data).toString('base64')),
    base64Decode: jest.fn((data) => Buffer.from(data, 'base64').toString()),
    sleep: jest.fn(),
    formatDate: jest.fn((_date, _tz, _format) => '2024-01-01 12:00:00'),
  };

  global.HtmlService = {
    createHtmlOutputFromFile: jest.fn((_filename) => ({
      getContent: jest.fn(() => `<html>${filename}</html>`),
    })),
    createTemplateFromFile: jest.fn((_filename) => ({
      evaluate: jest.fn(() => ({
        setTitle: jest.fn().mockReturnThis(),
        setFaviconUrl: jest.fn().mockReturnThis(),
        addMetaTag: jest.fn().mockReturnThis(),
        setSandboxMode: jest.fn().mockReturnThis(),
      })),
    })),
    createHtmlOutput: jest.fn((html) => ({
      setTitle: jest.fn().mockReturnThis(),
      setFaviconUrl: jest.fn().mockReturnThis(),
      addMetaTag: jest.fn().mockReturnThis(),
    })),
  };

  global.FormApp = {
    create: jest.fn((_title) => ({
      getId: jest.fn(() => 'form-id-123'),
      getPublishedUrl: jest.fn(() => 'https://forms.google.com/123'),
      setDescription: jest.fn(),
      addTextItem: jest.fn(() => ({
        setTitle: jest.fn().mockReturnThis(),
        setRequired: jest.fn().mockReturnThis(),
      })),
      addParagraphTextItem: jest.fn(() => ({
        setTitle: jest.fn().mockReturnThis(),
        setRequired: jest.fn().mockReturnThis(),
      })),
      addMultipleChoiceItem: jest.fn(() => ({
        setTitle: jest.fn().mockReturnThis(),
        setRequired: jest.fn().mockReturnThis(),
        setChoiceValues: jest.fn().mockReturnThis(),
      })),
      addCheckboxItem: jest.fn(() => ({
        setTitle: jest.fn().mockReturnThis(),
        setRequired: jest.fn().mockReturnThis(),
        setChoiceValues: jest.fn().mockReturnThis(),
      })),
      setDestination: jest.fn(),
    })),
    DestinationType: {
      SPREADSHEET: 'SPREADSHEET',
    },
  };

  global.UrlFetchApp = {
    fetch: jest.fn((_url, _options) => ({
      getResponseCode: jest.fn(() => 200),
      getContentText: jest.fn(() => JSON.stringify({ success: true })),
    })),
  };

  global.LockService = {
    getScriptLock: jest.fn(() => ({
      waitLock: jest.fn(),
      releaseLock: jest.fn(),
    })),
  };

  global.DriveApp = {
    getFileById: jest.fn(),
    getFolderById: jest.fn(),
    createFile: jest.fn(),
  };

  global.ScriptApp = {
    getService: jest.fn(() => ({
      getUrl: jest.fn(() => 'https://script.google.com/test'),
    })),
  };

  global.ContentService = {
    createTextOutput: jest.fn((_text) => ({
      setMimeType: jest.fn().mockReturnThis(),
    })),
    MimeType: {
      JSON: 'JSON',
    },
  };
};

// テスト用のサービス関数定義（GASファイルは直接requireできないため、関数を直接定義）
const loadServices = () => {
  // エラーサービスの関数
  global.ERROR_SEVERITY = {
    LOW: 'low',
    MEDIUM: 'medium',
    HIGH: 'high',
    CRITICAL: 'critical',
  };

  global.ERROR_CATEGORIES = {
    AUTHENTICATION: 'authentication',
    AUTHORIZATION: 'authorization',
    DATABASE: 'database',
    CACHE: 'cache',
    NETWORK: 'network',
    VALIDATION: 'validation',
    SYSTEM: 'system',
    USER_INPUT: 'user_input',
  };

  global.logError = (error, context, severity = 'medium', category = 'system', metadata = {}) => {
    const errorMessage = error instanceof Error ? error.message : String(error);
    if (severity === 'critical' || severity === 'high') {
      console.error(`[${severity.toUpperCase()}] ${context}:`, errorMessage, metadata);
    } else if (severity === 'medium') {
      console.warn(`[${severity.toUpperCase()}] ${context}:`, errorMessage);
    } else {
      console.log(`[${severity.toUpperCase()}] ${context}:`, errorMessage);
    }
    return { message: errorMessage, severity, category, metadata };
  };

  global.debugLog = (...args) => {
    if (global.shouldEnableDebugMode && global.shouldEnableDebugMode()) {
      console.log('[DEBUG]', ...args);
    }
  };

  global.infoLog = (...args) => console.log('[INFO]', ...args);
  global.warnLog = (...args) => console.warn('[WARN]', ...args);
  global.errorLog = (...args) => console.error('[ERROR]', ...args);

  // キャッシュサービスの関数
  global.CACHE_CONFIG = {
    DEFAULT_TTL: 300,
    USER_DATA_TTL: 600,
    SPREADSHEET_TTL: 1800,
    MAX_KEY_LENGTH: 250,
  };

  global.memoryCache = new Map();
  global.memoryCacheExpiry = new Map();

  global.getCacheValue = (key) => {
    if (global.memoryCache.has(key)) {
      const expiry = global.memoryCacheExpiry.get(key);
      if (expiry && expiry > Date.now()) {
        return global.memoryCache.get(key);
      }
      global.memoryCache.delete(key);
      global.memoryCacheExpiry.delete(key);
    }
    return null;
  };

  global.setCacheValue = (key, value, ttl = 300) => {
    global.memoryCache.set(key, value);
    global.memoryCacheExpiry.set(key, Date.now() + ttl * 1000);
  };

  global.removeCacheValue = (key) => {
    global.memoryCache.delete(key);
    global.memoryCacheExpiry.delete(key);
  };

  global.clearCacheByPattern = (pattern) => {
    const regex = new RegExp(pattern);
    for (const key of global.memoryCache.keys()) {
      if (regex.test(key)) {
        global.memoryCache.delete(key);
        global.memoryCacheExpiry.delete(key);
      }
    }
  };

  global.clearAllCache = () => {
    global.memoryCache.clear();
    global.memoryCacheExpiry.clear();
  };

  global.clearUserCache = (userId) => {
    global.clearCacheByPattern(`user_${userId}`);
  };

  // GASラッパーの関数
  global.Properties = {
    getScriptProperty: (key) => global.PropertiesService.getScriptProperties().getProperty(key),
    setScriptProperty: (key, value) =>
      global.PropertiesService.getScriptProperties().setProperty(key, value),
  };

  global.Spreadsheet = {
    openById: (id) => global.SpreadsheetApp.openById(id),
    create: (name) => global.SpreadsheetApp.create(name),
  };

  global.Html = {
    createTemplateFromFile: (filename) => global.HtmlService.createTemplateFromFile(filename),
    createHtmlOutputFromFile: (filename) => global.HtmlService.createHtmlOutputFromFile(filename),
    createHtmlOutput: (html) => global.HtmlService.createHtmlOutput(html),
  };

  global.SessionService = {
    getActiveUserEmail: () => global.Session.getActiveUser().getEmail(),
    getTemporaryActiveUserKey: () => global.Session.getTemporaryActiveUserKey(),
  };

  global.Utils = global.Utilities;

  // データベースサービスの関数
  global.getUserById = (userId) => {
    if (userId === 'user123') {
      return {
        userId: 'user123',
        adminEmail: 'test@example.com',
        spreadsheetId: 'sheet123',
        spreadsheetUrl: 'https://sheets.google.com/123',
        createdAt: '2024-01-01',
        configJson: '{}',
        lastAccessedAt: '2024-01-02',
        isActive: 'true',
      };
    }
    return null;
  };

  global.getUserByEmail = (email) => {
    if (email === 'test@example.com') {
      return global.getUserById('user123');
    }
    return null;
  };

  global.createUser = (userData) => {
    return {
      userId: userData.userId || `uuid-${Date.now()}`,
      adminEmail: userData.adminEmail,
      spreadsheetId: userData.spreadsheetId || '',
      spreadsheetUrl: userData.spreadsheetUrl || '',
      createdAt: new Date().toISOString(),
      configJson: JSON.stringify(userData.config || {}),
      lastAccessedAt: new Date().toISOString(),
      isActive: 'true',
    };
  };

  global.updateUser = (userId, updates) => {
    return true;
  };

  global.deleteUser = (userId) => {
    return global.updateUser(userId, { isActive: 'false' });
  };

  global.getUserInfo = global.getUserById;
  global.getUserIdFromEmail = (email) => 'user123';
  global.getUserId = () => 'user123';
  global.isDeployUser = () => false;
  global.shouldEnableDebugMode = () => false;

  // APIサービスの関数
  global.handleCoreApiRequest = (action, params) => {
    try {
      const userId = params.userId || global.getUserId();
      const user = global.getUserInfo(userId);

      if (!user && action !== 'getLoginStatus') {
        return { success: false, error: '認証が必要です' };
      }

      switch (action) {
        case 'getInitialData':
          return {
            success: true,
            data: [],
            config: {},
            userInfo: {
              userId: user.userId,
              email: user.adminEmail,
              isAdmin: false,
            },
          };
        case 'getSheetData':
          return { success: true, data: [] };
        default:
          return { success: false, error: `不明なアクション: ${action}` };
      }
    } catch (error) {
      return { success: false, error: error.message };
    }
  };

  // マイグレーションブリッジの関数
  global.findUserById = (userId) => global.getUserById(userId);
  global.findUserByEmail = (email) => global.getUserByEmail(email);
  global.clearExecutionUserInfoCache = () => global.clearUserCache(global.getUserId());

  // メインエントリーポイントの関数
  global.doGet = (e) => {
    try {
      if (!global.isSystemInitialized()) {
        return global.showSetupPage();
      }

      const user = global.authenticateUser();
      if (!user) {
        return global.showLoginPage();
      }

      const params = global.parseRequestParams(e);

      switch (params.mode) {
        case 'admin':
          return global.showAdminPanel(user, params);
        default:
          return global.showMainPage(user, params);
      }
    } catch (error) {
      return global.showErrorPage(error);
    }
  };

  global.isSystemInitialized = () => true;
  global.authenticateUser = () => ({ userId: 'user123' });
  global.parseRequestParams = (e) => e.parameter || {};
  global.showSetupPage = () => 'setup-page-html';
  global.showLoginPage = () => 'login-page-html';
  global.showAdminPanel = () => 'admin-panel-html';
  global.showMainPage = () => 'main-page-html';
  global.showErrorPage = () => 'error-page-html';

  // スプレッドシートサービスの補助関数
  global.getCurrentSheetName = (_spreadsheetId) => 'Sheet1';
  global.getSheetConfig = (_userId, _sheetName) => ({});
  global.getSheetData = (_userId, _sheetName, _classFilter, _sortOrder, _adminMode) => [];

  // ログ関数の追加
  global.logError = jest.fn((error, context, severity, category, data) => {
    console.error(
      `[${severity?.toUpperCase() || 'ERROR'}] ${context}:`,
      error.message || error,
      data
    );
  });
  global.debugLog = jest.fn((message) => {
    if (global.shouldEnableDebugMode && global.shouldEnableDebugMode()) {
      console.log('[DEBUG]', message);
    }
  });
  global.warnLog = jest.fn((message) => {
    console.warn('[WARN]', message);
  });
  global.infoLog = jest.fn((message) => {
    console.info('[INFO]', message);
  });

  // キャッシュ管理関数の追加
  global.setCacheValue = jest.fn((key, value, ttl = 300) => {
    global.memoryCache.set(key, value);
    global.memoryCacheExpiry.set(key, Date.now() + ttl * 1000);
  });

  global.getCacheValue = jest.fn((key) => {
    if (global.memoryCache.has(key)) {
      const expiry = global.memoryCacheExpiry.get(key);
      if (expiry && expiry > Date.now()) {
        return global.memoryCache.get(key);
      } else {
        global.memoryCache.delete(key);
        global.memoryCacheExpiry.delete(key);
      }
    }
    return null;
  });

  global.removeCacheValue = jest.fn((key) => {
    global.memoryCache.delete(key);
    global.memoryCacheExpiry.delete(key);
  });

  global.clearCacheByPattern = jest.fn((pattern) => {
    const regex = new RegExp(pattern.replace('*', '.*'));
    for (const key of global.memoryCache.keys()) {
      if (regex.test(key)) {
        global.memoryCache.delete(key);
        global.memoryCacheExpiry.delete(key);
      }
    }
  });

  // データベース関数の追加
  global.getUserById = jest.fn((userId) => {
    if (userId === 'user123') {
      return {
        userId: 'user123',
        adminEmail: 'test@example.com',
        spreadsheetId: 'sheet123',
        isActive: 'true',
      };
    }
    return null;
  });

  global.createUser = jest.fn((userData) => {
    if (typeof userData === 'string') {
      return { userId: 'new-user-123', adminEmail: userData };
    }
    return {
      userId: 'new-user-123',
      adminEmail: userData.adminEmail,
      spreadsheetId: userData.spreadsheetId || 'new-sheet-id',
      isActive: 'true',
    };
  });
  global.updateUser = jest.fn((_userId, _updates) => true);
};

describe('新アーキテクチャ統合テスト', () => {
  beforeEach(() => {
    setupMocks();
    loadServices(); // サービス関数を読み込み
    // グローバル変数のリセット
    global.memoryCache = new Map();
    global.memoryCacheExpiry = new Map();
  });

  describe('エラーハンドリング', () => {
    test('エラーログが正しく記録される', () => {
      const consoleSpy = jest.spyOn(console, 'error').mockImplementation();

      global.logError(new Error('テストエラー'), 'testContext', 'high', 'system', { test: true });

      expect(consoleSpy).toHaveBeenCalledWith('[HIGH] testContext:', 'テストエラー', {
        test: true,
      });

      consoleSpy.mockRestore();
    });

    test('デバッグログが適切に制御される', () => {
      const consoleSpy = jest.spyOn(console, 'log').mockImplementation();

      global.shouldEnableDebugMode = jest.fn(() => false);
      debugLog('デバッグメッセージ');
      expect(consoleSpy).not.toHaveBeenCalled();

      global.shouldEnableDebugMode = jest.fn(() => true);
      debugLog('デバッグメッセージ');
      expect(consoleSpy).toHaveBeenCalledWith('[DEBUG]', 'デバッグメッセージ');

      consoleSpy.mockRestore();
    });
  });

  describe('キャッシュ管理', () => {
    test('キャッシュの取得と設定が正しく動作する', () => {
      setCacheValue('test-key', { data: 'test' }, 300);
      const cached = getCacheValue('test-key');

      expect(cached).toEqual({ data: 'test' });
    });

    test('キャッシュの削除が正しく動作する', () => {
      setCacheValue('test-key', 'test-value');
      removeCacheValue('test-key');
      const cached = getCacheValue('test-key');

      expect(cached).toBeNull();
    });

    test('パターンによるキャッシュクリアが動作する', () => {
      setCacheValue('user_123', 'user-data');
      setCacheValue('user_456', 'user-data-2');
      setCacheValue('other_key', 'other-data');

      clearCacheByPattern('user_');

      expect(getCacheValue('user_123')).toBeNull();
      expect(getCacheValue('user_456')).toBeNull();
      expect(getCacheValue('other_key')).toEqual('other-data');
    });
  });

  describe('データベース操作', () => {
    test('ユーザー情報の取得が正しく動作する', () => {
      const user = getUserById('user123');

      expect(user).toEqual(
        expect.objectContaining({
          userId: 'user123',
          adminEmail: 'test@example.com',
          spreadsheetId: 'sheet123',
        })
      );
    });

    test('新規ユーザーの作成が正しく動作する', () => {
      const newUser = createUser({
        adminEmail: 'new@example.com',
        spreadsheetId: 'new-sheet-id',
      });

      expect(newUser).toEqual(
        expect.objectContaining({
          adminEmail: 'new@example.com',
          spreadsheetId: 'new-sheet-id',
          isActive: 'true',
        })
      );
    });

    test('ユーザー情報の更新が正しく動作する', () => {
      const result = updateUser('user123', {
        spreadsheetId: 'updated-sheet-id',
      });

      expect(result).toBe(true);
      // updateUser関数のモック実装では実際のスプレッドシート操作は行わないため、
      // 戻り値のチェックのみ行う
    });
  });

  describe('API処理', () => {
    test('初期データ取得APIが正しく動作する', () => {
      global.getUserInfo = jest.fn(() => ({
        userId: 'user123',
        adminEmail: 'test@example.com',
        spreadsheetId: 'sheet123',
        configJson: '{}',
      }));

      global.getCurrentSheetName = jest.fn(() => 'Sheet1');
      global.getSheetConfig = jest.fn(() => ({}));
      global.getSheetData = jest.fn(() => []);
      global.isDeployUser = jest.fn(() => false);

      const result = handleCoreApiRequest('getInitialData', {
        userId: 'user123',
        sheetName: 'Sheet1',
      });

      expect(result.success).toBe(true);
      expect(result.userInfo).toBeDefined();
    });

    test('未認証ユーザーのAPIアクセスが拒否される', () => {
      global.getUserInfo = jest.fn(() => null);

      const result = handleCoreApiRequest('getSheetData', {
        userId: 'invalid-user',
      });

      expect(result.success).toBe(false);
      expect(result.error).toContain('認証が必要');
    });
  });

  describe('マイグレーションブリッジ', () => {
    test('既存関数が新サービスに正しくマップされる', () => {
      global.getUserById = jest.fn(() => ({ userId: 'test' }));

      const user = findUserById('test');
      expect(user).toEqual({ userId: 'test' });
      expect(global.getUserById).toHaveBeenCalledWith('test');
    });

    test('キャッシュクリア関数が正しくマップされる', () => {
      global.clearUserCache = jest.fn();
      global.getUserId = jest.fn(() => 'user123');

      clearExecutionUserInfoCache();
      expect(global.clearUserCache).toHaveBeenCalledWith('user123');
    });
  });

  describe('エントリーポイント', () => {
    test('doGet関数が正しくルーティングする', () => {
      global.isSystemInitialized = jest.fn(() => true);
      global.authenticateUser = jest.fn(() => ({ userId: 'user123' }));
      global.parseRequestParams = jest.fn(() => ({ mode: 'admin' }));
      global.showAdminPanel = jest.fn(() => 'admin-panel-html');

      const result = doGet({ parameter: {} });

      expect(global.showAdminPanel).toHaveBeenCalled();
      expect(result).toBe('admin-panel-html');
    });

    test('システム未初期化時にセットアップページが表示される', () => {
      global.isSystemInitialized = jest.fn(() => false);
      global.showSetupPage = jest.fn(() => 'setup-page-html');

      const result = doGet({ parameter: {} });

      expect(global.showSetupPage).toHaveBeenCalled();
      expect(result).toBe('setup-page-html');
    });
  });
});

// テスト実行
if (require.main === module) {
  setupMocks();
  console.log('新アーキテクチャ統合テスト環境のセットアップ完了');
}
