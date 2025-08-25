/**
 * @fileoverview 新アーキテクチャの統合テスト
 */

// モックのセットアップ
const setupMocks = () => {
  // GAS グローバルオブジェクトのモック
  global.PropertiesService = {
    getScriptProperties: jest.fn(() => ({
      getProperty: jest.fn((key) => {
        const props = {
          'DATABASE_SPREADSHEET_ID': 'test-db-id',
          'ADMIN_EMAIL': 'admin@example.com',
          'DEBUG_MODE': 'false'
        };
        return props[key];
      }),
      setProperty: jest.fn()
    })),
    getUserProperties: jest.fn(() => ({
      getProperty: jest.fn(),
      setProperty: jest.fn()
    }))
  };

  global.SpreadsheetApp = {
    openById: jest.fn((id) => ({
      getSheetByName: jest.fn((name) => ({
        getDataRange: jest.fn(() => ({
          getValues: jest.fn(() => [
            ['userId', 'adminEmail', 'spreadsheetId', 'spreadsheetUrl', 'createdAt', 'configJson', 'lastAccessedAt', 'isActive'],
            ['user123', 'test@example.com', 'sheet123', 'https://sheets.google.com/123', '2024-01-01', '{}', '2024-01-02', 'true']
          ])
        })),
        appendRow: jest.fn(),
        getRange: jest.fn(() => ({
          setValues: jest.fn(),
          getValue: jest.fn(),
          setValue: jest.fn()
        })),
        setFrozenRows: jest.fn(),
        insertSheet: jest.fn(),
        getName: jest.fn(() => name),
        getLastColumn: jest.fn(() => 8),
        getLastRow: jest.fn(() => 2)
      })),
      insertSheet: jest.fn((name) => ({
        getRange: jest.fn(() => ({
          setValues: jest.fn()
        })),
        setFrozenRows: jest.fn()
      })),
      getSheets: jest.fn(() => [
        { getName: jest.fn(() => 'Sheet1'), getSheetId: jest.fn(() => 1) },
        { getName: jest.fn(() => 'Sheet2'), getSheetId: jest.fn(() => 2) }
      ]),
      getActiveSheet: jest.fn(() => ({
        getName: jest.fn(() => 'Sheet1')
      }))
    })),
    create: jest.fn((name) => ({
      getId: jest.fn(() => 'new-sheet-id'),
      getUrl: jest.fn(() => 'https://sheets.google.com/new'),
      getActiveSheet: jest.fn(() => ({
        setName: jest.fn()
      }))
    }))
  };

  global.CacheService = {
    getScriptCache: jest.fn(() => ({
      get: jest.fn(),
      put: jest.fn(),
      remove: jest.fn()
    })),
    getUserCache: jest.fn(() => ({
      get: jest.fn(),
      put: jest.fn(),
      remove: jest.fn()
    }))
  };

  global.Session = {
    getActiveUser: jest.fn(() => ({
      getEmail: jest.fn(() => 'test@example.com')
    })),
    getTemporaryActiveUserKey: jest.fn(() => 'temp-key-123')
  };

  global.Utilities = {
    getUuid: jest.fn(() => 'uuid-' + Date.now()),
    base64Encode: jest.fn((data) => Buffer.from(data).toString('base64')),
    base64Decode: jest.fn((data) => Buffer.from(data, 'base64').toString()),
    sleep: jest.fn(),
    formatDate: jest.fn((date, tz, format) => '2024-01-01 12:00:00')
  };

  global.HtmlService = {
    createHtmlOutputFromFile: jest.fn((filename) => ({
      getContent: jest.fn(() => `<html>${filename}</html>`)
    })),
    createTemplateFromFile: jest.fn((filename) => ({
      evaluate: jest.fn(() => ({
        setTitle: jest.fn().mockReturnThis(),
        setFaviconUrl: jest.fn().mockReturnThis(),
        addMetaTag: jest.fn().mockReturnThis(),
        setSandboxMode: jest.fn().mockReturnThis()
      }))
    })),
    createHtmlOutput: jest.fn((html) => ({
      setTitle: jest.fn().mockReturnThis(),
      setFaviconUrl: jest.fn().mockReturnThis(),
      addMetaTag: jest.fn().mockReturnThis()
    }))
  };

  global.FormApp = {
    create: jest.fn((title) => ({
      getId: jest.fn(() => 'form-id-123'),
      getPublishedUrl: jest.fn(() => 'https://forms.google.com/123'),
      setDescription: jest.fn(),
      addTextItem: jest.fn(() => ({
        setTitle: jest.fn().mockReturnThis(),
        setRequired: jest.fn().mockReturnThis()
      })),
      addParagraphTextItem: jest.fn(() => ({
        setTitle: jest.fn().mockReturnThis(),
        setRequired: jest.fn().mockReturnThis()
      })),
      addMultipleChoiceItem: jest.fn(() => ({
        setTitle: jest.fn().mockReturnThis(),
        setRequired: jest.fn().mockReturnThis(),
        setChoiceValues: jest.fn().mockReturnThis()
      })),
      addCheckboxItem: jest.fn(() => ({
        setTitle: jest.fn().mockReturnThis(),
        setRequired: jest.fn().mockReturnThis(),
        setChoiceValues: jest.fn().mockReturnThis()
      })),
      setDestination: jest.fn()
    })),
    DestinationType: {
      SPREADSHEET: 'SPREADSHEET'
    }
  };

  global.UrlFetchApp = {
    fetch: jest.fn((url, options) => ({
      getResponseCode: jest.fn(() => 200),
      getContentText: jest.fn(() => JSON.stringify({ success: true }))
    }))
  };

  global.LockService = {
    getScriptLock: jest.fn(() => ({
      waitLock: jest.fn(),
      releaseLock: jest.fn()
    }))
  };

  global.DriveApp = {
    getFileById: jest.fn(),
    getFolderById: jest.fn(),
    createFile: jest.fn()
  };

  global.ScriptApp = {
    getService: jest.fn(() => ({
      getUrl: jest.fn(() => 'https://script.google.com/test')
    }))
  };

  global.ContentService = {
    createTextOutput: jest.fn((text) => ({
      setMimeType: jest.fn().mockReturnThis()
    })),
    MimeType: {
      JSON: 'JSON'
    }
  };
};

// テスト用のサービス読み込み（実際のファイルパスに合わせて調整が必要）
const loadServices = () => {
  // エラーサービス
  require('../../src/services/ErrorService.gs');
  // キャッシュサービス
  require('../../src/services/CacheService.gs');
  // GASラッパー
  require('../../src/services/GASWrapper.gs');
  // データベースサービス
  require('../../src/services/DatabaseService.gs');
  // スプレッドシートサービス
  require('../../src/services/SpreadsheetService.gs');
  // APIサービス
  require('../../src/services/APIService.gs');
  // メインエントリーポイント
  require('../../src/main-new.gs');
  // マイグレーションブリッジ
  require('../../src/migration/bridge.gs');
};

describe('新アーキテクチャ統合テスト', () => {
  beforeEach(() => {
    setupMocks();
    // グローバル変数のリセット
    global.memoryCache = new Map();
    global.memoryCacheExpiry = new Map();
  });

  describe('エラーハンドリング', () => {
    test('エラーログが正しく記録される', () => {
      const consoleSpy = jest.spyOn(console, 'error').mockImplementation();
      
      logError(new Error('テストエラー'), 'testContext', 'high', 'system', { test: true });
      
      expect(consoleSpy).toHaveBeenCalledWith(
        expect.stringContaining('[HIGH]'),
        expect.stringContaining('testContext'),
        expect.any(String),
        expect.any(Object)
      );
      
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
      
      expect(user).toEqual(expect.objectContaining({
        userId: 'user123',
        adminEmail: 'test@example.com',
        spreadsheetId: 'sheet123'
      }));
    });

    test('新規ユーザーの作成が正しく動作する', () => {
      const newUser = createUser({
        adminEmail: 'new@example.com',
        spreadsheetId: 'new-sheet-id'
      });
      
      expect(newUser).toEqual(expect.objectContaining({
        adminEmail: 'new@example.com',
        spreadsheetId: 'new-sheet-id',
        isActive: 'true'
      }));
    });

    test('ユーザー情報の更新が正しく動作する', () => {
      const result = updateUser('user123', {
        spreadsheetId: 'updated-sheet-id'
      });
      
      expect(result).toBe(true);
      expect(SpreadsheetApp.openById().getSheetByName().getRange().setValues).toHaveBeenCalled();
    });
  });

  describe('API処理', () => {
    test('初期データ取得APIが正しく動作する', () => {
      global.getUserInfo = jest.fn(() => ({
        userId: 'user123',
        adminEmail: 'test@example.com',
        spreadsheetId: 'sheet123',
        configJson: '{}'
      }));
      
      global.getCurrentSheetName = jest.fn(() => 'Sheet1');
      global.getSheetConfig = jest.fn(() => ({}));
      global.getSheetData = jest.fn(() => []);
      global.isDeployUser = jest.fn(() => false);
      
      const result = handleCoreApiRequest('getInitialData', {
        userId: 'user123',
        sheetName: 'Sheet1'
      });
      
      expect(result.success).toBe(true);
      expect(result.userInfo).toBeDefined();
    });

    test('未認証ユーザーのAPIアクセスが拒否される', () => {
      global.getUserInfo = jest.fn(() => null);
      
      const result = handleCoreApiRequest('getSheetData', {
        userId: 'invalid-user'
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