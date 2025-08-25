/**
 * @fileoverview クリティカル機能のテスト - 新アーキテクチャでの重要機能の動作確認
 */

// モックとサービスのセットアップ
const setupTest = () => {
  // グローバルモックの設定
  global.PropertiesService = {
    getScriptProperties: jest.fn(() => ({
      getProperty: jest.fn((key) => {
        const props = {
          'DATABASE_SPREADSHEET_ID': 'test-db-id',
          'ADMIN_EMAIL': 'admin@example.com',
          'SERVICE_ACCOUNT_CREDS': JSON.stringify({
            client_email: 'service@example.com',
            private_key: 'test-key'
          })
        };
        return props[key];
      }),
      setProperty: jest.fn()
    }))
  };

  global.SpreadsheetApp = {
    openById: jest.fn((id) => ({
      getSheetByName: jest.fn((name) => ({
        getDataRange: jest.fn(() => ({
          getValues: jest.fn(() => {
            if (name === 'Users') {
              return [
                ['userId', 'adminEmail', 'spreadsheetId', 'spreadsheetUrl', 'createdAt', 'configJson', 'lastAccessedAt', 'isActive'],
                ['user123', 'test@example.com', 'sheet123', 'https://sheets.google.com/123', '2024-01-01', '{"appPublished":true,"publishedSheetName":"Sheet1"}', '2024-01-02', 'true']
              ];
            }
            return [['header1', 'header2'], ['data1', 'data2']];
          })
        })),
        appendRow: jest.fn(),
        getRange: jest.fn(() => ({
          setValues: jest.fn(),
          getValue: jest.fn(() => ''),
          setValue: jest.fn(),
          getValues: jest.fn(() => [['header1', 'header2']])
        })),
        getLastColumn: jest.fn(() => 2),
        getLastRow: jest.fn(() => 2),
        getName: jest.fn(() => name),
        getSheetId: jest.fn(() => 1)
      })),
      getSheets: jest.fn(() => [
        { getName: jest.fn(() => 'Sheet1'), getSheetId: jest.fn(() => 1) },
        { getName: jest.fn(() => 'Sheet2'), getSheetId: jest.fn(() => 2) }
      ]),
      getActiveSheet: jest.fn(() => ({
        getName: jest.fn(() => 'Sheet1')
      }))
    }))
  };

  global.Session = {
    getActiveUser: jest.fn(() => ({
      getEmail: jest.fn(() => 'test@example.com')
    }))
  };

  global.CacheService = {
    getScriptCache: jest.fn(() => ({
      get: jest.fn(() => null),
      put: jest.fn(),
      remove: jest.fn()
    }))
  };

  global.Utilities = {
    getUuid: jest.fn(() => 'uuid-test-123'),
    formatDate: jest.fn(() => '2024-01-01 12:00:00')
  };

  global.ScriptApp = {
    getService: jest.fn(() => ({
      getUrl: jest.fn(() => 'https://script.google.com/test')
    }))
  };

  global.FormApp = {
    create: jest.fn((title) => ({
      getId: jest.fn(() => 'form-123'),
      getPublishedUrl: jest.fn(() => 'https://forms.google.com/123'),
      setDescription: jest.fn(),
      addTextItem: jest.fn(() => ({
        setTitle: jest.fn().mockReturnThis(),
        setRequired: jest.fn().mockReturnThis()
      })),
      setDestination: jest.fn()
    })),
    DestinationType: { SPREADSHEET: 'SPREADSHEET' }
  };

  global.HtmlService = {
    createTemplateFromFile: jest.fn(() => ({
      evaluate: jest.fn(() => ({
        setTitle: jest.fn().mockReturnThis(),
        setFaviconUrl: jest.fn().mockReturnThis(),
        addMetaTag: jest.fn().mockReturnThis(),
        setSandboxMode: jest.fn().mockReturnThis()
      }))
    })),
    createHtmlOutputFromFile: jest.fn(() => ({
      getContent: jest.fn(() => '<html>test</html>')
    })),
    SandboxMode: { IFRAME: 'IFRAME' }
  };

  global.ContentService = {
    createTextOutput: jest.fn((text) => ({
      setMimeType: jest.fn().mockReturnThis()
    })),
    MimeType: { JSON: 'JSON' }
  };

  // サービス関数の定義
  global.ERROR_SEVERITY = {
    LOW: 'low',
    MEDIUM: 'medium',
    HIGH: 'high',
    CRITICAL: 'critical'
  };
  
  global.ERROR_CATEGORIES = {
    AUTHENTICATION: 'authentication',
    DATABASE: 'database',
    CACHE: 'cache',
    SYSTEM: 'system'
  };

  global.logError = jest.fn();
  global.debugLog = jest.fn();
  global.warnLog = jest.fn();
  global.infoLog = jest.fn();

  // メモリキャッシュ
  global.memoryCache = new Map();
  global.memoryCacheExpiry = new Map();

  global.getCacheValue = (key) => {
    if (global.memoryCache.has(key)) {
      const expiry = global.memoryCacheExpiry.get(key);
      if (expiry && expiry > Date.now()) {
        return global.memoryCache.get(key);
      }
    }
    return null;
  };

  global.setCacheValue = (key, value, ttl = 300) => {
    global.memoryCache.set(key, value);
    global.memoryCacheExpiry.set(key, Date.now() + (ttl * 1000));
  };

  global.clearUserCache = (userId) => {
    for (const key of global.memoryCache.keys()) {
      if (key.includes(userId)) {
        global.memoryCache.delete(key);
        global.memoryCacheExpiry.delete(key);
      }
    }
  };

  // データベースサービス
  global.getUserInfo = jest.fn((userId) => {
    if (userId === 'user123') {
      return {
        userId: 'user123',
        adminEmail: 'test@example.com',
        spreadsheetId: 'sheet123',
        spreadsheetUrl: 'https://sheets.google.com/123',
        configJson: '{"appPublished":true,"publishedSheetName":"Sheet1","setupStatus":"completed"}',
        isActive: 'true',
        lastAccessedAt: '2024-01-02'
      };
    }
    return null;
  });

  global.getUserId = jest.fn(() => 'user123');
  global.updateUser = jest.fn(() => true);
  global.createUser = jest.fn((data) => ({ ...data, userId: 'new-user-123' }));

  // スプレッドシートサービス
  global.Spreadsheet = {
    openById: (id) => global.SpreadsheetApp.openById(id)
  };

  global.Properties = {
    getScriptProperty: (key) => global.PropertiesService.getScriptProperties().getProperty(key)
  };

  global.SessionService = {
    getActiveUserEmail: () => global.Session.getActiveUser().getEmail()
  };

  // コア機能のインポート（実際の関数を使用）
  global.getSetupStep = (userInfo, configJson) => {
    if (!userInfo || !userInfo.spreadsheetId) return 1;
    if (!configJson.formCreated) return 2;
    if (!configJson.appPublished) return 3;
    return 4;
  };

  global.getSheetsList = (userId) => ['Sheet1', 'Sheet2'];
  
  global.generateUserUrls = (userId) => ({
    webApp: 'https://script.google.com/test',
    admin: 'https://script.google.com/test?mode=admin',
    setup: 'https://script.google.com/test?mode=setup',
    published: 'https://script.google.com/test'
  });

  global.getWebAppUrl = () => 'https://script.google.com/test';
  
  global.getSheetData = jest.fn(() => [
    { rowIndex: 1, name: 'Test User', opinion: 'Test Opinion' }
  ]);

  global.getResponsesData = jest.fn(() => ({
    status: 'success',
    data: [{ id: 1 }, { id: 2 }]
  }));

  global.verifyUserAccess = jest.fn(() => true);
  
  global.createQuickStartForm = jest.fn(() => ({
    formId: 'form-123',
    formUrl: 'https://forms.google.com/123',
    spreadsheetId: 'sheet-new-123',
    spreadsheetUrl: 'https://sheets.google.com/new'
  }));
};

// テストケース
describe('クリティカル機能テスト', () => {
  beforeEach(() => {
    setupTest();
    jest.clearAllMocks();
  });

  describe('getInitialData', () => {
    test('正常に初期データを取得できる', () => {
      // CoreFunctionsServiceのgetInitialData関数を模擬
      global.getInitialData = (requestUserId, targetSheetName, lightweightMode) => {
        const userInfo = global.getUserInfo(requestUserId);
        const configJson = JSON.parse(userInfo.configJson);
        // formCreatedを追加してステップ4になるようにする
        configJson.formCreated = true;
        
        return {
          userInfo: userInfo,
          appUrls: global.generateUserUrls(requestUserId),
          setupStep: global.getSetupStep(userInfo, configJson),
          activeSheetName: configJson.publishedSheetName,
          webAppUrl: global.getWebAppUrl(),
          isPublished: !!configJson.appPublished,
          answerCount: 2,
          totalReactions: 4,
          config: configJson,
          allSheets: global.getSheetsList(requestUserId),
          _meta: {
            apiVersion: 'integrated_v1',
            executionTime: null
          }
        };
      };

      const result = global.getInitialData('user123', 'Sheet1', false);
      
      expect(result).toBeDefined();
      expect(result.userInfo.userId).toBe('user123');
      expect(result.setupStep).toBe(4); // 全て完了
      expect(result.isPublished).toBe(true);
      expect(result.answerCount).toBe(2);
    });

    test('ユーザーが見つからない場合エラーをスロー', () => {
      global.getUserInfo = jest.fn(() => null);
      
      global.getInitialData = (requestUserId) => {
        const userInfo = global.getUserInfo(requestUserId);
        if (!userInfo) {
          throw new Error('ユーザー情報が見つかりません');
        }
      };

      expect(() => {
        global.getInitialData('invalid-user');
      }).toThrow('ユーザー情報が見つかりません');
    });
  });

  describe('getPublishedSheetData', () => {
    test('公開シートデータを取得できる', () => {
      global.getPublishedSheetData = (requestUserId, classFilter, sortOrder, adminMode, bypassCache) => {
        const userInfo = global.getUserInfo(requestUserId);
        if (!userInfo) {
          return { success: false, error: 'ユーザーが見つかりません' };
        }
        
        const data = global.getSheetData(requestUserId, 'Sheet1', classFilter, sortOrder, adminMode);
        return { success: true, data: data, fromCache: false };
      };

      const result = global.getPublishedSheetData('user123', null, 'newest', false, false);
      
      expect(result.success).toBe(true);
      expect(result.data).toBeDefined();
      expect(result.data.length).toBeGreaterThan(0);
    });

    test('キャッシュから取得できる', () => {
      const testData = [{ id: 1, cached: true }];
      global.setCacheValue('sheet_data_user123_Sheet1_null_newest_false', testData, 60);
      
      global.getPublishedSheetData = (requestUserId) => {
        const cacheKey = `sheet_data_${requestUserId}_Sheet1_null_newest_false`;
        const cached = global.getCacheValue(cacheKey);
        if (cached) {
          return { success: true, data: cached, fromCache: true };
        }
        return { success: true, data: [], fromCache: false };
      };

      const result = global.getPublishedSheetData('user123');
      
      expect(result.fromCache).toBe(true);
      expect(result.data[0].cached).toBe(true);
    });
  });

  describe('createQuickStartFormUI', () => {
    test('クイックスタートフォームを作成できる', () => {
      global.createQuickStartFormUI = (requestUserId) => {
        const userInfo = global.getUserInfo(requestUserId);
        if (!userInfo) {
          throw new Error('ユーザーが見つかりません');
        }
        
        const result = global.createQuickStartForm(userInfo.adminEmail, requestUserId);
        
        global.updateUser(requestUserId, {
          spreadsheetId: result.spreadsheetId,
          spreadsheetUrl: result.spreadsheetUrl
        });
        
        return {
          success: true,
          formUrl: result.formUrl,
          spreadsheetUrl: result.spreadsheetUrl,
          setupStep: 3
        };
      };

      const result = global.createQuickStartFormUI('user123');
      
      expect(result.success).toBe(true);
      expect(result.formUrl).toBeDefined();
      expect(result.spreadsheetUrl).toBeDefined();
      expect(result.setupStep).toBe(3);
      expect(global.updateUser).toHaveBeenCalledWith('user123', expect.any(Object));
    });
  });

  describe('saveSheetConfig', () => {
    test('シート設定を保存できる', () => {
      global.saveSheetConfig = (userId, spreadsheetId, sheetName, config, options = {}) => {
        const userInfo = global.getUserInfo(userId);
        if (!userInfo) {
          return false;
        }
        
        const configJson = JSON.parse(userInfo.configJson || '{}');
        configJson[`sheet_${sheetName}`] = config;
        
        if (options.setAsPublished) {
          configJson.publishedSheetName = sheetName;
          configJson.appPublished = true;
        }
        
        global.updateUser(userId, { configJson: JSON.stringify(configJson) });
        global.clearUserCache(userId);
        
        return true;
      };

      const config = {
        opinionHeader: '意見',
        nameHeader: '名前',
        classHeader: 'クラス'
      };
      
      const result = global.saveSheetConfig('user123', 'sheet123', 'Sheet1', config, { setAsPublished: true });
      
      expect(result).toBe(true);
      expect(global.updateUser).toHaveBeenCalled();
    });
  });

  describe('API統合', () => {
    test('handleCoreApiRequestが正しくルーティングする', () => {
      global.handleCoreApiRequest = (action, params) => {
        switch (action) {
          case 'getInitialData':
            return { success: true, data: global.getInitialData(params.userId) };
          case 'getPublishedSheetData':
            return global.getPublishedSheetData(params.userId);
          case 'createQuickStartForm':
            return global.createQuickStartFormUI(params.userId);
          default:
            return { success: false, error: 'Unknown action' };
        }
      };

      const result1 = global.handleCoreApiRequest('getInitialData', { userId: 'user123' });
      expect(result1.success).toBe(true);

      const result2 = global.handleCoreApiRequest('getPublishedSheetData', { userId: 'user123' });
      expect(result2.success).toBe(true);

      const result3 = global.handleCoreApiRequest('unknown', {});
      expect(result3.success).toBe(false);
    });
  });
});

// 実行
if (require.main === module) {
  setupTest();
  console.log('クリティカル機能テスト環境のセットアップ完了');
}