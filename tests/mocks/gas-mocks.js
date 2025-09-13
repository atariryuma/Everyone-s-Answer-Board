/**
 * GAS API モック定義 - CLAUDE.md準拠版
 * 🚀 Google Apps Script API の完全モック実装
 * 
 * 特徴:
 * - 全GAS APIの包括的モック
 * - configJSON中心型データ構造対応
 * - 5フィールドデータベースモック
 * - リアルな動作シミュレーション
 */

/**
 * 🔧 SpreadsheetApp 完全モック
 * configJSON中心型設計に対応した高機能モック
 */
const createSpreadsheetAppMock = () => {
  const mockSheets = new Map();
  const mockSpreadsheets = new Map();
  
  // デフォルトのテストスプレッドシート作成
  const createDefaultTestSheet = (spreadsheetId = 'test-spreadsheet-id') => {
    const mockData = [
      ['userId', 'userEmail', 'isActive', 'configJson', 'lastModified'],
      [
        'test-user-1',
        'test1@example.com',
        true,
        JSON.stringify({
          spreadsheetId: 'user-sheet-1',
          sheetName: '回答データ',
          setupStatus: 'completed',
          appPublished: true,
          createdAt: '2025-01-01T00:00:00Z'
        }),
        '2025-01-01T12:00:00Z'
      ],
      [
        'test-user-2', 
        'test2@example.com',
        true,
        JSON.stringify({
          spreadsheetId: 'user-sheet-2',
          sheetName: '回答データ',
          setupStatus: 'data_source_set',
          appPublished: false,
          createdAt: '2025-01-01T01:00:00Z'
        }),
        '2025-01-01T13:00:00Z'
      ]
    ];
    
    return {
      getId: () => spreadsheetId,
      getName: () => `テストスプレッドシート ${spreadsheetId}`,
      getSheets: jest.fn(() => [
        {
          getName: () => 'Users',
          getSheetId: () => 0
        },
        {
          getName: () => '回答データ',
          getSheetId: () => 1
        }
      ]),
      getSheetByName: jest.fn((name) => {
        if (name === 'Users' || name === DB_CONFIG.SHEET_NAME) {
          return {
            getName: () => name,
            getSheetId: () => 0,
            getLastRow: jest.fn(() => mockData.length),
            getLastColumn: jest.fn(() => mockData[0].length),
            getRange: jest.fn((range) => {
              // 範囲指定の解析（簡易版）
              if (typeof range === 'string') {
                if (range === 'A1:E1') {
                  return {
                    getValues: jest.fn(() => [mockData[0]]),
                    setValues: jest.fn()
                  };
                }
                if (range.startsWith('A1:')) {
                  return {
                    getValues: jest.fn(() => mockData),
                    setValues: jest.fn((values) => {
                      // データ更新のシミュレーション
                      mockData.length = 0;
                      mockData.push(...values);
                    })
                  };
                }
              }
              
              // 数値指定の場合
              return {
                getValue: jest.fn(() => mockData[1] ? mockData[1][0] : ''),
                setValue: jest.fn((value) => {
                  if (mockData[1]) mockData[1][0] = value;
                }),
                getValues: jest.fn(() => mockData.slice(1)),
                setValues: jest.fn((values) => {
                  mockData.splice(1, mockData.length - 1, ...values);
                })
              };
            }),
            appendRow: jest.fn((row) => {
              mockData.push(row);
            }),
            insertRowAfter: jest.fn((afterRow) => {
              mockData.splice(afterRow + 1, 0, new Array(mockData[0].length).fill(''));
            }),
            deleteRow: jest.fn((rowIndex) => {
              mockData.splice(rowIndex - 1, 1);
            }),
            getDataRange: jest.fn(() => ({
              getValues: jest.fn(() => mockData),
              getLastRow: jest.fn(() => mockData.length),
              getLastColumn: jest.fn(() => mockData[0] ? mockData[0].length : 0)
            }))
          };
        }
        
        // その他のシート
        return {
          getName: () => name,
          getSheetId: () => 999,
          getLastRow: jest.fn(() => 10),
          getLastColumn: jest.fn(() => 6),
          getRange: jest.fn(() => ({
            getValues: jest.fn(() => [
              ['タイムスタンプ', 'メールアドレス', 'クラス', '回答', '理由', '名前'],
              ['2025-01-01 10:00:00', 'student1@example.com', '1年A組', 'テスト回答1', 'テスト理由1', '学生1'],
              ['2025-01-01 11:00:00', 'student2@example.com', '1年B組', 'テスト回答2', 'テスト理由2', '学生2']
            ]),
            setValues: jest.fn()
          })),
          appendRow: jest.fn(),
          getDataRange: jest.fn(() => ({
            getValues: jest.fn(() => [])
          }))
        };
      }),
      insertSheet: jest.fn((name) => ({
        getName: () => name,
        getSheetId: () => Math.floor(Math.random() * 1000)
      }))
    };
  };
  
  return {
    openById: jest.fn((id) => {
      if (!mockSpreadsheets.has(id)) {
        mockSpreadsheets.set(id, createDefaultTestSheet(id));
      }
      return mockSpreadsheets.get(id);
    }),
    
    create: jest.fn((name) => {
      const newId = `created-${Date.now()}`;
      const newSheet = createDefaultTestSheet(newId);
      mockSpreadsheets.set(newId, newSheet);
      return newSheet;
    }),
    
    getActiveSpreadsheet: jest.fn(() => {
      return createDefaultTestSheet('active-spreadsheet');
    }),
    
    flush: jest.fn()
  };
};

/**
 * 🗂️ PropertiesService 高機能モック
 * システム設定とユーザー設定の完全シミュレーション
 */
const createPropertiesServiceMock = () => {
  const scriptProperties = new Map();
  const userProperties = new Map();
  
  // Properties Service keys (for testing)
  const PROPS_KEYS = {
    DATABASE_SPREADSHEET_ID: 'DATABASE_SPREADSHEET_ID',
    ADMIN_EMAIL: 'ADMIN_EMAIL',
    SERVICE_ACCOUNT_CREDS: 'SERVICE_ACCOUNT_CREDS'
  };
  
  // デフォルト設定の初期化
  scriptProperties.set(PROPS_KEYS.DATABASE_SPREADSHEET_ID, 'test-database-sheet-id');
  scriptProperties.set(PROPS_KEYS.ADMIN_EMAIL, 'admin@example.com');
  scriptProperties.set(PROPS_KEYS.SERVICE_ACCOUNT_CREDS, JSON.stringify({
    type: 'service_account',
    project_id: 'test-project',
    private_key_id: 'test-key-id',
    client_email: 'test-service@test-project.iam.gserviceaccount.com'
  }));
  
  return {
    getScriptProperties: jest.fn(() => ({
      getProperty: jest.fn((key) => scriptProperties.get(key) || null),
      setProperty: jest.fn((key, value) => {
        scriptProperties.set(key, value);
        return this;
      }),
      deleteProperty: jest.fn((key) => {
        scriptProperties.delete(key);
        return this;
      }),
      getProperties: jest.fn(() => Object.fromEntries(scriptProperties)),
      setProperties: jest.fn((props) => {
        Object.entries(props).forEach(([key, value]) => {
          scriptProperties.set(key, value);
        });
        return this;
      })
    })),
    
    getUserProperties: jest.fn(() => ({
      getProperty: jest.fn((key) => userProperties.get(key) || null),
      setProperty: jest.fn((key, value) => {
        userProperties.set(key, value);
        return this;
      }),
      deleteProperty: jest.fn((key) => {
        userProperties.delete(key);
        return this;
      }),
      getProperties: jest.fn(() => Object.fromEntries(userProperties))
    }))
  };
};

/**
 * 💾 CacheService 高性能モック
 * TTL対応の本格的なキャッシュシミュレーション
 */
const createCacheServiceMock = () => {
  const cache = new Map();
  const expiry = new Map();
  
  const isExpired = (key) => {
    const expiryTime = expiry.get(key);
    return expiryTime && Date.now() > expiryTime;
  };
  
  return {
    getScriptCache: jest.fn(() => ({
      get: jest.fn((key) => {
        if (isExpired(key)) {
          cache.delete(key);
          expiry.delete(key);
          return null;
        }
        return cache.get(key) || null;
      }),
      
      put: jest.fn((key, value, expirationInSeconds = 3600) => {
        cache.set(key, value);
        expiry.set(key, Date.now() + (expirationInSeconds * 1000));
        return this;
      }),
      
      putAll: jest.fn((values, expirationInSeconds = 3600) => {
        Object.entries(values).forEach(([key, value]) => {
          cache.set(key, value);
          expiry.set(key, Date.now() + (expirationInSeconds * 1000));
        });
        return this;
      }),
      
      remove: jest.fn((key) => {
        cache.delete(key);
        expiry.delete(key);
        return this;
      }),
      
      removeAll: jest.fn(() => {
        cache.clear();
        expiry.clear();
        return this;
      })
    }))
  };
};

/**
 * 🌐 UrlFetchApp 実用的モック
 * HTTP通信の詳細シミュレーション
 */
const createUrlFetchAppMock = () => {
  const mockResponses = new Map();
  
  // デフォルトレスポンスの設定
  mockResponses.set('https://sheets.googleapis.com', {
    responseCode: 200,
    contentText: JSON.stringify({
      values: [
        ['タイムスタンプ', 'メールアドレス', '回答', '理由'],
        ['2025-01-01 10:00:00', 'test@example.com', 'テスト回答', 'テスト理由']
      ]
    }),
    headers: {
      'Content-Type': 'application/json'
    }
  });
  
  return {
    fetch: jest.fn((url, options = {}) => {
      // URLに基づくレスポンス決定
      let response = mockResponses.get(url);
      
      if (!response) {
        // デフォルトレスポンス
        if (url.includes('googleapis.com')) {
          response = {
            responseCode: 200,
            contentText: JSON.stringify({ success: true }),
            headers: { 'Content-Type': 'application/json' }
          };
        } else {
          response = {
            responseCode: 404,
            contentText: 'Not Found',
            headers: { 'Content-Type': 'text/plain' }
          };
        }
      }
      
      return {
        getResponseCode: jest.fn(() => response.responseCode),
        getContentText: jest.fn(() => response.contentText),
        getBlob: jest.fn(() => ({
          getBytes: () => new TextEncoder().encode(response.contentText),
          getName: () => 'response.json'
        })),
        getHeaders: jest.fn(() => response.headers),
        getAllHeaders: jest.fn(() => response.headers)
      };
    }),
    
    fetchAll: jest.fn((requests) => {
      return requests.map(request => {
        return exports.UrlFetchApp.fetch(
          typeof request === 'string' ? request : request.url,
          typeof request === 'object' ? request : {}
        );
      });
    })
  };
};

/**
 * 📝 HtmlService 包括的モック
 * テンプレート処理とHTMLサービスの完全実装
 */
const createHtmlServiceMock = () => {
  const templates = new Map();
  
  // テンプレートファイルのモック内容
  templates.set('Page', '<html><body>回答ボードページ</body></html>');
  templates.set('AdminPanel', '<html><body>管理パネルページ</body></html>');
  templates.set('AppSetupPage', '<html><body>セットアップページ</body></html>');
  
  return {
    createTemplateFromFile: jest.fn((filename) => {
      const content = templates.get(filename) || `<html><body>Template: ${filename}</body></html>`;
      
      return {
        evaluate: jest.fn(() => ({
          setXFrameOptionsMode: jest.fn(function(mode) { return this; }),
          getContent: jest.fn(() => content),
          setTitle: jest.fn(function(title) { return this; })
        })),
        getRawContent: jest.fn(() => content),
        
        // テンプレート変数の設定
        data: {},
        setData: jest.fn(function(data) {
          this.data = { ...this.data, ...data };
          return this;
        })
      };
    }),
    
    createHtmlOutput: jest.fn((content = '<html><body>Mock HTML Output</body></html>') => ({
      setXFrameOptionsMode: jest.fn(function(mode) { return this; }),
      getContent: jest.fn(() => content),
      setTitle: jest.fn(function(title) { return this; }),
      append: jest.fn(function(content) { return this; })
    })),
    
    createTemplate: jest.fn((content) => ({
      evaluate: jest.fn(() => ({
        setXFrameOptionsMode: jest.fn(function(mode) { return this; }),
        getContent: jest.fn(() => content)
      })),
      getRawContent: jest.fn(() => content)
    })),
    
    XFrameOptionsMode: {
      ALLOWALL: 'ALLOWALL',
      SAMEORIGIN: 'SAMEORIGIN', 
      DENY: 'DENY'
    }
  };
};

/**
 * 🚀 統合エクスポート
 * 全GAS APIモックの統一提供
 */
module.exports = {
  SpreadsheetApp: createSpreadsheetAppMock(),
  PropertiesService: createPropertiesServiceMock(),
  CacheService: createCacheServiceMock(),
  UrlFetchApp: createUrlFetchAppMock(),
  HtmlService: createHtmlServiceMock(),
  
  // 基本サービス
  Session: {
    getActiveUser: jest.fn(() => ({
      getEmail: jest.fn(() => 'test@example.com'),
      getName: jest.fn(() => 'テストユーザー')
    })),
    getEffectiveUser: jest.fn(() => ({
      getEmail: jest.fn(() => 'test@example.com')
    })),
    getTemporaryActiveUserKey: jest.fn(() => `temp-key-${Date.now()}`)
  },
  
  Utilities: {
    getUuid: jest.fn(() => `test-uuid-${Date.now()}`),
    sleep: jest.fn((ms) => {
      // 実際にはsleepしない（テスト高速化）
      return Promise.resolve();
    }),
    formatDate: jest.fn((date, timezone, format) => {
      return date.toISOString();
    }),
    jsonStringify: jest.fn((obj) => JSON.stringify(obj)),
    jsonParse: jest.fn((str) => JSON.parse(str)),
    base64Encode: jest.fn((input) => Buffer.from(input).toString('base64')),
    base64Decode: jest.fn((input) => Buffer.from(input, 'base64').toString()),
    computeHmacSha256Signature: jest.fn(() => 'mock-hmac-signature')
  },
  
  ScriptApp: {
    getProjectId: jest.fn(() => 'test-gas-project-id'),
    newTrigger: jest.fn(() => ({
      timeBased: jest.fn(() => ({
        everyMinutes: jest.fn(() => ({ create: jest.fn() })),
        everyHours: jest.fn(() => ({ create: jest.fn() })),
        everyDays: jest.fn(() => ({ create: jest.fn() })),
        at: jest.fn(() => ({ create: jest.fn() }))
      })),
      handlerFunction: jest.fn(function(name) { return this; }),
      create: jest.fn()
    }))
  },
  
  ContentService: {
    createTextOutput: jest.fn((content = '') => ({
      setMimeType: jest.fn(function(mimeType) { return this; }),
      getContent: jest.fn(() => content),
      setContent: jest.fn(function(content) { return this; })
    })),
    MimeType: {
      JSON: 'application/json',
      TEXT: 'text/plain',
      HTML: 'text/html',
      XML: 'application/xml'
    }
  },
  
  Logger: {
    log: jest.fn((message) => {
      console.log(`[GAS Logger] ${message}`);
    })
  },
  
  LockService: {
    getScriptLock: jest.fn(() => ({
      tryLock: jest.fn(() => true),
      waitLock: jest.fn((timeoutInMillis) => true),
      releaseLock: jest.fn(),
      hasLock: jest.fn(() => true)
    })),
    getUserLock: jest.fn(() => ({
      tryLock: jest.fn(() => true),
      waitLock: jest.fn(() => true),
      releaseLock: jest.fn()
    }))
  },

  // Test Data
  testUserData: {
    userId: 'test-user-123',
    userEmail: 'test@example.com',
    isActive: true,
    configJson: JSON.stringify({
      spreadsheetId: 'test-sheet-id',
      sheetName: '回答データ',
      setupStatus: 'completed',
      appPublished: true
    }),
    lastModified: '2025-01-15T10:00:00Z'
  },

  testSheetData: [
    {
      id: 'row_2',
      timestamp: '2025-01-15 10:00',
      email: 'student1@example.com',
      answer: 'テスト回答1',
      reason: 'テスト理由1',
      className: 'A組',
      name: '学生1',
      reactions: { understand: 1, like: 2, curious: 0 },
      isHighlighted: false
    },
    {
      id: 'row_3', 
      timestamp: '2025-01-15 10:05',
      email: 'student2@example.com',
      answer: 'テスト回答2',
      reason: 'テスト理由2',
      className: 'B組',
      name: '学生2',
      reactions: { understand: 0, like: 1, curious: 1 },
      isHighlighted: true
    }
  ],

  testConfigData: {
    spreadsheetId: 'test-sheet-id',
    sheetName: '回答データ',
    setupStatus: 'completed',
    appPublished: true,
    displaySettings: {
      showNames: false,
      showReactions: true
    }
  },

  createMockSpreadsheet: (id = 'mock-sheet-id') => ({
    getId: jest.fn(() => id),
    getName: jest.fn(() => 'テストスプレッドシート'),
    getSheetByName: jest.fn(() => ({
      getLastRow: jest.fn(() => 10),
      getLastColumn: jest.fn(() => 7),
      getRange: jest.fn(() => ({
        getValues: jest.fn(() => [['ヘッダー1', 'ヘッダー2']]),
        getValue: jest.fn(() => 'テスト値'),
        setValue: jest.fn()
      }))
    }))
  })
};