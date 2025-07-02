/**
 * Shared Mock Environment for Service Account Database Tests
 * 共有モック環境 - 全てのテストで一貫した環境を提供
 */

const fs = require('fs');
const path = require('path');

// Mock property storage
let mockScriptProperties = {
  'SERVICE_ACCOUNT_CREDS': JSON.stringify({
    "type": "service_account",
    "project_id": "test-project",
    "private_key_id": "test-key-id",
    "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQC...\n-----END PRIVATE KEY-----\n",
    "client_email": "test-service@test-project.iam.gserviceaccount.com",
    "client_id": "123456789",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token"
  }),
  'DATABASE_SPREADSHEET_ID': '1BcD3fGhIjKlMnOpQrStUvWxYz0123456789ABCDEFGH'
};

// Mock user property storage
let mockUserProperties = {
  'CURRENT_USER_ID': 'user-001'
};

// More detailed mock database for advanced tests
let mockDatabase = [
  ['userId', 'adminEmail', 'spreadsheetId', 'spreadsheetUrl', 'createdAt', 'configJson', 'lastAccessedAt', 'isActive'],
  ['test-user-id-12345', 'test@example.com', 'test-spreadsheet-id', 'https://docs.google.com/spreadsheets/d/test', '2024-01-01T00:00:00.000Z', '{"formUrl":"https://forms.google.com/test"}', '2024-01-01T12:00:00.000Z', 'true'],
  ['user-001', 'teacher1@school.edu', 'spreadsheet-001', 'https://docs.google.com/spreadsheets/d/001', '2024-01-01T00:00:00.000Z', '{"formUrl":"https://forms.google.com/001","publishedSheet":"フォームの回答 1","displayMode":"anonymous"}', '2024-01-01T12:00:00.000Z', 'true'],
  ['user-002', 'teacher2@school.edu', 'spreadsheet-002', 'https://docs.google.com/spreadsheets/d/002', '2024-01-02T00:00:00.000Z', '{"formUrl":"https://forms.google.com/002","publishedSheet":"フォームの回答 1","displayMode":"named"}', '2024-01-02T12:00:00.000Z', 'true'],
  ['user-003', 'teacher3@school.edu', 'spreadsheet-003', 'https://docs.google.com/spreadsheets/d/003', '2024-01-03T00:00:00.000Z', '{"formUrl":"https://forms.google.com/003","publishedSheet":"フォームの回答 1","displayMode":"anonymous","appPublished":true}', '2024-01-03T12:00:00.000Z', 'false']
];

let mockSpreadsheetData = {
  'spreadsheet-001': [
    ['タイムスタンプ', 'メールアドレス', 'クラス', '回答', '理由', '名前', 'なるほど！', 'いいね！', 'もっと知りたい！', 'ハイライト'],
    ['2024-01-01 10:00:00', 'student1@school.edu', '6-1', '宇宙は無限だと思います', '星がたくさんあるから', '田中太郎', '', 'teacher1@school.edu, student2@school.edu', '', 'false'],
    ['2024-01-01 10:05:00', 'student2@school.edu', '6-1', '宇宙には終わりがあると思います', '物理的な限界があるはず', '佐藤花子', 'teacher1@school.edu', '', 'student1@school.edu', 'true'],
    ['2024-01-01 10:10:00', 'student3@school.edu', '6-2', '分からないです', 'まだ勉強中だから', '鈴木次郎', '', 'teacher1@school.edu', 'teacher1@school.edu', 'false']
  ]
};

function setupGlobalMocks() {
  // Ensure console is available
  global.console = console;

  // PropertiesService mock
  global.PropertiesService = {
    getScriptProperties: () => ({
      getProperty: (key) => {
        return mockScriptProperties[key] || null;
      },
      setProperty: (key, value) => {
        console.log(`Setting property: ${key} = ${value}`);
        mockScriptProperties[key] = value;
        return true;
      }
    }),
    getUserProperties: () => ({
      getProperty: (key) => {
        return mockUserProperties[key] || null;
      },
      setProperty: (key, value) => {
        console.log(`Setting user property: ${key} = ${value}`);
        mockUserProperties[key] = value;
        return true;
      }
    })
  };

  // Utilities mock
  global.Utilities = {
    getUuid: () => 'mock-uuid-' + Date.now(),
    base64EncodeWebSafe: (str) => Buffer.from(str).toString('base64'),
    computeRsaSha256Signature: (input, key) => 'mock-signature-' + input.length
  };

  // UrlFetchApp mock
  global.UrlFetchApp = {
    fetch: (url, options) => {
      console.log(`Mock API call to: ${url}`);
      
      if (url.includes('oauth2/v4/token')) {
        return {
          getContentText: () => JSON.stringify({
            access_token: 'mock-access-token-' + Date.now(),
            expires_in: 3600
          })
        };
      }
      
      if (url.includes('sheets.googleapis.com')) {
        const spreadsheetId = url.match(/\/spreadsheets\/([^\/]+)/)?.[1];
        
        if (url.includes('/spreadsheets/') && !url.includes('/values/')) {
          return {
            getContentText: () => JSON.stringify({
              properties: { title: `Test Database Spreadsheet` },
              sheets: [
                { properties: { title: 'Users', sheetId: 0 } },
                { properties: { title: 'フォームの回答 1', sheetId: 1 } },
                { properties: { title: '名簿', sheetId: 2 } }
              ]
            })
          };
        }
        
        if (url.includes('/values/')) {
          const range = url.match(/\/values\/([^?]+)/)?.[1];
          
          if (range && range.includes('Users!')) {
            return {
              getContentText: () => JSON.stringify({
                values: mockDatabase
              })
            };
          }
          
          if (range && range.includes('フォームの回答 1!')) {
            const data = mockSpreadsheetData[spreadsheetId] || [['No data']];
            return {
              getContentText: () => JSON.stringify({
                values: data
              })
            };
          }
          
          if (range && range.includes('NonExistentSheet!')) {
            return {
              getContentText: () => JSON.stringify({
                values: []
              })
            };
          }
          
          if (range && range.includes('名簿!')) {
            return {
              getContentText: () => JSON.stringify({
                values: [
                  ['メールアドレス', '名前', 'クラス'],
                  ['student1@school.edu', '田中太郎', '6-1'],
                  ['student2@school.edu', '佐藤花子', '6-1'],
                  ['student3@school.edu', '鈴木次郎', '6-2'],
                  ['teacher1@school.edu', '山田先生', '教員']
                ]
              })
            };
          }
        }
        
        if (url.includes(':append')) {
          const newData = JSON.parse(options.payload);
          if (newData.values && newData.values[0]) {
            mockDatabase.push(newData.values[0]);
          }
          return {
            getContentText: () => JSON.stringify({
              updates: { updatedRows: 1, updatedCells: newData.values[0].length }
            })
          };
        }
        
        if (url.includes(':batchUpdate')) {
          return {
            getContentText: () => JSON.stringify({
              replies: [{ updatedRows: 1 }]
            })
          };
        }
        
        if (url.includes('/values/') && options && options.method === 'put') {
          return {
            getContentText: () => JSON.stringify({
              updatedRows: 1,
              updatedCells: 1
            })
          };
        }
      }
      
      throw new Error(`Unexpected API call: ${url}`);
    }
  };

  // Session mock
  global.Session = {
    getActiveUser: () => ({
      getEmail: () => 'teacher1@school.edu'
    })
  };

  // ScriptApp mock
  global.ScriptApp = {
    getService: () => ({
      getUrl: () => 'https://script.google.com/macros/s/test-app-id/exec'
    })
  };

  // LockService mock
  global.LockService = {
    getScriptLock: () => ({
      waitLock: (timeout) => {
        console.log(`Mock lock acquired with timeout: ${timeout}ms`);
      },
      releaseLock: () => {
        console.log('Mock lock released');
      }
    })
  };
}

// Code.gs evaluation with proper global scope
let codeEvaluated = false;

function ensureCodeEvaluated() {
  if (!codeEvaluated) {
    console.log('🔧 Evaluating GAS source files...');

    const srcDir = path.join(__dirname, '../src');
    const files = fs.readdirSync(srcDir).filter(f => f.endsWith('.gs'));
    files.forEach(file => {
      const content = fs.readFileSync(path.join(srcDir, file), 'utf8');
      (1, eval)(content);
    });

    codeEvaluated = true;
    console.log('✅ GAS source evaluation completed');
  }
}

function loadCode() {
  setupGlobalMocks();
  ensureCodeEvaluated();
  return global;
}

function resetMocks() {
  // Reset state for clean tests
  mockUserProperties = {
    'CURRENT_USER_ID': 'user-001'
  };
  codeEvaluated = false;
}

module.exports = {
  setupGlobalMocks,
  ensureCodeEvaluated,
  loadCode,
  resetMocks,
  mockDatabase,
  mockSpreadsheetData
};