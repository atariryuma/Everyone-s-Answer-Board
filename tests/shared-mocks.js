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
}

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
      },
      deleteProperty: (key) => {
        console.log(`Deleting script property: ${key}`);
        delete mockScriptProperties[key];
        return true;
      },
      setProperties: (props) => {
        Object.assign(mockScriptProperties, props);
        console.log(`Setting script properties: ${JSON.stringify(props)}`);
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
      },
      deleteProperty: (key) => {
        console.log(`Deleting user property: ${key}`);
        delete mockUserProperties[key];
        return true;
      },
      setProperties: (props) => {
        Object.assign(mockUserProperties, props);
        console.log(`Setting user properties: ${JSON.stringify(props)}`);
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
      // console.log(`Mock API call to: ${url}`); // Commented out to reduce noise
      
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
        
        if (url.includes('/values:batchGet')) {
          const rangesParam = url.split('?')[1];
          const ranges = rangesParam.split('&').map(p => decodeURIComponent(p.split('=')[1].split('&')[0])); // Extract range correctly
          
          const valueRanges = ranges.map(range => {
            console.log(`  Mock batchGetSheetsData: Processing range: ${range}`);
            if (range.includes('Users!')) {
              console.log(`    Returning mockDatabase for Users!`);
              return { range: range, values: mockDatabase };
            } else if (range.includes('フォームの回答 1!')) {
              const data = mockSpreadsheetData[spreadsheetId] || [['No data']];
              console.log(`    Returning mockSpreadsheetData for フォームの回答 1!:`, data);
              return { range: range, values: data };
            } else if (range.includes('NonExistentSheet!')) {
              console.log(`    Returning empty array for NonExistentSheet!`);
              return { range: range, values: [] };
            } else if (range.includes('名簿!')) {
              const data = [
                ['メールアドレス', '名前', 'クラス'],
                ['student1@school.edu', '田中太郎', '6-1'],
                ['student2@school.edu', '佐藤花子', '6-1'],
                ['student3@school.edu', '鈴木次郎', '6-2'],
                ['teacher1@school.edu', '山田先生', '教員']
              ];
              console.log(`    Returning mock data for 名簿!:`, data);
              return { range: range, values: data };
            }
            console.log(`    Returning empty array for unknown range: ${range}`);
            return { range: range, values: [] };
          });

          console.log(`  Mock batchGetSheetsData: Final valueRanges:`, JSON.stringify(valueRanges));
          return {
            getContentText: () => JSON.stringify({
              valueRanges: valueRanges
            })
          };
        } else if (url.includes('/values/')) {
          const range = url.match(/\/values\/([^?]+)/)?.[1];
          console.log(`  Mock /values/ call for range: ${range}`);
          
          if (range && range.includes('Users!')) {
            console.log(`    Returning mockDatabase for Users!`);
            return {
              getContentText: () => JSON.stringify({
                values: mockDatabase
              })
            };
          }
          
          if (range && range.includes('フォームの回答 1!')) {
            const data = mockSpreadsheetData[spreadsheetId] || [['No data']];
            console.log(`    Returning mockSpreadsheetData for フォームの回答 1!:`, data);
            return {
              getContentText: () => JSON.stringify({
                values: data
              })
            };
          }
          
          if (range && range.includes('NonExistentSheet!')) {
            console.log(`    Returning empty array for NonExistentSheet!`);
            return {
              getContentText: () => JSON.stringify({
                values: []
              })
            };
          }
          
          if (range && range.includes('名簿!')) {
            const data = [
              ['メールアドレス', '名前', 'クラス'],
              ['student1@school.edu', '田中太郎', '6-1'],
              ['student2@school.edu', '佐藤花子', '6-1'],
              ['student3@school.edu', '鈴木次郎', '6-2'],
              ['teacher1@school.edu', '山田先生', '教員']
            ];
            console.log(`    Returning mock data for 名簿!:`, data);
            return {
              getContentText: () => JSON.stringify({
                values: data
              })
            };
          }
        }
        
        if (url.includes(':append')) {
          const newData = JSON.parse(options.payload);
          if (newData.values && newData.values[0]) {
            const sheetName = decodeURIComponent(url.match(/\/values\/([^!]+)!/)[1]);
            if (mockSpreadsheetData[sheetName]) {
              mockSpreadsheetData[sheetName].push(newData.values[0]);
            } else {
              mockDatabase.push(newData.values[0]); // Default to mockDatabase if sheet not found
            }
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
          const range = decodeURIComponent(url.match(/\/values\/([^?]+)/)[1]);
          const sheetName = range.split('!')[0];
          const rowCol = range.split('!')[1];
          const newData = JSON.parse(options.payload).values;

          if (mockSpreadsheetData[sheetName]) {
            // Simple update logic for mockSpreadsheetData
            const startRow = parseInt(rowCol.match(/\d+/)[0]) - 1;
            const startCol = rowCol.match(/[A-Z]+/)[0].charCodeAt(0) - 'A'.charCodeAt(0);
            
            for (let r = 0; r < newData.length; r++) {
              for (let c = 0; c < newData[r].length; c++) {
                if (mockSpreadsheetData[sheetName][startRow + r] && mockSpreadsheetData[sheetName][startRow + r][startCol + c] !== undefined) {
                  mockSpreadsheetData[sheetName][startRow + r][startCol + c] = newData[r][c];
                }
              }
            }
          } else if (sheetName === 'Users') {
            // Simple update logic for mockDatabase (assuming it's the 'Users' sheet)
            const startRow = parseInt(rowCol.match(/\d+/)[0]) - 1;
            const startCol = rowCol.match(/[A-Z]+/)[0].charCodeAt(0) - 'A'.charCodeAt(0);

            for (let r = 0; r < newData.length; r++) {
              for (let c = 0; c < newData[r].length; c++) {
                if (mockDatabase[startRow + r] && mockDatabase[startRow + r][startCol + c] !== undefined) {
                  mockDatabase[startRow + r][startCol + c] = newData[r][c];
                }
              }
            }
          }
          return {
            getContentText: () => JSON.stringify({
              updatedRows: newData.length,
              updatedCells: newData.reduce((sum, row) => sum + row.length, 0)
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
      getUrl: () => 'https://example.com/exec'
    })
  };

  // LockService mock
  global.LockService = {
    getScriptLock: () => ({
      waitLock: (timeout) => {
        // console.log(`Mock lock acquired with timeout: ${timeout}ms`); // Commented out to reduce noise
      },
      releaseLock: () => {
        // console.log('Mock lock released'); // Commented out to reduce noise
      }
    })
  };

  // CacheService mock for AdvancedCacheManager
  const mockCache = new Map();
  global.CacheService = {
    getScriptCache: () => ({
      get: (key) => mockCache.get(key) || null,
      put: (key, value, ttl) => {
        mockCache.set(key, value);
        // console.log(`Cache put: ${key} (TTL: ${ttl}s)`); // Commented out to reduce noise
      },
      remove: (key) => {
        mockCache.delete(key);
        // console.log(`Cache remove: ${key}`); // Commented out to reduce noise
      },
      getAll: (keys) => {
        const result = {};
        keys.forEach(key => {
          if (mockCache.has(key)) {
            result[key] = mockCache.get(key);
          }
        });
        return result;
      },
      putAll: (values, ttl) => {
        for (const key in values) {
          mockCache.set(key, values[key]);
        }
      }
    }),
    getDocumentCache: () => global.CacheService.getScriptCache(),
    getUserCache: () => global.CacheService.getScriptCache()
  };

  // SpreadsheetApp mock
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({
      getId: () => mockScriptProperties['DATABASE_SPREADSHEET_ID'],
      getUrl: () => `https://docs.google.com/spreadsheets/d/${mockScriptProperties['DATABASE_SPREADSHEET_ID']}`,
      getSheets: () => {
        // Return mock sheets based on mockSpreadsheetData keys
        return Object.keys(mockSpreadsheetData).map(sheetName => ({
          getName: () => sheetName,
          getDataRange: () => ({
            getValues: () => mockSpreadsheetData[sheetName],
            getLastColumn: () => mockSpreadsheetData[sheetName][0] ? mockSpreadsheetData[sheetName][0].length : 0,
            getLastRow: () => mockSpreadsheetData[sheetName].length
          }),
          getRange: (a1Notation) => {
            const values = mockSpreadsheetData[sheetName];
            if (!values || values.length === 0) return { getValues: () => [[]], setValues: () => {} };

            const parseA1 = (a1) => {
              const colMatch = a1.match(/([A-Z]+)/);
              const rowMatch = a1.match(/([0-9]+)/);
              let col = colMatch ? colMatch[1].split('').reduce((sum, char) => sum * 26 + (char.charCodeAt(0) - 'A' + 1), 0) - 1 : 0;
              let row = rowMatch ? parseInt(rowMatch[1], 10) - 1 : 0;
              return { col, row };
            };

            let start = parseA1(a1Notation.split(':')[0]);
            let end = a1Notation.includes(':') ? parseA1(a1Notation.split(':')[1]) : start;

            const selectedRows = values.slice(start.row, end.row + 1);
            const result = selectedRows.map(row => row.slice(start.col, end.col + 1));

            return {
              getValues: () => result,
              setValues: (newValues) => {
                for (let r = 0; r < newValues.length; r++) {
                  for (let c = 0; c < newValues[r].length; c++) {
                    if (values[start.row + r] && values[start.row + r][start.col + c] !== undefined) {
                      values[start.row + r][start.col + c] = newValues[r][c];
                    }
                  }
                }
              }
            };
          },
          getLastColumn: () => mockSpreadsheetData[sheetName][0] ? mockSpreadsheetData[sheetName][0].length : 0,
          getLastRow: () => mockSpreadsheetData[sheetName].length,
          appendRow: (row) => {
            mockSpreadsheetData[sheetName].push(row);
          },
          addEditor: (email) => {
            // console.log(`Mock: Added ${email} as editor to sheet ${sheetName}`); // Commented out to reduce noise
          },
          deleteRow: (rowPosition) => {
            if (mockSpreadsheetData[sheetName] && rowPosition > 0 && rowPosition <= mockSpreadsheetData[sheetName].length) {
              mockSpreadsheetData[sheetName].splice(rowPosition - 1, 1);
            }
          }
        }));
      },
      getSheetByName: (name) => {
        const sheets = global.SpreadsheetApp.getActiveSpreadsheet().getSheets();
        return sheets.find(sheet => sheet.getName() === name) || null;
      },
      openById: (id) => {
        if (id === mockScriptProperties['DATABASE_SPREADSHEET_ID']) {
          return global.SpreadsheetApp.getActiveSpreadsheet();
        }
        return null;
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

    // Recursively get all .gs files
    const getAllGsFiles = (dir) => {
      let gsFiles = [];
      const files = fs.readdirSync(dir);
      for (const file of files) {
        const filePath = path.join(dir, file);
        const stat = fs.statSync(filePath);
        if (stat.isDirectory()) {
          gsFiles = gsFiles.concat(getAllGsFiles(filePath));
        } else if (file.endsWith('.gs')) {
          gsFiles.push(filePath);
        }
      }
      return gsFiles;
    };

    const gsFiles = getAllGsFiles(srcDir);

    // Define a load order for critical files to ensure dependencies are met
    const loadOrder = [
      'main.gs',
      'config.gs',
      'cache.gs',
      'auth.gs',
      'database.gs',
      'url.gs',
      'Core.gs',
      'monitor.gs',
      'optimizer.gs',
      'reactionService.gs'
    ];

    // Create a map for quick lookup of file paths by base name
    const fileMap = new Map();
    gsFiles.forEach(filePath => {
      fileMap.set(path.basename(filePath), filePath);
    });

    // Load files in the specified order
    loadOrder.forEach(filename => {
      const filePath = fileMap.get(filename);
      if (filePath) {
        const content = fs.readFileSync(filePath, 'utf8');
        try {
          (1, eval)(content);
          console.log(`✅ Loaded: ${filePath.replace(srcDir + path.sep, '')}`);
        } catch (error) {
          console.warn(`⚠️ Error loading ${filePath.replace(srcDir + path.sep, '')}:`, error.message);
        }
        fileMap.delete(filename); // Remove from map once loaded
      }
    });

    // Load any remaining .gs files (those not in loadOrder or in other subdirectories)
    fileMap.forEach(filePath => {
      const content = fs.readFileSync(filePath, 'utf8');
      try {
        (1, eval)(content);
        console.log(`✅ Loaded: ${filePath.replace(srcDir + path.sep, '')}`);
      } catch (error) {
        console.warn(`⚠️ Error loading ${filePath.replace(srcDir + path.sep, '')}:`, error.message);
      }
    });

    // Add compatibility functions for tests
    setupCompatibilityFunctions();

    codeEvaluated = true;
    console.log('✅ GAS source evaluation completed');
  }
}

function loadCode() {
  setupGlobalMocks();
  ensureCodeEvaluated();
  return global;
}

// Add compatibility functions for tests
function setupCompatibilityFunctions() {
  // Note: Constants are now defined in main.gs, so we don't need to redefine them
  // unless they're missing in the test environment
  
  if (typeof SCRIPT_PROPS_KEYS === 'undefined') {
    global.SCRIPT_PROPS_KEYS = {
      SERVICE_ACCOUNT_CREDS: 'SERVICE_ACCOUNT_CREDS',
      DATABASE_SPREADSHEET_ID: 'DATABASE_SPREADSHEET_ID'
    };
  }

  global.DB_SHEET_CONFIG = {
    SHEET_NAME: 'Users',
    HEADERS: [
      'userId', 'adminEmail', 'spreadsheetId', 'spreadsheetUrl',
      'createdAt', 'configJson', 'lastAccessedAt', 'isActive'
    ]
  };

  global.LOG_SHEET_CONFIG = {
    SHEET_NAME: 'Logs',
    HEADERS: ['timestamp', 'userId', 'action', 'details']
  };

  global.COLUMN_HEADERS = {
    TIMESTAMP: 'タイムスタンプ',
    EMAIL: 'メールアドレス',
    CLASS: 'クラス',
    OPINION: '回答',
    REASON: '理由',
    NAME: '名前',
    UNDERSTAND: 'なるほど！',
    LIKE: 'いいね！',
    CURIOUS: 'もっと知りたい！',
    HIGHLIGHT: 'ハイライト'
  };

  global.EMAIL_REGEX = /^[^\n@]+@[^\n@]+\.[^\n@]+$/;
  global.DEBUG = true;

  // Compatibility functions for old test references
  global.getServiceAccountToken = getServiceAccountTokenCached;
  global.getWebAppUrl = getWebAppUrlCached;
  global.getHeaderIndices = getHeadersCached;
  global.saveDeployId = function(id) { /* mock implementation if needed */ }; // This function is not in the refactored code, so mock it
  global.checkAdmin = checkAdmin;
  global.findHeaderIndices = function(headers, requiredHeaders) { /* mock implementation if needed */ }; // This function is not in the refactored code, so mock it
  global.guessHeadersFromArray = autoMapHeaders;
  global.prepareSheetForBoard = function() { /* mock implementation if needed */ }; // This function is not in the refactored code, so mock it
  global.buildBoardData = function() { /* mock implementation if needed */ }; // This function is not in the refactored code, so mock it
  global.parseReactionString = parseReactionString;
  global.toggleHighlight = toggleHighlight;
  
  global.setCachedUserInfo = function(userId, userInfo) {
    // Use AdvancedCacheManager for caching
    if (typeof cacheManager !== 'undefined') {
      return cacheManager.get(
        `user_${userId}`,
        () => userInfo,
        { ttl: 3600, enableMemoization: true }
      );
    }
    return userInfo;
  };
  
  global.getCachedUserInfo = function(userId) {
    // Use AdvancedCacheManager for retrieval
    if (typeof cacheManager !== 'undefined') {
      return cacheManager.get(
        `user_${userId}`,
        () => null,
        { enableMemoization: true }
      );
    }
    return null;
  };
  
  global.safeSpreadsheetOperation = function(operation, fallbackValue) {
    try {
      return operation();
    } catch (error) {
      console.warn('Safe operation failed:', error.message);
      return fallbackValue;
    }
  };
  
  global.clearAllCaches = clearAllCaches;

  // Add performCacheCleanup compatibility
  global.performCacheCleanup = clearAllCaches;

  // Missing utility functions
  global.isValidEmail = isValidEmail;

  global.getEmailDomain = getEmailDomain;

  global.getAndCacheHeaderIndices = function(spreadsheetId, sheetName) {
    // Simplified version for testing
    return {
      'タイムスタンプ': 0,
      'メールアドレス': 1,
      'クラス': 2,
      '回答': 3,
      '理由': 4,
      '名前': 5,
      'なるほど！': 6,
      'いいね！': 7,
      'もっと知りたい！': 8,
      'ハイライト': 9
    };
  };

  global.saveDisplayMode = function(mode) {
    mockUserProperties['DISPLAY_MODE'] = mode;
    return { success: true };
  };

  // Add AdvancedCacheManager compatibility for old tests
  global.AdvancedCacheManager = {
    smartGet: function(key, fetchFunction, options) {
      if (typeof cacheManager !== 'undefined') {
        return cacheManager.get(key, fetchFunction, options);
      }
      return fetchFunction ? fetchFunction() : null;
    },
    getHealth: function() {
      if (typeof cacheManager !== 'undefined') {
        return cacheManager.getHealth();
      }
      return { status: 'ok', memoCacheSize: 0 };
    },
    conditionalClear: function(pattern) {
      if (typeof cacheManager !== 'undefined') {
        cacheManager.clearByPattern(pattern);
      }
    }
  };
}

function resetMocks() {
  // Reset state for clean tests
  mockUserProperties = {}; // Reset all user properties
  mockUserProperties['CURRENT_USER_ID'] = 'user-001';
  mockDatabase = [
    ['userId', 'adminEmail', 'spreadsheetId', 'spreadsheetUrl', 'createdAt', 'configJson', 'lastAccessedAt', 'isActive'],
    ['test-user-id-12345', 'test@example.com', 'test-spreadsheet-id', 'https://docs.google.com/spreadsheets/d/test', '2024-01-01T00:00:00.000Z', '{"formUrl":"https://forms.google.com/test"}', '2024-01-01T12:00:00.000Z', 'true'],
    ['user-001', 'teacher1@school.edu', 'spreadsheet-001', 'https://docs.google.com/spreadsheets/d/001', '2024-01-01T00:00:00.000Z', '{"formUrl":"https://forms.google.com/001","publishedSheet":"フォームの回答 1","displayMode":"anonymous"}', '2024-01-01T12:00:00.000Z', 'true'],
    ['user-002', 'teacher2@school.edu', 'spreadsheet-002', 'https://docs.google.com/spreadsheets/d/002', '2024-01-02T00:00:00:000Z', '{"formUrl":"https://forms.google.com/002","publishedSheet":"フォームの回答 1","displayMode":"named"}', '2024-01-02T12:00:00.000Z', 'true'],
    ['user-003', 'teacher3@school.edu', 'spreadsheet-003', 'https://docs.google.com/spreadsheets/d/003', '2024-01-03T00:00:00.000Z', '{"formUrl":"https://forms.google.com/003","publishedSheet":"フォームの回答 1","displayMode":"anonymous","appPublished":true}', '2024-01-03T12:00:00.000Z', 'false']
  ];
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