const fs = require('fs');
const vm = require('vm');

describe('getInitialData header extraction', () => {
  const urlCode = fs.readFileSync('src/url.gs', 'utf8');
  const mainCode = fs.readFileSync('src/main.gs', 'utf8');
  const coreCode = fs.readFileSync('src/Core.gs', 'utf8');
  const databaseCode = fs.readFileSync('src/database.gs', 'utf8');
  const unifiedUtilitiesCode = fs.readFileSync('src/unifiedUtilities.gs', 'utf8');
  let context;

  beforeEach(() => {
    const store = {};
    context = {
      debugLog: () => {},
      logError: () => {},
      logWarn: () => {},
      logDebug: () => {},
      errorLog: () => {},
      warnLog: () => {},
      infoLog: () => {},
      console,
      PropertiesService: {
        getScriptProperties: () => ({ 
          getProperty: (key) => {
            if (key === 'DATABASE_SPREADSHEET_ID') return 'test-spreadsheet-id';
            if (key === 'ADMIN_EMAIL') return 'admin@example.com';
            if (key === 'SERVICE_ACCOUNT_CREDS') return '{}';
            return null;
          }
        }),
      },
      CacheService: {
        getUserCache: () => ({ get: () => null, put: () => {} }),
        getScriptCache: () => ({ get: (k) => store[k] || null, put: (k, v) => { store[k] = v; }, remove: (k) => { delete store[k]; } })
      },
      cacheManager: {
        store,
        get(key, fn) {
          if (this.store[key]) return this.store[key];
          const val = fn();
          this.store[key] = val;
          return val;
        },
        remove(key) { delete this.store[key]; },
        clearByPattern: (pattern) => {
          for (const key in store) {
            if (key.startsWith(pattern.replace('*', ''))) {
              delete store[key];
            }
          }
        }
      },
      ScriptApp: {
        getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/ID/exec' }),
        getScriptId: () => 'ID'
      },
      Session: { getActiveUser: () => ({ getEmail: () => 'test@example.com' }) },
      Utilities: { 
        getUuid: () => 'uid', 
        computeDigest: () => [], 
        Charset: { UTF_8: 'UTF-8' },
        sleep: (ms) => {} // モックのsleep関数
      },
      getSecureDatabaseId: () => '1234567890123456789012345678901234567890', // 有効なスプレッドシートID形式に修正
      generateNewServiceAccountToken: jest.fn(() => 'mock-service-account-token'), // generateNewServiceAccountToken のモック
    };
    vm.createContext(context);
    vm.runInContext(unifiedUtilitiesCode, context);
    vm.runInContext(urlCode, context);
    vm.runInContext(mainCode, context);
    vm.runInContext(databaseCode, context);
    vm.runInContext(coreCode, context);

    Object.assign(context, {
      verifyUserAccess: jest.fn(),
      getOrFetchUserInfo: jest.fn(() => ({
        userId: 'U',
        adminEmail: 'test@example.com',
        spreadsheetId: '',
        spreadsheetUrl: '',
        configJson: JSON.stringify({
          publishedSheetName: 'ClassA',
          sheet_ClassA: {
            opinionHeader: 'テーマ',
            nameHeader: '氏名',
            classHeader: 'クラス',
          },
        }),
      })),
      findUserById: jest.fn(() => ({
        userId: 'U',
        adminEmail: 'test@example.com',
        spreadsheetId: '',
        spreadsheetUrl: '',
        configJson: JSON.stringify({
          publishedSheetName: 'ClassA',
          sheet_ClassA: {
            opinionHeader: 'テーマ',
            nameHeader: '氏名',
            classHeader: 'クラス',
          },
        }),
      })),
      getSheetsList: jest.fn(() => []),
      generateUserUrls: jest.fn(() => ({ webApp: 'web', adminUrl: 'admin', viewUrl: 'view' })),
      getResponsesData: jest.fn(() => ({ status: 'success', data: [] })),
      determineSetupStep: jest.fn(() => 3),
      getServiceAccountTokenCached: jest.fn(() => 'mock-token'),
      getSheetsService: jest.fn(() => ({ baseUrl: 'mock-base-url' })), // baseUrl を追加
      getCachedSheetsService: jest.fn(() => ({ baseUrl: 'mock-base-url' })), // baseUrl を追加
      unifiedBatchProcessor: { // unifiedBatchProcessor のモック
        batchGet: jest.fn(() => ({
          valueRanges: [{
            values: [
              ['userId', 'adminEmail', 'spreadsheetId', 'spreadsheetUrl', 'configJson'],
              ['U', 'test@example.com', 'mock-sheet-id', 'mock-sheet-url', JSON.stringify({
                publishedSheetName: 'ClassA',
                sheet_ClassA: {
                  opinionHeader: 'テーマ',
                  nameHeader: '氏名',
                  classHeader: 'クラス',
                },
              })]
            ]
          }]
        })),
        batchUpdate: jest.fn(() => ({})),
        batchUpdateSpreadsheet: jest.fn(() => ({})),
        batchCacheOperation: jest.fn(() => ({})),
        getMetrics: jest.fn(() => ({})),
      },
      synchronizeCacheAfterCriticalUpdate: jest.fn(), // synchronizeCacheAfterCriticalUpdate のモック
    });
  });

  test('returns headers from nested sheet config', () => {
    const res = context.getInitialData('U');
    expect(res.config.opinionHeader).toBe('テーマ');
    expect(res.config.nameHeader).toBe('氏名');
    expect(res.config.classHeader).toBe('クラス');
  });
});
