const fs = require('fs');
const vm = require('vm');

describe('getDataCount reflects new rows', () => {
  const coreCode = fs.readFileSync('src/Core.gs', 'utf8');
  const mainCode = fs.readFileSync('src/main.gs', 'utf8');
  const unifiedCacheManagerCode = fs.readFileSync('src/unifiedCacheManager.gs', 'utf8');
  const debugConfigCode = fs.readFileSync('src/debugConfig.gs', 'utf8');
  let context;
  let sheetData;

  beforeEach(() => {
    sheetData = [
      ['Timestamp', 'クラス'],
      ['2024-01-01', 'A'],
    ];
    const getFreshSheet = () => ({
      getLastRow: () => sheetData.length,
      getRange: (row, col, numRows, numCols) => ({
        getValues: () => sheetData
          .slice(row - 1, row - 1 + numRows)
          .map(r => r.slice(col - 1, col - 1 + numCols)),
      }),
    });
    context = {
      console,
      debugLog: () => {},
      PropertiesService: {
        getScriptProperties: () => ({
          getProperty: () => null,
        }),
      },
      Session: { getActiveUser: () => ({ getEmail: () => 'user@example.com' }) },
      SpreadsheetApp: {
        openById: jest.fn(() => ({
          getSheetByName: () => getFreshSheet(),
        })),
      },
      Utilities: { getUuid: () => 'test-uuid-' + Math.random() },
      CacheService: (() => {
        const scriptCacheStore = {};
        return {
          getScriptCache: () => ({
            get: jest.fn(key => (scriptCacheStore[key] === undefined ? null : scriptCacheStore[key])),
            put: jest.fn((key, value) => { scriptCacheStore[key] = value; }),
            remove: jest.fn(key => { delete scriptCacheStore[key]; }),
            removeAll: jest.fn(() => { Object.keys(scriptCacheStore).forEach(key => delete scriptCacheStore[key]); }),
            getAll: jest.fn(keys => {
              const result = {};
              keys.forEach(key => { if (Object.prototype.hasOwnProperty.call(scriptCacheStore, key)) result[key] = scriptCacheStore[key]; });
              return result;
            }),
            putAll: jest.fn(values => { Object.assign(scriptCacheStore, values); }),
          }),
          getUserCache: () => ({
            get: jest.fn(),
            put: jest.fn(),
            remove: jest.fn(),
            removeAll: jest.fn(),
          }),
        };
      })(),
      verifyUserAccess: jest.fn(() => {}), // 後方互換性確保
      findUserById: jest.fn(() => ({
        userId: 'U1',
        configJson: JSON.stringify({
          publishedSpreadsheetId: 'SID',
          publishedSheetName: 'Sheet1',
        }),
      })),
      getHeaderIndices: () => ({ クラス: 1 }),
      COLUMN_HEADERS: { CLASS: 'クラス' },
      getCurrentUserEmail: () => 'user@example.com',
      // マルチテナント対応関数のモック
      validateTenantAccess: jest.fn(() => true),
      auditTenantAccess: jest.fn(),
      auditSecurityViolation: jest.fn(),
      buildSecureUserScopedKey: jest.fn((prefix, userId, key) => `${prefix}_${userId}_${key}`),
      buildUserScopedKey: jest.fn((prefix, userId, key) => `${prefix}_${userId}_${key}`),
      // スプレッドシート関数のモック
      openSpreadsheetOptimized: jest.fn((id) => ({
        getSheetByName: (name) => getFreshSheet(),
      })),
    };
    context.global = context;
    vm.createContext(context);
    vm.runInContext(debugConfigCode, context);
    vm.runInContext(unifiedCacheManagerCode, context);
    vm.runInContext(mainCode, context);
    vm.runInContext(coreCode, context);
    context.verifyUserAccess = jest.fn();
  });

  test('returns updated count after adding row', () => {
    const first = context.getDataCount('U1', 'すべて', 'newest', false);
    expect(first.count).toBe(1); // ヘッダー行 + 1データ行

    sheetData.push(['2024-01-02', 'A']);
    // マルチテナント対応: すべてのキャッシュをクリア
    context.cacheManager.clearAll();

    const second = context.getDataCount('U1', 'すべて', 'newest', false);
    expect(second.count).toBe(2);
  });
});
