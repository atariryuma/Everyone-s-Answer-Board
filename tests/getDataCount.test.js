const fs = require('fs');
const vm = require('vm');

describe('getDataCount reflects new rows', () => {
  const coreCode = fs.readFileSync('src/Core.gs', 'utf8');
  const mainCode = fs.readFileSync('src/main.gs', 'utf8');
  const unifiedCacheManagerCode = fs.readFileSync('src/unifiedCacheManager.gs', 'utf8');
  const unifiedSheetDataManagerCode = fs.readFileSync('src/unifiedSheetDataManager.gs', 'utf8');
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
        getValues: () =>
          sheetData
            .slice(row - 1, row - 1 + numRows)
            .map((r) => r.slice(col - 1, col - 1 + numCols)),
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
      Utilities: { getUuid: () => `test-uuid-${Math.random()}` },
      CacheService: (() => {
        const scriptCacheStore = {};
        return {
          getScriptCache: () => ({
            get: jest.fn((key) =>
              scriptCacheStore[key] === undefined ? null : scriptCacheStore[key]
            ),
            put: jest.fn((key, value) => {
              scriptCacheStore[key] = value;
            }),
            remove: jest.fn((key) => {
              delete scriptCacheStore[key];
            }),
            removeAll: jest.fn(() => {
              Object.keys(scriptCacheStore).forEach((key) => delete scriptCacheStore[key]);
            }),
            getAll: jest.fn((keys) => {
              const result = {};
              keys.forEach((key) => {
                if (Object.prototype.hasOwnProperty.call(scriptCacheStore, key))
                  result[key] = scriptCacheStore[key];
              });
              return result;
            }),
            putAll: jest.fn((values) => {
              Object.assign(scriptCacheStore, values);
            }),
          }),
          getUserCache: () => ({
            get: jest.fn(),
            put: jest.fn(),
            remove: jest.fn(),
            removeAll: jest.fn(),
          }),
        };
      })(),
      verifyUserAccess: jest.fn(),
      findUserById: jest.fn(() => ({
        userId: 'U1',
        configJson: JSON.stringify({
          publishedSpreadsheetId: 'SID',
          publishedSheetName: 'Sheet1',
          appPublished: true,
        }),
      })),
      getHeaderIndices: () => ({ クラス: 1 }),
      COLUMN_HEADERS: { CLASS: 'クラス' },
      getCurrentUserEmail: () => 'user@example.com',
      // 統一システム必須関数の事前定義
      ULog: {
        debug: jest.fn(),
        info: jest.fn(),
        warn: jest.fn(),
        error: jest.fn(),
      },
      createSuccessResponse: (data) => ({ status: 'success', data }),
      createErrorResponse: (error) => ({ status: 'error', error }),
      cacheManager: {
        get: jest.fn(),
        put: jest.fn(),
        remove: jest.fn(),
      },
    };
    context.global = context;
    vm.createContext(context);
    vm.runInContext(debugConfigCode, context);
    vm.runInContext(unifiedCacheManagerCode, context);
    vm.runInContext(unifiedSheetDataManagerCode, context);
    vm.runInContext(mainCode, context);
    vm.runInContext(coreCode, context);
    context.verifyUserAccess = jest.fn();
  });

  test('returns updated count after adding row', () => {
    const first = context.getDataCount('U1', 'すべて', 'newest', false);
    expect(first.count).toBe(1);

    sheetData.push(['2024-01-02', 'A']);
    context.cacheManager.remove('rowCount_SID_Sheet1_すべて');

    const second = context.getDataCount('U1', 'すべて', 'newest', false);
    expect(second.count).toBe(2);
  });
});
