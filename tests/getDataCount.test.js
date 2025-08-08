const fs = require('fs');
const vm = require('vm');

describe('getDataCount reflects new rows', () => {
  const coreCode = fs.readFileSync('src/Core.gs', 'utf8');
  const mainCode = fs.readFileSync('src/main.gs', 'utf8');
  const spreadsheetCacheCode = fs.readFileSync('src/spreadsheetCache.gs', 'utf8');
  const debugConfigCode = fs.readFileSync('src/debugConfig.gs', 'utf8');
  let context;
  let sheetData;

  beforeEach(() => {
    sheetData = [
      ['Timestamp', 'クラス'],
      ['2024-01-01', 'A'],
    ];
    const sheet = {
      getLastRow: () => sheetData.length,
      getRange: (row, col, numRows, numCols) => ({
        getValues: () => sheetData
          .slice(row - 1, row - 1 + numRows)
          .map(r => r.slice(col - 1, col - 1 + numCols)),
      }),
    };
    context = {
      console,
      debugLog: () => {},
      cacheManager: {
        store: {},
        get(key, fn) { return this.store.hasOwnProperty(key) ? this.store[key] : (this.store[key] = fn()); },
        remove(key) { delete this.store[key]; },
      },
      PropertiesService: {
        getScriptProperties: () => ({
          getProperty: () => null,
        }),
      },
      Session: { getActiveUser: () => ({ getEmail: () => 'user@example.com' }) },
      SpreadsheetApp: {
        openById: () => ({ getSheetByName: () => sheet }),
      },
      Utilities: { getUuid: () => 'test-uuid-' + Math.random() },
      verifyUserAccess: jest.fn(),
      findUserById: jest.fn(() => ({
        userId: 'U1',
        configJson: JSON.stringify({
          publishedSpreadsheetId: 'SID',
          publishedSheetName: 'Sheet1',
        }),
      })),
      getHeaderIndices: () => ({ クラス: 1 }),
      COLUMN_HEADERS: { CLASS: 'クラス' },
    };
    vm.createContext(context);
    vm.runInContext(debugConfigCode, context);
    vm.runInContext(spreadsheetCacheCode, context);
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

