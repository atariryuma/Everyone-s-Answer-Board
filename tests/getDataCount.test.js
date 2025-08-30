const fs = require('fs');
const vm = require('vm');

describe('getDataCount reflects new rows', () => {
  const coreCode = fs.readFileSync('src/Core.gs', 'utf8');
  const mainCode = fs.readFileSync('src/main.gs', 'utf8');
  let context;
  let sheetData;

  beforeEach(() => {
    sheetData = [
      ['Timestamp', 'クラス'],
      // 初期状態はヘッダー行のみでデータ行なし
    ];
    const sheet = {
      getLastRow: () => sheetData.length,
      getRange: (row, col, numRows, numCols) => ({
        getValues: () =>
          sheetData
            .slice(row - 1, row - 1 + numRows)
            .map((r) => r.slice(col - 1, col - 1 + numCols)),
      }),
    };
    context = {
      console,
      debugLog: () => {},
      Log: {
        debug: jest.fn(),
        info: jest.fn(),
        warn: jest.fn(),
        error: jest.fn(),
      },
      cacheManager: {
        store: {},
        get(key, fn) {
          return Object.prototype.hasOwnProperty.call(this.store, key)
            ? this.store[key]
            : (this.store[key] = fn());
        },
        remove(key) {
          delete this.store[key];
        },
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
      verifyUserAccess: jest.fn(),
      DB: {
        findUserById: jest.fn(() => ({
          userId: 'U1',
          configJson: JSON.stringify({
            publishedSpreadsheetId: 'SID',
            publishedSheetName: 'Sheet1',
          }),
        })),
      },
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
    vm.runInContext(mainCode, context);
    vm.runInContext(coreCode, context);
    context.verifyUserAccess = jest.fn();
  });

  test('returns updated count after adding row', () => {
    // 初期状態：ヘッダー行のみ（データ行数は0）
    const first = context.getDataCount('U1', 'すべて', 'newest', false);
    expect(first.count).toBe(0); // ヘッダー行を除くと0行

    // データ行を追加してカウント確認
    sheetData.push(['2024-01-02', 'A']);
    context.cacheManager.remove('rowCount_SID_Sheet1_すべて');
    const second = context.getDataCount('U1', 'すべて', 'newest', false);
    expect(second.count).toBe(1); // ヘッダー行を除くと1行
  });
});
