const fs = require('fs');
const vm = require('vm');

describe('deleteUserAccount cache handling', () => {
  const dbCode = fs.readFileSync('src/database.gs', 'utf8');
  let context;

  beforeEach(() => {
    context = {
      console,
      debugLog: () => {},
      infoLog: () => {},
      warnLog: () => {},
      errorLog: () => {},
      cacheManager: { invalidateRelated: jest.fn(), remove: jest.fn(), clearByPattern: jest.fn() },
      invalidateUserCache: jest.fn(),
      executeWithStandardizedLock: async (lock, name, fn) => await fn(),
      PropertiesService: {
        getScriptProperties: () => ({ getProperty: () => 'db123' }),
        getUserProperties: () => ({ deleteProperty: jest.fn() })
      },
      DB_SHEET_CONFIG: { SHEET_NAME: 'Users', HEADERS: ['userId'] },
      SCRIPT_PROPS_KEYS: { DATABASE_SPREADSHEET_ID: 'DATABASE_SPREADSHEET_ID' },
      getSecureDatabaseId: () => 'db123',
      getResilientScriptProperties: () => ({ getProperty: () => 'db123' }),
      resilientExecutor: { execute: (fn) => fn() }
    };
    vm.createContext(context);
    vm.runInContext(dbCode, context);
    context.findUserById = jest.fn(() => ({ adminEmail: 'admin@example.com', spreadsheetId: 'userSheet' }));
    context.getSheetsServiceCached = () => ({});
    context.getSpreadsheetsData = () => ({ sheets: [{ properties: { title: 'Users', sheetId: 0 } }] });
    // 最初の呼び出しでは削除対象のユーザーを返し、検証時には空の結果を返す
    let callCount = 0;
    context.batchGetSheetsData = () => {
      callCount++;
      if (callCount === 1) {
        // 削除前: ユーザーが存在
        return { valueRanges: [{ values: [['userId'], ['U1']] }] };
      } else {
        // 削除後の検証: ユーザーが削除済み
        return { valueRanges: [{ values: [['userId']] }] };
      }
    };
    context.batchUpdateSpreadsheet = jest.fn().mockResolvedValue({});
    context.logAccountDeletion = jest.fn();
    context.getServiceAccountTokenCached = jest.fn(() => 'mock-token');
    context.UrlFetchApp = {
      fetch: jest.fn(() => ({
        getResponseCode: () => 200,
        getContentText: () => JSON.stringify({ replies: [] })
      }))
    };
  });

  test('passes database id to invalidateUserCache', async () => {
    await context.deleteUserAccount('U1');
    expect(context.invalidateUserCache).toHaveBeenCalledWith('U1', 'admin@example.com', 'userSheet', false, 'db123');
  });
});
