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
    context.batchGetSheetsData = () => ({ valueRanges: [{ values: [['userId'], ['U1']] }] });
    context.batchUpdateSpreadsheet = jest.fn().mockResolvedValue({});
    context.logAccountDeletion = jest.fn();
  });

  test('passes database id to invalidateUserCache', async () => {
    await context.deleteUserAccount('U1');
    expect(context.invalidateUserCache).toHaveBeenCalledWith('U1', 'admin@example.com', 'userSheet', false, 'db123');
  });
});
