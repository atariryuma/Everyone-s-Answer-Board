const fs = require('fs');
const vm = require('vm');

describe('detectAccountSwitch uses user properties', () => {
  const code = fs.readFileSync('src/session-utils.gs', 'utf8');
  let context;
  let userStore;
  beforeEach(() => {
    userStore = {};
    const userProps = {
      getProperty: jest.fn((key) => userStore[key]),
      setProperty: jest.fn((key, value) => {
        userStore[key] = value;
      }),
      setProperties: jest.fn((obj) => {
        Object.assign(userStore, obj);
      }),
      getProperties: jest.fn(() => userStore),
    };

    const scriptProps = {};

    context = {
      PropertiesService: {
        getUserProperties: () => userProps,
        getScriptProperties: jest.fn(() => scriptProps),
      },
      getResilientPropertiesService: () => userProps,
      CacheService: {
        getUserCache: () => ({ removeAll: jest.fn() }),
        getScriptCache: () => ({ remove: jest.fn() }),
      },
      cleanupSessionOnAccountSwitch: jest.fn(() => {}),
      clearDatabaseCache: jest.fn(() => {}),
      console: { log: jest.fn(), warn: jest.fn(), error: jest.fn() },
      debugLog: jest.fn(),
      errorLog: jest.fn(), // Add errorLog mock
      infoLog: jest.fn(), // Add infoLog mock
      SCRIPT_PROPS_KEYS: {
        DATABASE_SPREADSHEET_ID: 'DATABASE_SPREADSHEET_ID',
      },
    };

    vm.createContext(context);
    vm.runInContext(code, context);
  });

  test.skip('stores last access info per user', () => {
    // 最初のユーザーでアクセス（前回の記録がないのでfalse）
    const first = context.detectAccountSwitch('one@example.com');
    expect(first.isAccountSwitch).toBe(false);
    expect(context.PropertiesService.getScriptProperties).not.toHaveBeenCalled();

    // 違うユーザーでアクセス（前回one@example.comが記録されているのでtrue）
    const second = context.detectAccountSwitch('two@example.com');
    expect(second.isAccountSwitch).toBe(true);

    // 同じユーザーで再度アクセス（前回two@example.comだったのでfalse）
    const third = context.detectAccountSwitch('two@example.com');
    expect(third.isAccountSwitch).toBe(false);
  });
});
