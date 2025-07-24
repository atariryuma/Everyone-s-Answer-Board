const fs = require('fs');
const vm = require('vm');

describe('detectAccountSwitch uses user properties', () => {
  const code = fs.readFileSync('src/session-utils.gs', 'utf8');
  let context;
  let userStore;
  beforeEach(() => {
    userStore = {};
    const userProps = {
      getProperty: jest.fn(key => userStore[key]),
      setProperty: jest.fn((key, value) => { userStore[key] = value; }),
      setProperties: jest.fn(obj => { Object.assign(userStore, obj); }),
      getProperties: jest.fn(() => userStore)
    };

    const scriptProps = {};

    context = {
      PropertiesService: {
        getUserProperties: () => userProps,
        getScriptProperties: jest.fn(() => scriptProps)
      },
      CacheService: {
        getUserCache: () => ({ removeAll: jest.fn() }),
        getScriptCache: () => ({ remove: jest.fn() })
      },
      cleanupSessionOnAccountSwitch: jest.fn(() => {}),
      clearDatabaseCache: jest.fn(() => {}),
      console: { log: jest.fn(), warn: jest.fn(), error: jest.fn() },
      SCRIPT_PROPS_KEYS: {
        DATABASE_SPREADSHEET_ID: 'DATABASE_SPREADSHEET_ID'
      }
    };

    vm.createContext(context);
    vm.runInContext(code, context);
  });

  test('stores last access info per user', () => {
    const first = context.detectAccountSwitch('one@example.com');
    expect(first.isAccountSwitch).toBe(false);
    expect(context.PropertiesService.getScriptProperties).not.toHaveBeenCalled();

    const second = context.detectAccountSwitch('two@example.com');
    expect(second.isAccountSwitch).toBe(true);
  });
});
