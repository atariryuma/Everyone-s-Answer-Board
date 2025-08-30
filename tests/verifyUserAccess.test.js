const fs = require('fs');
const vm = require('vm');

describe('verifyUserAccess security checks', () => {
  const coreCode = fs.readFileSync('src/Core.gs', 'utf8');
  // ConfigurationManagerとAccessControllerは同じファイルに統合
  const accessControllerCode = fs.readFileSync('src/AccessController.gs', 'utf8');
  let context;

  beforeEach(() => {
    context = {
      console,
      debugLog: () => {},
      Log: {
        debug: jest.fn(),
        info: jest.fn(),
        warn: jest.fn(),
        error: jest.fn(),
      },
      PropertiesService: {
        getScriptProperties: () => ({
          getProperty: jest.fn((key) => {
            if (key === 'user_config_U1') {
              return JSON.stringify({
                tenantId: 'U1',
                ownerEmail: 'admin@example.com',
                isPublic: false,
                allowAnonymous: false
              });
            }
            return null;
          })
        })
      },
      CacheService: {
        getScriptCache: () => ({
          get: jest.fn(() => null),
          put: jest.fn(),
          remove: jest.fn()
        })
      },
      DB: {
        findUserById: jest.fn(() => ({
          ownerEmail: 'admin@example.com',
          configJson: JSON.stringify({ appPublished: false }),
        })),
      },
      User: {
        email: jest.fn(() => 'admin@example.com'),
        info: jest.fn(() => ({
          tenantId: 'U1',
          ownerEmail: 'admin@example.com',
          spreadsheetId: 'testSpreadsheetId',
        })),
      },
      clearExecutionUserInfoCache: jest.fn(),
      Session: { getActiveUser: () => ({ getEmail: () => 'admin@example.com' }) },
      findUserById: jest.fn(() => ({
        ownerEmail: 'admin@example.com',
        configJson: JSON.stringify({ appPublished: false }),
      })),
    };
    vm.createContext(context);
    vm.runInContext(accessControllerCode, context);
    vm.runInContext(coreCode, context);
  });

  test('allows access for matching admin email', () => {
    expect(() => context.verifyUserAccess('U1')).not.toThrow();
  });

  test('denies access when emails differ and board not published', () => {
    context.Session.getActiveUser = () => ({ getEmail: () => 'other@example.com' });
    expect(() => context.verifyUserAccess('U1')).toThrow('認証エラー');
  });

  test('allows read-only access for published board', () => {
    // 新しいコンテキストで公開設定をテスト
    const publicContext = {
      console,
      debugLog: () => {},
      Log: {
        debug: jest.fn(),
        info: jest.fn(),
        warn: jest.fn(),
        error: jest.fn(),
      },
      PropertiesService: {
        getScriptProperties: () => ({
          getProperty: jest.fn((key) => {
            if (key === 'user_config_U1') {
              return JSON.stringify({
                tenantId: 'U1',
                ownerEmail: 'admin@example.com',
                isPublic: true,
                allowAnonymous: true
              });
            }
            return null;
          })
        })
      },
      CacheService: {
        getScriptCache: () => ({
          get: jest.fn(() => null),
          put: jest.fn(),
          remove: jest.fn()
        })
      },
      DB: {
        findUserById: jest.fn(() => ({
          ownerEmail: 'admin@example.com',
          configJson: JSON.stringify({ appPublished: false }),
        })),
      },
      User: {
        email: jest.fn(() => 'admin@example.com'),
        info: jest.fn(() => ({
          tenantId: 'U1',
          ownerEmail: 'admin@example.com',
          spreadsheetId: 'testSpreadsheetId',
        })),
      },
      clearExecutionUserInfoCache: jest.fn(),
      Session: { getActiveUser: () => ({ getEmail: () => 'viewer@example.com' }) },
      findUserById: jest.fn(() => ({
        ownerEmail: 'admin@example.com',
        configJson: JSON.stringify({ appPublished: false }),
      })),
    };
    
    vm.createContext(publicContext);
    vm.runInContext(accessControllerCode, publicContext);
    vm.runInContext(coreCode, publicContext);
    
    expect(() => publicContext.verifyUserAccess('U1')).not.toThrow();
  });
});
