const fs = require('fs');
const vm = require('vm');

describe('verifyUserAccess security checks', () => {
  const coreCode = fs.readFileSync('src/Core.gs', 'utf8');
  // ConfigurationManagerとAccessControllerは新しいBase.gsに統合
  const baseCode = fs.readFileSync('src/Base.gs', 'utf8');
  const appCode = fs.readFileSync('src/App.gs', 'utf8');
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
                userId: 'U1',
                userEmail: 'admin@example.com',
                isPublic: false,
                allowAnonymous: false
              });
            }
            if (key === 'ADMIN_EMAIL') {
              return 'admin@example.com';
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
          userEmail: 'admin@example.com',
          configJson: JSON.stringify({ 
            userId: 'U1',
            userEmail: 'admin@example.com',
            appPublished: false,
            isPublic: false,
            allowAnonymous: false
          }),
        })),
      },
      User: {
        email: jest.fn(() => 'admin@example.com'),
        info: jest.fn(() => ({
          userId: 'U1',
          userEmail: 'admin@example.com',
          spreadsheetId: 'testSpreadsheetId',
        })),
      },
      clearExecutionUserInfoCache: jest.fn(),
      Session: { getActiveUser: () => ({ getEmail: () => 'admin@example.com' }) },
      findUserById: jest.fn(() => ({
        userEmail: 'admin@example.com',
        configJson: JSON.stringify({ 
          userId: 'U1',
          userEmail: 'admin@example.com',
          appPublished: false,
          isPublic: false,
          allowAnonymous: false
        }),
      })),
    };
    vm.createContext(context);
    vm.runInContext(baseCode, context);
    vm.runInContext(appCode, context);
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
    context.DB.findUserById = jest.fn(() => ({
      userEmail: 'admin@example.com',
      configJson: JSON.stringify({ 
        userId: 'U1',
        userEmail: 'admin@example.com',
        appPublished: true, 
        isPublic: true,
        allowAnonymous: true
      }),
    }));
    
    // PropertiesServiceも更新
    context.PropertiesService.getScriptProperties = () => ({
      getProperty: jest.fn((key) => {
        if (key === 'user_config_U1') {
          return JSON.stringify({
            userId: 'U1',
            userEmail: 'admin@example.com',
            appPublished: true,
            isPublic: true,
            allowAnonymous: true
          });
        }
        if (key === 'ADMIN_EMAIL') {
          return 'admin@example.com';
        }
        return null;
      })
    });

    // 他のユーザーからのアクセス
    context.Session.getActiveUser = () => ({ getEmail: () => 'other@example.com' });
    expect(() => context.verifyUserAccess('U1')).not.toThrow();
  });
});