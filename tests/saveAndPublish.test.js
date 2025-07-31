const fs = require('fs');
const vm = require('vm');

describe('saveAndPublish with mocked getCachedUserInfo', () => {
  const code = fs.readFileSync('src/config.gs', 'utf8');
  let context;

  beforeEach(() => {
    context = {
      console,
      debugLog: () => {},
      errorLog: jest.fn(), // Add errorLog mock
      infoLog: jest.fn(),  // Add infoLog mock
      LockService: {
        getScriptLock: () => ({
          waitLock: jest.fn(),
          releaseLock: jest.fn(),
        }),
      },
    };
    vm.createContext(context);
    vm.runInContext(code, context);

    Object.assign(context, {
      verifyUserAccess: jest.fn(),
      saveSheetConfigInContext: jest.fn(),
      switchToSheetInContext: jest.fn(),
      setDisplayOptionsInContext: jest.fn(),
      commitAllChanges: jest.fn(),
      synchronizeCacheAfterCriticalUpdate: jest.fn(),
      findUserByIdFresh: jest.fn(() => ({
        userId: 'U',
        adminEmail: 'admin@example.com',
        spreadsheetId: 'SS',
      })),
      buildResponseFromContext: jest.fn(() => ({ _meta: {} })),
    });
  });

  test('returns response when user info exists', () => {
    context.getCachedUserInfo = jest.fn(() => ({
      userId: 'U',
      adminEmail: 'admin@example.com',
      spreadsheetId: 'SS',
    }));
    context.createExecutionContext = jest.fn((uid) => ({
      requestUserId: uid,
      userInfo: context.getCachedUserInfo(uid),
      stats: { sheetsServiceCreations: 0, dbQueries: 0, operationsCount: 0 },
    }));

    const res = context.saveAndPublish('U', 'Sheet1', {});
    expect(res._meta).toBeDefined();
  });

  test('throws error when user info missing', () => {
    context.getCachedUserInfo = jest.fn(() => null);
    context.createExecutionContext = jest.fn((uid) => {
      const userInfo = context.getCachedUserInfo(uid);
      if (!userInfo) throw new Error('ユーザー情報が取得できません');
      return { requestUserId: uid, userInfo, stats: { sheetsServiceCreations: 0, dbQueries: 0, operationsCount: 0 } };
    });

    expect(() => context.saveAndPublish('U', 'Sheet1', {})).toThrow(
      '設定の保存と公開中にサーバーエラーが発生しました: ユーザー情報が取得できません',
    );
  });
});
