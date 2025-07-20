const fs = require('fs');
const vm = require('vm');

describe('getInitialData header extraction', () => {
  const coreCode = fs.readFileSync('src/Core.gs', 'utf8');
  const mainCode = fs.readFileSync('src/main.gs', 'utf8');
  let context;

  beforeEach(() => {
    context = {
      debugLog: () => {},
      console,
      PropertiesService: {
        getScriptProperties: () => ({ getProperty: () => null }),
      },
      CacheService: {
        getUserCache: () => ({ get: () => null, put: () => {} }),
      },
      Session: { getActiveUser: () => ({ getEmail: () => 'test@example.com' }) },
      Utilities: { getUuid: () => 'uid', computeDigest: () => [], Charset: { UTF_8: 'UTF-8' } },
    };
    vm.createContext(context);
    vm.runInContext(mainCode, context);
    vm.runInContext(coreCode, context);

    Object.assign(context, {
      verifyUserAccess: jest.fn(),
      getCachedUserInfo: jest.fn(() => ({
        userId: 'U',
        adminEmail: 'test@example.com',
        spreadsheetId: '',
        spreadsheetUrl: '',
        configJson: JSON.stringify({
          publishedSheetName: 'ClassA',
          sheet_ClassA: {
            opinionHeader: 'テーマ',
            nameHeader: '氏名',
            classHeader: 'クラス',
          },
        }),
      })),
      getSheetsList: jest.fn(() => []),
      generateAppUrls: jest.fn(() => ({ webApp: 'web', adminUrl: 'admin', viewUrl: 'view' })),
      getResponsesData: jest.fn(() => ({ status: 'success', data: [] })),
      determineSetupStep: jest.fn(() => 3),
    });
  });

  test('returns headers from nested sheet config', () => {
    const res = context.getInitialData('U');
    expect(res.config.opinionHeader).toBe('テーマ');
    expect(res.config.nameHeader).toBe('氏名');
    expect(res.config.classHeader).toBe('クラス');
  });
});
