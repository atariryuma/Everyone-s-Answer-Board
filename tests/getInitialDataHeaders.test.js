const fs = require('fs');
const vm = require('vm');

describe('getInitialData header extraction', () => {
  const urlCode = fs.readFileSync('src/url.gs', 'utf8');
  const mainCode = fs.readFileSync('src/main.gs', 'utf8');
  const coreCode = fs.readFileSync('src/Core.gs', 'utf8');
  let context;

  beforeEach(() => {
    const store = {};
    context = {
      debugLog: () => {},
      console,
      PropertiesService: {
        getScriptProperties: () => ({ getProperty: () => null }),
      },
      CacheService: {
        getUserCache: () => ({ get: () => null, put: () => {} }),
        getScriptCache: () => ({ get: (k) => store[k] || null, put: (k, v) => { store[k] = v; }, remove: (k) => { delete store[k]; } })
      },
      cacheManager: {
        store,
        get(key, fn) {
          if (this.store[key]) return this.store[key];
          const val = fn();
          this.store[key] = val;
          return val;
        },
        remove(key) { delete this.store[key]; },
        clearByPattern: (pattern) => {
          for (const key in store) {
            if (key.startsWith(pattern.replace('*', ''))) {
              delete store[key];
            }
          }
        }
      },
      ScriptApp: {
        getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/ID/exec' }),
        getScriptId: () => 'ID'
      },
      Session: { getActiveUser: () => ({ getEmail: () => 'test@example.com' }) },
      Utilities: { getUuid: () => 'uid', computeDigest: () => [], Charset: { UTF_8: 'UTF-8' } },
    };
    vm.createContext(context);
    vm.runInContext(urlCode, context);
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
      generateUserUrls: jest.fn(() => ({ webApp: 'web', adminUrl: 'admin', viewUrl: 'view' })),
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
