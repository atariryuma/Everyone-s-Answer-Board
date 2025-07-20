const fs = require('fs');
const vm = require('vm');

describe('generateAppUrls admin url', () => {
  const mainCode = fs.readFileSync('src/main.gs', 'utf8');
  let context;

  beforeEach(() => {
    context = {
      cacheManager: {
        store: {},
        get(key, fn) { return fn(); },
        remove() {}
      },
      ScriptApp: {
        getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/ID/dev' })
      },
      console: { error: () => {}, log: () => {}, warn: () => {} },
      PropertiesService: {
        getScriptProperties: () => ({
          getProperty: (key) => {
            if (key === 'DEBUG_MODE') return 'false';
            return null;
          }
        })
      },
      CacheService: {
        getUserCache: () => ({
          get: () => null,
          put: () => {}
        })
      },
      Session: {
        getActiveUser: () => ({ getEmail: () => 'test@example.com' })
      },
      Utilities: {
        getUuid: () => 'mock-uuid',
        computeDigest: () => [],
        Charset: { UTF_8: 'UTF-8' }
      }
    };
    vm.createContext(context);
    vm.runInContext(mainCode, context);
  });

  test('returns adminUrl with userId and mode parameter', () => {
    const urls = context.generateAppUrls('abc');
    expect(urls.adminUrl).toBe('https://script.google.com/macros/s/ID/exec?userId=abc&mode=admin');
  });
});
