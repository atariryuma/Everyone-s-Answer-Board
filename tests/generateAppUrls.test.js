const fs = require('fs');
const vm = require('vm');

describe('generateAppUrls admin url', () => {
  const mainCode = fs.readFileSync('src/main.gs', 'utf8');
  const urlCode = fs.readFileSync('src/url.gs', 'utf8');
  let context;

  beforeEach(() => {
    const store = {};
    context = {
      cacheManager: {
        store,
        get(key, fn) { return fn(); },
        remove() {}
      },
      ScriptApp: {
        getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/ID/dev' }),
        getScriptId: () => 'ID'
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
        }),
        getScriptCache: () => ({
          get: (key) => store[key] || null,
          put: (key, val) => { store[key] = val; },
          remove: (key) => { delete store[key]; }
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
    vm.runInContext(urlCode, context);
  });

  test('returns adminUrl with userId and mode parameter', () => {
    const urls = context.generateAppUrls('abc');
    expect(urls.adminUrl).toBe('https://script.google.com/macros/s/ID/exec?userId=abc&mode=admin');
  });
});
