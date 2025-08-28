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
        getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/ID/exec' }),
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

      CacheService: (() => {
        const store = {};
        return {
          getUserCache: () => ({
            get: () => null,
            put: () => {}
          }),
          getScriptCache: () => ({
            get: (k) => store[k] || null,
            put: (k, v) => { store[k] = v; },
            remove: (k) => { delete store[k]; }
          })
        };
      })(),
        
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
    vm.runInContext(urlCode, context);
    vm.runInContext(mainCode, context);
    vm.runInContext(urlCode, context);
  });

  test('returns adminUrl with userId and mode parameter', () => {
    const urls = context.generateUserUrls('abc');
    expect(urls.adminUrl).toBe('https://script.google.com/macros/s/ID/exec?mode=admin&userId=abc');
  });
});
