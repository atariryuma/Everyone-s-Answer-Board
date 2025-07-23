const fs = require('fs');
const vm = require('vm');

describe('getWebAppBaseUrl caching', () => {
  const mainCode = fs.readFileSync('src/main.gs', 'utf8');
  const urlCode = fs.readFileSync('src/url.gs', 'utf8');
  let context;

  beforeEach(() => {
    const store = {};
    context = {
      cacheManager: {
        store,
        get(key, fn) {
          if (this.store[key]) return this.store[key];
          const val = fn();
          this.store[key] = val;
          return val;
        },
        remove(key) { delete this.store[key]; }
      },
      ScriptApp: {
        getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/ID/exec' }),
        getScriptId: () => 'ID'
      },
      console: { warn: () => {}, log: () => {}, error: () => {} },
      PropertiesService: {
        getScriptProperties: () => ({
          getProperty: (key) => {
            if (key === 'DEBUG_MODE') return 'false';
            return null;
          }
        })
      },

      CacheService: (() => {
        const store = { webapp_base_url: 'https://script.google.com/macros/s/ID/exec' };
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


  test('returns the correct web app base URL and caches it', () => {
    // Clear cache to ensure it's fetched
    context.cacheManager.remove('webapp_base_url');
    
    // First call fetches and caches
    const url1 = context.getWebAppBaseUrl();
    expect(url1).toBe('https://script.google.com/macros/s/ID/exec');

    // Change the underlying service URL, but cache should still return old value
    context.ScriptApp.getService = () => ({ getUrl: () => 'https://script.google.com/macros/s/NEW_ID/exec' });
    const url2 = context.getWebAppBaseUrl();
    expect(url2).toBe('https://script.google.com/macros/s/ID/exec'); // Still the cached value

    // Clear cache and call again, should get new URL
    context.cacheManager.remove('webapp_base_url');
    const url3 = context.getWebAppBaseUrl();
    expect(url3).toBe('https://script.google.com/macros/s/NEW_ID/exec');
  });
});