const fs = require('fs');
const vm = require('vm');

describe('getWebAppUrlCached upgrade', () => {
  const mainCode = fs.readFileSync('src/main.gs', 'utf8');
  let context;

  beforeEach(() => {
    context = {
      cacheManager: {
        store: {},
        get(key, fn) {
          if (this.store[key]) return this.store[key];
          const val = fn();
          this.store[key] = val;
          return val;
        },
        remove(key) { delete this.store[key]; }
      },
      ScriptApp: {
        getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/ID/dev' }),
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

  test('updates cached dev url when production url available', () => {
    expect(context.getWebAppUrlCached()).toMatch('/dev');
    context.ScriptApp.getService = () => ({ getUrl: () => 'https://script.google.com/macros/s/ID/exec' });
    expect(context.getWebAppUrlCached()).toMatch('/exec');
  });

  test('normalizes domain placement in url', () => {
    context.ScriptApp.getService = () => ({ getUrl: () => 'https://script.google.com/a/example.com/macros/s/ID/exec' });
    const normalized = context.computeWebAppUrl();
    expect(normalized).toBe('https://script.google.com/a/macros/example.com/s/ID/exec');
  });

  test('updates cache when domain format differs', () => {
    context.cacheManager.store['WEB_APP_URL'] = 'https://script.google.com/a/example.com/macros/s/ID/exec';
    context.ScriptApp.getService = () => ({ getUrl: () => 'https://script.google.com/a/macros/example.com/s/ID/exec' });
    const updated = context.getWebAppUrlCached();
    expect(updated).toBe('https://script.google.com/a/macros/example.com/s/ID/exec');
  });
});
