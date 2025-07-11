const fs = require('fs');
const vm = require('vm');

describe('generateAppUrls admin url', () => {
  const code = fs.readFileSync('src/url.gs', 'utf8');
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
      console: { error: () => {}, log: () => {}, warn: () => {} }
    };
    vm.createContext(context);
    vm.runInContext(code, context);
  });

  test('returns adminUrl with userId and mode parameter', () => {
    const urls = context.generateAppUrls('abc');
    expect(urls.adminUrl).toBe('https://script.google.com/macros/s/ID/exec?userId=abc&mode=admin');
  });
});
