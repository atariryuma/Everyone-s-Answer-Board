const fs = require('fs');
const vm = require('vm');

describe('getWebAppUrl', () => {
  const urlCode = fs.readFileSync('src/url.gs', 'utf8');

  test('retrieves URL via Apps Script API', () => {
    const context = {
      ScriptApp: { getScriptId: () => 'ID' },
      AppsScript: {
        Script: {
          Deployments: {
            list: () => ({
              deployments: [
                { deploymentConfig: { webApp: { url: 'https://script.google.com/macros/s/ID/exec' } } }
              ]
            })
          }
        }
      },
      console: { log: () => {}, error: () => {}, warn: () => {} },
      debugLog: () => {},
      errorLog: () => {},
      warnLog: () => {},
      infoLog: () => {},
      CacheService: {
        getScriptCache: () => ({
          get: () => null,
          put: () => {}
        })
      }
    };
    vm.createContext(context);
    vm.runInContext(urlCode, context);
    expect(context.getWebAppUrl()).toBe('https://script.google.com/macros/s/ID/exec');
  });

  test('falls back to ScriptApp service URL when API fails', () => {
    const context = {
      ScriptApp: {
        getScriptId: () => 'ID',
        getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/ID/exec' })
      },
      AppsScript: {
        Script: {
          Deployments: {
            list: () => { throw new Error('API disabled'); }
          }
        }
      },
      console: { log: () => {}, error: () => {}, warn: () => {} },
      debugLog: () => {},
      errorLog: () => {},
      warnLog: () => {},
      infoLog: () => {},
      CacheService: {
        getScriptCache: () => ({
          get: () => null,
          put: () => {}
        })
      }
    };
    vm.createContext(context);
    vm.runInContext(urlCode, context);
    expect(context.getWebAppUrl()).toBe('https://script.google.com/macros/s/ID/exec');
  });
});
