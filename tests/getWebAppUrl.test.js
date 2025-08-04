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
      console: { log: () => {}, error: () => {}, warn: () => {} }
    };
    vm.createContext(context);
    vm.runInContext(urlCode, context);
    expect(context.getWebAppUrl()).toBe('https://script.google.com/macros/s/ID/exec');
  });
});
