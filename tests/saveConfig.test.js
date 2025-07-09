const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { JSDOM } = require('jsdom');

describe('AdminPanel saveConfig', () => {
  test('calls google.script.run.saveConfig and stores pending config', () => {
    const htmlPath = path.resolve(__dirname, '../src/AdminPanel.html');
    const html = fs.readFileSync(htmlPath, 'utf8');

    const dom = new JSDOM(html, { runScripts: "dangerously", resources: "usable" });

    const ctx = {
      window: dom.window,
      document: dom.window.document,
      google: {
        script: {
          run: {
            withSuccessHandler: function(fn) { this.success = fn; return this; },
            withFailureHandler: function(fn) { this.failure = fn; return this; },
            saveConfig: jest.fn(function (sheetName, cfg) { if (this.success) this.success('ok'); }),
          }
        }
      },
      showMessage: jest.fn(),
      showError: jest.fn(),
      setLoading: jest.fn(),
      showPublishModal: jest.fn(),
      unifiedCache: {
        getOrSet: (key, fn) => fn()
      },
      GASOptimizer: class {
        call() {
          return Promise.resolve();
        }
      }
    };

    vm.createContext(ctx);
    
    const scriptTags = dom.window.document.getElementsByTagName('script');
    for (let i = 0; i < scriptTags.length; i++) {
      vm.runInContext(scriptTags[i].textContent, ctx);
    }

    // Set up the DOM for the test
    const sheetSelect = dom.window.document.getElementById('sheet-select');
    const option = dom.window.document.createElement('option');
    option.value = 'Sheet1';
    sheetSelect.appendChild(option);
    sheetSelect.value = 'Sheet1';
    
    dom.window.document.getElementById('opinionHeader').value = 'opinion';
    dom.window.document.getElementById('reasonHeader').value = 'reason';
    dom.window.document.getElementById('nameHeader').value = 'name';
    dom.window.document.getElementById('classHeader').value = 'class';
    dom.window.document.getElementById('showNames').checked = false;
    dom.window.document.getElementById('showCounts').checked = true;

    // Mock the necessary functions and objects that are called within saveConfig
    ctx.selectedSheet = 'Sheet1';
    ctx.currentStatus = { userInfo: { spreadsheetId: 'id123' } };
    ctx.elements = {
        sheetSelect: sheetSelect,
        opinionHeader: dom.window.document.getElementById('opinionHeader'),
        reasonHeader: dom.window.document.getElementById('reasonHeader'),
        nameHeader: dom.window.document.getElementById('nameHeader'),
        classHeader: dom.window.document.getElementById('classHeader'),
        showNames: dom.window.document.getElementById('showNames'),
        showCounts: dom.window.document.getElementById('showCounts'),
        saveConfigBtn: dom.window.document.getElementById('save-config-btn'),
        saveConfigText: { textContent: '' },
    };
    ctx.showConfirmationModal = (title, msg, callback) => callback();


    // Trigger the click event
    dom.window.document.getElementById('save-config-btn').click();

    expect(ctx.google.script.run.saveConfig).toHaveBeenCalledWith('Sheet1', {
      opinionHeader: 'opinion',
      reasonHeader: 'reason',
      nameHeader: 'name',
      classHeader: 'class',
      showNames: false,
      showCounts: true,
    });
  });
});