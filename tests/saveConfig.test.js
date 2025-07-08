const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { JSDOM } = require('jsdom');

function extractSaveConfigScript(html) {
  const start = html.indexOf('function saveConfig()');
  if (start === -1) return null;
  let depth = 0;
  let end = start;
  for (; end < html.length; end++) {
    const c = html[end];
    if (c === '{') depth++;
    else if (c === '}') {
      depth--;
      if (depth === 0) { end++; break; }
    }
  }
  return html.slice(start, end);
}

describe('AdminPanel saveConfig', () => {
  test('calls google.script.run.saveConfig and stores pending config', () => {
    const jsPath = path.resolve(__dirname, '../src/client/scripts/AdminPanel.js.html');
    const js = fs.readFileSync(jsPath, 'utf8');
    const script = extractSaveConfigScript(js);
    expect(script).toBeTruthy();

    const dom = new JSDOM(`<!DOCTYPE html><body>
      <select id="sheet-select"><option value="Sheet1"></option></select>
      <input id="opinionHeader" value="opinion" />
      <input id="reasonHeader" value="reason" />
      <input id="nameHeader" value="name" />
      <input id="classHeader" value="class" />
      <input type="checkbox" id="showNames" />
      <input type="checkbox" id="showCounts" checked />
    </body>`);

    const ctx = {
      window: dom.window,
      document: dom.window.document,
      elements: {
        sheetSelect: dom.window.document.getElementById('sheet-select'),
        opinionHeader: dom.window.document.getElementById('opinionHeader'),
        reasonHeader: dom.window.document.getElementById('reasonHeader'),
        nameHeader: dom.window.document.getElementById('nameHeader'),
        classHeader: dom.window.document.getElementById('classHeader'),
        saveConfigBtn: {},
        saveConfigText: { textContent: '' },
      },
      currentStatus: { userInfo: { spreadsheetId: 'id123' } },
      showMessage: jest.fn(),
      addVisualFeedback: jest.fn(),
      showError: jest.fn(),
      setLoading: jest.fn(),
      showPublishModal: jest.fn(),
    };

    const runStub = {
      success: null,
      failure: null,
      withSuccessHandler(fn) { this.success = fn; return this; },
      withFailureHandler(fn) { this.failure = fn; return this; },
      saveConfig: jest.fn(function (sheetName, cfg) { if (this.success) this.success('ok'); }),
    };
    ctx.google = { script: { run: runStub } };

    vm.createContext(ctx);
    vm.runInContext(script, ctx);
    ctx.saveConfig();

    expect(runStub.saveConfig).toHaveBeenCalledWith('Sheet1', {
      opinionHeader: 'opinion',
      reasonHeader: 'reason',
      nameHeader: 'name',
      classHeader: 'class',
      showNames: false,
      showCounts: true,
    });

    expect(ctx.window.pendingConfig).toEqual({
      spreadsheetId: 'id123',
      sheetName: 'Sheet1',
      config: {
        opinionHeader: 'opinion',
        reasonHeader: 'reason',
        nameHeader: 'name',
        classHeader: 'class',
        showNames: false,
        showCounts: true,
      },
    });

    expect(ctx.showPublishModal).toHaveBeenCalled();
  });
});
