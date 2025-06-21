const { doGet } = require('../src/Code.gs');

afterEach(() => {
  delete global.PropertiesService;
  delete global.Session;
  delete global.HtmlService;
  delete global.SpreadsheetApp;
  delete global.getConfig;
});

function setup() {
  global.PropertiesService = {
    getScriptProperties: () => ({
      getProperty: (key) => {
        if (key === 'IS_PUBLISHED') return 'false';
        return null;
      }
    })
  };
  global.getConfig = jest.fn(() => ({ questionHeader: 'Q', answerHeader: 'A', reasonHeader: 'R', nameMode: '同一シート' }));
  global.Session = { getActiveUser: () => ({ getEmail: () => { throw new Error('no user'); } }) };
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({
      getSheets: () => [
        { getName: () => 'Sheet1', isSheetHidden: () => false }
      ]
    })
  };
  const output = { setTitle: jest.fn(() => output) };
  let template = { evaluate: () => output };
  global.HtmlService = {
    createTemplateFromFile: jest.fn(() => template)
  };
  return template;
}

test('doGet uses Unpublished template when email fails', () => {
  const template = setup();
  doGet({});
  expect(HtmlService.createTemplateFromFile).toHaveBeenCalledWith('Unpublished');
  expect(template.userEmail).toBe('');
});
