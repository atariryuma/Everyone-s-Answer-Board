const { doGet } = require('../src/Code.gs');

afterEach(() => {
  delete global.PropertiesService;
  delete global.Session;
  delete global.HtmlService;
  delete global.SpreadsheetApp;
  delete global.getUserInfo;
});

function setup() {
  global.PropertiesService = {
    getScriptProperties: () => ({
      getProperty: (key) => {
        if (key === 'IS_PUBLISHED') return 'false';
        if (key === 'USER_DB_ID') return 'db';
        return null;
      },
      setProperty: jest.fn()
    }),
    getUserProperties: () => ({ getProperty: (k) => k === 'CURRENT_USER_ID' ? 'user1' : null })
  };
  global.Session = { getActiveUser: () => ({ getEmail: () => { throw new Error('no user'); } }) };
  global.getUserInfo = () => ({ adminEmail: '', spreadsheetId: 'id123', configJson: { isPublished: false, sheetName: 'Sheet1' } });
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({
      getSheets: () => [
        { getName: () => 'Sheet1', isSheetHidden: () => false }
      ]
    }),
    openById: () => ({ getSheetByName: () => ({}) }),
    create: () => ({
      getActiveSheet: () => ({ setName: jest.fn(), getRange: () => ({ setValues: jest.fn() }) }),
      getId: () => 'db'
    })
  };
  const output = { setTitle: jest.fn(() => output), addMetaTag: jest.fn(() => output) };
  let template = { evaluate: () => output };
  global.HtmlService = {
    createTemplateFromFile: jest.fn(() => template)
  };
  return template;
}

test('doGet uses Unpublished template when email fails', () => {
  const template = setup();
  doGet({ parameter: { userId: 'user1' } });
  expect(HtmlService.createTemplateFromFile).toHaveBeenCalledWith('Unpublished');
  expect(template.userEmail).toBe('');
});
