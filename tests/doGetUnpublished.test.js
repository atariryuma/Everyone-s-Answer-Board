const { doGet } = require('../src/Code.gs');

afterEach(() => {
  delete global.PropertiesService;
  delete global.Session;
  delete global.HtmlService;
  delete global.SpreadsheetApp;
  delete global.getCurrentSpreadsheet;
});

function setup() {
  global.PropertiesService = {
    getScriptProperties: () => ({
      getProperty: (key) => {
        if (key === 'IS_PUBLISHED') return 'false';
        if (key === 'USER_DB_ID') return 'db1';
        return null;
      }
    }),
    getUserProperties: () => ({
      getProperty: jest.fn(),
      setProperty: jest.fn(),
      setProperties: jest.fn()
    })
  };
  global.Session = { getActiveUser: () => ({ getEmail: () => { throw new Error('no user'); } }) };
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({
      getSheets: () => [
        { getName: () => 'Sheet1', isSheetHidden: () => false }
      ]
    }),
    create: jest.fn(() => ({
      getActiveSheet: () => ({
        setName: jest.fn(),
        getRange: jest.fn(() => ({ setValues: jest.fn() }))
      }),
      getId: () => 'id1'
    }))
  };
  global.getCurrentSpreadsheet = () => global.SpreadsheetApp.getActiveSpreadsheet();
  const output = { setTitle: jest.fn(() => output), addMetaTag: jest.fn(() => output) };
  let template = { evaluate: () => output };
  global.HtmlService = {
    createTemplateFromFile: jest.fn(() => template)
  };
  const dbSheet = {
    getDataRange: () => ({
      getValues: () => [
        ['userId','spreadsheetId','adminEmail','configJson','lastAccessedAt'],
        ['u1','id1','admin@example.com', JSON.stringify({ isPublished: false, sheetName: 'Sheet1' }), '']
      ]
    }),
    getRange: jest.fn(() => ({ setValue: jest.fn() }))
  };
  global.SpreadsheetApp.openById = jest.fn(() => ({ getSheetByName: () => dbSheet }));
  return template;
}

test('doGet uses Unpublished template when email fails', () => {
  const template = setup();
  doGet({ parameter: { userId: 'u1' } });
  expect(HtmlService.createTemplateFromFile).toHaveBeenCalledWith('Unpublished');
  expect(template.userEmail).toBe('admin@example.com');
});
