const { doGet } = require('../src/Code.gs');

afterEach(() => {
  delete global.HtmlService;
  delete global.getActiveUserEmail;
  delete global.Session;
  delete global.PropertiesService;
  delete global.SpreadsheetApp;
  delete global.getCurrentSpreadsheet;
});

function setup({ userEmail = 'admin@example.com', adminEmails = 'admin@example.com' }) {
  global.PropertiesService = {
    getScriptProperties: () => ({
      getProperty: (key) => {
        if (key === 'ADMIN_EMAILS') return adminEmails;
        if (key === 'USER_DB_ID') return 'db1';
        return null;
      },
      setProperty: jest.fn()
    }),
    getUserProperties: () => ({
      getProperty: jest.fn(),
      setProperty: jest.fn(),
      setProperties: jest.fn()
    })
  };
  const dbSheet = {
    getDataRange: () => ({
      getValues: () => [
        ['userId','spreadsheetId','adminEmail','configJson','lastAccessedAt'],
        ['u1','id1','admin@example.com', JSON.stringify({ isPublished: true, sheetName: 'Sheet1' }), '']
      ]
    }),
    getRange: jest.fn(() => ({ setValue: jest.fn() }))
  };
  global.getActiveUserEmail = () => userEmail;
  global.Session = { getActiveUser: () => ({ getEmail: () => userEmail }) };
  global.DriveApp = { getFileById: jest.fn(() => ({ setSharing: jest.fn() })) };
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({
      getSheets: () => [
        { getName: () => 'Sheet1', isSheetHidden: () => false }
      ]
    }),
    openById: jest.fn(() => ({ getSheetByName: () => dbSheet })),
    create: jest.fn(() => ({
      getActiveSheet: () => ({
        setName: jest.fn(),
        getRange: jest.fn(() => ({ setValues: jest.fn() }))
      }),
      getId: () => 'id1'
    }))
  };
  global.getCurrentSpreadsheet = () => global.SpreadsheetApp.getActiveSpreadsheet();
  const output = { setTitle: jest.fn(() => output), addMetaTag: jest.fn(() => output), setSandboxMode: jest.fn(() => output) };
  let template = { evaluate: () => output };
  global.HtmlService = {
    createTemplateFromFile: jest.fn(() => template),
    SandboxMode: { IFRAME: 'IFRAME' }
  };
  return { getTemplate: () => template };
}

test('page parameter is ignored and admin mode is based on user role', () => {
  const { getTemplate } = setup({});
  doGet({ parameter: { page: 'admin', userId: 'u1' } });
  const tpl = getTemplate();
  expect(tpl.isAdminUser).toBe(true);
  expect(tpl.showAdminFeatures).toBe(false);
});

test('admin user starts in admin mode', () => {
  const { getTemplate } = setup({});
  doGet({ parameter: { userId: 'u1' } });
  const tpl = getTemplate();
  expect(tpl.isAdminUser).toBe(true);
  expect(tpl.showAdminFeatures).toBe(false);
});

test('non admin user starts in viewer mode', () => {
  const { getTemplate } = setup({ userEmail: 'user@example.com' });
  doGet({ parameter: { userId: 'u1' } });
  const tpl = getTemplate();
  expect(tpl.isAdminUser).toBe(false);
  expect(tpl.showAdminFeatures).toBe(false);
});

test('different domain user receives permission error', () => {
  setup({ userEmail: 'user@other.com' });
  const output = { setTitle: jest.fn(() => output) };
  global.HtmlService.createHtmlOutput = jest.fn(() => output);
  doGet({ parameter: { userId: 'u1' } });
  expect(HtmlService.createHtmlOutput).toHaveBeenCalledWith(expect.stringContaining('権限'));
});
