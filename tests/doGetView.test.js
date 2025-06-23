const { doGet } = require('../src/Code.gs');

afterEach(() => {
  delete global.HtmlService;
  delete global.getActiveUserEmail;
  delete global.Session;
  delete global.PropertiesService;
  delete global.SpreadsheetApp;
  delete global.getUserInfo;
});

function setup({ userEmail = 'admin@example.com', adminEmails = 'admin@example.com' }) {
  global.PropertiesService = {
    getScriptProperties: () => ({
      getProperty: (key) => {
        if (key === 'ADMIN_EMAILS') return adminEmails;
        if (key === 'USER_DB_ID') return 'db';
        return null;
      },
      setProperty: jest.fn()
    }),
    getUserProperties: () => ({
      getProperty: (key) => key === 'CURRENT_USER_ID' ? 'user1' : null
    })
  };
  global.getUserInfo = () => ({
    adminEmail: 'admin@example.com',
    spreadsheetId: 'id123',
    configJson: { isPublished: true, sheetName: 'Sheet1', showDetails: false }
  });
  global.getActiveUserEmail = () => userEmail;
  global.Session = { getActiveUser: () => ({ getEmail: () => userEmail }) };
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({
      getSheets: () => [
        { getName: () => 'Sheet1', isSheetHidden: () => false }
      ]
    }),
    openById: () => ({ getSheetByName: () => ({}) }),
    create: () => ({
      getActiveSheet: () => ({
        setName: jest.fn(),
        getRange: () => ({ setValues: jest.fn() })
      }),
      getId: () => 'db'
    })
  };
  const output = { setTitle: jest.fn(() => output), addMetaTag: jest.fn(() => output) };
  let template = { evaluate: () => output };
  global.HtmlService = {
    createTemplateFromFile: jest.fn(() => template)
  };
  return { getTemplate: () => template };
}

test('page parameter is ignored and admin mode is based on user role', () => {
  const { getTemplate } = setup({});
  doGet({ parameter: { page: 'admin', userId: 'user1' } });
  const tpl = getTemplate();
  expect(tpl.isAdminUser).toBe(true);
  expect(tpl.showAdminFeatures).toBe(false);
});

test('admin user starts in admin mode', () => {
  const { getTemplate } = setup({});
  doGet({ parameter: { userId: 'user1' } });
  const tpl = getTemplate();
  expect(tpl.isAdminUser).toBe(true);
  expect(tpl.showAdminFeatures).toBe(false);
});

test('non admin user starts in viewer mode', () => {
  const { getTemplate } = setup({ userEmail: 'user@example.com' });
  doGet({ parameter: { userId: 'user1' } });
  const tpl = getTemplate();
  expect(tpl.isAdminUser).toBe(false);
  expect(tpl.showAdminFeatures).toBe(false);
});
