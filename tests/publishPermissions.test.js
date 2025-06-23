const gas = require('../src/Code.gs');
const { publishApp, unpublishApp } = gas;

function setup(userEmail, adminEmails) {
  const props = {
    getProperty: (key) => {
      if (key === 'ADMIN_EMAILS') return adminEmails.join(',');
      if (key === 'USER_DB_ID') return 'db1';
      return null;
    },
    setProperty: jest.fn()
  };
  global.PropertiesService = {
    getScriptProperties: () => props,
    getUserProperties: () => ({
      getProperty: jest.fn(() => 'u1'),
      setProperty: jest.fn(),
      setProperties: jest.fn()
    })
  };
  global.getUserInfo = jest.fn(() => ({
    adminEmail: adminEmails[0],
    spreadsheetId: 'id1'
  }));
  global.getUserDatabase = jest.fn(() => ({
    getDataRange: () => ({
      getValues: () => [
        ['userId','spreadsheetId','adminEmail','configJson','lastAccessedAt'],
        ['u1','id1','admin@example.com', JSON.stringify({ isPublished: true, sheetName: 'Sheet1' }), '']
      ]
    }),
    getRange: jest.fn(() => ({ setValue: jest.fn() }))
  }));
  global.Session = { getActiveUser: () => ({ getEmail: () => userEmail }) };
  const sheet = {
    getLastColumn: () => 2,
    getRange: () => ({ getValues: () => [['a','b']], setValue: jest.fn(), setValues: jest.fn() }),
    getDataRange: () => ({ getValues: () => [
      ['userId','spreadsheetId','adminEmail','configJson','lastAccessedAt'],
      ['u1','id1','admin@example.com', JSON.stringify({ isPublished: true, sheetName: 'Sheet1' }), '']
    ] }),
    insertColumnAfter: jest.fn(),
    setName: jest.fn()
  };
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({ getSheetByName: () => sheet }),
    openById: jest.fn(() => ({ getSheetByName: () => sheet })),
    create: jest.fn(() => ({
      getActiveSheet: () => sheet,
      getId: () => 'newid'
    }))
  };
  global.getCurrentSpreadsheet = () => global.SpreadsheetApp.getActiveSpreadsheet();
  return props;
}

afterEach(() => {
  delete global.PropertiesService;
  delete global.Session;
  delete global.SpreadsheetApp;
  delete global.getUserDatabase;
  delete global.getCurrentSpreadsheet;
  jest.restoreAllMocks();
});

test('publishApp succeeds for admin user', () => {
  const props = setup('admin@example.com', ['admin@example.com']);
  const msg = publishApp('Sheet1');
  expect(msg).toBe('「Sheet1」を公開しました。');
});

test('publishApp throws for non-admin user', () => {
  setup('user@example.com', ['admin@example.com']);
  expect(() => publishApp('Sheet1')).toThrow('権限');
});

test('unpublishApp succeeds for admin user', () => {
  const props = setup('admin@example.com', ['admin@example.com']);
  const msg = unpublishApp();
  expect(props.setProperty).toHaveBeenCalledWith('IS_PUBLISHED', 'false');
  expect(msg).toBe('アプリを非公開にしました。');
});

test('unpublishApp throws for non-admin user', () => {
  setup('user@example.com', ['admin@example.com']);
  expect(() => unpublishApp()).toThrow('権限');
});
