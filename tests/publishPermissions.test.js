const { publishApp, unpublishApp } = require('../src/Code.gs');

function setup(userEmail, adminEmails) {
  const props = {
    getProperty: (key) => key === 'ADMIN_EMAILS' ? adminEmails.join(',') : null,
    setProperty: jest.fn()
  };
  global.PropertiesService = {
    getScriptProperties: () => ({
      ...props,
      getProperty: (key) => {
        if (key === 'ADMIN_EMAILS') return adminEmails.join(',');
        if (key === 'USER_DB_ID') return 'db';
        return null;
      },
      setProperty: jest.fn()
    }),
    getUserProperties: () => ({
      getProperty: (key) => key === 'CURRENT_USER_ID' ? 'user1' : null
    })
  };
  global.getUserInfo = () => ({ adminEmail: userEmail, spreadsheetId: 'id123' });
  global.Session = { getActiveUser: () => ({ getEmail: () => userEmail }) };
  const sheet = {
    getLastColumn: () => 2,
    getRange: () => ({ getValues: () => [['a','b']], setValue: jest.fn() }),
    insertColumnAfter: jest.fn()
  };
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({ getSheetByName: () => sheet }),
    openById: () => ({ getSheetByName: () => sheet }),
    create: () => ({
      getActiveSheet: () => ({
        setName: jest.fn(),
        getRange: () => ({ setValues: jest.fn() })
      }),
      getId: () => 'db'
    })
  };
  return props;
}

afterEach(() => {
  delete global.PropertiesService;
  delete global.Session;
  delete global.SpreadsheetApp;
  delete global.getUserInfo;
});

test('publishApp succeeds for admin user', () => {
  const props = setup('admin@example.com', ['admin@example.com']);
  const msg = publishApp('Sheet1');
  expect(props.setProperty).toHaveBeenCalledWith('IS_PUBLISHED', 'true');
  expect(props.setProperty).toHaveBeenCalledWith('ACTIVE_SHEET_NAME', 'Sheet1');
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
