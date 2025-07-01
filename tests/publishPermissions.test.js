

function setup(userEmail, adminEmails) {
  const props = {
    getProperty: (key) => {
      if (key === 'ADMIN_EMAILS') return adminEmails.join(',');
      if (key === 'DATABASE_ID') return 'db1';
      return null;
    },
    setProperty: jest.fn()
  };
  global.PropertiesService = {
    getScriptProperties: () => props,
    getUserProperties: () => ({
      getProperty: jest.fn(() => 'user1234567'),
      setProperty: jest.fn(),
      setProperties: jest.fn()
    })
  };
  global.getUserInfo = jest.fn(() => ({
    adminEmail: adminEmails[0],
    spreadsheetId: 'id1'
  }));
  global.getDatabase = jest.fn(() => ({
    getSheetByName: () => ({
      getDataRange: () => ({
        getValues: () => [
          ['userId','spreadsheetId','adminEmail','configJson','lastAccessedAt'],
          ['user1234567','id1','admin@example.com', JSON.stringify({ isPublished: true, sheetName: 'Sheet1' }), '']
        ]
      }),
      getRange: jest.fn(() => ({ setValue: jest.fn() }))
    })
  }));
  global.DriveApp = { getFileById: jest.fn(() => ({ setSharing: jest.fn() })) };
  global.Session = { getActiveUser: () => ({ getEmail: () => userEmail }) };
  global.publishApp = jest.fn();
  global.unpublishApp = jest.fn();
  global.checkAdmin = jest.fn(() => userEmail.includes('admin'));
  const sheet = {
    getLastColumn: () => 2,
    getRange: () => ({ getValues: () => [['a','b']], setValue: jest.fn(), setValues: jest.fn() }),
    getDataRange: () => ({ getValues: () => [
      ['userId','spreadsheetId','adminEmail','configJson','lastAccessedAt'],
      ['user1234567','id1','admin@example.com', JSON.stringify({ isPublished: true, sheetName: 'Sheet1' }), '']
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
  delete global.DriveApp;
  delete global.getCurrentSpreadsheet;
  jest.restoreAllMocks();
});

test('publishApp succeeds for admin user', () => {
  const props = setup('admin@example.com', ['admin@example.com']);
  global.publishApp.mockImplementation((sheetName) => {
    if (!global.checkAdmin()) {
      throw new Error('権限がありません。');
    }
    props.setProperty('IS_PUBLISHED', 'true');
    props.setProperty('PUBLISHED_SHEET_NAME', sheetName);
    return `「${sheetName}」を公開しました。`;
  });
  const msg = publishApp('Sheet1');
  expect(msg).toBe('「Sheet1」を公開しました。');
});

test('publishApp throws for non-admin user', () => {
  setup('user@example.com', ['admin@example.com']);
  global.publishApp.mockImplementation((sheetName) => {
    if (!global.checkAdmin()) {
      throw new Error('権限がありません。');
    }
    props.setProperty('IS_PUBLISHED', 'true');
    props.setProperty('PUBLISHED_SHEET_NAME', sheetName);
    return `「${sheetName}」を公開しました。`;
  });
  expect(() => publishApp('Sheet1')).toThrow('権限');
});

test('unpublishApp succeeds for admin user', () => {
  const props = setup('admin@example.com', ['admin@example.com']);
  global.unpublishApp.mockImplementation(() => {
    if (!global.checkAdmin()) {
      throw new Error('権限がありません。');
    }
    props.setProperty('IS_PUBLISHED', 'false');
    if (props.deleteProperty) props.deleteProperty('PUBLISHED_SHEET_NAME');
    return 'アプリを非公開にしました。';
  });
  const msg = unpublishApp();
  expect(props.setProperty).toHaveBeenCalledWith('IS_PUBLISHED', 'false');
  expect(msg).toBe('アプリを非公開にしました。');
});

test('unpublishApp throws for non-admin user', () => {
  setup('user@example.com', ['admin@example.com']);
  global.unpublishApp.mockImplementation(() => {
    if (!global.checkAdmin()) {
      throw new Error('権限がありません。');
    }
    props.setProperty('IS_PUBLISHED', 'false');
    if (props.deleteProperty) props.deleteProperty('PUBLISHED_SHEET_NAME');
    return 'アプリを非公開にしました。';
  });
  expect(() => unpublishApp()).toThrow('権限');
});
