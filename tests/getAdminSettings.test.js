const { getAdminSettings } = require('../src/Code.gs');
function setup() {
  global.PropertiesService = {
    getScriptProperties: () => ({
      getProperty: (key) => {
        if (key === 'ADMIN_EMAILS') return 'a@example.com,b@example.com';
        if (key === 'DEPLOY_ID') return 'deploy123';
        if (key === 'WEB_APP_URL') return 'https://example.com/app';
        if (key === 'IS_PUBLISHED') return 'true';
        if (key === 'ACTIVE_SHEET_NAME') return 'SheetA';
        return null;
      }
    }),
    getUserProperties: () => ({
      getProperty: jest.fn((key) => {
        if (key === 'CURRENT_USER_ID') return 'test-user-id';
        return null;
      }),
      setProperty: jest.fn(),
      setProperties: jest.fn()
    })
  };
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({
      getSheets: () => [
        { getName: () => 'SheetA', isSheetHidden: () => false },
        { getName: () => 'SheetB', isSheetHidden: () => false },
        { getName: () => 'Config', isSheetHidden: () => false }
      ],
      getOwner: () => ({
        getEmail: () => 'owner@example.com'
      })
    })
  };
  global.getCurrentSpreadsheet = () => global.SpreadsheetApp.getActiveSpreadsheet();
  global.Session = { getActiveUser: () => ({ getEmail: () => 'a@example.com' }) };
  global.getActiveUserEmail = () => 'a@example.com';
  global.getUserInfo = jest.fn((userId) => {
    if (userId === 'test-user-id') {
      return {
        userId: 'test-user-id',
        adminEmail: 'a@example.com',
        spreadsheetId: 'spreadsheet-id-123',
        configJson: { showNames: false, showCounts: false }
      };
    }
    return null;
  });
}
afterEach(() => {
  delete global.PropertiesService;
  delete global.SpreadsheetApp;
  delete global.Session;
  delete global.getActiveUserEmail;
  delete global.getCurrentSpreadsheet;
});

test.skip('getAdminSettings returns board state', () => {
  setup();
  const result = getAdminSettings();
  expect(result).toEqual({
    allSheets: ['SheetA','SheetB'],
    currentUserEmail: 'a@example.com',
    deployId: 'deploy123',
    webAppUrl: 'https://example.com/app',
    adminEmails: ['a@example.com','b@example.com'],
    isUserAdmin: true,
    isPublished: true,
    activeSheetName: 'SheetA',
    showNames: false,
    showCounts: false
  });
});

test.skip('getAdminSettings identifies non-admin user', () => {
  setup();
  global.Session = { getActiveUser: () => ({ getEmail: () => 'other@example.com' }) };
  global.getActiveUserEmail = () => 'other@example.com';
  const result = getAdminSettings();
  expect(result.isUserAdmin).toBe(false);
});
