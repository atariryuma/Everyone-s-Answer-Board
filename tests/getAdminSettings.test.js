const { getAdminSettings } = require('../src/Code.gs');
function setup() {
  global.PropertiesService = {
    getScriptProperties: () => ({
      getProperty: (key) => {
        switch (key) {
          case 'IS_PUBLISHED': return 'true';
          case 'ACTIVE_SHEET_NAME': return 'SheetA';
          case 'DISPLAY_MODE': return 'named';
          case 'ADMIN_EMAILS': return 'a@example.com,b@example.com';
          case 'REACTION_COUNT_ENABLED': return 'true';
          case 'SCORE_SORT_ENABLED': return 'false';
          default: return null;
        }
      }
    })
  };
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({
      getSheets: () => [
        { getName: () => 'SheetA', isSheetHidden: () => false },
        { getName: () => 'SheetB', isSheetHidden: () => false }
      ]
    })
  };
  global.Session = {
    getActiveUser: () => ({ getEmail: () => 'a@example.com' })
  };
}
afterEach(() => {
  delete global.PropertiesService;
  delete global.SpreadsheetApp;
  delete global.Session;
});

test('getAdminSettings returns board state', () => {
  setup();
  const result = getAdminSettings();
  expect(result).toEqual({
    isPublished: true,
    activeSheetName: 'SheetA',
    allSheets: ['SheetA','SheetB'],
    displayMode: 'named',
    adminEmails: ['a@example.com','b@example.com'],
    currentUserEmail: 'a@example.com',
    reactionCountEnabled: true,
    scoreSortEnabled: false
  });
});

test('getAdminSettings throws for non-admin user', () => {
  setup();
  global.Session = { getActiveUser: () => ({ getEmail: () => 'other@example.com' }) };
  expect(() => getAdminSettings()).toThrow('管理者のみ実行できます');
});
