const { loadCode } = require('./shared-mocks');
const { isUserAdmin } = loadCode();

function setup({ currentEmail, userId, boardId, adminLists }) {
  const scriptProps = {
    getProperty: (key) => {
      if (key === 'DATABASE_ID') return 'db1';
      if (key.startsWith('ADMIN_EMAILS_')) {
        const id = key.replace('ADMIN_EMAILS_', '');
        const emails = adminLists[id];
        return emails ? emails.join(',') : null;
      }
      return null;
    },
    setProperty: jest.fn()
  };

  const userProps = {
    getProperty: jest.fn((k) => (k === 'CURRENT_USER_ID' ? userId : null)),
    setProperty: jest.fn(),
    setProperties: jest.fn()
  };

  const sheet = {
    getDataRange: () => ({
      getValues: () => [
        ['userId','spreadsheetId','adminEmail','configJson','lastAccessedAt'],
        [userId, boardId, currentEmail, '{}', '']
      ]
    }),
    getRange: jest.fn(() => ({ setValue: jest.fn() }))
  };

  global.PropertiesService = {
    getScriptProperties: () => scriptProps,
    getUserProperties: () => userProps
  };
  global.DriveApp = {
    getFileById: jest.fn(() => ({ setSharing: jest.fn() })),
    Access: { DOMAIN_WITH_LINK: 'link' },
    Permission: { VIEW: 'view' }
  };
  global.SpreadsheetApp = {
    openById: jest.fn(() => ({ getSheetByName: () => sheet }))
  };

  global.Session = { getActiveUser: () => ({ getEmail: () => currentEmail }) };
}

afterEach(() => {
  delete global.PropertiesService;
  delete global.Session;
  delete global.SpreadsheetApp;
  delete global.DriveApp;
});

test.skip('returns true when user is admin for current board', () => {
  setup({
    currentEmail: 'board1admin@example.com',
    userId: 'user1234567',
    boardId: 'board1',
    adminLists: {
      board1: ['board1admin@example.com'],
      board2: ['board2admin@example.com']
    }
  });
  expect(isUserAdmin()).toBe(true);
});

test.skip('returns false when user is admin for a different board', () => {
  setup({
    currentEmail: 'board2admin@example.com',
    userId: 'user1234567',
    boardId: 'board1',
    adminLists: {
      board1: ['board1admin@example.com'],
      board2: ['board2admin@example.com']
    }
  });
  expect(isUserAdmin()).toBe(false);
});
