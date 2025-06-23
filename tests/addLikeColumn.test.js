const { addReaction, COLUMN_HEADERS } = require('../src/Code.gs');

function buildSheet() {
  const headerRow = [
    COLUMN_HEADERS.EMAIL,
    COLUMN_HEADERS.CLASS,
    COLUMN_HEADERS.LIKE,
    COLUMN_HEADERS.OPINION,
    COLUMN_HEADERS.REASON,
    COLUMN_HEADERS.TIMESTAMP,
    COLUMN_HEADERS.UNDERSTAND,
    COLUMN_HEADERS.CURIOUS,
    COLUMN_HEADERS.HIGHLIGHT
  ];
  const values = {
    UNDERSTAND: 'a@example.com',
    LIKE: '',
    CURIOUS: ''
  };
  const sheet = {
    getLastColumn: () => headerRow.length,
    getRange: jest.fn((row, col, numRows, numCols) => {
      if (row === 1 && col === 1) {
        return { getValues: () => [headerRow] };
      }
      if (row === 2 && col === 7 && numRows === 1 && numCols === 3) {
        return {
          getValues: () => [[values.UNDERSTAND, values.LIKE, values.CURIOUS]],
          setValues: (rows) => {
            [values.UNDERSTAND, values.LIKE, values.CURIOUS] = rows[0];
          }
        };
      }
      throw new Error('Unexpected range request');
    }),
    isSheetHidden: () => false,
    getName: () => 'Sheet1',
    values
  };
  return sheet;
}

function setupMocks(email, sheet) {
  global.LockService = { getScriptLock: () => ({ tryLock: jest.fn(() => true), releaseLock: jest.fn() }) };
  global.Session = { getActiveUser: () => ({ getEmail: () => email }) };
  global.PropertiesService = {
    getScriptProperties: () => ({}),
    getUserProperties: () => ({
      getProperty: jest.fn(),
      setProperty: jest.fn(),
      setProperties: jest.fn()
    })
  };
  global.CacheService = { getScriptCache: () => ({ get: () => null, put: () => null }) };
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({
      getSheetByName: () => sheet,
      getSheets: () => [sheet]
    })
  };
}

afterEach(() => {
  delete global.LockService;
  delete global.Session;
  delete global.PropertiesService;
  delete global.CacheService;
  delete global.SpreadsheetApp;
});

test('addReaction updates value in LIKE column', () => {
  const sheet = buildSheet();
  setupMocks('like@example.com', sheet);
  const result = addReaction(2, 'LIKE', 'Sheet1');
  expect(result.status).toBe('ok');
  expect(sheet.getRange.mock.calls[1][0]).toBe(2);
  expect(sheet.getRange.mock.calls[1][1]).toBe(7);
});

test('addReaction handles failure to get user email', () => {
  const sheet = buildSheet();
  global.LockService = { getScriptLock: () => ({ tryLock: jest.fn(() => true), releaseLock: jest.fn() }) };
  global.Session = { getActiveUser: () => ({ getEmail: () => { throw new Error('fail'); } }) };
  global.PropertiesService = {
    getScriptProperties: () => ({}),
    getUserProperties: () => ({
      getProperty: jest.fn(),
      setProperty: jest.fn(),
      setProperties: jest.fn()
    })
  };
  global.CacheService = { getScriptCache: () => ({ get: () => null, put: () => null }) };
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({
      getSheetByName: () => sheet,
      getSheets: () => [sheet]
    })
  };
  const result = addReaction(2, 'LIKE', 'Sheet1');
  expect(result.status).toBe('error');
});
