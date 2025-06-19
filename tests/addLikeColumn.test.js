const { addLike, COLUMN_HEADERS } = require('../src/Code.gs');

function buildSheet() {
  const headerRow = [
    COLUMN_HEADERS.EMAIL,
    COLUMN_HEADERS.CLASS,
    COLUMN_HEADERS.LIKE,
    COLUMN_HEADERS.OPINION,
    COLUMN_HEADERS.REASON,
    COLUMN_HEADERS.UNDERSTAND,
    COLUMN_HEADERS.CURIOUS,
    COLUMN_HEADERS.HIGHLIGHT
  ];
  let likeVal = '';
  const sheet = {
    getLastColumn: () => headerRow.length,
    getRange: jest.fn((row, col, numRows, numCols) => {
      if (row === 1 && col === 1) {
        return { getValues: () => [headerRow] };
      }
      if (row === 2 && col === 3) {
        return {
          getValue: () => likeVal,
          setValue: (v) => { likeVal = v; }
        };
      }
      throw new Error('Unexpected range request');
    }),
    isSheetHidden: () => false,
    getName: () => 'Sheet1'
  };
  return sheet;
}

function setupMocks(email, sheet) {
  global.LockService = { getScriptLock: () => ({ waitLock: jest.fn(), releaseLock: jest.fn() }) };
  global.Session = { getActiveUser: () => ({ getEmail: () => email }) };
  global.PropertiesService = { getScriptProperties: () => ({}) };
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

test('addLike updates value in LIKE column', () => {
  const sheet = buildSheet();
  setupMocks('like@example.com', sheet);
  const result = addLike(2);
  expect(result.status).toBe('ok');
  expect(sheet.getRange.mock.calls[1][0]).toBe(2);
  expect(sheet.getRange.mock.calls[1][1]).toBe(3);
});

test('addLike handles failure to get user email', () => {
  const sheet = buildSheet();
  global.LockService = { getScriptLock: () => ({ waitLock: jest.fn(), releaseLock: jest.fn() }) };
  global.Session = { getActiveUser: () => ({ getEmail: () => { throw new Error('fail'); } }) };
  global.PropertiesService = { getScriptProperties: () => ({}) };
  global.CacheService = { getScriptCache: () => ({ get: () => null, put: () => null }) };
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({
      getSheetByName: () => sheet,
      getSheets: () => [sheet]
    })
  };
  const result = addLike(2);
  expect(result.status).toBe('error');
});
