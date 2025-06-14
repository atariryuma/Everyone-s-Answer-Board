const { addLike, COLUMN_HEADERS } = require('../src/Code.gs');

function buildSheet() {
  const headerRow = [COLUMN_HEADERS.EMAIL, COLUMN_HEADERS.CLASS, COLUMN_HEADERS.OPINION, COLUMN_HEADERS.REASON, COLUMN_HEADERS.LIKES];
  let likeValue = 'a@example.com';
  return {
    getLastColumn: () => headerRow.length,
    getRange: jest.fn((row, col, numRows, numCols) => {
      if (row === 1) {
        return { getValues: () => [headerRow] };
      }
      return {
        getValue: () => likeValue,
        setValue: (val) => { likeValue = val; }
      };
    })
  };
}

function setupMocks(userEmail, sheet) {
  global.LockService = { getScriptLock: () => ({ waitLock: jest.fn(), releaseLock: jest.fn() }) };
  global.Session = { getActiveUser: () => ({ getEmail: () => userEmail }) };
  global.PropertiesService = { getScriptProperties: () => ({ getProperty: () => 'Sheet1' }) };
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({ getSheetByName: () => sheet })
  };
}

afterEach(() => {
  delete global.LockService;
  delete global.Session;
  delete global.PropertiesService;
  delete global.SpreadsheetApp;
});

test('addLike toggles user in list', () => {
  const sheet = buildSheet();
  setupMocks('b@example.com', sheet);

  const result1 = addLike(2);
  expect(result1.status).toBe('ok');
  expect(result1.newScore).toBe(2);

  const result2 = addLike(2);
  expect(result2.status).toBe('ok');
  expect(result2.newScore).toBe(1);
});

test('addLike errors when user email is empty', () => {
  const sheet = buildSheet();
  setupMocks('', sheet);

  const result = addLike(2);
  expect(result.status).toBe('error');
});
