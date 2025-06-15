const { addReaction, COLUMN_HEADERS } = require('../src/Code.gs');

function buildSheet() {
  const headerRow = [
    COLUMN_HEADERS.EMAIL,
    COLUMN_HEADERS.CLASS,
    COLUMN_HEADERS.OPINION,
    COLUMN_HEADERS.REASON,
    COLUMN_HEADERS.UNDERSTAND,
    COLUMN_HEADERS.LIKE,
    COLUMN_HEADERS.CURIOUS
  ];
  const values = {
    UNDERSTAND: 'a@example.com',
    LIKE: '',
    CURIOUS: ''
  };
  return {
    getLastColumn: () => headerRow.length,
    getRange: jest.fn((row, col, numRows, numCols) => {
      if (row === 1) {
        return { getValues: () => [headerRow] };
      }
      const header = headerRow[col - 1];
      return {
        getValue: () => values[Object.keys(COLUMN_HEADERS).find(k => COLUMN_HEADERS[k] === header)] || '',
        setValue: (val) => {
          const key = Object.keys(COLUMN_HEADERS).find(k => COLUMN_HEADERS[k] === header);
          values[key] = val;
        }
      };
    }),
    values
  };
}

function setupMocks(userEmail, sheet) {
  global.LockService = { getScriptLock: () => ({ tryLock: jest.fn(() => true), releaseLock: jest.fn() }) };
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

test('addReaction toggles user in list', () => {
  const sheet = buildSheet();
  setupMocks('b@example.com', sheet);

  const result1 = addReaction(2, 'UNDERSTAND');
  expect(result1.status).toBe('ok');
  expect(result1.reactions.UNDERSTAND.count).toBe(2);
  expect(result1.reactions.UNDERSTAND.reacted).toBe(true);

  const result2 = addReaction(2, 'UNDERSTAND');
  expect(result2.status).toBe('ok');
  expect(result2.reactions.UNDERSTAND.count).toBe(1);
  expect(result2.reactions.UNDERSTAND.reacted).toBe(false);
});

test('addReaction switches reaction columns', () => {
  const sheet = buildSheet();
  setupMocks('c@example.com', sheet);

  const res1 = addReaction(2, 'LIKE');
  expect(res1.status).toBe('ok');
  expect(sheet.values.LIKE).toBe('c@example.com');
  expect(sheet.values.UNDERSTAND).toBe('a@example.com');
  expect(res1.reactions.LIKE.reacted).toBe(true);

  const res2 = addReaction(2, 'CURIOUS');
  expect(res2.status).toBe('ok');
  expect(sheet.values.LIKE).toBe('');
  expect(sheet.values.CURIOUS).toBe('c@example.com');
  expect(res2.reactions.LIKE.reacted).toBe(false);
  expect(res2.reactions.CURIOUS.reacted).toBe(true);
});

test('addReaction errors when user email is empty', () => {
  const sheet = buildSheet();
  setupMocks('', sheet);

  const result = addReaction(2, 'UNDERSTAND');
  expect(result.status).toBe('error');
});
