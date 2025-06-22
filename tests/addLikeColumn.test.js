const { addReaction, COLUMN_HEADERS } = require('../src/Code.gs');

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
  const values = {
    UNDERSTAND: 'a@example.com',
    LIKE: '',
    CURIOUS: ''
  };
  const sheet = {
    getLastColumn: () => headerRow.length,
    getRange: jest.fn((row, col, numRows, numCols) => {
      if (row === 1 && col === 1) {
        const len = numCols || headerRow.length;
        return { getValues: () => [headerRow.slice(0, len)] };
      }
      const header = headerRow[col - 1];
      if (numRows === 1 && numCols === 1 || numRows === undefined) {
        return {
          getValue: () => values[Object.keys(COLUMN_HEADERS).find(k => COLUMN_HEADERS[k] === header)] || '',
          setValue: (val) => {
            const key = Object.keys(COLUMN_HEADERS).find(k => COLUMN_HEADERS[k] === header);
            values[key] = val;
          }
        };
      }
      if (row === 2 && numRows === 1 && numCols === 3) {
        const headers = headerRow.slice(col - 1, col - 1 + numCols);
        return {
          getValues: () => [headers.map(h => { const key = Object.keys(COLUMN_HEADERS).find(k => COLUMN_HEADERS[k] === h); return values[key] || ''; })],
          setValues: (rows) => {
            headers.forEach((h, i) => { const key = Object.keys(COLUMN_HEADERS).find(k => COLUMN_HEADERS[k] === h); values[key] = rows[0][i]; });
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

test('addReaction updates value in LIKE column', () => {
  const sheet = buildSheet();
  setupMocks('like@example.com', sheet);
  const result = addReaction(2, 'LIKE');
  expect(result.status).toBe('ok');
  expect(sheet.values.LIKE).toBe('like@example.com');
});

test('addReaction handles failure to get user email', () => {
  const sheet = buildSheet();
  global.LockService = { getScriptLock: () => ({ tryLock: jest.fn(() => true), releaseLock: jest.fn() }) };
  global.Session = { getActiveUser: () => ({ getEmail: () => { throw new Error('fail'); } }) };
  global.PropertiesService = { getScriptProperties: () => ({}) };
  global.CacheService = { getScriptCache: () => ({ get: () => null, put: () => null }) };
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({
      getSheetByName: () => sheet,
      getSheets: () => [sheet]
    })
  };
  const result = addReaction(2, 'LIKE');
  expect(result.status).toBe('error');
});
