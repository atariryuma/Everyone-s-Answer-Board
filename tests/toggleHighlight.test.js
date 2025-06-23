const { toggleHighlight, COLUMN_HEADERS } = require('../src/Code.gs');

function buildSheet() {
  const headerRow = [
    COLUMN_HEADERS.EMAIL,
    COLUMN_HEADERS.CLASS,
    COLUMN_HEADERS.OPINION,
    COLUMN_HEADERS.REASON,
    COLUMN_HEADERS.TIMESTAMP,
    COLUMN_HEADERS.UNDERSTAND,
    COLUMN_HEADERS.LIKE,
    COLUMN_HEADERS.CURIOUS,
    COLUMN_HEADERS.HIGHLIGHT
  ];
  let highlight = false;
  return {
    getLastColumn: () => headerRow.length,
    getRange: jest.fn((row, col, numRows, numCols) => {
      if (row === 1) {
        return { getValues: () => [headerRow] };
      }
      return {
        getValue: () => highlight,
        setValue: (val) => { highlight = val; }
      };
    }),
    isSheetHidden: () => false,
    getName: () => 'Sheet1'
  };
}

function setupMocks(sheet, cacheImpl, userEmail = 'admin@example.com', adminEmails = 'admin@example.com') {
  global.LockService = { getScriptLock: () => ({ waitLock: jest.fn(), releaseLock: jest.fn() }) };
  global.Session = { getActiveUser: () => ({ getEmail: () => userEmail }) };
  global.PropertiesService = {
    getScriptProperties: () => ({
      getProperty: (key) => key === 'ADMIN_EMAILS' ? adminEmails : null
    }),
    getUserProperties: () => ({
      getProperty: jest.fn(),
      setProperty: jest.fn(),
      setProperties: jest.fn()
    })
  };
  global.CacheService = { getScriptCache: () => cacheImpl || ({ get: () => null, put: () => null }) };
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

test('toggleHighlight flips stored value', () => {
  const sheet = buildSheet();
  setupMocks(sheet);

  const result1 = toggleHighlight(2, 'Sheet1');
  expect(result1).toEqual({ status: 'ok', highlight: true });

  const result2 = toggleHighlight(2, 'Sheet1');
  expect(result2).toEqual({ status: 'ok', highlight: false });
});

test('toggleHighlight reads headers only once with cache', () => {
  const sheet = buildSheet();
  const store = {};
  const cache = {
    get: jest.fn(key => store[key] || null),
    put: jest.fn((key, val) => { store[key] = val; })
  };
  setupMocks(sheet, cache);

  toggleHighlight(2, 'Sheet1');
  toggleHighlight(2, 'Sheet1');

  const headerCalls = sheet.getRange.mock.calls.filter(c => c[0] === 1).length;
  expect(headerCalls).toBe(1);
});

test('toggleHighlight returns error for non-admin user', () => {
  const sheet = buildSheet();
  setupMocks(sheet, null, 'user@example.com', 'admin@example.com');
  const result = toggleHighlight(2, 'Sheet1');
  expect(result.status).toBe('error');
});
