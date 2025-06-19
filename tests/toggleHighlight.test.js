const { toggleHighlight, COLUMN_HEADERS } = require('../src/Code.gs');

function buildSheet() {
  const headerRow = [
    COLUMN_HEADERS.EMAIL,
    COLUMN_HEADERS.CLASS,
    COLUMN_HEADERS.OPINION,
    COLUMN_HEADERS.REASON,
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

function setupMocks(sheet) {
  global.LockService = { getScriptLock: () => ({ waitLock: jest.fn(), releaseLock: jest.fn() }) };
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
  delete global.PropertiesService;
  delete global.CacheService;
  delete global.SpreadsheetApp;
});

test('toggleHighlight flips stored value', () => {
  const sheet = buildSheet();
  setupMocks(sheet);

  const result1 = toggleHighlight(2);
  expect(result1).toEqual({ status: 'ok', highlight: true });

  const result2 = toggleHighlight(2);
  expect(result2).toEqual({ status: 'ok', highlight: false });
});
