const { toggleHighlight, COLUMN_HEADERS } = require('../src/Code.gs');

function buildSheet() {
  const headerRow = [COLUMN_HEADERS.HIGHLIGHT];
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
    })
  };
}

function setupMocks(sheet) {
  global.LockService = { getScriptLock: () => ({ waitLock: jest.fn(), releaseLock: jest.fn() }) };
  global.PropertiesService = {
    getScriptProperties: () => ({
      getProperty: (key) => {
        if (key === 'ADMIN_EMAILS') return 'admin@example.com';
        if (key === 'ACTIVE_SHEET_NAME') return 'Sheet1';
        return 'Sheet1';
      }
    })
  };
  global.getActiveUserEmail = () => 'admin@example.com';
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({ getSheetByName: () => sheet })
  };
}

afterEach(() => {
  delete global.LockService;
  delete global.PropertiesService;
  delete global.SpreadsheetApp;
  delete global.getActiveUserEmail;
});

test('toggleHighlight flips stored value', () => {
  const sheet = buildSheet();
  setupMocks(sheet);

  const result1 = toggleHighlight(2);
  expect(result1).toEqual({ status: 'ok', highlight: true });

  const result2 = toggleHighlight(2);
  expect(result2).toEqual({ status: 'ok', highlight: false });
});
