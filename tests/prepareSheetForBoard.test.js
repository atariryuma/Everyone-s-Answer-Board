const { prepareSheetForBoard, COLUMN_HEADERS } = require('../src/Code.gs');

function setup(headers) {
  const sheet = {
    getLastColumn: () => headers.length,
    insertColumnAfter: jest.fn(() => { headers.push(''); }),
    getRange: jest.fn((row, col) => ({
      getValues: () => [headers],
      setValue: (val) => { headers[col - 1] = val; }
    }))
  };
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({ getSheetByName: () => sheet })
  };
  return { sheet, headers };
}

afterEach(() => {
  delete global.SpreadsheetApp;
});

test('prepareSheetForBoard appends missing reaction columns', () => {
  const base = [COLUMN_HEADERS.EMAIL, COLUMN_HEADERS.CLASS];
  const { headers } = setup(base);
  prepareSheetForBoard('Sheet1');
  expect(headers).toContain(COLUMN_HEADERS.TIMESTAMP);
  expect(headers).toContain(COLUMN_HEADERS.UNDERSTAND);
  expect(headers).toContain(COLUMN_HEADERS.LIKE);
  expect(headers).toContain(COLUMN_HEADERS.CURIOUS);
  expect(headers).toContain(COLUMN_HEADERS.HIGHLIGHT);
});
