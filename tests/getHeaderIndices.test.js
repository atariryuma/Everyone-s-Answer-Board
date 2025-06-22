const { getHeaderIndices, COLUMN_HEADERS } = require('../src/Code.gs');

function setup(headers) {
  const sheet = {
    getLastColumn: () => headers.length,
    getRange: jest.fn(() => ({ getValues: () => [headers] })),
    getName: () => 'Sheet1'
  };
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({ getSheetByName: () => sheet })
  };
  global.CacheService = { getScriptCache: () => ({ get: () => null, put: () => null }) };
  return { sheet };
}

afterEach(() => {
  delete global.SpreadsheetApp;
  delete global.CacheService;
});

test('getHeaderIndices ignores unrelated header differences', () => {
  const headers = [
    'Email',
    'Class',
    'Question',
    COLUMN_HEADERS.UNDERSTAND,
    COLUMN_HEADERS.LIKE,
    COLUMN_HEADERS.CURIOUS,
    COLUMN_HEADERS.HIGHLIGHT
  ];
  setup(headers);
  const indices = getHeaderIndices('Sheet1');
  expect(indices).toEqual({
    [COLUMN_HEADERS.UNDERSTAND]: 3,
    [COLUMN_HEADERS.LIKE]: 4,
    [COLUMN_HEADERS.CURIOUS]: 5,
    [COLUMN_HEADERS.HIGHLIGHT]: 6
  });
  expect(Object.keys(indices)).toHaveLength(4);
});
