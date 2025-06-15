const { getSheetUpdates, COLUMN_HEADERS } = require('../src/Code.gs');

function setupMocks(rows, userEmail) {
  const rosterHeaders = ['姓', '名', 'ニックネーム', 'Googleアカウント'];
  const rosterRows = [['A', 'Alice', '', 'a@example.com']];
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({
      getSheetByName: (name) => {
        if (name === 'Sheet1') {
          return { getDataRange: () => ({ getValues: () => rows }) };
        }
        if (name === 'sheet 1') {
          return { getDataRange: () => ({ getValues: () => [rosterHeaders, ...rosterRows] }) };
        }
        return null;
      }
    })
  };
  global.Session = { getActiveUser: () => ({ getEmail: () => userEmail }) };
  global.CacheService = { getScriptCache: () => ({ get: () => null, put: () => null }) };
  global.PropertiesService = {
    getScriptProperties: () => ({
      getProperty: (key) => {
        if (key === 'DISPLAY_MODE') return 'named';
        if (key === 'ACTIVE_SHEET_NAME') return 'Sheet1';
        return null;
      }
    })
  };
  global.Utilities = {
    base64Encode: (b) => Buffer.from(b).toString('base64'),
    computeDigest: (alg, str) => Buffer.from(str),
    DigestAlgorithm: { MD5: 'md5' }
  };
}

afterEach(() => {
  delete global.SpreadsheetApp;
  delete global.Session;
  delete global.CacheService;
  delete global.PropertiesService;
  delete global.Utilities;
});

test('getSheetUpdates detects no changes when hashes match', () => {
  const data = [
    [
      COLUMN_HEADERS.EMAIL,
      COLUMN_HEADERS.CLASS,
      COLUMN_HEADERS.OPINION,
      COLUMN_HEADERS.REASON,
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ],
    ['a@example.com', '1-1', 'Opinion1', 'Reason1', '', '', '', 'false']
  ];
  setupMocks(data, 'a@example.com');
  const first = getSheetUpdates(undefined, 'newest', '{}');
  const second = getSheetUpdates(undefined, 'newest', JSON.stringify(first.hashMap));
  expect(first.changedRows.length).toBe(1);
  expect(second.changedRows.length).toBe(0);
  expect(second.removedRows.length).toBe(0);
});
