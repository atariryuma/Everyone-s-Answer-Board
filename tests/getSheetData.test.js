const { getSheetData } = require('../src/Code.gs');

// Japanese column headers copied from Code.gs
const HEADERS = {
  EMAIL: 'メールアドレス',
  CLASS: 'クラスを選択してください。',
  OPINION: 'これまでの学んだことや、経験したことから、根からとり入れた水は、植物のからだのどこを通るのか予想しましょう。',
  REASON: '予想したわけを書きましょう。',
  LIKES: 'いいね！'
};

function setupMocks(rows, userEmail) {
  const rosterHeaders = ['姓', '名', 'ニックネーム', 'Googleアカウント'];
  const rosterRows = [
    ['A', 'Alice', '', 'a@example.com'],
    ['B', 'Bob', '', 'b@example.com']
  ];

  // Mock SpreadsheetApp
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

  // Mock other global services
  global.Session = {
    getActiveUser: () => ({ getEmail: () => userEmail })
  };
  global.CacheService = { getScriptCache: () => ({ get: () => null, put: () => null }) };
}

afterEach(() => {
  delete global.SpreadsheetApp;
  delete global.Session;
  delete global.CacheService;
});

test('getSheetData filters and scores rows', () => {
  const data = [
    [HEADERS.EMAIL, HEADERS.CLASS, HEADERS.OPINION, HEADERS.REASON, HEADERS.LIKES],
    ['a@example.com', '1-1', 'Opinion1', 'Reason1', 'b@example.com'],
    ['b@example.com', '1-1', 'Opinion2', 'Reason2', ''],
    ['', '', '', '', ''] // ignored
  ];
  setupMocks(data, 'b@example.com');

  const result = getSheetData('Sheet1');

  expect(result.header).toBe(HEADERS.OPINION);
  expect(result.rows).toHaveLength(2);
  expect(result.rows[0].name).toBe('A Alice');
  expect(result.rows[0].hasLiked).toBe(true);
  expect(result.rows[1].name).toBe('B Bob');
  expect(result.rows[1].hasLiked).toBe(false);
});
