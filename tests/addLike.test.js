const { addLike } = require('../src/Code.gs');

// headers copied from Code.gs
const HEADERS = {
  EMAIL: 'メールアドレス',
  CLASS: 'クラスを選択してください。',
  OPINION: 'これまでの学んだことや、経験したことから、根からとり入れた水は、植物のからだのどこを通るのか予想しましょう。',
  REASON: '予想したわけを書きましょう。',
  LIKES: 'いいね！'
};

function buildSheet() {
  const headerRow = [HEADERS.EMAIL, HEADERS.CLASS, HEADERS.OPINION, HEADERS.REASON, HEADERS.LIKES];
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
