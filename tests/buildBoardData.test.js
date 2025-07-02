const { loadCode } = require('./shared-mocks');
const { buildBoardData, COLUMN_HEADERS } = loadCode();

function setup({configRows, dataRows}) {
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({
      getSheetByName: (name) => {
        if (name === 'Config') {
          return { getDataRange: () => ({ getValues: () => configRows }) };
        }
        if (name === 'Sheet1') {
          return { getDataRange: () => ({ getValues: () => dataRows }) };
        }
        return null;
      }
    })
  };
  global.getCurrentSpreadsheet = () => global.SpreadsheetApp.getActiveSpreadsheet();
  global.CacheService = { getScriptCache: () => ({ get: () => null, put: () => null, remove: () => null }) };
  global.PropertiesService = {
    getScriptProperties: () => ({ getProperty: () => null }),
    getUserProperties: () => ({
      getProperty: jest.fn(),
      setProperty: jest.fn(),
      setProperties: jest.fn()
    })
  };
  global.getConfig = (sheetName) => {
    const row = configRows.find((r, idx) => idx > 0 && r[0] === sheetName);
    if (!row) throw new Error('missing');
    return { questionHeader: row[1], answerHeader: row[2], reasonHeader: row[3] || null, nameHeader: row[4] || null, classHeader: row[5] || null };
  };
}

afterEach(() => {
  delete global.SpreadsheetApp;
  delete global.CacheService;
  delete global.PropertiesService;
  delete global.getConfig;
  delete global.getCurrentSpreadsheet;
});

test('buildBoardData reads names from same sheet', () => {
  const configRows = [
    ['表示シート名','Q','A','R','名前','クラス'],
    ['Sheet1','Q','A','R','名前','クラス']
  ];
  const dataRows = [
    ['Q','A','R','名前','クラス', COLUMN_HEADERS.EMAIL],
    ['q1','ans1','why1','Tanaka','1-A','t@example.com'],
    ['q1','ans2','why2','Suzuki','1-B','s@example.com']
  ];
  setup({ configRows, dataRows });
  const result = buildBoardData('Sheet1');
  expect(result.header).toBe('Q');
  expect(result.entries).toEqual([
    { answer:'ans1', reason:'why1', name:'Tanaka', class:'1-A' },
    { answer:'ans2', reason:'why2', name:'Suzuki', class:'1-B' }
  ]);
});

test('buildBoardData uses question column when answer header missing', () => {
  const configRows = [
    ['表示シート名','Q','','R','名前','クラス'],
    ['Sheet1','Q','','R','名前','クラス']
  ];
  const dataRows = [
    ['Q','R','名前','クラス', COLUMN_HEADERS.EMAIL],
    ['ans1','why1','Tanaka','1-A','t@example.com'],
    ['ans2','why2','Suzuki','1-B','s@example.com']
  ];
  setup({ configRows, dataRows });
  const result = buildBoardData('Sheet1');
  expect(result.header).toBe('Q');
  expect(result.entries).toEqual([
    { answer:'ans1', reason:'why1', name:'Tanaka', class:'1-A' },
    { answer:'ans2', reason:'why2', name:'Suzuki', class:'1-B' }
  ]);
});

test('buildBoardData uses answer column when question header missing', () => {
  const configRows = [
    ['表示シート名','Q','A','R','名前','クラス'],
    ['Sheet1','','A','R','名前','クラス']
  ];
  const dataRows = [
    ['Q','A','R','名前','クラス', COLUMN_HEADERS.EMAIL],
    ['q1','ans1','why1','Tanaka','1-A','t@example.com']
  ];
  setup({ configRows, dataRows });
  const result = buildBoardData('Sheet1');
  expect(result.header).toBe('A');
  expect(result.entries).toEqual([
    { answer:'ans1', reason:'why1', name:'Tanaka', class:'1-A' }
  ]);
});

