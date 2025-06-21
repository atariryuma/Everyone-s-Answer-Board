const { buildBoardData } = require('../src/Code.gs');
const { COLUMN_HEADERS } = require('../src/Code.gs');

function setup({configRows, dataRows, rosterRows}) {
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({
      getSheetByName: (name) => {
        if (name === 'Config') {
          return { getDataRange: () => ({ getValues: () => configRows }) };
        }
        if (name === 'Sheet1') {
          return { getDataRange: () => ({ getValues: () => dataRows }) };
        }
        if (name === 'sheet 1') {
          return { getDataRange: () => ({ getValues: () => rosterRows }) };
        }
        return null;
      }
    })
  };
  global.CacheService = { getScriptCache: () => ({ get: () => null, put: () => null, remove: () => null }) };
  global.PropertiesService = {
    getScriptProperties: () => ({ getProperty: () => null })
  };
  global.getConfig = (sheetName) => {
    const row = configRows.find((r, idx) => idx > 0 && r[0] === sheetName);
    if (!row) throw new Error('missing');
    return { questionHeader: row[1], answerHeader: row[2], reasonHeader: row[3] || null, nameMode: row[4], nameHeader: row[5] || null, classHeader: row[6] || null };
  };
  global.getRosterMap = () => {
    if (!rosterRows) return {};
    const headers = rosterRows[0];
    const lastIdx = headers.indexOf('姓');
    const firstIdx = headers.indexOf('名');
    const emailIdx = headers.indexOf('Googleアカウント');
    const map = {};
    rosterRows.slice(1).forEach(r => {
      map[r[emailIdx]] = `${r[lastIdx]} ${r[firstIdx]}`;
    });
    return map;
  };
}

afterEach(() => {
  delete global.SpreadsheetApp;
  delete global.CacheService;
  delete global.PropertiesService;
  delete global.getConfig;
  delete global.getRosterMap;
});

test('buildBoardData reads names from same sheet', () => {
  const configRows = [
    ['表示シート名','Q','A','R','同一シート','名前','クラス'],
    ['Sheet1','Q','A','R','同一シート','名前','クラス']
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

test('buildBoardData uses roster when name mode is separate sheet', () => {
  const configRows = [
    ['表示シート名','Q','A','R','別シート','',''],
    ['Sheet1','Q','A','R','別シート','','']
  ];
  const dataRows = [
    [COLUMN_HEADERS.EMAIL,'Q','A','R'],
    ['t@example.com','q1','ans1','why1']
  ];
  const rosterRows = [
    ['姓','名','ニックネーム','Googleアカウント'],
    ['田','太郎','', 't@example.com']
  ];
  setup({ configRows, dataRows, rosterRows });
  const result = buildBoardData('Sheet1');
  expect(result.entries[0].name).toBe('田 太郎');
});
