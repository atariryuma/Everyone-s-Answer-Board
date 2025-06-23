let saveSheetConfig;

beforeEach(() => {
  jest.resetModules();
});

function setup(initialRows) {
  const rows = initialRows.slice();
  const sheet = {
    getDataRange: () => ({ getValues: () => rows }),
    appendRow: jest.fn(row => rows.push(row)),
    getRange: jest.fn((r,c,n,m)=>({ setValues: vals => { rows[r-1] = vals[0]; } }))
  };
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({
      getSheetByName: () => sheet,
      insertSheet: () => sheet
    })
  };
  global.PropertiesService = {
    getUserProperties: () => ({ getProperty: () => null })
  };
  return { sheet, rows };
}

afterEach(() => {
  delete global.SpreadsheetApp;
  delete global.PropertiesService;
  delete global.getCurrentSpreadsheet;
});

test('saveSheetConfig appends new row', () => {
  const headers = ['表示シート名','問題文ヘッダー','回答ヘッダー','理由ヘッダー','名前取得モード','名前列ヘッダー','クラス列ヘッダー'];
  const { sheet, rows } = setup([headers]);
  global.getCurrentSpreadsheet = () => ({
    getSheetByName: () => sheet,
    insertSheet: () => sheet
  });
  ({ saveSheetConfig } = require('../src/Code.gs'));
  const cfg = { questionHeader:'Q', answerHeader:'A', reasonHeader:'R', nameMode:'同一シート', nameHeader:'名前', classHeader:'クラス' };
  saveSheetConfig('Sheet1', cfg);
  expect(sheet.appendRow).toHaveBeenCalled();
  expect(rows[1]).toEqual(['Sheet1','Q','A','R','同一シート','名前','クラス']);
});
