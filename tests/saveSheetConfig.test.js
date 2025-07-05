const { loadCode } = require('./shared-mocks');
const { saveSheetConfig } = loadCode();

function setup(initialRows) {
  const rows = initialRows.slice();
  const sheet = {
    getDataRange: () => ({ getValues: () => rows }),
    appendRow: jest.fn(row => rows.push(row)),
    getRange: jest.fn((r,c,n,m)=>({ setValues: vals => { rows[r-1] = vals[0]; } }))
  };
  global.SpreadsheetApp = {
    getActive: () => ({
      getSheetByName: () => sheet,
      insertSheet: () => sheet
    })
  };
  global.getCurrentSpreadsheet = () => global.SpreadsheetApp.getActive();
  global.PropertiesService = {
    getUserProperties: () => ({
      getProperty: jest.fn((key) => {
        if (key === 'CURRENT_USER_ID') return 'test-user-id';
        return null;
      }),
      setProperty: jest.fn(),
      setProperties: jest.fn()
    })
  };
  global.getUserInfo = jest.fn((userId) => {
    if (userId === 'test-user-id') {
      return {
        userId: 'test-user-id',
        adminEmail: 'admin@example.com',
        spreadsheetId: 'spreadsheet-id-123',
        configJson: {},
      };
    }
    return null;
  });
  global.updateUserConfig = jest.fn();
  global.auditLog = jest.fn();
  global.Session = {
    getActiveUser: () => ({ getEmail: () => 'admin@example.com' })
  };
  return { sheet, rows };
}

afterEach(() => {
  delete global.SpreadsheetApp;
  delete global.getCurrentSpreadsheet;
  delete global.PropertiesService;
  delete global.updateUserConfig;
  delete global.auditLog;
});

test.skip('saveSheetConfig appends new row', () => {
  const headers = ['表示シート名','問題文ヘッダー','回答ヘッダー','理由ヘッダー','名前列ヘッダー','クラス列ヘッダー'];
  const { sheet, rows } = setup([headers]);
  const cfg = { questionHeader:'Q', answerHeader:'A', reasonHeader:'R', nameHeader:'名前', classHeader:'クラス' };
  saveSheetConfig('spreadsheet-id-123', 'Sheet1', cfg);
  expect(sheet.appendRow).toHaveBeenCalled();
  expect(rows[1]).toEqual(['Sheet1','Q','A','R','名前','クラス']);
});
