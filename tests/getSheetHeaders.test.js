const { loadCode } = require('./shared-mocks');
const { getSheetHeaders } = loadCode();

function setup(headers) {
  const sheet = {
    getLastColumn: jest.fn(() => headers.length),
    getRange: jest.fn(() => ({ getValues: () => [headers] }))
  };
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({
      getSheetByName: jest.fn(() => sheet)
    })
  };
  global.getCurrentSpreadsheet = () => global.SpreadsheetApp.getActiveSpreadsheet();
  global.PropertiesService = {
    getUserProperties: () => ({
      getProperty: jest.fn(),
      setProperty: jest.fn(),
      setProperties: jest.fn()
    })
  };
  return { sheet };
}

afterEach(() => {
  delete global.SpreadsheetApp;
  delete global.PropertiesService;
  delete global.getCurrentSpreadsheet;
});

test('getSheetHeaders returns header row', () => {
  const headers = ['A', 'B', 'C'];
  setup(headers);
  expect(getSheetHeaders('Sheet1')).toEqual(headers);
});

test('getSheetHeaders throws when sheet missing', () => {
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({ getSheetByName: jest.fn(() => null) })
  };
  global.getCurrentSpreadsheet = () => global.SpreadsheetApp.getActiveSpreadsheet();
  global.PropertiesService = {
    getUserProperties: () => ({
      getProperty: jest.fn(),
      setProperty: jest.fn(),
      setProperties: jest.fn()
    })
  };
  expect(() => getSheetHeaders('Missing')).toThrow("シート 'Missing' が見つかりません。");
});
