const { getSheetHeaders } = require('../src/Code.gs');

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
  global.PropertiesService = {
    getUserProperties: () => ({ getProperty: () => null })
  };
  return { sheet };
}

afterEach(() => {
  delete global.SpreadsheetApp;
  delete global.PropertiesService;
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
  global.PropertiesService = { getUserProperties: () => ({ getProperty: () => null }) };
  expect(() => getSheetHeaders('Missing')).toThrow("シート 'Missing' が見つかりません。");
});
