const { getRosterMap } = require('../src/Code.gs');

function buildSheet() {
  const data = [
    ['姓', '名', 'ニックネーム', 'Googleアカウント'],
    ['A', 'Alice', '', 'a@example.com'],
    ['B', 'Bob', 'Bobby', 'b@example.com']
  ];
  return {
    getDataRange: jest.fn(() => ({ getValues: () => data }))
  };
}

function setupMocks() {
  const sheet = buildSheet();
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({
      getSheetByName: jest.fn(() => sheet)
    })
  };
  global.PropertiesService = {
    getScriptProperties: () => ({ getProperty: () => null })
  };
  return { sheet };
}

afterEach(() => {
  delete global.SpreadsheetApp;
  delete global.PropertiesService;
});

test('getRosterMap builds map from sheet', () => {
  const { sheet } = setupMocks();

  const result = getRosterMap();

  expect(result).toEqual({
    'a@example.com': 'A Alice',
    'b@example.com': 'B Bob (Bobby)'
  });
  expect(sheet.getDataRange).toBeDefined();
});

test('getRosterMap uses ROSTER_SHEET_NAME property when set', () => {
  const { sheet } = setupMocks();
  const spy = jest.fn(() => sheet);
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({ getSheetByName: spy })
  };
  global.PropertiesService = {
    getScriptProperties: () => ({ getProperty: () => 'RosterSheet' })
  };

  getRosterMap();

  expect(spy).toHaveBeenCalledWith('RosterSheet');
});

