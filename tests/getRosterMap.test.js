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

function setupMocks(cacheValue) {
  const sheet = buildSheet();
  const cacheObj = {
    get: jest.fn(() => cacheValue),
    put: jest.fn()
  };
  global.CacheService = {
    getScriptCache: () => cacheObj
  };
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({
      getSheetByName: jest.fn(() => sheet)
    })
  };
  global.PropertiesService = {
    getScriptProperties: () => ({ getProperty: () => null })
  };
  return { sheet, cache: cacheObj };
}

afterEach(() => {
  delete global.CacheService;
  delete global.SpreadsheetApp;
  delete global.PropertiesService;
});

test('getRosterMap builds map and caches it', () => {
  const { sheet, cache } = setupMocks(null);

  const result = getRosterMap();

  expect(result).toEqual({
    'a@example.com': 'A Alice',
    'b@example.com': 'B Bob (Bobby)'
  });
  expect(sheet.getDataRange).toBeDefined();
  expect(cache.put).toHaveBeenCalledWith(
    'roster_name_map_v3',
    JSON.stringify(result),
    21600
  );
});

test('getRosterMap returns cached map when available', () => {
  const cached = { 'x@example.com': 'X X' };
  const { cache } = setupMocks(JSON.stringify(cached));
  const sheetCall = SpreadsheetApp.getActiveSpreadsheet().getSheetByName;

  const result = getRosterMap();

  expect(result).toEqual(cached);
  expect(sheetCall).not.toHaveBeenCalled();
  expect(cache.put).not.toHaveBeenCalled();
});

test('getRosterMap uses ROSTER_SHEET_NAME property when set', () => {
  const { sheet } = setupMocks(null);
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

