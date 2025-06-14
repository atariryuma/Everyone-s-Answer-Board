const code = require('../src/Code.gs');

afterEach(() => {
  jest.restoreAllMocks();
});

test('getPublishedSheetData forwards parameters and returns data', () => {
  jest.spyOn(code, 'getAppSettings').mockReturnValue({ activeSheetName: 'Sheet1' });
  jest.spyOn(code, 'getSheetData').mockReturnValue({ header: 'HEADER', rows: [1, 2, 3] });

  const result = code.getPublishedSheetData('classA', 'newest');

  expect(code.getSheetData).toHaveBeenCalledWith('Sheet1', 'classA', 'newest');
  expect(result).toEqual({ sheetName: 'Sheet1', header: 'HEADER', rows: [1, 2, 3] });
});
