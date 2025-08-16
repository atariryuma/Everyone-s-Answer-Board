const fs = require('fs');
const vm = require('vm');

describe('getSheetsList error handling', () => {
  let context;
  const coreCode = fs.readFileSync('src/Core.gs', 'utf8');

  beforeEach(() => {
    context = {
      console,
      debugLog: () => {},
      warnLog: jest.fn(),
      errorLog: jest.fn(),
      infoLog: () => {},
      Utilities: { sleep: jest.fn(), getUuid: () => 'uuid' },
      Session: { getActiveUser: () => ({ getEmail: () => 'admin@example.com' }) },
    };
    vm.createContext(context);
    vm.runInContext(coreCode, context);

    context.findUserById = jest.fn(() => ({
      spreadsheetId: 'sheet1',
      adminEmail: 'admin@example.com',
    }));
    context.getSheetsServiceCached = jest.fn(() => ({}));
    context.addServiceAccountToSpreadsheet = jest.fn();
    context.repairUserSpreadsheetAccess = jest.fn();
    context.getSpreadsheetsData = jest.fn();
  });

  test('repairs only when a 403 error occurs', () => {
    context.getSpreadsheetsData
      .mockImplementationOnce(() => { throw new Error('Sheets API error: 403 - {}'); })
      .mockImplementationOnce(() => ({ sheets: [{ properties: { title: 'Sheet1', sheetId: 1 } }] }));

    const result = context.getSheetsList('user1');

    expect(context.addServiceAccountToSpreadsheet).toHaveBeenCalledTimes(1);
    expect(context.getSpreadsheetsData).toHaveBeenCalledTimes(2);
    expect(result).toEqual([{ name: 'Sheet1', id: 1 }]);
  });

  test('retries with backoff for server errors without repairing access', () => {
    context.getSpreadsheetsData.mockImplementation(() => {
      throw new Error('Sheets API error: 500 - {}');
    });

    const result = context.getSheetsList('user1');

    expect(result).toEqual([]);
    expect(context.addServiceAccountToSpreadsheet).not.toHaveBeenCalled();
    expect(context.getSpreadsheetsData).toHaveBeenCalledTimes(3);
    expect(context.Utilities.sleep).toHaveBeenCalledTimes(2);
    expect(context.Utilities.sleep).toHaveBeenNthCalledWith(1, 1000);
    expect(context.Utilities.sleep).toHaveBeenNthCalledWith(2, 2000);
  });
});

