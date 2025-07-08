const fs = require('fs');
const path = require('path');
const vm = require('vm');

describe('Server config functions', () => {
  const configFile = fs.readFileSync(path.resolve(__dirname, '../src/config.gs'), 'utf8');
  const saveConfigScript = /function saveConfig\([\s\S]*?\n\}/.exec(configFile)[0];

  test('saveConfig calls saveSheetConfig with user spreadsheet id', () => {
    const ctx = {
      getUserInfo: jest.fn(() => ({ spreadsheetId: 'sheet123' })),
      saveSheetConfig: jest.fn(),
      console,
    };
    vm.createContext(ctx);
    vm.runInContext(saveConfigScript, ctx);
    const result = ctx.saveConfig('MySheet', { a: 1 });

    expect(ctx.getUserInfo).toHaveBeenCalled();
    expect(ctx.saveSheetConfig).toHaveBeenCalledWith('sheet123', 'MySheet', { a: 1 });
    expect(result).toEqual({ success: true, message: '設定が正常に保存されました。' });
  });

  test('saveConfig throws error when sheetName is missing', () => {
    const ctx = {
      getUserInfo: jest.fn(() => ({ spreadsheetId: 'sheet123' })),
      saveSheetConfig: jest.fn(),
      console,
    };
    vm.createContext(ctx);
    vm.runInContext(saveConfigScript, ctx);
    expect(() => ctx.saveConfig('', { a: 1 })).toThrow('sheetName');
  });
});
