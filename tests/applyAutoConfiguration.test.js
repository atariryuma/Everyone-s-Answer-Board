const fs = require('fs');
const vm = require('vm');

describe('applyAutoConfiguration', () => {
  let context;
  beforeEach(() => {
    const errorHandlerCode = fs.readFileSync('src/errorHandler.gs', 'utf8');
    const coreCode = fs.readFileSync('src/Core.gs', 'utf8');
    context = {
      console,
      debugLog: () => {},
      infoLog: jest.fn(),
      errorLog: jest.fn(),
      Utilities: { getUuid: () => 'test-uuid' }
    };
    vm.createContext(context);
    vm.runInContext(errorHandlerCode, context);
    vm.runInContext(coreCode, context);
    context.saveSheetConfig = jest.fn();
  });

  test('returns success when saveSheetConfig succeeds', () => {
    context.saveSheetConfig.mockReturnValue({ status: 'success', message: 'ok' });
    const res = context.applyAutoConfiguration('U', 'SID', 'Sheet1', {
      success: true,
      guessedConfig: { opinionHeader: 'Q', nameHeader: 'N' }
    });
    expect(res.success).toBe(true);
    expect(context.saveSheetConfig).toHaveBeenCalledTimes(1);
  });

  test('retries and fails when saveSheetConfig keeps failing', () => {
    context.saveSheetConfig.mockReturnValue({ status: 'error', message: 'fail' });
    const res = context.applyAutoConfiguration('U', 'SID', 'Sheet1', {
      success: true,
      guessedConfig: { opinionHeader: 'Q', nameHeader: 'N' }
    });
    expect(res.success).toBe(false);
    expect(context.saveSheetConfig).toHaveBeenCalledTimes(2);
  });
});
