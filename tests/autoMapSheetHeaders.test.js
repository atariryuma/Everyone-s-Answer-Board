const fs = require('fs');
const vm = require('vm');

describe('autoMapSheetHeaders override', () => {
  const code = fs.readFileSync('src/config.gs', 'utf8');
  const context = {
    debugLog: () => {},
    errorLog: () => {},
    infoLog: () => {},
    warnLog: () => {},
    console,
    Utilities: {
      sleep: () => {}, // Add Utilities.sleep mock
    },
    getUnifiedExecutionCache: () => ({
      setUserInfo: () => {},
      syncWithUnifiedCache: () => {},
    }),
  };
  vm.createContext(context);
  vm.runInContext(code, context);
  // Override dependencies after evaluation
  context.verifyUserAccess = jest.fn();
  context.getCurrentSpreadsheet = jest.fn(() => ({ getId: () => 'ID' }));
  context.getSheetHeaders = jest.fn(() => [
    'クラス',
    '名前',
    '質問X',
    'そう考える理由や体験があれば教えてください（任意）',
  ]);
  context.autoMapHeaders = jest.fn(() => ({
    opinionHeader: '',
    reasonHeader: '',
    nameHeader: '',
    classHeader: '',
  }));

  test('uses overrides for mapping', () => {
    const mapping = context.autoMapSheetHeaders('uid1', 'フォームの回答 1', {
      mainQuestion: '質問X',
      reasonQuestion: 'そう考える理由や体験があれば教えてください（任意）',
      nameQuestion: '名前',
      classQuestion: 'クラス',
    });
    expect(mapping.mainHeader).toBe('質問X');
    expect(mapping.classHeader).toBe('クラス');
    expect(mapping.nameHeader).toBe('名前');
    expect(mapping.rHeader).toBe('そう考える理由や体験があれば教えてください（任意）');
  });
});
