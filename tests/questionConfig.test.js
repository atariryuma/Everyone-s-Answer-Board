const fs = require('fs');
const vm = require('vm');

describe('getQuestionConfig simple', () => {
  const code = fs.readFileSync('src/Core.gs', 'utf8');
  const context = {
    Utilities: { getUuid: () => `test-uuid-${Math.random()}` },
  };
  vm.createContext(context);
  vm.runInContext(code, context);

  test('returns simple config', () => {
    const cfg = context.getQuestionConfig('simple');
    expect(cfg.classQuestion.title).toBe('クラス');
    expect(cfg.classQuestion.choices.length).toBe(4);
    expect(cfg.nameQuestion.title).toBe('名前');
  });
});
