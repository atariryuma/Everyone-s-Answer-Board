const fs = require('fs');
const vm = require('vm');

function createMockItem() {
  return {
    setTitle: jest.fn(),
    setRequired: jest.fn(),
    setChoiceValues: jest.fn(),
    showOtherOption: jest.fn()
  };
}

describe('addUnifiedQuestions other option', () => {
  const code = fs.readFileSync('src/Core.gs', 'utf8');
  const context = {
    console,
    debugLog: () => {},
    Utilities: { getUuid: () => 'test-uuid-' + Math.random() }
  };
  vm.createContext(context);
  vm.runInContext(code, context);

  test('shows other option for multiple choice type', () => {
    const mainItem = createMockItem();
    const form = {
      setCollectEmail: jest.fn(),
      addCheckboxItem: jest.fn(() => mainItem),
      addParagraphTextItem: jest.fn(() => createMockItem()),
      addTextItem: jest.fn(() => createMockItem()),
      addMultipleChoiceItem: jest.fn(() => createMockItem()),
      addListItem: jest.fn(() => createMockItem())
    };

    context.addUnifiedQuestions(form, 'custom', {
      mainQuestion: {
        type: 'multiple',
        choices: ['A', 'B'],
        includeOthers: true
      },
      enableClass: false
    });

    expect(mainItem.showOtherOption).toHaveBeenCalledWith(true);
  });

  test('shows other option for single choice type', () => {
    const mainItem = createMockItem();
    const form = {
      setCollectEmail: jest.fn(),
      addMultipleChoiceItem: jest.fn(() => mainItem),
      addParagraphTextItem: jest.fn(() => createMockItem()),
      addTextItem: jest.fn(() => createMockItem()),
      addCheckboxItem: jest.fn(() => createMockItem()),
      addListItem: jest.fn(() => createMockItem())
    };

    context.addUnifiedQuestions(form, 'custom', {
      mainQuestion: {
        type: 'choice',
        choices: ['A', 'B'],
        includeOthers: true
      },
      enableClass: false
    });

    expect(mainItem.showOtherOption).toHaveBeenCalledWith(true);
  });
});
