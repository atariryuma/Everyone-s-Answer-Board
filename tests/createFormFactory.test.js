const fs = require('fs');
const vm = require('vm');

describe('createFormFactory returns URLs', () => {
  const code = fs.readFileSync('src/Core.gs', 'utf8');
  let context;

  beforeEach(() => {
    const mockItem = {
      setTitle: jest.fn(),
      setRequired: jest.fn(),
      setChoiceValues: jest.fn(),
      showOtherOption: jest.fn(),
      setHelpText: jest.fn()
    };
    const mockForm = {
      setDescription: jest.fn(),
      setCollectEmail: jest.fn(),
      setEmailCollectionType: jest.fn(),
      addTextItem: jest.fn(() => mockItem),
      addCheckboxItem: jest.fn(() => mockItem),
      addListItem: jest.fn(() => mockItem),
      addMultipleChoiceItem: jest.fn(() => mockItem),
      addParagraphTextItem: jest.fn(() => mockItem),
      getId: jest.fn(() => 'FORM_ID'),
      getPublishedUrl: jest.fn(() => 'https://example.com/view'),
      getEditUrl: jest.fn(() => 'https://example.com/edit')
    };

    context = {
      console,
      debugLog: () => {},
      FormApp: {
        create: jest.fn(() => mockForm),
        EmailCollectionType: { VERIFIED: 'VERIFIED' }
      },
      Utilities: { 
        formatDate: jest.fn(() => '2025/01/01 00:00:00'),
        getUuid: () => 'test-uuid-' + Math.random()
      }
    };
    vm.createContext(context);
    vm.runInContext(code, context);
    context.createLinkedSpreadsheet = jest.fn(() => ({
      spreadsheetId: 'SS_ID',
      spreadsheetUrl: 'https://example.com/ss',
      sheetName: 'Sheet1'
    }));
    context.addReactionColumnsToSpreadsheet = jest.fn();
  });

  test('returns editFormUrl and viewFormUrl', () => {
    const result = context.createFormFactory({
      userEmail: 'u@example.com',
      userId: 'uid',
      formTitle: 'Title'
    });
    expect(result.editFormUrl).toBe('https://example.com/edit');
    expect(result.viewFormUrl).toBe('https://example.com/view');
    expect(result.formUrl).toBe('https://example.com/view');
  });
});

