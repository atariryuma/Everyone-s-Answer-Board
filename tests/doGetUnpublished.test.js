const { doGet } = require('../src/Code.gs');

afterEach(() => {
  delete global.PropertiesService;
  delete global.Session;
  delete global.HtmlService;
});

function setup() {
  global.PropertiesService = {
    getScriptProperties: () => ({
      getProperty: (key) => {
        if (key === 'IS_PUBLISHED') return 'false';
        return null;
      }
    })
  };
  global.Session = { getActiveUser: () => ({ getEmail: () => { throw new Error('no user'); } }) };
  const output = { setTitle: jest.fn(() => output) };
  let template = { evaluate: () => output };
  global.HtmlService = {
    createTemplateFromFile: jest.fn(() => template)
  };
  return template;
}

test('doGet uses Unpublished template when email fails', () => {
  const template = setup();
  doGet({});
  expect(HtmlService.createTemplateFromFile).toHaveBeenCalledWith('Unpublished');
  expect(template.userEmail).toBe('');
});
