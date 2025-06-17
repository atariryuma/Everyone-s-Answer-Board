const { doGet } = require('../src/Code.gs');

afterEach(() => {
  delete global.HtmlService;
  delete global.getActiveUserEmail;
  delete global.Session;
  delete global.PropertiesService;
});

function setup({userEmail='admin@example.com', adminEmails='admin@example.com'}) {
  global.PropertiesService = {
    getScriptProperties: () => ({
      getProperty: (key) => {
        if (key === 'ADMIN_EMAILS') return adminEmails;
        return null;
      }
    })
  };
  global.getActiveUserEmail = () => userEmail;
  global.Session = { getActiveUser: () => ({ getEmail: () => userEmail }) };
  const output = { setTitle: jest.fn(() => output), addMetaTag: jest.fn(() => output) };
  let template = { evaluate: () => output };
  global.HtmlService = {
    createTemplateFromFile: jest.fn(() => template)
  };
  return { output, getTemplate: () => template };
}

test('admin view=board shows Page template', () => {
  const { output } = setup({});
  const e = { parameter: { view: 'board' } };
  doGet(e);
  expect(HtmlService.createTemplateFromFile).toHaveBeenCalledWith('Page');
});

test('admin default shows Page', () => {
  setup({});
  const e = { parameter: {} };
  doGet(e);
  expect(HtmlService.createTemplateFromFile).toHaveBeenCalledWith('Page');
});

test('admin default uses Page template regardless of publish status', () => {
  setup({});
  const e = { parameter: {} };
  doGet(e);
  expect(HtmlService.createTemplateFromFile).toHaveBeenCalledWith('Page');
});

test('admin=true enables admin view', () => {
  const { getTemplate } = setup({});
  const e = { parameter: { admin: 'true' } };
  doGet(e);
  const template = getTemplate();
  expect(HtmlService.createTemplateFromFile).toHaveBeenCalledWith('Page');
  expect(template.displayMode).toBe('named');
});

