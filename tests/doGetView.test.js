const { doGet } = require('../src/Code.gs');

afterEach(() => {
  delete global.HtmlService;
  delete global.Session;
  delete global.PropertiesService;
});

function setup({isPublished, userEmail='admin@example.com', adminEmails='admin@example.com'}) {
  global.PropertiesService = {
    getScriptProperties: () => ({
      getProperty: (key) => {
        if (key === 'IS_PUBLISHED') return isPublished ? 'true' : 'false';
        if (key === 'ACTIVE_SHEET_NAME') return 'Sheet1';
        if (key === 'ADMIN_EMAILS') return adminEmails;
        return null;
      }
    })
  };
  global.Session = { getActiveUser: () => ({ getEmail: () => userEmail }) };
  const output = { setTitle: jest.fn(() => output), addMetaTag: jest.fn(() => output) };
  global.HtmlService = {
    createTemplateFromFile: jest.fn(() => ({ evaluate: () => output }))
  };
  return { output };
}

test('admin view=board shows Page template even when unpublished', () => {
  const { output } = setup({ isPublished: false });
  const e = { parameter: { view: 'board' } };
  doGet(e);
  expect(HtmlService.createTemplateFromFile).toHaveBeenCalledWith('Page');
});

test('admin without board view shows Unpublished', () => {
  setup({ isPublished: true });
  const e = {};
  doGet(e);
  expect(HtmlService.createTemplateFromFile).toHaveBeenCalledWith('Unpublished');
});

