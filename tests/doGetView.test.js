const { doGet } = require('../src/Code.gs');

afterEach(() => {
  delete global.HtmlService;
  delete global.getActiveUserEmail;
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
  global.getActiveUserEmail = () => userEmail;
  const output = { setTitle: jest.fn(() => output), addMetaTag: jest.fn(() => output) };
  let template = { evaluate: () => output };
  global.HtmlService = {
    createTemplateFromFile: jest.fn(() => template)
  };
  return { output, getTemplate: () => template };
}

test('admin view=board shows Page template even when unpublished', () => {
  const { output } = setup({ isPublished: false });
  const e = { parameter: { view: 'board' } };
  doGet(e);
  expect(HtmlService.createTemplateFromFile).toHaveBeenCalledWith('Page');
});

test('admin default shows Page when published', () => {
  setup({ isPublished: true });
  const e = { parameter: {} };
  doGet(e);
  expect(HtmlService.createTemplateFromFile).toHaveBeenCalledWith('Page');
});

test('admin default shows Unpublished when not published', () => {
  setup({ isPublished: false });
  const e = { parameter: {} };
  doGet(e);
  expect(HtmlService.createTemplateFromFile).toHaveBeenCalledWith('Unpublished');
});

test('mode=admin enables admin view', () => {
  const { getTemplate } = setup({ isPublished: true });
  const e = { parameter: { mode: 'admin' } };
  doGet(e);
  const template = getTemplate();
  expect(HtmlService.createTemplateFromFile).toHaveBeenCalledWith('Page');
  expect(template.adminMode).toBe(true);
});

