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

test('admin user with ?page=admin sees admin view', () => {
  const { getTemplate } = setup({});
  const e = { parameter: { page: 'admin' } };
  doGet(e);
  const template = getTemplate();
  expect(HtmlService.createTemplateFromFile).toHaveBeenCalledWith('Page');
  expect(template.displayMode).toBe('named');
});

test('non-admin default shows student view', () => {
  const { getTemplate } = setup({ userEmail: 'user@example.com' });
  const e = { parameter: {} };
  doGet(e);
  const template = getTemplate();
  expect(HtmlService.createTemplateFromFile).toHaveBeenCalledWith('Page');
  expect(template.displayMode).toBe('anonymous');
});

test('admin email without page parameter shows student view', () => {
  const { getTemplate } = setup({});
  const e = { parameter: {} };
  doGet(e);
  const template = getTemplate();
  expect(HtmlService.createTemplateFromFile).toHaveBeenCalledWith('Page');
  expect(template.displayMode).toBe('anonymous');
});



test('doGet sets isAdminUser based on isUserAdmin when in student mode', () => {
  const { getTemplate } = setup({ userEmail: 'user@example.com', adminEmails: 'admin@example.com' });
  const e = { parameter: {} };
  doGet(e);
  const template = getTemplate();
  expect(template.isAdminUser).toBe(false);
});

