const { doGet } = require('../src/Code.gs');

afterEach(() => {
  delete global.HtmlService;
  delete global.getActiveUserEmail;
  delete global.Session;
  delete global.PropertiesService;
});

function setup({ userEmail = 'admin@example.com', adminEmails = 'admin@example.com' }) {
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
  return { getTemplate: () => template };
}

test('page parameter is ignored and admin mode is based on user role', () => {
  const { getTemplate } = setup({});
  doGet({ parameter: { page: 'admin' } });
  const tpl = getTemplate();
  expect(tpl.isAdminUser).toBe(true);
  expect(tpl.showAdminFeatures).toBe(false);
});

test('admin user starts in admin mode', () => {
  const { getTemplate } = setup({});
  doGet({ parameter: {} });
  const tpl = getTemplate();
  expect(tpl.isAdminUser).toBe(true);
  expect(tpl.showAdminFeatures).toBe(false);
});

test('non admin user starts in viewer mode', () => {
  const { getTemplate } = setup({ userEmail: 'user@example.com' });
  doGet({ parameter: {} });
  const tpl = getTemplate();
  expect(tpl.isAdminUser).toBe(false);
  expect(tpl.showAdminFeatures).toBe(false);
});
