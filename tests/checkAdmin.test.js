const { checkAdmin } = require('../src/Code.gs');

function setup(admins, currentEmail) {
  global.PropertiesService = {
    getScriptProperties: () => ({ getProperty: (key) => key === 'ADMIN_EMAILS' ? admins.join(',') : null }),
    getUserProperties: () => ({ getProperty: () => null })
  };
  global.Session = { getActiveUser: () => ({ getEmail: () => currentEmail }) };
}

afterEach(() => {
  delete global.PropertiesService;
  delete global.Session;
});

test('checkAdmin returns true for admin user', () => {
  setup(['a@example.com','b@example.com'], 'a@example.com');
  expect(checkAdmin()).toBe(true);
});

test('checkAdmin returns false for non-admin', () => {
  setup(['a@example.com','b@example.com'], 'c@example.com');
  expect(checkAdmin()).toBe(false);
});
