const fs = require('fs');
const vm = require('vm');

describe('isDeployUser uses ADMIN_EMAIL property', () => {
  const code = fs.readFileSync('src/Core.gs', 'utf8');
  let context;
  beforeEach(() => {
    const scriptProps = { getProperty: jest.fn((key) => {
      if (key === 'ADMIN_EMAIL') return 'admin@example.com';
      return null;
    }) };
    context = {
      PropertiesService: { getScriptProperties: () => scriptProps },
      Session: { getActiveUser: () => ({ getEmail: () => 'admin@example.com' }) },
      SCRIPT_PROPS_KEYS: { ADMIN_EMAIL: 'ADMIN_EMAIL' },
      console
    };
    vm.createContext(context);
    vm.runInContext(code, context);
  });

  test('returns true when email matches property', () => {
    context.Session.getActiveUser = () => ({ getEmail: () => 'admin@example.com' });
    expect(context.isDeployUser()).toBe(true);
  });

  test('returns false when email does not match', () => {
    context.Session.getActiveUser = () => ({ getEmail: () => 'other@example.com' });
    expect(context.isDeployUser()).toBe(false);
  });
});
