const fs = require('fs');
const vm = require('vm');

describe('isSystemSetup requires ADMIN_EMAIL', () => {
  const code = fs.readFileSync('src/main.gs', 'utf8');
  let context;

  let scriptProps;
  beforeEach(() => {
    scriptProps = { getProperty: jest.fn(() => null) };
    context = {
      console,
      PropertiesService: { getScriptProperties: () => scriptProps }
    };
    vm.createContext(context);
    vm.runInContext(code, context);
  });

  test('returns false when admin email missing', () => {
    scriptProps.getProperty = (key) =>
      key === 'DATABASE_SPREADSHEET_ID' ? 'id' : null;
    expect(context.isSystemSetup()).toBe(false);
  });

  test('returns true when both properties exist', () => {
    scriptProps.getProperty = (key) => {
      if (key === 'DATABASE_SPREADSHEET_ID') return 'id';
      if (key === 'ADMIN_EMAIL') return 'admin@example.com';
      return null;
    };
    expect(context.isSystemSetup()).toBe(true);
  });
});
