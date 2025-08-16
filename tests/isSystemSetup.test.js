const fs = require('fs');
const vm = require('vm');

describe('isSystemSetup requires ADMIN_EMAIL and SERVICE_ACCOUNT_CREDS', () => {
  const code = fs.readFileSync('src/main.gs', 'utf8');
  let context;

  let scriptProps;
  beforeEach(() => {
    scriptProps = {getProperty: jest.fn(() => null)};
    context = {
      console,
      PropertiesService: {getScriptProperties: () => scriptProps},
    };
    vm.createContext(context);
    vm.runInContext(code, context);
  });

  test('returns false when admin email missing', () => {
    scriptProps.getProperty = key =>
      key === 'DATABASE_SPREADSHEET_ID' ? 'id' : null;
    expect(context.isSystemSetup()).toBe(false);
  });

  test('returns false when service account missing', () => {
    scriptProps.getProperty = key => {
      if (key === 'DATABASE_SPREADSHEET_ID') return 'id';
      if (key === 'ADMIN_EMAIL') return 'admin@example.com';
      return null;
    };
    expect(context.isSystemSetup()).toBe(false);
  });

  test('returns true when all properties exist', () => {
    scriptProps.getProperty = key => {
      if (key === 'DATABASE_SPREADSHEET_ID') return 'id';
      if (key === 'ADMIN_EMAIL') return 'admin@example.com';
      if (key === 'SERVICE_ACCOUNT_CREDS')
        return '{"client_email":"a","private_key":"b"}';
      return null;
    };
    expect(context.isSystemSetup()).toBe(true);
  });
});
