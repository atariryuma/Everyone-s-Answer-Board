const fs = require('fs');
const vm = require('vm');

describe('isDeployUser checks spreadsheet editors', () => {
  const code = fs.readFileSync('src/Core.gs', 'utf8');
  let context;
  beforeEach(() => {
    const scriptProps = { getProperty: jest.fn(() => 'SPREADSHEET_ID') };
    const file = {
      getEditors: () => [{ getEmail: () => 'editor@example.com' }],
      getOwner: () => ({ getEmail: () => 'owner@example.com' })
    };
    context = {
      PropertiesService: { getScriptProperties: () => scriptProps },
      Session: { getActiveUser: () => ({ getEmail: () => 'editor@example.com' }) },
      DriveApp: { getFileById: jest.fn(() => file) },
      SCRIPT_PROPS_KEYS: { DATABASE_SPREADSHEET_ID: 'DATABASE_SPREADSHEET_ID' },
      console
    };
    vm.createContext(context);
    vm.runInContext(code, context);
  });

  test('returns true for editor email', () => {
    context.Session.getActiveUser = () => ({ getEmail: () => 'editor@example.com' });
    expect(context.isDeployUser()).toBe(true);
  });

  test('returns false for non editor', () => {
    context.Session.getActiveUser = () => ({ getEmail: () => 'other@example.com' });
    expect(context.isDeployUser()).toBe(false);
  });
});
