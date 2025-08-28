const fs = require('fs');
const vm = require('vm');

describe('verifyUserAccess security checks', () => {
  const coreCode = fs.readFileSync('src/Core.gs', 'utf8');
  const unifiedUtilitiesCode = fs.readFileSync('src/unifiedUtilities.gs', 'utf8');
  let context;

  beforeEach(() => {
    context = {
      console,
      debugLog: () => {},
      clearExecutionUserInfoCache: jest.fn(),
      Session: { getActiveUser: () => ({ getEmail: () => 'admin@example.com' }) },
      Utilities: { getUuid: () => `test-uuid-${Math.random()}` },
      findUserById: jest.fn(() => ({
        adminEmail: 'admin@example.com',
        configJson: JSON.stringify({ appPublished: false }),
      })),
      fetchUserFromDatabase: jest.fn(() => ({ adminEmail: 'admin@example.com' })),
    };
    vm.createContext(context);
    vm.runInContext(unifiedUtilitiesCode, context);
    vm.runInContext(coreCode, context);
  });

  test('allows access for matching admin email', () => {
    expect(() => context.verifyUserAccess('U1')).not.toThrow();
  });

  test('denies access when emails differ and board not published', () => {
    context.Session.getActiveUser = () => ({ getEmail: () => 'other@example.com' });
    expect(() => context.verifyUserAccess('U1')).toThrow('権限エラー');
  });

  test('allows read-only access for published board', () => {
    context.Session.getActiveUser = () => ({ getEmail: () => 'viewer@example.com' });
    context.findUserById.mockReturnValue({
      adminEmail: 'admin@example.com',
      configJson: JSON.stringify({ appPublished: true }),
    });
    expect(() => context.verifyUserAccess('U1')).not.toThrow();
  });

  test('denies access when user missing in database', () => {
    context.fetchUserFromDatabase.mockReturnValue(null);
    expect(() => context.verifyUserAccess('U1')).toThrow('認証エラー');
    expect(context.fetchUserFromDatabase).toHaveBeenCalled();
  });
});
