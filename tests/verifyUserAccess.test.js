const fs = require('fs');
const vm = require('vm');

describe('verifyUserAccess security checks', () => {
  const coreCode = fs.readFileSync('src/Core.gs', 'utf8');
  let context;

  beforeEach(() => {
    context = {
      console,
      debugLog: () => {},
      Log: {
        debug: jest.fn(),
        info: jest.fn(),
        warn: jest.fn(),
        error: jest.fn()
      },
      DB: {
        findUserById: jest.fn(() => ({
          adminEmail: 'admin@example.com',
          configJson: JSON.stringify({ appPublished: false }),
        }))
      },
      clearExecutionUserInfoCache: jest.fn(),
      Session: { getActiveUser: () => ({ getEmail: () => 'admin@example.com' }) },
      findUserById: jest.fn(() => ({
        adminEmail: 'admin@example.com',
        configJson: JSON.stringify({ appPublished: false }),
      })),
    };
    vm.createContext(context);
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
    context.DB.findUserById.mockReturnValue({
      adminEmail: 'admin@example.com',
      configJson: JSON.stringify({ appPublished: true }),
    });
    context.findUserById.mockReturnValue({
      adminEmail: 'admin@example.com',
      configJson: JSON.stringify({ appPublished: true }),
    });
    expect(() => context.verifyUserAccess('U1')).not.toThrow();
  });
});

