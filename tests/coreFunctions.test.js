const fs = require('fs');
const vm = require('vm');

describe('Core.gs utilities', () => {
  const code = fs.readFileSync('src/Core.gs', 'utf8');
  let context;
  beforeEach(() => {
    context = { 
      debugLog: () => {}, 
      console,
      cacheManager: {
        remove: jest.fn(),
        invalidateDependents: jest.fn(),
        clearByPattern: jest.fn()
      },
      Utilities: {
        sleep: jest.fn(),
        getUuid: () => 'UUID'
      },
      fetchUserFromDatabase: jest.fn(),
      comprehensiveUserSearch: jest.fn(),
      registerNewUser: jest.fn(() => ({ userId: 'UUID', adminUrl: 'admin', viewUrl: 'view' }))
    };
    vm.createContext(context);
    vm.runInContext(code, context);
  });

  describe('getOpinionHeaderSafely', () => {
    beforeEach(() => {
      context.findUserById = jest.fn();
    });

    test('returns default when user not found', () => {
      context.findUserById.mockReturnValue(null);
      const header = context.getOpinionHeaderSafely('uid', 'Sheet1');
      expect(header).toBe('お題');
    });

    test('returns configured header', () => {
      context.findUserById.mockReturnValue({
        configJson: JSON.stringify({
          publishedSheetName: 'Sheet1',
          sheet_Sheet1: { opinionHeader: 'テーマ' }
        })
      });
      const header = context.getOpinionHeaderSafely('uid', 'Sheet1');
      expect(header).toBe('テーマ');
    });
  });

  describe('registerNewUser', () => {
    beforeEach(() => {
      Object.assign(context, {
        Session: { getActiveUser: () => ({ getEmail: () => 'admin@example.com' }) },
        findUserByEmail: jest.fn(),
        createUser: jest.fn(),
        updateUser: jest.fn(),
        invalidateUserCache: jest.fn(),
        generateAppUrls: jest.fn(() => ({ adminUrl: 'admin', viewUrl: 'view' })),
        getDeployUserDomainInfo: jest.fn(() => ({ 
          currentDomain: 'example.com', 
          deployDomain: 'example.com', 
          isDomainMatch: true 
        })),
        findUserById: jest.fn(() => ({ userId: 'UUID', adminEmail: 'admin@example.com' }))
      });
    });

    test('creates user when none exists', () => {
      context.findUserByEmail.mockReturnValue(null);
      const res = context.registerNewUser('admin@example.com');
      expect(context.createUser).toHaveBeenCalled();
      expect(res.isExistingUser).toBe(false);
    });

    test('updates existing user', () => {
      context.findUserByEmail.mockReturnValue({ userId: 'U', configJson: '{}', spreadsheetId: 'SS' });
      const res = context.registerNewUser('admin@example.com');
      expect(context.updateUser).toHaveBeenCalled();
      expect(res.isExistingUser).toBe(true);
    });
  });
});
