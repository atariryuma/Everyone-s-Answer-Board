const fs = require('fs');
const vm = require('vm');

describe('Core.gs utilities', () => {
  const code = fs.readFileSync('src/Core.gs', 'utf8');
  let context;
  beforeEach(() => {
    context = { debugLog: () => {}, console };
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
        Utilities: { getUuid: () => 'UUID' },
        findUserByEmail: jest.fn(),
        createUser: jest.fn(),
        updateUser: jest.fn(),
        invalidateUserCache: jest.fn(),
        generateAppUrls: jest.fn(() => ({ adminUrl: 'admin', viewUrl: 'view' }))
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
