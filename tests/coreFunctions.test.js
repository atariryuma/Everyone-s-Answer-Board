const fs = require('fs');
const vm = require('vm');

describe('Core.gs utilities', () => {
  const urlCode = fs.readFileSync('src/url.gs', 'utf8');
  const mainCode = fs.readFileSync('src/main.gs', 'utf8');
  const unifiedUtilitiesCode = fs.readFileSync('src/unifiedUtilities.gs', 'utf8');
  const coreCode = fs.readFileSync('src/Core.gs', 'utf8');
  let context;
  beforeEach(() => {
    const store = {};
    context = {
      debugLog: () => {},
      errorLog: () => {},
      infoLog: () => {},
      warnLog: () => {},
      console,
      PropertiesService: {
        getScriptProperties: () => ({
          getProperty: (key) => {
            if (key === 'DEBUG_MODE') return 'false';
            if (key === 'DATABASE_SPREADSHEET_ID') return 'mock-spreadsheet-id';
            return null;
          },
        }),
      },
      CacheService: {
        getUserCache: () => ({
          get: () => null,
          put: () => {},
        }),
        getScriptCache: () => ({
          get: (k) => store[k] || null,
          put: (k, v) => {
            store[k] = v;
          },
          remove: (k) => {
            delete store[k];
          },
        }),
      },
      cacheManager: {
        store,
        get(key, fn) {
          if (this.store[key]) return this.store[key];
          const val = fn();
          this.store[key] = val;
          return val;
        },
        remove(key) {
          delete this.store[key];
        },
        clearByPattern: (pattern) => {
          for (const key in store) {
            if (key.startsWith(pattern.replace('*', ''))) {
              delete store[key];
            }
          }
        },
      },
      ScriptApp: {
        getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/ID/exec' }),
        getScriptId: () => 'ID',
      },
      Session: {
        getActiveUser: () => ({ getEmail: () => 'test@example.com' }),
      },
      Utilities: {
        getUuid: () => 'mock-uuid',
        computeDigest: () => [],
        Charset: { UTF_8: 'UTF-8' },
      },
    };
    vm.createContext(context);
    vm.runInContext(urlCode, context);
    vm.runInContext(mainCode, context);
    vm.runInContext(unifiedUtilitiesCode, context);
    vm.runInContext(coreCode, context);
  });

  describe('getOpinionHeaderSafely', () => {
    beforeEach(() => {
      context.findUserById = jest.fn();
      // Clear cache before each test
      context.cacheManager.store = {};
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
          sheet_Sheet1: { opinionHeader: 'テーマ' },
        }),
      });
      const header = context.getOpinionHeaderSafely('uid', 'Sheet1');
      expect(header).toBe('テーマ');
    });
  });

  describe('registerNewUser', () => {
    beforeEach(() => {
      let createdUserData = null; // To store the user created by createUser

      Object.assign(context, {
        debugLog: (...args) => console.log('[DEBUG]', ...args), // Add debugLog
        infoLog: (...args) => console.log('[INFO]', ...args), // Add infoLog
        errorLog: (...args) => console.log('[ERROR]', ...args), // Add errorLog
        warnLog: (...args) => console.log('[WARN]', ...args), // Add warnLog
        console, // Keep console for direct logging

        Session: { getActiveUser: () => ({ getEmail: () => 'admin@example.com' }) },
        Utilities: {
          getUuid: () => 'UUID',
          computeDigest: () => [],
          Charset: { UTF_8: 'UTF-8' },
          sleep: jest.fn(),
        },
        findUserByEmail: jest.fn(),
        createUser: jest.fn((userData) => {
          // Modify createUser mock
          console.log('MOCK: createUser called with:', userData); // Debug log
          createdUserData = userData; // Store the created user data
        }),
        updateUser: jest.fn(),
        invalidateUserCache: jest.fn(),
        generateUserUrls: jest.fn(() => ({ adminUrl: 'admin', viewUrl: 'view' })),
        getDeployUserDomainInfo: jest.fn(() => ({ deployDomain: '', isDomainMatch: true })),
        logDatabaseError: jest.fn(),
        waitForUserRecord: jest.fn(() => true),
        fetchUserFromDatabase: jest.fn((field, value, options) => {
          // Modify fetchUserFromDatabase mock
          console.log('MOCK: fetchUserFromDatabase called with:', { field, value, options }); // Debug log
          if (
            createdUserData &&
            ((field === 'userId' && value === createdUserData.userId) ||
              (field === 'adminEmail' && value === createdUserData.adminEmail))
          ) {
            console.log('MOCK: fetchUserFromDatabase returning createdUserData:', createdUserData); // Debug log
            return createdUserData; // Return the created user if found
          }
          console.log('MOCK: fetchUserFromDatabase returning null'); // Debug log
          return null; // Otherwise, return null
        }),
      });
    });

    test('creates user when none exists', () => {
      context.findUserByEmail.mockReturnValue(null);
      const res = context.registerNewUser('admin@example.com');
      expect(context.createUser).toHaveBeenCalled();
      expect(res.isExistingUser).toBe(false);
    });

    test('updates existing user', () => {
      context.findUserByEmail.mockReturnValue({
        userId: 'U',
        configJson: '{}',
        spreadsheetId: 'SS',
      });
      const res = context.registerNewUser('admin@example.com');
      expect(context.updateUser).toHaveBeenCalled();
      expect(res.isExistingUser).toBe(true);
    });

    test('logs and throws when createUser fails', () => {
      context.findUserByEmail.mockReturnValue(null);
      context.createUser.mockImplementation(() => {
        throw new Error('General DB Error');
      });
      expect(() => context.registerNewUser('admin@example.com')).toThrow(
        'ユーザー登録に失敗しました。システム管理者に連絡してください。'
      );
    });
  });
});
