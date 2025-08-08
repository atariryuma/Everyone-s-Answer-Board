const fs = require('fs');
const vm = require('vm');

describe('Core.gs utilities', () => {
  const urlCode = fs.readFileSync('src/url.gs', 'utf8');
  const mainCode = fs.readFileSync('src/main.gs', 'utf8');
  const commonUtilitiesCode = fs.readFileSync('src/commonUtilities.gs', 'utf8');
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
          }
        })
      },
      CacheService: {
        getUserCache: () => ({
          get: () => null,
          put: () => {}
        }),
        getScriptCache: () => ({
          get: (k) => store[k] || null,
          put: (k, v) => { store[k] = v; },
          remove: (k) => { delete store[k]; }
        })
      },
      cacheManager: {
        store,
        get(key, fn) {
          if (this.store[key]) return this.store[key];
          const val = fn();
          this.store[key] = val;
          return val;
        },
        remove(key) { delete this.store[key]; },
        clearByPattern: (pattern) => {
          for (const key in store) {
            if (key.startsWith(pattern.replace('*', ''))) {
              delete store[key];
            }
          }
        }
      },
      ScriptApp: {
        getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/ID/exec' }),
        getScriptId: () => 'ID'
      },
      Session: {
        getActiveUser: () => ({ getEmail: () => 'test@example.com' })
      },
      Utilities: {
        getUuid: () => 'mock-uuid',
        computeDigest: () => [],
        Charset: { UTF_8: 'UTF-8' }
      }
    };
    vm.createContext(context);
    vm.runInContext(urlCode, context);
    vm.runInContext(mainCode, context);
    vm.runInContext(commonUtilitiesCode, context);
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
        generateUserUrls: jest.fn(() => ({ adminUrl: 'admin', viewUrl: 'view' })),
        getDeployUserDomainInfo: jest.fn(() => ({ deployDomain: '', isDomainMatch: true }))
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
