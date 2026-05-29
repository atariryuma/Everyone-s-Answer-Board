const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function createMockCache(initial = {}) {
  const store = new Map(Object.entries(initial));
  return {
    get: (k) => (store.has(k) ? store.get(k) : null),
    put: (k, v) => { store.set(k, v); },
    remove: (k) => { store.delete(k); },
    _store: store
  };
}

function loadDatabaseContext(overrides = {}) {
  const cache = overrides.cache || createMockCache();
  const propsStore = overrides.propsStore || {};

  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    CacheService: { getScriptCache: () => cache },
    CACHE_DURATION: {
      SHORT: 10, MEDIUM: 30, FORM_DATA: 30, LONG: 300,
      DATABASE_LONG: 600, USER_INDIVIDUAL: 900, EXTRA_LONG: 3600
    },
    TIMEOUT_MS: { DEFAULT: 5000 },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (k) => (k in propsStore ? propsStore[k] : null),
        setProperty: (k, v) => { propsStore[k] = v; },
        getProperties: () => ({ ...propsStore })
      })
    },
    SpreadsheetApp: overrides.SpreadsheetApp || {
      openById: () => { throw new Error('SpreadsheetApp.openById not stubbed'); }
    },
    UrlFetchApp: overrides.UrlFetchApp || {
      fetch: () => { throw new Error('UrlFetchApp.fetch not stubbed'); }
    },
    Utilities: overrides.Utilities || {
      base64EncodeWebSafe: (s) => Buffer.from(s).toString('base64').replace(/\+/g, '-').replace(/\//g, '_'),
      computeRsaSha256Signature: () => new Uint8Array([1, 2, 3, 4]),
      sleep: () => {}
    },
    Session: {
      getActiveUser: () => ({ getEmail: () => overrides.currentEmail || 'actor@example.com' })
    },
    // Real validateEmail returns { isValid, sanitized, errors }; stub matches that shape.
    validateEmail: (email) => ({
      isValid: typeof email === 'string' && /.+@.+/.test(email),
      sanitized: email,
      errors: []
    }),
    getCurrentEmail: overrides.getCurrentEmail || (() => overrides.currentEmail || 'actor@example.com'),
    isAdministrator: overrides.isAdministrator || (() => false),
    getUserConfig: () => ({ success: true, config: {} }),
    executeWithRetry: (fn) => fn(),
    getCachedProperty: (k) => (k in propsStore ? propsStore[k] : null),
    clearPropertyCache: () => {},
    simpleHash: (s) => String(s).length,
    saveToCacheWithSizeCheck: (k, data) => { cache.put(k, JSON.stringify(data)); return true; },
    safeJsonParse_: (text, fallback) => {
      const fb = (fallback === undefined ? null : fallback);
      if (text == null || text === '') return fb;
      if (typeof text === 'object') return text;
      try { return JSON.parse(text); } catch (_) { return fb; }
    },
    logError_: () => {},
    ...overrides
  };

  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/DatabaseCore.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'DatabaseCore.js' });
  context._cache = cache;
  context._propsStore = propsStore;
  return context;
}

// =====================================================================
// openSpreadsheet — input validation
// =====================================================================

test('openSpreadsheet: returns null for invalid spreadsheet id', () => {
  const ctx = loadDatabaseContext();
  assert.equal(ctx.openSpreadsheet(null), null);
  assert.equal(ctx.openSpreadsheet(''), null);
  assert.equal(ctx.openSpreadsheet(undefined), null);
  assert.equal(ctx.openSpreadsheet(42), null);
});

// =====================================================================
// openSpreadsheet — normal (non-SA) path
// =====================================================================

test('openSpreadsheet: non-SA path returns dataAccess with spreadsheet', () => {
  const mockSpreadsheet = { id: 'ss', getSheetByName: (n) => ({ name: n }) };
  const ctx = loadDatabaseContext({
    SpreadsheetApp: { openById: () => mockSpreadsheet }
  });

  const result = ctx.openSpreadsheet('ss-abc', { useServiceAccount: false });

  assert.ok(result);
  assert.equal(result.spreadsheet, mockSpreadsheet);
  assert.equal(typeof result.getSheet, 'function');
});

test('openSpreadsheet: returns null when SpreadsheetApp.openById throws', () => {
  const ctx = loadDatabaseContext({
    SpreadsheetApp: {
      openById: () => { throw new Error('spreadsheet not found'); }
    }
  });

  const result = ctx.openSpreadsheet('ss-abc', { useServiceAccount: false });
  assert.equal(result, null);
});

test('openSpreadsheet: dataAccess.getSheet delegates to spreadsheet.getSheetByName', () => {
  let calledWith = null;
  const mockSpreadsheet = {
    getSheetByName: (name) => { calledWith = name; return { name }; }
  };
  const ctx = loadDatabaseContext({
    SpreadsheetApp: { openById: () => mockSpreadsheet }
  });

  const dataAccess = ctx.openSpreadsheet('ss-abc', { useServiceAccount: false });
  const sheet = dataAccess.getSheet('Sheet1');

  assert.equal(calledWith, 'Sheet1');
  assert.equal(sheet.name, 'Sheet1');
});

test('openSpreadsheet: getSheet returns null when sheet name missing', () => {
  const ctx = loadDatabaseContext({
    SpreadsheetApp: { openById: () => ({ getSheetByName: () => ({}) }) }
  });

  const dataAccess = ctx.openSpreadsheet('ss', { useServiceAccount: false });
  assert.equal(dataAccess.getSheet(''), null);
  assert.equal(dataAccess.getSheet(null), null);
});

test('openSpreadsheet: getSheet returns null when getSheetByName throws', () => {
  const ctx = loadDatabaseContext({
    SpreadsheetApp: {
      openById: () => ({
        getSheetByName: () => { throw new Error('no access'); }
      })
    }
  });

  const dataAccess = ctx.openSpreadsheet('ss', { useServiceAccount: false });
  assert.equal(dataAccess.getSheet('Sheet1'), null);
});

// =====================================================================
// openSpreadsheet — service account gating (validateServiceAccountUsage)
// =====================================================================

test('openSpreadsheet: non-DB sheet without owner/admin is denied (unpublished board)', () => {
  // v2782+: 非 admin & 非 owner で未公開ボードへの SA アクセスは validation で deny される。
  const ctx = loadDatabaseContext({
    propsStore: { DATABASE_SPREADSHEET_ID: 'db-id-123' },
    SpreadsheetApp: { openById: () => ({ id: 'user-sheet' }) },
    isAdministrator: () => false
  });
  // vm.runInContext で source の関数宣言が override を上書きするため、 load 後に再設定。
  ctx.findUserBySpreadsheetId = () => ({ userId: 'u1', userEmail: 'owner@example.com' });
  ctx.getUserConfig = () => ({ success: true, config: { isPublished: false } });

  const result = ctx.openSpreadsheet('different-sheet-id', { useServiceAccount: true });

  assert.equal(result, null, 'unpublished cross-user access must be denied');
});

test('openSpreadsheet: non-DB sheet for owner uses own access (openById)', () => {
  // v2782+: owner が自分の board にアクセス → accessMode='own' → openById (own OAuth) を使う。
  const mockSpreadsheet = { id: 'owner-sheet', getSheetByName: () => null };
  const ctx = loadDatabaseContext({
    propsStore: { DATABASE_SPREADSHEET_ID: 'db-id-123' },
    SpreadsheetApp: { openById: () => mockSpreadsheet },
    isAdministrator: () => false,
    currentEmail: 'owner@example.com'
  });
  ctx.findUserBySpreadsheetId = () => ({ userId: 'u1', userEmail: 'owner@example.com' });

  const result = ctx.openSpreadsheet('owner-sheet-id', { useServiceAccount: true });

  assert.ok(result, 'owner access should succeed via openById');
  assert.equal(result.spreadsheet, mockSpreadsheet);
  assert.equal(result.accessMode, 'own');
});

test('openSpreadsheet: non-DB user board refuses owner openById fallback when SA pool unavailable', () => {
  // v2782+: 公開 board への非 owner 非 admin アクセス → accessMode='sa'。
  // セキュリティ: SA pool 全滅時に user board を deploy 主の openById で読むのは
  // 「viewer は board SS に直接権限を持たない」モデルを破る fail-open。
  // user board (= DB id と異なる SS) では owner fallback を拒否し null を返す。
  const mockSpreadsheet = { id: 'viewer-sheet', getSheetByName: () => null };
  const ctx = loadDatabaseContext({
    propsStore: { DATABASE_SPREADSHEET_ID: 'db-id-123' }, // SERVICE_ACCOUNT_CREDS absent → SA pool unavailable
    SpreadsheetApp: { openById: () => mockSpreadsheet },
    isAdministrator: () => false,
    currentEmail: 'viewer@example.com'
  });
  ctx.findUserBySpreadsheetId = () => ({ userId: 'u1', userEmail: 'owner@example.com' });
  ctx.getUserConfig = () => ({ success: true, config: { isPublished: true } });

  const result = ctx.openSpreadsheet('different-sheet-id', { useServiceAccount: true });

  assert.equal(result, null, 'user board access must NOT fall back to owner openById when SA pool is down');
});

test('openSpreadsheet: matching DB id + useServiceAccount=true falls through to SpreadsheetApp when no credentials', () => {
  const mockSpreadsheet = { id: 'db', getSheetByName: () => null };
  const ctx = loadDatabaseContext({
    propsStore: { DATABASE_SPREADSHEET_ID: 'db-id-123' }, // SERVICE_ACCOUNT_CREDS absent
    SpreadsheetApp: { openById: () => mockSpreadsheet }
  });

  const result = ctx.openSpreadsheet('db-id-123', { useServiceAccount: true });

  assert.ok(result);
  assert.equal(result.spreadsheet, mockSpreadsheet);
  assert.equal(result.auth.isValid, false);
});

// =====================================================================
// getServiceAccount
// =====================================================================

test('getServiceAccount: returns invalid when no credentials configured', () => {
  const ctx = loadDatabaseContext();
  const result = ctx.getServiceAccount();
  assert.equal(result.isValid, false);
});

test('getServiceAccount: returns invalid when credentials are malformed JSON', () => {
  const ctx = loadDatabaseContext({
    propsStore: { SERVICE_ACCOUNT_CREDS: 'not-json' }
  });
  const result = ctx.getServiceAccount();
  assert.equal(result.isValid, false);
});

test('getServiceAccount: rejects credentials missing required fields', () => {
  const ctx = loadDatabaseContext({
    propsStore: {
      SERVICE_ACCOUNT_CREDS: JSON.stringify({ client_email: 'sa@ex.com' }) // no private_key, no type
    }
  });
  const result = ctx.getServiceAccount();
  assert.equal(result.isValid, false);
});

test('getServiceAccount: rejects credentials with invalid email format', () => {
  const ctx = loadDatabaseContext({
    propsStore: {
      SERVICE_ACCOUNT_CREDS: JSON.stringify({
        client_email: 'not-an-email',
        private_key: '-----BEGIN PRIVATE KEY-----\nXXX\n-----END PRIVATE KEY-----',
        type: 'service_account'
      })
    }
  });
  const result = ctx.getServiceAccount();
  assert.equal(result.isValid, false);
});

test('getServiceAccount: rejects credentials with malformed private key', () => {
  const ctx = loadDatabaseContext({
    propsStore: {
      SERVICE_ACCOUNT_CREDS: JSON.stringify({
        client_email: 'sa@ex.com',
        private_key: 'not-a-key',
        type: 'service_account'
      })
    }
  });
  const result = ctx.getServiceAccount();
  assert.equal(result.isValid, false);
});

test('getServiceAccount: accepts valid credentials', () => {
  const ctx = loadDatabaseContext({
    propsStore: {
      SERVICE_ACCOUNT_CREDS: JSON.stringify({
        client_email: 'sa@project.iam.gserviceaccount.com',
        private_key: '-----BEGIN PRIVATE KEY-----\nMIIE...\n-----END PRIVATE KEY-----',
        type: 'service_account'
      })
    }
  });
  const result = ctx.getServiceAccount();
  assert.equal(result.isValid, true);
  assert.equal(result.email, 'sa@project.iam.gserviceaccount.com');
});

// =====================================================================
// validateServiceAccountUsage
// =====================================================================

test('validateServiceAccountUsage: useServiceAccount=false always allowed', () => {
  const ctx = loadDatabaseContext();
  const result = ctx.validateServiceAccountUsage('any-id', false);
  assert.equal(result.allowed, true);
});

test('validateServiceAccountUsage: database spreadsheet id always allowed', () => {
  const ctx = loadDatabaseContext({
    propsStore: { DATABASE_SPREADSHEET_ID: 'db-id' }
  });
  const result = ctx.validateServiceAccountUsage('db-id', true);
  assert.equal(result.allowed, true);
});

test('validateServiceAccountUsage: admin allowed on any sheet (cross-user via SA)', () => {
  // v2782+: admin が他人のボードにアクセスするときは accessMode='sa'。
  // 自分が owner のときは 'own' に上書きされる (owner check が admin より優先)。
  const ctx = loadDatabaseContext({
    propsStore: { DATABASE_SPREADSHEET_ID: 'db-id' },
    isAdministrator: () => true,
    currentEmail: 'admin@example.com'
  });
  ctx.findUserBySpreadsheetId = () => ({ userId: 'u1', userEmail: 'someone-else@example.com' });
  const result = ctx.validateServiceAccountUsage('other-sheet', true);
  assert.equal(result.allowed, true);
  assert.equal(result.accessMode, 'sa');
});

// =====================================================================
// openDatabase
// =====================================================================

test('openDatabase: returns null when DATABASE_SPREADSHEET_ID not configured', () => {
  const ctx = loadDatabaseContext();
  const result = ctx.openDatabase();
  assert.equal(result, null);
});

test('openDatabase: returns spreadsheet via SpreadsheetApp fallback when SA fails', () => {
  const mockSpreadsheet = { id: 'db-sheet' };
  const ctx = loadDatabaseContext({
    propsStore: { DATABASE_SPREADSHEET_ID: 'db-id' },
    SpreadsheetApp: { openById: () => mockSpreadsheet }
    // No SERVICE_ACCOUNT_CREDS → SA path fails, falls back
  });

  const result = ctx.openDatabase();
  assert.equal(result, mockSpreadsheet);
});

test('openDatabase: returns null when both SA and SpreadsheetApp fail', () => {
  const ctx = loadDatabaseContext({
    propsStore: { DATABASE_SPREADSHEET_ID: 'db-id' },
    SpreadsheetApp: {
      openById: () => { throw new Error('denied'); }
    }
  });

  const result = ctx.openDatabase();
  assert.equal(result, null);
});
