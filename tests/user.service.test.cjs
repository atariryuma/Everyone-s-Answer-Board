const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function createMockCache(initial = {}) {
  const store = new Map(Object.entries(initial));
  const removed = [];
  return {
    get: (k) => (store.has(k) ? store.get(k) : null),
    put: (k, v) => { store.set(k, v); },
    remove: (k) => { store.delete(k); removed.push(k); },
    _store: store,
    _removed: removed
  };
}

function loadUserServiceContext(overrides = {}) {
  const cache = overrides.cache || createMockCache();
  const activeEmail = overrides.activeEmail !== undefined ? overrides.activeEmail : 'actor@example.com';

  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    logError_: () => {},
    CacheService: { getScriptCache: () => cache },
    CACHE_DURATION: { LONG: 300 },
    SYSTEM_LIMITS: { MAX_LOCK_ROWS: 100 },
    Session: {
      getActiveUser: () => ({ getEmail: () => activeEmail })
    },
    ScriptApp: { getIdentityToken: () => null },
    Utilities: {
      newBlob: () => ({ getDataAsString: () => '{}' }),
      base64DecodeWebSafe: () => new Uint8Array()
    },
    validateUrl: (url) => ({ isValid: typeof url === 'string' && url.startsWith('https://') }),
    validateEmail: (email) => typeof email === 'string' && /.+@.+/.test(email),
    getCurrentEmail: () => activeEmail,
    findUserByEmail_: overrides.findUserByEmail_ || (() => null),
    findUserById_: () => null,
    findUserByGoogleId_: () => null,
    openSpreadsheet: () => null,
    updateUser_: () => ({ success: true }),
    getUserConfig: () => ({ success: true, config: {} }),
    getConfigOrDefault: overrides.getConfigOrDefault || (() => ({})),
    isAdministrator: overrides.isAdministrator || (() => false),
    clearConfigCache_: overrides.clearConfigCache_ || (() => {}),
    createExceptionResponse: (error) => ({ success: false, message: error.message }),
    getWebAppUrl: overrides.getWebAppUrl || (() => 'https://script.google.com/macros/s/abc/exec'),
    getUserCacheKeys_: overrides.getUserCacheKeys_ || ((userId) => [`user_kv_${userId}`]),
    ...overrides
  };

  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/UserService.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'UserService.js' });
  context._cache = cache;
  return context;
}

// =====================================================================
// getUserInfoCacheKey
// =====================================================================

test('getUserInfoCacheKey: builds a key scoped by email', () => {
  const ctx = loadUserServiceContext();
  assert.equal(ctx.getUserInfoCacheKey('a@x.com'), 'current_user_info_a@x.com');
});

// =====================================================================
// isAdmin
// =====================================================================

test('isAdmin: returns true for administrator email', () => {
  const ctx = loadUserServiceContext({
    isAdministrator: (email) => email === 'admin@example.com',
    activeEmail: 'admin@example.com'
  });
  assert.equal(ctx.isAdmin(), true);
});

test('isAdmin: returns false for non-admin email', () => {
  const ctx = loadUserServiceContext({
    isAdministrator: () => false
  });
  assert.equal(ctx.isAdmin(), false);
});

test('isAdmin: returns false when no active email', () => {
  const ctx = loadUserServiceContext({
    getCurrentEmail: () => null
  });
  assert.equal(ctx.isAdmin(), false);
});

test('isAdmin: returns false on exception from isAdministrator', () => {
  const ctx = loadUserServiceContext({
    isAdministrator: () => { throw new Error('boom'); }
  });
  assert.equal(ctx.isAdmin(), false);
});
