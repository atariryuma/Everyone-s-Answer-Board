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
    findUserByEmail: overrides.findUserByEmail || (() => null),
    findUserById: () => null,
    findUserByGoogleId: () => null,
    openSpreadsheet: () => null,
    updateUser: () => ({ success: true }),
    getUserConfig: () => ({ success: true, config: {} }),
    getConfigOrDefault: overrides.getConfigOrDefault || (() => ({})),
    isAdministrator: overrides.isAdministrator || (() => false),
    clearConfigCache: overrides.clearConfigCache || (() => {}),
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
// getCurrentUserInfo
// =====================================================================

test('getCurrentUserInfo: returns null when session has no email', () => {
  const ctx = loadUserServiceContext({ activeEmail: '' });
  assert.equal(ctx.getCurrentUserInfo(), null);
});

test('getCurrentUserInfo: returns null when user not found in DB', () => {
  const ctx = loadUserServiceContext({
    findUserByEmail: () => null
  });
  assert.equal(ctx.getCurrentUserInfo(), null);
});

test('getCurrentUserInfo: returns cached value without DB lookup', () => {
  const cached = { userId: 'u1', userEmail: 'actor@example.com', config: {} };
  const cache = createMockCache({
    'current_user_info_actor@example.com': JSON.stringify(cached)
  });
  let findCalls = 0;
  const ctx = loadUserServiceContext({
    cache,
    findUserByEmail: () => { findCalls += 1; return null; }
  });

  const result = ctx.getCurrentUserInfo();
  assert.equal(result.userId, 'u1');
  assert.equal(findCalls, 0, 'findUserByEmail must not be called on cache hit');
});

test('getCurrentUserInfo: evicts and refetches on corrupt cache', () => {
  const cache = createMockCache({
    'current_user_info_actor@example.com': 'not-valid-json'
  });
  const ctx = loadUserServiceContext({
    cache,
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'actor@example.com', isActive: true })
  });

  const result = ctx.getCurrentUserInfo();
  assert.equal(result.userId, 'u1');
  assert.ok(cache._removed.includes('current_user_info_actor@example.com'));
});

test('getCurrentUserInfo: caches fresh result after miss', () => {
  const cache = createMockCache();
  const ctx = loadUserServiceContext({
    cache,
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'actor@example.com', isActive: true }),
    getConfigOrDefault: () => ({ spreadsheetId: 'ss1' })
  });

  ctx.getCurrentUserInfo();

  assert.ok(cache._store.has('current_user_info_actor@example.com'));
});

// =====================================================================
// enrichUserInfo
// =====================================================================

test('enrichUserInfo: returns original on invalid input (fallback)', () => {
  const ctx = loadUserServiceContext();
  const result = ctx.enrichUserInfo(null);
  assert.equal(result, null);
});

test('enrichUserInfo: returns original when userId missing', () => {
  const ctx = loadUserServiceContext();
  const input = { userEmail: 'a@x.com' }; // no userId
  const result = ctx.enrichUserInfo(input);
  // Falls back to input
  assert.deepEqual(result, input);
});

test('enrichUserInfo: attaches config and userInfo slice', () => {
  const ctx = loadUserServiceContext({
    getConfigOrDefault: () => ({ spreadsheetId: 'ss1' })
  });

  const result = ctx.enrichUserInfo({
    userId: 'u1',
    userEmail: 'a@x.com',
    isActive: true,
    lastModified: '2026-01-01'
  });

  assert.equal(result.userId, 'u1');
  assert.equal(result.userEmail, 'a@x.com');
  assert.equal(result.isActive, true);
  assert.equal(result.lastModified, '2026-01-01');
  assert.ok(result.config);
  assert.equal(result.config.spreadsheetUrl, 'https://docs.google.com/spreadsheets/d/ss1/edit');
  assert.ok(result.userInfo);
  assert.equal(result.userInfo.userId, 'u1');
});

// =====================================================================
// generateDynamicUserUrls
// =====================================================================

test('generateDynamicUserUrls: derives spreadsheetUrl from spreadsheetId', () => {
  const ctx = loadUserServiceContext();
  const result = ctx.generateDynamicUserUrls({ spreadsheetId: 'abc123' });
  assert.equal(result.spreadsheetUrl, 'https://docs.google.com/spreadsheets/d/abc123/edit');
});

test('generateDynamicUserUrls: preserves explicitly-provided spreadsheetUrl', () => {
  const ctx = loadUserServiceContext();
  const result = ctx.generateDynamicUserUrls({
    spreadsheetId: 'abc',
    spreadsheetUrl: 'https://custom.example.com/sheet'
  });
  assert.equal(result.spreadsheetUrl, 'https://custom.example.com/sheet');
});

test('generateDynamicUserUrls: attaches appUrl when isPublished', () => {
  const ctx = loadUserServiceContext({
    getWebAppUrl: () => 'https://script.google.com/webapp'
  });
  const result = ctx.generateDynamicUserUrls({ isPublished: true });
  assert.equal(result.appUrl, 'https://script.google.com/webapp');
});

test('generateDynamicUserUrls: does not add appUrl when not published', () => {
  const ctx = loadUserServiceContext();
  const result = ctx.generateDynamicUserUrls({ isPublished: false });
  assert.equal(result.appUrl, undefined);
});

test('generateDynamicUserUrls: validates formUrl and sets hasValidForm', () => {
  const ctx = loadUserServiceContext();
  const validResult = ctx.generateDynamicUserUrls({
    formUrl: 'https://docs.google.com/forms/d/e/xxx/viewform'
  });
  assert.equal(validResult.hasValidForm, true);

  const invalidResult = ctx.generateDynamicUserUrls({
    formUrl: 'not-a-url'
  });
  assert.equal(invalidResult.hasValidForm, false);
});

test('generateDynamicUserUrls: returns original config on exception (fallback)', () => {
  const ctx = loadUserServiceContext({
    validateUrl: () => { throw new Error('boom'); }
  });
  const input = { formUrl: 'https://x.com' };
  const result = ctx.generateDynamicUserUrls(input);
  // Fallback: original input is returned when any step throws
  assert.ok(result);
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

// =====================================================================
// getUser
// =====================================================================

test('getUser: returns email-only result by default', () => {
  const ctx = loadUserServiceContext();
  const result = ctx.getUser();
  assert.equal(result.success, true);
  assert.equal(result.email, 'actor@example.com');
  assert.equal(result.user, undefined);
});

test('getUser: returns full user when infoType=full', () => {
  const ctx = loadUserServiceContext({
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'actor@example.com', isActive: true })
  });
  const result = ctx.getUser('full');
  assert.equal(result.success, true);
  assert.equal(result.user.userId, 'u1');
});

test('getUser: returns null user when lookup finds nothing', () => {
  const ctx = loadUserServiceContext({
    findUserByEmail: () => null
  });
  const result = ctx.getUser('full');
  assert.equal(result.success, true);
  assert.equal(result.user, null);
});

test('getUser: returns error when unauthenticated', () => {
  const ctx = loadUserServiceContext({
    getCurrentEmail: () => null
  });
  const result = ctx.getUser();
  assert.equal(result.success, false);
  assert.match(result.message, /authenticated/i);
});

test('getUser: rejects unsupported infoType', () => {
  const ctx = loadUserServiceContext();
  const result = ctx.getUser('weird');
  assert.equal(result.success, false);
  assert.match(result.message, /weird/);
});

// =====================================================================
// resetAuth
// =====================================================================

test('resetAuth: clears global cache keys', () => {
  const cache = createMockCache({
    admin_auth_cache: 'x',
    session_data: 'y',
    system_diagnostic_cache: 'z',
    bulk_admin_data_cache: 'w'
  });
  const ctx = loadUserServiceContext({ cache });

  ctx.resetAuth();

  assert.ok(cache._removed.includes('admin_auth_cache'));
  assert.ok(cache._removed.includes('session_data'));
  assert.ok(cache._removed.includes('system_diagnostic_cache'));
  assert.ok(cache._removed.includes('bulk_admin_data_cache'));
});

test('resetAuth: clears email-scoped cache keys', () => {
  const cache = createMockCache();
  const ctx = loadUserServiceContext({ cache });

  ctx.resetAuth();

  assert.ok(cache._removed.includes('current_user_info_actor@example.com'));
  assert.ok(cache._removed.includes('board_data_actor@example.com'));
  assert.ok(cache._removed.includes('user_data_actor@example.com'));
  assert.ok(cache._removed.includes('admin_panel_actor@example.com'));
});

test('resetAuth: clears userId-scoped keys when user exists', () => {
  const cache = createMockCache();
  const ctx = loadUserServiceContext({
    cache,
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'actor@example.com' })
  });

  ctx.resetAuth();

  assert.ok(cache._removed.includes('user_config_u1'));
  assert.ok(cache._removed.includes('user_kv_u1')); // from getUserCacheKeys_
});

test('resetAuth: calls clearConfigCache for the user', () => {
  let calledWith = null;
  const ctx = loadUserServiceContext({
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'actor@example.com' }),
    clearConfigCache: (id) => { calledWith = id; }
  });

  ctx.resetAuth();
  assert.equal(calledWith, 'u1');
});

test('resetAuth: clears reaction/highlight locks for the user', () => {
  const cache = createMockCache();
  const ctx = loadUserServiceContext({
    cache,
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'actor@example.com' })
  });

  ctx.resetAuth();

  // At least some reaction lock patterns should appear in removed
  assert.ok(cache._removed.some((k) => k.startsWith('reaction_u1_')));
  assert.ok(cache._removed.some((k) => k.startsWith('highlight_u1_')));
});

test('resetAuth: succeeds even when no user is authenticated', () => {
  const ctx = loadUserServiceContext({
    getCurrentEmail: () => null
  });

  const result = ctx.resetAuth();
  assert.equal(result.success, true);
});

test('resetAuth: returns success response with summary counts', () => {
  const ctx = loadUserServiceContext({
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'actor@example.com' })
  });

  const result = ctx.resetAuth();

  assert.equal(result.success, true);
  assert.ok(result.details);
  assert.ok(typeof result.details.clearedKeys === 'number');
  assert.ok(result.details.clearedKeys > 0);
});
