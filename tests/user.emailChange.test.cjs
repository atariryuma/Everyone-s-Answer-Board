const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('path');
const vm = require('vm');

// --- getGoogleIdFromToken tests ---

function loadUserServiceContext(overrides = {}) {
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    validateUrl: () => ({ isValid: true }),
    validateEmail: (e) => ({ isValid: !!e && e.includes('@') }),
    getCurrentEmail: () => 'user@example.com',
    findUserByEmail: () => null,
    findUserById: () => null,
    findUserByGoogleId: () => null,
    openSpreadsheet: () => null,
    updateUser: () => ({ success: true }),
    getUserConfig: () => ({ success: true, config: {} }),
    isAdministrator: () => false,
    CACHE_DURATION: { LONG: 600, USER_INDIVIDUAL: 300 },
    clearConfigCache: () => {},
    SYSTEM_LIMITS: { MAX_USERS: 100 },
    createExceptionResponse: (e) => ({ success: false, message: e.message }),
    CacheService: { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) },
    ScriptApp: { getIdentityToken: () => null },
    Utilities: {
      base64DecodeWebSafe: (str) => Buffer.from(str, 'base64url'),
      newBlob: (data) => ({ getDataAsString: () => Buffer.from(data).toString('utf8') })
    },
    ...overrides
  };
  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/UserService.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'UserService.js' });
  return context;
}

function createMockJwt(payload) {
  const header = Buffer.from(JSON.stringify({ alg: 'RS256', typ: 'JWT' })).toString('base64url');
  const body = Buffer.from(JSON.stringify(payload)).toString('base64url');
  return `${header}.${body}.fake-signature`;
}

test('getGoogleIdFromToken: returns null when token is not available', () => {
  const ctx = loadUserServiceContext({
    ScriptApp: { getIdentityToken: () => null }
  });
  const result = ctx.getGoogleIdFromToken();
  assert.equal(result, null);
});

test('getGoogleIdFromToken: extracts sub and email from valid JWT', () => {
  const mockToken = createMockJwt({ sub: '123456789', email: 'user@example.com' });
  const ctx = loadUserServiceContext({
    ScriptApp: { getIdentityToken: () => mockToken }
  });
  const result = ctx.getGoogleIdFromToken();
  assert.equal(result.googleId, 'gid_123456789');
  assert.equal(result.email, 'user@example.com');
});

test('getGoogleIdFromToken: returns null when sub is missing', () => {
  const mockToken = createMockJwt({ email: 'user@example.com' });
  const ctx = loadUserServiceContext({
    ScriptApp: { getIdentityToken: () => mockToken }
  });
  const result = ctx.getGoogleIdFromToken();
  assert.equal(result, null);
});

test('getGoogleIdFromToken: handles malformed JWT gracefully', () => {
  const ctx = loadUserServiceContext({
    ScriptApp: { getIdentityToken: () => 'not-a-jwt' }
  });
  const result = ctx.getGoogleIdFromToken();
  assert.equal(result, null);
});

// --- processLoginAction email change detection tests ---

function loadUserApisContext(overrides = {}) {
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    getCurrentEmail: () => 'newmail@example.com',
    findUserByEmail: () => null,
    findUserByGoogleId: () => null,
    isAdministrator: () => false,
    getUserConfig: () => ({ success: true, config: {} }),
    createUser: (email, config, ctx) => ({
      userId: 'new-user-id',
      userEmail: email,
      googleId: ctx?.googleId || '',
      isActive: true
    }),
    updateUser: () => ({ success: true }),
    getGoogleIdFromToken: () => ({ googleId: 'google-id-123', email: 'newmail@example.com' }),
    getCachedProperty: () => null,
    setCachedProperty: () => {},
    getWebAppUrl: () => 'https://script.google.com/app',
    createAuthError: () => ({ success: false, message: 'auth required' }),
    createUserNotFoundError: () => ({ success: false, message: 'user not found' }),
    createExceptionResponse: (e) => ({ success: false, message: e.message }),
    ScriptApp: { getIdentityToken: () => null },
    Utilities: {},
    shouldEnforceDomainRestrictions: () => false,
    validateDomainAccess: () => ({ allowed: true }),
    ...overrides
  };
  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/UserApis.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'UserApis.js' });
  return context;
}

test('processLoginAction: detects email change via googleId and updates user', () => {
  const updates = [];
  const oldUser = {
    userId: 'existing-user-id',
    userEmail: 'oldmail@example.com',
    googleId: 'google-id-123',
    isActive: true
  };

  const ctx = loadUserApisContext({
    getCurrentEmail: () => 'newmail@example.com',
    findUserByEmail: () => null, // new email not found
    findUserByGoogleId: (gid, ctx) => {
      assert.equal(ctx.skipAccessCheck, true, 'skipAccessCheck must be true for email migration');
      return gid === 'google-id-123' ? { ...oldUser } : null;
    },
    updateUser: (userId, data) => { updates.push({ userId, ...data }); return { success: true }; },
    getGoogleIdFromToken: () => ({ googleId: 'google-id-123', email: 'newmail@example.com' })
  });

  const result = ctx.processLoginAction();
  assert.equal(result.success, true);
  assert.ok(result.redirectUrl.includes('existing-user-id'));
  // Should have updated userEmail
  const emailUpdate = updates.find(u => u.userEmail === 'newmail@example.com');
  assert.ok(emailUpdate, 'userEmail should be updated to new email');
  assert.equal(emailUpdate.userId, 'existing-user-id');
});

test('processLoginAction: updates ADMIN_EMAIL when admin email changes', () => {
  const updates = [];
  let savedAdminEmail = null;
  const oldUser = {
    userId: 'admin-user-id',
    userEmail: 'old-admin@example.com',
    googleId: 'google-id-admin',
    isActive: true
  };

  const ctx = loadUserApisContext({
    getCurrentEmail: () => 'new-admin@example.com',
    findUserByEmail: () => null,
    findUserByGoogleId: (gid, ctx) => gid === 'google-id-admin' ? { ...oldUser } : null,
    updateUser: (userId, data) => { updates.push({ userId, ...data }); return { success: true }; },
    getGoogleIdFromToken: () => ({ googleId: 'google-id-admin', email: 'new-admin@example.com' }),
    getCachedProperty: (key) => key === 'ADMIN_EMAIL' ? 'old-admin@example.com' : null,
    setCachedProperty: (key, value) => { if (key === 'ADMIN_EMAIL') savedAdminEmail = value; }
  });

  const result = ctx.processLoginAction();
  assert.equal(result.success, true);
  assert.equal(savedAdminEmail, 'new-admin@example.com');
});

test('processLoginAction: does not update ADMIN_EMAIL for non-admin email change', () => {
  let savedAdminEmail = null;
  const oldUser = {
    userId: 'regular-user-id',
    userEmail: 'old-user@example.com',
    googleId: 'google-id-regular',
    isActive: true
  };

  const ctx = loadUserApisContext({
    getCurrentEmail: () => 'new-user@example.com',
    findUserByEmail: () => null,
    findUserByGoogleId: (gid, ctx) => gid === 'google-id-regular' ? { ...oldUser } : null,
    updateUser: () => ({ success: true }),
    getGoogleIdFromToken: () => ({ googleId: 'google-id-regular', email: 'new-user@example.com' }),
    getCachedProperty: (key) => key === 'ADMIN_EMAIL' ? 'admin@example.com' : null,
    setCachedProperty: (key, value) => { if (key === 'ADMIN_EMAIL') savedAdminEmail = value; }
  });

  const result = ctx.processLoginAction();
  assert.equal(result.success, true);
  assert.equal(savedAdminEmail, null, 'ADMIN_EMAIL should not be changed for non-admin users');
});

test('processLoginAction: backfills googleId for existing user without it', () => {
  const updates = [];
  const existingUser = {
    userId: 'user-1',
    userEmail: 'user@example.com',
    googleId: '', // not set yet
    isActive: true
  };

  const ctx = loadUserApisContext({
    getCurrentEmail: () => 'user@example.com',
    findUserByEmail: () => ({ ...existingUser }),
    getGoogleIdFromToken: () => ({ googleId: 'google-id-456', email: 'user@example.com' }),
    updateUser: (userId, data) => { updates.push({ userId, ...data }); return { success: true }; }
  });

  const result = ctx.processLoginAction();
  assert.equal(result.success, true);
  const googleIdUpdate = updates.find(u => u.googleId === 'google-id-456');
  assert.ok(googleIdUpdate, 'googleId should be backfilled');
});

test('processLoginAction: creates new user with googleId when no match found', () => {
  let createdContext = null;

  const ctx = loadUserApisContext({
    getCurrentEmail: () => 'brand-new@example.com',
    findUserByEmail: () => null,
    findUserByGoogleId: () => null,
    getGoogleIdFromToken: () => ({ googleId: 'google-id-789', email: 'brand-new@example.com' }),
    createUser: (email, config, context) => {
      createdContext = context;
      return { userId: 'new-id', userEmail: email, googleId: context?.googleId || '' };
    }
  });

  const result = ctx.processLoginAction();
  assert.equal(result.success, true);
  assert.equal(createdContext.googleId, 'google-id-789');
});

test('processLoginAction: works gracefully when identity token is unavailable', () => {
  const ctx = loadUserApisContext({
    getCurrentEmail: () => 'user@example.com',
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'user@example.com', isActive: true }),
    getGoogleIdFromToken: () => null
  });

  const result = ctx.processLoginAction();
  assert.equal(result.success, true);
});
