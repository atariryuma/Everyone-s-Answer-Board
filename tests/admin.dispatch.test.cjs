const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('path');
const vm = require('vm');

function loadAdminContext(overrides = {}) {
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    createErrorResponse: (msg, data, extra) => ({ success: false, message: msg, ...extra }),
    createSuccessResponse: (msg, data) => ({ success: true, message: msg, ...(data && { data }) }),
    createAdminRequiredError: () => ({ success: false, message: 'admin required' }),
    createAuthError: () => ({ success: false, message: 'auth required' }),
    createUserNotFoundError: () => ({ success: false, message: 'user not found' }),
    createExceptionResponse: (e) => ({ success: false, message: e.message }),
    getCurrentEmail: () => 'admin@example.com',
    isAdministrator: () => true,
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'admin@example.com' }),
    findUserById: () => ({ userId: 'u1', userEmail: 'admin@example.com', isActive: true }),
    getAllUsers: () => [{ userId: 'u1', userEmail: 'admin@example.com', isActive: true }],
    updateUser: () => ({ success: true }),
    getUserConfig: () => ({ success: true, config: {} }),
    saveUserConfig: () => ({ success: true }),
    requireAdmin: () => ({ email: 'admin@example.com', isAdmin: true }),
    getConfigOrDefault: () => ({}),
    getBatchedAdminAuth: () => ({ success: true, email: 'admin@example.com', isAdmin: true }),
    // SystemController functions (called by dispatch)
    testSystemDiagnosis: () => ({ success: true, message: 'diagnosis ok' }),
    performAutoRepair: () => ({ success: true, message: 'repair ok' }),
    forceUrlSystemReset: () => ({ success: true, message: 'reset ok' }),
    getPerformanceMetrics: (cat) => ({ success: true, category: cat }),
    diagnosePerformance: () => ({ success: true, message: 'perf ok' }),
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (k) => k === 'APP_NAME' ? 'test_value' : k === 'APP_DISABLED' ? null : null,
        setProperty: () => {},
        deleteProperty: () => {},
        getProperties: () => ({ 'APP_NAME': 'test', 'ADMIN_API_KEY': 'secret123' })
      })
    },
    getCachedProperty: () => null,
    ...overrides
  };
  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/AdminApis.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'AdminApis.js' });
  return context;
}

// --- Routing ---

test('dispatch: missing operation returns error', () => {
  const ctx = loadAdminContext();
  const result = ctx.dispatchAdminOperation(null, {});
  assert.equal(result.success, false);
  assert.equal(result.error, 'MISSING_OPERATION');
});

test('dispatch: unknown operation returns error', () => {
  const ctx = loadAdminContext();
  const result = ctx.dispatchAdminOperation('badOp', {});
  assert.equal(result.success, false);
  assert.equal(result.error, 'UNKNOWN_OPERATION');
});

test('dispatch: trims whitespace from operation', () => {
  const ctx = loadAdminContext();
  const result = ctx.dispatchAdminOperation('  getAppStatus  ', {});
  assert.equal(result.success, true);
});

// --- User Management ---

test('dispatch: getUsers returns success', () => {
  const ctx = loadAdminContext();
  const result = ctx.dispatchAdminOperation('getUsers', {});
  assert.equal(result.success, true);
});

test('dispatch: toggleUserActive requires targetUserId', () => {
  const ctx = loadAdminContext();
  const result = ctx.dispatchAdminOperation('toggleUserActive', {});
  assert.equal(result.success, false);
  assert.match(result.message, /targetUserId/);
});

test('dispatch: toggleUserBoard requires targetUserId', () => {
  const ctx = loadAdminContext();
  const result = ctx.dispatchAdminOperation('toggleUserBoard', {});
  assert.equal(result.success, false);
  assert.match(result.message, /targetUserId/);
});

// --- App Control ---

test('dispatch: getAppStatus returns success', () => {
  const ctx = loadAdminContext();
  const result = ctx.dispatchAdminOperation('getAppStatus', {});
  assert.equal(result.success, true);
});

test('dispatch: enableApp returns success', () => {
  const ctx = loadAdminContext();
  const result = ctx.dispatchAdminOperation('enableApp', {});
  assert.equal(result.success, true);
});

// --- System ---

test('dispatch: systemDiagnosis returns success', () => {
  const ctx = loadAdminContext();
  const result = ctx.dispatchAdminOperation('systemDiagnosis', {});
  assert.equal(result.success, true);
});

test('dispatch: perfMetrics defaults category to all', () => {
  const ctx = loadAdminContext();
  const result = ctx.dispatchAdminOperation('perfMetrics', {});
  assert.equal(result.success, true);
  assert.equal(result.category, 'all');
});

test('dispatch: perfMetrics passes custom category', () => {
  const ctx = loadAdminContext();
  const result = ctx.dispatchAdminOperation('perfMetrics', { category: 'api' });
  assert.equal(result.success, true);
  assert.equal(result.category, 'api');
});

// --- Property Management ---

test('dispatch: getProperty returns value', () => {
  const ctx = loadAdminContext();
  const result = ctx.dispatchAdminOperation('getProperty', { key: 'APP_NAME' });
  assert.equal(result.success, true);
  assert.equal(result.data.value, 'test_value');
});

test('dispatch: getProperty requires key', () => {
  const ctx = loadAdminContext();
  const result = ctx.dispatchAdminOperation('getProperty', {});
  assert.equal(result.success, false);
});

test('dispatch: setProperty requires key and string value', () => {
  const ctx = loadAdminContext();
  assert.equal(ctx.dispatchAdminOperation('setProperty', {}).success, false);
  assert.equal(ctx.dispatchAdminOperation('setProperty', { key: 'k' }).success, false);
  assert.equal(ctx.dispatchAdminOperation('setProperty', { key: 'k', value: 123 }).success, false);
});

test('dispatch: listProperties masks sensitive keys', () => {
  const ctx = loadAdminContext();
  const result = ctx.dispatchAdminOperation('listProperties', {});
  assert.equal(result.success, true);
  // Non-sensitive key: shown as-is
  assert.equal(result.data.properties['APP_NAME'], 'test');
  // Sensitive key (contains KEY): must be fully masked, not contain actual value
  const masked = result.data.properties['ADMIN_API_KEY'];
  assert.ok(masked.startsWith('***'), 'Should start with ***');
  assert.ok(masked.includes('chars'), 'Should indicate character count');
  assert.ok(!masked.includes('secret123'), 'Must not contain actual value');
});

// --- Protected Property Denylist ---

test('dispatch: getProperty rejects ADMIN_API_KEY', () => {
  const ctx = loadAdminContext();
  const result = ctx.dispatchAdminOperation('getProperty', { key: 'ADMIN_API_KEY' });
  assert.equal(result.success, false);
  assert.equal(result.error, 'PROTECTED_PROPERTY');
});

test('dispatch: getProperty rejects SERVICE_ACCOUNT_CREDS', () => {
  const ctx = loadAdminContext();
  const result = ctx.dispatchAdminOperation('getProperty', { key: 'SERVICE_ACCOUNT_CREDS' });
  assert.equal(result.success, false);
  assert.equal(result.error, 'PROTECTED_PROPERTY');
});

test('dispatch: getProperty rejects ADMIN_EMAIL (account takeover vector)', () => {
  const ctx = loadAdminContext();
  const result = ctx.dispatchAdminOperation('getProperty', { key: 'ADMIN_EMAIL' });
  assert.equal(result.success, false);
  assert.equal(result.error, 'PROTECTED_PROPERTY');
});

test('dispatch: getProperty rejects DATABASE_SPREADSHEET_ID (DB redirect vector)', () => {
  const ctx = loadAdminContext();
  const result = ctx.dispatchAdminOperation('getProperty', { key: 'DATABASE_SPREADSHEET_ID' });
  assert.equal(result.success, false);
  assert.equal(result.error, 'PROTECTED_PROPERTY');
});

test('dispatch: getProperty rejects any key containing KEY/CREDS/SECRET/TOKEN/PASSWORD', () => {
  const ctx = loadAdminContext();
  for (const key of ['MY_KEY', 'foo_CREDS', 'anything_secret', 'OAUTH_TOKEN', 'DB_PASSWORD']) {
    const result = ctx.dispatchAdminOperation('getProperty', { key });
    assert.equal(result.success, false, `Expected ${key} to be protected`);
    assert.equal(result.error, 'PROTECTED_PROPERTY');
  }
});

test('dispatch: getProperty is case-insensitive for protection', () => {
  const ctx = loadAdminContext();
  const result = ctx.dispatchAdminOperation('getProperty', { key: 'admin_api_key' });
  assert.equal(result.success, false);
  assert.equal(result.error, 'PROTECTED_PROPERTY');
});

for (const protectedKey of ['ADMIN_API_KEY', 'SERVICE_ACCOUNT_CREDS', 'ADMIN_EMAIL', 'DATABASE_SPREADSHEET_ID']) {
  test(`dispatch: setProperty rejects ${protectedKey} overwrite`, () => {
    let setCalls = 0;
    const ctx = loadAdminContext({
      PropertiesService: {
        getScriptProperties: () => ({
          getProperty: () => null,
          setProperty: () => { setCalls += 1; },
          deleteProperty: () => {},
          getProperties: () => ({})
        })
      }
    });
    const result = ctx.dispatchAdminOperation('setProperty', { key: protectedKey, value: 'attacker' });
    assert.equal(result.success, false);
    assert.equal(result.error, 'PROTECTED_PROPERTY');
    assert.equal(setCalls, 0, 'setProperty must not be called for protected keys');
  });
}

test('dispatch: setProperty allows non-protected keys', () => {
  let captured = null;
  const ctx = loadAdminContext({
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: () => null,
        setProperty: (k, v) => { captured = { k, v }; },
        deleteProperty: () => {},
        getProperties: () => ({})
      })
    }
  });
  const result = ctx.dispatchAdminOperation('setProperty', { key: 'APP_DISABLED', value: 'false' });
  assert.equal(result.success, true);
  assert.deepEqual(captured, { k: 'APP_DISABLED', v: 'false' });
});

test('dispatch: getProperty allows non-protected keys', () => {
  const ctx = loadAdminContext({
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (k) => k === 'APP_DISABLED' ? 'true' : null,
        setProperty: () => {},
        deleteProperty: () => {},
        getProperties: () => ({})
      })
    }
  });
  const result = ctx.dispatchAdminOperation('getProperty', { key: 'APP_DISABLED' });
  assert.equal(result.success, true);
  assert.equal(result.data.value, 'true');
});
