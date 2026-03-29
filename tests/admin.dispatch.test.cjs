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
    getBatchedAdminAuth: () => ({ success: true, email: 'admin@example.com', isAdmin: true }),
    // SystemController functions (called by dispatch)
    testSystemDiagnosis: () => ({ success: true, message: 'diagnosis ok' }),
    performAutoRepair: () => ({ success: true, message: 'repair ok' }),
    forceUrlSystemReset: () => ({ success: true, message: 'reset ok' }),
    getPerformanceMetrics: (cat) => ({ success: true, category: cat }),
    diagnosePerformance: () => ({ success: true, message: 'perf ok' }),
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (k) => k === 'TEST_KEY' ? 'test_value' : k === 'APP_DISABLED' ? null : null,
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
  const result = ctx.dispatchAdminOperation('getProperty', { key: 'TEST_KEY' });
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
