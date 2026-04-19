const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('path');
const vm = require('vm');

function loadHelpersContext(overrides = {}) {
  const propsStore = { ADMIN_EMAIL: 'admin@test.com', DB_ID: 'spreadsheet-123' };
  const propsCalls = [];
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (k) => { propsCalls.push(['get', k]); return propsStore[k] || null; },
        setProperty: (k, v) => { propsCalls.push(['set', k, v]); propsStore[k] = v; }
      })
    },
    CacheService: {
      getScriptCache: () => ({
        get: () => null,
        put: () => {},
        remove: () => {},
        removeAll: () => {}
      })
    },
    PROPERTY_CACHE_TTL: 30000,
    Date: { now: () => overrides._mockTime || Date.now() },
    ...overrides
  };
  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/helpers.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'helpers.js' });
  return { context, propsCalls, propsStore };
}

// --- getCachedProperty ---

test('getCachedProperty: fetches from PropertiesService on first call', () => {
  const { context, propsCalls } = loadHelpersContext();
  const value = context.getCachedProperty('ADMIN_EMAIL');
  assert.equal(value, 'admin@test.com');
  assert.deepEqual(propsCalls, [['get', 'ADMIN_EMAIL']]);
});

test('getCachedProperty: returns cached value on second call within TTL', () => {
  const { context, propsCalls } = loadHelpersContext();
  context.getCachedProperty('ADMIN_EMAIL');
  context.getCachedProperty('ADMIN_EMAIL');
  // PropertiesService should be called only once
  assert.equal(propsCalls.filter(c => c[0] === 'get').length, 1);
});

test('getCachedProperty: refetches after TTL expires', () => {
  let mockTime = 1000000;
  const { context, propsCalls } = loadHelpersContext({
    _mockTime: mockTime,
    Date: { now: () => mockTime }
  });
  context.getCachedProperty('ADMIN_EMAIL');
  assert.equal(propsCalls.filter(c => c[0] === 'get').length, 1);

  // Advance past 30s TTL
  mockTime += 31000;
  context.Date = { now: () => mockTime };
  context.getCachedProperty('ADMIN_EMAIL');
  assert.equal(propsCalls.filter(c => c[0] === 'get').length, 2);
});

// --- setCachedProperty ---

test('setCachedProperty: writes to PropertiesService', () => {
  const { context, propsCalls } = loadHelpersContext();
  context.setCachedProperty('NEW_KEY', 'new_value');
  assert.deepEqual(propsCalls[0], ['set', 'NEW_KEY', 'new_value']);
});

test('setCachedProperty: invalidates memory cache', () => {
  const { context, propsCalls } = loadHelpersContext();
  // Cache the value
  context.getCachedProperty('ADMIN_EMAIL');
  assert.equal(propsCalls.filter(c => c[0] === 'get').length, 1);

  // Write new value — should invalidate cache
  context.setCachedProperty('ADMIN_EMAIL', 'new@test.com');

  // Next read should fetch from PropertiesService again
  context.getCachedProperty('ADMIN_EMAIL');
  assert.equal(propsCalls.filter(c => c[0] === 'get').length, 2);
});

// --- clearPropertyCache ---

test('clearPropertyCache: clears specific key', () => {
  const { context, propsCalls } = loadHelpersContext();
  context.getCachedProperty('ADMIN_EMAIL');
  context.clearPropertyCache('ADMIN_EMAIL');
  context.getCachedProperty('ADMIN_EMAIL');
  // Should have fetched twice (cache was cleared)
  assert.equal(propsCalls.filter(c => c[0] === 'get').length, 2);
});

test('clearPropertyCache: clears all when no key', () => {
  const { context, propsCalls } = loadHelpersContext();
  context.getCachedProperty('ADMIN_EMAIL');
  context.getCachedProperty('DB_ID');
  context.clearPropertyCache();
  context.getCachedProperty('ADMIN_EMAIL');
  context.getCachedProperty('DB_ID');
  // Each key fetched twice
  assert.equal(propsCalls.filter(c => c[0] === 'get').length, 4);
});

// --- setCachedProperty ---

test('setCachedProperty: writes through and clears memory cache for that key', () => {
  const { context, propsCalls } = loadHelpersContext();
  context.getCachedProperty('ADMIN_EMAIL');
  context.setCachedProperty('ADMIN_EMAIL', 'new@example.com');
  // After set, the cached entry is invalidated — next read hits PropertiesService
  context.getCachedProperty('ADMIN_EMAIL');
  const gets = propsCalls.filter((c) => c[0] === 'get').length;
  const sets = propsCalls.filter((c) => c[0] === 'set').length;
  assert.equal(sets, 1);
  assert.equal(gets, 2, 'Cache should be invalidated after setCachedProperty');
});

// --- timingSafeEqual ---

test('timingSafeEqual: true for identical strings', () => {
  const { context } = loadHelpersContext();
  assert.equal(context.timingSafeEqual('hello', 'hello'), true);
});

test('timingSafeEqual: false for different strings of same length', () => {
  const { context } = loadHelpersContext();
  assert.equal(context.timingSafeEqual('hello', 'world'), false);
});

test('timingSafeEqual: false when lengths differ', () => {
  const { context } = loadHelpersContext();
  assert.equal(context.timingSafeEqual('abc', 'abcd'), false);
  assert.equal(context.timingSafeEqual('', 'x'), false);
});

test('timingSafeEqual: false for non-string inputs', () => {
  const { context } = loadHelpersContext();
  assert.equal(context.timingSafeEqual(null, 'x'), false);
  assert.equal(context.timingSafeEqual('x', null), false);
  assert.equal(context.timingSafeEqual(42, 42), false);
  assert.equal(context.timingSafeEqual(undefined, undefined), false);
});

test('timingSafeEqual: true for empty strings', () => {
  const { context } = loadHelpersContext();
  assert.equal(context.timingSafeEqual('', ''), true);
});

// --- createErrorResponse ---

test('createErrorResponse: minimal shape', () => {
  const { context } = loadHelpersContext();
  const r = context.createErrorResponse('something bad');
  assert.equal(r.success, false);
  assert.equal(r.message, 'something bad');
  assert.equal(r.error, 'something bad');
  assert.equal(r.data, undefined);
});

test('createErrorResponse: attaches data when provided', () => {
  const { context } = loadHelpersContext();
  const r = context.createErrorResponse('bad', { attemptId: 'x' });
  assert.deepEqual({ ...r.data }, { attemptId: 'x' });
});

test('createErrorResponse: spreads extra fields', () => {
  const { context } = loadHelpersContext();
  const r = context.createErrorResponse('bad', null, { error: 'CODE_X', retry: true });
  assert.equal(r.error, 'CODE_X');
  assert.equal(r.retry, true);
});

// --- createSuccessResponse ---

test('createSuccessResponse: minimal shape', () => {
  const { context } = loadHelpersContext();
  const r = context.createSuccessResponse('ok');
  assert.equal(r.success, true);
  assert.equal(r.message, 'ok');
  assert.equal(r.data, undefined);
});

test('createSuccessResponse: attaches data when provided', () => {
  const { context } = loadHelpersContext();
  const r = context.createSuccessResponse('ok', { users: [1, 2] });
  assert.equal(r.data.users.length, 2);
});

test('createSuccessResponse: spreads extra fields', () => {
  const { context } = loadHelpersContext();
  const r = context.createSuccessResponse('ok', null, { etag: 'abc', count: 5 });
  assert.equal(r.etag, 'abc');
  assert.equal(r.count, 5);
});

// --- createAuthError / createUserNotFoundError / createAdminRequiredError ---

test('createAuthError: stable shape', () => {
  const { context } = loadHelpersContext();
  const r = context.createAuthError();
  assert.equal(r.success, false);
  assert.match(r.message, /認証/);
});

test('createUserNotFoundError: stable shape', () => {
  const { context } = loadHelpersContext();
  const r = context.createUserNotFoundError();
  assert.equal(r.success, false);
  assert.match(r.message, /ユーザー/);
});

test('createAdminRequiredError: stable shape', () => {
  const { context } = loadHelpersContext();
  const r = context.createAdminRequiredError();
  assert.equal(r.success, false);
  assert.match(r.message, /管理者/);
});

// --- createExceptionResponse ---

test('createExceptionResponse: uses error.message when present', () => {
  const { context } = loadHelpersContext();
  const r = context.createExceptionResponse(new Error('boom'));
  assert.equal(r.success, false);
  assert.equal(r.message, 'boom');
});

test('createExceptionResponse: falls back to "Unknown error"', () => {
  const { context } = loadHelpersContext();
  const r = context.createExceptionResponse({});
  assert.equal(r.message, 'Unknown error');
});

// --- createDataServiceErrorResponse ---

test('createDataServiceErrorResponse: includes sheetName and empty arrays', () => {
  const { context } = loadHelpersContext();
  const r = context.createDataServiceErrorResponse('fetch failed', 'Sheet1');
  assert.equal(r.success, false);
  assert.equal(r.sheetName, 'Sheet1');
  assert.equal(Array.isArray(r.headers), true);
  assert.equal(r.headers.length, 0);
});

// --- simpleHash ---

test('simpleHash: stable for identical objects regardless of key order', () => {
  const { context } = loadHelpersContext();
  const a = context.simpleHash({ a: 1, b: 2, c: 3 });
  const b = context.simpleHash({ c: 3, a: 1, b: 2 });
  assert.equal(a, b);
});

test('simpleHash: differs for different values', () => {
  const { context } = loadHelpersContext();
  const a = context.simpleHash({ x: 1 });
  const b = context.simpleHash({ x: 2 });
  assert.notEqual(a, b);
});

test('simpleHash: empty string for null/non-object', () => {
  const { context } = loadHelpersContext();
  assert.equal(context.simpleHash(null), '');
  assert.equal(context.simpleHash('string'), '');
  assert.equal(context.simpleHash(42), '');
});

// --- saveToCacheWithSizeCheck ---

test('saveToCacheWithSizeCheck: saves within size limit', () => {
  let putKey = null;
  let putValue = null;
  const { context } = loadHelpersContext({
    CacheService: {
      getScriptCache: () => ({
        put: (k, v) => { putKey = k; putValue = v; },
        get: () => null,
        remove: () => {}
      })
    }
  });
  const saved = context.saveToCacheWithSizeCheck('k1', { a: 1 }, 60);
  assert.equal(saved, true);
  assert.equal(putKey, 'k1');
  assert.ok(putValue && putValue.includes('"a":1'));
});

test('saveToCacheWithSizeCheck: refuses data over maxSize', () => {
  let putCalled = false;
  const { context } = loadHelpersContext({
    CacheService: {
      getScriptCache: () => ({
        put: () => { putCalled = true; },
        get: () => null,
        remove: () => {}
      })
    }
  });
  const big = { x: 'a'.repeat(200000) };
  const saved = context.saveToCacheWithSizeCheck('k1', big, 60);
  assert.equal(saved, false);
  assert.equal(putCalled, false);
});

test('saveToCacheWithSizeCheck: returns false when cache.put throws', () => {
  const { context } = loadHelpersContext({
    CacheService: {
      getScriptCache: () => ({
        put: () => { throw new Error('cache failure'); },
        get: () => null,
        remove: () => {}
      })
    }
  });
  const saved = context.saveToCacheWithSizeCheck('k1', { a: 1 }, 60);
  assert.equal(saved, false);
});

// --- normalizeHeader ---

test('normalizeHeader: lowercases and trims', () => {
  const { context } = loadHelpersContext();
  assert.equal(context.normalizeHeader('  NAME  '), 'name');
});

test('normalizeHeader: returns empty string for null/undefined', () => {
  const { context } = loadHelpersContext();
  assert.equal(context.normalizeHeader(null), '');
  assert.equal(context.normalizeHeader(undefined), '');
});
