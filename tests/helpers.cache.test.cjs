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
