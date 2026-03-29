const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('path');
const vm = require('vm');

function loadConfigContext(overrides = {}) {
  const removedKeys = [];
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    CacheService: {
      getScriptCache: () => ({
        removeAll: (keys) => { removedKeys.push(...keys); },
        remove: () => {},
        get: () => null,
        put: () => {}
      })
    },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: () => null,
        setProperty: () => {}
      })
    },
    getCachedProperty: () => null,
    getWebAppUrl: () => 'https://example.com',
    validateConfig: () => ({ isValid: true, errors: [] }),
    openSpreadsheet: () => ({ spreadsheet: {} }),
    findUserByEmail: () => null,
    updateUser: () => ({ success: true }),
    getCurrentEmail: () => 'admin@test.com',
    isAdministrator: () => true,
    ...overrides
  };
  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/ConfigService.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'ConfigService.js' });
  return { context, removedKeys };
}

// --- getUserCacheKeys_ ---

test('getUserCacheKeys_: returns 5 keys for userId', () => {
  const { context } = loadConfigContext();
  const keys = context.getUserCacheKeys_('user-123');
  assert.equal(keys.length, 5);
  assert.ok(keys.includes('config_user-123'));
  assert.ok(keys.includes('user_user-123'));
  assert.ok(keys.includes('board_data_user-123'));
  assert.ok(keys.includes('admin_panel_user-123'));
  assert.ok(keys.includes('question_text_user-123'));
});

// --- clearConfigCache ---

test('clearConfigCache: removes 5 keys for user', () => {
  const { context, removedKeys } = loadConfigContext();
  context.clearConfigCache('user-1');
  assert.equal(removedKeys.length, 5);
  assert.ok(removedKeys.includes('config_user-1'));
});

// --- clearAllConfigCache ---

test('clearAllConfigCache: removes keys for multiple users', () => {
  const { context, removedKeys } = loadConfigContext();
  context.clearAllConfigCache(['user-1', 'user-2']);
  assert.equal(removedKeys.length, 10);
  assert.ok(removedKeys.includes('config_user-1'));
  assert.ok(removedKeys.includes('config_user-2'));
});

test('clearAllConfigCache: filters out non-string and empty ids', () => {
  const { context, removedKeys } = loadConfigContext();
  context.clearAllConfigCache(['user-1', '', null, 123, 'user-2']);
  assert.equal(removedKeys.length, 10); // only user-1 and user-2
});

test('clearAllConfigCache: does nothing for empty array', () => {
  const { context, removedKeys } = loadConfigContext();
  context.clearAllConfigCache([]);
  assert.equal(removedKeys.length, 0);
});

test('clearAllConfigCache: does nothing for all-invalid ids', () => {
  const { context, removedKeys } = loadConfigContext();
  context.clearAllConfigCache([null, '', 0]);
  assert.equal(removedKeys.length, 0);
});
