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
        getProperty: (k) => (overrides.props && k in overrides.props) ? overrides.props[k] : null,
        setProperty: () => {}
      })
    },
    SYSTEM_LIMITS: {
      PREVIEW_LENGTH: 50,
      DEFAULT_PAGE_SIZE: 20,
      MAX_PAGE_SIZE: 100,
      MAX_LOCK_ROWS: 100,
      RADIX_DECIMAL: 10
    },
    DEFAULT_DISPLAY_SETTINGS: {
      showNames: false,
      showReactions: true,
      theme: 'default',
      pageSize: 20
    },
    getCachedProperty: (k) => (overrides.props && k in overrides.props) ? overrides.props[k] : null,
    getWebAppUrl: () => 'https://example.com',
    validateConfig: () => ({ isValid: true, errors: [], sanitized: {} }),
    validateUrl: () => ({ isValid: true }),
    openSpreadsheet: () => ({ spreadsheet: {} }),
    findUserByEmail: () => null,
    findUserById: () => null,
    findUserBySpreadsheetId: () => null,
    updateUser: () => ({ success: true }),
    getCurrentEmail: () => 'admin@test.com',
    isAdministrator: () => true,
    validateEmail: (e) => ({ isValid: typeof e === 'string' && /.+@.+/.test(e) }),
    validateSpreadsheetId: () => true,
    getSheetInfo: () => ({ lastRow: 1, lastCol: 0, headers: [] }),
    UserService: {},
    SLEEP_MS: { SHORT: 100 },
    TIMEOUT_MS: { DEFAULT: 5000 },
    CACHE_DURATION: { EXTRA_LONG: 3600, LONG: 300 },
    URL,
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

// --- sanitizeDisplaySettings ---

test('sanitizeDisplaySettings: coerces booleans and enforces page size bounds', () => {
  const { context } = loadConfigContext();
  const result = context.sanitizeDisplaySettings({
    showNames: 'truthy-string',
    showReactions: 0,
    theme: 'dark',
    pageSize: 200
  });
  assert.equal(result.showNames, true);
  assert.equal(result.showReactions, false);
  assert.equal(result.theme, 'dark');
  assert.equal(result.pageSize, 100); // MAX_PAGE_SIZE cap
});

test('sanitizeDisplaySettings: clamps pageSize below minimum to 1', () => {
  const { context } = loadConfigContext();
  const result = context.sanitizeDisplaySettings({ pageSize: 0 });
  // 0 → falsy, so defaults to DEFAULT_PAGE_SIZE (20)
  assert.equal(result.pageSize, 20);
});

test('sanitizeDisplaySettings: clamps negative pageSize to 1', () => {
  const { context } = loadConfigContext();
  const result = context.sanitizeDisplaySettings({ pageSize: -5 });
  // Math.max(-5, 1) = 1
  assert.equal(result.pageSize, 1);
});

test('sanitizeDisplaySettings: truncates long theme string', () => {
  const { context } = loadConfigContext();
  const result = context.sanitizeDisplaySettings({ theme: 'x'.repeat(200) });
  assert.equal(result.theme.length, 50); // PREVIEW_LENGTH cap
});

test('sanitizeDisplaySettings: uses default theme when absent', () => {
  const { context } = loadConfigContext();
  const result = context.sanitizeDisplaySettings({});
  assert.equal(result.theme, 'default');
});

// --- sanitizeMapping ---

test('sanitizeMapping: keeps valid numeric indices only', () => {
  const { context } = loadConfigContext();
  const result = context.sanitizeMapping({
    answer: 2,
    reason: 3,
    class: 1,
    name: 0,
    email: 4,
    timestamp: 5
  });
  assert.deepEqual({ ...result }, { answer: 2, reason: 3, class: 1, name: 0, email: 4, timestamp: 5 });
});

test('sanitizeMapping: drops non-numeric fields', () => {
  const { context } = loadConfigContext();
  const result = context.sanitizeMapping({
    answer: '2',       // string — rejected
    reason: null,       // null — rejected
    class: true,        // boolean — rejected
    name: { x: 1 }      // object — rejected
  });
  assert.equal(Object.keys(result).length, 0);
});

test('sanitizeMapping: drops fields outside 0..99 range', () => {
  const { context } = loadConfigContext();
  const result = context.sanitizeMapping({
    answer: -1,   // rejected
    reason: 100,  // rejected (exclusive upper bound)
    class: 200,   // rejected
    name: 0       // accepted
  });
  assert.deepEqual({ ...result }, { name: 0 });
});

test('sanitizeMapping: ignores unknown fields', () => {
  const { context } = loadConfigContext();
  const result = context.sanitizeMapping({
    answer: 1,
    __proto__: { malicious: 'yes' },  // ignored, not in validFields
    adminNote: 5                       // ignored
  });
  assert.deepEqual({ ...result }, { answer: 1 });
});

// --- validateConfigUserId ---

test('validateConfigUserId: accepts non-empty strings up to 100 chars', () => {
  const { context } = loadConfigContext();
  assert.equal(context.validateConfigUserId('u1'), true);
  assert.equal(context.validateConfigUserId('x'.repeat(100)), true);
});

test('validateConfigUserId: rejects empty or too-long strings and non-strings', () => {
  const { context } = loadConfigContext();
  assert.equal(context.validateConfigUserId(''), false);
  assert.equal(context.validateConfigUserId('x'.repeat(101)), false);
  assert.equal(context.validateConfigUserId(null), false);
  assert.equal(context.validateConfigUserId(undefined), false);
  assert.equal(context.validateConfigUserId(42), false);
});

// --- parseAndRepairConfig ---

test('parseAndRepairConfig: parses valid JSON and ensures required fields', () => {
  const { context } = loadConfigContext();
  const config = context.parseAndRepairConfig(JSON.stringify({
    spreadsheetId: 'ss-1',
    sheetName: 'Sheet1',
    isPublished: true
  }), 'user-1');

  assert.equal(config.userId, 'user-1');
  assert.equal(config.spreadsheetId, 'ss-1');
  assert.equal(config.sheetName, 'Sheet1');
  assert.equal(config.isPublished, true);
  assert.ok(config.displaySettings);
});

test('parseAndRepairConfig: malformed JSON falls back to default config', () => {
  const { context } = loadConfigContext();
  const config = context.parseAndRepairConfig('not-valid-json', 'user-1');
  assert.equal(config.userId, 'user-1');
  assert.equal(config.setupStatus, 'pending');
  assert.equal(config.isPublished, false);
});

test('parseAndRepairConfig: empty string yields default config', () => {
  const { context } = loadConfigContext();
  const config = context.parseAndRepairConfig('', 'user-1');
  assert.equal(config.userId, 'user-1');
});

test('parseAndRepairConfig: broken displaySettings are repaired', () => {
  const { context } = loadConfigContext();
  const config = context.parseAndRepairConfig(JSON.stringify({
    displaySettings: 'not-an-object'
  }), 'user-1');
  assert.ok(config.displaySettings);
  assert.equal(typeof config.displaySettings, 'object');
});

// --- getDefaultConfig ---

test('getDefaultConfig: returns shape with required fields', () => {
  const { context } = loadConfigContext();
  const config = context.getDefaultConfig('user-1');
  assert.equal(config.userId, 'user-1');
  assert.equal(config.setupStatus, 'pending');
  assert.equal(config.isPublished, false);
  assert.equal(config.hasSeenWelcome, false);
  assert.equal(config.completionScore, 0);
  assert.ok(config.displaySettings);
});

// --- hasCoreSystemProps ---

test('hasCoreSystemProps: true when all three properties present', () => {
  const { context } = loadConfigContext({
    props: {
      ADMIN_EMAIL: 'admin@example.com',
      DATABASE_SPREADSHEET_ID: 'db-id',
      SERVICE_ACCOUNT_CREDS: '{"client_email":"sa@..."}'
    }
  });
  assert.equal(context.hasCoreSystemProps(), true);
});

test('hasCoreSystemProps: false when ADMIN_EMAIL missing', () => {
  const { context } = loadConfigContext({
    props: { DATABASE_SPREADSHEET_ID: 'db', SERVICE_ACCOUNT_CREDS: '{}' }
  });
  assert.equal(context.hasCoreSystemProps(), false);
});

test('hasCoreSystemProps: false when DATABASE_SPREADSHEET_ID missing', () => {
  const { context } = loadConfigContext({
    props: { ADMIN_EMAIL: 'a@x.com', SERVICE_ACCOUNT_CREDS: '{}' }
  });
  assert.equal(context.hasCoreSystemProps(), false);
});

test('hasCoreSystemProps: false when SERVICE_ACCOUNT_CREDS missing', () => {
  const { context } = loadConfigContext({
    props: { ADMIN_EMAIL: 'a@x.com', DATABASE_SPREADSHEET_ID: 'db' }
  });
  assert.equal(context.hasCoreSystemProps(), false);
});

test('hasCoreSystemProps: false when nothing configured', () => {
  const { context } = loadConfigContext();
  assert.equal(context.hasCoreSystemProps(), false);
});

// --- calculateCompletionScore ---

test('calculateCompletionScore: minimal config scores low', () => {
  const { context } = loadConfigContext();
  const score = context.calculateCompletionScore({});
  assert.ok(score >= 0 && score <= 100);
  assert.ok(score < 50, `Expected low score for empty config, got ${score}`);
});

test('calculateCompletionScore: fully configured scores high', () => {
  const { context } = loadConfigContext();
  const score = context.calculateCompletionScore({
    spreadsheetId: 'ss-1',
    sheetName: 'Sheet1',
    formUrl: 'https://docs.google.com/forms/x',
    columnMapping: { answer: 2, reason: 3, class: 1, name: 0 },
    displaySettings: { showNames: true, showReactions: true },
    isPublished: true
  });
  assert.ok(score >= 50, `Expected high score for full config, got ${score}`);
});
