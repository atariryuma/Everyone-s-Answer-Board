const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function loadDataApisContext(overrides = {}) {
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    URL,
    createErrorResponse: (message) => ({ success: false, message }),
    createExceptionResponse: (error) => ({ success: false, message: error.message }),
    createSuccessResponse: (message, data) => ({ success: true, message, ...(data && { data }) }),
    getCurrentEmail: () => 'actor@example.com',
    findUserByEmail: () => null,
    findUserById: () => null,
    findUserBySpreadsheetId: () => null,
    getUserConfig: () => ({ success: true, config: {} }),
    getConfigOrDefault: () => ({}),
    saveUserConfig: () => ({ success: true }),
    openSpreadsheet: () => null,
    SpreadsheetApp: { openById: () => { throw new Error('not stubbed'); } },
    UrlFetchApp: { fetch: () => { throw new Error('not stubbed'); } },
    ScriptApp: { getService: () => ({ getUrl: () => 'https://script.google.com/x' }) },
    Session: { getActiveUser: () => ({ getEmail: () => 'actor@example.com' }) },
    CacheService: { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {} }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
    CACHE_DURATION: { SHORT: 10, MEDIUM: 30, LONG: 300, DATABASE_LONG: 600, FORM_DATA: 30 },
    TIMEOUT_MS: { DEFAULT: 5000 },
    SYSTEM_LIMITS: {},
    getCachedProperty: () => null,
    saveToCacheWithSizeCheck: () => true,
    validateEmail: (e) => ({ isValid: typeof e === 'string' && /.+@.+/.test(e) }),
    validateUrl: () => ({ isValid: true }),
    validateSpreadsheetId: () => true,
    normalizeHeader: (h) => String(h || '').toLowerCase().trim(),
    extractFieldValueUnified: () => '',
    extractReactions: () => ({}),
    extractHighlight: () => false,
    formatTimestamp: (v) => String(v || ''),
    getQuestionText: () => '',
    createDataServiceErrorResponse: (message) => ({ success: false, message }),
    isAdministrator: () => false,
    ...overrides
  };

  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/DataApis.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'DataApis.js' });
  return context;
}

// =====================================================================
// isValidFormUrl
// =====================================================================

test('isValidFormUrl: accepts docs.google.com forms URL', () => {
  const ctx = loadDataApisContext();
  assert.equal(ctx.isValidFormUrl('https://docs.google.com/forms/d/e/1FAIpQLSd.../viewform'), true);
});

test('isValidFormUrl: accepts forms.gle short URL', () => {
  const ctx = loadDataApisContext();
  assert.equal(ctx.isValidFormUrl('https://forms.gle/abc123'), true);
});

test('isValidFormUrl: rejects http:// (requires https)', () => {
  const ctx = loadDataApisContext();
  assert.equal(ctx.isValidFormUrl('http://docs.google.com/forms/d/e/x/viewform'), false);
});

test('isValidFormUrl: rejects javascript: protocol', () => {
  const ctx = loadDataApisContext();
  assert.equal(ctx.isValidFormUrl('javascript:alert(1)'), false);
});

test('isValidFormUrl: rejects data: protocol', () => {
  const ctx = loadDataApisContext();
  assert.equal(ctx.isValidFormUrl('data:text/html,<script>alert(1)</script>'), false);
});

test('isValidFormUrl: rejects non-Google hosts', () => {
  const ctx = loadDataApisContext();
  assert.equal(ctx.isValidFormUrl('https://evil.example.com/forms'), false);
  assert.equal(ctx.isValidFormUrl('https://docs.google.com.evil.com/forms/x'), false);
});

test('isValidFormUrl: rejects docs.google.com without /forms/ or /viewform', () => {
  const ctx = loadDataApisContext();
  assert.equal(ctx.isValidFormUrl('https://docs.google.com/spreadsheets/d/xxx'), false);
  assert.equal(ctx.isValidFormUrl('https://docs.google.com/document/d/xxx'), false);
});

test('isValidFormUrl: rejects null, undefined, empty, non-string', () => {
  const ctx = loadDataApisContext();
  assert.equal(ctx.isValidFormUrl(null), false);
  assert.equal(ctx.isValidFormUrl(undefined), false);
  assert.equal(ctx.isValidFormUrl(''), false);
  assert.equal(ctx.isValidFormUrl(42), false);
  assert.equal(ctx.isValidFormUrl({}), false);
});

test('isValidFormUrl: trims whitespace before validation', () => {
  const ctx = loadDataApisContext();
  assert.equal(ctx.isValidFormUrl('  https://forms.gle/abc  '), true);
});

test('isValidFormUrl: rejects malformed URL strings', () => {
  const ctx = loadDataApisContext();
  assert.equal(ctx.isValidFormUrl('not a url'), false);
  assert.equal(ctx.isValidFormUrl('https://'), false);
});

// =====================================================================
// extractSpreadsheetInfo
// =====================================================================

test('extractSpreadsheetInfo: parses spreadsheet ID and gid', () => {
  const ctx = loadDataApisContext();
  const result = ctx.extractSpreadsheetInfo(
    'https://docs.google.com/spreadsheets/d/1abc_XYZ-123/edit#gid=456'
  );
  assert.equal(result.success, true);
  assert.equal(result.spreadsheetId, '1abc_XYZ-123');
  assert.equal(result.gid, '456');
});

test('extractSpreadsheetInfo: defaults gid to "0" when missing', () => {
  const ctx = loadDataApisContext();
  const result = ctx.extractSpreadsheetInfo(
    'https://docs.google.com/spreadsheets/d/abc123/edit'
  );
  assert.equal(result.success, true);
  assert.equal(result.spreadsheetId, 'abc123');
  assert.equal(result.gid, '0');
});

test('extractSpreadsheetInfo: handles gid after & separator', () => {
  const ctx = loadDataApisContext();
  const result = ctx.extractSpreadsheetInfo(
    'https://docs.google.com/spreadsheets/d/xyz/edit?foo=bar&gid=789'
  );
  assert.equal(result.success, true);
  assert.equal(result.gid, '789');
});

test('extractSpreadsheetInfo: rejects non-spreadsheet URLs', () => {
  const ctx = loadDataApisContext();
  const result = ctx.extractSpreadsheetInfo('https://docs.google.com/forms/d/abc');
  assert.equal(result.success, false);
});

test('extractSpreadsheetInfo: rejects null/empty/non-string', () => {
  const ctx = loadDataApisContext();
  assert.equal(ctx.extractSpreadsheetInfo(null).success, false);
  assert.equal(ctx.extractSpreadsheetInfo('').success, false);
  assert.equal(ctx.extractSpreadsheetInfo(42).success, false);
});

test('extractSpreadsheetInfo: extracts only alphanumeric IDs with hyphens/underscores', () => {
  const ctx = loadDataApisContext();
  // Valid ID chars include a-zA-Z0-9_-
  const result = ctx.extractSpreadsheetInfo(
    'https://docs.google.com/spreadsheets/d/1A-B_c-2DE/edit'
  );
  assert.equal(result.spreadsheetId, '1A-B_c-2DE');
});

test('extractSpreadsheetInfo: rejects URL lacking /spreadsheets/d/ prefix', () => {
  const ctx = loadDataApisContext();
  assert.equal(ctx.extractSpreadsheetInfo('https://example.com/d/abc').success, false);
});
