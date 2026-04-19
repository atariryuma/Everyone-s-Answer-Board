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
    DEFAULT_DISPLAY_SETTINGS: { showNames: false, showReactions: true, theme: 'default', pageSize: 20 },
    getBatchedAdminAuth: () => ({ success: true, authenticated: true, email: 'viewer@example.com', isAdmin: false }),
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

// =====================================================================
// getBoardInfo
// =====================================================================

test('getBoardInfo: returns auth error when no email', () => {
  const ctx = loadDataApisContext({ getCurrentEmail: () => null });
  const result = ctx.getBoardInfo();
  assert.equal(result.success, false);
});

test('getBoardInfo: returns user-not-found when lookup fails', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'nobody@example.com',
    findUserByEmail: () => null
  });
  const result = ctx.getBoardInfo();
  assert.equal(result.success, false);
  assert.match(result.message, /User not found/);
});

test('getBoardInfo: includes isPublished, URLs, and lastUpdated', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'owner@example.com', lastModified: '2026-01-01' }),
    getConfigOrDefault: () => ({ isPublished: true, publishedAt: '2026-04-19' }),
    getQuestionText: () => '今日の感想は？',
    getWebAppUrl: () => 'https://script.google.com/macros/s/abc/exec'
  });

  const result = ctx.getBoardInfo();
  assert.equal(result.success, true);
  assert.equal(result.isPublished, true);
  assert.equal(result.questionText, '今日の感想は？');
  assert.ok(result.urls.view.includes('userId=u1'));
  assert.ok(result.urls.view.includes('mode=view'));
  assert.ok(result.urls.admin.includes('mode=admin'));
  assert.equal(result.lastUpdated, '2026-04-19');
});

test('getBoardInfo: lastUpdated falls back to user.lastModified when not published', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'owner@example.com', lastModified: '2026-01-01' }),
    getConfigOrDefault: () => ({ isPublished: false }),
    getWebAppUrl: () => 'https://x.example.com'
  });
  const result = ctx.getBoardInfo();
  assert.equal(result.isPublished, false);
  assert.equal(result.lastUpdated, '2026-01-01');
});

// =====================================================================
// validateHeaderIntegrity
// =====================================================================

test('validateHeaderIntegrity: returns user-not-found without session or target', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => null,
    findUserById: () => null,
    findUserByEmail: () => null
  });
  const result = ctx.validateHeaderIntegrity(null);
  assert.equal(result.success, false);
});

test('validateHeaderIntegrity: returns error when spreadsheet config incomplete', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserById: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ spreadsheetId: '', sheetName: '' })
  });
  const result = ctx.validateHeaderIntegrity('u1');
  assert.equal(result.success, false);
  assert.match(result.error, /configuration incomplete/i);
});

test('validateHeaderIntegrity: returns error when target sheet not found', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserById: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ spreadsheetId: 'ss', sheetName: 'Missing' }),
    openSpreadsheet: () => ({ getSheet: () => null })
  });
  const result = ctx.validateHeaderIntegrity('u1');
  assert.equal(result.success, false);
  assert.match(result.error, /Sheet not found/);
});

test('validateHeaderIntegrity: returns valid=true for healthy headers', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserById: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ spreadsheetId: 'ss', sheetName: 'Sheet1' }),
    openSpreadsheet: () => ({ getSheet: () => ({ name: 'Sheet1' }) }),
    getSheetInfo: () => ({ headers: ['Q1', '理由', '名前', 'LIKE'], lastRow: 10, lastCol: 4 })
  });
  const result = ctx.validateHeaderIntegrity('u1');
  assert.equal(result.success, true);
  assert.equal(result.valid, true);
  assert.equal(result.headerCount, 4);
  assert.equal(result.emptyHeaderCount, 0);
});

test('validateHeaderIntegrity: reports empty-header count', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserById: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ spreadsheetId: 'ss', sheetName: 'Sheet1' }),
    openSpreadsheet: () => ({ getSheet: () => ({}) }),
    getSheetInfo: () => ({ headers: ['Q1', '', '  ', 'LIKE'], lastRow: 2, lastCol: 4 })
  });
  const result = ctx.validateHeaderIntegrity('u1');
  assert.equal(result.valid, true); // at least one non-empty header
  assert.equal(result.emptyHeaderCount, 2);
});

test('validateHeaderIntegrity: returns valid=false when all headers are empty', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserById: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ spreadsheetId: 'ss', sheetName: 'Sheet1' }),
    openSpreadsheet: () => ({ getSheet: () => ({}) }),
    getSheetInfo: () => ({ headers: ['', '  ', null], lastRow: 0, lastCol: 3 })
  });
  const result = ctx.validateHeaderIntegrity('u1');
  assert.equal(result.success, false);
  assert.equal(result.valid, false);
});

// =====================================================================
// getPublishedSheetData — authorization decisions
// =====================================================================

function ctxWithAuth(overrides = {}) {
  return loadDataApisContext({
    getBatchedAdminAuth: () => ({
      success: true,
      authenticated: true,
      email: 'viewer@example.com',
      isAdmin: false
    }),
    getUserSheetData: () => ({
      success: true,
      data: [{ id: 1 }],
      headers: ['Q1'],
      sheetName: 'Sheet1',
      header: '',
      showDetails: true
    }),
    getQuestionText: () => '',
    ...overrides
  });
}

test('getPublishedSheetData: rejects unauthenticated viewer', () => {
  const ctx = ctxWithAuth({
    getBatchedAdminAuth: () => ({ success: false, authenticated: false })
  });
  const result = ctx.getPublishedSheetData(null, 'newest', false, 'u1');
  assert.equal(result.success, false);
  assert.match(result.error, /Authentication/);
});

test('getPublishedSheetData: rejects when target user not found', () => {
  const ctx = ctxWithAuth({
    findUserById: () => null
  });
  const result = ctx.getPublishedSheetData(null, 'newest', false, 'ghost');
  assert.equal(result.success, false);
  assert.match(result.error, /Target user/);
});

test('getPublishedSheetData: blocks non-admin non-owner viewer from unpublished board', () => {
  let getUserSheetDataCalled = false;
  const ctx = ctxWithAuth({
    findUserById: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ spreadsheetId: 'ss', sheetName: 'Sheet1', isPublished: false }),
    getUserSheetData: () => { getUserSheetDataCalled = true; return { success: true }; }
  });
  const result = ctx.getPublishedSheetData(null, 'newest', false, 'u1');
  assert.equal(result.success, false);
  assert.match(result.error, /未公開/);
  assert.equal(getUserSheetDataCalled, false, 'Must not fetch data for unauthorized viewer');
});

test('getPublishedSheetData: allows non-admin on a published board', () => {
  const ctx = ctxWithAuth({
    findUserById: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ spreadsheetId: 'ss', sheetName: 'Sheet1', isPublished: true })
  });
  const result = ctx.getPublishedSheetData(null, 'newest', false, 'u1');
  assert.equal(result.success, true);
});

test('getPublishedSheetData: allows owner to view their own unpublished board', () => {
  const ctx = ctxWithAuth({
    getBatchedAdminAuth: () => ({
      success: true, authenticated: true,
      email: 'owner@example.com', isAdmin: false
    }),
    findUserById: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ spreadsheetId: 'ss', sheetName: 'Sheet1', isPublished: false })
  });
  const result = ctx.getPublishedSheetData(null, 'newest', false, 'u1');
  assert.equal(result.success, true);
});

test('getPublishedSheetData: allows system admin regardless of publish state', () => {
  const ctx = ctxWithAuth({
    getBatchedAdminAuth: () => ({
      success: true, authenticated: true,
      email: 'admin@example.com', isAdmin: true
    }),
    findUserById: () => ({ userId: 'u1', userEmail: 'other@example.com' }),
    getConfigOrDefault: () => ({ spreadsheetId: 'ss', sheetName: 'Sheet1', isPublished: false })
  });
  const result = ctx.getPublishedSheetData(null, 'newest', false, 'u1');
  assert.equal(result.success, true);
});

test('getPublishedSheetData: forwards classFilter and sortOrder to getUserSheetData', () => {
  let capturedOptions = null;
  const ctx = ctxWithAuth({
    findUserById: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ spreadsheetId: 'ss', sheetName: 'Sheet1', isPublished: true }),
    getUserSheetData: (_userId, options) => {
      capturedOptions = options;
      return { success: true, data: [], headers: [], sheetName: '' };
    }
  });
  ctx.getPublishedSheetData('3-1', 'oldest', false, 'u1');
  assert.equal(capturedOptions.classFilter, '3-1');
  assert.equal(capturedOptions.sortBy, 'oldest');
});

test('getPublishedSheetData: classFilter="すべて" is converted to undefined', () => {
  let capturedOptions = null;
  const ctx = ctxWithAuth({
    findUserById: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ spreadsheetId: 'ss', sheetName: 'Sheet1', isPublished: true }),
    getUserSheetData: (_userId, options) => {
      capturedOptions = options;
      return { success: true, data: [], headers: [], sheetName: '' };
    }
  });
  ctx.getPublishedSheetData('すべて', 'newest', false, 'u1');
  assert.equal(capturedOptions.classFilter, undefined);
});
