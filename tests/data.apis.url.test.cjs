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

// =====================================================================
// getSheetList
// =====================================================================

test('getSheetList: rejects empty spreadsheetId', () => {
  const ctx = loadDataApisContext();
  const result = ctx.getSheetList('');
  assert.equal(result.success, false);
  assert.match(result.message, /required/i);
});

test('getSheetList: rejects unauthenticated user', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => null
  });
  const result = ctx.getSheetList('ss-1');
  assert.equal(result.success, false);
  assert.match(result.error, /認証/);
});

test('getSheetList: returns error when openSpreadsheet throws', () => {
  const ctx = loadDataApisContext({
    openSpreadsheet: () => { throw new Error('no access'); }
  });
  const result = ctx.getSheetList('ss-1');
  assert.equal(result.success, false);
  assert.match(result.error, /アクセスできません|設定されているか/);
});

test('getSheetList: returns error when dataAccess is null', () => {
  const ctx = loadDataApisContext({
    openSpreadsheet: () => null
  });
  const result = ctx.getSheetList('ss-1234567890');
  assert.equal(result.success, false);
});

test('getSheetList: returns list of sheets with dimensions', () => {
  const mockSheets = [
    {
      getName: () => 'Sheet1',
      getSheetId: () => 0
    },
    {
      getName: () => 'Responses',
      getSheetId: () => 1
    }
  ];
  const ctx = loadDataApisContext({
    openSpreadsheet: () => ({ spreadsheet: { getSheets: () => mockSheets } }),
    getSheetInfo: (sheet) => ({
      lastRow: sheet.getName() === 'Sheet1' ? 10 : 5,
      lastCol: 7,
      headers: []
    })
  });
  const result = ctx.getSheetList('ss-1');
  assert.equal(result.success, true);
  assert.equal(result.sheets.length, 2);
  assert.equal(result.sheets[0].name, 'Sheet1');
  assert.equal(result.sheets[0].rowCount, 10);
  assert.equal(result.sheets[0].columnCount, 7);
  assert.equal(result.sheets[1].name, 'Responses');
  assert.equal(result.sheets[1].rowCount, 5);
});

// =====================================================================
// getDataCount
// =====================================================================

test('getDataCount: returns auth error without session', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => null
  });
  const result = ctx.getDataCount(null, 'newest', false);
  assert.match(result.error, /Authentication/);
  assert.equal(result.count, 0);
});

test('getDataCount: returns user-not-found when email unknown', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'ghost@example.com',
    findUserByEmail: () => null
  });
  const result = ctx.getDataCount(null, 'newest', false);
  assert.match(result.error, /User not found/);
  assert.equal(result.count, 0);
});

test('getDataCount: returns count and sheetName on success', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getUserSheetData: () => ({
      success: true,
      data: [{ id: 1 }, { id: 2 }, { id: 3 }],
      sheetName: 'Sheet1'
    })
  });
  const result = ctx.getDataCount(null, 'newest', false);
  assert.equal(result.success, true);
  assert.equal(result.count, 3);
  assert.equal(result.sheetName, 'Sheet1');
});

test('getDataCount: returns 0 count when getUserSheetData fails', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getUserSheetData: () => ({ success: false, message: 'sheet gone' })
  });
  const result = ctx.getDataCount(null, 'newest', false);
  assert.equal(result.count, 0);
  assert.match(result.error, /sheet gone/);
});

// =====================================================================
// getNotificationUpdate
// =====================================================================

test('getNotificationUpdate: rejects invalid request (no email or targetUserId)', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => null
  });
  assert.equal(ctx.getNotificationUpdate('u1', {}).success, false);

  const ctx2 = loadDataApisContext({
    getCurrentEmail: () => 'a@x.com'
  });
  assert.equal(ctx2.getNotificationUpdate(null, {}).success, false);
});

test('getNotificationUpdate: rejects unknown target user', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'a@x.com',
    findUserById: () => null
  });
  const result = ctx.getNotificationUpdate('ghost', {});
  assert.equal(result.success, false);
});

test('getNotificationUpdate: blocks non-admin non-owner on unpublished board', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'viewer@example.com',
    findUserById: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ isPublished: false }),
    isAdministrator: () => false
  });
  const result = ctx.getNotificationUpdate('u1', {});
  assert.equal(result.success, false);
  assert.match(result.message, /Access denied/);
});

test('getNotificationUpdate: reports hasNewContent=false when no items newer than lastUpdateTime', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'viewer@example.com',
    findUserById: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ isPublished: true }),
    getUserSheetData: () => ({
      success: true,
      data: [
        { timestamp: '2026-04-18T10:00:00Z' },
        { timestamp: '2026-04-18T11:00:00Z' }
      ]
    })
  });
  const result = ctx.getNotificationUpdate('u1', {
    lastUpdateTime: '2026-04-19T00:00:00Z'
  });
  assert.equal(result.success, true);
  assert.equal(result.hasNewContent, false);
  assert.equal(result.newItemsCount, 0);
});

test('getNotificationUpdate: reports hasNewContent=true when items newer than lastUpdateTime', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'viewer@example.com',
    findUserById: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ isPublished: true }),
    getUserSheetData: () => ({
      success: true,
      data: [
        { timestamp: '2026-04-18T10:00:00Z' }, // older
        { timestamp: '2026-04-19T12:00:00Z' }, // newer
        { timestamp: '2026-04-19T13:00:00Z' }  // newer
      ]
    })
  });
  const result = ctx.getNotificationUpdate('u1', {
    lastUpdateTime: '2026-04-19T00:00:00Z'
  });
  assert.equal(result.success, true);
  assert.equal(result.hasNewContent, true);
  assert.equal(result.newItemsCount, 2);
});

// =====================================================================
// processFormUrlInput — URL validation (doesn't require FormApp)
// =====================================================================

test('processFormUrlInput: rejects non-string input', () => {
  const ctx = loadDataApisContext();
  assert.equal(ctx.processFormUrlInput(null).success, false);
  assert.equal(ctx.processFormUrlInput(42).success, false);
  assert.equal(ctx.processFormUrlInput('').success, false);
});

test('processFormUrlInput: rejects URL without /forms/d/ or forms.gle/', () => {
  const ctx = loadDataApisContext();
  assert.equal(ctx.processFormUrlInput('https://docs.google.com/spreadsheets/d/abc').success, false);
  assert.equal(ctx.processFormUrlInput('https://example.com/path').success, false);
});

test('processFormUrlInput: accepts /forms/d/ and /forms.gle/ URLs', () => {
  const ctx = loadDataApisContext({
    FormApp: {
      openByUrl: () => { throw new Error('dummy - we only test URL gate'); }
    }
  });
  // These pass URL gate but fail at FormApp.openByUrl with a specific error
  const r1 = ctx.processFormUrlInput('https://docs.google.com/forms/d/abc/viewform');
  assert.equal(r1.success, false);
  assert.match(r1.error, /アクセスできません/); // Got past URL gate, failed at FormApp

  const r2 = ctx.processFormUrlInput('https://forms.gle/abc');
  assert.equal(r2.success, false);
  assert.match(r2.error, /アクセスできません/);
});

// =====================================================================
// saveConfig
// =====================================================================

test('saveConfig: returns auth error without session', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => null
  });
  const result = ctx.saveConfig({}, {});
  assert.equal(result.success, false);
  assert.match(result.message, /認証/);
});

test('saveConfig: returns user-not-found when email unknown', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'ghost@example.com',
    findUserByEmail: () => null
  });
  const result = ctx.saveConfig({}, {});
  assert.equal(result.success, false);
  assert.match(result.message, /ユーザー/);
});

test('saveConfig: delegates to saveUserConfig with isMainConfig by default', () => {
  let capturedOpts = null;
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    saveUserConfig: (_userId, _config, opts) => {
      capturedOpts = opts;
      return { success: true };
    }
  });
  ctx.saveConfig({ spreadsheetId: 'ss' });
  assert.equal(capturedOpts.isMainConfig, true);
  assert.equal(capturedOpts.isDraft, undefined);
});

test('saveConfig: uses isDraft when options.isDraft is truthy', () => {
  let capturedOpts = null;
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    saveUserConfig: (_userId, _config, opts) => {
      capturedOpts = opts;
      return { success: true };
    }
  });
  ctx.saveConfig({ spreadsheetId: 'ss' }, { isDraft: true });
  assert.equal(capturedOpts.isDraft, true);
  assert.equal(capturedOpts.isMainConfig, undefined);
});

test('saveConfig: forwards userId and config to saveUserConfig unchanged', () => {
  let capturedUserId = null;
  let capturedConfig = null;
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    saveUserConfig: (userId, config) => {
      capturedUserId = userId;
      capturedConfig = config;
      return { success: true };
    }
  });
  const input = { spreadsheetId: 'ss', sheetName: 'Sheet1', columnMapping: { answer: 2 } };
  ctx.saveConfig(input);
  assert.equal(capturedUserId, 'u1');
  assert.equal(capturedConfig.spreadsheetId, 'ss');
  assert.equal(capturedConfig.columnMapping.answer, 2);
});

test('saveConfig: propagates saveUserConfig error', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    saveUserConfig: () => ({ success: false, message: 'validation failed: invalid sheetName' })
  });
  const result = ctx.saveConfig({});
  assert.equal(result.success, false);
  assert.match(result.message, /validation failed/);
});

test('saveConfig: returns exception response when saveUserConfig throws', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    saveUserConfig: () => { throw new Error('storage down'); }
  });
  const result = ctx.saveConfig({});
  assert.equal(result.success, false);
  assert.match(result.message, /storage down/);
});

// =====================================================================
// getActiveFormInfo
// =====================================================================

test('getActiveFormInfo: returns auth error without session', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => null
  });
  const result = ctx.getActiveFormInfo(null);
  assert.equal(result.success, false);
  assert.equal(result.formUrl, null);
  assert.equal(result.formTitle, null);
});

test('getActiveFormInfo: resolves current user when no userId given', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({
      formUrl: 'https://docs.google.com/forms/d/abc/viewform',
      formTitle: 'クラスアンケート'
    })
  });
  const result = ctx.getActiveFormInfo(null);
  assert.equal(result.success, true);
  assert.equal(result.shouldShow, true);
  assert.equal(result.formTitle, 'クラスアンケート');
  assert.equal(result.isValidUrl, true);
});

test('getActiveFormInfo: returns user-not-found when self-lookup fails', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'ghost@example.com',
    findUserByEmail: () => null
  });
  const result = ctx.getActiveFormInfo(null);
  assert.equal(result.success, false);
  assert.match(result.message, /User not found/);
});

test('getActiveFormInfo: uses explicit userId over current user', () => {
  let findByEmailCalled = false;
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'viewer@example.com',
    findUserByEmail: () => { findByEmailCalled = true; return null; },
    getConfigOrDefault: (uid) => {
      assert.equal(uid, 'other-user-id');
      return { formUrl: 'https://forms.gle/xyz', formTitle: 'Other' };
    }
  });
  const result = ctx.getActiveFormInfo('other-user-id');
  assert.equal(result.success, true);
  assert.equal(findByEmailCalled, false, 'Must not look up current user when id given');
});

test('getActiveFormInfo: reports shouldShow=false when formUrl absent', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({})
  });
  const result = ctx.getActiveFormInfo(null);
  assert.equal(result.success, false);
  assert.equal(result.shouldShow, false);
  assert.equal(result.formUrl, null);
});

test('getActiveFormInfo: reports isValidUrl=false when formUrl is non-Google URL', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ formUrl: 'https://evil.example.com/fake-form' })
  });
  const result = ctx.getActiveFormInfo(null);
  assert.equal(result.shouldShow, true); // hasFormUrl=true → shouldShow
  assert.equal(result.isValidUrl, false); // but isValidFormUrl rejects it
});

test('getActiveFormInfo: whitespace-only formUrl treated as absent', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ formUrl: '   ' })
  });
  const result = ctx.getActiveFormInfo(null);
  assert.equal(result.shouldShow, false);
});

test('getActiveFormInfo: default formTitle when absent from config', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ formUrl: 'https://docs.google.com/forms/d/abc/viewform' })
  });
  const result = ctx.getActiveFormInfo(null);
  assert.equal(result.formTitle, 'フォーム');
});

// =====================================================================
// getSheetNameFromGid
// =====================================================================

test('getSheetNameFromGid: returns sheet name matching gid', () => {
  const mockSheets = [
    { getName: () => 'Sheet1', getSheetId: () => 0 },
    { getName: () => 'フォームの回答 1', getSheetId: () => 12345 },
    { getName: () => 'Data', getSheetId: () => 67890 }
  ];
  const ctx = loadDataApisContext({
    SpreadsheetApp: {
      openById: () => ({ getSheets: () => mockSheets })
    },
    executeWithRetry: (fn) => fn()
  });
  assert.equal(ctx.getSheetNameFromGid('ss-1', '12345'), 'フォームの回答 1');
  assert.equal(ctx.getSheetNameFromGid('ss-1', '67890'), 'Data');
});

test('getSheetNameFromGid: falls back to first sheet when gid not found', () => {
  const mockSheets = [
    { getName: () => 'FirstSheet', getSheetId: () => 0 },
    { getName: () => 'SecondSheet', getSheetId: () => 100 }
  ];
  const ctx = loadDataApisContext({
    SpreadsheetApp: {
      openById: () => ({ getSheets: () => mockSheets })
    },
    executeWithRetry: (fn) => fn()
  });
  assert.equal(ctx.getSheetNameFromGid('ss-1', '999'), 'FirstSheet');
});

test('getSheetNameFromGid: returns "Sheet1" when openById throws', () => {
  const ctx = loadDataApisContext({
    SpreadsheetApp: { openById: () => { throw new Error('denied'); } },
    executeWithRetry: (fn) => fn()
  });
  assert.equal(ctx.getSheetNameFromGid('ss-1', '0'), 'Sheet1');
});

test('getSheetNameFromGid: returns "Sheet1" when no sheets present', () => {
  const ctx = loadDataApisContext({
    SpreadsheetApp: {
      openById: () => ({ getSheets: () => [] })
    },
    executeWithRetry: (fn) => fn()
  });
  assert.equal(ctx.getSheetNameFromGid('ss-1', '0'), 'Sheet1');
});

// =====================================================================
// connectDataSource / processDataSourceOperations
//
// These call same-file functions (getColumnAnalysis, getFormInfoInternal),
// so we monkey-patch the loaded context after construction rather than
// via harness overrides — vm.runInContext's function hoisting would
// otherwise clobber any overrides we pass in.
// =====================================================================

test('connectDataSource: rejects unauthenticated user', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => null
  });
  const result = ctx.connectDataSource('ss-1', 'Sheet1');
  assert.equal(result.success, false);
  assert.match(result.error, /認証/);
});

test('connectDataSource: attempts domain-wide sharing but tolerates failure', () => {
  let sharingAttempted = false;
  let analysisCalled = false;
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    setupDomainWideSharing: () => {
      sharingAttempted = true;
      throw new Error('sharing denied');
    }
  });
  ctx.getColumnAnalysis = () => {
    analysisCalled = true;
    return { success: true, mapping: {}, headers: [] };
  };
  const result = ctx.connectDataSource('ss-1', 'Sheet1');
  assert.equal(sharingAttempted, true);
  assert.equal(analysisCalled, true, 'Sharing failure must not block the rest of the flow');
  assert.equal(result.success, true);
});

test('connectDataSource: delegates to getColumnAnalysis when no batch operations', () => {
  let capturedArgs = null;
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    setupDomainWideSharing: () => {}
  });
  ctx.getColumnAnalysis = (ss, sheet) => {
    capturedArgs = { ss, sheet };
    return { success: true };
  };
  ctx.connectDataSource('my-ss', 'my-sheet');
  assert.equal(capturedArgs.ss, 'my-ss');
  assert.equal(capturedArgs.sheet, 'my-sheet');
});

test('connectDataSource: delegates to processDataSourceOperations when batch given', () => {
  let columnCallCount = 0;
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    setupDomainWideSharing: () => {}
  });
  ctx.getColumnAnalysis = () => {
    columnCallCount += 1;
    return { success: true, mapping: {}, headers: [] };
  };
  ctx.getFormInfoInternal = () => ({ formData: { formUrl: 'https://docs.google.com/forms/d/x' } });
  const result = ctx.connectDataSource('ss-1', 'Sheet1', [
    { type: 'validateAccess' },
    { type: 'getFormInfo' },
    { type: 'connectDataSource' }
  ]);
  assert.equal(result.success, true);
  assert.equal(columnCallCount, 1, 'Column analysis should be cached across ops');
  assert.ok(result.batchResults.validation);
  assert.ok(result.batchResults.formInfo);
});

test('processDataSourceOperations: reports failure in validation branch', () => {
  const ctx = loadDataApisContext();
  ctx.getColumnAnalysis = () => ({ success: false, message: 'permission denied' });
  const result = ctx.processDataSourceOperations('ss-1', 'Sheet1', [
    { type: 'validateAccess' }
  ]);
  assert.equal(result.batchResults.validation.success, false);
  assert.match(result.batchResults.validation.details.connectionError, /permission denied/);
});

test('processDataSourceOperations: marks overall success=false when connect op fails', () => {
  const ctx = loadDataApisContext();
  ctx.getColumnAnalysis = () => ({
    success: false,
    message: 'header integrity issue',
    errorResponse: { message: 'header integrity issue' }
  });
  const result = ctx.processDataSourceOperations('ss-1', 'Sheet1', [
    { type: 'connectDataSource' }
  ]);
  assert.equal(result.success, false);
  assert.match(result.error, /header integrity issue/);
});

test('processDataSourceOperations: unknown op type is silently skipped', () => {
  const ctx = loadDataApisContext();
  ctx.getColumnAnalysis = () => ({ success: true });
  const result = ctx.processDataSourceOperations('ss-1', 'Sheet1', [
    { type: 'unknownOp' }
  ]);
  assert.equal(result.success, true);
  assert.deepEqual({ ...result.batchResults }, {});
});

test('processDataSourceOperations: handles empty operations array', () => {
  const ctx = loadDataApisContext();
  const result = ctx.processDataSourceOperations('ss-1', 'Sheet1', []);
  assert.equal(result.success, true);
  assert.deepEqual({ ...result.batchResults }, {});
});

// =====================================================================
// getColumnAnalysis
// =====================================================================

test('getColumnAnalysis: rejects unauthenticated user', () => {
  const ctx = loadDataApisContext({ getCurrentEmail: () => null });
  const result = ctx.getColumnAnalysis('ss-1', 'Sheet1');
  assert.equal(result.success, false);
  assert.match(result.error, /認証/);
});

test('getColumnAnalysis: returns error when openSpreadsheet returns null', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    openSpreadsheet: () => null
  });
  const result = ctx.getColumnAnalysis('ss-1', 'Sheet1');
  assert.equal(result.success, false);
  assert.match(result.error, /アクセスできません/);
});

test('getColumnAnalysis: returns error when openSpreadsheet throws', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    openSpreadsheet: () => { throw new Error('denied'); }
  });
  const result = ctx.getColumnAnalysis('ss-1', 'Sheet1');
  assert.equal(result.success, false);
  assert.match(result.error, /アクセス権がありません/);
});

test('getColumnAnalysis: returns sheet-not-found when getSheet returns null', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    openSpreadsheet: () => ({ getSheet: () => null })
  });
  const result = ctx.getColumnAnalysis('ss-1', 'Missing');
  assert.equal(result.success, false);
  assert.match(result.message, /Sheet not found/);
});

test('getColumnAnalysis: returns headers and mapping on success', () => {
  const sheet = {
    getDataRange: () => ({ getValues: () => [['Q1', '理由', '名前']] }),
    getRange: () => ({ getValues: () => [] })
  };
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    openSpreadsheet: () => ({ getSheet: () => sheet })
  });
  // Stub performIntegratedColumnDiagnostics via in-context patching to avoid
  // requiring the real ColumnMappingService in this test.
  ctx.performIntegratedColumnDiagnostics = () => ({
    recommendedMapping: { answer: 0, reason: 1, name: 2 },
    confidence: { answer: 90 }
  });
  // Suppress column-setup side effect path by forcing no columns added
  ctx.setupReactionAndHighlightColumns = () => ({ columnsAdded: [] });

  const result = ctx.getColumnAnalysis('ss-1', 'Sheet1');
  assert.equal(result.success, true);
  assert.equal(result.mapping.answer, 0);
  assert.ok(result.headers.length === 3);
});

// =====================================================================
// getFormInfoInternal
// =====================================================================

test('getFormInfoInternal: delegates to getFormInfo when present', () => {
  let capturedArgs = null;
  const ctx = loadDataApisContext();
  ctx.getFormInfo = (ss, sheet) => {
    capturedArgs = { ss, sheet };
    return { success: true, formData: { formUrl: 'https://docs.google.com/forms/d/abc' } };
  };
  const result = ctx.getFormInfoInternal('ss-1', 'Sheet1');
  assert.equal(result.success, true);
  assert.equal(capturedArgs.ss, 'ss-1');
  assert.equal(capturedArgs.sheet, 'Sheet1');
});

test('getFormInfoInternal: returns not-initialized error when getFormInfo missing', () => {
  const ctx = loadDataApisContext();
  ctx.getFormInfo = undefined;
  const result = ctx.getFormInfoInternal('ss-1', 'Sheet1');
  assert.equal(result.success, false);
  assert.match(result.message, /初期化されていません/);
});

test('getFormInfoInternal: catches getFormInfo exception', () => {
  const ctx = loadDataApisContext();
  ctx.getFormInfo = () => { throw new Error('form api down'); };
  const result = ctx.getFormInfoInternal('ss-1', 'Sheet1');
  assert.equal(result.success, false);
  assert.match(result.message, /form api down/);
});
