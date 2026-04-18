const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function createMockCache(initial = {}) {
  const store = new Map(Object.entries(initial));
  return {
    get: (k) => (store.has(k) ? store.get(k) : null),
    put: (k, v, _ttl) => { store.set(k, v); },
    remove: (k) => { store.delete(k); },
    _store: store,
    _puts: []
  };
}

function createCacheWithPutSpy() {
  const store = new Map();
  const puts = [];
  return {
    get: (k) => (store.has(k) ? store.get(k) : null),
    put: (k, v, ttl) => { store.set(k, v); puts.push({ k, v, ttl }); },
    remove: (k) => { store.delete(k); },
    _store: store,
    _puts: puts
  };
}

function createMockSheet({
  headers = [],
  rows = [],
  sheetName = 'Sheet1',
  spreadsheetId = 'ss-1'
} = {}) {
  const data = [headers.slice(), ...rows.map((r) => r.slice())];
  return {
    getName: () => sheetName,
    getParent: () => ({ getId: () => spreadsheetId }),
    getDataRange: () => ({
      getValues: () => data.map((r) => r.slice())
    }),
    getLastRow: () => data.length,
    getRange: () => ({ getValues: () => [], setValues: () => {} })
  };
}

function loadDataServiceContext(overrides = {}) {
  const cache = overrides.cache || createMockCache();

  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    CacheService: { getScriptCache: () => cache },
    CACHE_DURATION: { DATABASE_LONG: 600, FORM_DATA: 30 },
    createErrorResponse: (message, data, extra) => ({ success: false, message, ...extra }),
    createExceptionResponse: (error) => ({ success: false, message: error.message }),
    createDataServiceErrorResponse: (message, sheetName = '') => ({
      success: false, message, data: [], headers: [], sheetName
    }),
    getCurrentEmail: () => 'actor@example.com',
    findUserByEmail: () => null,
    findUserById: () => null,
    findUserBySpreadsheetId: () => null,
    openSpreadsheet: () => null,
    getUserConfig: () => ({ success: true, config: {} }),
    getConfigOrDefault: () => ({}),
    normalizeHeader: (h) => String(h || '').toLowerCase().trim(),
    extractFieldValueUnified: () => '',
    extractReactions: () => ({}),
    extractHighlight: () => false,
    getQuestionText: () => '',
    formatTimestamp: (v) => String(v || ''),
    getCachedProperty: () => null,
    ...overrides
  };

  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/DataService.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'DataService.js' });
  context._cache = cache;
  return context;
}

// =====================================================================
// getSheetHeaders
// =====================================================================

test('getSheetHeaders: fetches from sheet when cache empty and caches result', () => {
  const cache = createCacheWithPutSpy();
  const ctx = loadDataServiceContext({ cache });
  const sheet = createMockSheet({
    headers: ['Q1', 'UNDERSTAND', 'LIKE', 'CURIOUS', 'HIGHLIGHT']
  });

  const info = ctx.getSheetHeaders(sheet);

  assert.equal(info.lastCol, 5);
  assert.deepEqual(info.headers, ['Q1', 'UNDERSTAND', 'LIKE', 'CURIOUS', 'HIGHLIGHT']);
  assert.equal(cache._puts.length, 1);
  assert.equal(cache._puts[0].ttl, 600); // DATABASE_LONG
});

test('getSheetHeaders: returns cached value without touching sheet', () => {
  const cachedInfo = { lastCol: 3, headers: ['A', 'B', 'C'] };
  const cache = createMockCache({ 'sheet_headers_ss-1_Sheet1': JSON.stringify(cachedInfo) });
  const ctx = loadDataServiceContext({ cache });

  let dataRangeCalls = 0;
  const sheet = {
    getName: () => 'Sheet1',
    getParent: () => ({ getId: () => 'ss-1' }),
    getDataRange: () => { dataRangeCalls += 1; return { getValues: () => [[]] }; },
    getLastRow: () => 0
  };

  const info = ctx.getSheetHeaders(sheet);

  assert.equal(info.lastCol, 3);
  assert.deepEqual([...info.headers], ['A', 'B', 'C']);
  assert.equal(dataRangeCalls, 0, 'Must not call getDataRange on cache hit');
});

test('getSheetHeaders: cache parse failure falls back to fresh fetch', () => {
  const cache = createMockCache({ 'sheet_headers_ss-1_Sheet1': 'not valid json' });
  const ctx = loadDataServiceContext({ cache });
  const sheet = createMockSheet({ headers: ['A', 'B'] });

  const info = ctx.getSheetHeaders(sheet);

  assert.deepEqual(info.headers, ['A', 'B']);
  assert.equal(info.lastCol, 2);
});

test('getSheetHeaders: cache write failure does not break fetch', () => {
  const cache = {
    get: () => null,
    put: () => { throw new Error('cache write failed'); },
    remove: () => {}
  };
  const ctx = loadDataServiceContext({ cache });
  const sheet = createMockSheet({ headers: ['A'] });

  const info = ctx.getSheetHeaders(sheet);

  assert.deepEqual(info.headers, ['A']);
  assert.equal(info.lastCol, 1);
});

test('getSheetHeaders: falls back to "unknown" id when sheet has no getParent', () => {
  const cache = createCacheWithPutSpy();
  const ctx = loadDataServiceContext({ cache });
  const sheet = {
    getName: () => 'Sheet1',
    getDataRange: () => ({ getValues: () => [['A', 'B']] }),
    getLastRow: () => 1
  };

  ctx.getSheetHeaders(sheet);

  assert.equal(cache._puts[0].k, 'sheet_headers_unknown_Sheet1');
});

test('getSheetHeaders: empty sheet yields zero cols and empty headers', () => {
  const cache = createMockCache();
  const ctx = loadDataServiceContext({ cache });
  const sheet = {
    getName: () => 'Empty',
    getParent: () => ({ getId: () => 'ss-1' }),
    getDataRange: () => ({ getValues: () => [] }),
    getLastRow: () => 0
  };

  const info = ctx.getSheetHeaders(sheet);

  assert.equal(info.lastCol, 0);
  assert.equal(info.headers.length, 0);
});

// =====================================================================
// invalidateSheetHeadersCache
// =====================================================================

test('invalidateSheetHeadersCache: removes the exact cache key', () => {
  const cache = createMockCache({
    'sheet_headers_ss-1_Sheet1': 'cached',
    'sheet_headers_ss-1_Other': 'other',
    'sheet_rows_ss-1_Sheet1': 'rows'
  });
  const ctx = loadDataServiceContext({ cache });

  ctx.invalidateSheetHeadersCache('ss-1', 'Sheet1');

  assert.equal(cache._store.has('sheet_headers_ss-1_Sheet1'), false);
  assert.equal(cache._store.has('sheet_headers_ss-1_Other'), true);
  assert.equal(cache._store.has('sheet_rows_ss-1_Sheet1'), true);
});

test('invalidateSheetHeadersCache: no-op when either arg is missing', () => {
  const cache = createMockCache({ 'sheet_headers_ss-1_Sheet1': 'cached' });
  const ctx = loadDataServiceContext({ cache });

  ctx.invalidateSheetHeadersCache('', 'Sheet1');
  ctx.invalidateSheetHeadersCache('ss-1', '');
  ctx.invalidateSheetHeadersCache(null, null);

  assert.equal(cache._store.has('sheet_headers_ss-1_Sheet1'), true);
});

test('invalidateSheetHeadersCache: swallows remove errors', () => {
  const cache = {
    get: () => null,
    put: () => {},
    remove: () => { throw new Error('remove failed'); }
  };
  const ctx = loadDataServiceContext({ cache });

  // Should not throw
  ctx.invalidateSheetHeadersCache('ss-1', 'Sheet1');
});

// =====================================================================
// getSheetRowCount
// =====================================================================

test('getSheetRowCount: fetches from sheet.getLastRow when cache empty', () => {
  const cache = createCacheWithPutSpy();
  const ctx = loadDataServiceContext({ cache });
  const sheet = createMockSheet({ headers: ['A'], rows: [['1'], ['2'], ['3']] });

  const rows = ctx.getSheetRowCount(sheet);

  assert.equal(rows, 4); // 1 header + 3 data
  assert.equal(cache._puts.length, 1);
  assert.equal(cache._puts[0].k, 'sheet_rows_ss-1_Sheet1');
  assert.equal(cache._puts[0].v, '4');
  assert.equal(cache._puts[0].ttl, 30); // FORM_DATA
});

test('getSheetRowCount: returns cached integer value', () => {
  const cache = createMockCache({ 'sheet_rows_ss-1_Sheet1': '42' });
  const ctx = loadDataServiceContext({ cache });

  let getLastRowCalls = 0;
  const sheet = {
    getName: () => 'Sheet1',
    getParent: () => ({ getId: () => 'ss-1' }),
    getLastRow: () => { getLastRowCalls += 1; return 0; }
  };

  assert.equal(ctx.getSheetRowCount(sheet), 42);
  assert.equal(getLastRowCalls, 0);
});

test('getSheetRowCount: invalid cached value falls through to fresh read', () => {
  const cache = createMockCache({ 'sheet_rows_ss-1_Sheet1': 'not-a-number' });
  const ctx = loadDataServiceContext({ cache });
  const sheet = createMockSheet({ headers: ['A'], rows: [['1']] });

  assert.equal(ctx.getSheetRowCount(sheet), 2);
});

// =====================================================================
// getSheetInfo
// =====================================================================

test('getSheetInfo: combines headers and row count', () => {
  const cache = createMockCache();
  const ctx = loadDataServiceContext({ cache });
  const sheet = createMockSheet({
    headers: ['Q1', 'LIKE'],
    rows: [['a', ''], ['b', 'x@x.com']]
  });

  const info = ctx.getSheetInfo(sheet);

  assert.equal(info.lastRow, 3);
  assert.equal(info.lastCol, 2);
  assert.deepEqual(info.headers, ['Q1', 'LIKE']);
});

test('getSheetInfo: uses cached header and cached row count when both present', () => {
  const cache = createMockCache({
    'sheet_headers_ss-1_Sheet1': JSON.stringify({ lastCol: 3, headers: ['X', 'Y', 'Z'] }),
    'sheet_rows_ss-1_Sheet1': '99'
  });
  const ctx = loadDataServiceContext({ cache });

  let dataRangeCalls = 0;
  let getLastRowCalls = 0;
  const sheet = {
    getName: () => 'Sheet1',
    getParent: () => ({ getId: () => 'ss-1' }),
    getDataRange: () => { dataRangeCalls += 1; return { getValues: () => [[]] }; },
    getLastRow: () => { getLastRowCalls += 1; return 0; }
  };

  const info = ctx.getSheetInfo(sheet);

  assert.equal(info.lastRow, 99);
  assert.equal(info.lastCol, 3);
  assert.deepEqual([...info.headers], ['X', 'Y', 'Z']);
  assert.equal(dataRangeCalls, 0);
  assert.equal(getLastRowCalls, 0);
});

test('getSheetInfo: row count refresh while headers cached (typical after new form post)', () => {
  // Only headers cached — rows should refresh (simulates 30s TTL expiry)
  const cache = createMockCache({
    'sheet_headers_ss-1_Sheet1': JSON.stringify({ lastCol: 2, headers: ['Q1', 'LIKE'] })
  });
  const ctx = loadDataServiceContext({ cache });
  const sheet = createMockSheet({
    headers: ['Q1', 'LIKE'],
    rows: [['a', ''], ['b', ''], ['c', '']]
  });

  const info = ctx.getSheetInfo(sheet);

  assert.equal(info.lastRow, 4);
  assert.equal(info.lastCol, 2);
  assert.deepEqual([...info.headers], ['Q1', 'LIKE']);
});

// =====================================================================
// getAdaptiveBatchSize (pure helper)
// =====================================================================

test('getAdaptiveBatchSize: 0 errors → 100', () => {
  const ctx = loadDataServiceContext();
  assert.equal(ctx.getAdaptiveBatchSize(0), 100);
});

test('getAdaptiveBatchSize: 1 error → 70', () => {
  const ctx = loadDataServiceContext();
  assert.equal(ctx.getAdaptiveBatchSize(1), 70);
});

test('getAdaptiveBatchSize: consecutive errors → 50', () => {
  const ctx = loadDataServiceContext();
  assert.equal(ctx.getAdaptiveBatchSize(2), 50);
  assert.equal(ctx.getAdaptiveBatchSize(5), 50);
  assert.equal(ctx.getAdaptiveBatchSize(100), 50);
});

// =====================================================================
// connectToSpreadsheetSheet
// =====================================================================

test('connectToSpreadsheetSheet: returns { sheet, spreadsheet } on success', () => {
  const mockSheet = createMockSheet({ headers: ['Q1'] });
  const ctx = loadDataServiceContext({
    openSpreadsheet: () => ({
      spreadsheet: { getSheetByName: (name) => (name === 'Sheet1' ? mockSheet : null) }
    })
  });

  const result = ctx.connectToSpreadsheetSheet({ spreadsheetId: 'ss', sheetName: 'Sheet1' });

  assert.equal(result.sheet, mockSheet);
});

test('connectToSpreadsheetSheet: throws when openSpreadsheet returns null', () => {
  const ctx = loadDataServiceContext({ openSpreadsheet: () => null });

  assert.throws(
    () => ctx.connectToSpreadsheetSheet({ spreadsheetId: 'ss', sheetName: 'Sheet1' }),
    /スプレッドシートアクセス/
  );
});

test('connectToSpreadsheetSheet: throws when dataAccess has no spreadsheet', () => {
  const ctx = loadDataServiceContext({ openSpreadsheet: () => ({}) });

  assert.throws(
    () => ctx.connectToSpreadsheetSheet({ spreadsheetId: 'ss', sheetName: 'Sheet1' }),
    /スプレッドシートアクセス/
  );
});

test('connectToSpreadsheetSheet: throws when sheet not found', () => {
  const ctx = loadDataServiceContext({
    openSpreadsheet: () => ({
      spreadsheet: { getSheetByName: () => null }
    })
  });

  assert.throws(
    () => ctx.connectToSpreadsheetSheet({ spreadsheetId: 'ss', sheetName: 'Missing' }),
    /'Missing'/
  );
});

// =====================================================================
// getUserSheetData: early returns (light-touch integration)
// =====================================================================

test('getUserSheetData: returns error when user not found', () => {
  const ctx = loadDataServiceContext({ findUserById: () => null });
  const result = ctx.getUserSheetData('ghost', {});
  assert.equal(result.success, false);
  assert.match(result.message, /ユーザー/);
});

test('getUserSheetData: returns error when preloadedConfig missing spreadsheetId', () => {
  const ctx = loadDataServiceContext({
    findUserById: () => ({ userId: 'u1', userEmail: 'a@x.com' })
  });
  const result = ctx.getUserSheetData('u1', {}, null, { spreadsheetId: '' });
  assert.equal(result.success, false);
  assert.match(result.message, /スプレッドシート/);
});

test('getUserSheetData: uses preloadedUser and preloadedConfig without DB hits', () => {
  let findCalls = 0;
  let configCalls = 0;
  const ctx = loadDataServiceContext({
    findUserById: () => { findCalls += 1; return null; },
    getConfigOrDefault: () => { configCalls += 1; return {}; },
    openSpreadsheet: () => null // Will fail downstream, but that's after the preload check
  });

  ctx.getUserSheetData(
    'u1',
    {},
    { userId: 'u1', userEmail: 'a@x.com' },
    { spreadsheetId: 'ss', sheetName: 'Sheet1' }
  );

  assert.equal(findCalls, 0, 'findUserById must not be called when preloadedUser provided');
  assert.equal(configCalls, 0, 'getConfigOrDefault must not be called when preloadedConfig provided');
});
