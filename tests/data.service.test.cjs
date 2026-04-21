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
    getLastColumn: () => (data[0] ? data[0].length : 0),
    getRange: (row, col, numRows = 1, numCols = 1) => ({
      getValues: () => {
        const out = [];
        for (let r = row - 1; r < row - 1 + numRows; r += 1) {
          const rowArr = [];
          for (let c = col - 1; c < col - 1 + numCols; c += 1) {
            rowArr.push(data[r] ? (data[r][c] !== undefined ? data[r][c] : '') : '');
          }
          out.push(rowArr);
        }
        return out;
      },
      setValues: () => {}
    })
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
    getLastColumn: () => 2,
    getRange: () => ({ getValues: () => [['A', 'B']] }),
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
    getLastColumn: () => 0,
    getRange: () => ({ getValues: () => [[]] }),
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

// =====================================================================
// processBatchData: regression — at least one batch attempted
// Ensures the 20s time-check does not short-circuit before any batch when
// setup (connectToSpreadsheetSheet + getSheetInfo) consumed the budget.
// =====================================================================

test('processBatchData: processes first batch even when startTime is already past budget', () => {
  const ctx = loadDataServiceContext({
    resolveColumnIndex: (headers, type, mapping) => {
      const explicit = mapping && typeof mapping[type] === 'number' ? mapping[type] : -1;
      return { index: explicit };
    }
  });
  const sheet = createMockSheet({
    headers: ['Timestamp', 'Answer'],
    rows: [
      ['2026-04-21', 'first'],
      ['2026-04-21', 'second'],
      ['2026-04-21', 'third']
    ]
  });
  // startTime が budget(20s) を超えた状態を模擬する。
  // バグ前: 先頭で time-check に引っかかって 0 行で break する。
  // 修正後: 少なくとも 1 バッチは試行されて全 3 行が処理される。
  const staleStartTime = Date.now() - 30000;

  const result = ctx.processBatchData(
    sheet,
    ['Timestamp', 'Answer'],
    4, // lastRow (1 header + 3 data)
    2, // lastCol
    { columnMapping: { answer: 1 } },
    {},
    null,
    staleStartTime
  );

  assert.equal(result.length, 3, 'すべてのデータ行が処理されるべき');
});

// =====================================================================
// deleteLinkedFormResponseByTimestamp
// =====================================================================

test('deleteLinkedFormResponseByTimestamp: deletes response matching timestamp', () => {
  const deletedIds = [];
  const response = {
    getTimestamp: () => new Date('2026-04-21T12:00:00.500Z'),
    getId: () => 'resp-1'
  };
  const form = {
    getResponses: () => [response],
    deleteResponse: (id) => { deletedIds.push(id); }
  };
  const ctx = loadDataServiceContext({
    FormApp: { openByUrl: () => form }
  });
  const sheet = {
    getFormUrl: () => 'https://docs.google.com/forms/d/1ABC/edit'
  };

  // 1 秒以内の誤差は許容
  const result = ctx.deleteLinkedFormResponseByTimestamp(sheet, new Date('2026-04-21T12:00:00.000Z'));

  assert.equal(result.success, true);
  assert.deepEqual(deletedIds, ['resp-1']);
});

test('deleteLinkedFormResponseByTimestamp: returns note when sheet has no linked form', () => {
  const ctx = loadDataServiceContext({
    FormApp: { openByUrl: () => { throw new Error('should not be called'); } }
  });
  const sheet = { getFormUrl: () => null };

  const result = ctx.deleteLinkedFormResponseByTimestamp(sheet, new Date());

  assert.equal(result.success, false);
  assert.equal(result.message, 'no linked form');
});

test('deleteLinkedFormResponseByTimestamp: returns note when no response matches timestamp', () => {
  const form = {
    getResponses: () => [{
      getTimestamp: () => new Date('2020-01-01'),
      getId: () => 'old-resp'
    }],
    deleteResponse: () => { throw new Error('should not delete mismatched response'); }
  };
  const ctx = loadDataServiceContext({
    FormApp: { openByUrl: () => form }
  });
  const sheet = { getFormUrl: () => 'https://docs.google.com/forms/d/1X/edit' };

  const result = ctx.deleteLinkedFormResponseByTimestamp(sheet, new Date('2026-04-21T12:00:00Z'));

  assert.equal(result.success, false);
  assert.equal(result.message, 'no matching form response');
});

test('deleteLinkedFormResponseByTimestamp: swallows FormApp errors and reports note', () => {
  const ctx = loadDataServiceContext({
    FormApp: { openByUrl: () => { throw new Error('permission denied'); } }
  });
  const sheet = { getFormUrl: () => 'https://docs.google.com/forms/d/1X/edit' };

  const result = ctx.deleteLinkedFormResponseByTimestamp(sheet, new Date());

  assert.equal(result.success, false);
  assert.equal(result.message, 'permission denied');
});

// =====================================================================
// applySortAndLimit — sort and limit are applied independently
// =====================================================================

test('applySortAndLimit: applies limit even when sortBy is undefined', () => {
  const ctx = loadDataServiceContext();
  const data = Array.from({ length: 5 }, (_, i) => ({ id: i, timestamp: `2026-04-0${i + 1}` }));
  const result = ctx.applySortAndLimit(data, { limit: 2 });
  assert.equal(result.length, 2);
});

test('applySortAndLimit: no sortBy + no limit returns data unchanged', () => {
  const ctx = loadDataServiceContext();
  const data = [{ id: 1 }, { id: 2 }];
  const result = ctx.applySortAndLimit(data, {});
  const ids = Array.from(result, (r) => r.id);
  assert.deepEqual(ids, [1, 2]);
});

test('applySortAndLimit: sortBy=score orders by highlight > reactionSum > timestamp', () => {
  const ctx = loadDataServiceContext();
  const data = [
    { id: 'a', highlight: false, reactions: { LIKE: { count: 1 }, UNDERSTAND: { count: 0 }, CURIOUS: { count: 0 } }, timestamp: '2026-04-10' },
    { id: 'b', highlight: false, reactions: { LIKE: { count: 3 }, UNDERSTAND: { count: 2 }, CURIOUS: { count: 0 } }, timestamp: '2026-04-05' },
    { id: 'c', highlight: true,  reactions: { LIKE: { count: 0 }, UNDERSTAND: { count: 0 }, CURIOUS: { count: 0 } }, timestamp: '2026-04-01' },
    { id: 'd', highlight: false, reactions: { LIKE: { count: 3 }, UNDERSTAND: { count: 2 }, CURIOUS: { count: 0 } }, timestamp: '2026-04-08' }
  ];
  const result = ctx.applySortAndLimit(data, { sortBy: 'score' });
  // highlight first, then reactionSum desc (d > b by newer ts), then last remaining
  const ids = Array.from(result, (r) => r.id);
  assert.deepEqual(ids, ['c', 'd', 'b', 'a']);
});

test('applySortAndLimit: sortBy with limit both apply', () => {
  const ctx = loadDataServiceContext();
  const data = [
    { id: 1, timestamp: '2026-04-01' },
    { id: 2, timestamp: '2026-04-03' },
    { id: 3, timestamp: '2026-04-02' }
  ];
  const result = ctx.applySortAndLimit(data, { sortBy: 'newest', limit: 2 });
  const ids = Array.from(result, (r) => r.id);
  assert.deepEqual(ids, [2, 3]); // newest-first then top 2
});
