const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function createMockSheet({ headers = [], rows = [], name = 'Sheet1' } = {}) {
  const data = [headers.slice(), ...rows.map((r) => r.slice())];
  const writes = [];

  const makeRange = (row, col, numRows, numCols) => ({
    getValues() {
      const result = [];
      for (let r = row - 1; r < row - 1 + numRows; r += 1) {
        const rowArr = [];
        for (let c = col - 1; c < col - 1 + numCols; c += 1) {
          rowArr.push(data[r] ? (data[r][c] !== undefined ? data[r][c] : '') : '');
        }
        result.push(rowArr);
      }
      return result;
    },
    setValues(values) {
      writes.push({ row, col, numRows, numCols, values: values.map((r) => r.slice()) });
      for (let r = 0; r < values.length; r += 1) {
        if (!data[row - 1 + r]) data[row - 1 + r] = [];
        for (let c = 0; c < values[r].length; c += 1) {
          data[row - 1 + r][col - 1 + c] = values[r][c];
        }
      }
    }
  });

  return {
    getName: () => name,
    getDataRange: () => makeRange(1, 1, data.length, data[0] ? data[0].length : 0),
    getRange: (row, col, numRows = 1, numCols = 1) => makeRange(row, col, numRows, numCols),
    _data: data,
    _writes: writes
  };
}

function createMockCache(initial = {}) {
  const store = new Map(Object.entries(initial));
  return {
    get: (k) => (store.has(k) ? store.get(k) : null),
    put: (k, v) => { store.set(k, v); },
    remove: (k) => { store.delete(k); },
    _store: store
  };
}

function createMockLock({ tryLockResult = true, throwOnRelease = false } = {}) {
  let held = false;
  return {
    tryLock: () => {
      if (tryLockResult) { held = true; return true; }
      return false;
    },
    releaseLock: () => {
      if (throwOnRelease) throw new Error('lock release failed');
      held = false;
    },
    isHeld: () => held
  };
}

function loadReactionContext(overrides = {}) {
  const cache = overrides.cache || createMockCache();
  const lock = overrides.lock || createMockLock();

  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    CacheService: { getScriptCache: () => cache },
    LockService: { getScriptLock: () => lock },
    CACHE_DURATION: { MEDIUM: 30 },
    SYSTEM_LIMITS: {},
    createErrorResponse: (message, data, extra) => ({ success: false, message, ...extra }),
    createExceptionResponse: (error) => ({ success: false, message: error.message, error: error.message }),
    getCurrentEmail: () => 'actor@example.com',
    isAdministrator: (email) => email === 'admin@example.com',
    findUserBySpreadsheetId: () => null,
    findUserById: () => ({ userId: 'owner-1', userEmail: 'owner@example.com' }),
    getUserConfig: () => ({ success: true, config: {} }),
    getConfigOrDefault: () => ({
      spreadsheetId: 'sheet-123',
      sheetName: 'Sheet1',
      isPublished: true
    }),
    openSpreadsheet: () => ({ spreadsheet: { getSheetByName: () => null } }),
    ...overrides
  };

  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/ReactionService.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'ReactionService.js' });
  context._cache = cache;
  context._lock = lock;
  return context;
}

// =====================================================================
// processReactionDirect — core logic (pure, no external calls)
// =====================================================================

test('processReactionDirect: adds LIKE when user has no prior reaction', () => {
  const ctx = loadReactionContext();
  const sheet = createMockSheet({
    headers: ['Q1', 'UNDERSTAND', 'LIKE', 'CURIOUS', 'HIGHLIGHT'],
    rows: [['answer-a', '', '', '', '']]
  });

  const result = ctx.processReactionDirect(sheet, 2, 'LIKE', 'actor@example.com');

  assert.equal(result.action, 'added');
  assert.equal(result.userReaction, 'LIKE');
  assert.equal(result.reactions.LIKE.count, 1);
  assert.equal(result.reactions.LIKE.reacted, true);
  assert.equal(result.reactions.UNDERSTAND.count, 0);
  assert.equal(sheet._writes.length, 1);
  assert.equal(sheet._writes[0].values[0][1], 'actor@example.com'); // LIKE column (offset from minCol=2)
});

test('processReactionDirect: removes LIKE when user toggles same reaction', () => {
  const ctx = loadReactionContext();
  const sheet = createMockSheet({
    headers: ['Q1', 'UNDERSTAND', 'LIKE', 'CURIOUS'],
    rows: [['answer-a', '', 'actor@example.com', '']]
  });

  const result = ctx.processReactionDirect(sheet, 2, 'LIKE', 'actor@example.com');

  assert.equal(result.action, 'removed');
  assert.equal(result.userReaction, null);
  assert.equal(result.reactions.LIKE.count, 0);
  assert.equal(result.reactions.LIKE.reacted, false);
});

test('processReactionDirect: switches reaction from LIKE to CURIOUS', () => {
  const ctx = loadReactionContext();
  const sheet = createMockSheet({
    headers: ['Q1', 'UNDERSTAND', 'LIKE', 'CURIOUS'],
    rows: [['answer-a', '', 'actor@example.com', '']]
  });

  const result = ctx.processReactionDirect(sheet, 2, 'CURIOUS', 'actor@example.com');

  assert.equal(result.action, 'added');
  assert.equal(result.userReaction, 'CURIOUS');
  assert.equal(result.reactions.LIKE.count, 0);
  assert.equal(result.reactions.LIKE.reacted, false);
  assert.equal(result.reactions.CURIOUS.count, 1);
  assert.equal(result.reactions.CURIOUS.reacted, true);
});

test('processReactionDirect: preserves other users reactions', () => {
  const ctx = loadReactionContext();
  const sheet = createMockSheet({
    headers: ['Q1', 'UNDERSTAND', 'LIKE', 'CURIOUS'],
    rows: [['answer-a', '', 'other@example.com|third@example.com', '']]
  });

  const result = ctx.processReactionDirect(sheet, 2, 'LIKE', 'actor@example.com');

  assert.equal(result.reactions.LIKE.count, 3);
  assert.equal(result.reactions.LIKE.reacted, true);
  // Written value must still contain other users
  const writtenLike = sheet._data[1][2];
  assert.ok(writtenLike.includes('other@example.com'));
  assert.ok(writtenLike.includes('third@example.com'));
  assert.ok(writtenLike.includes('actor@example.com'));
});

test('processReactionDirect: rejects invalid reaction type', () => {
  const ctx = loadReactionContext();
  const sheet = createMockSheet({
    headers: ['Q1', 'UNDERSTAND', 'LIKE', 'CURIOUS'],
    rows: [['answer-a', '', '', '']]
  });

  assert.throws(
    () => ctx.processReactionDirect(sheet, 2, 'INVALID', 'actor@example.com'),
    /Invalid reaction type/
  );
});

test('processReactionDirect: graceful degradation when headers unavailable', () => {
  const ctx = loadReactionContext();
  const sheet = {
    getDataRange: () => ({ getValues: () => [[]] }),
    getRange: () => ({ getValues: () => [[]], setValues: () => {} })
  };

  const result = ctx.processReactionDirect(sheet, 2, 'LIKE', 'actor@example.com');

  assert.equal(result.action, 'unavailable');
  assert.equal(result.userReaction, null);
  assert.equal(result.reactions.LIKE.count, 0);
  assert.match(result.message, /一時的/);
});

test('processReactionDirect: throws when reaction column is missing', () => {
  const ctx = loadReactionContext();
  const sheet = createMockSheet({
    headers: ['Q1', 'UNDERSTAND', 'CURIOUS'], // no LIKE
    rows: [['answer-a', '', '']]
  });

  assert.throws(
    () => ctx.processReactionDirect(sheet, 2, 'LIKE', 'actor@example.com'),
    /LIKE/
  );
});

test('processReactionDirect: handles header case and whitespace', () => {
  const ctx = loadReactionContext();
  const sheet = createMockSheet({
    headers: ['Q1', '  understand  ', ' Like ', 'CURIOUS'],
    rows: [['answer-a', '', '', '']]
  });

  const result = ctx.processReactionDirect(sheet, 2, 'LIKE', 'actor@example.com');

  assert.equal(result.action, 'added');
  assert.equal(result.reactions.LIKE.count, 1);
});

test('processReactionDirect: filters empty entries from pipe-serialized users', () => {
  const ctx = loadReactionContext();
  const sheet = createMockSheet({
    headers: ['Q1', 'UNDERSTAND', 'LIKE', 'CURIOUS'],
    rows: [['answer-a', '', '|actor@example.com||other@example.com|', '']]
  });

  const result = ctx.processReactionDirect(sheet, 2, 'LIKE', 'actor@example.com');

  // actor is toggling off. Should be removed, other preserved.
  assert.equal(result.action, 'removed');
  assert.equal(result.reactions.LIKE.count, 1);
  assert.equal(sheet._data[1][2], 'other@example.com');
});

// =====================================================================
// processHighlightDirect
// =====================================================================

test('processHighlightDirect: toggles FALSE → TRUE', () => {
  const ctx = loadReactionContext();
  const sheet = createMockSheet({
    headers: ['Q1', 'UNDERSTAND', 'LIKE', 'CURIOUS', 'HIGHLIGHT'],
    rows: [['answer-a', '', '', '', 'FALSE']]
  });

  const result = ctx.processHighlightDirect(sheet, 2);

  assert.equal(result.highlighted, true);
  assert.equal(sheet._data[1][4], 'TRUE');
});

test('processHighlightDirect: toggles TRUE → FALSE', () => {
  const ctx = loadReactionContext();
  const sheet = createMockSheet({
    headers: ['Q1', 'HIGHLIGHT'],
    rows: [['answer-a', 'TRUE']]
  });

  const result = ctx.processHighlightDirect(sheet, 2);

  assert.equal(result.highlighted, false);
  assert.equal(sheet._data[1][1], 'FALSE');
});

test('processHighlightDirect: empty cell treated as FALSE, toggles to TRUE', () => {
  const ctx = loadReactionContext();
  const sheet = createMockSheet({
    headers: ['Q1', 'HIGHLIGHT'],
    rows: [['answer-a', '']]
  });

  const result = ctx.processHighlightDirect(sheet, 2);

  assert.equal(result.highlighted, true);
  assert.equal(sheet._data[1][1], 'TRUE');
});

test('processHighlightDirect: graceful degradation on empty headers', () => {
  const ctx = loadReactionContext();
  const sheet = {
    getDataRange: () => ({ getValues: () => [[]] }),
    getRange: () => ({ getValues: () => [[]], setValues: () => {} })
  };

  const result = ctx.processHighlightDirect(sheet, 2);

  assert.equal(result.highlighted, false);
  assert.match(result.message, /一時的/);
});

test('processHighlightDirect: throws when HIGHLIGHT column missing', () => {
  const ctx = loadReactionContext();
  const sheet = createMockSheet({
    headers: ['Q1', 'UNDERSTAND'],
    rows: [['answer-a', '']]
  });

  assert.throws(
    () => ctx.processHighlightDirect(sheet, 2),
    /HIGHLIGHT/
  );
});

// =====================================================================
// extractReactions
// =====================================================================

test('extractReactions: counts users in pipe-separated cells', () => {
  const ctx = loadReactionContext();
  const headers = ['Q1', 'UNDERSTAND', 'LIKE', 'CURIOUS'];
  const row = ['answer', 'a@x.com|b@x.com', 'c@x.com', ''];

  const r = ctx.extractReactions(row, headers, 'c@x.com');

  assert.equal(r.UNDERSTAND.count, 2);
  assert.equal(r.UNDERSTAND.reacted, false);
  assert.equal(r.LIKE.count, 1);
  assert.equal(r.LIKE.reacted, true);
  assert.equal(r.CURIOUS.count, 0);
});

test('extractReactions: returns zero counts when columns missing', () => {
  const ctx = loadReactionContext();
  const r = ctx.extractReactions(['answer'], ['Q1'], 'x@x.com');

  assert.equal(r.UNDERSTAND.count, 0);
  assert.equal(r.LIKE.count, 0);
  assert.equal(r.CURIOUS.count, 0);
});

test('extractReactions: reacted=false when userEmail is null', () => {
  const ctx = loadReactionContext();
  const r = ctx.extractReactions(['a', 'x@x.com', '', ''], ['Q1', 'UNDERSTAND', 'LIKE', 'CURIOUS']);
  assert.equal(r.UNDERSTAND.reacted, false);
  assert.equal(r.UNDERSTAND.count, 1);
});

// =====================================================================
// extractHighlight
// =====================================================================

test('extractHighlight: TRUE string → true', () => {
  const ctx = loadReactionContext();
  assert.equal(ctx.extractHighlight(['a', 'TRUE'], ['Q1', 'HIGHLIGHT']), true);
});

test('extractHighlight: accepts 1 and YES as truthy', () => {
  const ctx = loadReactionContext();
  assert.equal(ctx.extractHighlight(['a', '1'], ['Q1', 'HIGHLIGHT']), true);
  assert.equal(ctx.extractHighlight(['a', 'yes'], ['Q1', 'HIGHLIGHT']), true);
});

test('extractHighlight: empty or FALSE → false', () => {
  const ctx = loadReactionContext();
  assert.equal(ctx.extractHighlight(['a', ''], ['Q1', 'HIGHLIGHT']), false);
  assert.equal(ctx.extractHighlight(['a', 'FALSE'], ['Q1', 'HIGHLIGHT']), false);
});

test('extractHighlight: returns false when column missing', () => {
  const ctx = loadReactionContext();
  assert.equal(ctx.extractHighlight(['a'], ['Q1']), false);
});

// =====================================================================
// validateReactionPermissionWithPreloadedData
// =====================================================================

test('validateReactionPermissionWithPreloadedData: admin always has permission', () => {
  const ctx = loadReactionContext({ isAdministrator: (email) => email === 'admin@example.com' });
  const ok = ctx.validateReactionPermissionWithPreloadedData(
    'admin@example.com',
    { userEmail: 'other@example.com' },
    { isPublished: false }
  );
  assert.equal(ok, true);
});

test('validateReactionPermissionWithPreloadedData: board owner has permission on unpublished', () => {
  const ctx = loadReactionContext();
  const ok = ctx.validateReactionPermissionWithPreloadedData(
    'owner@example.com',
    { userEmail: 'owner@example.com' },
    { isPublished: false }
  );
  assert.equal(ok, true);
});

test('validateReactionPermissionWithPreloadedData: any user has permission on published', () => {
  const ctx = loadReactionContext();
  const ok = ctx.validateReactionPermissionWithPreloadedData(
    'viewer@example.com',
    { userEmail: 'owner@example.com' },
    { isPublished: true }
  );
  assert.equal(ok, true);
});

test('validateReactionPermissionWithPreloadedData: non-owner blocked on unpublished', () => {
  const ctx = loadReactionContext();
  const ok = ctx.validateReactionPermissionWithPreloadedData(
    'viewer@example.com',
    { userEmail: 'owner@example.com' },
    { isPublished: false }
  );
  assert.equal(ok, false);
});

test('validateReactionPermissionWithPreloadedData: null actor → false', () => {
  const ctx = loadReactionContext();
  assert.equal(
    ctx.validateReactionPermissionWithPreloadedData(null, { userEmail: 'x' }, { isPublished: true }),
    false
  );
});

test('validateReactionPermissionWithPreloadedData: null targetUser → false', () => {
  const ctx = loadReactionContext();
  assert.equal(
    ctx.validateReactionPermissionWithPreloadedData('a@x.com', null, { isPublished: true }),
    false
  );
});

// =====================================================================
// addReaction — integration
// =====================================================================

function buildAddReactionContext({ sheet, cache, lock, overrides = {} } = {}) {
  const actualSheet = sheet || createMockSheet({
    headers: ['Q1', 'UNDERSTAND', 'LIKE', 'CURIOUS'],
    rows: [['answer-a', '', '', '']]
  });
  return loadReactionContext({
    cache: cache || createMockCache(),
    lock: lock || createMockLock(),
    openSpreadsheet: () => ({
      spreadsheet: { getSheetByName: (name) => (name === 'Sheet1' ? actualSheet : null) }
    }),
    ...overrides,
    _sheet: actualSheet
  });
}

test('addReaction: happy path returns success and writes to sheet', () => {
  const sheet = createMockSheet({
    headers: ['Q1', 'UNDERSTAND', 'LIKE', 'CURIOUS'],
    rows: [['answer-a', '', '', '']]
  });
  const ctx = buildAddReactionContext({ sheet });

  const result = ctx.addReaction('owner-1', 2, 'LIKE');

  assert.equal(result.success, true);
  assert.equal(result.action, 'added');
  assert.equal(result.reactions.LIKE.count, 1);
  assert.equal(sheet._writes.length, 1);
});

test('addReaction: releases lock and removes cache key on success', () => {
  const cache = createMockCache();
  const lock = createMockLock();
  const ctx = buildAddReactionContext({ cache, lock });

  ctx.addReaction('owner-1', 2, 'LIKE');

  assert.equal(lock.isHeld(), false);
  assert.equal(cache._store.size, 0);
});

test('addReaction: releases lock on processing error', () => {
  const sheet = createMockSheet({
    headers: ['Q1', 'UNDERSTAND', 'CURIOUS'], // missing LIKE → throws in processReactionDirect
    rows: [['answer-a', '', '']]
  });
  const lock = createMockLock();
  const cache = createMockCache();
  const ctx = buildAddReactionContext({ sheet, cache, lock });

  const result = ctx.addReaction('owner-1', 2, 'LIKE');

  assert.equal(result.success, false);
  assert.equal(lock.isHeld(), false);
  assert.equal(cache._store.size, 0);
});

test('addReaction: rejects invalid row id', () => {
  const ctx = buildAddReactionContext();
  assert.equal(ctx.addReaction('owner-1', 1, 'LIKE').success, false); // row 1 = header
  assert.equal(ctx.addReaction('owner-1', 0, 'LIKE').success, false);
  assert.equal(ctx.addReaction('owner-1', 'not-a-number', 'LIKE').success, false);
});

test('addReaction: parses "row_N" format', () => {
  const sheet = createMockSheet({
    headers: ['Q1', 'UNDERSTAND', 'LIKE', 'CURIOUS'],
    rows: [['a', '', '', ''], ['b', '', '', '']]
  });
  const ctx = buildAddReactionContext({ sheet });

  const result = ctx.addReaction('owner-1', 'row_3', 'LIKE');

  assert.equal(result.success, true);
  // Write should target row 3 in sheet
  assert.equal(sheet._writes[0].row, 3);
});

test('addReaction: rejects when target user not found', () => {
  const ctx = buildAddReactionContext({
    overrides: { findUserById: () => null }
  });
  const result = ctx.addReaction('ghost', 2, 'LIKE');
  assert.equal(result.success, false);
  assert.match(result.message, /Target user/);
});

test('addReaction: rejects when config missing spreadsheet', () => {
  const ctx = buildAddReactionContext({
    overrides: { getConfigOrDefault: () => ({ spreadsheetId: '', sheetName: '' }) }
  });
  const result = ctx.addReaction('owner-1', 2, 'LIKE');
  assert.equal(result.success, false);
  assert.match(result.message, /configuration/);
});

test('addReaction: denies non-owner on unpublished board', () => {
  const ctx = buildAddReactionContext({
    overrides: {
      getCurrentEmail: () => 'stranger@example.com',
      getConfigOrDefault: () => ({ spreadsheetId: 'x', sheetName: 'Sheet1', isPublished: false }),
      findUserById: () => ({ userId: 'owner-1', userEmail: 'owner@example.com' })
    }
  });
  const result = ctx.addReaction('owner-1', 2, 'LIKE');
  assert.equal(result.success, false);
  assert.match(result.message, /Access denied/);
});

test('addReaction: short-circuits when cache lock already held', () => {
  const cache = createMockCache({
    'reaction_sheet-123_2': 'other@example.com'
  });
  const ctx = buildAddReactionContext({ cache });

  const result = ctx.addReaction('owner-1', 2, 'LIKE');

  assert.equal(result.success, false);
  assert.match(result.message, /同時/);
  // Cache key untouched (belongs to the other holder)
  assert.equal(cache._store.get('reaction_sheet-123_2'), 'other@example.com');
});

test('addReaction: returns error when LockService tryLock fails', () => {
  const lock = createMockLock({ tryLockResult: false });
  const ctx = buildAddReactionContext({ lock });

  const result = ctx.addReaction('owner-1', 2, 'LIKE');

  assert.equal(result.success, false);
  assert.match(result.message, /同時/);
});

test('addReaction: returns error when openSpreadsheet fails', () => {
  const ctx = loadReactionContext({ openSpreadsheet: () => null });
  const result = ctx.addReaction('owner-1', 2, 'LIKE');
  assert.equal(result.success, false);
});

test('addReaction: uses cached headers from getSheetHeaders and avoids per-call getDataRange', () => {
  // Concrete headers from the "cache" — reactionType columns match these.
  const cachedHeaders = ['Q1', 'UNDERSTAND', 'LIKE', 'CURIOUS', 'HIGHLIGHT'];

  // Sheet returns EMPTY headers via getDataRange to prove we're NOT using that path.
  // If the new cached-header optimization works, this test still succeeds because
  // processReactionDirect falls back to the preloadedHeaders.
  let dataRangeCalls = 0;
  const sheet = {
    getName: () => 'Sheet1',
    getParent: () => ({ getId: () => 'sheet-123' }),
    getDataRange: () => { dataRangeCalls += 1; return { getValues: () => [[]] }; },
    getRange: (row, col, numRows = 1, numCols = 1) => ({
      getValues: () => {
        const out = [];
        for (let r = 0; r < numRows; r += 1) {
          const rowArr = [];
          for (let c = 0; c < numCols; c += 1) rowArr.push('');
          out.push(rowArr);
        }
        return out;
      },
      setValues: () => {}
    })
  };

  const ctx = loadReactionContext({
    getSheetHeaders: () => ({ headers: cachedHeaders, lastCol: cachedHeaders.length }),
    openSpreadsheet: () => ({
      spreadsheet: { getSheetByName: () => sheet }
    })
  });

  const result = ctx.addReaction('owner-1', 2, 'LIKE');

  assert.equal(result.success, true, 'Should succeed using cached headers');
  assert.equal(dataRangeCalls, 0, 'processReactionDirect must NOT call sheet.getDataRange when preloadedHeaders provided');
});

// =====================================================================
// toggleHighlight — integration
// =====================================================================

function buildToggleHighlightContext({ sheet, cache, lock, overrides = {} } = {}) {
  const actualSheet = sheet || createMockSheet({
    headers: ['Q1', 'HIGHLIGHT'],
    rows: [['answer-a', 'FALSE']]
  });
  return loadReactionContext({
    cache: cache || createMockCache(),
    lock: lock || createMockLock(),
    openSpreadsheet: () => ({
      spreadsheet: { getSheetByName: (name) => (name === 'Sheet1' ? actualSheet : null) }
    }),
    ...overrides,
    _sheet: actualSheet
  });
}

test('toggleHighlight: happy path flips FALSE → TRUE', () => {
  const sheet = createMockSheet({
    headers: ['Q1', 'HIGHLIGHT'],
    rows: [['answer-a', 'FALSE']]
  });
  const ctx = buildToggleHighlightContext({ sheet });

  const result = ctx.toggleHighlight('owner-1', 2);

  assert.equal(result.success, true);
  assert.equal(result.highlighted, true);
  assert.equal(sheet._data[1][1], 'TRUE');
});

test('toggleHighlight: releases lock on success', () => {
  const lock = createMockLock();
  const cache = createMockCache();
  const ctx = buildToggleHighlightContext({ cache, lock });

  ctx.toggleHighlight('owner-1', 2);

  assert.equal(lock.isHeld(), false);
  assert.equal(cache._store.size, 0);
});

test('toggleHighlight: rejects unauthorized actor', () => {
  const ctx = buildToggleHighlightContext({
    overrides: {
      getCurrentEmail: () => 'stranger@example.com',
      getConfigOrDefault: () => ({ spreadsheetId: 'x', sheetName: 'Sheet1', isPublished: false }),
      findUserById: () => ({ userId: 'owner-1', userEmail: 'owner@example.com' })
    }
  });
  const result = ctx.toggleHighlight('owner-1', 2);
  assert.equal(result.success, false);
  assert.match(result.message, /Access denied/);
});

test('toggleHighlight: returns error on missing HIGHLIGHT column', () => {
  const sheet = createMockSheet({
    headers: ['Q1', 'UNDERSTAND'],
    rows: [['answer-a', '']]
  });
  const lock = createMockLock();
  const ctx = buildToggleHighlightContext({ sheet, lock });

  const result = ctx.toggleHighlight('owner-1', 2);

  assert.equal(result.success, false);
  assert.equal(lock.isHeld(), false);
});

test('toggleHighlight: short-circuits when cache lock already held', () => {
  const cache = createMockCache({
    'highlight_sheet-123_2': 'other@example.com'
  });
  const ctx = buildToggleHighlightContext({ cache });

  const result = ctx.toggleHighlight('owner-1', 2);

  assert.equal(result.success, false);
  assert.match(result.message, /同時/);
});
