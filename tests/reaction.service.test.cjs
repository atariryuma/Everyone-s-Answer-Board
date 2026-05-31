const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { gasResponseStubs } = require('./_helpers.cjs');

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
    },
    // 単一セル書込み (ReactionService は reaction 列を個別 setValue する。
    // span setValues で中間列を stale 上書きしないため)。
    setValue(value) {
      writes.push({ row, col, numRows: 1, numCols: 1, values: [[value]] });
      if (!data[row - 1]) data[row - 1] = [];
      data[row - 1][col - 1] = value;
    }
  });

  return {
    getName: () => name,
    getDataRange: () => makeRange(1, 1, data.length, data[0] ? data[0].length : 0),
    getRange: (row, col, numRows = 1, numCols = 1) => makeRange(row, col, numRows, numCols),
    getLastColumn: () => (data[0] ? data[0].length : 0),
    getParent: () => ({ getId: () => 'mock-ss-id' }),
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
    ...gasResponseStubs(),
    getCurrentEmail: () => 'actor@example.com',
    isAdministrator: (email) => email === 'admin@example.com',
    findUserBySpreadsheetId: () => null,
    findUserById: () => ({ userId: 'owner-1', userEmail: 'owner@example.com' }),
    // findPublishedBoardOwner は DatabaseCore の helper。テストでは findUserById と同等でよい。
    findPublishedBoardOwner: (userId, viewerEmail, extra = {}) =>
      (context.findUserById ? context.findUserById(userId, { ...extra, requestingUser: viewerEmail, allowPublishedRead: true }) : null),
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
  // reaction 3 列を個別 setValue する (span setValues 廃止: 非連続列で中間列を stale 上書きしないため)
  assert.equal(sheet._writes.length, 3);
  assert.equal(sheet._data[1][2], 'actor@example.com'); // LIKE column (0-based index 2)
  assert.equal(sheet._data[1][1], ''); // UNDERSTAND 未変更
  assert.equal(sheet._data[1][3], ''); // CURIOUS 未変更
  assert.equal(sheet._data[1][0], 'answer-a'); // 非 reaction 列 (Q1) は触らない
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

test('processReactionDirect: case-insensitive toggle removes own mixed-case reaction (M1)', () => {
  // Why (M1): stored cell が別 case (Actor@Example.com) でも、 actor@example.com の toggle で
  //   自分の reaction を認識して外せる。 旧実装 (完全一致) では外せず二重カウントしていた。
  const ctx = loadReactionContext();
  const sheet = createMockSheet({
    headers: ['Q1', 'UNDERSTAND', 'LIKE', 'CURIOUS'],
    rows: [['answer-a', '', 'Actor@Example.com', '']]
  });

  const result = ctx.processReactionDirect(sheet, 2, 'LIKE', 'actor@example.com');

  assert.equal(result.action, 'removed', 'mixed-case stored email must be recognized as own reaction');
  assert.equal(result.reactions.LIKE.count, 0);
  assert.equal(result.reactions.LIKE.reacted, false);
  assert.equal(sheet._data[1][2], '', 'own reaction (any case) is removed');
});

test('processReactionDirect: stores canonical (lowercased) email on add (M1)', () => {
  const ctx = loadReactionContext();
  const sheet = createMockSheet({
    headers: ['Q1', 'UNDERSTAND', 'LIKE', 'CURIOUS'],
    rows: [['answer-a', '', '', '']]
  });
  ctx.processReactionDirect(sheet, 2, 'LIKE', 'Actor@Example.com');
  assert.equal(sheet._data[1][2], 'actor@example.com', 'added email is normalized to lowercase');
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

test('processReactionDirect: lazy-provisions missing reaction columns (new behavior)', () => {
  // Why (2026-05-14): 旧来は throw して教師に再接続を要求していたが、
  //   新規シートで初回リアクション必ず失敗の UX 破壊。lazy provisioning に変更。
  const ctx = loadReactionContext();
  const sheet = createMockSheet({
    headers: ['Q1', 'UNDERSTAND', 'CURIOUS'], // no LIKE / HIGHLIGHT
    rows: [['answer-a', '', '']]
  });

  // 不足列が auto-append され、リアクションが成功する
  const result = ctx.processReactionDirect(sheet, 2, 'LIKE', 'actor@example.com');
  assert.equal(result.action, 'added');
  assert.equal(result.reactions.LIKE.count, 1);
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

test('processHighlightDirect: lazy-provisions missing HIGHLIGHT column (new behavior)', () => {
  // Why (2026-05-14): 旧来は throw、新規ボード初回 highlight 必ず失敗の UX 破壊。lazy provisioning に変更。
  const ctx = loadReactionContext();
  const sheet = createMockSheet({
    headers: ['Q1', 'UNDERSTAND'],
    rows: [['answer-a', '']]
  });

  // 不足 HIGHLIGHT 列が auto-append され、トグルが成功する
  const result = ctx.processHighlightDirect(sheet, 2);
  assert.equal(result.highlighted, true);
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

test('extractReactions: reacted is case-insensitive (M1)', () => {
  // Why (M1): identity は case-insensitive。stored cell が別 case でも自分の reaction を認識する。
  const ctx = loadReactionContext();
  const headers = ['Q1', 'UNDERSTAND', 'LIKE', 'CURIOUS'];
  const row = ['answer', 'Alice@Example.com', '', ''];
  const r = ctx.extractReactions(row, headers, 'alice@example.com');
  assert.equal(r.UNDERSTAND.reacted, true, 'mixed-case stored email must match normalized viewer');
});

test('extractReactions: uses precomputed indices when provided (M2)', () => {
  // Why (M2): batch hot path は列 index を 1 度だけ解決して渡す。precomputed を尊重すること。
  const ctx = loadReactionContext();
  // headers をわざと UNDERSTAND を含まない形にし、precomputed の方を使うことを示す。
  const headers = ['Q1', 'colA', 'colB', 'colC'];
  const row = ['answer', 'a@x.com|b@x.com', 'c@x.com', ''];
  const precomputed = { UNDERSTAND: 1, LIKE: 2, CURIOUS: 3 };
  const r = ctx.extractReactions(row, headers, 'c@x.com', precomputed);
  assert.equal(r.UNDERSTAND.count, 2);
  assert.equal(r.LIKE.count, 1);
  assert.equal(r.LIKE.reacted, true);
});

test('resolveReactionColumns_: resolves all 4 columns in one pass (M2)', () => {
  const ctx = loadReactionContext();
  const map = ctx.resolveReactionColumns_(['ts', 'UNDERSTAND', 'q', 'like', 'CURIOUS', 'Highlight']);
  assert.equal(map.UNDERSTAND, 1);
  assert.equal(map.LIKE, 3);      // case-insensitive
  assert.equal(map.CURIOUS, 4);
  assert.equal(map.HIGHLIGHT, 5); // case-insensitive
});

test('resolveReactionColumns_: missing columns are -1', () => {
  const ctx = loadReactionContext();
  const map = ctx.resolveReactionColumns_(['ts', 'answer']);
  assert.equal(map.UNDERSTAND, -1);
  assert.equal(map.HIGHLIGHT, -1);
});

test('extractHighlight: uses precomputed index when provided (M2)', () => {
  const ctx = loadReactionContext();
  // headers に HIGHLIGHT 名が無くても precomputed index で解決。
  assert.equal(ctx.extractHighlight(['a', 'TRUE'], ['Q1', 'colX'], 1), true);
  assert.equal(ctx.extractHighlight(['a', ''], ['Q1', 'colX'], 1), false);
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
// canActOnTargetBoard
// =====================================================================

test('canActOnTargetBoard: admin always has permission', () => {
  const ctx = loadReactionContext({ isAdministrator: (email) => email === 'admin@example.com' });
  const ok = ctx.canActOnTargetBoard(
    'admin@example.com',
    { userEmail: 'other@example.com' },
    { isPublished: false }
  );
  assert.equal(ok, true);
});

test('canActOnTargetBoard: board owner has permission on unpublished', () => {
  const ctx = loadReactionContext();
  const ok = ctx.canActOnTargetBoard(
    'owner@example.com',
    { userEmail: 'owner@example.com' },
    { isPublished: false }
  );
  assert.equal(ok, true);
});

test('canActOnTargetBoard: any user has permission on published (viewer op)', () => {
  const ctx = loadReactionContext();
  const ok = ctx.canActOnTargetBoard(
    'viewer@example.com',
    { userEmail: 'owner@example.com' },
    { isPublished: true }
  );
  assert.equal(ok, true);
});

test('canActOnTargetBoard: non-owner blocked on unpublished', () => {
  const ctx = loadReactionContext();
  const ok = ctx.canActOnTargetBoard(
    'viewer@example.com',
    { userEmail: 'owner@example.com' },
    { isPublished: false }
  );
  assert.equal(ok, false);
});

test('canActOnTargetBoard: requireEditor blocks viewer on published (editor op)', () => {
  const ctx = loadReactionContext();
  const ok = ctx.canActOnTargetBoard(
    'viewer@example.com',
    { userEmail: 'owner@example.com' },
    { isPublished: true },
    { requireEditor: true }
  );
  assert.equal(ok, false, 'Highlight/editor操作は公開ボードの閲覧者でも拒否されるべき');
});

test('canActOnTargetBoard: requireEditor allows owner', () => {
  const ctx = loadReactionContext();
  const ok = ctx.canActOnTargetBoard(
    'owner@example.com',
    { userEmail: 'owner@example.com' },
    { isPublished: true },
    { requireEditor: true }
  );
  assert.equal(ok, true);
});

test('canActOnTargetBoard: collaborator (SS editor) allowed on requireEditor (v2855)', () => {
  // ボード SS の editor として共有された共同教師は highlight/unpublish 等の editor 操作を実行可能。
  const ctx = loadReactionContext({
    isBoardCollaborator: (targetUser, email) => email === 'collab@example.com'
  });
  const ok = ctx.canActOnTargetBoard(
    'collab@example.com',
    { userEmail: 'owner@example.com' },
    { isPublished: true },
    { requireEditor: true }
  );
  assert.equal(ok, true);
});

test('canActOnTargetBoard: non-collaborator viewer still blocked on requireEditor', () => {
  // SS editor として登録されていない viewer は requireEditor を pass しない。
  const ctx = loadReactionContext({
    isBoardCollaborator: () => false
  });
  const ok = ctx.canActOnTargetBoard(
    'viewer@example.com',
    { userEmail: 'owner@example.com' },
    { isPublished: true },
    { requireEditor: true }
  );
  assert.equal(ok, false);
});

test('canActOnTargetBoard: requireEditor allows admin via preloaded isAdmin (perf path)', () => {
  // isAdministrator を呼ばせない（preloaded を信用する）ことを検証
  let calls = 0;
  const ctx = loadReactionContext({
    isAdministrator: () => { calls += 1; return false; }
  });
  const ok = ctx.canActOnTargetBoard(
    'someone@example.com',
    { userEmail: 'owner@example.com' },
    { isPublished: false },
    { requireEditor: true, isAdmin: true }
  );
  assert.equal(ok, true);
  assert.equal(calls, 0, 'options.isAdmin=true の時に isAdministrator を再計算しないこと');
});

test('canActOnTargetBoard: null actor → false', () => {
  const ctx = loadReactionContext();
  assert.equal(
    ctx.canActOnTargetBoard(null, { userEmail: 'x' }, { isPublished: true }),
    false
  );
});

test('canActOnTargetBoard: null targetUser → false', () => {
  const ctx = loadReactionContext();
  assert.equal(
    ctx.canActOnTargetBoard('a@x.com', null, { isPublished: true }),
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
  assert.equal(sheet._writes.length, 3); // reaction 3 列を個別 write
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
  // Why (2026-05-14): 旧版は「列欠落 throw」で error トリガーしていたが、
  //   lazy-provisioning で missing 列は auto-append される。代わりに provisioning 自体が
  //   失敗するケース (setValues throws) を error トリガーに使う。
  const sheet = createMockSheet({
    headers: ['Q1', 'UNDERSTAND', 'CURIOUS'], // missing LIKE → provisioning attempted
    rows: [['answer-a', '', '']]
  });
  // setValues を強制失敗させて processReactionDirect 内の lazy-provisioning を失敗させる
  const originalGetRange = sheet.getRange;
  sheet.getRange = function (row, col, numRows, numCols) {
    const range = originalGetRange.call(sheet, row, col, numRows, numCols);
    if (row === 1 && numRows === 1) {
      // header 書き込み (lazy provisioning) のみ失敗させる
      const origSetValues = range.setValues;
      range.setValues = () => { throw new Error('mocked provisioning failure'); };
      // 復旧後に戻せるように本物の setValues も保持（実害ない）
      range._origSetValues = origSetValues;
    }
    return range;
  };
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
      setValues: () => {},
      setValue: () => {}
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
    // toggleHighlight は editor-only なので、デフォルトでは actor を owner にして
    // 権限チェックを通過させる。権限バグを検証したい個別テストは overrides で
    // getCurrentEmail / findUserById を差し替える。
    getCurrentEmail: () => 'owner@example.com',
    findUserById: () => ({ userId: 'owner-1', userEmail: 'owner@example.com' }),
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

test('toggleHighlight: rejects published-board viewer (editor-only gate)', () => {
  // Why: UI で highlight-btn を出していなくても、google.script.run で直接叩かれる
  //      経路を塞ぐため、サーバー側で editor 判定が必須。以前は config.isPublished
  //      だけで通してしまい生徒がハイライトを自由に操作できるバグがあった。
  const sheet = createMockSheet({
    headers: ['Q1', 'HIGHLIGHT'],
    rows: [['answer-a', 'FALSE']]
  });
  const ctx = buildToggleHighlightContext({
    sheet,
    overrides: {
      getCurrentEmail: () => 'student@example.com',
      isAdministrator: () => false,
      getConfigOrDefault: () => ({ spreadsheetId: 'x', sheetName: 'Sheet1', isPublished: true }),
      findUserById: () => ({ userId: 'owner-1', userEmail: 'teacher@example.com' })
    }
  });
  const result = ctx.toggleHighlight('owner-1', 2);
  assert.equal(result.success, false);
  assert.match(result.message, /Access denied/);
  // シートに書き込みが発生していないこと
  assert.equal(sheet._writes.length, 0, '権限エラー時はシートに書き込まない');
  assert.equal(sheet._data[1][1], 'FALSE', '元の HIGHLIGHT 値が維持される');
});

test('toggleHighlight: allows board owner even on unpublished board', () => {
  const sheet = createMockSheet({
    headers: ['Q1', 'HIGHLIGHT'],
    rows: [['answer-a', 'FALSE']]
  });
  const ctx = buildToggleHighlightContext({
    sheet,
    overrides: {
      getCurrentEmail: () => 'teacher@example.com',
      isAdministrator: () => false,
      getConfigOrDefault: () => ({ spreadsheetId: 'x', sheetName: 'Sheet1', isPublished: false }),
      findUserById: () => ({ userId: 'owner-1', userEmail: 'teacher@example.com' })
    }
  });
  const result = ctx.toggleHighlight('owner-1', 2);
  assert.equal(result.success, true);
  assert.equal(sheet._data[1][1], 'TRUE');
});

test('addReaction: still allows published-board viewer (viewer op is open)', () => {
  // Regression: editor-only 化を highlight だけに適用できていることを保証する。
  // addReaction は requireEditor を渡さないので、公開ボード閲覧者は従来どおり許可。
  const sheet = createMockSheet({
    headers: ['Q1', 'UNDERSTAND', 'LIKE', 'CURIOUS'],
    rows: [['answer-a', '', '', '']]
  });
  const ctx = buildAddReactionContext({
    sheet,
    overrides: {
      getCurrentEmail: () => 'student@example.com',
      isAdministrator: () => false,
      getConfigOrDefault: () => ({ spreadsheetId: 'x', sheetName: 'Sheet1', isPublished: true }),
      findUserById: () => ({ userId: 'owner-1', userEmail: 'teacher@example.com' })
    }
  });
  const result = ctx.addReaction('owner-1', 2, 'LIKE');
  assert.equal(result.success, true);
  assert.equal(result.reactions.LIKE.count, 1);
});

test('toggleHighlight: lazy-provisions missing HIGHLIGHT column (new behavior)', () => {
  // Why (2026-05-14): 旧版は missing 列で error 返却。lazy provisioning に変更したので、
  //   missing でも auto-append して成功する。lock も正しく解放される。
  const sheet = createMockSheet({
    headers: ['Q1', 'UNDERSTAND'],
    rows: [['answer-a', '']]
  });
  const lock = createMockLock();
  const ctx = buildToggleHighlightContext({ sheet, lock });

  const result = ctx.toggleHighlight('owner-1', 2);

  assert.equal(result.success, true);
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
