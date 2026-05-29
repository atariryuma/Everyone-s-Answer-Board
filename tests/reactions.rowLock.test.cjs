// v2785+: per-row CAS lock (ReactionService.js acquireRowLock_) のテスト。
// 旧 設計 (LockService.getScriptLock を process() 全体で保持) は全 board が serialize する
// bottleneck だった。 新設計は ScriptLock を ~5ms の critical section だけで使い、
// row 単位での 並列性を確保する。

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function loadReactionCtx(overrides = {}) {
  const cacheStore = new Map();
  const cache = {
    get: (k) => cacheStore.has(k) ? cacheStore.get(k) : null,
    put: (k, v) => { cacheStore.set(k, v); },
    remove: (k) => { cacheStore.delete(k); }
  };
  let lockHolder = null;
  const lock = {
    tryLock: (timeoutMs) => {
      if (lockHolder) return false;
      lockHolder = 'held';
      return true;
    },
    releaseLock: () => { lockHolder = null; }
  };

  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    CacheService: { getScriptCache: () => cache },
    LockService: { getScriptLock: () => lock },
    Utilities: { sleep: () => {} },
    getCurrentEmail: () => 'actor@example.com',
    findPublishedBoardOwner: () => null,
    getConfigOrDefault: () => ({}),
    openSpreadsheet: () => null,
    createErrorResponse: (msg) => ({ success: false, error: msg, message: msg }),
    createExceptionResponse: (err) => ({ success: false, error: err && err.message }),
    isAdministrator: () => false,
    invalidateSheetHeadersCache: () => {},
    bumpBoardDataVersion_: () => {},
    sameEmail_: (a, b) => String(a || '').toLowerCase().trim() === String(b || '').toLowerCase().trim(),
    ...overrides
  };
  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/ReactionService.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'ReactionService.js' });
  context._cache = cache;
  context._cacheStore = cacheStore;
  return context;
}

test('acquireRowLock_: returns true on first acquisition', () => {
  const ctx = loadReactionCtx();
  const ok = ctx.acquireRowLock_(ctx._cache, 'reaction_ss-1_5', 'a@x.com');
  assert.equal(ok, true);
  assert.equal(ctx._cache.get('reaction_ss-1_5'), 'a@x.com');
});

test('acquireRowLock_: rejects when same row is already locked (fast path)', () => {
  const ctx = loadReactionCtx();
  ctx._cache.put('reaction_ss-1_5', 'a@x.com', 10);
  const ok = ctx.acquireRowLock_(ctx._cache, 'reaction_ss-1_5', 'b@x.com');
  assert.equal(ok, false);
});

test('acquireRowLock_: different rows do NOT block each other', () => {
  const ctx = loadReactionCtx();
  const ok1 = ctx.acquireRowLock_(ctx._cache, 'reaction_ss-1_5', 'a@x.com');
  const ok2 = ctx.acquireRowLock_(ctx._cache, 'reaction_ss-1_6', 'b@x.com');
  const ok3 = ctx.acquireRowLock_(ctx._cache, 'reaction_ss-2_5', 'c@x.com');
  assert.equal(ok1, true);
  assert.equal(ok2, true);
  assert.equal(ok3, true);
});

test('acquireRowLock_: ScriptLock fail returns false (caller surfaces concurrent error)', () => {
  let lockHolder = 'someone else';
  const lock = {
    tryLock: () => !lockHolder,
    releaseLock: () => { lockHolder = null; }
  };
  const ctx = loadReactionCtx({
    LockService: { getScriptLock: () => lock }
  });
  const ok = ctx.acquireRowLock_(ctx._cache, 'reaction_ss-1_5', 'a@x.com');
  assert.equal(ok, false, 'when ScriptLock is held by another tx, acquisition fails');
});

test('acquireRowLock_: ScriptLock critical section re-checks cache (race protection)', () => {
  // ScriptLock 取得後、 別 thread が cache を埋めた状態をシミュレート。
  // critical section 内の 2nd cache.get で false になる。
  let calls = 0;
  const cacheStore = new Map();
  const cache = {
    get: (k) => {
      calls++;
      // 1st call (fast path) → 空、 2nd call (critical section) → 既に埋まっている
      if (calls === 1) return null;
      if (calls === 2) return 'other-actor';
      return cacheStore.has(k) ? cacheStore.get(k) : null;
    },
    put: (k, v) => { cacheStore.set(k, v); },
    remove: (k) => { cacheStore.delete(k); }
  };
  const ctx = loadReactionCtx({
    CacheService: { getScriptCache: () => cache }
  });
  const ok = ctx.acquireRowLock_(cache, 'reaction_ss-1_5', 'a@x.com');
  assert.equal(ok, false, 'race detected in critical section → reject');
});

test('acquireRowLock_: releases ScriptLock even if cache.put throws', () => {
  let releaseCount = 0;
  const lock = {
    tryLock: () => true,
    releaseLock: () => { releaseCount++; }
  };
  const cache = {
    get: () => null,
    put: () => { throw new Error('cache full'); },
    remove: () => {}
  };
  const ctx = loadReactionCtx({
    LockService: { getScriptLock: () => lock },
    CacheService: { getScriptCache: () => cache }
  });
  assert.throws(() => ctx.acquireRowLock_(cache, 'reaction_ss-1_5', 'a@x.com'), /cache full/);
  assert.equal(releaseCount, 1, 'ScriptLock must always release');
});
