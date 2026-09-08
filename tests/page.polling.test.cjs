const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

// Extract the StudyQuestApp class source from page.js.html and instantiate
// it in a minimal vm context with DOM/browser stubs. This lets us exercise
// the pure-logic methods (getPollingInterval, debounceReactionByRow) that
// contain recently-fixed regressions we want to lock in.
function loadStudyQuestAppClass() {
  const src = fs.readFileSync(path.resolve(__dirname, '../src/page.js.html'), 'utf8');

  // Extract the first <script>...</script> block containing the class.
  const scriptMatch = src.match(/<script[^>]*>([\s\S]*?)<\/script>/);
  if (!scriptMatch) throw new Error('<script> tag not found in page.js.html');
  const js = scriptMatch[1];
  // The script's bottom-of-file DOMContentLoaded + beforeunload registrations
  // are harmless in a vm context: document.addEventListener is a no-op stub
  // and the listeners never fire. We leave the whole block intact to keep
  // the extracted source syntactically complete.

  // Minimal DOM/browser stubs. Class methods we test don't read the DOM,
  // but the class body references several globals at parse time so they
  // need to exist.
  const timers = new Map();
  let nextTimerId = 1;
  const ctx = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    Map, Set, Date, Math, JSON, Number, Array, Object, Promise, Error,
    Symbol, WeakMap, WeakSet, Reflect, Proxy,
    Buffer,
    setTimeout: (fn, delay) => {
      const id = nextTimerId++;
      timers.set(id, { fn, delay, createdAt: Date.now() });
      return id;
    },
    clearTimeout: (id) => { timers.delete(id); },
    setInterval: () => 0,
    clearInterval: () => {},
    requestAnimationFrame: () => 0,
    cancelAnimationFrame: () => {},
    requestIdleCallback: (fn) => { if (fn) fn({ timeRemaining: () => 50 }); return 0; },
    cancelIdleCallback: () => {}
  };
  ctx.document = {
    createElement: () => ({
      classList: { add: () => {}, remove: () => {}, contains: () => false },
      setAttribute: () => {},
      querySelector: () => null,
      querySelectorAll: () => [],
      addEventListener: () => {},
      removeEventListener: () => {},
      appendChild: () => {},
      remove: () => {},
      dataset: {},
      style: {},
      textContent: '',
      innerHTML: ''
    }),
    createDocumentFragment: () => ({ appendChild: () => {} }),
    getElementById: () => null,
    querySelector: () => null,
    querySelectorAll: () => [],
    addEventListener: () => {},
    removeEventListener: () => {},
    contains: () => true,
    body: { classList: { add: () => {}, remove: () => {} } }
  };
  ctx.window = {
    addEventListener: () => {},
    removeEventListener: () => {},
    location: { reload: () => {}, href: '', search: '' },
    sharedUtilities: { security: { escapeHtml: (s) => String(s || '') } },
    UNIFIED_CONFIG: {},
    studyQuestApp: undefined,
    notifications: null
  };
  ctx.navigator = { userAgent: 'test', hardwareConcurrency: 4 };
  ctx.localStorage = { getItem: () => null, setItem: () => {}, removeItem: () => {}, clear: () => {} };
  ctx.sessionStorage = { getItem: () => null, setItem: () => {}, removeItem: () => {}, clear: () => {} };
  ctx.google = { script: { run: {} } };
  ctx.globalThis = ctx;
  ctx.self = ctx;

  vm.createContext(ctx);
  vm.runInContext(js, ctx, { filename: 'page.js.html' });

  // StudyQuestApp is defined as a top-level class in the script. vm evaluates
  // class declarations as local bindings of the script, so re-expose it.
  const classBinding = vm.runInContext('(typeof StudyQuestApp !== "undefined") ? StudyQuestApp : null', ctx);
  if (!classBinding) throw new Error('StudyQuestApp was not defined after evaluating page.js.html');
  return { StudyQuestApp: classBinding, ctx, timers };
}

// Create a lightweight instance bypassing the full constructor side effects
// (which would try to initialize polling, find DOM elements, etc.).
function makeInstance(overrides = {}) {
  const { StudyQuestApp, ctx, timers } = loadStudyQuestAppClass();
  // Bypass the constructor by creating a plain object with StudyQuestApp's
  // prototype, then populating only the fields the methods under test need.
  const instance = Object.create(StudyQuestApp.prototype);
  instance.polling = { isActive: false, errorCount: 0, timerId: null, resumeTimerId: null };
  instance.reactionDebounceTimeouts = new Map();
  instance.pendingReactions = new Map();
  instance.lastActivityTime = Date.now();
  instance.state = { userId: 'u1' };
  instance.elements = { answersContainer: null };
  Object.assign(instance, overrides);
  return { instance, ctx, timers };
}

// =====================================================================
// getPollingInterval — activity-based + error backoff
// =====================================================================

test('__basePollingInterval: <1min of activity → 5s (授業中)', () => {
  const { instance } = makeInstance({ lastActivityTime: Date.now() });
  assert.equal(instance.__basePollingInterval(), 5000);
});

test('__basePollingInterval: 1-5min → 15s', () => {
  const { instance } = makeInstance({ lastActivityTime: Date.now() - 120000 });
  assert.equal(instance.__basePollingInterval(), 15000);
});

test('__basePollingInterval: 5-15min → 1min', () => {
  const { instance } = makeInstance({ lastActivityTime: Date.now() - 600000 });
  assert.equal(instance.__basePollingInterval(), 60000);
});

test('__basePollingInterval: >15min → 5min', () => {
  const { instance } = makeInstance({ lastActivityTime: Date.now() - 3600000 });
  assert.equal(instance.__basePollingInterval(), 300000);
});

test('__basePollingInterval: errorCount=1 → 30s (exponential backoff kicks in)', () => {
  const { instance } = makeInstance();
  instance.polling.errorCount = 1;
  // Even if user is currently active, error backoff takes precedence
  assert.equal(instance.__basePollingInterval(), 30000);
});

test('__basePollingInterval: errorCount=2 → 60s (exponential doubles)', () => {
  const { instance } = makeInstance();
  instance.polling.errorCount = 2;
  assert.equal(instance.__basePollingInterval(), 60000);
});

test('__basePollingInterval: errorCount=3 → 120s (capped)', () => {
  const { instance } = makeInstance();
  instance.polling.errorCount = 3;
  assert.equal(instance.__basePollingInterval(), 120000);
});

test('__basePollingInterval: errorCount=10 → 120s (cap holds)', () => {
  const { instance } = makeInstance();
  instance.polling.errorCount = 10;
  assert.equal(instance.__basePollingInterval(), 120000);
});

test('__basePollingInterval: errorCount overrides activity-based rules', () => {
  const { instance } = makeInstance({ lastActivityTime: Date.now() - 3600000 });
  instance.polling.errorCount = 1;
  // Activity rule would say 5min (300000), error backoff says 30s — error wins
  assert.equal(instance.__basePollingInterval(), 30000);
});

// =====================================================================

// =====================================================================
// jitter + 授業中の backoff + 止めない — 2026-09-08 の 429 storm 対策
// =====================================================================

test('getPollingInterval: 基準値の ±20% にばらける (30 台の位相を揃えない)', () => {
  const { instance } = makeInstance({ lastActivityTime: Date.now() });
  for (let i = 0; i < 50; i++) {
    const v = instance.getPollingInterval();
    assert.ok(v >= 4000 && v <= 6000, `5000 の ±20% 内: ${v}`);
  }
});

test('__basePollingInterval: 授業中の失敗は 8s → 16s → 20s 上限 (フェーズ検知を 20 秒以上遅らせない)', () => {
  const { instance } = makeInstance({ lastActivityTime: Date.now() });
  instance.state.lessonPhase = { lessonId: 'l1', phaseIndex: 0, screenRole: 'input' };
  instance.polling.errorCount = 1;
  assert.equal(instance.__basePollingInterval(), 8000);
  instance.polling.errorCount = 2;
  assert.equal(instance.__basePollingInterval(), 16000);
  instance.polling.errorCount = 3;
  assert.equal(instance.__basePollingInterval(), 20000);
  instance.polling.errorCount = 10;
  assert.equal(instance.__basePollingInterval(), 20000);
});

test('schedulePollingCheck: 3 回連続で失敗しても polling を止めない (以前は 5 分停止して授業が止まった)', async () => {
  const { instance, timers } = makeInstance({ lastActivityTime: Date.now() });
  instance.state.lessonPhase = { lessonId: 'l1', phaseIndex: 0, screenRole: 'input' };
  instance.state.lastSeenTimestamp = 0;
  instance.runGas = async () => { throw new Error('Quota exceeded (429)'); };
  instance.polling.isActive = true;
  instance.polling.errorCount = 2;

  instance.schedulePollingCheck();
  const scheduled = Array.from(timers.values()).pop();
  assert.ok(scheduled, 'polling timer が登録される');
  await scheduled.fn();

  assert.equal(instance.polling.errorCount, 3);
  assert.equal(instance.polling.isActive, true, '止めない');
  assert.equal(instance.polling.resumeTimerId, null, '5 分後の再開 timer は作らない');
  const next = Array.from(timers.values()).pop();
  assert.ok(next && next !== scheduled, '次の polling が予約される');
  assert.ok(next.delay >= 16000 && next.delay <= 24000, `授業中の backoff (20s ±20%): ${next.delay}`);
});

// stopSimplePolling — both regular and cooldown timers cleared
// =====================================================================

test('stopSimplePolling: clears both active timer and cooldown resume timer', () => {
  const { instance, timers } = makeInstance();
  instance.polling.isActive = true;
  instance.polling.timerId = 42;
  instance.polling.resumeTimerId = 43;
  // Register them in the timer store so clearTimeout can remove them
  timers.set(42, { fn: () => {}, delay: 0, createdAt: 0 });
  timers.set(43, { fn: () => {}, delay: 0, createdAt: 0 });

  instance.stopSimplePolling();

  assert.equal(instance.polling.isActive, false);
  assert.equal(instance.polling.timerId, null);
  assert.equal(instance.polling.resumeTimerId, null);
  assert.equal(timers.has(42), false);
  assert.equal(timers.has(43), false);
});

test('stopSimplePolling: safe to call with no timers set', () => {
  const { instance } = makeInstance();
  // Should not throw
  instance.stopSimplePolling();
  assert.equal(instance.polling.isActive, false);
});

// =====================================================================
// debounceReactionByRow — race-condition fix locked in
// =====================================================================

test('debounceReactionByRow: first call registers timeout + reactionKey', () => {
  const { instance, timers } = makeInstance();
  instance.pendingReactions.set('row_5_LIKE', { reactionKey: 'row_5_LIKE' });

  instance.debounceReactionByRow('row_5', 5, 'LIKE', 'row_5_LIKE');

  assert.equal(instance.reactionDebounceTimeouts.size, 1);
  const stored = instance.reactionDebounceTimeouts.get('row_5');
  assert.ok(stored.timeoutId);
  assert.equal(stored.reactionKey, 'row_5_LIKE');
  assert.equal(timers.size, 1);
});

test('debounceReactionByRow: superseded click cleans up superseded pending entry', () => {
  const { instance, timers } = makeInstance();
  // User clicks LIKE on row 5 — creates pending entry and schedules debounce
  instance.pendingReactions.set('row_5_LIKE', { reactionKey: 'row_5_LIKE' });
  instance.debounceReactionByRow('row_5', 5, 'LIKE', 'row_5_LIKE');
  assert.equal(timers.size, 1);
  const firstTimerId = instance.reactionDebounceTimeouts.get('row_5').timeoutId;

  // User quickly clicks CURIOUS on the same row — should supersede LIKE
  instance.pendingReactions.set('row_5_CURIOUS', { reactionKey: 'row_5_CURIOUS' });
  instance.debounceReactionByRow('row_5', 5, 'CURIOUS', 'row_5_CURIOUS');

  // The old timeout must be cleared (not both alive)
  assert.equal(timers.has(firstTimerId), false, 'Old timeout should have been cleared');
  assert.equal(timers.size, 1);

  // Key fix: the superseded pendingReactions entry must be removed so future
  // clicks on LIKE for this row aren't blocked by stale "already in flight" state
  assert.equal(instance.pendingReactions.has('row_5_LIKE'), false,
    'Superseded pendingReactions entry for row_5_LIKE must be deleted');
  assert.equal(instance.pendingReactions.has('row_5_CURIOUS'), true,
    'Current pending entry must still be present');
});

test('debounceReactionByRow: same reactionKey re-click does NOT delete pending entry', () => {
  const { instance } = makeInstance();
  instance.pendingReactions.set('row_5_LIKE', { reactionKey: 'row_5_LIKE' });
  instance.debounceReactionByRow('row_5', 5, 'LIKE', 'row_5_LIKE');

  // Click LIKE again before 300ms — same reactionKey, just reset the timer
  instance.debounceReactionByRow('row_5', 5, 'LIKE', 'row_5_LIKE');

  assert.equal(instance.pendingReactions.has('row_5_LIKE'), true,
    'Same-reaction re-click must preserve the pending entry');
});

test('debounceReactionByRow: different rows are independent', () => {
  const { instance, timers } = makeInstance();
  instance.pendingReactions.set('row_5_LIKE', {});
  instance.pendingReactions.set('row_7_LIKE', {});

  instance.debounceReactionByRow('row_5', 5, 'LIKE', 'row_5_LIKE');
  instance.debounceReactionByRow('row_7', 7, 'LIKE', 'row_7_LIKE');

  assert.equal(timers.size, 2);
  assert.equal(instance.reactionDebounceTimeouts.size, 2);
  assert.equal(instance.pendingReactions.has('row_5_LIKE'), true);
  assert.equal(instance.pendingReactions.has('row_7_LIKE'), true);
});

test('debounceReactionByRow: storing reactionKey allows identifying which reaction was scheduled', () => {
  const { instance } = makeInstance();
  instance.debounceReactionByRow('row_5', 5, 'UNDERSTAND', 'row_5_UNDERSTAND');

  const stored = instance.reactionDebounceTimeouts.get('row_5');
  assert.equal(stored.reactionKey, 'row_5_UNDERSTAND');
});

// =====================================================================
// runContainerAction — data-action delegation routing
// =====================================================================

test('runContainerAction: reload action calls window.safeReload when available', () => {
  const { instance, ctx } = makeInstance();
  let called = false;
  ctx.window.safeReload = () => { called = true; };

  instance.runContainerAction('reload', null);
  assert.equal(called, true);
});

test('runContainerAction: reload falls back to location.reload when safeReload missing', () => {
  const { instance, ctx } = makeInstance();
  let locationReloadCalled = false;
  ctx.window.safeReload = undefined;
  ctx.window.location = { reload: () => { locationReloadCalled = true; } };

  instance.runContainerAction('reload', null);
  assert.equal(locationReloadCalled, true);
});

test('runContainerAction: retry-load calls loadSheetData with bypassCache + isInitialLoad', () => {
  const { instance } = makeInstance();
  let captured = null;
  instance.loadSheetData = (options) => { captured = options; };

  instance.runContainerAction('retry-load', null);
  assert.deepEqual({ ...captured }, { bypassCache: true, isInitialLoad: true });
});

test('runContainerAction: retry-load is safe when loadSheetData is absent', () => {
  const { instance } = makeInstance();
  instance.loadSheetData = undefined;
  // Should not throw
  instance.runContainerAction('retry-load', null);
});

test('runContainerAction: unknown action logs warning, does not throw', () => {
  const { instance } = makeInstance();
  // Should not throw
  instance.runContainerAction('totally-unknown', null);
});

// =====================================================================
// simulateServerExclusiveReaction — optimistic update for mutual-exclusion
// =====================================================================

function makeReactionInstance(itemReactions) {
  const { instance } = makeInstance();
  // reactionTypes is set in the constructor; re-set it on the instance for isolation.
  instance.reactionTypes = [
    { key: 'LIKE' },
    { key: 'UNDERSTAND' },
    { key: 'CURIOUS' }
  ];
  const item = { rowIndex: 1, reactions: itemReactions || {} };
  return { instance, item };
}

test('simulateServerExclusiveReaction: adds reaction when user had none', () => {
  const { instance, item } = makeReactionInstance({
    LIKE: { count: 3, reacted: false },
    UNDERSTAND: { count: 1, reacted: false },
    CURIOUS: { count: 0, reacted: false }
  });

  const result = instance.simulateServerExclusiveReaction(item, 'LIKE');
  assert.equal(result.changed, true);
  assert.equal(result.action, 'added');
  assert.equal(result.userReaction, 'LIKE');
  assert.equal(result.reactions.LIKE.count, 4);
  assert.equal(result.reactions.LIKE.reacted, true);
});

test('simulateServerExclusiveReaction: toggles off when user clicks same reaction', () => {
  const { instance, item } = makeReactionInstance({
    LIKE: { count: 4, reacted: true },
    UNDERSTAND: { count: 0, reacted: false },
    CURIOUS: { count: 0, reacted: false }
  });

  const result = instance.simulateServerExclusiveReaction(item, 'LIKE');
  assert.equal(result.action, 'removed');
  assert.equal(result.userReaction, null);
  assert.equal(result.reactions.LIKE.count, 3);
  assert.equal(result.reactions.LIKE.reacted, false);
});

test('simulateServerExclusiveReaction: switches from one reaction to another', () => {
  const { instance, item } = makeReactionInstance({
    LIKE: { count: 5, reacted: true },
    UNDERSTAND: { count: 2, reacted: false },
    CURIOUS: { count: 1, reacted: false }
  });

  const result = instance.simulateServerExclusiveReaction(item, 'CURIOUS');
  assert.equal(result.action, 'changed');
  assert.equal(result.userReaction, 'CURIOUS');
  // Old reaction decremented
  assert.equal(result.reactions.LIKE.count, 4);
  assert.equal(result.reactions.LIKE.reacted, false);
  // New reaction incremented
  assert.equal(result.reactions.CURIOUS.count, 2);
  assert.equal(result.reactions.CURIOUS.reacted, true);
});

test('simulateServerExclusiveReaction: count never goes below zero on decrement', () => {
  const { instance, item } = makeReactionInstance({
    LIKE: { count: 0, reacted: true }, // edge: reacted but count=0 (corrupt state)
    UNDERSTAND: { count: 0, reacted: false },
    CURIOUS: { count: 0, reacted: false }
  });

  const result = instance.simulateServerExclusiveReaction(item, 'LIKE');
  assert.equal(result.reactions.LIKE.count, 0); // Math.max(0, -1) = 0
});

test('simulateServerExclusiveReaction: initializes missing reaction types', () => {
  const { instance } = makeReactionInstance();
  const item = { rowIndex: 1, reactions: { LIKE: { count: 1, reacted: false } } };
  // UNDERSTAND and CURIOUS are missing from the input

  const result = instance.simulateServerExclusiveReaction(item, 'CURIOUS');
  // CURIOUS should be initialized and then incremented
  assert.equal(result.reactions.CURIOUS.count, 1);
  assert.equal(result.reactions.CURIOUS.reacted, true);
  assert.equal(result.reactions.UNDERSTAND.count, 0);
});

test('simulateServerExclusiveReaction: handles item with no reactions object', () => {
  const { instance } = makeReactionInstance();
  const item = { rowIndex: 1 };
  const result = instance.simulateServerExclusiveReaction(item, 'LIKE');
  assert.equal(result.reactions.LIKE.count, 1);
  assert.equal(result.reactions.LIKE.reacted, true);
});

test('simulateServerExclusiveReaction: does not mutate input item.reactions', () => {
  const input = { LIKE: { count: 5, reacted: false } };
  const { instance } = makeReactionInstance();
  const item = { rowIndex: 1, reactions: input };

  instance.simulateServerExclusiveReaction(item, 'LIKE');
  // The function writes to predictedReactions (a deep clone), so input must
  // be untouched. If this ever regresses, optimistic UI would corrupt the
  // canonical state before the server response arrives.
  assert.equal(input.LIKE.count, 5);
  assert.equal(input.LIKE.reacted, false);
});

// =====================================================================
// clearCache — option-gated cache clearing
// =====================================================================

test('clearCache: clears this.cache unconditionally', () => {
  const { instance } = makeInstance();
  let cleared = false;
  instance.cache = { clear: () => { cleared = true; } };
  instance.clearCache();
  assert.equal(cleared, true);
});

test('clearCache: tolerates missing this.cache', () => {
  const { instance } = makeInstance();
  instance.cache = null;
  // Should not throw
  instance.clearCache();
});

test('clearCache: clears reactionCache only when includeReactions=true', () => {
  const { instance } = makeInstance();
  let reactionCleared = false;
  instance.cache = { clear: () => {} };
  instance.reactionCache = { clear: () => { reactionCleared = true; } };

  instance.clearCache({ includeReactions: false });
  assert.equal(reactionCleared, false);

  instance.clearCache({ includeReactions: true });
  assert.equal(reactionCleared, true);
});

test('clearCache: swallows errors so callers don\'t break', () => {
  const { instance } = makeInstance();
  instance.cache = { clear: () => { throw new Error('cache dead'); } };
  // Should not throw
  instance.clearCache();
});

// =====================================================================
// enhanceError — annotates an error with function-specific user message
// =====================================================================

test('enhanceError: wraps error with originalError reference and metadata', () => {
  const { instance } = makeInstance();
  const original = new Error('network timeout');
  const enhanced = instance.enhanceError(original, 'addReaction', [1, 'LIKE']);

  assert.equal(enhanced instanceof Error, true);
  assert.equal(enhanced.message, 'network timeout');
  assert.equal(enhanced.originalError, original);
  assert.equal(enhanced.functionName, 'addReaction');
  assert.deepEqual([...enhanced.arguments], [1, 'LIKE']);
});

test('enhanceError: addReaction → リアクション処理エラー userMessage', () => {
  const { instance } = makeInstance();
  const e = instance.enhanceError(new Error('x'), 'addReaction', []);
  assert.equal(e.userMessage, 'リアクション処理エラー');
});

test('enhanceError: toggleHighlight → ハイライト処理エラー userMessage', () => {
  const { instance } = makeInstance();
  const e = instance.enhanceError(new Error('x'), 'toggleHighlight', []);
  assert.equal(e.userMessage, 'ハイライト処理エラー');
});

test('enhanceError: unknown funcName → default 通信エラー userMessage', () => {
  const { instance } = makeInstance();
  const e = instance.enhanceError(new Error('x'), 'getData', []);
  assert.equal(e.userMessage, '通信エラー');
});

test('enhanceError: accepts a string error (not an Error instance)', () => {
  const { instance } = makeInstance();
  const e = instance.enhanceError('raw error string', 'addReaction', []);
  assert.equal(e.message, 'raw error string');
});

test('enhanceError: accepts an error with falsy message, falls back to stringified', () => {
  const { instance } = makeInstance();
  // Object-like error without a .message
  const fakeError = { toString: () => 'toString fallback' };
  const e = instance.enhanceError(fakeError, 'addReaction', []);
  // new Error(error.message || error) — since .message is undefined, falls to the object,
  // and new Error(obj) coerces to string via toString()
  assert.match(e.message, /toString fallback/);
});

// =====================================================================
// generateRequestId — unique request tracking identifier
// =====================================================================

test('generateRequestId: returns a string with the req_ prefix', () => {
  const { instance } = makeInstance();
  const id = instance.generateRequestId();
  assert.equal(typeof id, 'string');
  assert.ok(id.startsWith('req_'));
});

test('generateRequestId: produces unique IDs across rapid calls', () => {
  const { instance } = makeInstance();
  const ids = new Set();
  for (let i = 0; i < 100; i += 1) {
    ids.add(instance.generateRequestId());
  }
  assert.equal(ids.size, 100, 'All 100 generated IDs must be unique');
});

test('generateRequestId: structure is req_<timestamp36>_<random9>', () => {
  const { instance } = makeInstance();
  const id = instance.generateRequestId();
  // req_ + at least one char + _ + 9 chars (substr(2, 9))
  assert.match(id, /^req_[0-9a-z]+_[0-9a-z]{1,9}$/);
});


test('loadSheetData: bypasses gate when isInitialLoad=true (初回起動経路は素通り)', async () => {
  const { instance } = makeInstance();
  let pastReloadCalled = false;
  let initialLoadStarted = false;
  instance.state = {
    userId: 'u1', viewingPastProfile: '導入', isLoading: false,
    currentAnswers: []
  };
  instance.__vizLoadPastProfile = () => { pastReloadCalled = true; };
  // performDataLoad を stub 化（実 fetch しない）
  instance.performDataLoad = async () => { initialLoadStarted = true; };
  instance.showLoadingOverlay = () => {};
  instance.hideLoadingOverlay = () => {};
  instance.clearCache = () => {};
  instance.shouldClearCache = () => false;
  instance.handleError = () => {};

  await instance.loadSheetData({ isInitialLoad: true });
  assert.equal(pastReloadCalled, false, '初回起動では gate しない');
  assert.equal(initialLoadStarted, true, '通常経路に進む');
});


// =====================================================================
// populateClassFilter: シンプル化版 (Option B 周辺の cleanup)
// 設計: profile 切替時に classFilter='すべて' リセットを掛けるので、ここは表記揺れを
//   気にせず「現在 DOM 値が new uniqueClasses にあれば保持、なければ 'すべて'」のシンプル動作。
// =====================================================================

function makeFilterMock() {
  let _value = '';
  let _html = '';
  return {
    get value() { return _value; },
    set value(v) { _value = v; },
    get innerHTML() { return _html; },
    set innerHTML(v) { _html = v; },
    classList: { remove: () => {}, add: () => {} }
  };
}

test('populateClassFilter: keeps previous selection when present in new data', () => {
  const { instance } = makeInstance();
  const cf = makeFilterMock();
  cf.value = '4組';
  instance.elements = { classFilter: cf };
  instance.loadPersistedClassFilter = () => '4組';
  instance.persistClassFilter = () => {};
  instance.populateClassFilter([{ class: '4組' }, { class: '5組' }]);
  assert.equal(cf.value, '4組');
});

test('populateClassFilter: falls back to すべて when selection missing + syncs sessionStorage', () => {
  const { instance } = makeInstance();
  const cf = makeFilterMock();
  cf.value = '4組';
  let persisted = null;
  instance.elements = { classFilter: cf };
  instance.loadPersistedClassFilter = () => '4組';
  instance.persistClassFilter = (v) => { persisted = v; };
  instance.populateClassFilter([{ class: '1組' }, { class: '2組' }]);
  assert.equal(cf.value, 'すべて');
  assert.equal(persisted, 'すべて', 'sessionStorage も同期されて次回 fetch も整合');
});


// =====================================================================
// updateDisplaySettingsFromAPI: profile 切替時の UNIFIED_CONFIG 同期
// Why: 管理モード ON/OFF が profile 跨ぎで正しく動くため、admin mode 中でも
//   UNIFIED_CONFIG.displaySettings は常に最新値で同期される必要がある。
//   旧コードは admin mode 中に早期 return して UNIFIED_CONFIG が bootstrap 時の値で
//   stale → 管理モード OFF にしたとき初期 profile の設定に戻ってしまうバグだった。
// =====================================================================

test('updateDisplaySettingsFromAPI: 通常モード時は globals + UNIFIED_CONFIG を両方更新', () => {
  const { instance, ctx } = makeInstance();
  ctx.window.UNIFIED_CONFIG = { displaySettings: { showNames: false, showReactions: false } };
  ctx.window.showAdminFeatures = false;
  ctx.window.showCounts = false;
  ctx.window.displayMode = 'anonymous';
  instance.state = { showCounts: false, displayMode: 'anonymous' };
  instance.clearCache = () => {};
  instance.renderWithCurrentData = () => {};

  instance.updateDisplaySettingsFromAPI({ showNames: true, showReactions: true });

  assert.equal(ctx.window.UNIFIED_CONFIG.displaySettings.showNames, true);
  assert.equal(ctx.window.UNIFIED_CONFIG.displaySettings.showReactions, true);
  assert.equal(ctx.window.showCounts, true);
  assert.equal(ctx.window.displayMode, 'named');
});

test('updateDisplaySettingsFromAPI: admin mode 中も UNIFIED_CONFIG は同期するが globals は触らない', () => {
  // Regression: admin mode 中に profile 切替 → updateDisplaySettingsFromAPI 早期 return →
  //   UNIFIED_CONFIG stale → admin OFF で初期値に戻る、というバグの修正テスト。
  const { instance, ctx } = makeInstance();
  ctx.window.UNIFIED_CONFIG = { displaySettings: { showNames: false, showReactions: false } };
  ctx.window.showAdminFeatures = true;        // admin mode 中
  ctx.window.showCounts = true;               // admin forced
  ctx.window.displayMode = 'named';           // admin forced
  instance.state = { showCounts: true, displayMode: 'named' };
  instance.clearCache = () => {};
  instance.renderWithCurrentData = () => {};

  instance.updateDisplaySettingsFromAPI({ showNames: true, showReactions: true });

  // UNIFIED_CONFIG は最新値に同期される (将来 admin OFF にしたとき参照されるため)
  assert.equal(ctx.window.UNIFIED_CONFIG.displaySettings.showNames, true);
  assert.equal(ctx.window.UNIFIED_CONFIG.displaySettings.showReactions, true);
  // 表示用 globals は admin forced 値のまま (触らない)
  assert.equal(ctx.window.showCounts, true);
  assert.equal(ctx.window.displayMode, 'named');
});

test('updateDisplaySettingsFromAPI: 無効入力は何もしない', () => {
  const { instance, ctx } = makeInstance();
  ctx.window.UNIFIED_CONFIG = { displaySettings: { showNames: false } };
  instance.updateDisplaySettingsFromAPI(null);
  instance.updateDisplaySettingsFromAPI(undefined);
  instance.updateDisplaySettingsFromAPI('not-object');
  assert.equal(ctx.window.UNIFIED_CONFIG.displaySettings.showNames, false, '不変');
});

// =====================================================================
// 授業モードの画面: 教師は答えない / 送信後は「送りました」画面 / フェーズ切替で戻す
// =====================================================================

function makeLessonInstance() {
  const { instance, ctx } = makeInstance();
  const calls = [];
  const overlay = { parentNode: { removeChild: () => { calls.push('teardown'); } } };
  let overlayPresent = false;
  ctx.document.getElementById = (id) => (id === 'lessonScreen' && overlayPresent) ? overlay : null;
  ctx.document.body.appendChild = () => { overlayPresent = true; calls.push('overlay'); };
  ctx.document.body.classList = { add: () => {}, remove: () => {}, toggle: () => {} };
  instance.state = { userId: 'u1', lessonPhase: null, lessonDraft: null, lessonSent: null, lessonEditing: false, lessonRecordKey: null, currentAnswers: [] };
  instance.__renderLessonInput = () => { calls.push('input'); };
  instance.__renderLessonDiscuss = () => { calls.push('discuss'); };
  instance.__renderLessonReflect = () => { calls.push('reflect'); };
  instance.showToast = () => {};
  instance.loadSheetData = () => {};
  instance.runGas = () => Promise.resolve({ success: true, data: { phases: [] } });
  return { instance, calls, ctx };
}

test('授業画面: 教師 (isEditor) には児童の入力画面を出さない', () => {
  const { instance, calls, ctx } = makeLessonInstance();
  ctx.document.body.classList.contains = () => false;
  instance.state.isEditor = true;
  instance.state.currentAnswers = [];
  instance.__applyLessonPhase({ lessonId: 'l1', phaseIndex: 0, screenRole: 'input', phaseName: '考える' });
  assert.ok(!calls.includes('input'), '教師に児童の入力画面が出ない');
  // 出会う: 待機画面も外れて分布そのもの
  instance.__applyLessonPhase({ lessonId: 'l1', phaseIndex: 1, screenRole: 'browse', phaseName: '出会う' });
  assert.ok(calls.includes('teardown'));
});

test('授業画面: 児童には入力画面が出る', () => {
  const { instance, calls } = makeLessonInstance();
  instance.state.isEditor = false;
  instance.__applyLessonPhase({ lessonId: 'l1', phaseIndex: 0, screenRole: 'input', phaseName: '考える' });
  assert.ok(calls.includes('input'));
});

test('授業画面: フェーズが変わると送信済み・下書き・置き直し中の状態を捨てる', () => {
  const { instance } = makeLessonInstance();
  instance.state.isEditor = false;
  instance.__applyLessonPhase({ lessonId: 'l1', phaseIndex: 0, screenRole: 'input' });
  instance.state.lessonSent = { lessonId: 'l1', phaseIndex: 0, numericX: 2, numericY: 4, reason: 'r' };
  instance.state.lessonDraft = { numericX: 2, numericY: 4, reason: 'r' };
  instance.state.lessonEditing = true;
  instance.state.lessonRecordKey = 'l1:0';
  instance.__applyLessonPhase({ lessonId: 'l1', phaseIndex: 1, screenRole: 'browse' });
  assert.equal(instance.state.lessonSent, null);
  assert.equal(instance.state.lessonDraft, null);
  assert.equal(instance.state.lessonEditing, false);
  assert.equal(instance.state.lessonRecordKey, null);
});

test('授業画面: 同じフェーズの polling では送信済み状態を保つ (入力画面に戻されない)', () => {
  const { instance, calls } = makeLessonInstance();
  instance.state.isEditor = false;
  instance.__applyLessonPhase({ lessonId: 'l1', phaseIndex: 0, screenRole: 'input' });
  instance.state.lessonSent = { lessonId: 'l1', phaseIndex: 0, numericX: 2, numericY: 4, reason: 'r' };
  const before = calls.length;
  instance.__applyLessonPhase({ lessonId: 'l1', phaseIndex: 0, screenRole: 'input' });
  assert.equal(calls.length, before, '描き直さない');
  assert.ok(instance.state.lessonSent, '送信済みのまま');
});

test('授業のクラス: 登録クラスが 1 つなら自動でそれを送る (児童はフィルタを触れない)', () => {
  const { instance } = makeLessonInstance();
  instance.state.lessonPhase = { lessonId: 'l1', phaseIndex: 0, screenRole: 'input', classes: ['6年1組'] };
  instance.elements = { classFilter: { value: 'すべて' } };
  assert.equal(instance.__lessonClass(), '6年1組');
});

test('授業のクラス: 複数クラスなら端末に覚えた選択を使い、無ければ空', () => {
  const { instance, ctx } = makeLessonInstance();
  instance.state.lessonPhase = { lessonId: 'l1', phaseIndex: 0, screenRole: 'input', classes: ['6年1組', '6年2組'] };
  instance.elements = { classFilter: { value: 'すべて' } };
  assert.equal(instance.__lessonClass(), '');
  ctx.localStorage.getItem = (k) => k === 'lessonClass:u1' ? '6年2組' : null;
  assert.equal(instance.__lessonClass(), '6年2組');
});

// =====================================================================
// 教師の入力フェーズ: 分布を出さず、問い + 送信済み人数 (明示のボタンでだけ開く)
// =====================================================================

function makeTeacherInstance() {
  const { instance, calls, ctx } = makeLessonInstance();
  instance.state.isEditor = true;
  instance.state.currentAnswers = [];
  instance.__renderLessonTeacherWait = (phase) => { calls.push('teacherWait:' + phase.screenRole); };
  return { instance, calls, ctx };
}

test('教師: 考える (input) では分布ではなく待機画面 (問い + 送信済み人数) を出す', () => {
  const { instance, calls } = makeTeacherInstance();
  instance.__applyLessonPhase({ lessonId: 'l1', phaseIndex: 0, screenRole: 'input', phaseName: '考える' });
  assert.ok(calls.includes('teacherWait:input'));
  assert.ok(!calls.includes('teardown'), '待機画面は分布の上にかぶせる');
});

test('教師: もう一度考える (reinput) も同じく待機画面', () => {
  const { instance, calls } = makeTeacherInstance();
  instance.__applyLessonPhase({ lessonId: 'l1', phaseIndex: 3, screenRole: 'reinput', phaseName: 'もう一度考える' });
  assert.ok(calls.includes('teacherWait:reinput'));
});

test('教師: 出会う (browse) では待機画面を外して分布を見せる', () => {
  const { instance, calls } = makeTeacherInstance();
  instance.__applyLessonPhase({ lessonId: 'l1', phaseIndex: 1, screenRole: 'browse', phaseName: '出会う' });
  assert.ok(!calls.some(c => c.startsWith('teacherWait')));
});


test('送信済み人数: 行数とクラス別に集計する', () => {
  const { instance } = makeTeacherInstance();
  instance.state.currentAnswers = [
    { class: '6年1組' }, { class: '6年1組' }, { class: '6年2組' }, { class: '' }
  ];
  const stats = instance.__lessonSubmissionStats();
  assert.equal(stats.total, 4);
  assert.deepEqual(JSON.parse(JSON.stringify(stats.byClass)), { '6年1組': 2, '6年2組': 1, '': 1 });
});

// =====================================================================
// 教師の操作面 (フェーズ切替) が polling と同期し、どの画面からも進める
// =====================================================================

function makeNavInstance() {
  const { instance, calls, ctx } = makeTeacherInstance();
  ctx.document.body.classList.contains = () => false;
  ctx.document.querySelectorAll = () => [];
  instance.__renderBoardPhaseNav = () => { calls.push('nav:' + (instance.state.boardPhaseNav ? instance.state.boardPhaseNav.activePhaseIndex : 'none')); };
  instance.__initBoardPhaseNav = () => { calls.push('nav:init'); return Promise.resolve(); };
  return { instance, calls };
}

test('操作面: 管理パネルで進めた変更が polling で届くと現在地が追従する', () => {
  const { instance, calls } = makeNavInstance();
  instance.state.boardPhaseNav = { lessonId: 'l1', activePhaseIndex: 0, phases: [{ index: 0, name: '考える' }, { index: 1, name: '出会う' }] };
  instance.__applyLessonPhase({ lessonId: 'l1', phaseIndex: 1, screenRole: 'browse', phaseName: '出会う' });
  assert.equal(instance.state.boardPhaseNav.activePhaseIndex, 1);
  assert.ok(calls.includes('nav:1'));
});

test('操作面: ボードを開いたあとに授業が始まると操作面を読み直す', () => {
  const { instance, calls } = makeNavInstance();
  instance.state.boardPhaseNav = null;
  instance.__applyLessonPhase({ lessonId: 'l1', phaseIndex: 0, screenRole: 'input', phaseName: '考える' });
  assert.ok(calls.includes('nav:init'));
});

test('操作面: 授業が終わると操作面を畳む', () => {
  const { instance, calls } = makeNavInstance();
  instance.state.boardPhaseNav = { lessonId: 'l1', activePhaseIndex: 4, phases: [] };
  instance.__applyLessonPhase(null);
  assert.equal(instance.state.boardPhaseNav, null);
  assert.ok(calls.includes('nav:none'));
});

test('操作面: 教師には「進みました」トーストを出さない (切替の success と二重になる)', () => {
  const { instance } = makeNavInstance();
  const toasts = [];
  instance.showToast = (m) => toasts.push(m);
  instance.state.boardPhaseNav = { lessonId: 'l1', activePhaseIndex: 0, phases: [{ index: 0 }, { index: 1 }] };
  instance.__applyLessonPhase({ lessonId: 'l1', phaseIndex: 0, screenRole: 'input' });
  instance.__applyLessonPhase({ lessonId: 'l1', phaseIndex: 1, screenRole: 'browse' });
  assert.equal(toasts.length, 0);
});

test('操作面: 児童には「進みました」トーストを出す', () => {
  const { instance } = makeLessonInstance();
  instance.state.isEditor = false;
  const toasts = [];
  instance.showToast = (m) => toasts.push(m);
  instance.__applyLessonPhase({ lessonId: 'l1', phaseIndex: 0, screenRole: 'input' });
  instance.__applyLessonPhase({ lessonId: 'l1', phaseIndex: 1, screenRole: 'browse', phaseName: '出会う' });
  assert.equal(toasts.length, 1);
});


test('送信の前提: 4 象限は縦横の両方、数直線は横だけ', async () => {
  const { instance } = makeLessonInstance();
  const toasts = [];
  instance.showToast = (m) => toasts.push(m);
  instance.state.lessonPhase = { lessonId: 'l1', phaseIndex: 0, screenRole: 'input', formTemplate: 'matrix' };
  instance.state.lessonDraft = { numericX: 3, numericY: null };
  await instance.__submitLessonAnswer();
  assert.match(toasts[0], /位置/, '4 象限で縦が無ければ位置の選択を求める');
  instance.state.lessonPhase = { lessonId: 'l1', phaseIndex: 0, screenRole: 'input', formTemplate: 'numberline' };
  instance.state.lessonDraft = { numericX: 3, numericY: null };
  await instance.__submitLessonAnswer();
  // 位置は通り、次の検証 (理由の未入力) に進んでいる
  assert.equal(toasts.length, 2);
  assert.doesNotMatch(toasts[1], /位置/);
});

test('教師の待機画面: クラス別の内訳はクラスが 2 つ以上あるときだけ集計に意味がある', () => {
  const { instance } = makeTeacherInstance();
  instance.state.currentAnswers = [{ class: '' }, { class: '' }];
  const stats = instance.__lessonSubmissionStats();
  assert.equal(Object.keys(stats.byClass).filter(c => c).length, 0, '空のクラスは内訳に出さない');
});

test('フェーズ切替: 押した瞬間に「切り替えています」、応答後は「読み込んでいます」、データが来たら次の画面', async () => {
  const { instance, calls, ctx } = makeTeacherInstance();
  const stages = [];
  instance.__renderLessonTransition = (t) => { stages.push(t.stage + ':' + t.to.name); };
  instance.__renderBoardPhaseNav = () => {};
  instance.loadSheetData = () => { stages.push('load'); return Promise.resolve(); };
  instance.elements = {};
  instance.state.boardPhaseNav = { lessonId: 'l1', activePhaseIndex: 0, phases: [
    { index: 0, name: '考える', screenRole: 'input' }, { index: 1, name: '出会う', screenRole: 'browse' }
  ] };
  instance.state.lessonPhase = { lessonId: 'l1', phaseIndex: 0, screenRole: 'input', phaseName: '考える' };
  ctx.runServer = () => { stages.push('server'); return Promise.resolve({ success: true, data: { activePhaseIndex: 1 } }); };
  ctx.window.notifications = { error: () => {}, success: () => { stages.push('toast'); } };
  await instance.__switchBoardPhase(instance.state.boardPhaseNav.phases[1]);
  assert.deepEqual(stages, ['switching:出会う', 'server', 'loading:出会う', 'load']);
  assert.equal(instance.state.lessonTransition, null);
  assert.equal(instance.state.lessonPhase.phaseIndex, 1);
  assert.equal(instance.state.lessonPhase.screenRole, 'browse', '出会う = 分布そのもの');
  assert.ok(!stages.includes('toast'), '完了のトーストは出さない');
});

// =====================================================================
// 回答への操作列: 1 つの部品で、権限に応じて リアクション / ハイライト / 削除 を並べる
// =====================================================================

function makeActionsInstance(isEditor, showAdminFeatures) {
  const { instance } = makeInstance();
  instance.state = { userId: 'u1', isEditor, showAdminFeatures, showCounts: true };
  instance.reactionTypes = [{ key: 'LIKE', icon: 'like' }, { key: 'UNDERSTAND', icon: 'understand' }, { key: 'CURIOUS', icon: 'curious' }];
  instance.getIcon = (name) => '<i data-icon="' + name + '"></i>';
  return instance;
}

test('操作列: 児童はリアクション 3 つだけ', () => {
  const html = makeActionsInstance(false, false).__buildAnswerActionsHtml({ rowIndex: 5, reactions: {}, highlight: false });
  assert.equal((html.match(/reaction-btn/g) || []).length, 3);
  assert.ok(!html.includes('highlight-btn'));
  assert.ok(!html.includes('delete-answer-btn'));
});

test('操作列: 編集者にはハイライト、管理モードでは削除も同じ列に出る', () => {
  const editor = makeActionsInstance(true, false).__buildAnswerActionsHtml({ rowIndex: 5, reactions: {}, highlight: true });
  assert.ok(editor.includes('highlight-btn'));
  assert.ok(editor.includes('aria-pressed="true"'));
  assert.ok(!editor.includes('delete-answer-btn'));
  const admin = makeActionsInstance(true, true).__buildAnswerActionsHtml({ rowIndex: 5, reactions: {}, highlight: false }, { size: 'modal' });
  assert.ok(admin.includes('delete-answer-btn'));
  assert.ok(admin.includes('answer-actions--lg'), 'モーダルは大きいサイズ');
  assert.ok(admin.includes('data-row-index="5"'));
});

test('操作列: 分布の意見一覧はリアクションだけ (編集者でも)', () => {
  const html = makeActionsInstance(true, true).createReactionButtons({ rowIndex: 5, reactions: {} });
  assert.ok(!html.includes('highlight-btn') && !html.includes('delete-answer-btn'));
});

// =====================================================================
// refreshAfterDelete — 削除後は必ずサーバから読み直す (2 個目の削除が古い行番号で止まらない)
// =====================================================================

test('refreshAfterDelete: 手元の一覧から消したあと、必ず bypassCache で loadSheetData を予約する', async () => {
  const { instance, timers } = makeInstance();
  instance.state.currentAnswers = [{ rowIndex: 2 }, { rowIndex: 3 }, { rowIndex: 4 }];
  instance.state.allAnswers = [{ rowIndex: 2 }, { rowIndex: 3 }, { rowIndex: 4 }];
  instance.state.showAdminFeatures = true;
  instance.updateAnswerCount = () => {};
  instance.updateAdminButtonUI = () => {};
  const loads = [];
  instance.loadSheetData = (opts) => { loads.push(opts); return Promise.resolve(); };

  await instance.refreshAfterDelete(3);

  assert.deepEqual(instance.state.currentAnswers.map((a) => a.rowIndex), [2, 4], '消した行は手元からも消える');
  const scheduled = Array.from(timers.values());
  assert.ok(scheduled.length >= 1, '再読込が予約される');
  scheduled.forEach((t) => t.fn());
  assert.equal(loads.length, 1);
  assert.equal(loads[0].bypassCache, true, 'サーバの行番号で取り直す');
  assert.equal(loads[0].preserveAdminMode, true, '管理モードは維持する');
  assert.equal(loads[0].showLoading, false, '一致しているときは静かに読み直す');
});
