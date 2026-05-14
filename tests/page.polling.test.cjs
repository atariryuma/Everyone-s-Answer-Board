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

test('getPollingInterval: <1min of activity → 5s (授業中)', () => {
  const { instance } = makeInstance({ lastActivityTime: Date.now() });
  assert.equal(instance.getPollingInterval(), 5000);
});

test('getPollingInterval: 1-5min → 15s', () => {
  const { instance } = makeInstance({ lastActivityTime: Date.now() - 120000 });
  assert.equal(instance.getPollingInterval(), 15000);
});

test('getPollingInterval: 5-15min → 1min', () => {
  const { instance } = makeInstance({ lastActivityTime: Date.now() - 600000 });
  assert.equal(instance.getPollingInterval(), 60000);
});

test('getPollingInterval: >15min → 5min', () => {
  const { instance } = makeInstance({ lastActivityTime: Date.now() - 3600000 });
  assert.equal(instance.getPollingInterval(), 300000);
});

test('getPollingInterval: errorCount=1 → 30s (exponential backoff kicks in)', () => {
  const { instance } = makeInstance();
  instance.polling.errorCount = 1;
  // Even if user is currently active, error backoff takes precedence
  assert.equal(instance.getPollingInterval(), 30000);
});

test('getPollingInterval: errorCount=2 → 60s (exponential doubles)', () => {
  const { instance } = makeInstance();
  instance.polling.errorCount = 2;
  assert.equal(instance.getPollingInterval(), 60000);
});

test('getPollingInterval: errorCount=3 → 120s (capped)', () => {
  const { instance } = makeInstance();
  instance.polling.errorCount = 3;
  assert.equal(instance.getPollingInterval(), 120000);
});

test('getPollingInterval: errorCount=10 → 120s (cap holds)', () => {
  const { instance } = makeInstance();
  instance.polling.errorCount = 10;
  assert.equal(instance.getPollingInterval(), 120000);
});

test('getPollingInterval: errorCount overrides activity-based rules', () => {
  const { instance } = makeInstance({ lastActivityTime: Date.now() - 3600000 });
  instance.polling.errorCount = 1;
  // Activity rule would say 5min (300000), error backoff says 30s — error wins
  assert.equal(instance.getPollingInterval(), 30000);
});

// =====================================================================
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

// =====================================================================
// Option B: loadSheetData gate during past-phase view
// Why: 過去フェーズ閲覧中に classFilter/sortOrder 変更 / refreshContent ボタンが押されると
//   従来は loadSheetData 経由で「現在フェーズ」に戻ってしまっていた。viewingPastProfile が
//   立っている間は __vizLoadPastProfile に分岐して同じフェーズで再フェッチさせるガードを追加。
//   (page.js の loadSheetData 入口)
// =====================================================================

test('loadSheetData: routes to __vizLoadPastProfile when viewing past phase', async () => {
  const { instance } = makeInstance();
  let pastReloadName = null;
  instance.state = { userId: 'u1', viewingPastProfile: '導入アンケート', isLoading: false };
  instance.__vizLoadPastProfile = (name) => { pastReloadName = name; };

  await instance.loadSheetData({ bypassCache: true, isInitialLoad: false });

  assert.equal(pastReloadName, '導入アンケート');
  // viewingPastProfile は __vizLoadPastProfile 内で再 set されるべきだが、ガード解除直後は null
  assert.equal(instance.state.viewingPastProfile, null);
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

test('loadSheetData: bypasses gate when viewingPastProfile not set', async () => {
  const { instance } = makeInstance();
  let pastReloadCalled = false;
  let loadStarted = false;
  instance.state = {
    userId: 'u1', viewingPastProfile: null, isLoading: false,
    currentAnswers: []
  };
  instance.__vizLoadPastProfile = () => { pastReloadCalled = true; };
  instance.performDataLoad = async () => { loadStarted = true; };
  instance.showLoadingOverlay = () => {};
  instance.hideLoadingOverlay = () => {};
  instance.clearCache = () => {};
  instance.shouldClearCache = () => false;
  instance.handleError = () => {};

  await instance.loadSheetData({ bypassCache: true, isInitialLoad: false });
  assert.equal(pastReloadCalled, false);
  assert.equal(loadStarted, true);
});

// =====================================================================
// normalizeClassToken / populateClassFilter (Option B L1)
// Why: profile 切替で dropdown が silent に 'すべて' に降格して、データが空になる
//   不整合を解消。同一クラスの別表記（"6年1組" ↔ "1組"）は normalize で同一視。
//   真に該当データがない場合だけ toast + auto-recover で 'すべて' に切替える。
// =====================================================================

test('normalizeClassToken: client-side mirror of server (string ↔ digit)', () => {
  const { instance } = makeInstance();
  assert.equal(instance.normalizeClassToken('1組'), '1');
  assert.equal(instance.normalizeClassToken('6年1組'), '1');
  assert.equal(instance.normalizeClassToken('11組'), '11');
  assert.equal(instance.normalizeClassToken('特進'), '特進');
  assert.equal(instance.normalizeClassToken(''), '');
  assert.equal(instance.normalizeClassToken(null), '');
});

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

test('populateClassFilter: keeps previous selection when exact match exists', () => {
  const { instance } = makeInstance();
  const cf = makeFilterMock();
  cf.value = '4組';
  instance.elements = { classFilter: cf };
  instance.loadPersistedClassFilter = () => '4組';
  instance.persistClassFilter = () => {};
  instance.showToast = () => {};
  instance.loadSheetData = async () => {};

  instance.populateClassFilter([{ class: '4組' }, { class: '5組' }]);
  assert.equal(cf.value, '4組');
});

test('populateClassFilter: ALIAS — "6年1組" desired but new data has "1組" → adopts "1組"', () => {
  // 真の同一クラスを別表記で持つ profile 跨ぎケース
  const { instance } = makeInstance();
  const cf = makeFilterMock();
  cf.value = '6年1組';
  let persisted = null;
  let toasted = null;
  let reloaded = false;
  instance.elements = { classFilter: cf };
  instance.loadPersistedClassFilter = () => '6年1組';
  instance.persistClassFilter = (v) => { persisted = v; };
  instance.showToast = (m) => { toasted = m; };
  instance.loadSheetData = async () => { reloaded = true; };

  instance.populateClassFilter([{ class: '1組' }, { class: '2組' }]);
  assert.equal(cf.value, '1組', 'dropdown は新データの表記に揃う');
  assert.equal(persisted, '1組', 'sessionStorage も同期');
  assert.equal(toasted, null, '同一クラスの別表記なら toast 不要');
});

test('populateClassFilter: MISMATCH — "4組" desired but no 4組 anywhere → fallback + toast + auto-recover', async () => {
  const { instance, timers } = makeInstance();
  const cf = makeFilterMock();
  cf.value = '4組';
  let persisted = null;
  let toasted = null;
  let reloaded = false;
  instance.elements = { classFilter: cf };
  instance.loadPersistedClassFilter = () => '4組';
  instance.persistClassFilter = (v) => { persisted = v; };
  instance.showToast = (m) => { toasted = m; };
  instance.loadSheetData = async () => { reloaded = true; };

  instance.populateClassFilter([{ class: '1組' }, { class: '2組' }]);
  assert.equal(cf.value, 'すべて', '別クラスなので "すべて" に降格');
  assert.equal(persisted, 'すべて', 'sessionStorage も "すべて" に');
  assert.match(toasted || '', /4組/, 'toast でユーザーに通知');

  // setTimeout で deferred 登録された loadSheetData が timer に乗っていることを確認
  // (実発火は別 task。timer の存在で間接確認)
  assert.ok(timers.size >= 1, 'auto-recover refetch が予約されている');
});

test('populateClassFilter: no previous selection → silent default to すべて', () => {
  const { instance } = makeInstance();
  const cf = makeFilterMock();
  cf.value = '';
  let toasted = null;
  instance.elements = { classFilter: cf };
  instance.loadPersistedClassFilter = () => null;
  instance.persistClassFilter = () => {};
  instance.showToast = (m) => { toasted = m; };
  instance.loadSheetData = async () => {};

  instance.populateClassFilter([{ class: '1組' }, { class: '2組' }]);
  assert.equal(cf.value, 'すべて');
  assert.equal(toasted, null, '初期状態は toast 出さない');
});
