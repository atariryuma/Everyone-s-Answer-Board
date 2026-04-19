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

test('getPollingInterval: <1min of activity → 15s', () => {
  const { instance } = makeInstance({ lastActivityTime: Date.now() });
  assert.equal(instance.getPollingInterval(), 15000);
});

test('getPollingInterval: 1-5min → 30s', () => {
  const { instance } = makeInstance({ lastActivityTime: Date.now() - 120000 });
  assert.equal(instance.getPollingInterval(), 30000);
});

test('getPollingInterval: 5-15min → 2min', () => {
  const { instance } = makeInstance({ lastActivityTime: Date.now() - 600000 });
  assert.equal(instance.getPollingInterval(), 120000);
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
