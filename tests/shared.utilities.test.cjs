const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

// Extract SharedUtilities.html <script> and evaluate in a vm context with
// minimal browser stubs. Exposes the internal classes (Cache, DebounceManager,
// ThrottleManager) for direct testing — they're pure logic with no DOM needs.
function loadSharedUtilities() {
  const src = fs.readFileSync(path.resolve(__dirname, '../src/SharedUtilities.html'), 'utf8');
  const scriptMatch = src.match(/<script[^>]*>([\s\S]*?)<\/script>/);
  if (!scriptMatch) throw new Error('<script> tag not found');
  const js = scriptMatch[1];

  const timers = new Map();
  let nextTimerId = 1;

  const ctx = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    Map, Set, Date, Math, JSON, Number, Array, Object, Promise, Error, Symbol,
    setTimeout: (fn, delay) => {
      const id = nextTimerId++;
      const entry = { fn, delay, scheduledAt: Date.now() };
      timers.set(id, entry);
      return id;
    },
    clearTimeout: (id) => { timers.delete(id); },
    window: {
      addEventListener: () => {},
      removeEventListener: () => {},
      console: { warn: () => {}, error: () => {} }
    },
    document: {
      addEventListener: () => {},
      removeEventListener: () => {},
      createElement: () => ({
        classList: { add: () => {}, remove: () => {} },
        textContent: '',
        innerHTML: '',
        style: {},
        setAttribute: () => {}
      }),
      getElementById: () => null
    },
    navigator: { userAgent: 'test' },
    localStorage: { getItem: () => null, setItem: () => {}, removeItem: () => {}, clear: () => {} },
    sessionStorage: { getItem: () => null, setItem: () => {}, removeItem: () => {}, clear: () => {} },
    location: { search: '', href: '' },
    google: { script: { run: {} } }
  };
  ctx.window.console = ctx.console;
  ctx.globalThis = ctx;
  ctx.self = ctx;

  vm.createContext(ctx);
  vm.runInContext(js, ctx, { filename: 'SharedUtilities.html' });

  return {
    Cache: vm.runInContext('Cache', ctx),
    DebounceManager: vm.runInContext('DebounceManager', ctx),
    ThrottleManager: vm.runInContext('ThrottleManager', ctx),
    SecurityUtilities: vm.runInContext('SecurityUtilities', ctx),
    ButtonBusyManager: vm.runInContext('ButtonBusyManager', ctx),
    timers,
    ctx
  };
}

function makeBtnMock() {
  const dataset = {};
  let textContent = '元の文字';
  const classes = new Set();
  return {
    disabled: false,
    get textContent() { return textContent; },
    set textContent(v) { textContent = v; },
    dataset,
    classList: {
      add: (c) => classes.add(c),
      remove: (c) => classes.delete(c),
      contains: (c) => classes.has(c),
      _has: classes
    }
  };
}

// =====================================================================
// Cache — LRU with TTL
// =====================================================================

test('Cache: set + get returns value within TTL', () => {
  const { Cache } = loadSharedUtilities();
  const c = new Cache();
  c.set('key', 'value');
  assert.equal(c.get('key'), 'value');
});

test('Cache: get returns null for missing key', () => {
  const { Cache } = loadSharedUtilities();
  const c = new Cache();
  assert.equal(c.get('nope'), null);
});

test('Cache: expired entry returns null and is removed', () => {
  const { Cache } = loadSharedUtilities();
  const c = new Cache(100, 1); // 1ms TTL
  c.set('k', 'v');
  // Wait for expiry — just busy-check with time elapsed
  const start = Date.now();
  while (Date.now() - start < 3) { /* tight wait to cross 1ms */ }
  assert.equal(c.get('k'), null);
  assert.equal(c.cache.has('k'), false, 'Expired entry must be purged from backing map');
});

test('Cache: set respects per-call TTL override', () => {
  const { Cache } = loadSharedUtilities();
  const c = new Cache(100, 1000);
  c.set('short', 'v', 1); // 1ms TTL
  c.set('long', 'v', 10000); // 10s TTL
  const start = Date.now();
  while (Date.now() - start < 3) { /* wait */ }
  assert.equal(c.get('short'), null);
  assert.equal(c.get('long'), 'v');
});

test('Cache: delete removes entry and all its metadata', () => {
  const { Cache } = loadSharedUtilities();
  const c = new Cache();
  c.set('k', 'v');
  c.delete('k');
  assert.equal(c.cache.has('k'), false);
  assert.equal(c.timestamps.has('k'), false);
  assert.equal(c.ttls.has('k'), false);
});

test('Cache: clear empties all internal maps', () => {
  const { Cache } = loadSharedUtilities();
  const c = new Cache();
  c.set('a', 1);
  c.set('b', 2);
  c.clear();
  assert.equal(c.cache.size, 0);
  assert.equal(c.timestamps.size, 0);
  assert.equal(c.ttls.size, 0);
});

test('Cache: has() mirrors get() liveness check', () => {
  const { Cache } = loadSharedUtilities();
  const c = new Cache();
  c.set('k', 'v');
  assert.equal(c.has('k'), true);
  c.delete('k');
  assert.equal(c.has('k'), false);
});

test('Cache: maxSize eviction removes oldest on overflow', () => {
  const { Cache } = loadSharedUtilities();
  const c = new Cache(3, 10000);
  c.set('a', 1);
  // Ensure monotonically increasing timestamps (some systems have ms-granular Date.now)
  const sleep = () => { const s = Date.now(); while (Date.now() === s) { /* spin */ } };
  sleep();
  c.set('b', 2);
  sleep();
  c.set('c', 3);
  sleep();
  // Adding 'd' should evict 'a' (oldest)
  c.set('d', 4);
  assert.equal(c.get('a'), null);
  assert.equal(c.get('b'), 2);
  assert.equal(c.get('c'), 3);
  assert.equal(c.get('d'), 4);
});

test('Cache: setting existing key does not trigger eviction', () => {
  const { Cache } = loadSharedUtilities();
  const c = new Cache(2, 10000);
  c.set('a', 1);
  c.set('b', 2);
  // Update 'a' — at max size, but key exists, so no eviction
  c.set('a', 99);
  assert.equal(c.get('a'), 99);
  assert.equal(c.get('b'), 2);
});

// =====================================================================
// SecurityUtilities.isValidEmail
// =====================================================================

test('SecurityUtilities.isValidEmail: accepts standard addresses', () => {
  const { SecurityUtilities } = loadSharedUtilities();
  const s = new SecurityUtilities();
  assert.equal(s.isValidEmail('user@example.com'), true);
  assert.equal(s.isValidEmail('first.last+tag@sub.example.co.jp'), true);
});

test('SecurityUtilities.isValidEmail: rejects malformed addresses', () => {
  const { SecurityUtilities } = loadSharedUtilities();
  const s = new SecurityUtilities();
  assert.equal(s.isValidEmail('no-at-sign'), false);
  assert.equal(s.isValidEmail('@example.com'), false);
  assert.equal(s.isValidEmail('user@'), false);
  assert.equal(s.isValidEmail('user @example.com'), false);
  assert.equal(s.isValidEmail(''), false);
});

// =====================================================================
// ButtonBusyManager — 全ページのボタン共通の処理中 UX
// =====================================================================

test('ButtonBusyManager.withBusy: disabled + busy class + 処理後復元', async () => {
  const { ButtonBusyManager } = loadSharedUtilities();
  const mgr = new ButtonBusyManager();
  const btn = makeBtnMock();
  let inside = null;
  await mgr.withBusy(btn, async () => {
    inside = { disabled: btn.disabled, busy: btn.classList.contains('btn-busy') };
    await Promise.resolve();
  });
  // 処理中: disabled + btn-busy class
  assert.equal(inside.disabled, true);
  assert.equal(inside.busy, true);
  // 処理後: 復元
  assert.equal(btn.disabled, false);
  assert.equal(btn.classList.contains('btn-busy'), false);
});

test('ButtonBusyManager.withBusy: 連打抑制 (実行中の重複 click は skip)', async () => {
  const { ButtonBusyManager } = loadSharedUtilities();
  const mgr = new ButtonBusyManager();
  const btn = makeBtnMock();
  let count = 0;
  // 1 回目: 50ms 待つ実行
  const p1 = mgr.withBusy(btn, async () => {
    count += 1;
    await new Promise(r => setTimeout(r, 50));
  });
  // すぐ 2 回目: 1 回目が完了する前 → skip されて undefined
  const p2 = mgr.withBusy(btn, async () => { count += 1; });
  const r2 = await p2;
  await p1;
  assert.equal(count, 1, '2 回目は実行されない');
  assert.equal(r2, undefined);
});

test('ButtonBusyManager.withBusy: busyText 指定でラベル一時置換 + 復元', async () => {
  const { ButtonBusyManager } = loadSharedUtilities();
  const mgr = new ButtonBusyManager();
  const btn = makeBtnMock();
  btn.textContent = '保存';
  let insideText = null;
  await mgr.withBusy(btn, async () => {
    insideText = btn.textContent;
  }, { busyText: '保存中…' });
  assert.equal(insideText, '保存中…');
  assert.equal(btn.textContent, '保存');
});

test('ButtonBusyManager.withBusy: 例外時も状態は復元される', async () => {
  const { ButtonBusyManager } = loadSharedUtilities();
  const mgr = new ButtonBusyManager();
  const btn = makeBtnMock();
  await assert.rejects(
    mgr.withBusy(btn, async () => { throw new Error('fail'); }),
    /fail/
  );
  assert.equal(btn.disabled, false);
  assert.equal(btn.classList.contains('btn-busy'), false);
});

test('ButtonBusyManager.navOnce: ms 内の連続呼出は false 返却 (タブ重複防止)', () => {
  const { ButtonBusyManager } = loadSharedUtilities();
  const mgr = new ButtonBusyManager();
  const btn = makeBtnMock();
  assert.equal(mgr.navOnce(btn, 500), true);
  assert.equal(mgr.navOnce(btn, 500), false, 'すぐの 2 回目は false');
});

test('ButtonBusyManager.navOnce: ms 経過後は再び true (gate 解除)', async () => {
  const { ButtonBusyManager } = loadSharedUtilities();
  const mgr = new ButtonBusyManager();
  const btn = makeBtnMock();
  assert.equal(mgr.navOnce(btn, 30), true);
  await new Promise(r => setTimeout(r, 40));
  assert.equal(mgr.navOnce(btn, 30), true, '30ms 後は通る');
});

test('ButtonBusyManager.withBusy: null btn の場合は asyncFn だけ実行', async () => {
  const { ButtonBusyManager } = loadSharedUtilities();
  const mgr = new ButtonBusyManager();
  let executed = false;
  const result = await mgr.withBusy(null, async () => { executed = true; return 42; });
  assert.equal(executed, true);
  assert.equal(result, 42);
});
