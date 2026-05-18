// SharedUtilities.html 内 window.themeManager のテスト。
// HTML から script 部分を抜き出して vm sandbox で実行し、 themeManager の挙動を検証する。

const test = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function setupSandbox({ initialTheme = null, prefersDark = true, hasBody = true } = {}) {
  const storage = {};
  if (initialTheme !== null) storage['app-theme'] = initialTheme;
  const subscribers = new Map();
  const listeners = [];

  const documentEl = {
    body: hasBody ? makeEl() : null,
    documentElement: makeEl(),
    readyState: 'complete',
    addEventListener: (ev, cb) => { listeners.push({ ev, cb }); },
    getElementById: () => null,
    createElement: () => makeEl()
  };

  function makeEl() {
    const classes = new Set();
    return {
      classList: {
        add: (...c) => c.forEach(x => classes.add(x)),
        remove: (...c) => c.forEach(x => classes.delete(x)),
        contains: (c) => classes.has(c),
        toString: () => Array.from(classes).join(' ')
      },
      setAttribute: function (k, v) { this[k] = v; },
      getAttribute: function (k) { return this[k]; },
      _classes: classes
    };
  }

  const ctx = {
    window: null,
    document: documentEl,
    localStorage: {
      getItem: (k) => (storage[k] === undefined ? null : storage[k]),
      setItem: (k, v) => { storage[k] = v; },
      removeItem: (k) => { delete storage[k]; }
    },
    matchMedia: (q) => ({
      matches: q.includes('prefers-color-scheme: dark') ? prefersDark : !prefersDark,
      addEventListener: (ev, cb) => subscribers.set(ev, cb),
      removeEventListener: () => {},
      addListener: () => {},
      removeListener: () => {}
    }),
    console: console,
    setTimeout: setTimeout,
    clearTimeout: clearTimeout,
    Promise: Promise,
    Date: Date,
    Math: Math,
    requestIdleCallback: (cb) => setTimeout(cb, 0),
    requestAnimationFrame: (cb) => setTimeout(cb, 0)
  };
  ctx.window = ctx;
  vm.createContext(ctx);

  // SharedUtilities.html を読込、 <script> ブロックだけ抽出して themeManager 部分を実行
  const htmlPath = path.resolve(__dirname, '../src/SharedUtilities.html');
  const html = fs.readFileSync(htmlPath, 'utf8');
  // themeManager の IIFE だけ抽出
  const startMarker = '(function setupThemeManager() {';
  const startIdx = html.indexOf(startMarker);
  if (startIdx < 0) throw new Error('setupThemeManager IIFE not found');
  // 関数の終端を中括弧マッチでスキャン (IIFE は `})();` で終わる)
  let depth = 0;
  let endIdx = -1;
  for (let i = startIdx; i < html.length; i++) {
    if (html[i] === '{') depth++;
    else if (html[i] === '}') {
      depth--;
      if (depth === 0) {
        // 後続の `)();` をスキップ
        endIdx = html.indexOf(';', i) + 1;
        break;
      }
    }
  }
  if (endIdx < 0) throw new Error('setupThemeManager IIFE end not found');

  const iifeCode = html.slice(startIdx, endIdx);
  vm.runInContext(iifeCode, ctx, { filename: 'SharedUtilities.themeManager.js' });

  return { ctx, storage, mqlSubscribers: subscribers };
}

test('themeManager: default は dark (localStorage 未設定時)', () => {
  const { ctx } = setupSandbox({ initialTheme: null });
  assert.strictEqual(ctx.window.themeManager.get(), 'dark');
  assert.strictEqual(ctx.window.themeManager.resolved(), 'dark');
});

test('themeManager: set("light") で localStorage と body class が更新される', () => {
  const { ctx, storage } = setupSandbox();
  ctx.window.themeManager.set('light');
  assert.strictEqual(storage['app-theme'], 'light');
  assert.ok(ctx.document.body._classes.has('theme-light'));
  assert.ok(!ctx.document.body._classes.has('theme-dark'));
  assert.strictEqual(ctx.document.documentElement['data-theme'], 'light');
});

test('themeManager: set("dark") で body class が theme-dark になる', () => {
  const { ctx } = setupSandbox({ initialTheme: 'light' });
  ctx.window.themeManager.set('dark');
  assert.ok(ctx.document.body._classes.has('theme-dark'));
  assert.ok(!ctx.document.body._classes.has('theme-light'));
});

test('themeManager: toggle() で dark ↔ light 反転', () => {
  const { ctx } = setupSandbox({ initialTheme: 'dark' });
  ctx.window.themeManager.set('dark');
  ctx.window.themeManager.toggle();
  assert.strictEqual(ctx.window.themeManager.get(), 'light');
  ctx.window.themeManager.toggle();
  assert.strictEqual(ctx.window.themeManager.get(), 'dark');
});

test('themeManager: auto で OS prefers-color-scheme に追従 (dark)', () => {
  const { ctx } = setupSandbox({ prefersDark: true });
  ctx.window.themeManager.set('auto');
  assert.strictEqual(ctx.window.themeManager.get(), 'auto');
  assert.strictEqual(ctx.window.themeManager.resolved(), 'dark');
  assert.ok(ctx.document.body._classes.has('theme-dark'));
});

test('themeManager: auto で OS prefers-color-scheme に追従 (light)', () => {
  const { ctx } = setupSandbox({ prefersDark: false });
  ctx.window.themeManager.set('auto');
  assert.strictEqual(ctx.window.themeManager.resolved(), 'light');
  assert.ok(ctx.document.body._classes.has('theme-light'));
});

test('themeManager: subscribe コールバックが set/toggle で呼ばれる', () => {
  const { ctx } = setupSandbox();
  const calls = [];
  const unsub = ctx.window.themeManager.subscribe((resolved, raw) => {
    calls.push({ resolved, raw });
  });
  ctx.window.themeManager.set('light');
  ctx.window.themeManager.set('dark');
  assert.strictEqual(calls.length, 2);
  assert.strictEqual(calls[0].resolved, 'light');
  assert.strictEqual(calls[1].resolved, 'dark');
  unsub();
  ctx.window.themeManager.set('light');
  assert.strictEqual(calls.length, 2);  // unsub 後は呼ばれない
});

test('themeManager: 不正な値は dark に正規化', () => {
  const { ctx } = setupSandbox();
  ctx.window.themeManager.set('invalid-theme');
  assert.strictEqual(ctx.window.themeManager.get(), 'dark');
});

test('themeManager: localStorage が壊れていても dark default に落ちる', () => {
  const { ctx } = setupSandbox({ initialTheme: 'corrupted-value' });
  // 'corrupted-value' は VALID に含まれない → dark
  assert.strictEqual(ctx.window.themeManager.get(), 'dark');
});
