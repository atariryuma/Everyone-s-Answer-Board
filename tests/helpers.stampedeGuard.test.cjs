/**
 * withStampedeGuard_ (helpers.js): cache stampede 防止の contract を pin する。
 *
 * Why: 2026-09-08 の 429 storm は「version bump の直後に 30 人が同時に miss して全員が
 *   Sheets API を読む」構造だった。1 件だけが読む / 残りは直前の結果か短い待ちで済ます、
 *   という約束が崩れると同じ事故が再発するので、ここで固定する。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const SOURCE = fs.readFileSync(path.resolve(__dirname, '../src/helpers.js'), 'utf8');

function loadContext(overrides = {}) {
  const store = new Map();
  const ops = [];
  const sleeps = [];
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {} }) },
    CacheService: {
      getScriptCache: () => ({
        get: (k) => { ops.push('get:' + k); return store.has(k) ? store.get(k) : null; },
        put: (k, v) => { ops.push('put:' + k); store.set(k, v); },
        remove: (k) => { ops.push('remove:' + k); store.delete(k); }
      })
    },
    Utilities: {
      sleep: (ms) => { sleeps.push(ms); if (overrides.onSleep) overrides.onSleep(store, ms); }
    },
    ...overrides
  };
  vm.createContext(context);
  vm.runInContext(SOURCE, context, { filename: 'helpers.js' });
  return { context, store, ops, sleeps };
}

test('withStampedeGuard_: cache hit なら loader を呼ばない', () => {
  const { context, store } = loadContext();
  store.set('k1', JSON.stringify({ v: 1 }));
  let loads = 0;
  const out = context.withStampedeGuard_({ key: 'k1', ttl: 10, loader: () => { loads++; return { v: 2 }; } });
  assert.equal(out.v, 1);
  assert.equal(loads, 0);
});

test('withStampedeGuard_: miss で claim した 1 件が読み、key と latest を書き、flight を外す', () => {
  const { context, store, ops } = loadContext();
  let loads = 0;
  const out = context.withStampedeGuard_({
    key: 'k1', ttl: 10, latestKey: 'k1:latest', loader: () => { loads++; return { v: 2 }; }
  });
  assert.equal(out.v, 2);
  assert.equal(loads, 1);
  assert.equal(JSON.parse(store.get('k1')).v, 2);
  assert.equal(JSON.parse(store.get('k1:latest')).v, 2);
  assert.ok(ops.includes('put:k1:flight'), 'flight を立てる');
  assert.ok(ops.includes('remove:k1:flight'), '読み終えたら flight を外す');
  assert.equal(store.has('k1:flight'), false);
});

test('withStampedeGuard_: 他が読んでいる間 (flight 中) は allowStale なら直前の結果を返し、読まない', () => {
  const { context, store, sleeps } = loadContext();
  store.set('k1:flight', '1');
  store.set('k1:latest', JSON.stringify({ v: 'stale' }));
  let loads = 0;
  const out = context.withStampedeGuard_({
    key: 'k1', ttl: 10, latestKey: 'k1:latest', allowStale: true, waitMs: 100,
    loader: () => { loads++; return { v: 'fresh' }; }
  });
  assert.equal(out.v, 'stale');
  assert.equal(loads, 0);
  assert.equal(sleeps.length, 0, 'stale があるなら待たない');
});

test('withStampedeGuard_: noStale (allowStale=false) は flight 中でも直前の結果を返さず、埋まるのを待つ', () => {
  const { context, store, sleeps } = loadContext({
    // 1 回目の sleep の間に「読んでいた誰か」が key を埋める
    onSleep: (s) => { s.set('k1', JSON.stringify({ v: 'filled' })); }
  });
  store.set('k1:flight', '1');
  store.set('k1:latest', JSON.stringify({ v: 'stale' }));
  let loads = 0;
  const out = context.withStampedeGuard_({
    key: 'k1', ttl: 10, latestKey: 'k1:latest', allowStale: false, waitMs: 100, waitTries: 3,
    loader: () => { loads++; return { v: 'fresh' }; }
  });
  assert.equal(out.v, 'filled');
  assert.equal(loads, 0);
  assert.deepEqual(sleeps, [100]);
});

test('withStampedeGuard_: 待っても埋まらなければ自分で読む (止まるより読む)', () => {
  const { context, store, sleeps } = loadContext();
  store.set('k1:flight', '1');
  let loads = 0;
  const out = context.withStampedeGuard_({
    key: 'k1', ttl: 10, waitMs: 50, waitTries: 2, loader: () => { loads++; return { v: 'mine' }; }
  });
  assert.equal(out.v, 'mine');
  assert.equal(loads, 1);
  assert.deepEqual(sleeps, [50, 50]);
  assert.equal(store.has('k1:flight'), true, '自分が立てた flight ではないので外さない');
});

test('withStampedeGuard_: loader が落ちても flight は外れる (次の read が永久に待たない)', () => {
  const { context, store } = loadContext();
  assert.throws(() => context.withStampedeGuard_({ key: 'k1', ttl: 10, loader: () => { throw new Error('boom'); } }), /boom/);
  assert.equal(store.has('k1:flight'), false);
});

test('withStampedeGuard_: isCacheable が false の結果は cache しない (失敗応答を固定しない)', () => {
  const { context, store } = loadContext();
  const out = context.withStampedeGuard_({
    key: 'k1', ttl: 10, latestKey: 'k1:latest',
    isCacheable: (v) => Boolean(v && v.success),
    loader: () => ({ success: false })
  });
  assert.equal(out.success, false);
  assert.equal(store.has('k1'), false);
  assert.equal(store.has('k1:latest'), false);
});

test('withStampedeGuard_: CacheService が無ければそのまま loader', () => {
  const { context } = loadContext({ CacheService: undefined });
  assert.equal(context.withStampedeGuard_({ key: 'k', ttl: 1, loader: () => ({ v: 1 }) }).v, 1);
});
