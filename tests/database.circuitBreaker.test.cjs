/**
 * fetchSheetsAPIWithRetry circuit breaker tests
 *
 * Why: 那覇市版 GCP project (proud-will-462613-h7) は別 GAS アプリと同居しており、
 *      Sheets API の 429 storm に対する load-bearing なガード。閾値 (7/10/15 連続エラー)
 *      とロックアウト時間 (30s/60s/120s) の段階遷移、Cache への永続化、回復時のリセットは
 *      従来 audit で「全テストに欠落」と HIGH 指定されていた。本ファイルで pin する。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const DB_SOURCE = fs.readFileSync(path.resolve(__dirname, '../src/DatabaseCore.js'), 'utf8');
const DB_SCRIPT = new vm.Script(DB_SOURCE, { filename: 'DatabaseCore.js' });

function makeCache(initial = {}) {
  const store = new Map(Object.entries(initial));
  return {
    get: (k) => (store.has(k) ? store.get(k) : null),
    put: (k, v) => { store.set(k, v); },
    remove: (k) => { store.delete(k); },
    _store: store
  };
}

function makeResponse(code, body = '') {
  return { getResponseCode: () => code, getContentText: () => body };
}

function loadCtx(overrides = {}) {
  const cache = overrides.cache || makeCache();
  const sleeps = [];
  const fetchCalls = [];
  const propsStore = overrides.propsStore || {};
  const fetchSequence = overrides.fetchSequence || [];
  let fetchIdx = 0;

  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    CacheService: { getScriptCache: () => cache },
    CACHE_DURATION: { SHORT: 10, MEDIUM: 30, FORM_DATA: 30, LONG: 300, DATABASE_LONG: 600, USER_INDIVIDUAL: 900, EXTRA_LONG: 3600 },
    TIMEOUT_MS: { DEFAULT: 5000 },
    PropertiesService: { getScriptProperties: () => ({
      getProperty: (k) => (k in propsStore ? propsStore[k] : null),
      setProperty: (k, v) => { propsStore[k] = v; },
      getProperties: () => ({ ...propsStore })
    }) },
    SpreadsheetApp: { openById: () => { throw new Error('not stubbed'); } },
    Utilities: {
      sleep: (ms) => sleeps.push(ms),
      base64EncodeWebSafe: (s) => Buffer.from(s).toString('base64'),
      computeRsaSha256Signature: () => new Uint8Array([1, 2, 3, 4])
    },
    UrlFetchApp: {
      fetch: (url, opts) => {
        fetchCalls.push({ url, opts });
        if (fetchIdx >= fetchSequence.length) {
          return makeResponse(500, 'no more stubs');
        }
        const next = fetchSequence[fetchIdx++];
        if (next instanceof Error) throw next;
        return next;
      }
    },
    Session: { getActiveUser: () => ({ getEmail: () => 'a@example.com' }) },
    validateEmail: (e) => ({ isValid: typeof e === 'string' && /.+@.+/.test(e), sanitized: e, errors: [] }),
    getCurrentEmail: () => 'a@example.com',
    isAdministrator: () => false,
    getUserConfig: () => ({ success: true, config: {} }),
    executeWithRetry: overrides.executeWithRetry || ((fn, opts) => {
      // 実装と同じインターフェイスのまま fn を直接呼ぶ。
      // circuit-breaker テストでは fetchSheetsAPIWithRetry が executeWithRetry の中で fn を呼ぶ
      // 関係上、ここはリトライ無しで一発 fn() でよい(fnの内部で例外が出れば throw する)。
      // ただし source 内では 例外時のリトライを期待する分岐があるので、
      // テスト側でリトライしたい場合は overrides.executeWithRetry を差し込む。
      try { return fn(); } catch (e) { throw e; }
    }),
    getCachedProperty: (k) => (k in propsStore ? propsStore[k] : null),
    clearPropertyCache: () => {},
    simpleHash: (s) => String(s).length,
    saveToCacheWithSizeCheck: (k, data) => { cache.put(k, JSON.stringify(data)); return true; },
    __sleeps: sleeps,
    __fetchCalls: fetchCalls,
    __cache: cache
  };
  Object.assign(context, overrides);
  vm.createContext(context);
  DB_SCRIPT.runInContext(context);
  return context;
}

// =====================================================================
// happy path: 200 OK で circuit state はクリア
// =====================================================================

test('fetchSheetsAPIWithRetry: 200 OK 成功で consecutiveErrors を 1 だけ減衰させる (M3 decay)', () => {
  // Why (M3): 旧実装は success で hard-reset (=0) していたが、 部分 storm で並行する
  //   任意の 1 成功が global counter を 0 に戻す race で breaker が閾値 7 に到達できなかった。
  //   decay (= -1) なら net error pressure が蓄積する。
  const cache = makeCache({
    circuit_breaker_state: JSON.stringify({ consecutiveErrors: 5, circuitOpenUntil: 0 })
  });
  const ctx = loadCtx({
    cache,
    fetchSequence: [makeResponse(200, '{"ok":1}')]
  });

  const resp = ctx.fetchSheetsAPIWithRetry('https://x/api', { method: 'get' }, 'TestOp');
  assert.equal(resp.getResponseCode(), 200);

  const state = JSON.parse(cache._store.get('circuit_breaker_state'));
  assert.equal(state.consecutiveErrors, 4, 'success decays error counter by 1 (not hard reset)');
  assert.equal(state.circuitOpenUntil, 0);
});

test('fetchSheetsAPIWithRetry: success decay は 0 未満にならない (clamp)', () => {
  const cache = makeCache({
    circuit_breaker_state: JSON.stringify({ consecutiveErrors: 0, circuitOpenUntil: 0 })
  });
  const ctx = loadCtx({ cache, fetchSequence: [makeResponse(200, '{}')] });
  ctx.fetchSheetsAPIWithRetry('https://x/api', { method: 'get' }, 'TestOp');
  const state = JSON.parse(cache._store.get('circuit_breaker_state'));
  assert.equal(state.consecutiveErrors, 0, 'decay clamps at 0');
});

// =====================================================================
// H3: 非冪等 write (:append) は 5xx / network 断で retry させない (重複行防止)
// =====================================================================

test('fetchSheetsAPIWithRetry: idempotent=false + 5xx は "non-idempotent" マーカー付きで throw', () => {
  const ctx = loadCtx({
    fetchSequence: [makeResponse(500, 'backend error')],
    executeWithRetry: (fn) => fn()  // single attempt
  });
  assert.throws(
    () => ctx.fetchSheetsAPIWithRetry('https://x/api', { method: 'post' }, 'appendRow', 'sa@x', null, { idempotent: false }),
    /non-idempotent/
  );
});

test('fetchSheetsAPIWithRetry: idempotent=false + network throw も "non-idempotent" で throw', () => {
  const ctx = loadCtx({
    fetchSequence: [new Error('Connection reset')],
    executeWithRetry: (fn) => fn()
  });
  assert.throws(
    () => ctx.fetchSheetsAPIWithRetry('https://x/api', { method: 'post' }, 'appendRow', 'sa@x', null, { idempotent: false }),
    /non-idempotent/
  );
});

test('fetchSheetsAPIWithRetry: idempotent (default) + 5xx は通常の "API returned" エラー (retryable)', () => {
  const ctx = loadCtx({
    fetchSequence: [makeResponse(503, 'service unavailable')],
    executeWithRetry: (fn) => fn()
  });
  assert.throws(
    () => ctx.fetchSheetsAPIWithRetry('https://x/api', { method: 'get' }, 'getValues'),
    (err) => /API returned 503/.test(err.message) && !/non-idempotent/.test(err.message)
  );
});

test('fetchSheetsAPIWithRetry: idempotent=false でも 429 は retry 可能 (quota=未実行で安全)', () => {
  // 429 は確実に reject (=未実行) なので append でも retry してよい。 message に "non-idempotent"
  //   を含めず、 従来の "Quota exceeded (429)" を維持することで isRetryableError が retry を許す。
  const cache = makeCache();
  const ctx = loadCtx({
    cache,
    fetchSequence: [makeResponse(429, 'rate limited')],
    executeWithRetry: (fn) => fn()
  });
  assert.throws(
    () => ctx.fetchSheetsAPIWithRetry('https://x/api', { method: 'post' }, 'appendRow', 'sa@x', null, { idempotent: false }),
    (err) => /Quota exceeded \(429\)/.test(err.message) && !/non-idempotent/.test(err.message)
  );
});

// =====================================================================
// circuit open: 既に circuit が開いている時は即時 throw、fetch 自体しない
// =====================================================================

test('fetchSheetsAPIWithRetry: circuit がまだ open ならば fetch せず throw', () => {
  const futureUntil = Date.now() + 30000;
  const cache = makeCache({
    circuit_breaker_state: JSON.stringify({ consecutiveErrors: 7, circuitOpenUntil: futureUntil })
  });
  const ctx = loadCtx({ cache, fetchSequence: [makeResponse(200)] });

  assert.throws(
    () => ctx.fetchSheetsAPIWithRetry('https://x/api', { method: 'get' }, 'TestOp'),
    /Circuit breaker open/
  );
  assert.equal(ctx.__fetchCalls.length, 0, 'fetch must not be invoked while circuit is open');
});

test('fetchSheetsAPIWithRetry: 期限切れ circuit (circuitOpenUntil < now) は再開を許可', () => {
  const cache = makeCache({
    circuit_breaker_state: JSON.stringify({ consecutiveErrors: 7, circuitOpenUntil: Date.now() - 1000 })
  });
  const ctx = loadCtx({ cache, fetchSequence: [makeResponse(200)] });

  const resp = ctx.fetchSheetsAPIWithRetry('https://x/api', { method: 'get' }, 'TestOp');
  assert.equal(resp.getResponseCode(), 200);
  // 成功で counter は decay (7 → 6)。 circuitOpenUntil は 0 にクリアされ再開を許可。
  const state = JSON.parse(cache._store.get('circuit_breaker_state'));
  assert.equal(state.consecutiveErrors, 6);
  assert.equal(state.circuitOpenUntil, 0);
});

// =====================================================================
// 429 → consecutiveErrors インクリメント, 7 未満では circuitOpenUntil = 0 のまま
// =====================================================================

test('fetchSheetsAPIWithRetry: 429 で consecutiveErrors++ するが 7 未満は circuit は開かない', () => {
  const cache = makeCache({
    circuit_breaker_state: JSON.stringify({ consecutiveErrors: 3, circuitOpenUntil: 0 })
  });
  // 内側の retry は 1 回だけで諦めるようにする
  const ctx = loadCtx({
    cache,
    fetchSequence: [makeResponse(429, 'quota')],
    executeWithRetry: (fn) => fn()  // single attempt
  });

  assert.throws(
    () => ctx.fetchSheetsAPIWithRetry('https://x/api', {}, 'TestOp'),
    /Quota exceeded/
  );

  const state = JSON.parse(cache._store.get('circuit_breaker_state'));
  assert.equal(state.consecutiveErrors, 4, '3 → 4');
  assert.equal(state.circuitOpenUntil, 0, 'no lockout yet at <7');
});

// =====================================================================
// 段階的バックオフ: 7 → 30s, 10 → 60s, 15 → 120s
// =====================================================================

test('fetchSheetsAPIWithRetry: consecutiveErrors 7 到達で 30s ロックアウト', () => {
  const cache = makeCache({
    circuit_breaker_state: JSON.stringify({ consecutiveErrors: 6, circuitOpenUntil: 0 })
  });
  const t0 = Date.now();
  const ctx = loadCtx({
    cache,
    fetchSequence: [makeResponse(429)],
    executeWithRetry: (fn) => fn()
  });
  assert.throws(() => ctx.fetchSheetsAPIWithRetry('https://x', {}, 'op'), /Quota exceeded/);

  const state = JSON.parse(cache._store.get('circuit_breaker_state'));
  assert.equal(state.consecutiveErrors, 7);
  const lockoutDuration = state.circuitOpenUntil - t0;
  // 30000ms ± 数 ms 許容
  assert.ok(lockoutDuration >= 29000 && lockoutDuration <= 31000,
    `expected ~30000ms lockout, got ${lockoutDuration}ms`);
});

test('fetchSheetsAPIWithRetry: consecutiveErrors 10 到達で 60s ロックアウト', () => {
  const cache = makeCache({
    circuit_breaker_state: JSON.stringify({ consecutiveErrors: 9, circuitOpenUntil: 0 })
  });
  const t0 = Date.now();
  const ctx = loadCtx({
    cache,
    fetchSequence: [makeResponse(429)],
    executeWithRetry: (fn) => fn()
  });
  assert.throws(() => ctx.fetchSheetsAPIWithRetry('https://x', {}, 'op'), /Quota exceeded/);

  const state = JSON.parse(cache._store.get('circuit_breaker_state'));
  assert.equal(state.consecutiveErrors, 10);
  const lockoutDuration = state.circuitOpenUntil - t0;
  assert.ok(lockoutDuration >= 59000 && lockoutDuration <= 61000,
    `expected ~60000ms lockout, got ${lockoutDuration}ms`);
});

test('fetchSheetsAPIWithRetry: consecutiveErrors 15 到達で 120s ロックアウト', () => {
  const cache = makeCache({
    circuit_breaker_state: JSON.stringify({ consecutiveErrors: 14, circuitOpenUntil: 0 })
  });
  const t0 = Date.now();
  const ctx = loadCtx({
    cache,
    fetchSequence: [makeResponse(429)],
    executeWithRetry: (fn) => fn()
  });
  assert.throws(() => ctx.fetchSheetsAPIWithRetry('https://x', {}, 'op'), /Quota exceeded/);

  const state = JSON.parse(cache._store.get('circuit_breaker_state'));
  assert.equal(state.consecutiveErrors, 15);
  const lockoutDuration = state.circuitOpenUntil - t0;
  assert.ok(lockoutDuration >= 119000 && lockoutDuration <= 121000,
    `expected ~120000ms lockout, got ${lockoutDuration}ms`);
});

// =====================================================================
// non-200 / non-429: error throw だが circuit には触れない
// =====================================================================

test('fetchSheetsAPIWithRetry: 500 では throw するが consecutiveErrors は増やさない', () => {
  // Why: 500 は server-side エラーで quota とは別。circuit breaker は quota 専用。
  const cache = makeCache({
    circuit_breaker_state: JSON.stringify({ consecutiveErrors: 2, circuitOpenUntil: 0 })
  });
  const ctx = loadCtx({
    cache,
    fetchSequence: [makeResponse(500, 'server error')],
    executeWithRetry: (fn) => fn()
  });
  assert.throws(
    () => ctx.fetchSheetsAPIWithRetry('https://x', {}, 'op'),
    /API returned 500/
  );

  // cache はそのまま (counter 増えない)
  const state = JSON.parse(cache._store.get('circuit_breaker_state'));
  assert.equal(state.consecutiveErrors, 2);
});

// =====================================================================
// キャッシュ未初期化 (fresh state) — デフォルトは consecutiveErrors=0
// =====================================================================

test('fetchSheetsAPIWithRetry: 初回呼び出し時 (cache 未設定) はデフォルト state でスタート', () => {
  const cache = makeCache(); // empty
  const ctx = loadCtx({ cache, fetchSequence: [makeResponse(200)] });
  ctx.fetchSheetsAPIWithRetry('https://x', {}, 'op');

  const state = JSON.parse(cache._store.get('circuit_breaker_state'));
  assert.equal(state.consecutiveErrors, 0);
  assert.equal(state.circuitOpenUntil, 0);
});

// =====================================================================
// 429 後の backoff: ms 値は Math.min(15000 + 15000*retry, 60000)
// =====================================================================

test('fetchSheetsAPIWithRetry: 429 時の Utilities.sleep は 15s 〜 60s の範囲で adaptive', () => {
  const cache = makeCache();
  const ctx = loadCtx({
    cache,
    fetchSequence: [makeResponse(429)],
    executeWithRetry: (fn) => fn()
  });
  assert.throws(() => ctx.fetchSheetsAPIWithRetry('https://x', {}, 'op'), /Quota exceeded/);

  // 初回の sleep は 15000ms (retry=0)
  assert.equal(ctx.__sleeps.length, 1);
  assert.equal(ctx.__sleeps[0], 15000);
});
