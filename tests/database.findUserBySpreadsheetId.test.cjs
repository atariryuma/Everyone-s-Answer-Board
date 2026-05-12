/**
 * findUserBySpreadsheetId — cache versioning と JSON parse 耐性のテスト
 *
 * Why: commit 52bdd495 で「version findUserBySpreadsheetId cache key」を導入。
 *      USER_CACHE_VERSION の bump が cache key を変更し、stale な cache を実質的に
 *      無効化する契約が race 防止の core。これが従来 untested だった (audit HIGH)。
 *      また configJson が壊れている user を skip して次へ進む挙動も pin する
 *      (実 prod で過去に発生)。
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

function loadCtx(overrides = {}) {
  const cache = overrides.cache || makeCache();
  const propsStore = overrides.propsStore || {};
  const users = overrides.users || [];

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
    Utilities: { sleep: () => {}, base64EncodeWebSafe: (s) => Buffer.from(s).toString('base64'), computeRsaSha256Signature: () => new Uint8Array([1, 2, 3, 4]) },
    UrlFetchApp: { fetch: () => { throw new Error('not stubbed'); } },
    Session: { getActiveUser: () => ({ getEmail: () => 'a@example.com' }) },
    validateEmail: (e) => typeof e === 'string' && /.+@.+/.test(e),
    getCurrentEmail: () => 'a@example.com',
    isAdministrator: () => false,
    getUserConfig: () => ({ success: true, config: {} }),
    executeWithRetry: (fn) => fn(),
    getCachedProperty: (k) => (k in propsStore ? propsStore[k] : null),
    clearPropertyCache: () => {},
    simpleHash: (s) => String(s).length,
    saveToCacheWithSizeCheck: (k, data) => { cache.put(k, JSON.stringify(data)); return true; },
    __cache: cache,
    __propsStore: propsStore
  };
  Object.assign(context, overrides);
  vm.createContext(context);
  DB_SCRIPT.runInContext(context);
  // Why post-load override: DatabaseCore.js は function getAllUsers() を内部で宣言しており、
  //   vm.runInContext 後にグローバルとして context に bind される。pre-load override は
  //   その宣言で上書きされてしまうため、post-load で書き換える必要がある。
  //   getAllUsers の参照は呼び出し時に scope chain を辿るので、後で代入したスタブも反映される。
  context.getAllUsers = overrides.getAllUsers || (() => users);
  return context;
}

function user(id, spreadsheetId, extraConfig = {}) {
  return {
    userId: id,
    userEmail: `${id}@example.com`,
    configJson: JSON.stringify({ spreadsheetId, ...extraConfig })
  };
}

// =====================================================================
// 入力バリデーション
// =====================================================================

test('findUserBySpreadsheetId: 非string / 空文字は null', () => {
  const ctx = loadCtx({ users: [user('u1', 'sheet-1')] });
  assert.equal(ctx.findUserBySpreadsheetId(null), null);
  assert.equal(ctx.findUserBySpreadsheetId(''), null);
  assert.equal(ctx.findUserBySpreadsheetId(undefined), null);
  assert.equal(ctx.findUserBySpreadsheetId(42), null);
});

// =====================================================================
// happy path: 一致する user を返し、cache に保存
// =====================================================================

test('findUserBySpreadsheetId: 一致 user を返す + cache に保存', () => {
  const cache = makeCache();
  const ctx = loadCtx({
    cache,
    users: [user('u1', 'sheet-A'), user('u2', 'sheet-B')]
  });
  const found = ctx.findUserBySpreadsheetId('sheet-B');
  assert.ok(found);
  assert.equal(found.userId, 'u2');

  // cache に user_by_sheet_v0_sheet-B として保存されている
  const cacheKey = 'user_by_sheet_v0_sheet-B';
  assert.ok(cache._store.has(cacheKey), 'should populate cache key');
  const cached = JSON.parse(cache._store.get(cacheKey));
  assert.equal(cached.userId, 'u2');
});

// =====================================================================
// 見つからない場合: null を返し、negative-result も短 TTL で cache
// =====================================================================

test('findUserBySpreadsheetId: 見つからない場合 null + negative-cache', () => {
  const cache = makeCache();
  const ctx = loadCtx({
    cache,
    users: [user('u1', 'sheet-A')]
  });
  const result = ctx.findUserBySpreadsheetId('sheet-NONEXISTENT');
  assert.equal(result, null);

  const cacheKey = 'user_by_sheet_v0_sheet-NONEXISTENT';
  assert.ok(cache._store.has(cacheKey));
  assert.equal(cache._store.get(cacheKey), JSON.stringify(null));
});

// =====================================================================
// cache hit: getAllUsers を呼ばずに cached value を返す
// =====================================================================

test('findUserBySpreadsheetId: cache hit ならば getAllUsers を呼ばない', () => {
  const cachedUser = { userId: 'u-cached', userEmail: 'cached@x.com' };
  const cache = makeCache({ 'user_by_sheet_v0_sheet-X': JSON.stringify(cachedUser) });
  let getAllCalled = 0;
  const ctx = loadCtx({
    cache,
    getAllUsers: () => { getAllCalled++; return []; }
  });

  const result = ctx.findUserBySpreadsheetId('sheet-X');
  assert.equal(result.userId, 'u-cached');
  assert.equal(getAllCalled, 0, 'cache hit must not hit DB');
});

// =====================================================================
// USER_CACHE_VERSION 切替: 古い cache key を実質的に無効化
// =====================================================================

test('findUserBySpreadsheetId: USER_CACHE_VERSION bump で古い cache が見えなくなる', () => {
  // 古い version の cache に「user-OLD」が入っているが、現 version は "7"
  const cache = makeCache({
    'user_by_sheet_v0_sheet-X': JSON.stringify({ userId: 'OLD', userEmail: 'old@x.com' })
  });
  const ctx = loadCtx({
    cache,
    propsStore: { USER_CACHE_VERSION: '7' },
    users: [user('NEW', 'sheet-X')]
  });

  const result = ctx.findUserBySpreadsheetId('sheet-X');
  assert.equal(result.userId, 'NEW', 'version bump must invalidate old cache');

  // 新しい version の cache key が書き込まれる
  assert.ok(cache._store.has('user_by_sheet_v7_sheet-X'));
});

test('findUserBySpreadsheetId: 同じ version (同じ key 名前空間) なら従来通り cache を信用', () => {
  const cache = makeCache({
    'user_by_sheet_v3_sheet-Y': JSON.stringify({ userId: 'CACHED', userEmail: 'c@x.com' })
  });
  let getAllCalled = 0;
  const ctx = loadCtx({
    cache,
    propsStore: { USER_CACHE_VERSION: '3' },
    getAllUsers: () => { getAllCalled++; return []; }
  });

  const result = ctx.findUserBySpreadsheetId('sheet-Y');
  assert.equal(result.userId, 'CACHED');
  assert.equal(getAllCalled, 0);
});

// =====================================================================
// skipCache: cache を無視して常に DB
// =====================================================================

test('findUserBySpreadsheetId: { skipCache: true } で cache 読み書きをバイパス', () => {
  const cache = makeCache({
    'user_by_sheet_v0_sheet-Z': JSON.stringify({ userId: 'CACHED', userEmail: 'c@x.com' })
  });
  let getAllCalled = 0;
  const ctx = loadCtx({
    cache,
    users: [user('FRESH', 'sheet-Z')],
    getAllUsers: () => { getAllCalled++; return [user('FRESH', 'sheet-Z')]; }
  });

  const result = ctx.findUserBySpreadsheetId('sheet-Z', { skipCache: true });
  assert.equal(result.userId, 'FRESH', 'skipCache must bypass cached value');
  assert.equal(getAllCalled, 1, 'skipCache must hit DB');

  // skipCache=true は cache を更新しない (旧キャッシュはそのまま残る)
  const stillCached = JSON.parse(cache._store.get('user_by_sheet_v0_sheet-Z'));
  assert.equal(stillCached.userId, 'CACHED', 'skipCache must not write back');
});

// =====================================================================
// 壊れた configJson の user は skip して次へ
// =====================================================================

test('findUserBySpreadsheetId: 壊れた configJson の user は skip', () => {
  const ctx = loadCtx({
    users: [
      { userId: 'corrupted', userEmail: 'c@x.com', configJson: '{not valid json' },
      user('healthy', 'sheet-OK')
    ]
  });
  const found = ctx.findUserBySpreadsheetId('sheet-OK');
  assert.ok(found, 'should not crash on corrupted neighbor');
  assert.equal(found.userId, 'healthy');
});

test('findUserBySpreadsheetId: configJson が undefined の user は default {} 扱い', () => {
  const ctx = loadCtx({
    users: [
      { userId: 'no-config', userEmail: 'n@x.com' },  // no configJson
      user('with-sheet', 'sheet-W')
    ]
  });
  const found = ctx.findUserBySpreadsheetId('sheet-W');
  assert.equal(found.userId, 'with-sheet');
});

// =====================================================================
// getAllUsers が失敗 / null を返した場合
// =====================================================================

test('findUserBySpreadsheetId: getAllUsers が非配列を返したら null', () => {
  const ctx = loadCtx({
    getAllUsers: () => null
  });
  assert.equal(ctx.findUserBySpreadsheetId('sheet-X'), null);
});

test('findUserBySpreadsheetId: getAllUsers が throw したら null を返す (全体 try/catch)', () => {
  const ctx = loadCtx({
    getAllUsers: () => { throw new Error('DB error'); }
  });
  assert.equal(ctx.findUserBySpreadsheetId('sheet-X'), null);
});

// =====================================================================
// cache write 失敗は warn ログのみ、結果はそのまま返す (best-effort)
// =====================================================================

test('findUserBySpreadsheetId: cache.put が throw しても見つけた user を返す', () => {
  const badCache = {
    get: () => null,
    put: () => { throw new Error('cache write failed'); },
    remove: () => {},
    _store: new Map()
  };
  const ctx = loadCtx({
    cache: badCache,
    users: [user('u1', 'sheet-A')]
  });
  const found = ctx.findUserBySpreadsheetId('sheet-A');
  assert.ok(found);
  assert.equal(found.userId, 'u1');
});
