/**
 * 2026-10 の同時アクセス監査で直した挙動を固定する。
 *
 *  - SA proxy の読み失敗 ([] が返る) を「空のボード」と混同して成功 cache しない (DataService)
 *  - 共有 cache の reacted を閲覧者ごとに付け直す (ReactionService.applyViewerReactions_)
 *  - key 単位の cache lock (helpers.acquireCacheLock_)
 *  - createUser_ は失敗で error envelope を返さず null、存在確認の前に version memo を捨てる
 *  - escapeHtml は " と ' もエスケープする (SharedUtilities)
 *  - polling の世代番号 (page.js): stop → start の間に終わった poll がチェーンを増やさない
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { gasResponseStubs } = require('./_helpers.cjs');

const SRC = (name) => fs.readFileSync(path.resolve(__dirname, '../src/' + name), 'utf8');

function makeCache() {
  const store = new Map();
  return {
    get: (k) => (store.has(k) ? store.get(k) : null),
    put: (k, v) => { store.set(k, v); },
    remove: (k) => { store.delete(k); },
    _store: store
  };
}

// =====================================================================
// DataService: proxy の読み失敗は BUSY (cache されない)
// =====================================================================

function loadDataService(overrides = {}) {
  const cache = makeCache();
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    CacheService: { getScriptCache: () => cache },
    CACHE_DURATION: { DATABASE_LONG: 600, FORM_DATA: 30 },
    Utilities: { sleep: () => {} },
    ...gasResponseStubs(),
    getCurrentEmail: () => 'viewer@example.com',
    findUserById: () => null,
    openSpreadsheet: () => null,
    getConfigOrDefault: () => ({}),
    getQuestionText: () => 'Q',
    normalizeHeader: (h) => String(h || '').toLowerCase().trim(),
    resolveColumnIndex: (headers, field) => ({ index: headers.indexOf(field) }),
    extractReactions: () => ({}),
    extractHighlight: () => false,
    formatTimestamp: (v) => String(v || ''),
    ...overrides
  };
  vm.createContext(context);
  vm.runInContext(SRC('DataService.js'), context, { filename: 'DataService.js' });
  context._cache = cache;
  return context;
}

// SA proxy と同じ契約: 失敗時は [] を返し、失敗 slot にメモを残す。
function makeProxyLikeSheet({ headers, rows, failRows = false, failHeaders = false }, ctx) {
  return {
    getName: () => 'Sheet1',
    getParent: () => ({ getId: () => 'ss-1' }),
    getLastRow: () => 1000,                   // グリッド行数 (データ行数ではない)
    getLastColumn: () => headers.length || 26,
    getRange: (row, col, numRows, numCols) => ({
      getValues: () => {
        if (row === 1 && numRows === 1) {
          if (failHeaders) { ctx.noteProxyReadFailure_('getRange', new Error('Quota exceeded (429)')); return []; }
          return [headers.slice()];
        }
        if (failRows) { ctx.noteProxyReadFailure_('getRange', new Error('Quota exceeded (429)')); return []; }
        return rows.map((r) => r.slice());
      }
    })
  };
}

test('fetchSpreadsheetData: 見出しが読めない (429) ときも BUSY', () => {
  const ctx = loadDataService();
  let slot = null;
  ctx.noteProxyReadFailure_ = (label, err) => { slot = { label, message: err.message }; };
  ctx.takeProxyReadFailure_ = () => { const f = slot; slot = null; return f; };
  const sheet = makeProxyLikeSheet({ headers: [], rows: [], failHeaders: true }, ctx);
  ctx.openSpreadsheet = () => ({ spreadsheet: { getSheetByName: () => sheet } });
  const res = ctx.fetchSpreadsheetData({ spreadsheetId: 'ss-1', sheetName: 'Sheet1' }, {});
  assert.equal(res.success, false);
  assert.equal(res.busy, true);
});

test('fetchSpreadsheetData: 本当に空のボード (見出しだけ) は success:true, data:[]', () => {
  const ctx = loadDataService();
  ctx.noteProxyReadFailure_ = () => {};
  ctx.takeProxyReadFailure_ = () => null;
  const sheet = makeProxyLikeSheet({ headers: ['timestamp', 'answer'], rows: [] }, ctx);
  ctx.openSpreadsheet = () => ({ spreadsheet: { getSheetByName: () => sheet } });
  const res = ctx.fetchSpreadsheetData({ spreadsheetId: 'ss-1', sheetName: 'Sheet1' }, {});
  assert.equal(res.success, true);
  assert.deepEqual(JSON.parse(JSON.stringify(res.data)), []);
});

test('fetchSpreadsheetData: 正常に読めたら従来どおり行を返す', () => {
  const ctx = loadDataService();
  ctx.noteProxyReadFailure_ = () => {};
  ctx.takeProxyReadFailure_ = () => null;
  const sheet = makeProxyLikeSheet({ headers: ['timestamp', 'answer'], rows: [['t1', 'a'], ['t2', 'b']] }, ctx);
  ctx.openSpreadsheet = () => ({ spreadsheet: { getSheetByName: () => sheet } });
  const res = ctx.fetchSpreadsheetData({ spreadsheetId: 'ss-1', sheetName: 'Sheet1', columnMapping: { answer: 1 } }, {});
  assert.equal(res.success, true);
  assert.equal(res.data.length, 2);
});

// =====================================================================
// ReactionService: 共有 cache の reacted を閲覧者ごとに付け直す
// =====================================================================

function loadReactionService() {
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    ...gasResponseStubs(),
    CacheService: { getScriptCache: () => makeCache() },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
    Utilities: { sleep: () => {} },
    getCurrentEmail: () => 'a@example.com',
    isAdministrator: () => false,
    findPublishedBoardOwner: () => null,
    getConfigOrDefault: () => ({}),
    openSpreadsheet: () => null,
    resolveReactionMode: () => 'on'
  };
  vm.createContext(context);
  vm.runInContext(SRC('ReactionService.js'), context, { filename: 'ReactionService.js' });
  return context;
}

test('extractReactions(includeVoters): voters に正規化済メールを残し、reacted は読んだ人基準', () => {
  const ctx = loadReactionService();
  const headers = ['Q', 'UNDERSTAND', 'LIKE', 'CURIOUS'];
  const row = ['ans', '', 'A@Example.com|b@example.com', ''];
  const r = ctx.extractReactions(row, headers, 'a@example.com', null, { includeVoters: true });
  assert.equal(r.LIKE.count, 2);
  assert.equal(r.LIKE.reacted, true);
  assert.deepEqual(JSON.parse(JSON.stringify(r.LIKE.voters)), ['a@example.com', 'b@example.com']);
  // 既定 (includeVoters 無し) は voters を持たない (wire 互換)
  const plain = ctx.extractReactions(row, headers, 'a@example.com');
  assert.equal(plain.LIKE.voters, undefined);
});

test('applyViewerReactions_: 最初に読んだ児童の reacted が他の児童に漏れない (voters から付け直し、voters は落とす)', () => {
  const ctx = loadReactionService();
  const cachedRows = [{
    rowIndex: 2,
    reactions: {
      UNDERSTAND: { count: 0, reacted: false, voters: [] },
      LIKE: { count: 2, reacted: true, voters: ['a@example.com', 'b@example.com'] },   // 読んだのは a
      CURIOUS: { count: 1, reacted: false, voters: ['c@example.com'] }
    }
  }];
  const forC = ctx.applyViewerReactions_(cachedRows, 'C@Example.com');
  assert.equal(forC[0].reactions.LIKE.reacted, false, 'c は LIKE していない');
  assert.equal(forC[0].reactions.CURIOUS.reacted, true, 'c は CURIOUS している');
  assert.equal(forC[0].reactions.LIKE.count, 2);
  assert.equal(forC[0].reactions.LIKE.voters, undefined, 'メール一覧は wire に載せない');
  // 元の cache 行は mutate しない
  assert.ok(Array.isArray(cachedRows[0].reactions.LIKE.voters));
  // voters を持たない行 (旧 cache / 直読み) はそのまま
  const legacy = [{ rowIndex: 3, reactions: { LIKE: { count: 1, reacted: true } } }];
  assert.equal(ctx.applyViewerReactions_(legacy, 'x@example.com')[0].reactions.LIKE.reacted, true);
});

test('processReactionDirect: 隣り合う 2 セルが変わる切替 (LIKE → CURIOUS) は 1 回の setValues', () => {
  const ctx = loadReactionService();
  const data = [['Q', 'UNDERSTAND', 'LIKE', 'CURIOUS'], ['ans', '', 'a@example.com', '']];
  const writes = [];
  const sheet = {
    getName: () => 'S', getParent: () => ({ getId: () => 'ss' }),
    getLastColumn: () => 4,
    getRange: (row, col, numRows = 1, numCols = 1) => ({
      getValues: () => [data[row - 1].slice(col - 1, col - 1 + numCols)],
      setValue: (v) => { writes.push({ row, col, v }); data[row - 1][col - 1] = v; },
      setValues: (vs) => { writes.push({ row, col, vs }); vs[0].forEach((v, i) => { data[row - 1][col - 1 + i] = v; }); }
    })
  };
  const r = ctx.processReactionDirect_(sheet, 2, 'CURIOUS', 'a@example.com', data[0]);
  assert.equal(r.action, 'added');
  assert.equal(writes.length, 1, 'LIKE (3 列目) と CURIOUS (4 列目) は隣なので 1 回で書く');
  assert.equal(data[1][2], '');
  assert.equal(data[1][3], 'a@example.com');
});

// =====================================================================
// helpers: key 単位の cache lock
// =====================================================================

function loadHelpers(lockBehavior) {
  const cache = makeCache();
  let held = false;
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    PROPERTY_CACHE_TTL: 30000,
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {} }) },
    CacheService: { getScriptCache: () => cache },
    LockService: { getScriptLock: () => ({
      tryLock: () => (lockBehavior === 'busy' ? false : (held ? false : (held = true))),
      releaseLock: () => { held = false; }
    }) },
    Utilities: { sleep: () => {} }
  };
  vm.createContext(context);
  vm.runInContext(SRC('helpers.js'), context, { filename: 'helpers.js' });
  context._cache = cache;
  return context;
}

test('acquireCacheLock_: 取得 → 同じ key は取れない → release で取れる', () => {
  const ctx = loadHelpers();
  const t1 = ctx.acquireCacheLock_('users_row_lock:u1', { ttlSec: 30 });
  assert.ok(t1 && t1 !== 'nolock');
  assert.equal(ctx.acquireCacheLock_('users_row_lock:u1', { ttlSec: 30 }), null, '保持中は取れない');
  // 別の key は独立
  assert.ok(ctx.acquireCacheLock_('users_row_lock:u2', { ttlSec: 30 }));
  ctx.releaseCacheLock_('users_row_lock:u1', 'wrong-token');
  assert.equal(ctx.acquireCacheLock_('users_row_lock:u1'), null, '他人の token では外れない');
  ctx.releaseCacheLock_('users_row_lock:u1', t1);
  assert.ok(ctx.acquireCacheLock_('users_row_lock:u1'), 'release 後は取れる');
});

test('acquireCacheLock_: blockedBy の key が立っている間は取らない', () => {
  const ctx = loadHelpers();
  ctx._cache.put('users_struct_lock', 'x', 60);
  assert.equal(ctx.acquireCacheLock_('users_row_lock:u1', { blockedBy: ['users_struct_lock'] }), null);
});

test('acquireCacheLock_: ScriptLock が取れない (混雑) なら待たない呼び出しは null', () => {
  const ctx = loadHelpers('busy');
  assert.equal(ctx.acquireCacheLock_('k', { waitMs: 0 }), null);
});

test('withStampedeGuard_: 直前の結果 (latest) を返すときは stale:true を付ける', () => {
  const ctx = loadHelpers();
  ctx._cache.put('k:flight', '1', 5);
  ctx._cache.put('k:latest', JSON.stringify({ success: true, data: [1] }), 60);
  const v = ctx.withStampedeGuard_({ key: 'k', ttl: 10, loader: () => ({ success: true, data: [2] }), latestKey: 'k:latest', allowStale: true });
  assert.equal(v.stale, true);
  assert.deepEqual(JSON.parse(JSON.stringify(v.data)), [1]);
});

test('withStampedeGuard_: key を書いてから flight を外す (待っている側が「無い」瞬間を見ない)', () => {
  const ctx = loadHelpers();
  const order = [];
  const origPut = ctx._cache.put; const origRemove = ctx._cache.remove;
  ctx._cache.put = (k, v, t) => { order.push('put:' + k); return origPut(k, v, t); };
  ctx._cache.remove = (k) => { order.push('remove:' + k); return origRemove(k); };
  ctx.withStampedeGuard_({ key: 'k', ttl: 10, loader: () => ({ ok: true }), flightKey: 'k:flight' });
  assert.ok(order.indexOf('put:k') < order.indexOf('remove:k:flight'));
});

// =====================================================================
// DatabaseCore.createUser_: null 契約 + version memo の破棄
// =====================================================================

test('createUser: users シートが読めない (429 で []) ときは error envelope ではなく null', () => {
  const cache = makeCache();
  const cleared = [];
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    ...gasResponseStubs(),
    CacheService: { getScriptCache: () => cache },
    CACHE_DURATION: { SHORT: 10, MEDIUM: 30, LONG: 300, DATABASE_LONG: 600, USER_INDIVIDUAL: 900 },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {} }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
    Utilities: { getUuid: () => 'uuid-1', sleep: () => {} },
    SpreadsheetApp: { openById: () => { throw new Error('not stubbed'); } },
    validateEmail: (e) => ({ isValid: /.+@.+/.test(String(e)) }),
    getCurrentEmail: () => 'new@example.com',
    isAdministrator: () => false,
    getUserConfig: () => ({ success: true, config: {} }),
    executeWithRetry: (fn) => fn(),
    getCachedProperty_: () => null,
    clearPropertyCache_: (k) => cleared.push(k),
    simpleHash: () => 'h',
    saveToCacheWithSizeCheck: () => true,
    DEFAULT_DISPLAY_SETTINGS: {},
    withStampedeGuard_: (o) => o.loader(),
    acquireCacheLock_: () => 'tok',
    releaseCacheLock_: () => {}
  };
  vm.createContext(context);
  vm.runInContext(SRC('DatabaseCore.js'), context, { filename: 'DatabaseCore.js' });
  // users シートは読めない ([] が返る)
  const sheet = { getName: () => 'users', getDataRange: () => ({ getValues: () => [] }), getLastRow: () => 1000, getLastColumn: () => 7, getRange: () => ({ getValues: () => [] }) };
  context.openDatabase = () => ({ getSheetByName: () => sheet });
  const user = context.createUser_('new@example.com', {}, {});
  assert.equal(user, null, '呼び出し側の `if (!user)` で拾える');
  assert.ok(cleared.includes('USER_CACHE_VERSION'), '存在確認の前に version memo を捨てる (別インスタンスの作成を見逃さない)');
});

// =====================================================================
// SharedUtilities.escapeHtml: 属性に入れても抜けられない
// =====================================================================

test('escapeHtml: & < > " \' の 5 文字をすべてエスケープする', () => {
  const src = SRC('SharedUtilities.html');
  const js = src.match(/<script[^>]*>([\s\S]*?)<\/script>/)[1];
  const ctx = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    Map, Set, Date, Math, JSON, Number, Array, Object, Promise, Error, Symbol,
    setTimeout: () => 1, clearTimeout: () => {},
    window: { addEventListener: () => {}, removeEventListener: () => {}, console: { warn: () => {}, error: () => {} } },
    document: {
      addEventListener: () => {}, removeEventListener: () => {},
      createElement: () => ({ classList: { add: () => {}, remove: () => {} }, textContent: '', innerHTML: '', setAttribute: () => {}, appendChild: () => {}, style: {} }),
      getElementById: () => null, querySelector: () => null, querySelectorAll: () => [],
      body: { appendChild: () => {}, classList: { add: () => {}, remove: () => {} } },
      readyState: 'complete'
    },
    navigator: { userAgent: '' },
    localStorage: { getItem: () => null, setItem: () => {}, removeItem: () => {} },
    sessionStorage: { getItem: () => null, setItem: () => {}, removeItem: () => {} }
  };
  ctx.window.document = ctx.document;
  ctx.globalThis = ctx;
  vm.createContext(ctx);
  try { vm.runInContext(js, ctx, { filename: 'SharedUtilities.html' }); } catch (_) { /* DOM 依存の init は無視 */ }
  const esc = ctx.window.sharedUtilities && ctx.window.sharedUtilities.security
    ? ctx.window.sharedUtilities.security.escapeHtml.bind(ctx.window.sharedUtilities.security)
    : null;
  assert.ok(esc, 'sharedUtilities.security.escapeHtml が公開されている');
  assert.equal(esc('x" onmouseover="alert(1)'), 'x&quot; onmouseover=&quot;alert(1)');
  assert.equal(esc("<a href='x'>&</a>"), '&lt;a href=&#39;x&#39;&gt;&amp;&lt;/a&gt;');
  assert.equal(esc('3年1組'), '3年1組', '日本語は素通し');
  assert.equal(esc(''), '');
  assert.equal(esc(null), '');
});
