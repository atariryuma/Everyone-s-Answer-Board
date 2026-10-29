/**
 * 行番号のずれ防御 (ReactionService.resolveRowIdentity_ / executeBoardRowOperation_ の expect) と
 * users 一覧 cache の圧縮 (DatabaseCore.getAllUsers_ の pack / unpack) を固定する。
 *
 * Why: 教師が回答を削除すると行番号が繰り上がり、次の poll までの数秒間、児童の画面の行番号は
 *   隣の回答を指す。児童の画面が持つ emailHash (塩入り) で同じ回答か確かめ、違えば書かない。
 *   users 一覧は人数が 30〜50 人になると素の JSON が 100KB を超えて cache に入らなくなる。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const zlib = require('node:zlib');
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

// GAS Utilities の gzip 系を Node の zlib で模す (Byte[] は signed int 配列)。
function gasGzipUtilities() {
  const toBytes = (buf) => Array.from(buf, (b) => (b > 127 ? b - 256 : b));
  const fromBytes = (bytes) => Buffer.from(bytes.map((b) => (b < 0 ? b + 256 : b)));
  const blob = (buf) => ({ getBytes: () => toBytes(buf), getDataAsString: () => buf.toString('utf8') });
  return {
    sleep: () => {},
    getUuid: () => 'uuid',
    newBlob: (data) => blob(typeof data === 'string' ? Buffer.from(data, 'utf8') : fromBytes(data)),
    gzip: (b) => blob(zlib.gzipSync(fromBytes(b.getBytes()))),
    ungzip: (b) => blob(zlib.gunzipSync(fromBytes(b.getBytes()))),
    base64Encode: (bytes) => fromBytes(bytes).toString('base64'),
    base64Decode: (s) => toBytes(Buffer.from(String(s), 'base64'))
  };
}

// =====================================================================
// resolveRowIdentity_
// =====================================================================

function loadReaction(overrides = {}) {
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    ...gasResponseStubs(),
    CacheService: { getScriptCache: () => makeCache() },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
    Utilities: { sleep: () => {} },
    getCurrentEmail: () => 'viewer@example.com',
    isAdministrator: () => false,
    findPublishedBoardOwner_: () => null,
    getConfigOrDefault: () => ({}),
    openSpreadsheet: () => null,
    resolveReactionMode: () => 'count',
    // 本物と同じ契約: mapping 優先、無ければ見出し名で探す (email だけ分かればよい)
    resolveColumnIndex: (headers, field, mapping) => {
      if (mapping && typeof mapping[field] === 'number') return { index: mapping[field] };
      const idx = headers.findIndex((h) => String(h).toLowerCase().includes(field === 'email' ? 'mail' : field));
      return { index: idx };
    },
    emailToShortHash_: (email) => (email ? 'h_' + String(email).trim().toLowerCase() : null),
    ...overrides
  };
  vm.createContext(context);
  vm.runInContext(SRC('ReactionService.js'), context, { filename: 'ReactionService.js' });
  return context;
}

test('resolveRowIdentity_: email ハッシュは email 列のハッシュと比べる (大文字小文字は無視)', () => {
  const ctx = loadReaction();
  const headers = ['timestamp', 'email', 'answer'];
  const id = ctx.resolveRowIdentity_({ field: 'email', hash: 'h_a@example.com' }, headers, {});
  assert.equal(id.column, 2);
  assert.equal(id.matches('A@Example.com'), true);
  assert.equal(id.matches('b@example.com'), false, '別の児童の行 = ずれている');
  assert.equal(id.matches(''), true, '匿名行 (email 無し) は判定できないので通す');
});

test('resolveRowIdentity_: email 列が無いボードでは照合しない (null)', () => {
  const ctx = loadReaction();
  assert.equal(ctx.resolveRowIdentity_({ field: 'email', hash: 'h_x' }, ['timestamp', 'answer'], {}), null);
  assert.equal(ctx.resolveRowIdentity_({ field: 'email', hash: 'h_x' }, [], {}), null, '見出しが無ければ照合しない');
});

test('resolveRowIdentity_: timestamp は時刻として比べる (Date と文字列の差を吸収)', () => {
  const ctx = loadReaction();
  const headers = ['タイムスタンプ', 'answer'];
  const id = ctx.resolveRowIdentity_({ field: 'timestamp', value: '2026-10-16T01:02:03.000Z' }, headers, {});
  assert.equal(id.column, 1);
  assert.equal(id.matches(new Date('2026-10-16T01:02:03.000Z')), true);
  assert.equal(id.matches('2026-10-16T01:02:04.000Z'), false);
});

test('resolveRowIdentity_: 旧形式 { column, value } はそのまま文字列一致', () => {
  const ctx = loadReaction();
  const id = ctx.resolveRowIdentity_({ column: 2, value: 'A@x.com' }, ['t', 'email'], {});
  assert.equal(id.column, 2);
  assert.equal(id.matches(' a@x.com '), true);
});

test('executeBoardRowOperation_: 行がずれていたら書かずに ROW_CHANGED を返す', () => {
  const writes = [];
  const data = [['timestamp', 'email', 'answer', 'UNDERSTAND', 'LIKE', 'CURIOUS'], ['t1', 'a@example.com', 'A', '', '', ''], ['t2', 'b@example.com', 'B', '', '', '']];
  const sheet = {
    getName: () => 'S', getParent: () => ({ getId: () => 'ss' }),
    getLastColumn: () => 6,
    getRange: (row, col, numRows = 1, numCols = 1) => ({
      getValues: () => [data[row - 1].slice(col - 1, col - 1 + numCols)],
      setValue: (v) => { writes.push({ row, col, v }); },
      setValues: (vs) => { writes.push({ row, col, vs }); }
    })
  };
  const ctx = loadReaction({
    findPublishedBoardOwner_: () => ({ userId: 'owner', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ spreadsheetId: 'ss', sheetName: 'S', isPublished: true, displaySettings: { reactionMode: 'count' } }),
    openSpreadsheet: () => ({ spreadsheet: { getSheetByName: () => sheet } }),
    getSheetHeaders: () => ({ headers: data[0] })
  });
  // シートの 2 行目 = a さん。児童の画面が「2 行目 = a さん」と思っていれば書ける。
  const res = ctx.addReaction('owner', 2, 'LIKE', { rowIdentity: { field: 'email', hash: 'h_a@example.com' } });
  assert.equal(res.success, true, '正体が合えば書く: ' + JSON.stringify(res));
  // 1 行消えて児童の画面の「2 行目」が実は b さんだった (画面は b さんのつもりで押した) → 書かない
  const res2 = ctx.addReaction('owner', 2, 'LIKE', { rowIdentity: { field: 'email', hash: 'h_b@example.com' } });
  assert.equal(res2.success, false);
  assert.equal(res2.error, 'ROW_CHANGED');
  assert.equal(writes.length, 1, 'ずれていた 2 回目は何も書かない');
});

test('addReaction: 画面から来た { column, value } は照合に使わない (隠れた列の中身を当てさせない)', () => {
  const data = [['timestamp', 'email', 'answer', 'UNDERSTAND', 'LIKE', 'CURIOUS'], ['t1', 'a@example.com', 'A', '', '', '']];
  const sheet = {
    getName: () => 'S', getParent: () => ({ getId: () => 'ss' }),
    getLastColumn: () => 6,
    getRange: (row, col, numRows = 1, numCols = 1) => ({
      getValues: () => [data[row - 1].slice(col - 1, col - 1 + numCols)],
      setValue: () => {}, setValues: () => {}
    })
  };
  const ctx = loadReaction({
    findPublishedBoardOwner_: () => ({ userId: 'owner', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ spreadsheetId: 'ss', sheetName: 'S', isPublished: true, displaySettings: { reactionMode: 'count' } }),
    openSpreadsheet: () => ({ spreadsheet: { getSheetByName: () => sheet } }),
    getSheetHeaders: () => ({ headers: data[0] })
  });
  // メール列 (2 列目) に別の値を当てても ROW_CHANGED にならない = 当たり外れが分からない
  const guess = ctx.addReaction('owner', 2, 'LIKE', { rowIdentity: { column: 2, value: 'b@example.com' } });
  assert.notEqual(guess.error, 'ROW_CHANGED', JSON.stringify(guess));
  assert.equal(ctx.__clientRowExpect_({ rowIdentity: { column: 2, value: 'x' } }).rowIdentity.column, undefined);
  assert.deepEqual(JSON.parse(JSON.stringify(ctx.__clientRowExpect_({ rowIdentity: { field: 'email', hash: 'h' } }))),
    { rowIdentity: { field: 'email', hash: 'h' } }, '画面が送る形はそのまま');
});

// =====================================================================
// getAllUsers_: 一覧 cache を gzip で書く
// =====================================================================

test('getAllUsers: 一覧は圧縮して cache し、読み返すと同じ形 (100KB 超でも入る)', () => {
  const cache = makeCache();
  const helpersCtx = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    PROPERTY_CACHE_TTL: 30000,
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {} }) },
    CacheService: { getScriptCache: () => cache },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
    Utilities: gasGzipUtilities(),
    ...gasResponseStubs()
  };
  vm.createContext(helpersCtx);
  vm.runInContext(SRC('helpers.js'), helpersCtx, { filename: 'helpers.js' });
  // 50 人分、configJson 4KB ずつ = 素の JSON で 200KB 超
  const rows = [];
  for (let i = 0; i < 50; i++) {
    rows.push(['u' + i, 'user' + i + '@example.com', '', true, JSON.stringify({ isPublished: true, note: 'x'.repeat(4000), i }), '', '']);
  }
  const data = [['userId', 'userEmail', 'googleId', 'isActive', 'configJson', 'lastModified', 'createdAt'], ...rows];
  let reads = 0;
  const sheet = { getName: () => 'users', getDataRange: () => ({ getValues: () => { reads++; return data.map((r) => r.slice()); } }) };
  Object.assign(helpersCtx, {
    CACHE_DURATION: { SHORT: 10, MEDIUM: 30, LONG: 300, DATABASE_LONG: 600, USER_INDIVIDUAL: 900 },
    SpreadsheetApp: { openById: () => { throw new Error('not stubbed'); } },
    validateEmail: (e) => ({ isValid: /.+@.+/.test(String(e)) }),
    getCurrentEmail: () => 'admin@example.com',
    isAdministrator: () => true,
    getUserConfig: () => ({ success: true, config: {} }),
    executeWithRetry_: (fn) => fn(),
    DEFAULT_DISPLAY_SETTINGS: {}
  });
  vm.runInContext(SRC('DatabaseCore.js'), helpersCtx, { filename: 'DatabaseCore.js' });
  helpersCtx.openDatabase = () => ({ getSheetByName: () => sheet });

  const first = helpersCtx.getAllUsers_({ activeOnly: false }, { systemLookup: true });
  assert.equal(first.length, 50);
  assert.equal(reads, 1);
  const stored = [...cache._store.entries()].find(([k]) => k.startsWith('all_users_v'));
  assert.ok(stored, '一覧が cache に入る');
  assert.ok(stored[1].length < 100000, `cache 値は 100KB 未満 (${stored[1].length} 文字)`);
  assert.ok(JSON.parse(stored[1]).__gz, '圧縮表現で書かれている');

  const second = helpersCtx.getAllUsers_({ activeOnly: false }, { systemLookup: true });
  assert.equal(reads, 1, '2 回目は cache から (シートを読まない)');
  assert.equal(JSON.stringify(second), JSON.stringify(first), '読み返した形は同じ');
});
