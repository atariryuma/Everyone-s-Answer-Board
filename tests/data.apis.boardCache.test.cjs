/**
 * board data cache の圧縮表現 (DataApis.js packBoardDataForCache_ / unpackBoardDataFromCache_)。
 *
 * Why: ScriptCache は 1 value 100KB。行に answer と同じ opinion、常に false の isEmpty、
 *   timestamp から作れる formattedTimestamp が載っていて、150〜250 行で JSON が 100KB を超え
 *   「cache に入らない → 30 人の毎 poll が Sheets を読む」になっていた (2026-10 監査)。
 *   cache に書く形を圧縮しても、受け手が見る形 (wire) は変わらないことをここで固定する。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const zlib = require('node:zlib');
const { gasResponseStubs } = require('./_helpers.cjs');

const HELPERS = fs.readFileSync(path.resolve(__dirname, '../src/helpers.js'), 'utf8');
const DATA_APIS = fs.readFileSync(path.resolve(__dirname, '../src/DataApis.js'), 'utf8');

// GAS Utilities の gzip 系を Node の zlib で模す (Byte[] は signed int 配列)。
function gasGzipUtilities() {
  const toBytes = (buf) => Array.from(buf, (b) => (b > 127 ? b - 256 : b));
  const fromBytes = (bytes) => Buffer.from(bytes.map((b) => (b < 0 ? b + 256 : b)));
  const blob = (buf) => ({ getBytes: () => toBytes(buf), getDataAsString: () => buf.toString('utf8') });
  return {
    sleep: () => {},
    newBlob: (data) => blob(typeof data === 'string' ? Buffer.from(data, 'utf8') : fromBytes(data)),
    gzip: (b) => blob(zlib.gzipSync(fromBytes(b.getBytes()))),
    ungzip: (b) => blob(zlib.gunzipSync(fromBytes(b.getBytes()))),
    base64Encode: (bytes) => fromBytes(bytes).toString('base64'),
    base64Decode: (s) => toBytes(Buffer.from(s, 'base64'))
  };
}

function loadContext(opts = {}) {
  const store = new Map();
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    URL,
    ...gasResponseStubs(),
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {} }) },
    CacheService: {
      getScriptCache: () => ({
        get: (k) => (store.has(k) ? store.get(k) : null),
        put: (k, v) => { store.set(k, v); },
        remove: (k) => { store.delete(k); }
      })
    },
    Utilities: opts.gzip === false ? { sleep: () => {} } : gasGzipUtilities(),
    getCurrentEmail: () => 'viewer@example.com',
    isAdministrator: () => false,
    getUserConfig: () => ({ success: true, config: {} }),
    // formatters.js の formatTimestamp と同じ契約 (決定的な文字列)。
    formatTimestamp: (ts) => (ts ? new Date(ts).toISOString().slice(0, 16).replace('T', ' ') : '-')
  };
  vm.createContext(context);
  vm.runInContext(HELPERS, context, { filename: 'helpers.js' });
  vm.runInContext(DATA_APIS, context, { filename: 'DataApis.js' });
  return { context, store };
}

function makeRows(n, answerLen) {
  const jp = 'わたしは、うそをつくと信じてもらえなくなるので、自由には責任がともなうと思いました。';
  const text = (i) => (jp.repeat(Math.ceil(answerLen / jp.length)) + i).slice(0, answerLen);
  const rows = [];
  for (let i = 0; i < n; i++) {
    const ts = new Date(Date.UTC(2026, 9, 16, 1, i % 60, 0));
    rows.push({
      id: `row_${i + 2}`, rowIndex: i + 2, timestamp: ts, email: `student${i}@example.com`,
      answer: text(i), opinion: text(i), reason: '理由' + i, class: '6-1', name: '児童' + i,
      numericX: (i % 5) + 1, numericY: null,
      formattedTimestamp: ts.toISOString().slice(0, 16).replace('T', ' '), isEmpty: false,
      reactions: {
        UNDERSTAND: { count: 1, reacted: false, voters: ['a@example.com'] },
        LIKE: { count: 0, reacted: false, voters: [] },
        CURIOUS: { count: 0, reacted: false, voters: [] }
      },
      highlight: false
    });
  }
  return { success: true, data: rows, headers: ['h1', 'h2'], sheetName: 'フォームの回答 1' };
}

// JSON を通ったあとの形 (Date → ISO 文字列)。旧実装 (そのまま cache) でも受け手はこの形を見ていた。
//   vm 側の object は prototype が別 realm なので、比較は両辺ともこちらの realm の JSON に揃える。
const viaJson = (v) => JSON.parse(JSON.stringify(v));

test('board cache: 2 回目 (cache hit) は 1 回目と同じ wire の形 (voters / 派生 3 項目 / Date→ISO)', () => {
  const { context, store } = loadContext();
  let loads = 0;
  const fresh = makeRows(3, 40);
  const loader = () => { loads++; return fresh; };
  const first = context.withBoardDataCache_('u1', { classFilter: '6-1', sortBy: 'newest' }, loader);
  assert.equal(first, fresh, '1 回目は loader の結果そのもの');
  const second = context.withBoardDataCache_('u1', { classFilter: '6-1', sortBy: 'newest' }, loader);
  assert.equal(loads, 1, '2 回目は読まない');
  assert.deepEqual(viaJson(second), viaJson(fresh));
  assert.deepEqual(viaJson(second.data[0].reactions.UNDERSTAND.voters), ['a@example.com'], 'voters (閲覧者ごとの reacted の元) は残る');
  // cache には圧縮表現で入っている
  const key = [...store.keys()].find((k) => k.startsWith('board_data:u1:'));
  assert.ok(key, 'board_data key が書かれる');
  const raw = JSON.parse(store.get(key));
  assert.ok(typeof raw.__boardGz === 'string', 'gzip+base64 で書く');
  assert.ok(!raw.data, '生の行は cache に入れない');
});

test('board cache: 250 行 (JSON 100KB 超) でも cache に入り、毎 poll が Sheets を読まない', () => {
  const { context, store } = loadContext();
  let loads = 0;
  const fresh = makeRows(250, 150);
  const rawLen = JSON.stringify(fresh).length;
  assert.ok(rawLen > 100000, `前提: 生の JSON は 100KB 超 (${rawLen})`);
  const loader = () => { loads++; return fresh; };
  context.withBoardDataCache_('u1', { sortBy: 'newest' }, loader);
  const key = [...store.keys()].find((k) => k.startsWith('board_data:u1:'));
  assert.ok(key && store.get(key).length < 100000, `圧縮後は 100KB 未満 (${key ? store.get(key).length : 'none'})`);
  const second = context.withBoardDataCache_('u1', { sortBy: 'newest' }, loader);
  assert.equal(loads, 1, '2 回目は cache hit');
  assert.equal(second.data.length, 250);
  assert.deepEqual(viaJson(second.data[249]), viaJson(fresh.data[249]));
});

test('board cache: gzip が無い環境でも派生項目を落とした形で cache し、同じ形に戻す', () => {
  const { context, store } = loadContext({ gzip: false });
  const fresh = makeRows(2, 30);
  context.withBoardDataCache_('u1', {}, () => fresh);
  const key = [...store.keys()].find((k) => k.startsWith('board_data:u1:'));
  const raw = JSON.parse(store.get(key));
  assert.ok(raw.__boardSlim, '非圧縮 slim 形');
  assert.ok(!('opinion' in raw.__boardSlim.data[0]), 'opinion は落とす');
  assert.ok(!('isEmpty' in raw.__boardSlim.data[0]), 'isEmpty は落とす');
  assert.ok(!('formattedTimestamp' in raw.__boardSlim.data[0]), 'formattedTimestamp は落とす');
  const second = context.withBoardDataCache_('u1', {}, () => { throw new Error('must not load'); });
  assert.deepEqual(viaJson(second), viaJson(fresh));
});

test('board cache: includeTimestamp=false の行 (timestamp 無し) には formattedTimestamp を足さない', () => {
  const { context } = loadContext();
  const fresh = makeRows(1, 10);
  delete fresh.data[0].timestamp;
  delete fresh.data[0].formattedTimestamp;
  context.withBoardDataCache_('u1', {}, () => fresh);
  const second = context.withBoardDataCache_('u1', {}, () => { throw new Error('must not load'); });
  assert.ok(!('formattedTimestamp' in second.data[0]));
  assert.ok(!('timestamp' in second.data[0]));
  assert.equal(second.data[0].opinion, fresh.data[0].answer);
});

test('board cache: opinion が answer と違う行はそのまま保つ (落とすのは同値のときだけ)', () => {
  const { context } = loadContext();
  const fresh = makeRows(1, 10);
  fresh.data[0].opinion = '別の値';
  context.withBoardDataCache_('u1', {}, () => fresh);
  const second = context.withBoardDataCache_('u1', {}, () => { throw new Error('must not load'); });
  assert.equal(second.data[0].opinion, '別の値');
});

test('board cache: 他が読んでいる間の直前の結果 (latest) も圧縮表現から戻し stale を付ける', () => {
  const { context, store } = loadContext();
  const fresh = makeRows(2, 20);
  context.withBoardDataCache_('u1', {}, () => fresh); // key + latest を書く
  // version を上げて key を無効化し、誰かが読んでいる状態にする
  store.set('board_data_ver:u1', '1');
  store.set('board_data_flight:u1:_:newest', '1');
  const stale = context.withBoardDataCache_('u1', {}, () => { throw new Error('must not load'); });
  assert.equal(stale.stale, true);
  assert.deepEqual(viaJson(stale.data), viaJson(fresh).data);
});

test('board cache: 壊れた圧縮値は miss として読み直す', () => {
  const { context, store } = loadContext();
  let loads = 0;
  const fresh = makeRows(1, 10);
  context.withBoardDataCache_('u1', {}, () => { loads++; return fresh; });
  const key = [...store.keys()].find((k) => k.startsWith('board_data:u1:'));
  store.set(key, JSON.stringify({ __boardGz: 'not-gzip!!' }));
  store.delete('board_data_latest:u1:_:newest');
  const again = context.withBoardDataCache_('u1', {}, () => { loads++; return fresh; });
  assert.equal(loads, 2);
  assert.equal(again, fresh);
});
