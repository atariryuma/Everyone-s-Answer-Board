/**
 * 新しい版への引き継ぎと、前の URL からの転送 (src/Relocation.js) のテスト。
 *
 * Why: 転送は児童が開くたびに通る入口の処理。壊れると児童が開けなくなるので、失敗したら送らずに続けること、
 *   自分へ送って回り続けないこと、Apps Script のウェブアプリ以外へは送らないことを確かめる。
 *   引き継ぎは DB の付け替えなので、先生本人の回答ボード以外を DB にできないことを確かめる。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const SRC = fs.readFileSync(path.resolve(__dirname, '../src/Relocation.js'), 'utf8');
const exec = (id) => `https://script.google.com/a/macros/example.ed.jp/s/${id}/exec`;
const OLD_ID = 'a'.repeat(44);
const OWN_SS = 'b'.repeat(44);
const ADMIN = 'teacher@example.ed.jp';
const plain = (v) => JSON.parse(JSON.stringify(v));

function fakeBoard({ owner = ADMIN, sheets = ['users', 'lessons'], userEmails = [ADMIN], lessonRows = 3, name = 'みんなの回答ボード のコピー', tabHead = '', tabRows = null } = {}) {
  const written = {};
  const made = {};
  const sheetObj = (n) => ({
    getLastRow: () => {
      if (n === 'users') return userEmails.length + 1;
      if (n === 'lessons') return lessonRows + 1;
      if (tabRows !== null) return tabRows;
      return written[n] || tabHead ? 3 : 0;
    },
    getRange: () => {
      const chain = {
        getValues: () => userEmails.map((e) => ['uid', e]),
        getValue: () => (written[n] ? written[n][0][0] : tabHead),
        setNumberFormat: () => chain,
        setValues: (v) => { written[n] = v; return chain; }
      };
      return chain;
    }
  });
  const ss = {
    getSheetByName: (n) => {
      if (!sheets.includes(n) && !made[n]) return null;
      made[n] = made[n] || sheetObj(n);
      return made[n];
    },
    insertSheet: (n) => { made[n] = sheetObj(n); return made[n]; }
  };
  return { ss, written, owner, name };
}

function load({ baked = false, ownUrl = exec('NEW'), marker = '', teacher = true, board = fakeBoard(), files = [OLD_ID],
  dbFails = false, markerWriteFails = false, activeNull = false, containerKnown = true } = {}) {
  const props = { ADMIN_EMAIL: ADMIN, DATABASE_SPREADSHEET_ID: OWN_SS };
  if (containerKnown) props.CONTAINER_SPREADSHEET_ID = OWN_SS;
  const container = fakeBoard({ sheets: [] });
  const cache = {};
  const calls = { dbOpens: 0, userCacheCleared: 0, guide: [] };
  if (markerWriteFails) board.ss.insertSheet = () => { throw new Error('cannot write'); };
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    isBakedDeployment_: () => baked,
    getWebAppUrl: () => ownUrl,
    teacherOnlyError_: () => (teacher ? null : { success: false, message: '管理者権限が必要です' }),
    createSuccessResponse: (message, data) => ({ success: true, message, data }),
    createErrorResponse: (message) => ({ success: false, message, error: message }),
    getCachedProperty_: (k) => props[k] || null,
    setCachedProperty_: (k, v) => { props[k] = v; },
    clearDatabaseUserCache_: () => { calls.userCacheCleared += 1; },
    invalidateAccessDecisionCache_: () => {},
    validateSpreadsheetId: (id) => ({ isValid: /^[a-zA-Z0-9-_]{40,50}$/.test(String(id)) }),
    markGuideAdopted_: (ss, name) => { calls.guide.push(name); },
    FIRST_STEPS_SHARED_KEY: 'FIRST_STEPS_URL_SHARED',
    FIRST_STEPS_CHILD_KEY: 'FIRST_STEPS_CHILD_OPENED',
    openDatabase: () => {
      calls.dbOpens += 1;
      if (dbFails) throw new Error('db down');
      return { getSheetByName: (n) => (n === '引っ越し先' && marker ? { getRange: () => ({ getValue: () => marker }) } : null) };
    },
    CacheService: { getScriptCache: () => ({ get: (k) => cache[k] || null, put: (k, v) => { cache[k] = v; }, remove: (k) => { delete cache[k]; } }) },
    PropertiesService: { getScriptProperties: () => ({
      getProperty: (k) => props[k] || null,
      setProperty: (k, v) => { props[k] = v; },
      deleteProperty: (k) => { delete props[k]; }
    }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
    SpreadsheetApp: {
      // 画面からの呼び出しでは紐づけ元を返さないこともある、という前提で作る (activeNull)
      getActiveSpreadsheet: () => (activeNull ? null : { getId: () => OWN_SS }),
      openById: (id) => (id === OWN_SS ? container.ss : board.ss)
    },
    DriveApp: {
      getFileById: () => ({
        getOwner: () => (board.owner ? { getEmail: () => board.owner } : null),
        getName: () => board.name,
        getLastUpdated: () => new Date('2026-11-01T00:00:00Z')
      }),
      searchFiles: () => {
        let i = 0;
        return { hasNext: () => i < files.length, next: () => { const id = files[i]; i += 1; return { getId: () => id }; } };
      }
    }
  };
  vm.createContext(context);
  vm.runInContext(SRC, context, { filename: 'Relocation.js' });
  return { context, props, cache, calls, board, container };
}

// ─── 転送 ───────────────────────────────────────────

test('転送: 引っ越し先が自分でなければ、その URL を返す', () => {
  const { context } = load({ ownUrl: exec('OLD'), marker: exec('NEW') });
  assert.equal(context.movedDestination_(), exec('NEW'));
});

test('転送: 引っ越し先が自分なら送らない (新しい配置は同じ DB の印を読む)', () => {
  const { context } = load({ ownUrl: exec('NEW'), marker: exec('NEW') });
  assert.equal(context.movedDestination_(), '');
});

test('転送: 引っ越していなければ送らず、開くたびに DB を読まない', () => {
  const { context, calls } = load({ marker: '' });
  for (let i = 0; i < 30; i += 1) assert.equal(context.movedDestination_(), '');
  assert.equal(calls.dbOpens, 1);
});

test('転送: Apps Script のウェブアプリ以外の URL には送らない', () => {
  const bads = ['https://evil.example.com/s/x/exec', 'http://script.google.com/macros/s/abc/exec',
    'https://script.google.com/macros/s/abc/exec?x=1', 'https://script.google.com/macros/s/abc/dev',
    'https://script.google.com.evil.example/macros/s/abc/exec', '=HYPERLINK("x")'];
  for (const bad of bads) {
    assert.equal(load({ ownUrl: exec('OLD'), marker: bad }).context.movedDestination_(), '', bad);
  }
});

test('転送: 自分の URL が分からないとき・DB を読めないときは送らない (児童の画面を止めない)', () => {
  assert.equal(load({ ownUrl: '', marker: exec('NEW') }).context.movedDestination_(), '');
  assert.equal(load({ ownUrl: exec('OLD'), marker: exec('NEW'), dbFails: true }).context.movedDestination_(), '');
});

test('転送: 焼き込みのある配置では DB も cache も読まない', () => {
  const { context, calls } = load({ baked: true, ownUrl: exec('OLD'), marker: exec('NEW') });
  assert.equal(context.movedDestination_(), '');
  assert.equal(calls.dbOpens, 0);
});

test('転送: 開こうとしていたページを付け、もう一度送らない印を付ける', () => {
  const { context } = load();
  assert.equal(context.movedTargetUrl_(exec('NEW'), { mode: 'view', userId: 'u 1&x', other: 'no' }),
    exec('NEW') + '?moved=NEW&mode=view&userId=u%201%26x');
  assert.equal(context.movedTargetUrl_(exec('NEW'), null), exec('NEW') + '?moved=NEW');
});

test('転送: 送られてきた先がいまの引っ越し先なら、もう送らない。次の引っ越しのあとは送る', () => {
  // 自分の URL を読み違えた新しい配置 (OLD と名乗る) でも、moved が合っていれば回り続けない
  assert.equal(load({ ownUrl: exec('OLD'), marker: exec('NEW') }).context.movedDestination_('NEW'), '');
  // 前の moved が付いたブックマーク: 送り先が NEWER に変わったら送る
  assert.equal(load({ ownUrl: exec('NEW'), marker: exec('NEWER') }).context.movedDestination_('NEW'), exec('NEWER'));
});

test('転送: よその組織のウェブアプリには送らない', () => {
  const other = 'https://script.google.com/a/macros/other.example.com/s/NEW/exec';
  assert.equal(load({ ownUrl: exec('OLD'), marker: other }).context.movedDestination_(), '');
});

// ─── 引き継ぎ ───────────────────────────────────────

test('引き継ぎ: 前のスプレッドシートを DB にし、引っ越し先を書き、4 歩を済みにする', () => {
  const { context, props, calls, board } = load();
  const res = context.adoptPreviousBoard(`https://docs.google.com/spreadsheets/d/${OLD_ID}/edit#gid=0`);
  assert.equal(res.success, true, res.message);
  assert.equal(props.DATABASE_SPREADSHEET_ID, OLD_ID);
  assert.equal(board.written['引っ越し先'][0][1], exec('NEW'));
  assert.equal(props.FIRST_STEPS_URL_SHARED, '1');
  assert.equal(props.FIRST_STEPS_CHILD_OPENED, '1');
  assert.equal(calls.userCacheCleared, 1);
  assert.deepEqual(calls.guide, [board.name]);
});

test('引き継ぎ: 先生でなければ・焼き込みのある配置では、何も変えない', () => {
  for (const opts of [{ teacher: false }, { baked: true }]) {
    const { context, props } = load(opts);
    assert.equal(context.adoptPreviousBoard(OLD_ID).success, false);
    assert.equal(context.findPreviousBoards().success, false);
    assert.equal(props.DATABASE_SPREADSHEET_ID, OWN_SS);
  }
});

test('引き継ぎ: 持ち主が違う・回答ボードでない・自分の登録が無いスプレッドシートは断る', () => {
  const boards = [fakeBoard({ owner: 'other@example.ed.jp' }), fakeBoard({ owner: '' }),
    fakeBoard({ sheets: ['users'] }), fakeBoard({ userEmails: ['someone@example.ed.jp'] })];
  for (const board of boards) {
    const { context, props } = load({ board });
    assert.equal(context.adoptPreviousBoard(OLD_ID).success, false);
    assert.equal(props.DATABASE_SPREADSHEET_ID, OWN_SS);
    assert.equal(board.written['引っ越し先'], undefined);
  }
});

test('引き継ぎ: 読み取れない入力・いま使っているスプレッドシート・テスト用の URL からは断る', () => {
  assert.equal(load().context.adoptPreviousBoard('こんにちは').success, false);
  assert.equal(load().context.adoptPreviousBoard({ id: OLD_ID }).success, false);
  assert.equal(load().context.adoptPreviousBoard(OWN_SS).success, false);
  const dev = load({ ownUrl: 'https://script.google.com/macros/s/abc/dev' });
  assert.equal(dev.context.adoptPreviousBoard(OLD_ID).success, false);
  assert.equal(dev.props.DATABASE_SPREADSHEET_ID, OWN_SS);
});

test('引き継ぎ: 引っ越し先を書けなかったら DB を戻す (送り先の無い引き継ぎを残さない)', () => {
  const { context, props } = load({ markerWriteFails: true });
  assert.equal(context.adoptPreviousBoard(OLD_ID).success, false);
  assert.equal(props.DATABASE_SPREADSHEET_ID, OWN_SS);
});

test('引き継ぎ: 2 つ目は断る (1 つ目の前の URL の児童が、2 つ目のデータへ送られないように)', () => {
  const { context, props } = load();
  assert.equal(context.adoptPreviousBoard(OLD_ID).success, true);
  const second = context.adoptPreviousBoard('c'.repeat(44));
  assert.equal(second.success, false);
  assert.equal(props.DATABASE_SPREADSHEET_ID, OLD_ID);
});

test('引き継ぎ: 人が作った同名のタブは上書きせずに断り、アプリが書いたタブは書き直す', () => {
  const mine = fakeBoard({ sheets: ['users', 'lessons', '引っ越し先'], tabHead: '来年の引っ越しのメモ' });
  const a = load({ board: mine });
  assert.equal(a.context.adoptPreviousBoard(OLD_ID).success, false);
  assert.equal(a.props.DATABASE_SPREADSHEET_ID, OWN_SS);
  assert.equal(mine.written['引っ越し先'], undefined);
  const ours = fakeBoard({ sheets: ['users', 'lessons', '引っ越し先'], tabHead: '新しい回答ボードの URL' });
  const b = load({ board: ours });
  assert.equal(b.context.adoptPreviousBoard(OLD_ID).success, true);
  assert.equal(ours.written['引っ越し先'][0][1], exec('NEW'));
});

test('引き継ぎ: 印を書いたあとの後片付けに失敗しても、DB は戻さない (前の URL の送り先と食い違わせない)', () => {
  const { context, props } = load();
  context.clearDatabaseUserCache_ = () => { throw new Error('cache down'); };
  context.invalidateAccessDecisionCache_ = () => { throw new Error('cache down'); };
  context.markGuideAdopted_ = () => { throw new Error('sheet down'); };
  assert.equal(context.adoptPreviousBoard(OLD_ID).success, true);
  assert.equal(props.DATABASE_SPREADSHEET_ID, OLD_ID);
});

test('引き継ぎ: 見出しの無いタブでも、中身があれば上書きしない (空のタブは使う)', () => {
  const filled = fakeBoard({ sheets: ['users', 'lessons', '引っ越し先'], tabHead: '', tabRows: 2 });
  assert.equal(load({ board: filled }).context.adoptPreviousBoard(OLD_ID).success, false);
  assert.equal(filled.written['引っ越し先'], undefined);
  const empty = fakeBoard({ sheets: ['users', 'lessons', '引っ越し先'], tabHead: '', tabRows: 0 });
  assert.equal(load({ board: empty }).context.adoptPreviousBoard(OLD_ID).success, true);
});

test('引き継ぎ: 画面からの呼び出しで紐づけ元を読めなくても、控えた ID で動く', () => {
  const { context, props, container } = load({ activeNull: true });
  assert.equal(context.adoptPreviousBoard(OLD_ID).success, true);
  assert.equal(props.DATABASE_SPREADSHEET_ID, OLD_ID);
  assert.equal(container.written['データの置き場'][0][0], '授業と回答が入っているスプレッドシート');
  // 控えも無く、紐づけ元も読めないときも、引き継ぎそのものは済ませる (案内のタブを書かないだけ)
  const none = load({ activeNull: true, containerKnown: false });
  assert.equal(none.context.adoptPreviousBoard(OLD_ID).success, true);
});

test('引き継ぎ: 引き継ぎを済ませたコピー (データを持たないアプリの本体) は、相手にできない', () => {
  const carrier = fakeBoard({ sheets: ['users', 'lessons', 'データの置き場'], tabHead: '授業と回答が入っているスプレッドシート' });
  const { context, props } = load({ board: carrier });
  assert.equal(context.adoptPreviousBoard(OLD_ID).success, false);
  assert.equal(props.DATABASE_SPREADSHEET_ID, OWN_SS);
  assert.deepEqual(plain(context.findPreviousBoards().data.boards), []);
});

test('探す: 自分の回答ボードだけを返し、いま使っているものは除く', () => {
  const res = load({ files: [OWN_SS, OLD_ID] }).context.findPreviousBoards();
  assert.equal(res.success, true);
  assert.deepEqual(plain(res.data.boards),
    [{ id: OLD_ID, name: 'みんなの回答ボード のコピー', updatedAt: '2026-11-01T00:00:00.000Z', lessonCount: 3 }]);
  const others = load({ board: fakeBoard({ owner: 'x@example.ed.jp' }) }).context.findPreviousBoards();
  assert.deepEqual(plain(others.data.boards), []);
});
