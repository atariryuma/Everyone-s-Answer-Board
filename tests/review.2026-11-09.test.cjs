/**
 * 2026-11-09 全体レビュー (Everyone-s-Answer-Board #2〜#17) の回帰テスト。
 * データはすべてダミー (example.com のメール・架空の名前)。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { gasResponseStubs } = require('./_helpers.cjs');

const src = (name) => fs.readFileSync(path.resolve(__dirname, '../src', name), 'utf8');
const SYSTEM_LIMITS = { MAX_LOCK_ROWS: 100, PREVIEW_LENGTH: 50, DEFAULT_PAGE_SIZE: 20, MAX_PAGE_SIZE: 100, RADIX_DECIMAL: 10 };
const quiet = { log: () => {}, warn: () => {}, error: () => {} };

function load(files, globals) {
  const context = { console: quiet, URL, ...gasResponseStubs(), ...globals };
  vm.createContext(context);
  for (const f of files) vm.runInContext(src(f), context, { filename: f });
  return context;
}

// ---------------------------------------------------------------- #3
test('#3 validateConfig: 文字列 "false" の showNames を「名前を出す」にしない', () => {
  const ctx = load(['validators.js'], { SYSTEM_LIMITS, getColumnAnalysis: () => ({ success: true }), getFormInfo: () => ({ success: true }) });
  const r = ctx.validateConfig({ displaySettings: { showNames: 'false', showReactions: 'true' } });
  assert.equal(r.sanitized.displaySettings.showNames, false);
  assert.equal(r.sanitized.displaySettings.showReactions, true);
  assert.equal(ctx.validateConfig({ displaySettings: { showNames: true } }).sanitized.displaySettings.showNames, true);
});

// ---------------------------------------------------------------- #5 / #13 / #16 (helpers.js)
function loadHelpers(overrides = {}) {
  const props = new Map();
  const puts = [];
  const ctx = load(['helpers.js'], {
    PROPERTY_CACHE_TTL: 30000,
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => (props.has(k) ? props.get(k) : null), setProperty: (k, v) => props.set(k, v) }) },
    CacheService: { getScriptCache: () => ({ get: () => null, put: (k, v) => puts.push([k, v]), remove: () => {} }) },
    ...overrides
  });
  return { ctx, props, puts };
}

test('#5 cache の上限はバイトで判定する (日本語 1 文字 = 3 バイト)', () => {
  const { ctx, puts } = loadHelpers();
  assert.equal(ctx.utf8ByteLength_('abc'), 3);
  assert.equal(ctx.utf8ByteLength_('あ'), 3);
  assert.equal(ctx.utf8ByteLength_('😀'), 4);
  const jp = 'あ'.repeat(40000);   // 4 万字 = 約 12 万バイト (文字数では 10 万未満)
  assert.equal(ctx.saveToCacheWithSizeCheck_('k1', jp, 60), false, '100KB を超えるので入れない');
  assert.equal(ctx.saveToCacheWithSizeCheck_('k2', 'a'.repeat(90000), 60), true);
  assert.deepEqual(puts.map((p) => p[0]), ['k2']);
});

test('#13 salt: lock の中では PropertiesService から読み直す (先に確定した salt を上書きしない)', () => {
  const { ctx, props } = loadHelpers({
    LockService: { getScriptLock: () => ({ waitLock: () => { props.set('EMAIL_HASH_SALT', 'salt-from-other'); }, releaseLock: () => {} }) },
    Utilities: {
      getUuid: () => 'new-salt-should-not-be-used',
      DigestAlgorithm: { SHA_1: 'SHA_1' }, Charset: { UTF_8: 'UTF_8' },
      computeDigest: (_a, s) => Array.from(String(s)).slice(0, 4).map((c) => c.charCodeAt(0))
    }
  });
  ctx.emailToShortHash_('child@example.com');
  assert.equal(props.get('EMAIL_HASH_SALT'), 'salt-from-other', '他の request が確定させた salt を残す');
});

test('#16 実行ログのメールは伏せる', () => {
  const { ctx } = loadHelpers();
  assert.equal(ctx.maskEmailForLog_('student123@example.com'), 'stu…@');
  assert.equal(ctx.maskEmailForLog_(''), '-');
  for (const f of ['DataApis.js', 'DataService.js']) {
    assert.doesNotMatch(src(f), /logError_\([^;]*\{\s*(?:targetUserId,\s*)?viewerEmail\s*\}/, `${f}: 伏せずに書かない`);
  }
});

// ---------------------------------------------------------------- #6 (main.js)
function loadMain(overrides = {}) {
  const reports = [];
  const store = new Map();
  const ctx = load(['main.js'], {
    CacheService: { getScriptCache: () => ({ get: (k) => store.get(k) || null, put: (k, v) => store.set(k, v) }) },
    Utilities: { DigestAlgorithm: { SHA_1: 'SHA_1' }, computeDigest: (_a, s) => Array.from({ length: 8 }, (_, i) => (String(s).length * 7 + i) & 0xff) },
    Session: { getActiveUser: () => ({ getEmail: () => 'child@example.com' }) },
    opsReport_: (where, err, place) => reports.push({ where, kind: err.name, place }),
    getCachedProperty_: () => null,
    ...overrides
  });
  return { ctx, reports };
}

test('#6 画面からの報告: 場所は固定の client・kind は決まった名前だけ・行は送らない', () => {
  const { ctx, reports } = loadMain();
  ctx.handleClientErrorReport_('child@example.com', { level: 'error', context: 'x/y/z-' + Math.random(), line: 1234, message: 'ダミーの文面 Secret: 値' });
  ctx.handleClientErrorReport_('child@example.com', { level: 'error', message: 'x', stack: 'TypeError: cannot read' });
  ctx.handleClientErrorReport_('child@example.com', { level: 'error', message: 'MyPrivateWord: ...' });
  assert.deepEqual(reports.map((r) => r.where), ['client', 'client', 'client']);
  assert.deepEqual(reports.map((r) => r.kind), ['Error', 'TypeError', 'Error'], '文面の単語を kind にしない');
  assert.ok(reports.every((r) => r.place.line === 0 && r.place.file === 'client'));
});

test('#6 画面からの報告: メールが取れないときは ops-monitor へ送らない・配置全体に上限がある', () => {
  const { ctx, reports } = loadMain();
  const r = ctx.handleClientErrorReport_('', { level: 'error', message: 'x' });
  assert.equal(r.success, true, '実行ログには残す');
  assert.equal(reports.length, 0);
  for (let i = 0; i < 40; i++) ctx.handleClientErrorReport_(`child${i}@example.com`, { level: 'error', message: 'x' });
  assert.equal(reports.length, 20, '1 時間に 20 件まで');
});

test('#12 JSON の構文エラーで本文の先頭を記録しない', () => {
  assert.doesNotMatch(src('main.js'), /dataPreview/);
});

// ---------------------------------------------------------------- #7 (AdminApis.js)
test('#7 markWelcomeSeen: 設定が壊れている (corrupted) ときは保存しない', () => {
  let saved = false;
  const ctx = load(['AdminApis.js'], {
    getCurrentEmail: () => 'teacher@example.com',
    findUserByEmail_: () => ({ userId: 'u1', userEmail: 'teacher@example.com' }),
    getUserConfig: () => ({ success: true, corrupted: true, config: {} }),
    saveUserConfig_: () => { saved = true; return { success: true }; }
  });
  const r = ctx.markWelcomeSeen();
  assert.equal(r.success, false);
  assert.equal(saved, false);
});

// ---------------------------------------------------------------- #2 / #8 / #9 / #15 (DataApis.js)
function fakeForm() {
  const form = {
    items: ['old1', 'old2'], deleted: [], destination: 'old-ss',
    getItems() { return this.items.slice(); },
    deleteItem(it) { this.deleted.push(it); },
    getDestinationId() { return this.destination; },
    setDestination(_t, id) { this.destination = id; },
    getId: () => 'form-1', getTitle: () => 'ダミーのフォーム',
    setTitle() {}, setDescription() {},
    addTextItem() { return { setTitle() { return this; }, setHelpText() { return this; }, setRequired() { return this; } }; }
  };
  return form;
}

function loadDataApis(overrides = {}) {
  return load(['DataApis.js'], {
    boardHashScope_: () => 'scope',
    getCurrentEmail: () => 'teacher@example.com',
    ...overrides
  });
}

test('#9 customizeForm_: 未対応の型があれば既存の設問を消さずに断る', () => {
  const form = fakeForm();
  const ctx = loadDataApis({ FormApp: { openById: () => form, openByUrl: () => form } });
  const r = ctx.customizeForm_('form-1', { questions: [{ type: 'unknownType', title: 'Q' }] });
  assert.equal(r.success, false);
  assert.deepEqual(form.deleted, [], '設問は消していない');
});

test('#8 customizeForm_: 前の回答のシートが読めないときは作り直さない (ゴミ箱へ移さない)', () => {
  const form = fakeForm();
  let created = false;
  let trashed = false;
  const ctx = loadDataApis({
    FormApp: { openById: () => form, openByUrl: () => form, DestinationType: { SPREADSHEET: 'SPREADSHEET' } },
    SpreadsheetApp: { openById: () => { throw new Error('一時的に読めない'); }, create: () => { created = true; return { getId: () => 'new-ss' }; } },
    DriveApp: { getFileById: () => ({ setTrashed: () => { trashed = true; }, setName() {} }) },
    CacheService: { getScriptCache: () => ({ get: () => null, put() {}, remove() {}, removeAll() {} }) }
  });
  ctx.customizeForm_('form-1', { questions: [{ type: 'text', title: 'Q' }] });
  assert.equal(created, false, '新しいシートを作らない');
  assert.equal(trashed, false, '前のシートをゴミ箱へ移さない');
  assert.equal(form.destination, 'old-ss');
});

test('#2 直前の結果 (LATEST) の key は接続先と設定の世代で分かれる', () => {
  const seen = [];
  const ctx = loadDataApis({
    CacheService: { getScriptCache: () => ({ get: () => null, put() {} }) },
    Utilities: { DigestAlgorithm: { MD5: 'MD5' }, Charset: { UTF_8: 'UTF_8' }, computeDigest: (_a, s) => Array.from(String(s).padStart(8, '_').slice(-8)).map((c) => c.charCodeAt(0)) },
    getBoardDataVersion_: () => 'v1',
    withStampedeGuard_: (opts) => { seen.push(opts.latestKey); return { success: true, data: [] }; }
  });
  const opts = { sortBy: 'newest' };
  ctx.withBoardDataCache_('u1', opts, () => ({ success: true, data: [] }), { config: { spreadsheetId: 'ss-A', sheetName: 'S', etag: 'e1' } });
  ctx.withBoardDataCache_('u1', opts, () => ({ success: true, data: [] }), { config: { spreadsheetId: 'ss-B', sheetName: 'S', etag: 'e1' } });
  ctx.withBoardDataCache_('u1', opts, () => ({ success: true, data: [] }), { config: { spreadsheetId: 'ss-A', sheetName: 'S', etag: 'e2' } });
  assert.equal(new Set(seen).size, 3, '接続先・設定が変われば別の key');
  assert.equal(ctx.boardSourceTag_(null), '');
});

test('#15 フェーズを判定できない (読み込み失敗) ときは、自分の行だけを返す', () => {
  const ROWS = { success: true, data: [{ email: 'child1@example.com' }, { email: 'child2@example.com' }] };
  const ctx = loadDataApis({ __getViewerLessonPhase_: () => null, __lessonPhaseUnknown_: () => true });
  const out = ctx.__maskOthersDuringInputPhase_(ROWS, 'u1', 'child1@example.com', false, false);
  assert.deepEqual(out.data.map((r) => r.email), ['child1@example.com']);
  const ctx2 = loadDataApis({ __getViewerLessonPhase_: () => null, __lessonPhaseUnknown_: () => false });
  assert.equal(ctx2.__maskOthersDuringInputPhase_(ROWS, 'u1', 'child1@example.com', false, false).data.length, 2, '授業中でなければそのまま');
});

test('#15 __getViewerLessonPhase_: 授業を指しているのにフェーズが読めず、直前の値も無ければ「判定できない」を立てる', () => {
  const ctx = load(['LessonService.js'], {
    CacheService: { getScriptCache: () => ({ get: () => null, put() {}, remove() {} }) },
    saveToCacheWithSizeCheck_: () => true
  });
  ctx.__computeViewerLessonPhase_ = () => undefined;   // lessons シートの読み込み失敗
  assert.equal(ctx.__getViewerLessonPhase_('u1'), null);
  assert.equal(ctx.__lessonPhaseUnknown_('u1'), true);
  ctx.__computeViewerLessonPhase_ = () => null;        // 授業中でない (読めた)
  assert.equal(ctx.__getViewerLessonPhase_('u1'), null);
  assert.equal(ctx.__lessonPhaseUnknown_('u1'), false);
});

test('#11 確認済みメールにできなかったら、作成の応答と記録で分かるようにする', () => {
  const s = src('DataApis.js');
  assert.match(s, /const emailVerified = setFormEmailCollectionVerified_\(form\.getId\(\)\);/);
  assert.match(s, /emailVerified: emailVerified,/);
});

// ---------------------------------------------------------------- #4 (DataService.js)
test('#4 フォームの回答の削除: 同じ秒の別の児童の回答は消さない (メールで照らし合わせ・複数なら止める)', () => {
  const resp = (id, ms, email) => ({ getId: () => id, getTimestamp: () => new Date(ms), getRespondentEmail: () => email });
  const t = Date.parse('2026-04-21T12:00:00Z');
  const deleted = [];
  const form = { responses: [], getResponses() { return this.responses; }, deleteResponse: (id) => deleted.push(id) };
  const ctx = load(['DataService.js'], { FormApp: { openByUrl: () => form } });
  const sheet = { getFormUrl: () => 'https://docs.google.com/forms/d/dummy/edit' };

  form.responses = [resp('a', t + 200, 'child1@example.com'), resp('b', t + 300, 'child2@example.com')];
  assert.equal(ctx.deleteLinkedFormResponseByTimestamp_(sheet, new Date(t), 'child2@example.com').success, true);
  assert.deepEqual(deleted, ['b'], 'メールが合う回答だけ');

  deleted.length = 0;
  const r = ctx.deleteLinkedFormResponseByTimestamp_(sheet, new Date(t), '');
  assert.equal(r.success, false, 'メールが分からず候補が 2 つ = 決められない');
  assert.deepEqual(deleted, []);
});

test('#4 回答の削除は「行の並び」の lock を取ってから行を消す', () => {
  const events = [];
  const data = [['タイムスタンプ', 'メールアドレス'], [new Date('2026-05-01T00:00:00Z'), 'child1@example.com']];
  const sheet = {
    getName: () => 'Sheet1', getParent: () => ({ getId: () => 'ss-1' }),
    getLastRow: () => data.length, getLastColumn: () => 2,
    getRange: (row, col, nr, nc) => ({ getValues: () => data.slice(row - 1, row - 1 + nr).map((r) => r.slice(col - 1, col - 1 + nc)) }),
    deleteRows: (start) => { events.push('delete'); data.splice(start - 1, 1); }
  };
  const ctx = load(['DataService.js'], {
    getCurrentEmail: () => 'teacher@example.com',
    findUserById_: () => ({ userId: 'u1', userEmail: 'teacher@example.com' }),
    isAdministrator: () => false,
    getUserConfig: () => ({ success: true, config: { spreadsheetId: 'ss-1', sheetName: 'Sheet1' } }),
    openSpreadsheet: () => ({ spreadsheet: { getSheetByName: () => sheet } }),
    getSheetInfo: () => ({ lastCol: 2 }),
    invalidateSheetRowCountCache_: () => {},
    answerSheetStructLockKey_: (ss, name) => `struct:${ss}:${name}`,
    acquireCacheLock_: (key) => { events.push('lock:' + key); return 'tok'; },
    releaseCacheLock_: (key) => events.push('release:' + key),
    CacheService: { getScriptCache: () => ({ get: () => null, put() {}, remove() {} }) },
    Utilities: { sleep: () => events.push('settle') }
  });
  const res = ctx.deleteAnswerRow('u1', 2);
  assert.equal(res.success, true, JSON.stringify(res));
  assert.deepEqual(events, ['lock:struct:ss-1:Sheet1', 'settle', 'delete', 'release:struct:ss-1:Sheet1']);
});

// ---------------------------------------------------------------- #14 / #4 (ReactionService.js)
test('#14 列を足すところはスクリプト全体の lock で包む・取れなければ書かない', () => {
  const ctx = load(['ReactionService.js'], {
    LockService: { getScriptLock: () => ({ tryLock: () => false, releaseLock() {} }) }
  });
  let ran = false;
  assert.throws(() => ctx.withColumnProvisionLock_(() => { ran = true; }), /混み合って/);
  assert.equal(ran, false);
  const ctx2 = load(['ReactionService.js'], {
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) }
  });
  assert.equal(ctx2.withColumnProvisionLock_(() => 42), 42);
});

test('#4 リアクション: 回答の削除中 (行の並びの lock) は書かない', () => {
  const s = src('ReactionService.js');
  assert.match(s, /cache\.get\(answerSheetStructLockKey_\(config\.spreadsheetId, config\.sheetName\)\)/);
});

// ---------------------------------------------------------------- #10
test('#10 publishApp はこの配置の先生だけ', () => {
  assert.match(src('SystemController.js'), /function publishApp\(publishConfig\) \{[\s\S]{0,600}teacherOnlyError_\(\)/);
});

// ---------------------------------------------------------------- #17 (page.js.html)
test('#17 リアクションは「書かずに断った」ときだけ再送する (押した結果を反転させない)', () => {
  const s = src('page.js.html');
  const block = s.slice(s.indexOf('async sendReactionToServerDirect('), s.indexOf('ensureReactionCleanup(rowIndex, reaction, reactionKey, error = null)'));
  assert.match(block, /retryCondition: \(error\) => \{\s*const msg = [^;]+;\s*return msg\.includes\('同時リアクション処理中'\) \|\| msg\.includes\('混み合って'\);/);
});
