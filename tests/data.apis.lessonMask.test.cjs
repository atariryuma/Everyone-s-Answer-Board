/**
 * 「考える」フェーズで他者の回答をサーバが返さないことを pin する。
 *
 * Why UI テストではなくここか: 「先入観より先に自分の考えを持つ」は授業の構造であり、
 *   オーバーレイで隠すだけでは (読み込みの一瞬 / DevTools で) 破れる。
 *   データを渡さないことが唯一の保証なので、その保証をここで固定する。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { gasResponseStubs } = require('./_helpers.cjs');

const SOURCE = fs.readFileSync(path.resolve(__dirname, '../src/DataApis.js'), 'utf8');

function loadContext(lessonPhase) {
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    URL,
    ...gasResponseStubs(),
    getCurrentEmail: () => 'child1@example.com',
    logError_: () => {},
    // 授業モードの現在フェーズ。null なら掲示板モード相当。
    __getViewerLessonPhase_: lessonPhase === undefined ? undefined : () => lessonPhase
  };
  vm.createContext(context);
  vm.runInContext(SOURCE, context, { filename: 'DataApis.js' });
  return context;
}

const ROWS = {
  success: true,
  data: [
    { email: 'child1@example.com', reason: '自分の考え', numericX: 4, numericY: 2 },
    { email: 'child2@example.com', reason: '友達の考え', numericX: 1, numericY: 5 },
    { email: 'child3@example.com', reason: 'もう一人', numericX: 3, numericY: 3 }
  ]
};

test('input フェーズ: 児童には自分の行だけを返す', () => {
  const ctx = loadContext({ screenRole: 'input', phaseIndex: 0, lessonId: 'l1' });
  const out = ctx.__maskOthersDuringInputPhase_(ROWS, 'u1', 'child1@example.com', false, false);
  assert.equal(out.data.length, 1);
  assert.equal(out.data[0].email, 'child1@example.com');
});

test('reinput フェーズ: 同じく自分の行だけ (もう一度考える時間も他者は見せない)', () => {
  const ctx = loadContext({ screenRole: 'reinput', phaseIndex: 3, lessonId: 'l1' });
  const out = ctx.__maskOthersDuringInputPhase_(ROWS, 'u1', 'child1@example.com', false, false);
  assert.equal(out.data.length, 1);
});

test('browse フェーズ: 全員の考えを返す (出会う時間)', () => {
  const ctx = loadContext({ screenRole: 'browse', phaseIndex: 1, lessonId: 'l1' });
  const out = ctx.__maskOthersDuringInputPhase_(ROWS, 'u1', 'child1@example.com', false, false);
  assert.equal(out.data.length, 3);
});

test('教師 (own board) は input フェーズでも全員を見られる (投影して授業を進める)', () => {
  const ctx = loadContext({ screenRole: 'input', phaseIndex: 0, lessonId: 'l1' });
  const out = ctx.__maskOthersDuringInputPhase_(ROWS, 'u1', 'teacher@example.com', true, false);
  assert.equal(out.data.length, 3);
});

test('管理者も対象外', () => {
  const ctx = loadContext({ screenRole: 'input', phaseIndex: 0, lessonId: 'l1' });
  const out = ctx.__maskOthersDuringInputPhase_(ROWS, 'u1', 'admin@example.com', false, true);
  assert.equal(out.data.length, 3);
});

test('授業をしていなければ素通し (掲示板モードに影響しない)', () => {
  const ctx = loadContext(null);
  const out = ctx.__maskOthersDuringInputPhase_(ROWS, 'u1', 'child1@example.com', false, false);
  assert.equal(out.data.length, 3);
});

test('LessonService 未ロードでも素通し (掲示板だけの環境で壊れない)', () => {
  const ctx = loadContext(undefined);
  const out = ctx.__maskOthersDuringInputPhase_(ROWS, 'u1', 'child1@example.com', false, false);
  assert.equal(out.data.length, 3);
});

test('フェーズ判定が例外でも掲示板を壊さない (fail-open で元データを返す)', () => {
  const ctx = loadContext(null);
  ctx.__getViewerLessonPhase_ = () => { throw new Error('boom'); };
  const out = ctx.__maskOthersDuringInputPhase_(ROWS, 'u1', 'child1@example.com', false, false);
  assert.equal(out.data.length, 3);
});

test('元の result を破壊しない (cache に入っている共有オブジェクトを汚さない)', () => {
  const ctx = loadContext({ screenRole: 'input', phaseIndex: 0, lessonId: 'l1' });
  const original = { success: true, data: ROWS.data.slice() };
  const out = ctx.__maskOthersDuringInputPhase_(original, 'u1', 'child1@example.com', false, false);
  assert.equal(out.data.length, 1);
  assert.equal(original.data.length, 3, 'cache 共有の元 result が書き換わっている');
});

// =====================================================================
// 見出し: 授業中はフェーズの問い (回答列ヘッダー「横軸」を出さない)
// =====================================================================

test('授業中の見出しはフェーズの問いになる (「横軸」ではない)', () => {
  const ctx = loadContext({ screenRole: 'browse', question: 'あなたならどうする？', phaseName: '出会う' });
  const out = ctx.__applyLessonHeader_({ success: true, header: '横軸', data: [] }, 'u1');
  assert.equal(out.header, 'あなたならどうする？');
});

test('問いが無ければフェーズ名、授業中でなければ元の見出し', () => {
  const withName = loadContext({ screenRole: 'browse', question: '', phaseName: '出会う' });
  assert.equal(withName.__applyLessonHeader_({ success: true, header: '横軸' }, 'u1').header, '出会う');
  const noLesson = loadContext(null);
  assert.equal(noLesson.__applyLessonHeader_({ success: true, header: '本時の問い' }, 'u1').header, '本時の問い');
  const noService = loadContext(undefined);
  assert.equal(noService.__applyLessonHeader_({ success: true, header: '本時の問い' }, 'u1').header, '本時の問い');
});

test('見出しの上書きは元の result を壊さない (cache 共有オブジェクト)', () => {
  const ctx = loadContext({ screenRole: 'input', question: 'Q' });
  const src = { success: true, header: '横軸', data: [] };
  ctx.__applyLessonHeader_(src, 'u1');
  assert.equal(src.header, '横軸');
});

test('応答に lessonPhase を同梱する (授業中はフェーズ、授業中でなければ null)', () => {
  const ctx = loadContext({ screenRole: 'input', question: 'Q', phaseName: '考える', phaseIndex: 0, lessonId: 'l1' });
  const out = ctx.__applyLessonHeader_({ success: true, header: '横軸', data: [] }, 'u1');
  assert.equal(out.lessonPhase.phaseIndex, 0);
  const none = loadContext(null);
  assert.equal(none.__applyLessonHeader_({ success: true, header: 'h' }, 'u1').lessonPhase, null);
  // buildSafePublishedDataResult は lessonPhase を落とさない
  const built = ctx.buildSafePublishedDataResult(out, { displaySettings: { showNames: false } }, { isAdmin: false, isOwnBoard: false });
  assert.equal(built.lessonPhase.phaseIndex, 0);
});
