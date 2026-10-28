/**
 * 管理パネルの授業画面: ボードにつながっていない授業は、切り替えるまで送れない・終了できない。
 *
 * Why: 開始中の授業が 2 つあるとき、つながっていない方のフェーズを送るとボードがそちらへ
 *   移っていた (2026-10-28、理科の授業中に道徳を送り、児童の画面が 10 分道徳になった)。
 *   サーバも NOT_ON_BOARD で断るが、画面でも「いまボードが映している授業」を示して押させない。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const SRC = fs.readFileSync(path.resolve(__dirname, '../src/AdminPanel.js.html'), 'utf8');
const BLOCK = SRC.slice(SRC.indexOf('// ---- lesson board link: begin ----'), SRC.indexOf('// ---- lesson board link: end ----'));

function harness({ onBoard, list }) {
  const box = { classes: new Set(['hidden']), classList: null };
  box.classList = { toggle: (c, on) => { if (on) box.classes.add(c); else box.classes.delete(c); } };
  const els = {
    'lesson-runner-board-link': box,
    'lesson-runner-board-link-text': { textContent: '' },
    'lesson-runner-end': { disabled: false }
  };
  const ctx = { document: { getElementById: (id) => els[id] || null } };
  vm.createContext(ctx);
  vm.runInContext(
    `let currentLessonOnBoard = ${JSON.stringify(onBoard)}; let lastLessonList = ${JSON.stringify(list)};\n` +
    BLOCK + '\nthis.api = { boardLinkOf, renderRunnerBoardLink };', ctx);
  return { api: ctx.api, els, box };
}

test('boardLinkOf: 応答の boardLessonId と比べる (無い・null は「分からない」)', () => {
  const { api } = harness({ onBoard: null, list: [] });
  assert.equal(api.boardLinkOf('a', { boardLessonId: 'a' }), true);
  assert.equal(api.boardLinkOf('a', { boardLessonId: 'b' }), false);
  assert.equal(api.boardLinkOf('a', { boardLessonId: '' }), false);
  assert.equal(api.boardLinkOf('a', { boardLessonId: null }), null);
  assert.equal(api.boardLinkOf('a', { lesson: {} }), null);
});

test('つながっていない授業: 案内にボードが映している授業名を出し、終了を押せなくする', () => {
  const h = harness({ onBoard: false, list: [{ lessonId: 'sci', name: '理科', onBoard: true }] });
  h.api.renderRunnerBoardLink({ lessonId: 'moral', name: '道徳' });
  assert.equal(h.box.classes.has('hidden'), false);
  assert.match(h.els['lesson-runner-board-link-text'].textContent, /「理科」を映しています/);
  assert.equal(h.els['lesson-runner-end'].disabled, true);
});

test('つながっている / 分からない授業: 案内を出さず、終了も押せる', () => {
  for (const onBoard of [true, null]) {
    const h = harness({ onBoard, list: [] });
    h.api.renderRunnerBoardLink({ lessonId: 'sci', name: '理科' });
    assert.equal(h.box.classes.has('hidden'), true, `onBoard=${onBoard}`);
    assert.equal(h.els['lesson-runner-end'].disabled, false, `onBoard=${onBoard}`);
  }
});
