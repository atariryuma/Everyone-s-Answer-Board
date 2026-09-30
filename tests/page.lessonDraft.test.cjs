/**
 * 授業モードの書きかけ (下書き) の端末保持。
 *
 * Why: 教師が進めた瞬間に送った児童が拒否され、書き直しになった (2026-09-30 の PHASE_MISMATCH)。
 *   教師はその場でフェーズを戻して回収するので、戻ったときに続きから送れればよい。
 *   サーバには何も送らない (1 文字ごとの Sheets 書きは 429 の原因になる)。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function loadApp() {
  const src = fs.readFileSync(path.resolve(__dirname, '../src/page.js.html'), 'utf8');
  const js = src.match(/<script[^>]*>([\s\S]*?)<\/script>/)[1];
  const store = new Map();
  const noop = () => {};
  const el = () => ({
    classList: { add: noop, remove: noop, contains: () => false },
    setAttribute: noop, querySelector: () => null, querySelectorAll: () => [],
    addEventListener: noop, removeEventListener: noop, appendChild: noop, remove: noop,
    dataset: {}, style: {}, textContent: '', innerHTML: ''
  });
  const ctx = {
    console: { log: noop, warn: noop, error: noop },
    Map, Set, Date, Math, JSON, Number, Array, Object, Promise, Error, String, Boolean,
    Symbol, WeakMap, WeakSet, Reflect, Proxy, Buffer,
    setTimeout: () => 1, clearTimeout: noop, setInterval: () => 0, clearInterval: noop,
    requestAnimationFrame: () => 0, cancelAnimationFrame: noop,
    requestIdleCallback: () => 0, cancelIdleCallback: noop,
    document: {
      createElement: el, createDocumentFragment: () => ({ appendChild: noop }),
      getElementById: () => null, querySelector: () => null, querySelectorAll: () => [],
      addEventListener: noop, removeEventListener: noop, contains: () => true,
      body: { classList: { add: noop, remove: noop } }
    },
    window: {
      addEventListener: noop, removeEventListener: noop,
      location: { reload: noop, href: '', search: '' },
      sharedUtilities: { security: { escapeHtml: (s) => String(s || '') } },
      UNIFIED_CONFIG: {}, notifications: null
    },
    navigator: { userAgent: 'test', hardwareConcurrency: 4 },
    localStorage: {
      getItem: (k) => (store.has(k) ? store.get(k) : null),
      setItem: (k, v) => store.set(k, String(v)),
      removeItem: (k) => store.delete(k),
      key: (i) => Array.from(store.keys())[i] || null,
      get length() { return store.size; }
    },
    sessionStorage: { getItem: () => null, setItem: noop, removeItem: noop },
    google: { script: { run: {} } }
  };
  ctx.globalThis = ctx; ctx.self = ctx;
  vm.createContext(ctx);
  vm.runInContext(js, ctx, { filename: 'page.js.html' });
  const App = vm.runInContext('StudyQuestApp', ctx);
  const app = Object.create(App.prototype);
  app.state = { userId: 'u1', lessonDraft: null, lessonPhase: null, lessonSent: null, lessonEditing: false, isEditor: false };
  app.toasts = [];
  app.showToast = (m) => app.toasts.push(m);
  app.loadSheetData = async () => {};   // 出会うに入った瞬間の読み直しは対象外
  return { app, store };
}

const think = { lessonId: 'L1', phaseIndex: 0, screenRole: 'input', phaseName: '考える' };
const meet = { lessonId: 'L1', phaseIndex: 1, screenRole: 'browse', phaseName: '出会う' };
const rethink = { lessonId: 'L1', phaseIndex: 3, screenRole: 'reinput', phaseName: 'もう一度考える' };

test('書きかけは 授業 × フェーズ の key で端末に残り、同じフェーズに戻ると復元される', () => {
  const { app, store } = loadApp();
  app.state.lessonPhase = think;
  app.state.lessonDraft = { numericX: 2, numericY: 4, reason: '書きかけの理由' };
  app.__persistLessonDraft();
  assert.ok(store.has('lessonDraft:u1:L1:0'));

  // vm の realm が違うので deepEqual ではなく JSON で比べる
  assert.equal(JSON.stringify(app.__restoreLessonDraft(think)), JSON.stringify({ numericX: 2, numericY: 4, reason: '書きかけの理由', addedInsight: '' }));
  assert.equal(app.__restoreLessonDraft(rethink), null, '別のフェーズには出ない');
  assert.equal(app.__restoreLessonDraft({ lessonId: 'L2', phaseIndex: 0 }), null, '別の授業には出ない');
});

test('見るだけ・話すフェーズでは保存しない (input / reinput だけ)', () => {
  const { app, store } = loadApp();
  app.state.lessonPhase = meet;
  app.state.lessonDraft = { reason: 'x' };
  app.__persistLessonDraft();
  assert.equal(store.size, 0);
});

test('送れたら端末の書きかけは消える (記録はサーバにある)', () => {
  const { app, store } = loadApp();
  app.state.lessonPhase = rethink;
  app.state.lessonDraft = { numericX: 5, numericY: 1, reason: 'r', addedInsight: 'i' };
  app.__persistLessonDraft();
  assert.equal(store.size, 1);
  app.__clearLessonDraft(rethink);
  assert.equal(store.size, 0);
});

test('教師が進めたとき、送っていない書きかけがあれば児童に「残しています」と伝える', () => {
  const { app } = loadApp();
  app.state.lessonPhase = think;
  app.state.lessonDraft = { numericX: 3, numericY: 3, reason: '途中まで' };
  app.__persistLessonDraft();
  assert.equal(app.__hasUnsentLessonDraft(think), true);

  // 送信済みなら「残しています」は出さない
  app.state.lessonSent = { lessonId: 'L1', phaseIndex: 0 };
  assert.equal(app.__hasUnsentLessonDraft(think), false);
  app.state.lessonSent = null;

  // 位置もことばも無い空の下書きは数えない
  app.state.lessonDraft = { reason: '   ' };
  app.__persistLessonDraft();
  assert.equal(app.__hasUnsentLessonDraft(think), false);
});

test('教師は対象外: 教師の画面ではフェーズが変わっても通知しない (isEditor)', () => {
  const { app } = loadApp();
  app.state.isEditor = true;
  app.state.lessonPhase = think;
  app.state.lessonDraft = { numericX: 1, numericY: 1, reason: 'x' };
  app.__persistLessonDraft();
  app.__renderLessonScreen = () => {};
  app.__syncBoardPhaseNav = () => {};
  app.__applyLessonPhase(meet);
  assert.equal(app.toasts.length, 0);
});

test('児童: フェーズが変わると state の下書きは外れるが端末には残り、通知が 1 回出る', () => {
  const { app, store } = loadApp();
  app.state.lessonPhase = think;
  app.state.lessonDraft = { numericX: 1, numericY: 1, reason: '書きかけ' };
  app.__persistLessonDraft();
  app.__renderLessonScreen = () => {};
  app.__applyLessonPhase(meet);
  assert.equal(app.state.lessonDraft, null);
  assert.ok(store.has('lessonDraft:u1:L1:0'), '端末には残る');
  assert.ok(app.toasts.some((m) => m.indexOf('書きかけは残しています') >= 0));
});

test('授業が終わると、その授業の書きかけだけを端末から消す', () => {
  const { app, store } = loadApp();
  store.set('lessonDraft:u1:L1:0', '{}');
  store.set('lessonDraft:u1:L1:3', '{}');
  store.set('lessonDraft:u1:L9:0', '{}');
  store.set('lessonName:u1', 'あおい');
  app.state.lessonPhase = rethink;
  app.__teardownLessonScreen = () => {};
  app.__applyLessonPhase(null);
  assert.deepEqual(Array.from(store.keys()).sort(), ['lessonDraft:u1:L9:0', 'lessonName:u1']);
});
