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

function loadApp(opts = {}) {
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
    setTimeout: opts.setTimeout || (() => 1), clearTimeout: noop, setInterval: () => 0, clearInterval: noop,
    requestAnimationFrame: () => 0, cancelAnimationFrame: noop,
    requestIdleCallback: () => 0, cancelIdleCallback: noop,
    document: opts.document || {
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

test('送っていない書きかけがあるかの判定 (送信済み・空の下書きは数えない)', () => {
  const { app } = loadApp();
  app.state.lessonPhase = think;
  app.state.lessonDraft = { numericX: 3, numericY: 3, reason: '途中まで' };
  app.__persistLessonDraft();
  assert.equal(app.__hasUnsentLessonDraft(think), true);

  // 送信済みなら数えない
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

test('児童: フェーズが変わると state の下書きは外れるが端末には残り、「まだ送れていません」が残る', () => {
  const { app, store } = loadApp();
  app.state.lessonPhase = think;
  app.state.lessonDraft = { numericX: 1, numericY: 1, reason: '書きかけ' };
  app.__persistLessonDraft();
  app.__renderLessonScreen = () => {};
  app.__applyLessonPhase(meet);
  assert.equal(app.state.lessonDraft, null);
  assert.ok(store.has('lessonDraft:u1:L1:0'), '端末には残る');
  const note = app.state.lessonUnsentNote;
  assert.equal(JSON.stringify(note), JSON.stringify({ lessonId: 'L1', phaseIndex: 0, phaseName: '考える', replaced: false }));
  // トーストは「進みました」の 1 回だけ。以前は「残しています」がこれに上書きされて見えなかった (2026-10-28)。
  assert.deepEqual(Array.from(app.toasts), ['「出会う」に進みました']);
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

// =====================================================================
// 「まだ送れていません」案内 (2026-10-28)
//   送る前に教師が進めると、児童からは書いたものが消えたようにしか見えなかった。
//   事実 (送れていない / 書いたことは端末にある) だけを次の画面に残す。
// =====================================================================

const discuss = { lessonId: 'L1', phaseIndex: 2, screenRole: 'discuss', phaseName: '議論する' };
const reflect = { lessonId: 'L1', phaseIndex: 4, screenRole: 'reflect', phaseName: 'ふりかえる' };

function withUnsentThink(app) {
  app.state.lessonPhase = think;
  app.state.lessonDraft = { numericX: 2, numericY: 2, reason: '途中' };
  app.__persistLessonDraft();
  app.__renderLessonScreen = () => {};
}

test('案内は見る・話すフェーズのあいだ持ち越し、書く画面に入ったら外れる', () => {
  const { app } = loadApp();
  withUnsentThink(app);
  app.__applyLessonPhase(meet);
  app.__applyLessonPhase(discuss);
  assert.equal(app.state.lessonUnsentNote && app.state.lessonUnsentNote.phaseIndex, 0, '議論するでも出ている');
  app.__applyLessonPhase(rethink);
  assert.equal(app.state.lessonUnsentNote, null, 'もう一度考えるでは新しく書くので外す');
});

test('教師が案内したフェーズに戻したら外れる (書きかけは画面に復元される)', () => {
  const { app } = loadApp();
  withUnsentThink(app);
  app.__applyLessonPhase(meet);
  app.__applyLessonPhase(think);
  assert.equal(app.state.lessonUnsentNote, null);
});

test('考える → もう一度考える と飛んだときは、考えるの案内を出す (● が無い理由が分かる)', () => {
  const { app } = loadApp();
  withUnsentThink(app);
  app.__applyLessonPhase(rethink);
  assert.equal(app.state.lessonUnsentNote && app.state.lessonUnsentNote.phaseName, '考える');
});

test('もう一度考えるで送らずに進んだら、ふりかえるでその案内を出す', () => {
  const { app } = loadApp();
  app.state.lessonPhase = rethink;
  app.state.lessonDraft = { numericX: 4, numericY: 4, reason: 'まだ' };
  app.__persistLessonDraft();
  app.__renderLessonScreen = () => {};
  app.__applyLessonPhase(reflect);
  assert.equal(app.state.lessonUnsentNote && app.state.lessonUnsentNote.phaseName, 'もう一度考える');
});

test('送ったあとの置き直しだけが送れていないときは、replaced として区別する', () => {
  const { app } = loadApp();
  app.state.lessonPhase = think;
  app.state.lessonSent = { lessonId: 'L1', phaseIndex: 0 };
  app.state.lessonEditing = true;
  app.state.lessonDraft = { numericX: 5, numericY: 5, reason: '置き直し中' };
  app.__persistLessonDraft();
  app.__renderLessonScreen = () => {};
  app.__applyLessonPhase(meet);
  assert.equal(app.state.lessonUnsentNote && app.state.lessonUnsentNote.replaced, true);
});

test('送ってから進んだとき・教師の画面・授業の終わりには案内を出さない', () => {
  const sent = loadApp().app;
  sent.state.lessonPhase = think;
  sent.state.lessonSent = { lessonId: 'L1', phaseIndex: 0 };
  sent.__renderLessonScreen = () => {};
  sent.__applyLessonPhase(meet);
  assert.equal(sent.state.lessonUnsentNote, null, '送信済み');

  const teacher = loadApp().app;
  withUnsentThink(teacher);
  teacher.state.isEditor = true;
  teacher.__syncBoardPhaseNav = () => {};
  teacher.__applyLessonPhase(meet);
  assert.equal(teacher.state.lessonUnsentNote, null, '教師');

  const ended = loadApp().app;
  withUnsentThink(ended);
  ended.__applyLessonPhase(meet);
  ended.__teardownLessonScreen = () => {};
  ended.__applyLessonPhase(null);
  assert.equal(ended.state.lessonUnsentNote, null, '授業の終わり (書きかけも消える)');
});

// 描画の確認用の最小 DOM。id で引けて、子要素と click を持つだけ。
function miniDocument() {
  const byId = new Map();
  const bodyClasses = new Set();
  const make = (tag) => {
    const node = {
      tagName: tag, children: [], listeners: {}, attrs: {}, parentNode: null, className: '', type: '',
      _text: '',
      get textContent() { return this._text + this.children.map((c) => c.textContent).join(''); },
      set textContent(v) { this._text = String(v); this.children = []; },
      get id() { return this.attrs.id || ''; },
      set id(v) { this.attrs.id = v; byId.set(v, this); },
      setAttribute(k, v) { this.attrs[k] = v; },
      appendChild(c) { c.parentNode = this; this.children.push(c); return c; },
      removeChild(c) { this.children = this.children.filter((x) => x !== c); c.parentNode = null; byId.delete(c.id); return c; },
      addEventListener(type, fn) { this.listeners[type] = fn; }
    };
    return node;
  };
  const body = make('body');
  body.classList = { add: (c) => bodyClasses.add(c), remove: (c) => bodyClasses.delete(c), contains: (c) => bodyClasses.has(c) };
  return {
    body, bodyClasses,
    createElement: make,
    getElementById: (id) => byId.get(id) || null,
    querySelector: () => null, querySelectorAll: () => [],
    addEventListener: () => {}, removeEventListener: () => {}, contains: () => true
  };
}

test('案内の文言: 送れていないこと・端末に残っていることだけを書き、教師の操作は予告しない', () => {
  const doc = miniDocument();
  const { app } = loadApp({ document: doc });
  app.state.lessonUnsentNote = { lessonId: 'L1', phaseIndex: 0, phaseName: '考える', replaced: false };
  app.__renderLessonUnsentNote();
  const el = doc.getElementById('lessonUnsentNote');
  assert.ok(el, '画面に出る');
  assert.equal(el.attrs.role, 'status');
  assert.match(el.textContent, /「考える」の考えは、まだ送れていません/);
  assert.match(el.textContent, /書いたことは、この端末に残っています/);
  assert.doesNotMatch(el.textContent, /先生/);
  assert.ok(doc.bodyClasses.has('lesson-has-note'), '下に余白を足す');

  // 置き直しは戻っても復元されないので「残っています」とは言わない
  app.state.lessonUnsentNote = { lessonId: 'L1', phaseIndex: 0, phaseName: '考える', replaced: true };
  app.__renderLessonUnsentNote();
  assert.match(el.textContent, /置き直しは、まだ送れていません/);
  assert.match(el.textContent, /前に送った考えは届いています/);
  assert.doesNotMatch(el.textContent, /端末に残っています/);

  // 「とじる」で消える
  const close = el.children.find((c) => c.tagName === 'button');
  close.listeners.click();
  assert.equal(app.state.lessonUnsentNote, null);
  assert.equal(doc.getElementById('lessonUnsentNote'), null);
  assert.equal(doc.bodyClasses.has('lesson-has-note'), false);
});

test('送る前に進んでいた (PHASE_MISMATCH): 読み込み直しではなく、端末に残っていることを伝える', async () => {
  const doc = miniDocument();
  const reasonEl = doc.createElement('textarea');
  reasonEl.id = 'lessonReason';
  reasonEl.value = '理由';
  const { app } = loadApp({ document: doc, setTimeout: (fn) => { fn(); return 1; } });
  app.state.lessonPhase = think;
  app.state.lessonDraft = { numericX: 2, numericY: 3, reason: '理由' };
  app.__lessonClass = () => '';
  app.__lessonStoredName = () => 'あおい';
  app.__reportLessonSubmit = () => {};
  app.runGas = async () => ({ success: false, error: 'PHASE_MISMATCH', message: 'PHASE_CHANGED: 送る前にフェーズが進んだので、送れませんでした' });
  await app.__submitLessonAnswer();
  assert.deepEqual(Array.from(app.toasts), ['送る前に次へ進んだので、送れませんでした。書いたことは、この端末に残っています']);

  // 置き直しの途中だったら、届いている方を伝える
  app.toasts.length = 0;
  app.state.lessonSent = { lessonId: 'L1', phaseIndex: 0 };
  await app.__submitLessonAnswer();
  assert.deepEqual(Array.from(app.toasts), ['送る前に次へ進んだので、置き直しは送れませんでした。前に送った考えは届いています']);

  // ほかの拒否は、機械コードの接頭辞を外してサーバの文言を出す
  app.toasts.length = 0;
  app.runGas = async () => ({ success: false, error: 'LESSON_MISMATCH', message: 'PHASE_CHANGED: 授業が切り替わりました。画面を読み込み直してください' });
  await app.__submitLessonAnswer();
  assert.deepEqual(Array.from(app.toasts), ['授業が切り替わりました。画面を読み込み直してください']);
});
