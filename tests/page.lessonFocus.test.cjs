/**
 * 議論する (教師の投影): ☆ の理由を大きく並べる。
 *
 * Why: 「出会う」で分布を見せたあと、議論に入るときに黒板に残るものが無かった。
 *   教師が ☆ を押す行為を「議論の焦点を決める」操作にする (押すものは増えない)。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function loadApp() {
  const src = fs.readFileSync(path.resolve(__dirname, '../src/page.js.html'), 'utf8');
  const js = src.match(/<script[^>]*>([\s\S]*?)<\/script>/)[1];
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
    requestAnimationFrame: () => 0, cancelAnimationFrame: noop, requestIdleCallback: () => 0, cancelIdleCallback: noop,
    document: {
      createElement: el, createDocumentFragment: () => ({ appendChild: noop }),
      getElementById: () => null, querySelector: () => null, querySelectorAll: () => [],
      addEventListener: noop, removeEventListener: noop, contains: () => true,
      body: { classList: { add: noop, remove: noop }, appendChild: noop }
    },
    window: {
      addEventListener: noop, removeEventListener: noop, location: { reload: noop, href: '', search: '' },
      sharedUtilities: { security: { escapeHtml: (s) => String(s || '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])) } },
      UNIFIED_CONFIG: {}, notifications: null
    },
    navigator: { userAgent: 'test', hardwareConcurrency: 4 },
    localStorage: { getItem: () => null, setItem: noop, removeItem: noop, key: () => null, length: 0 },
    sessionStorage: { getItem: () => null, setItem: noop, removeItem: noop },
    google: { script: { run: {} } }
  };
  ctx.globalThis = ctx; ctx.self = ctx;
  vm.createContext(ctx);
  vm.runInContext(js, ctx, { filename: 'page.js.html' });
  const App = vm.runInContext('StudyQuestApp', ctx);
  const app = Object.create(App.prototype);
  app.state = { userId: 'u1', isEditor: true, lessonPhase: null, lessonTransition: null, currentAnswers: [], boardMode: 'matrix', axisConfig: { xAxisLabels: { min: '言う', max: '言わない' }, yAxisLabels: { min: '迷いあり', max: '迷いなし' } } };
  app.calls = [];
  app.__renderLessonTeacherWait = () => app.calls.push('wait');
  app.__teardownLessonScreen = () => app.calls.push('teardown');
  app.__renderBoardPhaseNav = () => {};
  return { app, ctx };
}

const discuss = { lessonId: 'L1', phaseIndex: 2, screenRole: 'discuss', phaseName: '議論する', question: '考えがちがう友達と、理由をくらべて話そう' };
const rows = [
  { rowIndex: 2, numericX: 1, numericY: 2, reason: '自首を勧める。うそは相手をきずつけるから', highlight: true, name: 'あおい', class: '6年1組' },
  { rowIndex: 3, numericX: 3, numericY: 3, reason: '迷う', highlight: false },
  { rowIndex: 4, numericX: 5, numericY: 4, reason: '逃がす <b>', highlight: true }
];

test('☆ の行だけが焦点になる', () => {
  const { app } = loadApp();
  app.state.currentAnswers = rows;
  assert.deepEqual(app.__lessonFocusRows().map((r) => r.rowIndex), [2, 4]);
});

test('議論する + ☆あり: 教師の画面は焦点の一覧になり、分布 (teardown) にはしない', () => {
  const { app } = loadApp();
  app.state.lessonPhase = discuss;
  app.state.currentAnswers = rows;
  let focus = 0;
  app.__renderLessonDiscussFocus = () => { focus++; };
  app.__renderLessonScreen();
  assert.equal(focus, 1);
  assert.deepEqual(app.calls, []);
});

test('議論する + ☆なし: 従来どおり分布のまま (何も変わらない)', () => {
  const { app } = loadApp();
  app.state.lessonPhase = discuss;
  app.state.currentAnswers = rows.map((r) => Object.assign({}, r, { highlight: false }));
  app.__renderLessonScreen();
  assert.deepEqual(app.calls, ['teardown']);
});

test('児童の画面は変わらない (議論するは「画面をとじて、話そう」のまま)', () => {
  const { app } = loadApp();
  app.state.isEditor = false;
  app.state.lessonPhase = discuss;
  app.state.currentAnswers = rows;
  let quiet = 0, focus = 0;
  app.__renderLessonDiscuss = () => { quiet++; };
  app.__renderLessonDiscussFocus = () => { focus++; };
  app.__renderLessonScreen();
  assert.equal(quiet, 1);
  assert.equal(focus, 0);
});

test('データ更新のたびに焦点を描き直し、☆ が 0 になれば分布に戻す', () => {
  const { app } = loadApp();
  app.state.lessonPhase = discuss;
  app.state.currentAnswers = rows;
  let focus = 0;
  app.__renderLessonDiscussFocus = () => { focus++; };
  app.__refreshLessonTeacherWait();
  assert.equal(focus, 1);
  app.state.currentAnswers = rows.map((r) => Object.assign({}, r, { highlight: false }));
  app.__refreshLessonTeacherWait();
  assert.deepEqual(app.calls, ['teardown']);
});

test('焦点の HTML: 理由と軸の位置は出すが、名前・クラスは出さない。5 件目以降は数だけ', () => {
  const { app, ctx } = loadApp();
  app.state.lessonPhase = discuss;
  const many = [];
  for (let i = 0; i < 6; i++) many.push({ rowIndex: i + 2, numericX: 1 + (i % 5), numericY: 3, reason: '理由' + i, highlight: true, name: '名前' + i, class: '6年1組' });
  app.state.currentAnswers = many;
  let html = '';
  ctx.document.getElementById = (id) => (id === 'lessonScreen' ? { set innerHTML(v) { html = v; }, get innerHTML() { return html; } } : null);
  app.__renderLessonDiscussFocus(discuss);
  assert.ok(html.indexOf('理由0') >= 0 && html.indexOf('理由3') >= 0, '4 件目まで出る');
  assert.ok(html.indexOf('理由4') < 0, '5 件目は出ない');
  assert.ok(html.indexOf('ほか 2 件') >= 0);
  assert.ok(html.indexOf('名前0') < 0 && html.indexOf('6年1組') < 0, '名前とクラスは出さない');
  assert.ok(html.indexOf('言わない') >= 0, '軸ラベルつきの位置バーが出る');
  assert.ok(html.indexOf('is-dense') >= 0, '3 件以上は 2 列');
  assert.ok(html.indexOf('lesson-control-host') >= 0, '「次へ」の操作面がある');
});

test('焦点の HTML: 理由は escape される', () => {
  const { app, ctx } = loadApp();
  app.state.lessonPhase = discuss;
  app.state.currentAnswers = rows;
  let html = '';
  ctx.document.getElementById = (id) => (id === 'lessonScreen' ? { set innerHTML(v) { html = v; }, get innerHTML() { return html; } } : null);
  app.__renderLessonDiscussFocus(discuss);
  assert.ok(html.indexOf('&lt;b&gt;') >= 0);
  assert.ok(html.indexOf('<b>') < 0);
});

test('版番号の変化: 教師は焦点画面がかぶさっていても読み直す (☆ の増減を映す)、児童は読まない', () => {
  const { app, ctx } = loadApp();
  ctx.document.getElementById = (id) => (id === 'lessonScreen' ? {} : null);
  let loads = 0;
  app.loadSheetData = async () => { loads++; };
  app.state.lastBoardVersion = '1';
  assert.equal(app.__handleBoardVersion({ boardVersion: '2', hasNewContent: false }, false), true);
  assert.equal(loads, 1);
  app.state.isEditor = false;
  app.state.lastBoardVersion = '2';
  assert.equal(app.__handleBoardVersion({ boardVersion: '3', hasNewContent: false }, false), false);
  assert.equal(loads, 1);
});

// =====================================================================
// 授業の帯 (教師の投影): 5 フェーズの今ここ + いま児童の画面では
// Why: 投影からは構造 (遮断・停止・匿名・前後比較) が見えない。参観者にも読めるようにする。
// =====================================================================

function fakeEl(tag) {
  const el = {
    tag, className: '', textContent: '', children: [], attrs: {},
    classList: { add(c) { el.className += ' ' + c; }, remove(c) { el.className = el.className.replace(c, ''); }, contains: () => false },
    setAttribute(k, v) { el.attrs[k] = v; },
    appendChild(ch) { el.children.push(ch); return ch; }
  };
  return el;
}
function textOf(el) {
  if (typeof el === 'string') return el;
  if (el.text !== undefined) return el.text;
  return (el.textContent || '') + el.children.map(textOf).join('');
}

test('帯: フェーズ名を順に並べ、現在地だけ強調し、問いは出さない', () => {
  const { app, ctx } = loadApp();
  ctx.document.createElement = fakeEl;
  ctx.document.createTextNode = (t) => ({ text: t });
  app.state.boardPhaseNav = {
    lessonId: 'L1', activePhaseIndex: 2,
    phases: [
      { name: '考える', screenRole: 'input', question: '秘密の問い1' },
      { name: '出会う', screenRole: 'browse', question: '秘密の問い2' },
      { name: '議論する', screenRole: 'discuss', question: '秘密の問い3' },
      { name: 'もう一度考える', screenRole: 'reinput', question: '秘密の問い4' },
      { name: 'ふりかえる', screenRole: 'reflect', question: '秘密の問い5' }
    ]
  };
  const node = app.__buildLessonStripNode();
  const list = node.children[0];
  assert.equal(list.children.length, 5);
  assert.ok(list.children[2].className.indexOf('is-current') >= 0);
  assert.equal(list.children[2].attrs['aria-current'], 'step');
  assert.ok(list.children[0].className.indexOf('is-done') >= 0);
  assert.ok(list.children[4].className.indexOf('is-current') < 0 && list.children[4].className.indexOf('is-done') < 0);
  const all = textOf(node);
  assert.ok(all.indexOf('もう一度考える') >= 0);
  assert.ok(all.indexOf('秘密の問い') < 0, '未来の問いは出さない');
  assert.ok(all.indexOf('端末を閉じています') >= 0, '議論する = 端末を閉じている');
});

test('帯: 役割ごとの「いま児童の画面では」が構造を言い切る', () => {
  const { app } = loadApp();
  assert.ok(app.__lessonRoleCaption('input').indexOf('他の人の考えは見えません') >= 0);
  assert.ok(app.__lessonRoleCaption('browse').indexOf('書き込みはできません') >= 0);
  assert.ok(app.__lessonRoleCaption('reinput').indexOf('最初の自分') >= 0);
  assert.ok(app.__lessonRoleCaption('reflect').indexOf('学級の分布は出ません') >= 0);
  assert.equal(app.__lessonRoleCaption('unknown'), '');
});

test('帯: 授業が無ければ空 (何も出さない)', () => {
  const { app, ctx } = loadApp();
  ctx.document.createElement = fakeEl;
  app.state.boardPhaseNav = null;
  assert.equal(app.__buildLessonStripNode().children.length, 0);
});
