/**
 * 授業モード (native の授業が実行中) ではリアクションが存在しない。
 *
 * Why: 数を隠しても色のリングで「反応あり／なし」の二値と早い者勝ちの増幅は残る。
 *   設定は増やさず、lessonPhase の有無で機能ごと出さない。掲示板モードは従来どおり。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

// ---- page.js: ボタン生成とカードの色 ----
function loadApp() {
  const src = fs.readFileSync(path.resolve(__dirname, '../src/page.js.html'), 'utf8');
  const js = src.match(/<script[^>]*>([\s\S]*?)<\/script>/)[1];
  const noop = () => {};
  const el = () => ({ classList: { add: noop, remove: noop, contains: () => false }, setAttribute: noop, querySelector: () => null, querySelectorAll: () => [], addEventListener: noop, removeEventListener: noop, appendChild: noop, remove: noop, dataset: {}, style: {}, textContent: '', innerHTML: '' });
  class HTMLElement { constructor() { this.classes = new Set(); this.classList = { add: (c) => this.classes.add(c), remove: (c) => this.classes.delete(c), contains: (c) => this.classes.has(c) }; } }
  const ctx = {
    console: { log: noop, warn: noop, error: noop }, HTMLElement,
    Map, Set, Date, Math, JSON, Number, Array, Object, Promise, Error, String, Boolean, Symbol, WeakMap, WeakSet, Reflect, Proxy, Buffer,
    setTimeout: () => 1, clearTimeout: noop, setInterval: () => 0, clearInterval: noop, requestAnimationFrame: () => 0, cancelAnimationFrame: noop, requestIdleCallback: () => 0, cancelIdleCallback: noop,
    document: { createElement: el, createDocumentFragment: () => ({ appendChild: noop }), getElementById: () => null, querySelector: () => null, querySelectorAll: () => [], addEventListener: noop, removeEventListener: noop, contains: () => true, body: { classList: { add: noop, remove: noop }, appendChild: noop } },
    window: { addEventListener: noop, removeEventListener: noop, location: { reload: noop, href: '', search: '' }, sharedUtilities: { security: { escapeHtml: (s) => String(s || '') } }, UNIFIED_CONFIG: {}, notifications: null },
    navigator: { userAgent: 'test' }, localStorage: { getItem: () => null, setItem: noop, removeItem: noop, key: () => null, length: 0 },
    sessionStorage: { getItem: () => null, setItem: noop, removeItem: noop }, google: { script: { run: {} } }
  };
  ctx.globalThis = ctx; ctx.self = ctx;
  vm.createContext(ctx);
  vm.runInContext(js, ctx, { filename: 'page.js.html' });
  const App = vm.runInContext('StudyQuestApp', ctx);
  const app = Object.create(App.prototype);
  app.state = { userId: 'u1', isEditor: false, lessonPhase: null, showCounts: false, showAdminFeatures: false };
  app.reactionTypes = [{ key: 'LIKE', icon: 'hand-thumb-up' }, { key: 'UNDERSTAND', icon: 'lightbulb' }, { key: 'CURIOUS', icon: 'magnifying-glass-plus' }];
  app.getIcon = (name) => '<svg data-icon="' + name + '"></svg>';
  return { app, ctx };
}

const phase = { lessonId: 'L1', phaseIndex: 1, screenRole: 'browse' };
const row = { rowIndex: 2, reactions: { LIKE: { count: 3, reacted: false } }, highlight: false };

test('掲示板モード (授業なし): リアクションボタンが 3 つ出る', () => {
  const { app } = loadApp();
  const html = app.__buildAnswerActionsHtml(row, { size: 'card', reactionsOnly: true });
  assert.equal((html.match(/reaction-btn/g) || []).length, 3);
});

test('授業モード: リアクションボタンが出ない (教師の ☆ は残る)', () => {
  const { app } = loadApp();
  app.state.lessonPhase = phase;
  assert.equal(app.__buildAnswerActionsHtml(row, { size: 'card', reactionsOnly: true }), '');
  app.state.isEditor = true;
  const html = app.__buildAnswerActionsHtml(row, { size: 'modal' });
  assert.equal((html.match(/reaction-btn/g) || []).length, 0);
  assert.equal((html.match(/highlight-btn/g) || []).length, 1);
});

test('授業モード: カードにリアクションの色を付けない (☆ は付く)', () => {
  const { app, ctx } = loadApp();
  const card = new ctx.HTMLElement();
  app.applyReactionStyles(card, row);
  assert.ok(card.classes.has('reaction-bg-like') && card.classes.has('reaction-border-1'), '授業なしでは色が付く');

  app.state.lessonPhase = phase;
  const card2 = new ctx.HTMLElement();
  app.applyReactionStyles(card2, row);
  assert.equal(card2.classes.size, 0);
  const card3 = new ctx.HTMLElement();
  app.applyReactionStyles(card3, Object.assign({}, row, { highlight: true }));
  assert.ok(card3.classes.has('highlighted'));
});

// ---- page.viz.js: 点のリングと数バッジ ----
function loadViz() {
  const html = fs.readFileSync(path.resolve(__dirname, '../src/page.viz.js.html'), 'utf8');
  const source = html.match(/<script>([\s\S]*?)<\/script>/)[1];
  const StudyQuestApp = class StudyQuestApp {};
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    window: { StudyQuestApp, location: { search: '' } }, StudyQuestApp,
    document: { readyState: 'complete', addEventListener: () => {}, getElementById: () => null, createElement: () => ({ classList: { add: () => {}, remove: () => {}, toggle: () => {} }, addEventListener: () => {}, setAttribute: () => {}, appendChild: () => {}, style: {} }), createElementNS: () => ({ setAttribute: () => {}, appendChild: () => {} }), body: { classList: { add: () => {} } } },
    URLSearchParams, Map, Set, Promise, JSON, Math, Number, Object, Array, String, Boolean, Date, RegExp, Error, TypeError, parseInt, parseFloat, isNaN, isFinite,
    setTimeout: (fn, ms) => setTimeout(fn, ms), clearTimeout: (id) => clearTimeout(id)
  };
  vm.createContext(context);
  vm.runInContext(source, context, { filename: 'page.viz.js.html' });
  return { StudyQuestApp };
}

// d3 selection の最小 stub: style(name, fn) を各 datum に適用して記録する
function fakeSelection(data) {
  const styles = data.map(() => ({}));
  const sel = {
    style(name, fn) { data.forEach((d, i) => { styles[i][name] = typeof fn === 'function' ? fn(d) : fn; }); return sel; },
    styles
  };
  return sel;
}

test('点のリング: 授業モードでは反応があってもリングを描かない', () => {
  const { StudyQuestApp } = loadViz();
  const proto = StudyQuestApp.prototype;
  const data = [{ data: { rowIndex: 2, reactions: { LIKE: { count: 2 } }, highlight: false } }];
  const on = fakeSelection(data);
  proto.__applyReactionRingStyles(on, { __lessonReactionsOff: () => false });
  assert.ok(on.styles[0].stroke, '授業なしではリングが付く');

  const off = fakeSelection(data);
  proto.__applyReactionRingStyles(off, { __lessonReactionsOff: () => true });
  assert.equal(off.styles[0].stroke, null);
  assert.equal(off.styles[0]['stroke-width'], null);
  assert.equal(off.styles[0]['stroke-dasharray'], null);
});

test('数バッジ: 授業モードでは showCounts=true でも描かない', () => {
  const { StudyQuestApp } = loadViz();
  const proto = StudyQuestApp.prototype;
  const nodes = [{ x: 1, y: 1, data: { rowIndex: 2, reactions: { LIKE: { count: 2 } } } }];
  let bound = null;
  const g = { selectAll: () => ({ data: (arr) => { bound = arr; return { exit: () => ({ remove: () => {} }), enter: () => ({ append: () => ({ attr() { return this; }, merge() { return this; }, text() { return this; } }) }) }; }, raise: () => {} }) };
  proto.__renderDotLabels(g, nodes, { showCounts: true }, { __lessonReactionsOff: () => false });
  assert.equal(bound.length, 1);
  proto.__renderDotLabels(g, nodes, { showCounts: true }, { __lessonReactionsOff: () => true });
  assert.equal(bound.length, 0);
});
