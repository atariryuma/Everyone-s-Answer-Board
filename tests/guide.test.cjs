/**
 * ひな形をコピーした先生への案内 (src/Guide.js) のテスト。
 *
 * Why: onOpen と openBoardGuide は名前の都合で公開関数になり、児童の画面からも先生の権限で呼べる。
 *   スプレッドシートの画面の外から呼ばれたとき・焼き込みのある配置では、何も書かないことを確かめる。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const SRC = fs.readFileSync(path.resolve(__dirname, '../src/Guide.js'), 'utf8');

function fakeSheet(name) {
  const sheet = { name, cells: [], cleared: 0 };
  const chain = () => new Proxy({}, { get: (_, k) => (k === 'setValues'
    ? (v) => { sheet.lastWrite(v); return chain(); }
    : () => chain()) });
  let pending = null;
  sheet.lastWrite = (v) => { v.forEach((row, i) => { sheet.cells[pending.row - 1 + i] = sheet.cells[pending.row - 1 + i] || []; row.forEach((c, j) => { sheet.cells[pending.row - 1 + i][pending.col - 1 + j] = c; }); }); };
  sheet.getLastRow = () => sheet.cells.length;
  sheet.getRange = (row, col, n, m) => {
    pending = { row, col };
    const r = chain();
    return new Proxy({}, { get: (_, k) => (k === 'getValues'
      ? () => sheet.cells.slice(row - 1, row - 1 + n).map((x) => (x || []).slice(col - 1, col - 1 + m))
      : r[k]) });
  };
  sheet.getRangeList = () => chain();
  sheet.clear = () => { sheet.cells = []; sheet.cleared += 1; };
  sheet.setColumnWidth = () => {};
  sheet.setFrozenRows = () => {};
  return sheet;
}

function load({ admin = '', db = '', ui = true, deployment = true } = {}) {
  const sheets = {};
  const menus = [];
  let active = null;
  const ss = {
    getSheetByName: (n) => sheets[n] || null,
    insertSheet: (n, index) => { sheets[n] = fakeSheet(n); sheets[n].index = index; return sheets[n]; },
    setActiveSheet: (s) => { active = s; }
  };
  const menu = { addItem: (label, fn) => { menus.push([label, fn]); return menu; }, addToUi: () => {} };
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    SpreadsheetApp: {
      getUi: () => { if (!ui) throw new Error('Cannot call SpreadsheetApp.getUi() from this context.'); return { createMenu: () => menu }; },
      getActiveSpreadsheet: () => ss
    },
    getWebAppUrl: () => 'https://script.google.com/a/macros/example.com/s/XXX/exec'
  };
  vm.createContext(context);
  if (deployment) {
    vm.runInContext(`const DEPLOY_ADMIN_EMAIL = ${JSON.stringify(admin)};
      const DEPLOY_DATABASE_SPREADSHEET_ID = ${JSON.stringify(db)};`, context);
  }
  vm.runInContext(SRC, context, { filename: 'Guide.js' });
  return { context, sheets, menus, ss, getActive: () => active };
}

const read = (name) => vm.runInContext(name, load().context);

test('コピーした配置を開く: 最初のタブに「はじめに」を書き、前に出し、メニューを足す', () => {
  const { context, sheets, menus, getActive } = load();
  context.onOpen();
  const guide = sheets['はじめに'];
  assert.ok(guide, '「はじめに」ができる');
  assert.equal(guide.index, 0, '最初のタブ');
  assert.equal(guide.cells[0][0], 'みんなの回答ボード');
  assert.equal(guide.cells[2][1], read('GUIDE_STATUS_PENDING'));
  assert.equal(getActive(), guide);
  assert.deepEqual(menus, [['はじめに（使い方）を開く', 'openBoardGuide']]);
});

test('2 回目に開く: 今の版ならシートに触らず、前にも出さない', () => {
  const { context, sheets, getActive, ss } = load();
  context.onOpen();
  ss.setActiveSheet(null);
  const before = JSON.stringify(sheets['はじめに'].cells);
  context.onOpen();
  assert.equal(sheets['はじめに'].cleared, 0);
  assert.equal(JSON.stringify(sheets['はじめに'].cells), before);
  assert.equal(getActive(), null);
});

test('焼き込みのある配置・Deployment.js が無いとき: 何もしない (先生の DB にタブを足さない)', () => {
  for (const opts of [{ admin: 't@example.com', db: 'ss' }, { deployment: false }]) {
    const { context, sheets, menus } = load(opts);
    context.onOpen();
    context.openBoardGuide();
    assert.deepEqual(Object.keys(sheets), []);
    assert.deepEqual(menus, []);
  }
});

test('ウェブアプリから呼ばれた (児童の画面の google.script.run): 何も書かない', () => {
  const { context, sheets, menus } = load({ ui: false });
  context.onOpen();
  context.openBoardGuide();
  assert.deepEqual(Object.keys(sheets), []);
  assert.deepEqual(menus, []);
});

test('メニュー: 消してあっても書き直して前に出す', () => {
  const { context, sheets, getActive } = load();
  context.openBoardGuide();
  assert.ok(sheets['はじめに']);
  assert.equal(getActive(), sheets['はじめに']);
});

test('準備が済んだら、状態と管理パネルの URL を書く', () => {
  const { context, sheets, ss } = load();
  context.markGuideReady_(ss);
  const cells = sheets['はじめに'].cells;
  assert.equal(cells[2][1], read('GUIDE_STATUS_READY'));
  assert.equal(cells[3][1], 'https://script.google.com/a/macros/example.com/s/XXX/exec');
});

test('版が古い: 書き直すが、準備が済んだ印と URL は残す', () => {
  const { context, sheets, ss } = load();
  context.markGuideReady_(ss);
  const cells = sheets['はじめに'].cells;
  cells[cells.length - 1][1] = '0';   // 古い版の案内
  cells[10][1] = '古い手順の文';
  context.onOpen();
  const after = sheets['はじめに'].cells;
  assert.equal(sheets['はじめに'].cleared, 1);
  assert.notEqual(after[10][1], '古い手順の文');
  assert.equal(after[2][1], read('GUIDE_STATUS_READY'));
  assert.equal(after[3][1], 'https://script.google.com/a/macros/example.com/s/XXX/exec');
});

test('案内の行: 状態と URL の行の位置が定数と合っている', () => {
  const { context } = load();
  const lines = vm.runInContext('GUIDE_LINES', context);
  assert.equal(lines[vm.runInContext('GUIDE_STATUS_ROW', context) - 1][1], '今の状態');
  assert.equal(lines[vm.runInContext('GUIDE_URL_ROW', context) - 1][1], '管理パネルの URL');
  assert.equal(lines[lines.length - 1][1], vm.runInContext('GUIDE_VERSION_LABEL', context));
});
