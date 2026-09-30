/**
 * 「● → ★」の「ことばが変わった」の印 (管理パネル)。
 * Why: 位置が同じでも根拠が変わった児童を拾う。評価はしない (印だけ)。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function loadFn() {
  const src = fs.readFileSync(path.resolve(__dirname, '../src/AdminPanel.js.html'), 'utf8');
  const m = src.match(/function liveWordsChanged\(first, last\) \{[\s\S]*?\n    \}/);
  if (!m) throw new Error('liveWordsChanged not found');
  const ctx = { String };
  vm.createContext(ctx);
  return vm.runInContext('(' + m[0] + ')', ctx);
}

test('liveWordsChanged: 理由の文面が変わっていれば true', () => {
  const f = loadFn();
  assert.equal(f({ reason: '正直に言う' }, { reason: '友だちのために言う' }), true);
});

test('liveWordsChanged: 同じ文面、空白や改行だけの差は false', () => {
  const f = loadFn();
  assert.equal(f({ reason: '正直に言う' }, { reason: '正直に言う' }), false);
  assert.equal(f({ reason: '正直に 言う' }, { reason: '正直に\n言う' }), false);
});

test('liveWordsChanged: どちらかが無ければ false (未提出は印を出さない)', () => {
  const f = loadFn();
  assert.equal(f(null, { reason: 'x' }), false);
  assert.equal(f({ reason: 'x' }, null), false);
});
