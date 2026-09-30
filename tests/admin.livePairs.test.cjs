/**
 * 話すと面白い組 (管理パネルの回答一覧、教師の手元だけ)。
 * Why: 異なる立場と話す機会を、名簿と分布を見比べずに作れるようにする。自動で組ませない。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function loadFn() {
  const src = fs.readFileSync(path.resolve(__dirname, '../src/AdminPanel.js.html'), 'utf8');
  const m = src.match(/function livePickPairs\(rows, axis\) \{[\s\S]*?\n    \}\n/);
  if (!m) throw new Error('livePickPairs not found');
  const ctx = { String, Number, Set, Array };
  vm.createContext(ctx);
  return vm.runInContext('(' + m[0] + ')', ctx);
}
const matrix = { boardMode: 'matrix', min: 1, max: 5, x: { min: '言う', max: '言わない' }, y: { min: '迷いあり', max: '迷いなし' } };
const r = (email, x, y, reason) => ({ email, name: email, numericX: x, numericY: y, reason });

test('1 人以下なら何も出さない', () => {
  const f = loadFn();
  assert.equal(f([], matrix).length, 0);
  assert.equal(f([r("a", 1, 1, "x")], matrix).length, 0);
});

test('位置が一番遠い 2 人、同じ位置で理由が違う 2 人、縦軸の両端、を重複なく選ぶ', () => {
  const f = loadFn();
  const rows = [
    r('far1', 1, 1, '正直に言うから'),
    r('far2', 5, 5, 'にがすから'),
    r('same1', 3, 3, '迷うから'),
    r('same2', 3, 3, 'どちらも友だちのためだから'),
    r('lo', 2, 1, '迷いあり'),
    r('hi', 4, 5, '迷いなし')
  ];
  const pairs = f(rows, matrix);
  assert.equal(pairs.length, 3);
  assert.equal(pairs[0].title, '位置が一番遠い');
  assert.deepEqual([pairs[0].a.email, pairs[0].b.email].sort(), ['far1', 'far2']);
  assert.equal(pairs[1].title, '同じ位置で理由がちがう');
  assert.deepEqual([pairs[1].a.email, pairs[1].b.email].sort(), ['same1', 'same2']);
  assert.equal(pairs[2].title, '「迷いあり」と「迷いなし」');
  assert.deepEqual([pairs[2].a.email, pairs[2].b.email].sort(), ['hi', 'lo']);
  const all = pairs.flatMap((p) => [p.a.email, p.b.email]);
  assert.equal(new Set(all).size, all.length, '同じ児童が複数の組に出ない');
});

test('全員が同じ位置・同じ理由なら組は出ない (距離 0 は「遠い」にしない)', () => {
  const f = loadFn();
  const rows = [r('a', 3, 3, '同じ'), r('b', 3, 3, '同じ'), r('c', 3, 3, '同じ')];
  assert.equal(f(rows, matrix).length, 0);
});

test('数直線 (1 次元) は横だけで距離を測り、縦軸の組は出さない', () => {
  const f = loadFn();
  const line = { boardMode: 'numberline', min: 1, max: 5, x: { min: '反対', max: '賛成' }, y: null };
  const rows = [r('a', 1, null, 'x だから'), r('b', 5, null, 'y だから'), r('c', 3, null, 'z だから')];
  const pairs = f(rows, line);
  assert.equal(pairs.length, 1);
  assert.deepEqual([pairs[0].a.email, pairs[0].b.email].sort(), ['a', 'b']);
});
