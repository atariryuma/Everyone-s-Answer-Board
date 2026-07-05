// scripts/lib/theme-core.js (theme-*.js CLI 群の共通色/token ユーティリティ) のテスト。
// plain CommonJS モジュールなので vm sandbox 不要、直接 require する。

const test = require('node:test');
const assert = require('node:assert/strict');
const {
  parseColor,
  blendOnBg,
  relativeLuminance,
  extractBlockBody,
  extractTokensFromBody,
} = require('../scripts/lib/theme-core');

// ── parseColor ────────────────────────────────────────────────

test('parseColor: #rrggbb を {r,g,b,a:1} に', () => {
  assert.deepEqual(parseColor('#ff8040'), { r: 255, g: 128, b: 64, a: 1 });
});

test('parseColor: #rgb を展開', () => {
  assert.deepEqual(parseColor('#f84'), { r: 255, g: 136, b: 68, a: 1 });
});

test('parseColor: rgba() は alpha を保持', () => {
  assert.deepEqual(parseColor('rgba(10, 20, 30, 0.5)'), { r: 10, g: 20, b: 30, a: 0.5 });
});

test('parseColor: rgb() は alpha=1', () => {
  assert.deepEqual(parseColor('rgb(1, 2, 3)'), { r: 1, g: 2, b: 3, a: 1 });
});

test('parseColor: 不正 hex は null', () => {
  assert.equal(parseColor('#gggggg'), null);
  assert.equal(parseColor('#12'), null); // 2 桁は 6 桁化できず null
});

test('parseColor: null / 空 / 非色文字列は null', () => {
  assert.equal(parseColor(null), null);
  assert.equal(parseColor(''), null);
  assert.equal(parseColor('transparent'), null);
});

// ── blendOnBg ─────────────────────────────────────────────────

test('blendOnBg: alpha>=1 は fg をそのまま返す', () => {
  const fg = { r: 100, g: 100, b: 100, a: 1 };
  assert.equal(blendOnBg(fg, { r: 0, g: 0, b: 0, a: 1 }), fg);
});

test('blendOnBg: alpha=0.5 を白背景に合成', () => {
  // 黒 50% を白に重ねる → 中間グレー 128
  assert.deepEqual(
    blendOnBg({ r: 0, g: 0, b: 0, a: 0.5 }, { r: 255, g: 255, b: 255, a: 1 }),
    { r: 128, g: 128, b: 128, a: 1 }
  );
});

test('blendOnBg: fg / bg が null なら fg を返す (ガード)', () => {
  assert.equal(blendOnBg(null, { r: 0, g: 0, b: 0, a: 1 }), null);
});

// ── relativeLuminance ─────────────────────────────────────────

test('relativeLuminance: 白 = 1, 黒 = 0', () => {
  assert.equal(relativeLuminance({ r: 255, g: 255, b: 255 }), 1);
  assert.equal(relativeLuminance({ r: 0, g: 0, b: 0 }), 0);
});

test('relativeLuminance で白黒 contrast ratio = 21', () => {
  const l1 = relativeLuminance({ r: 255, g: 255, b: 255 });
  const l2 = relativeLuminance({ r: 0, g: 0, b: 0 });
  const ratio = (Math.max(l1, l2) + 0.05) / (Math.min(l1, l2) + 0.05);
  assert.equal(Math.round(ratio), 21);
});

// ── extractBlockBody / extractTokensFromBody ──────────────────

test('extractBlockBody: :root ブロックの中身をネスト対応で抽出', () => {
  const css = ':root { --a: 1; --b: 2; } body { --c: 3; }';
  assert.equal(extractBlockBody(css, /:root\s*\{/).trim(), '--a: 1; --b: 2;');
});

test('extractBlockBody: マッチ無しは空文字', () => {
  assert.equal(extractBlockBody('body {}', /:root\s*\{/), '');
});

test('extractTokensFromBody: --token: value を map に', () => {
  assert.deepEqual(
    extractTokensFromBody('--theme-bg: #1a1b26; --theme-fg: #e2e8f0;'),
    { '--theme-bg': '#1a1b26', '--theme-fg': '#e2e8f0' }
  );
});

test('extractTokensFromBody: 空 body は空 map', () => {
  assert.deepEqual(extractTokensFromBody(''), {});
});
