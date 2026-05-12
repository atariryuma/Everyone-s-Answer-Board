const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function loadCtx() {
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    normalizeHeader: (h) => String(h || '').toLowerCase().trim()
  };
  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/ColumnMappingService.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'ColumnMappingService.js' });
  return context;
}

// =====================================================================
// detectNumericScaleColumns
// Why: Forms「線形尺度」を整数値で識別する自動検出ロジック。M1/M2 のキー機能。
// =====================================================================

test('detectNumericScaleColumns: detects a 1-5 linear scale column', () => {
  const ctx = loadCtx();
  const headers = ['タイムスタンプ', 'メール', '回答', 'ポスター対応度'];
  const sample = [
    ['2026-01-01', 'a@x', 'はい', 1],
    ['2026-01-01', 'b@x', 'はい', 5],
    ['2026-01-01', 'c@x', 'いいえ', 3],
    ['2026-01-01', 'd@x', 'まあ', 4],
    ['2026-01-01', 'e@x', 'いいえ', 2]
  ];
  const result = ctx.detectNumericScaleColumns(headers, sample);
  assert.equal(result.length, 1);
  assert.equal(result[0].index, 3);
  assert.equal(result[0].header, 'ポスター対応度');
  assert.equal(result[0].min, 1);
  assert.equal(result[0].max, 5);
  assert.ok(result[0].confidence >= 80, `confidence ${result[0].confidence} should be ≥80`);
});

test('detectNumericScaleColumns: detects two scale columns for matrix mode (sorted by confidence)', () => {
  const ctx = loadCtx();
  const headers = ['ts', '効率', '倫理'];
  const sample = [
    ['2026', 1, 5],
    ['2026', 2, 4],
    ['2026', 3, 3],
    ['2026', 4, 2],
    ['2026', 5, 1],
    ['2026', 1, 1]
  ];
  const result = ctx.detectNumericScaleColumns(headers, sample);
  assert.equal(result.length, 2);
  // どちらも 1-5 で同等。VM realm 越境のため deepEqual ではなく値比較。
  const indices = result.map(r => r.index).sort((a, b) => a - b);
  assert.equal(indices[0], 1);
  assert.equal(indices[1], 2);
});

test('detectNumericScaleColumns: skips system columns (timestamp, reactions, highlight)', () => {
  const ctx = loadCtx();
  const headers = ['タイムスタンプ', 'UNDERSTAND', 'LIKE', 'CURIOUS', 'HIGHLIGHT', 'X'];
  const sample = [
    [new Date(), 0, 1, 0, 0, 3],
    [new Date(), 1, 0, 1, 1, 4],
    [new Date(), 0, 0, 0, 0, 5]
  ];
  const result = ctx.detectNumericScaleColumns(headers, sample);
  // X 列のみ候補
  assert.equal(result.length, 1);
  assert.equal(result[0].index, 5);
});

test('detectNumericScaleColumns: rejects non-integer columns', () => {
  const ctx = loadCtx();
  const headers = ['ts', '感想', '評価'];
  const sample = [
    ['2026', 'たのしい', 4],
    ['2026', 'ふつう', 3],
    ['2026', '微妙', 2.5]  // 小数あり → 全体として非整数列
  ];
  const result = ctx.detectNumericScaleColumns(headers, sample);
  // '評価' 列は 2.5 を含むので不採用
  assert.equal(result.length, 0);
});

test('detectNumericScaleColumns: rejects columns where all values are the same', () => {
  const ctx = loadCtx();
  const headers = ['ts', 'X'];
  const sample = [['2026', 3], ['2026', 3], ['2026', 3]];
  const result = ctx.detectNumericScaleColumns(headers, sample);
  assert.equal(result.length, 0);
});

test('detectNumericScaleColumns: rejects columns with range > 9 (likely free numeric, not scale)', () => {
  const ctx = loadCtx();
  const headers = ['ts', 'age'];
  const sample = [['2026', 7], ['2026', 50], ['2026', 25]];
  const result = ctx.detectNumericScaleColumns(headers, sample);
  assert.equal(result.length, 0);
});

test('detectNumericScaleColumns: parses numeric strings (Forms may export as string)', () => {
  const ctx = loadCtx();
  const headers = ['ts', 'X'];
  const sample = [['2026', '1'], ['2026', '5'], ['2026', '3'], ['2026', '4']];
  const result = ctx.detectNumericScaleColumns(headers, sample);
  assert.equal(result.length, 1);
  assert.equal(result[0].min, 1);
  assert.equal(result[0].max, 5);
});

test('detectNumericScaleColumns: requires at least 2 non-empty values', () => {
  const ctx = loadCtx();
  const headers = ['ts', 'X'];
  const sample = [['2026', 1], ['2026', ''], ['2026', null]];
  const result = ctx.detectNumericScaleColumns(headers, sample);
  // 有効データ 1 件のみ → スキップ
  assert.equal(result.length, 0);
});

test('detectNumericScaleColumns: empty inputs return empty array (no throw)', () => {
  const ctx = loadCtx();
  assert.equal(ctx.detectNumericScaleColumns([], []).length, 0);
  assert.equal(ctx.detectNumericScaleColumns(['X'], []).length, 0);
  assert.equal(ctx.detectNumericScaleColumns(null, null).length, 0);
});

test('detectNumericScaleColumns: higher confidence for 1-5 with many samples', () => {
  const ctx = loadCtx();
  const headers = ['ts', 'manySamples', 'fewSamples'];
  const sample = [];
  for (let i = 0; i < 20; i++) {
    sample.push(['2026', (i % 5) + 1, null]);
  }
  sample.push(['2026', null, 1]);
  sample.push(['2026', null, 5]);
  const result = ctx.detectNumericScaleColumns(headers, sample);
  const many = result.find(r => r.index === 1);
  const few = result.find(r => r.index === 2);
  assert.ok(many, 'manySamples should be detected');
  assert.ok(few, 'fewSamples should be detected');
  assert.ok(many.confidence > few.confidence, `${many.confidence} > ${few.confidence}`);
});

// =====================================================================
// performIntegratedColumnDiagnostics: 統合返却に candidates が含まれること
// =====================================================================

test('performIntegratedColumnDiagnostics: returns numericScaleCandidates field', () => {
  const ctx = loadCtx();
  const headers = ['タイムスタンプ', 'メール', '回答', '評価'];
  const sample = [
    ['2026', 'a@x', 'はい', 1],
    ['2026', 'b@x', 'いいえ', 5]
  ];
  const result = ctx.performIntegratedColumnDiagnostics(headers, { sampleData: sample });
  assert.equal(result.success, true);
  assert.ok(Array.isArray(result.numericScaleCandidates));
  assert.equal(result.numericScaleCandidates.length, 1);
  assert.equal(result.numericScaleCandidates[0].header, '評価');
});
