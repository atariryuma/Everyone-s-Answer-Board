const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function loadColumnMappingContext(overrides = {}) {
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    normalizeHeader: (h) => String(h || '').toLowerCase().trim(),
    ...overrides
  };

  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/ColumnMappingService.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'ColumnMappingService.js' });
  return context;
}

// =====================================================================
// filterSystemColumns — システム列の除外
// =====================================================================

test('filterSystemColumns: removes timestamp/reaction/highlight columns', () => {
  const ctx = loadColumnMappingContext();
  const { cleanHeaders, indexMap } = ctx.filterSystemColumns([
    'タイムスタンプ', 'メールアドレス', '名前', '回答', 'UNDERSTAND', 'LIKE', 'CURIOUS', 'HIGHLIGHT'
  ]);
  assert.deepEqual([...cleanHeaders], ['メールアドレス', '名前', '回答']);
  assert.deepEqual([...indexMap], [1, 2, 3]);
});

test('filterSystemColumns: removes English timestamp variants', () => {
  const ctx = loadColumnMappingContext();
  const { cleanHeaders } = ctx.filterSystemColumns(['timestamp', 'Name', 'Answer']);
  assert.deepEqual([...cleanHeaders], ['Name', 'Answer']);
});

test('filterSystemColumns: removes Japanese reaction label variants', () => {
  const ctx = loadColumnMappingContext();
  const { cleanHeaders } = ctx.filterSystemColumns(['質問', '理解', 'いいね', '気になる', 'ハイライト']);
  assert.deepEqual([...cleanHeaders], ['質問']);
});

test('filterSystemColumns: removes underscore-prefixed internal columns', () => {
  const ctx = loadColumnMappingContext();
  const { cleanHeaders } = ctx.filterSystemColumns(['_hidden', 'Name', '_debug']);
  assert.deepEqual([...cleanHeaders], ['Name']);
});

test('filterSystemColumns: skips empty/non-string headers', () => {
  const ctx = loadColumnMappingContext();
  const { cleanHeaders, indexMap } = ctx.filterSystemColumns(['Name', '', null, undefined, 42, '  ', 'Answer']);
  assert.deepEqual([...cleanHeaders], ['Name', 'Answer']);
  assert.deepEqual([...indexMap], [0, 6]);
});

test('filterSystemColumns: preserves original order and indices', () => {
  const ctx = loadColumnMappingContext();
  const { cleanHeaders, indexMap } = ctx.filterSystemColumns([
    'タイムスタンプ', 'A', 'B', 'LIKE', 'C'
  ]);
  assert.deepEqual([...cleanHeaders], ['A', 'B', 'C']);
  assert.deepEqual([...indexMap], [1, 2, 4]);
});

// =====================================================================
// resolveColumnIndex — 行処理 hot path (DataService から per-field 呼び出し)
// =====================================================================

test('resolveColumnIndex: existing mapping takes precedence over detection', () => {
  const ctx = loadColumnMappingContext();
  const headers = ['タイムスタンプ', '名前', '回答', '理由'];
  const result = ctx.resolveColumnIndex(headers, 'answer', { answer: 1 });
  assert.equal(result.index, 1);
  assert.equal(result.confidence, 100);
  assert.equal(result.method, 'existing_mapping');
});

test('resolveColumnIndex: invalid mapping index falls through to detection', () => {
  const ctx = loadColumnMappingContext();
  const headers = ['タイムスタンプ', '名前', '回答', '理由'];
  const result = ctx.resolveColumnIndex(headers, 'answer', { answer: 99 });
  assert.equal(result.index, 2);
  assert.notEqual(result.method, 'existing_mapping');
});

test('resolveColumnIndex: negative mapping index falls through to detection', () => {
  const ctx = loadColumnMappingContext();
  const headers = ['タイムスタンプ', '回答', '理由'];
  const result = ctx.resolveColumnIndex(headers, 'answer', { answer: -1 });
  assert.equal(result.index, 1);
});

test('resolveColumnIndex: detects 回答 column with high confidence', () => {
  const ctx = loadColumnMappingContext();
  const headers = ['タイムスタンプ', 'メールアドレス', '名前', '回答', '理由'];
  const result = ctx.resolveColumnIndex(headers, 'answer');
  assert.equal(result.index, 3);
  assert.ok(result.confidence > 80);
});

test('resolveColumnIndex: returns -1 when no match found', () => {
  const ctx = loadColumnMappingContext();
  const result = ctx.resolveColumnIndex(['foo', 'bar', 'baz'], 'answer');
  assert.equal(result.index, -1);
});

test('resolveColumnIndex: returns no_valid_headers when all filtered out', () => {
  const ctx = loadColumnMappingContext();
  const headers = ['タイムスタンプ', 'UNDERSTAND', 'LIKE', 'CURIOUS', 'HIGHLIGHT'];
  const result = ctx.resolveColumnIndex(headers, 'answer');
  assert.equal(result.index, -1);
  assert.equal(result.method, 'no_valid_headers');
});

test('resolveColumnIndex: unknown field type returns -1', () => {
  const ctx = loadColumnMappingContext();
  const result = ctx.resolveColumnIndex(['タイムスタンプ', '回答'], 'nonexistent_field');
  assert.equal(result.index, -1);
});
