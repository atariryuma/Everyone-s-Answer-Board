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
// filterSystemColumns
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
// resolveColumnIndex — existing mapping priority
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

  assert.equal(result.index, 2); // detects '回答'
  assert.notEqual(result.method, 'existing_mapping');
});

test('resolveColumnIndex: negative mapping index falls through to detection', () => {
  const ctx = loadColumnMappingContext();
  const headers = ['タイムスタンプ', '回答', '理由'];
  const result = ctx.resolveColumnIndex(headers, 'answer', { answer: -1 });

  // -1 is invalid, should detect '回答' via content; but the system column
  // filter keeps '回答' at index 1 in original array.
  assert.equal(result.index, 1);
});

// =====================================================================
// resolveColumnIndex — Japanese form detection (realistic scenarios)
// =====================================================================

test('resolveColumnIndex: detects 回答 column in standard Japanese form', () => {
  const ctx = loadColumnMappingContext();
  const headers = ['タイムスタンプ', 'メールアドレス', '名前', '回答', '理由', 'UNDERSTAND', 'LIKE', 'CURIOUS', 'HIGHLIGHT'];

  const result = ctx.resolveColumnIndex(headers, 'answer');
  assert.equal(result.index, 3);
  assert.ok(result.confidence > 80);
});

test('resolveColumnIndex: detects 理由 column after 回答', () => {
  const ctx = loadColumnMappingContext();
  const headers = ['タイムスタンプ', '名前', '回答', '理由'];

  const result = ctx.resolveColumnIndex(headers, 'reason');
  assert.equal(result.index, 3);
  assert.ok(result.confidence > 80);
});

test('resolveColumnIndex: detects free-text question as answer', () => {
  const ctx = loadColumnMappingContext();
  const headers = ['タイムスタンプ', '名前', '実験について気づいたことを書きましょう', 'そう思う理由'];

  const result = ctx.resolveColumnIndex(headers, 'answer');
  assert.equal(result.index, 2);
});

test('resolveColumnIndex: detects 名前 column', () => {
  const ctx = loadColumnMappingContext();
  const headers = ['タイムスタンプ', 'クラス', '名前', '回答'];

  const result = ctx.resolveColumnIndex(headers, 'name');
  assert.equal(result.index, 2);
});

test('resolveColumnIndex: detects クラス column', () => {
  const ctx = loadColumnMappingContext();
  const headers = ['タイムスタンプ', 'クラス', '名前', '回答'];

  const result = ctx.resolveColumnIndex(headers, 'class');
  assert.equal(result.index, 1);
});

test('resolveColumnIndex: detects email column via full-width alias', () => {
  const ctx = loadColumnMappingContext();
  const headers = ['タイムスタンプ', 'メールアドレス', '名前', '回答'];

  const result = ctx.resolveColumnIndex(headers, 'email');
  assert.equal(result.index, 1);
});

test('resolveColumnIndex: detects English headers', () => {
  const ctx = loadColumnMappingContext();
  const headers = ['Timestamp', 'Email', 'Name', 'Answer', 'Reason'];

  assert.equal(ctx.resolveColumnIndex(headers, 'email').index, 1);
  assert.equal(ctx.resolveColumnIndex(headers, 'name').index, 2);
  assert.equal(ctx.resolveColumnIndex(headers, 'answer').index, 3);
  assert.equal(ctx.resolveColumnIndex(headers, 'reason').index, 4);
});

// =====================================================================
// resolveColumnIndex — edge cases
// =====================================================================

test('resolveColumnIndex: returns -1 when no match found', () => {
  const ctx = loadColumnMappingContext();
  const headers = ['foo', 'bar', 'baz'];

  const result = ctx.resolveColumnIndex(headers, 'answer');
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
  const headers = ['タイムスタンプ', '回答'];

  const result = ctx.resolveColumnIndex(headers, 'nonexistent_field');
  assert.equal(result.index, -1);
});

// =====================================================================
// analyzeFieldRelationships
// =====================================================================

test('analyzeFieldRelationships: identifies answer/reason logical pairs', () => {
  const ctx = loadColumnMappingContext();
  const headers = ['名前', '回答', '理由'];
  const rels = ctx.analyzeFieldRelationships(headers);

  assert.ok(rels.answerCandidates.length >= 1);
  assert.ok(rels.reasonCandidates.length >= 1);
  assert.equal(rels.answerReasonPairs.length, 1);
  assert.equal(rels.answerReasonPairs[0].logicalOrder, true);
});

test('analyzeFieldRelationships: no pair when reason appears before answer', () => {
  const ctx = loadColumnMappingContext();
  const headers = ['名前', '理由', '回答'];
  const rels = ctx.analyzeFieldRelationships(headers);

  assert.equal(rels.answerReasonPairs.length, 0);
});

// =====================================================================
// validateFieldOrderLogic
// =====================================================================

test('validateFieldOrderLogic: removes reason when it appears before answer with low confidence', () => {
  const ctx = loadColumnMappingContext();
  const headers = ['名前', '理由', '回答'];
  const result = ctx.validateFieldOrderLogic(
    { answer: 2, reason: 1 },
    { answer: 90, reason: 70 },
    headers
  );

  assert.equal(result.mapping.reason, undefined);
  assert.equal(result.mapping.answer, 2);
  assert.ok(result.validation.corrections.length > 0);
});

test('validateFieldOrderLogic: preserves logical order when answer < reason', () => {
  const ctx = loadColumnMappingContext();
  const headers = ['名前', '回答', '理由'];
  const result = ctx.validateFieldOrderLogic(
    { answer: 1, reason: 2 },
    { answer: 90, reason: 85 },
    headers
  );

  assert.equal(result.mapping.answer, 1);
  assert.equal(result.mapping.reason, 2);
});

test('validateFieldOrderLogic: warns on potential field swap', () => {
  const ctx = loadColumnMappingContext();
  // Simulated mis-detection: position 2 ('理由のコメント') chosen as answer but
  // that header itself scores highly as reason (contains '理由'). In parallel
  // position 1 ('回答') is chosen as reason. answer confidence < reason - 10,
  // so the function should flag a possible swap.
  const headers = ['名前', '回答', '理由のコメント'];
  const result = ctx.validateFieldOrderLogic(
    { answer: 2, reason: 1 },
    { answer: 70, reason: 90 },
    headers
  );

  assert.ok(result.validation.warnings.some((w) => w.warning === 'potential_field_swap'));
});

// =====================================================================
// generateRecommendedMapping
// =====================================================================

test('generateRecommendedMapping: produces full mapping for standard form', () => {
  const ctx = loadColumnMappingContext();
  const headers = [
    'タイムスタンプ', 'メールアドレス', 'クラス', '名前',
    '問題について気づいたことを書きましょう', 'そう思う理由',
    'UNDERSTAND', 'LIKE', 'CURIOUS', 'HIGHLIGHT'
  ];

  const result = ctx.generateRecommendedMapping(headers);

  assert.equal(result.success, true);
  assert.equal(result.recommendedMapping.email, 1);
  assert.equal(result.recommendedMapping.class, 2);
  assert.equal(result.recommendedMapping.name, 3);
  assert.equal(result.recommendedMapping.answer, 4);
  assert.equal(result.recommendedMapping.reason, 5);
});

test('generateRecommendedMapping: avoids duplicate index assignment', () => {
  const ctx = loadColumnMappingContext();
  const headers = ['タイムスタンプ', '名前', '氏名', '回答'];
  const result = ctx.generateRecommendedMapping(headers);

  // Both '名前' and '氏名' match name, but each index can only be used once.
  const assignedIndices = Object.values(result.recommendedMapping);
  const uniqueIndices = [...new Set(assignedIndices)];
  assert.equal(assignedIndices.length, uniqueIndices.length);
});

test('generateRecommendedMapping: respects custom fields option', () => {
  const ctx = loadColumnMappingContext();
  const headers = ['タイムスタンプ', 'メールアドレス', '名前', '回答'];
  const result = ctx.generateRecommendedMapping(headers, { fields: ['email', 'name'] });

  assert.ok('email' in result.recommendedMapping);
  assert.ok('name' in result.recommendedMapping);
  assert.ok(!('answer' in result.recommendedMapping));
});

test('generateRecommendedMapping: overallScore reflects average confidence', () => {
  const ctx = loadColumnMappingContext();
  const headers = ['タイムスタンプ', '名前', '回答', '理由'];
  const result = ctx.generateRecommendedMapping(headers);

  assert.ok(result.analysis.overallScore > 80);
  assert.equal(result.analysis.resolvedFields, Object.keys(result.recommendedMapping).length);
});

test('generateRecommendedMapping: handles empty headers gracefully', () => {
  const ctx = loadColumnMappingContext();
  const result = ctx.generateRecommendedMapping([]);

  assert.equal(result.success, true);
  assert.equal(Object.keys(result.recommendedMapping).length, 0);
  assert.equal(result.analysis.overallScore, 0);
});

// =====================================================================
// performIntegratedColumnDiagnostics
// =====================================================================

test('performIntegratedColumnDiagnostics: wraps generateRecommendedMapping result', () => {
  const ctx = loadColumnMappingContext();
  const headers = ['タイムスタンプ', '名前', '回答', '理由'];
  const result = ctx.performIntegratedColumnDiagnostics(headers);

  assert.equal(result.success, true);
  assert.deepEqual([...result.headers], headers);
  assert.ok(result.recommendedMapping);
  assert.ok(result.confidence);
  assert.ok(result.aiAnalysis);
  assert.ok(result.timestamp);
});

// =====================================================================
// detectColumn (internal logic)
// =====================================================================

test('detectColumn: returns unknown_field for unregistered field types', () => {
  const ctx = loadColumnMappingContext();
  const result = ctx.detectColumn(['名前', '回答'], 'totally_unknown');

  assert.equal(result.index, -1);
  assert.equal(result.method, 'unknown_field');
});

test('detectColumn: confidence capped at 95', () => {
  const ctx = loadColumnMappingContext();
  const result = ctx.detectColumn(['回答'], 'answer');

  assert.ok(result.index === 0);
  assert.ok(result.confidence <= 95);
});
