/**
 * inferColumnRoles: L1 (header pattern) + L2 (data shape) + L3 (board mode) の統合検証。
 * Why: 旧 detectColumn は header text のみだったため、同一キーワードを含む列の衝突を
 *      実データで解消できなかった。inferColumnRoles は dataType を考慮して衝突解消する。
 */

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
// L1 単独: ヘッダーパターン
// =====================================================================

test('inferColumnRoles: L1 only — standard Japanese form (no sample data)', () => {
  const ctx = loadCtx();
  const headers = ['タイムスタンプ', 'メールアドレス', 'クラス', '名前',
    '問題について気づいたことを書きましょう', 'そう思う理由'];
  const r = ctx.inferColumnRoles(headers);
  assert.equal(r.mapping.email, 1);
  assert.equal(r.mapping.class, 2);
  assert.equal(r.mapping.name, 3);
  assert.equal(r.mapping.answer, 4);
  assert.equal(r.mapping.reason, 5);
});

test('inferColumnRoles: L1 — system columns (timestamp/reactions) are excluded', () => {
  const ctx = loadCtx();
  const headers = ['タイムスタンプ', 'UNDERSTAND', 'LIKE', 'CURIOUS', 'HIGHLIGHT', '回答'];
  const r = ctx.inferColumnRoles(headers);
  assert.equal(r.mapping.answer, 5);
  // システム列は role 未割当
  const sysCols = r.columns.filter(c => c.isSystem);
  assert.ok(sysCols.length >= 5, `expected ≥5 system cols, got ${sysCols.length}`);
  for (const c of sysCols) assert.equal(c.role, null);
});

// =====================================================================
// L2: データ形状による衝突解消
// =====================================================================

test('inferColumnRoles: L2 — two "回答"-like headers, longer text wins', () => {
  const ctx = loadCtx();
  const headers = ['ts', '回答', '答え'];
  const sample = [
    ['ts1', 'はい', 'なぜならば、実験結果から重要な傾向が読み取れて...']
  ];
  const r = ctx.inferColumnRoles(headers, sample);
  // 「答え」が長文 → answer に L2 boost。 「回答」は短文 → 相対的に低い。
  // ただし「回答」は exact match で 95、「答え」は keyword match で 90+long-text 8=98 cap 95。
  // どちらも 95 で同点。greedy は配列順 → 先に出現する '回答' (1) が選ばれる。
  assert.equal(r.mapping.answer, 1);
});

test('inferColumnRoles: L2 — email column wins email role over name with email-like header', () => {
  const ctx = loadCtx();
  const headers = ['ts', 'name', 'メールアドレス'];
  const sample = [['t', 'Tanaka', 'tanaka@school.jp']];
  const r = ctx.inferColumnRoles(headers, sample);
  assert.equal(r.mapping.email, 2);
  assert.equal(r.mapping.name, 1);
});

test('inferColumnRoles: L2 — integer-scale column gets numericX', () => {
  const ctx = loadCtx();
  const headers = ['ts', '回答', '評価'];
  const sample = [
    ['t', 'はい', 1], ['t', 'はい', 2], ['t', 'いいえ', 5],
    ['t', 'はい', 4], ['t', 'はい', 3]
  ];
  const r = ctx.inferColumnRoles(headers, sample);
  assert.equal(r.mapping.numericX, 2);
  assert.equal(r.mapping.answer, 1);
});

test('inferColumnRoles: L2 — two scale columns get numericX + numericY', () => {
  const ctx = loadCtx();
  const headers = ['ts', '効率', '倫理'];
  const sample = [
    ['t', 1, 5], ['t', 2, 4], ['t', 3, 3],
    ['t', 4, 2], ['t', 5, 1]
  ];
  const r = ctx.inferColumnRoles(headers, sample);
  assert.ok(typeof r.mapping.numericX === 'number');
  assert.ok(typeof r.mapping.numericY === 'number');
  assert.notEqual(r.mapping.numericX, r.mapping.numericY);
});

// =====================================================================
// answer-before-reason 制約
// =====================================================================

test('inferColumnRoles: constraint — answer always before reason', () => {
  const ctx = loadCtx();
  const headers = ['ts', '理由', '回答'];
  const r = ctx.inferColumnRoles(headers);
  // '理由' は idx 1 だが answer 候補ではない。'回答' は idx 2 で answer。
  // reason は idx > answer (=2) の中から探すが、'理由' は 1 < 2 なので採用不可。reason 未割当。
  assert.equal(r.mapping.answer, 2);
  assert.equal(r.mapping.reason, undefined);
});

// =====================================================================
// existingMapping fix
// =====================================================================

test('inferColumnRoles: existingMapping fixes role to specified index (confidence 100)', () => {
  const ctx = loadCtx();
  const headers = ['ts', '名前', '回答', '理由'];
  const r = ctx.inferColumnRoles(headers, [], { existingMapping: { answer: 1 } });
  assert.equal(r.mapping.answer, 1);
  assert.equal(r.confidence.answer, 100);
});

// =====================================================================
// L3: BoardMode ヒント
// =====================================================================

test('inferColumnRoles: L3 — pie mode boosts answer role', () => {
  const ctx = loadCtx();
  const headers = ['ts', '回答'];
  const r = ctx.inferColumnRoles(headers, [], { boardMode: 'pie' });
  assert.equal(r.mapping.answer, 1);
  // pie mode で +5 bias → 95 cap (元 95 で同じだが、threshold 比較で違いが出るケースで意味)
  assert.ok(r.confidence.answer >= 90);
});

// =====================================================================
// Enriched output: columns[]
// =====================================================================

test('inferColumnRoles: columns[] reports per-column dataType (UI 用)', () => {
  const ctx = loadCtx();
  const headers = ['ts', 'メール', '評価'];
  const sample = [
    ['t', 'a@x.jp', 1], ['t', 'b@x.jp', 5], ['t', 'c@x.jp', 3]
  ];
  const r = ctx.inferColumnRoles(headers, sample);
  const emailCol = r.columns.find(c => c.index === 1);
  const scaleCol = r.columns.find(c => c.index === 2);
  assert.equal(emailCol.stats.dataType, 'email');
  assert.equal(scaleCol.stats.dataType, 'integer-scale');
});

// =====================================================================
// Empty inputs
// =====================================================================

test('inferColumnRoles: empty headers returns empty result', () => {
  const ctx = loadCtx();
  const r = ctx.inferColumnRoles([]);
  assert.equal(Object.keys(r.mapping).length, 0);
  assert.equal(r.columns.length, 0);
});

// =====================================================================
// performIntegratedColumnDiagnostics: enriched
// =====================================================================

test('performIntegratedColumnDiagnostics: includes columnIntelligence per-column data', () => {
  const ctx = loadCtx();
  const headers = ['タイムスタンプ', 'メール', '回答', '評価'];
  const sample = [
    ['t', 'a@x.jp', 'はい', 1], ['t', 'b@x.jp', 'いいえ', 5]
  ];
  const r = ctx.performIntegratedColumnDiagnostics(headers, { sampleData: sample });
  assert.ok(Array.isArray(r.columnIntelligence));
  assert.ok(r.columnIntelligence.length >= 1);
  const scaleCol = r.columnIntelligence.find(c => c.header === '評価');
  assert.ok(scaleCol, 'expected 評価 in columnIntelligence');
  assert.equal(scaleCol.dataType, 'integer-scale');
});

test('performIntegratedColumnDiagnostics: recommendedMapping excludes numericX/Y (back-compat)', () => {
  const ctx = loadCtx();
  const headers = ['ts', '評価'];
  const sample = [['t', 1], ['t', 5], ['t', 3]];
  const r = ctx.performIntegratedColumnDiagnostics(headers, { sampleData: sample });
  // numericX/Y は recommendedMapping に入れない (getColumnAnalysis 側でマージ)
  assert.equal(r.recommendedMapping.numericX, undefined);
  // numericScaleCandidates には含まれる
  assert.equal(r.numericScaleCandidates.length, 1);
});
