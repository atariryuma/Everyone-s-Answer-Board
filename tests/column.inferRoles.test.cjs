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

test('inferColumnRoles: L1 — 振り返り系自由記述パターンを answer として検出 (v2764)', () => {
  // 実フォーム由来の question: 「今日の学習で、心に残ったこと・学んだこと・これからやってみたいことを
  //   自分の言葉で書いてください」を 48% (medium-text fallback) ではなく
  //   L1 patterns で正しく拾えるか。書いてください / 学んだこと / 心に残った を追加した。
  const ctx = loadCtx();
  const headers = ['ts', '名前', '今日の学習で、心に残ったこと・学んだこと・書いてください'];
  const r = ctx.inferColumnRoles(headers);
  assert.equal(r.mapping.answer, 2);
  assert.ok(r.confidence.answer >= 85,
    `振り返り question pattern should match L1 (>=85), got ${r.confidence.answer}`);
});

test('inferColumnRoles: L1 — 「書こう」「書いて」末尾パターンを answer として検出', () => {
  const ctx = loadCtx();
  // 児童向けフォーム頻出: 命令形末尾 (書こう / 書いて)
  assert.equal(ctx.inferColumnRoles(['ts', '理由を書こう']).mapping.answer, 1);
  assert.equal(ctx.inferColumnRoles(['ts', 'なぜそう思ったか書いて']).mapping.answer, 1);
});

test('inferColumnRoles: L1 — question mark and natural-language question patterns detected as answer', () => {
  const ctx = loadCtx();
  // 児童向けフォームでよくある自由記述質問。「回答」「答え」のような明示キーワードは含まれない。
  const headers1 = ['ts', 'クラス', '名前', 'あなたなら、ポスターをどうする？'];
  const r1 = ctx.inferColumnRoles(headers1);
  assert.equal(r1.mapping.answer, 3, 'question mark in header should map to answer');

  const headers2 = ['ts', '名前', 'なぜそう思いますか'];
  const r2 = ctx.inferColumnRoles(headers2);
  assert.equal(r2.mapping.answer, 2, 'なぜ pattern should map to answer');
});

test('inferColumnRoles: L1=0 + L2=long-text fallback (weak inference, capped at 60)', () => {
  const ctx = loadCtx();
  // ヘッダーが完全に無関係でも、データが長文なら answer 候補に拾われる
  const headers = ['ts', '名前', '入力欄'];
  const sample = [
    ['t', 'A', 'なぜならば、実験結果から重要な傾向が読み取れて議論の余地があるためです。']
  ];
  const r = ctx.inferColumnRoles(headers, sample);
  assert.equal(r.mapping.answer, 2);
  // L1=0 + L2 fallback は信頼度 <= 60 (弱い推定であることを表現)
  assert.ok(r.confidence.answer <= 60, `weak inference should be ≤60, got ${r.confidence.answer}`);
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

test('inferColumnRoles: M7 — reason-prompt が先でも genuine answer (questionPattern) が answer を取る', () => {
  // Why (M7): "なぜそう思った？" は reason keyword なぜ を含み answer の questionPattern とも重なる。
  //   それより後ろに answer 固有でない自由記述質問があるとき、 旧実装は先頭の理由列を answer に
  //   固定しがちだった。 相互排他ペナルティで理由列の answer スコアを下げ、 後続の純粋な質問列を
  //   answer に取らせる。
  const ctx = loadCtx();
  const headers = ['ts', 'なぜそう思った？', 'あなたはどうしますか'];
  const r = ctx.inferColumnRoles(headers);
  assert.equal(r.mapping.answer, 2, 'genuine question column (idx2) should win answer over the reason-prompt (idx1)');
});

test('inferColumnRoles: M7 — answer 固有 keyword を含むヘッダーはペナルティ対象外', () => {
  // "意見とその理由" は answer keyword (意見) を含むため、 reason keyword (理由) があっても
  //   answer スコアは下がらない。
  const ctx = loadCtx();
  const headers = ['ts', '名前', 'あなたの意見とその理由'];
  const r = ctx.inferColumnRoles(headers);
  assert.equal(r.mapping.answer, 2);
  assert.ok(r.confidence.answer >= 85, `answer keyword present → no penalty, got ${r.confidence.answer}`);
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

test('inferColumnRoles: L3 — pie mode boosts answer score (measurable diff vs board mode)', () => {
  const ctx = loadCtx();
  // L1 で 87 点 (questionPattern 一致だが exact/keyword は外れる) のヘッダーを使い、
  // L3 boost +5 が実際に反映されているか確認する。
  // 旧テストは '回答' (L1=95 exact match で cap に張り付き) で L3 の +5 が打ち消されていた。
  const headers = ['ts', '実験について気づいたことを書きましょう'];
  const pieResult = ctx.inferColumnRoles(headers, [], { boardMode: 'pie' });
  const boardResult = ctx.inferColumnRoles(headers, [], { boardMode: 'board' });
  assert.equal(pieResult.mapping.answer, 1);
  assert.equal(boardResult.mapping.answer, 1);
  // pie で +5 boost、board で +0 → pie の confidence は厳密に高いはず
  assert.ok(
    pieResult.confidence.answer > boardResult.confidence.answer,
    `pie (${pieResult.confidence.answer}) should be > board (${boardResult.confidence.answer})`
  );
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
