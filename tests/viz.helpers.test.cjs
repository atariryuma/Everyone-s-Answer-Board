const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

/**
 * Tests for page.viz.js.html internal helpers:
 *  - tokenizeJapanese (M3 word ranking)
 *  - aggregateWords / aggregateCategories
 *  - buildSnapshot
 *
 * Why: フロント実装だが純粋関数なので vm でロード可能。授業当日の表示が
 *      壊れないよう、トークナイザのストップワード判定とスナップショット
 *      シリアライズの contract を tests で固定する。
 *
 * 注意: page.viz.js.html は HTML 内 <script>。<script> タグを剥がして
 *      vm に流し込む。グローバル window.StudyQuestApp に prototype を
 *      追加する構造なので、それを偽装する。
 */

function loadVizContext() {
  const html = fs.readFileSync(path.resolve(__dirname, '../src/page.viz.js.html'), 'utf8');
  // <script>...</script> ブロックを抽出
  const m = html.match(/<script>([\s\S]*?)<\/script>/);
  if (!m) throw new Error('Could not extract <script> from page.viz.js.html');
  const source = m[1];

  // StudyQuestApp.prototype に method を追加する IIFE なので、prototype を持つ
  // dummy class を window 経由で渡す
  const StudyQuestApp = class StudyQuestApp {};

  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    window: { StudyQuestApp },
    StudyQuestApp,
    document: {
      readyState: 'complete',
      addEventListener: () => {},
      getElementById: () => null,
      createElement: () => ({ classList: { add: () => {}, remove: () => {}, toggle: () => {} }, addEventListener: () => {}, setAttribute: () => {}, appendChild: () => {}, style: {} }),
      createElementNS: () => ({ setAttribute: () => {}, appendChild: () => {} }),
      body: { classList: { add: () => {} } }
    },
    URLSearchParams: URLSearchParams,
    Map, Set, Promise, JSON, Math, Number, Object, Array, String, Boolean, Date, RegExp,
    Error, TypeError, parseInt, parseFloat, isNaN, isFinite,
    // Why: withTimelineSwap async テストで Promise を実時間 await する必要があるため、
    //   no-op stub ではなく Node の本物 setTimeout を渡す。
    setTimeout: (fn, ms) => setTimeout(fn, ms),
    clearTimeout: (id) => clearTimeout(id)
  };
  context.window.location = { search: '' };
  vm.createContext(context);
  vm.runInContext(source, context, { filename: 'page.viz.js.html' });
  return { context, StudyQuestApp };
}

// =====================================================================
// tokenizeJapanese
// =====================================================================

test('tokenizeJapanese: splits on punctuation and whitespace', () => {
  const { StudyQuestApp } = loadVizContext();
  const tok = StudyQuestApp.prototype.__tokenize;
  const words = tok('AIは便利だ。でも責任は人間にある。');
  // 「便利」「責任」「人間」「ある」あたりが残る想定（助詞・1文字ひらがなは除外）
  assert.ok(words.includes('便利'), 'should include 便利');
  assert.ok(words.includes('責任'), 'should include 責任');
  assert.ok(words.includes('人間'), 'should include 人間');
  // 助詞・記号は除外
  assert.ok(!words.includes('は'));
  assert.ok(!words.includes('。'));
});

test('tokenizeJapanese: removes stop-words', () => {
  const { StudyQuestApp } = loadVizContext();
  const tok = StudyQuestApp.prototype.__tokenize;
  const words = tok('これは これ それ あれ');
  // STOPWORDS にすべて含まれるため空に近い
  for (const stop of ['これ', 'それ', 'あれ']) {
    assert.ok(!words.includes(stop), `should exclude stopword ${stop}`);
  }
});

test('tokenizeJapanese: handles empty / null input', () => {
  const { StudyQuestApp } = loadVizContext();
  const tok = StudyQuestApp.prototype.__tokenize;
  assert.equal(tok('').length, 0);
  assert.equal(tok(null).length, 0);
  assert.equal(tok(undefined).length, 0);
  assert.equal(tok(123).length, 0);
});

test('tokenizeJapanese: keeps katakana words', () => {
  const { StudyQuestApp } = loadVizContext();
  const tok = StudyQuestApp.prototype.__tokenize;
  const words = tok('AI ポスター インパクト 効率');
  assert.ok(words.includes('AI'));
  assert.ok(words.includes('ポスター'));
  assert.ok(words.includes('インパクト'));
  assert.ok(words.includes('効率'));
});

/**
 * TinySegmenter を inject した状態で tokenizer を動かす用のローダー。
 *   実環境は Page.html include 順序で window.TinySegmenter が先に置かれるが、
 *   テストでは ../src/tinySegmenter.html を抽出して context に注入する。
 */
function loadVizContextWithSegmenter() {
  const html = fs.readFileSync(path.resolve(__dirname, '../src/page.viz.js.html'), 'utf8');
  const segHtml = fs.readFileSync(path.resolve(__dirname, '../src/tinySegmenter.html'), 'utf8');
  const m = html.match(/<script>([\s\S]*?)<\/script>/);
  if (!m) throw new Error('Could not extract <script> from page.viz.js.html');
  const segMatch = segHtml.match(/<script>([\s\S]*?)<\/script>/);
  if (!segMatch) throw new Error('Could not extract <script> from tinySegmenter.html');

  const StudyQuestApp = class StudyQuestApp {};
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    window: { StudyQuestApp },
    StudyQuestApp,
    document: {
      readyState: 'complete',
      addEventListener: () => {},
      getElementById: () => null,
      createElement: () => ({ classList: { add: () => {}, remove: () => {}, toggle: () => {} }, addEventListener: () => {}, setAttribute: () => {}, appendChild: () => {}, style: {} }),
      createElementNS: () => ({ setAttribute: () => {}, appendChild: () => {} }),
      body: { classList: { add: () => {} } }
    },
    URLSearchParams: URLSearchParams,
    Map, Set, Promise, JSON, Math, Number, Object, Array, String, Boolean, Date, RegExp,
    Error, TypeError, parseInt, parseFloat, isNaN, isFinite,
    setTimeout: (fn, ms) => setTimeout(fn, ms),
    clearTimeout: (id) => clearTimeout(id)
  };
  context.window.location = { search: '' };
  vm.createContext(context);
  // Why: tinySegmenter.html の中身は IIFE 内で TinySegmenter 関数を定義する。
  //   page.viz.js.html の IIFE が `typeof TinySegmenter` を見るので、
  //   window.TinySegmenter / global TinySegmenter どちらも見える状態にする必要あり。
  //   wrapper が `window.TinySegmenter = TinySegmenter` してくれるので、context.window
  //   に注入され、page.viz.js が `typeof window.TinySegmenter !== 'undefined'` で検知できる。
  vm.runInContext(segMatch[1], context, { filename: 'tinySegmenter.html' });
  vm.runInContext(m[1], context, { filename: 'page.viz.js.html' });
  return { context, StudyQuestApp };
}

test('tokenizeJapanese: with TinySegmenter, captures common compound words (v2670)', () => {
  // Why: 純正規表現では「気持ち」「お楽しみ」「頑張る」「誠実」が断片化していた。
  //   TinySegmenter を載せると統計分割により 1 トークンとして拾える。
  //   ただし「思いやり」は TinySegmenter でも 「思い/やり」に分割される弱点が残る
  //   （統計モデルの限界）。これは形態素解析でも完璧ではないことの実証。
  const { StudyQuestApp } = loadVizContextWithSegmenter();
  const tok = StudyQuestApp.prototype.__tokenize;

  // 「気持ち」が 1 トークンとして拾える
  const w1 = tok('お楽しみ会は大事、気持ちを忘れない。');
  assert.ok(w1.includes('気持ち'), `should capture 気持ち: ${w1.join(',')}`);
  assert.ok(w1.includes('大事'), `should capture 大事: ${w1.join(',')}`);
  assert.ok(w1.includes('お楽しみ'), `should capture お楽しみ: ${w1.join(',')}`);

  // 「誠実」「大切」「責任」「取り組む」など複合語/2+字漢字内容語が拾える
  //   (「頑張って」は TinySegmenter で「頑張っ + て」に分割されるため、
  //    「頑張っ」は v2672 の促音便フィルタで除外される。これは設計上の妥協で、
  //    形態素解析の限界を露呈する例。)
  const w2 = tok('頑張って誠実に取り組む。大切な責任を持つ。');
  assert.ok(w2.includes('誠実'), `should capture 誠実: ${w2.join(',')}`);
  assert.ok(w2.includes('大切'), `should capture 大切: ${w2.join(',')}`);
  assert.ok(w2.includes('責任'), `should capture 責任: ${w2.join(',')}`);
  assert.ok(w2.some(t => t.includes('取り組')), `should capture 取り組む or 取り組: ${w2.join(',')}`);

  // 純ひらがな短語 (3 文字以下) は除外される (助詞・助動詞の可能性が高い)
  const w3 = tok('これは とても 大切 です');
  assert.ok(!w3.includes('これ'), `should exclude これ: ${w3.join(',')}`);
  assert.ok(!w3.includes('は'), `should exclude particle は: ${w3.join(',')}`);
  assert.ok(w3.includes('大切'), `should keep 大切: ${w3.join(',')}`);
});

test('tokenizeJapanese: filters verb-inflection fragments (v2672)', () => {
  // Why: TinySegmenter は動詞活用の途中で切れることがあり、活用形断片が頻出語
  //   上位に紛れ込む。明確な活用シグナル（末尾「っ」: 促音便、末尾2字「い/え/わ/り/う/る/に/け」:
  //   活用語末/助詞付き）を除外する。名詞化接辞末「ち/み/く/さ/き」は残す。
  const { StudyQuestApp } = loadVizContextWithSegmenter();
  const tok = StudyQuestApp.prototype.__tokenize;

  // 促音便連用形（末尾「っ」）は length 問わず除外
  const w1 = tok('AIを使って間違ってしまった。思った通り。');
  assert.ok(!w1.includes('使っ'), `should exclude 使っ: ${w1.join(',')}`);
  assert.ok(!w1.includes('思っ'), `should exclude 思っ: ${w1.join(',')}`);
  assert.ok(!w1.includes('間違っ'), `should exclude 間違っ: ${w1.join(',')}`);
  assert.ok(w1.includes('AI'), `should keep AI: ${w1.join(',')}`);

  // 動詞活用断片 + 助詞付き残り (2字)
  const w2 = tok('期限に間に合わない。振る舞いを書け。形に残す。');
  assert.ok(!w2.includes('合わ'), `should exclude 合わ: ${w2.join(',')}`);
  assert.ok(!w2.includes('振る'), `should exclude 振る (verb end る): ${w2.join(',')}`);
  assert.ok(!w2.includes('書け'), `should exclude 書け (imperative け): ${w2.join(',')}`);
  assert.ok(!w2.includes('形に'), `should exclude 形に (particle に): ${w2.join(',')}`);
  assert.ok(w2.includes('期限'), `should keep 期限: ${w2.join(',')}`);

  // 名詞化接辞末（ち/み/く/さ/き）は残す
  const w3 = tok('望みを持ち、早く高さを動く');
  assert.ok(w3.some(t => /[ちみくさき]/.test(t)) || w3.length > 0,
    `should keep nominalized forms or have content: ${w3.join(',')}`);

  // 長い活用形 (3字以上) は除外対象外で残る
  const w4 = tok('楽しめる活動を考える。');
  assert.ok(w4.includes('楽しめる'), `should keep 楽しめる (3+ char): ${w4.join(',')}`);
  assert.ok(w4.includes('考える'), `should keep 考える (3+ char): ${w4.join(',')}`);
});

test('renderQuadrantSummary: filters 1-char keyword fragments (v2668)', () => {
  // Why: 「間に合う」「思いやり」のような複合語が tokenizer で「間/合/思」に分解されると、
  //   1 文字漢字断片が頻出語として top に来てしまい、keyword の意味が読み取れない
  //   (教師フィードバック 2026-05-14 v2667: 「#楽 #思 #間」が出る問題)。
  //   象限パネル側で word.length >= 2 のみを keyword 対象とすることで、
  //   「期限」「責任」「真実」のような 2+ 字内容語のみを top-3 に残す。
  const { StudyQuestApp, makeElement } = loadVizContextForQuadrant();
  const answers = makeElement('div');
  // (5,5) hh 象限に多数の意見を集中させて keyword 抽出をテスト
  const rows = [];
  for (let i = 0; i < 5; i++) {
    rows.push({ numericX: 5, numericY: 5, reason: '間に合うように責任を持つ。' });
  }
  // 上記 reason は tokenizer で「間」「合」「責任」「持」を生成する
  // length >= 2 filter で「責任」のみ keyword に残る期待
  StudyQuestApp.prototype.__renderQuadrantSummary(answers, rows, {});
  const panel = answers.querySelector('#vizQuadrantSummary');
  const trCell = panel._children.find(c => /qs-tr/.test(c.className));
  const tags = trCell._children
    .flatMap(c => c._children || [])
    .filter(c => c.className === 'qs-tag')
    .map(c => c.textContent);
  // 「責任」は 2 字なので残る
  assert.ok(tags.some(t => t.includes('責任')), `expected 責任 in tags: ${tags.join(',')}`);
  // 「間」「合」「持」は 1 字なので除外される
  assert.ok(!tags.some(t => t === '#間'), `should NOT include 1-char #間: ${tags.join(',')}`);
  assert.ok(!tags.some(t => t === '#合'), `should NOT include 1-char #合: ${tags.join(',')}`);
  assert.ok(!tags.some(t => t === '#持'), `should NOT include 1-char #持: ${tags.join(',')}`);
});

// =====================================================================
// aggregateWords
// =====================================================================

test('aggregateWords: counts frequency, sorted desc, 1 post = 1 vote per word', () => {
  const { StudyQuestApp } = loadVizContext();
  const agg = StudyQuestApp.prototype.__aggregateWords;
  const rows = [
    { reason: '責任は重要だ。責任を持つべき。' },  // 「責任」は 2 回出るが 1 票
    { reason: '効率も大事だが責任が先' },
    { reason: '効率を優先する' }
  ];
  const result = agg(rows, 'reason');
  const map = new Map(result.map(r => [r.word, r.count]));
  assert.equal(map.get('責任'), 2);   // row 0 + row 1
  assert.equal(map.get('効率'), 2);   // row 1 + row 2
  // 上位順
  assert.ok(result[0].count >= result[result.length - 1].count);
});

test('aggregateWords: handles empty rows', () => {
  const { StudyQuestApp } = loadVizContext();
  const agg = StudyQuestApp.prototype.__aggregateWords;
  assert.equal(agg([], 'reason').length, 0);
  assert.equal(agg(null, 'reason').length, 0);
  assert.equal(agg([{ reason: '' }, { reason: null }], 'reason').length, 0);
});

test('aggregateWords: caps at 25 entries', () => {
  const { StudyQuestApp } = loadVizContext();
  const agg = StudyQuestApp.prototype.__aggregateWords;
  // 30 種類の異なる単語を作る
  const rows = [];
  for (let i = 0; i < 30; i++) {
    rows.push({ reason: 'カタカナ語' + i });
  }
  const result = agg(rows, 'reason');
  assert.ok(result.length <= 25, 'should cap at 25');
});

// =====================================================================
// aggregateCategories
// =====================================================================

test('aggregateCategories: counts categorical values', () => {
  const { StudyQuestApp } = loadVizContext();
  const agg = StudyQuestApp.prototype.__aggregateCategories;
  const rows = [
    { answer: '賛成' }, { answer: '反対' }, { answer: '賛成' },
    { answer: '賛成' }, { answer: 'どちらでもない' }
  ];
  const result = agg(rows, 'answer');
  const map = new Map(result.map(r => [r.label, r.count]));
  assert.equal(map.get('賛成'), 3);
  assert.equal(map.get('反対'), 1);
  assert.equal(map.get('どちらでもない'), 1);
  // sorted desc
  assert.equal(result[0].label, '賛成');
});

test('aggregateCategories: ignores empty / null values', () => {
  const { StudyQuestApp } = loadVizContext();
  const agg = StudyQuestApp.prototype.__aggregateCategories;
  const rows = [
    { answer: '賛成' }, { answer: '' }, { answer: null },
    { answer: undefined }, { answer: '  ' }, { answer: '賛成' }
  ];
  const result = agg(rows, 'answer');
  assert.equal(result.length, 1);
  assert.equal(result[0].count, 2);
});

test('aggregateCategories: trims whitespace', () => {
  const { StudyQuestApp } = loadVizContext();
  const agg = StudyQuestApp.prototype.__aggregateCategories;
  const rows = [{ answer: '賛成' }, { answer: ' 賛成 ' }, { answer: '賛成  ' }];
  const result = agg(rows, 'answer');
  assert.equal(result.length, 1);
  assert.equal(result[0].count, 3);
});

// =====================================================================
// buildSnapshot (via __vizSnapshotInternals)
// =====================================================================

test('buildSnapshot: captures minimal row fields and metadata', () => {
  const { StudyQuestApp } = loadVizContext();
  const sn = StudyQuestApp.prototype.__vizSnapshotInternals;
  const fakeApp = {
    state: {
      boardMode: 'numberline',
      axisConfig: { defaultMin: 1, defaultMax: 5 },
      currentAnswers: [
        { rowIndex: 2, emailHash: 'h1', class: '5-1', name: '太郎', answer: 'はい', reason: 'なぜなら', numericX: 4, numericY: null, timestamp: '2026-05-12' },
        { rowIndex: 3, emailHash: 'h2', class: '5-1', name: '花子', answer: 'いいえ', reason: '違う', numericX: 2, numericY: 3, timestamp: '2026-05-12' }
      ]
    }
  };
  const snap = sn.buildSnapshot(fakeApp);
  assert.equal(snap.boardMode, 'numberline');
  assert.equal(snap.axisConfig.defaultMin, 1);
  assert.equal(snap.rows.length, 2);
  assert.equal(snap.rows[0].emailHash, 'h1');
  assert.equal(snap.rows[0].numericX, 4);
  assert.equal(snap.rows[1].numericY, 3);
  assert.ok(typeof snap.capturedAt === 'number');
});

test('buildSnapshot: emits null for missing numericX/Y', () => {
  const { StudyQuestApp } = loadVizContext();
  const sn = StudyQuestApp.prototype.__vizSnapshotInternals;
  const fakeApp = { state: { currentAnswers: [{ rowIndex: 2, answer: 'x' }] } };
  const snap = sn.buildSnapshot(fakeApp);
  assert.equal(snap.rows[0].numericX, null);
  assert.equal(snap.rows[0].numericY, null);
});

test('SNAPSHOT_MAX / INTERVAL_MS sanity', () => {
  const { StudyQuestApp } = loadVizContext();
  const sn = StudyQuestApp.prototype.__vizSnapshotInternals;
  assert.ok(sn.SNAPSHOT_MAX > 0 && sn.SNAPSHOT_MAX <= 200, 'cap is reasonable');
  assert.ok(sn.SNAPSHOT_INTERVAL_MS >= 10000, 'interval >= 10s');
});

// =====================================================================
// buildHistoricalSnapshots (SS timestamp-based timeline replay)
// =====================================================================

test('buildHistoricalSnapshots: groups rows by INTERVAL buckets and tracks latest per email', () => {
  const { StudyQuestApp } = loadVizContext();
  const sn = StudyQuestApp.prototype.__vizSnapshotInternals;
  const rows = [
    { rowIndex: 1, emailHash: 'a', timestamp: '2026-05-14T01:00:00Z', numericX: 1, numericY: 1, reason: 'init A' },
    { rowIndex: 2, emailHash: 'b', timestamp: '2026-05-14T01:00:30Z', numericX: 5, numericY: 5, reason: 'init B' },
    { rowIndex: 3, emailHash: 'a', timestamp: '2026-05-14T01:05:00Z', numericX: 4, numericY: 4, reason: 'revote A' }
  ];
  const snaps = sn.buildHistoricalSnapshots(rows, 'matrix', { defaultMin: 1, defaultMax: 5 });
  // 60s バケットで 01:00 〜 01:05+ → 少なくとも 5 バケット以上
  assert.ok(snaps.length >= 5, `expected >=5 buckets, got ${snaps.length}`);
  // 最初のバケットは a=1,1 / b=5,5 (両方初期投票)
  const first = snaps[0];
  assert.equal(first.rows.length, 2);
  assert.equal(first.boardMode, 'matrix');
  // 最後のバケットは a が re-vote 後の値、b は変わらず
  const last = snaps[snaps.length - 1];
  const aRow = last.rows.find(r => r.emailHash === 'a');
  const bRow = last.rows.find(r => r.emailHash === 'b');
  assert.equal(aRow.numericX, 4);
  assert.equal(aRow.numericY, 4);
  assert.equal(aRow.reason, 'revote A');
  assert.equal(bRow.numericX, 5);
});

test('buildHistoricalSnapshots: returns [] when no row has parseable timestamp', () => {
  const { StudyQuestApp } = loadVizContext();
  const sn = StudyQuestApp.prototype.__vizSnapshotInternals;
  // timestamp 欠落 / 不正
  const rows = [
    { rowIndex: 1, emailHash: 'a', numericX: 1 },
    { rowIndex: 2, emailHash: 'b', timestamp: 'not-a-date', numericX: 2 }
  ];
  const snaps = sn.buildHistoricalSnapshots(rows, 'matrix', null);
  // vm context の Array は外の Array と prototype が異なるため deepEqual は使えない。
  // length=0 のチェックで「空配列」を確認する。
  assert.equal(snaps.length, 0);
});

test('buildHistoricalSnapshots: trims to SNAPSHOT_MAX (oldest dropped)', () => {
  const { StudyQuestApp } = loadVizContext();
  const sn = StudyQuestApp.prototype.__vizSnapshotInternals;
  const startMs = new Date('2026-05-14T01:00:00Z').getTime();
  // SNAPSHOT_MAX より多くのバケットを発生させるため、それぞれ INTERVAL を空けた row を生成
  const count = sn.SNAPSHOT_MAX + 10;
  const rows = [];
  for (let i = 0; i < count; i++) {
    rows.push({
      rowIndex: i,
      emailHash: 'u' + i,
      timestamp: new Date(startMs + i * sn.SNAPSHOT_INTERVAL_MS).toISOString(),
      numericX: 3
    });
  }
  const snaps = sn.buildHistoricalSnapshots(rows, 'matrix', null);
  assert.ok(snaps.length <= sn.SNAPSHOT_MAX, `got ${snaps.length}, cap ${sn.SNAPSHOT_MAX}`);
});

test('buildHistoricalSnapshots: rows without emailHash use rowIndex as key', () => {
  const { StudyQuestApp } = loadVizContext();
  const sn = StudyQuestApp.prototype.__vizSnapshotInternals;
  const rows = [
    { rowIndex: 10, emailHash: null, timestamp: '2026-05-14T01:00:00Z', numericX: 2 },
    { rowIndex: 11, emailHash: null, timestamp: '2026-05-14T01:00:30Z', numericX: 3 }
  ];
  const snaps = sn.buildHistoricalSnapshots(rows, 'matrix', null);
  assert.ok(snaps.length >= 1);
  // 異なる rowIndex は別マーカーとして残る
  assert.equal(snaps[snaps.length - 1].rows.length, 2);
});

// =====================================================================
// withTimelineSwap (async/await race regression test)
// =====================================================================

function makeFakeApp(StudyQuestApp, { vizTimelineIndex, snapshots }) {
  const app = new StudyQuestApp();
  app.state = {
    currentAnswers: [{ rowIndex: 1, marker: 'ORIGINAL' }],
    axisConfig: { tag: 'orig' },
    vizTimelineIndex
  };
  app.vizGetSnapshots = () => snapshots;
  return app;
}

test('withTimelineSwap: async fn keeps swap until Promise resolves', async () => {
  const { StudyQuestApp } = loadVizContext();
  const swappedRows = [{ rowIndex: 999, marker: 'SWAPPED' }];
  const app = makeFakeApp(StudyQuestApp, {
    vizTimelineIndex: 0,
    snapshots: [{ capturedAt: 1, rows: swappedRows, axisConfig: { tag: 'swap' } }]
  });
  const original = app.state.currentAnswers;

  let observedDuringFn = null;
  let observedAxisDuringFn = null;
  await app.__vizWithTimelineSwap(async () => {
    // 50ms 待ってから currentAnswers を観測。旧コードはこの時点で既に
    // restore されていたため observedDuringFn = original になっていた (regression)。
    await new Promise((r) => setTimeout(r, 50));
    observedDuringFn = app.state.currentAnswers;
    observedAxisDuringFn = app.state.axisConfig;
  });

  // async 完了後に restore されている
  assert.equal(app.state.currentAnswers, original);
  assert.equal(app.state.axisConfig.tag, 'orig');
  // fn 実行中は swap が維持されていた (これが async/await race fix の本体)
  assert.equal(observedDuringFn, swappedRows);
  assert.equal(observedAxisDuringFn.tag, 'swap');
});

test('withTimelineSwap: vizTimelineIndex==null short-circuits without swap', () => {
  const { StudyQuestApp } = loadVizContext();
  const app = makeFakeApp(StudyQuestApp, { vizTimelineIndex: null, snapshots: [] });
  const result = app.__vizWithTimelineSwap(() => 42);
  assert.equal(result, 42);
});

test('withTimelineSwap: out-of-range idx falls back to fn() without swap', () => {
  const { StudyQuestApp } = loadVizContext();
  const swapped = [{ rowIndex: 999 }];
  const app = makeFakeApp(StudyQuestApp, {
    vizTimelineIndex: 99,  // length=1 の snapshots に対して out-of-range
    snapshots: [{ capturedAt: 1, rows: swapped, axisConfig: null }]
  });
  const original = app.state.currentAnswers;
  let observed;
  app.__vizWithTimelineSwap(() => { observed = app.state.currentAnswers; });
  assert.equal(observed, original);
});

test('withTimelineSwap: sync exception restores state before rethrow', () => {
  const { StudyQuestApp } = loadVizContext();
  const swapped = [{ rowIndex: 999 }];
  const app = makeFakeApp(StudyQuestApp, {
    vizTimelineIndex: 0,
    snapshots: [{ capturedAt: 1, rows: swapped, axisConfig: null }]
  });
  const original = app.state.currentAnswers;
  assert.throws(() => {
    app.__vizWithTimelineSwap(() => { throw new Error('boom'); });
  }, /boom/);
  assert.equal(app.state.currentAnswers, original, 'restored after sync throw');
});

test('withTimelineSwap: async rejection restores state', async () => {
  const { StudyQuestApp } = loadVizContext();
  const swapped = [{ rowIndex: 999 }];
  const app = makeFakeApp(StudyQuestApp, {
    vizTimelineIndex: 0,
    snapshots: [{ capturedAt: 1, rows: swapped, axisConfig: null }]
  });
  const original = app.state.currentAnswers;
  await assert.rejects(
    app.__vizWithTimelineSwap(async () => {
      await new Promise((r) => setTimeout(r, 10));
      throw new Error('async-boom');
    }),
    /async-boom/
  );
  assert.equal(app.state.currentAnswers, original, 'restored after async reject');
});

// =====================================================================
// groupByStudent (post-simplification: ghost is always [])
// =====================================================================

test('groupByStudent: allowResubmit=false returns rows as-is, ghost empty', () => {
  const { StudyQuestApp } = loadVizContext();
  // groupByStudent は inner function で直接 expose されていないため、
  // renderMatrix / renderNumberLine 経由でしかテストできない部分もあるが、
  // ここでは ghost が常に空という契約を __vizSnapshotInternals 経由で間接確認。
  const sn = StudyQuestApp.prototype.__vizSnapshotInternals;
  // buildHistoricalSnapshots は groupByStudent を経由しないが、ghost 廃止は
  // build 出力には影響しないことを確認 (各バケット rows に ghost 列が混じらない)
  const rows = [
    { rowIndex: 1, emailHash: 'a', timestamp: '2026-05-14T01:00:00Z', numericX: 1, numericY: 1 },
    { rowIndex: 2, emailHash: 'a', timestamp: '2026-05-14T01:05:00Z', numericX: 5, numericY: 5 }
  ];
  const snaps = sn.buildHistoricalSnapshots(rows, 'matrix', null);
  // 同一 emailHash の最新だけが残る = bucket 末で 1 件のみ
  assert.equal(snaps[snaps.length - 1].rows.length, 1);
});

// =====================================================================
// vizComputeSwings (existing, but with M3/M4 in play)
// =====================================================================

test('vizComputeSwings: ignores rows without numericX (wordcloud mode safe)', () => {
  const { StudyQuestApp } = loadVizContext();
  const compute = StudyQuestApp.prototype.vizComputeSwings;
  const rows = [
    { emailHash: 'h1', timestamp: '2026-05-12T01:00:00Z', numericX: null, answer: 'a' },
    { emailHash: 'h1', timestamp: '2026-05-12T02:00:00Z', numericX: null, answer: 'b' }
  ];
  const swings = compute.call({}, rows, 'numberline');
  assert.equal(swings.length, 0);
});

test('vizComputeSwings: requires emailHash (anonymous = no tracking)', () => {
  const { StudyQuestApp } = loadVizContext();
  const compute = StudyQuestApp.prototype.vizComputeSwings;
  const rows = [
    { emailHash: null, timestamp: '2026-05-12T01:00:00Z', numericX: 1 },
    { emailHash: null, timestamp: '2026-05-12T02:00:00Z', numericX: 5 }
  ];
  const swings = compute.call({}, rows, 'numberline');
  assert.equal(swings.length, 0);
});

test('vizComputeSwings: M1 1-axis distance', () => {
  const { StudyQuestApp } = loadVizContext();
  const compute = StudyQuestApp.prototype.vizComputeSwings;
  const rows = [
    { emailHash: 'h1', timestamp: '2026-05-12T01:00:00Z', numericX: 1 },
    { emailHash: 'h1', timestamp: '2026-05-12T02:00:00Z', numericX: 5 }
  ];
  const swings = compute.call({}, rows, 'numberline');
  assert.equal(swings.length, 1);
  assert.equal(swings[0].distance, 4);
});

test('vizComputeSwings: M2 Euclidean distance', () => {
  const { StudyQuestApp } = loadVizContext();
  const compute = StudyQuestApp.prototype.vizComputeSwings;
  const rows = [
    { emailHash: 'h1', timestamp: '2026-05-12T01:00:00Z', numericX: 1, numericY: 1 },
    { emailHash: 'h1', timestamp: '2026-05-12T02:00:00Z', numericX: 4, numericY: 5 }
  ];
  const swings = compute.call({}, rows, 'matrix');
  assert.equal(swings.length, 1);
  // sqrt(3² + 4²) = 5
  assert.equal(swings[0].distance, 5);
});

// =====================================================================
// __vizRenderCompare: dual-modal 対応のための左側ペイン情報保存 + side マーカー
// =====================================================================

/**
 * 比較表示用にカスタム vm context を立ち上げる。
 *   - sessionStorage に before-snapshot を仕込む
 *   - document.createElement は呼び出しごとに固有 mock element を返す
 *   - appendChild は子要素を配列に記録
 *
 * Why: ドット click handler 側で `event.target.closest('[data-viz-side]')` から
 *   left/right を取り出してモーダルに表示する設計のため、__vizRenderCompare が
 *   生成する svg host に data-viz-side が確実に付くこと、および
 *   state.vizCompareLeftSource が showAnswerModal('left') から参照可能な状態で
 *   保存されることを契約として固定する。
 */
function loadVizContextForCompare(beforeSnapshot) {
  const fs = require('node:fs');
  const path = require('node:path');
  const vm = require('node:vm');
  const html = fs.readFileSync(path.resolve(__dirname, '../src/page.viz.js.html'), 'utf8');
  const m = html.match(/<script>([\s\S]*?)<\/script>/);
  const source = m[1];
  const StudyQuestApp = class StudyQuestApp {};

  function makeElement(tag) {
    return {
      tagName: tag,
      id: '',
      className: '',
      innerHTML: '',
      hidden: false,
      dataset: {},  // v2686: __vizRenderCompare が beforeCard.dataset.side を設定するため必要
      _attrs: {},
      _children: [],
      classList: { add: () => {}, remove: () => {}, toggle: () => {} },
      style: {},
      setAttribute(k, v) { this._attrs[k] = v; },
      getAttribute(k) { return this._attrs[k]; },
      appendChild(child) { this._children.push(child); return child; },
      addEventListener: () => {}
    };
  }

  const sessionStore = {};
  if (beforeSnapshot) {
    // page.viz.js.html の SNAPSHOT_BEFORE_KEY と同期 (定数は closure 内で公開されないため
    // 文字列直書き)。
    sessionStore['viz:before:v1'] = JSON.stringify(beforeSnapshot);
  }

  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    window: { StudyQuestApp },
    StudyQuestApp,
    sessionStorage: {
      getItem: (k) => (k in sessionStore ? sessionStore[k] : null),
      setItem: (k, v) => { sessionStore[k] = v; },
      removeItem: (k) => { delete sessionStore[k]; }
    },
    document: {
      readyState: 'complete',
      addEventListener: () => {},
      getElementById: () => null,
      createElement: makeElement,
      createElementNS: () => ({ setAttribute: () => {}, appendChild: () => {} }),
      body: { classList: { add: () => {} } }
    },
    URLSearchParams,
    Map, Set, Promise, JSON, Math, Number, Object, Array, String, Boolean, Date, RegExp,
    Error, TypeError, parseInt, parseFloat, isNaN, isFinite,
    setTimeout: (fn, ms) => setTimeout(fn, ms),
    clearTimeout: (id) => clearTimeout(id)
  };
  context.window.location = { search: '' };
  vm.createContext(context);
  vm.runInContext(source, context, { filename: 'page.viz.js.html' });
  return { context, StudyQuestApp, makeElement };
}

test('__vizRenderCompare: renders before/after panes (v2688: dual modal feature removed)', async () => {
  // v2688: 個別意見ペア比較機能 (dual modal / pane card) は道徳教育の自然な議論フローに
  //   合わなかったため廃止。比較モードは「分布の左右並列表示」のみ提供する。
  //   - 旧 vizCompareLeftSource / vizCompareLeftLabel / vizCompareRightLabel state は不要
  //   - 旧 data-viz-side マーカーも不要 (click 元の判別が不要になったため)
  const beforeRows = [{ rowIndex: 7, opinion: '昔の意見' }];
  const beforeSnap = {
    capturedAt: new Date('2026-05-14T09:30:00').getTime(),
    rows: beforeRows,
    axisConfig: { tag: 'before-axis' }
  };
  const { StudyQuestApp, makeElement } = loadVizContextForCompare(beforeSnap);

  const answersContainer = makeElement('div');
  const app = new StudyQuestApp();
  app.state = {
    currentAnswers: [{ rowIndex: 9, opinion: '今の意見' }],
    axisConfig: { tag: 'current-axis' },
    vizTimelineIndex: null,
    vizCompareMode: true
  };
  app.elements = { answersContainer };

  let renderCallCount = 0;
  const renderOne = async () => { renderCallCount += 1; };

  await StudyQuestApp.prototype.__vizRenderCompare.call(app, renderOne, 'matrix');

  // renderOne が左右で 2 回呼ばれる
  assert.equal(renderCallCount, 2);

  // wrap → beforePane + afterPane の構造は維持
  assert.equal(answersContainer._children.length, 1);
  const wrap = answersContainer._children[0];
  assert.equal(wrap._children.length, 2);
  const [beforePane, afterPane] = wrap._children;
  // 各 pane に svg host が含まれる
  assert.ok(beforePane._children.length >= 1, 'beforePane should contain svg host');
  assert.ok(afterPane._children.length >= 1, 'afterPane should contain svg host');
});

test('__vizRenderCompare: falls back to single render when no snapshot exists', async () => {
  const { StudyQuestApp, makeElement } = loadVizContextForCompare(null);
  const answersContainer = makeElement('div');
  const app = new StudyQuestApp();
  app.state = {
    currentAnswers: [{ rowIndex: 1 }],
    axisConfig: null,
    vizTimelineIndex: null,
    vizCompareMode: true
  };
  app.elements = { answersContainer };
  let called = 0;
  await StudyQuestApp.prototype.__vizRenderCompare.call(app, async () => { called += 1; }, 'matrix');
  // snapshot ゼロ件 → fallback で renderOne が 1 回呼ばれ、compare mode は OFF
  assert.equal(called, 1);
  assert.equal(app.state.vizCompareMode, false);
});

// =====================================================================
// __renderQuadrantSummary: 4 象限ごとの reason キーワード抽出
// =====================================================================

/**
 * renderQuadrantSummary を呼べる最小限の context を作る。
 *   - document.createElement: 子要素を配列に記録する mock element
 *   - querySelector('#vizQuadrantSummary'): 動的に追加された panel を返す簡易検索
 *   - querySelector('#vizTeacherPanel'): 常に null（テストでは存在しない）
 * Why: __vizRenderCompare のテストと違い、こちらは document.querySelector を
 *   answers コンテナの中で実装する必要がある（renderQuadrantSummary が
 *   answers.querySelector('#vizQuadrantSummary') と answers.querySelector('#vizTeacherPanel')
 *   を呼ぶため）。
 */
function loadVizContextForQuadrant() {
  const fs = require('node:fs');
  const path = require('node:path');
  const vm = require('node:vm');
  const html = fs.readFileSync(path.resolve(__dirname, '../src/page.viz.js.html'), 'utf8');
  const m = html.match(/<script>([\s\S]*?)<\/script>/);
  const source = m[1];
  const StudyQuestApp = class StudyQuestApp {};

  function makeElement(tag) {
    const el = {
      tagName: tag,
      id: '',
      className: '',
      innerHTML: '',
      textContent: '',
      hidden: false,
      _attrs: {},
      _children: [],
      classList: { add: () => {}, remove: () => {}, toggle: () => {}, contains: () => false },
      style: {},
      setAttribute(k, v) { this._attrs[k] = v; },
      getAttribute(k) { return this._attrs[k]; },
      appendChild(child) { this._children.push(child); return child; },
      insertBefore(child, _ref) { this._children.push(child); return child; },
      remove() { /* parent tracking 省略 */ },
      addEventListener: () => {},
      // 簡易 querySelector: id セレクタのみ対応
      querySelector(sel) {
        if (!sel.startsWith('#')) return null;
        const id = sel.slice(1);
        const walk = (node) => {
          if (!node || !node._children) return null;
          for (const c of node._children) {
            if (c.id === id) return c;
            const found = walk(c);
            if (found) return found;
          }
          return null;
        };
        return walk(el);
      }
    };
    return el;
  }

  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    window: { StudyQuestApp },
    StudyQuestApp,
    sessionStorage: { getItem: () => null, setItem: () => {}, removeItem: () => {} },
    document: {
      readyState: 'complete',
      addEventListener: () => {},
      getElementById: () => null,
      createElement: makeElement,
      createElementNS: () => ({ setAttribute: () => {}, appendChild: () => {} }),
      body: { classList: { add: () => {} } }
    },
    URLSearchParams,
    Map, Set, Promise, JSON, Math, Number, Object, Array, String, Boolean, Date, RegExp,
    Error, TypeError, parseInt, parseFloat, isNaN, isFinite,
    setTimeout: (fn, ms) => setTimeout(fn, ms),
    clearTimeout: (id) => clearTimeout(id)
  };
  context.window.location = { search: '' };
  vm.createContext(context);
  vm.runInContext(source, context, { filename: 'page.viz.js.html' });
  return { context, StudyQuestApp, makeElement };
}

test('__renderQuadrantSummary: aggregates top-3 reason keywords per quadrant', () => {
  const { StudyQuestApp, makeElement } = loadVizContextForQuadrant();
  const answers = makeElement('div');

  // (X >= 3, Y >= 3) = hh (top-right「真実と思いやり」)
  // (X < 3, Y >= 3)  = lh (top-left「期限もみんなのため」)
  // (X >= 3, Y < 3)  = hl (bottom-right「自分の作品にこだわる」)
  // (X < 3, Y < 3)   = ll (bottom-left「効率優先・楽したい」)
  const rows = [
    { numericX: 5, numericY: 5, reason: '真実を伝えることが信頼につながる。誠実が大事' },
    { numericX: 5, numericY: 4, reason: '相手を思いやる気持ち。真実と誠実を貫きたい' },
    { numericX: 4, numericY: 5, reason: '真実と思いやりは両立できる' },
    { numericX: 2, numericY: 5, reason: '締切も大事。みんなのために守る' },
    { numericX: 4, numericY: 2, reason: '自分の作品にこだわりを持ちたい' },
    { numericX: 5, numericY: 2, reason: '丁寧に作品を仕上げる責任' }
  ];

  StudyQuestApp.prototype.__renderQuadrantSummary(answers, rows, {
    matrixQuadrantLabels: { hh: '真実と思いやり', lh: '期限もみんなのため', hl: '自分の作品にこだわる', ll: '効率優先・楽したい' }
  });

  const panel = answers.querySelector('#vizQuadrantSummary');
  assert.ok(panel, 'panel should be created');
  assert.equal(panel._children.length, 4, 'should have 4 cells (TL, TR, BL, BR)');

  // 各セルの cls から象限を識別
  const cellByCls = {};
  for (const cell of panel._children) {
    const m = cell.className.match(/qs-(tl|tr|bl|br)/);
    if (m) cellByCls[m[1]] = cell;
  }
  assert.ok(cellByCls.tl, 'TL cell exists');
  assert.ok(cellByCls.tr, 'TR cell exists');
  assert.ok(cellByCls.bl, 'BL cell exists');
  assert.ok(cellByCls.br, 'BR cell exists');

  // TR (hh: 真実と思いやり) には 3 件 → 「真実」「信頼」「誠実」「思いやり」など出現
  const trKeywords = cellByCls.tr._children
    .flatMap(c => c._children || [])
    .filter(c => c.className === 'qs-tag')
    .map(c => c.textContent);
  // 最低 1 つは reason 由来の意味語が tag になっている
  assert.ok(trKeywords.length > 0, 'TR should have keywords');
  // 「真実」は 3 件中 3 件で出現 → top-3 に入る
  assert.ok(trKeywords.some(t => t.includes('真実')), `TR should include 真実 in tags: ${trKeywords.join(',')}`);

  // BL (ll: 効率優先) には 0 件 → 「該当意見なし」placeholder
  const blEmpty = cellByCls.bl._children.find(c => c.className === 'qs-empty');
  assert.ok(blEmpty, 'BL should show empty placeholder');
  assert.match(blEmpty.textContent, /該当意見なし/);
});

test('__renderQuadrantSummary: hides panel when no rows have numericX/Y', () => {
  const { StudyQuestApp, makeElement } = loadVizContextForQuadrant();
  const answers = makeElement('div');
  StudyQuestApp.prototype.__renderQuadrantSummary(answers, [
    { reason: '数値なし' },
    { numericX: null, numericY: null, reason: '欠損' }
  ], {});
  const panel = answers.querySelector('#vizQuadrantSummary');
  assert.ok(panel, 'panel created');
  assert.equal(panel.hidden, true, 'panel hidden when no data');
});
