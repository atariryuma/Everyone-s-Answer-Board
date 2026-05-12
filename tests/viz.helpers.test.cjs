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
    setTimeout: (fn) => fn,
    clearTimeout: () => {}
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
// __vizSetMode: 授業中のモード即時切替（道徳 5 モード）
// =====================================================================

test('__vizSetMode: updates state.boardMode for valid mode', () => {
  const { StudyQuestApp } = loadVizContext();
  const setMode = StudyQuestApp.prototype.__vizSetMode;
  const fakeApp = {
    state: { boardMode: 'board' },
    renderBoard: () => {}
  };
  setMode.call(fakeApp, 'numberline');
  assert.equal(fakeApp.state.boardMode, 'numberline');
});

test('__vizSetMode: silently rejects invalid mode (no state change)', () => {
  const { StudyQuestApp } = loadVizContext();
  const setMode = StudyQuestApp.prototype.__vizSetMode;
  const fakeApp = {
    state: { boardMode: 'numberline' },
    renderBoard: () => {}
  };
  setMode.call(fakeApp, 'INVALID');
  assert.equal(fakeApp.state.boardMode, 'numberline');
});

test('__vizSetMode: resets timeline/compare on mode change', () => {
  const { StudyQuestApp } = loadVizContext();
  const setMode = StudyQuestApp.prototype.__vizSetMode;
  let renderCalled = false;
  const fakeApp = {
    state: { boardMode: 'numberline', vizTimelineIndex: 3, vizCompareMode: true },
    renderBoard: () => { renderCalled = true; }
  };
  setMode.call(fakeApp, 'pie');
  assert.equal(fakeApp.state.vizTimelineIndex, null);
  assert.equal(fakeApp.state.vizCompareMode, false);
  assert.equal(renderCalled, true);
});

test('__vizSetMode: accepts all 5 modes (board/numberline/matrix/wordcloud/pie)', () => {
  const { StudyQuestApp } = loadVizContext();
  const setMode = StudyQuestApp.prototype.__vizSetMode;
  const modes = ['board', 'numberline', 'matrix', 'wordcloud', 'pie'];
  for (const m of modes) {
    const fakeApp = { state: { boardMode: 'auto' }, renderBoard: () => {} };
    setMode.call(fakeApp, m);
    assert.equal(fakeApp.state.boardMode, m, `mode "${m}" should be accepted`);
  }
});
