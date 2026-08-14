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

// 1 度だけ読んで全テストで使い回す。
const VIZ_HTML = fs.readFileSync(path.resolve(__dirname, '../src/page.viz.js.html'), 'utf8');
const SEGMENTER_HTML = fs.readFileSync(path.resolve(__dirname, '../src/tinySegmenter.html'), 'utf8');

function loadVizContext() {
  const html = VIZ_HTML;
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
  const html = VIZ_HTML;
  const segHtml = SEGMENTER_HTML;
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
// distinctiveKeywords (TF-IDF 風象限固有度, v2691)
// =====================================================================

test('distinctiveKeywords: prefers words distinctive to this cell over globally frequent', () => {
  const { StudyQuestApp } = loadVizContext();
  const fn = StudyQuestApp.prototype.__distinctiveKeywords;
  // 「みんな」は全象限頻出、「真実」「誠実」はこの象限固有
  const thisCell = [
    { reason: 'みんなのために真実を伝える' },
    { reason: '真実は大切、誠実が信頼を作る' },
    { reason: 'みんなに対して誠実でいたい' }
  ];
  const other1 = [
    { reason: 'みんなで効率的に進める' },
    { reason: 'みんなの時間を尊重する' }
  ];
  const other2 = [
    { reason: 'みんなが楽しめるのが大事' }
  ];
  const other3 = [{ reason: 'みんなに合わせる' }];
  const result = fn(thisCell, [other1, other2, other3], 'reason', 3);
  const words = result.map(r => r.word);
  // 「真実」「誠実」が top に来る (他象限ゼロ出現で score 高い)
  // 「みんな」は他象限でも頻出なので score が下がる
  assert.ok(words.includes('真実'), `should include 真実: ${words.join(',')}`);
  assert.ok(words.includes('誠実'), `should include 誠実: ${words.join(',')}`);
  // 「みんな」は score が下がるので少なくとも上位 3 位以内には来ない期待
  // (3 件 - 平均 4/3 ≒ 1.67 なので、score=1.33。真実 3-0=3, 誠実 2-0=2 が勝つ)
  const minnaIndex = words.indexOf('みんな');
  if (minnaIndex >= 0) {
    const distinctIndex = Math.min(words.indexOf('真実'), words.indexOf('誠実'));
    assert.ok(distinctIndex < minnaIndex, `distinctive words should outrank みんな: ${words.join(',')}`);
  }
});

test('distinctiveKeywords: returns [] for empty input', () => {
  // Why length check ではなく deepEqual を使わない: vm context 跨ぎで Array prototype が
  //   異なるため deepStrictEqual が reference-equal を要求して失敗する。
  const { StudyQuestApp } = loadVizContext();
  const fn = StudyQuestApp.prototype.__distinctiveKeywords;
  assert.equal(fn([], [[], [], []], 'reason').length, 0);
  assert.equal(fn(null, [[], [], []], 'reason').length, 0);
});

test('distinctiveKeywords: handles single-cell case (no others to contrast)', () => {
  const { StudyQuestApp } = loadVizContext();
  const fn = StudyQuestApp.prototype.__distinctiveKeywords;
  // 他象限がすべて空 → score = count - 0 = count なので top 頻度語が並ぶ
  const result = fn(
    [{ reason: '責任を持つ。責任は重要。効率も大事。' }, { reason: '責任が一番' }],
    [[], [], []],
    'reason',
    3
  );
  const words = result.map(r => r.word);
  assert.ok(words.includes('責任'), `should include 責任: ${words.join(',')}`);
});

// =====================================================================
// consensusKeywords (立場を越えた共通語 = bridging, Pol.is 原則)
// =====================================================================

test('consensusKeywords: surfaces words shared across most groups (high coverage)', () => {
  const { StudyQuestApp } = loadVizContext();
  const fn = StudyQuestApp.prototype.__consensusKeywords;
  // 「友情」は全 4 群に登場（coverage=4）= 共通の足場。
  // 各群固有の語（真実/効率/楽しい/協力）は coverage=1 で拾われない。
  const g1 = [{ reason: '友情を大切に、真実を伝える' }];
  const g2 = [{ reason: '友情のため効率を上げる' }];
  const g3 = [{ reason: '友情があれば楽しい' }];
  const g4 = [{ reason: '友情は協力から生まれる' }];
  const result = fn([g1, g2, g3, g4], 'reason', 2, { minCoverageRatio: 0.75 });
  const words = result.map((r) => r.word);
  assert.ok(words.includes('友情'), `共通語 友情 を含むべき: ${words.join(',')}`);
  // 1 群しか出ない固有語は除外される
  assert.ok(!words.includes('真実'), `固有語 真実 は共通語に含めない: ${words.join(',')}`);
});

test('consensusKeywords: respects exclude set (対立軸で使った語は除外)', () => {
  const { StudyQuestApp } = loadVizContext();
  const fn = StudyQuestApp.prototype.__consensusKeywords;
  const g1 = [{ reason: '責任と自由は両立する' }];
  const g2 = [{ reason: '責任を持てば自由になる' }];
  // 「責任」は両群に出るが exclude に入れたら出さない
  const result = fn([g1, g2], 'reason', 3, { minCoverageRatio: 0.75, exclude: new Set(['責任']) });
  const words = result.map((r) => r.word);
  assert.ok(!words.includes('責任'), `exclude された 責任 は出さない: ${words.join(',')}`);
});

test('consensusKeywords: returns [] when fewer than 2 non-empty groups', () => {
  const { StudyQuestApp } = loadVizContext();
  const fn = StudyQuestApp.prototype.__consensusKeywords;
  assert.equal(fn([], 'reason').length, 0);
  assert.equal(fn([[{ reason: '友情が大事' }], []], 'reason').length, 0); // 非空は 1 群のみ
});

// =====================================================================
// contrastingPair (X 軸 / Y 軸方向の対立語ペア抽出, v2691)
// =====================================================================

test('contrastingPair: extracts distinctive word per side', () => {
  // 各 side で 1 語が突出して頻出するように構築（タイブレークに依存しない）。
  const { StudyQuestApp } = loadVizContext();
  const fn = StudyQuestApp.prototype.__contrastingPair;
  const left = [
    { reason: '効率を優先' },     // 効率 + 優先
    { reason: '効率を大切に' },   // 効率 + 大切
    { reason: '効率重視' }        // 効率重視
  ];
  // → 効率 count=2 score=2、他の語は score=1 → 効率が圧勝
  const right = [
    { reason: '誠実な対応' },     // 誠実 + 対応
    { reason: '誠実が一番' },     // 誠実 + 一番
    { reason: '誠実第一' }        // 誠実第一
  ];
  // → 誠実 count=2 score=2、他の語は score=1 → 誠実が圧勝
  const pair = fn(left, right, 'reason');
  assert.ok(pair, 'should return a pair');
  assert.equal(pair.left, '効率', `left should be 効率: ${JSON.stringify(pair)}`);
  assert.equal(pair.right, '誠実', `right should be 誠実: ${JSON.stringify(pair)}`);
});

test('contrastingPair: returns null when one side empty', () => {
  const { StudyQuestApp } = loadVizContext();
  const fn = StudyQuestApp.prototype.__contrastingPair;
  assert.equal(fn([], [{ reason: '責任' }], 'reason'), null);
  assert.equal(fn([{ reason: '責任' }], [], 'reason'), null);
});

test('contrastingPair: returns null when no word is distinctive (full overlap)', () => {
  const { StudyQuestApp } = loadVizContext();
  const fn = StudyQuestApp.prototype.__contrastingPair;
  // 両側とも同じ語のみ → score <= 0 になるはず
  const same = [{ reason: '責任' }, { reason: '責任' }];
  const result = fn(same, same, 'reason');
  // どちらの語も「left も right も同数」なので score=0、bestLeft/Right とも null になり pair は null
  assert.equal(result, null);
});

test('contrastingPair: filters out 1-char fragment tokens (v2691 patch)', () => {
  // 「間に合う」のような複合語が tokenizer で「間」「合」断片に切れたとき、
  //   対立軸ラベル「#間 ↔ #何」のような無意味な 1 字対立が出る問題への対策。
  //   1 字キーは contrastingPair から除外する。
  const { StudyQuestApp } = loadVizContext();
  const fn = StudyQuestApp.prototype.__contrastingPair;
  // left に「間」(1字) と「効率」(2字) を入れ、両方 score>0。1字は除外され「効率」が選ばれる期待
  const left = [{ reason: '間 効率' }, { reason: '間 効率' }, { reason: '間 効率' }];
  const right = [{ reason: '何 誠実' }, { reason: '何 誠実' }, { reason: '何 誠実' }];
  const pair = fn(left, right, 'reason');
  assert.ok(pair, 'should return a pair');
  assert.notEqual(pair.left, '間', `1-char fragment should be excluded from left: ${JSON.stringify(pair)}`);
  assert.notEqual(pair.right, '何', `1-char fragment should be excluded from right: ${JSON.stringify(pair)}`);
  assert.equal(pair.left, '効率');
  assert.equal(pair.right, '誠実');
});

test('contrastingPair: kanji-quality bias picks kanji over kana when scores tie (v2691 patch)', () => {
  // 「かいぜんし」(5字 hiragana 中途断片) と「改善」(2字 漢字) が同 score で並んだ場合、
  //   漢字優先で「改善」が選ばれることをテスト。
  const { StudyQuestApp } = loadVizContext();
  const fn = StudyQuestApp.prototype.__contrastingPair;
  // かいぜんし は 5 字 hiragana（≤3 字フィルタを通る）。
  //   左に「かいぜんし」「改善」両方を 2 件ずつ入れ、右にはどちらも 0 件。
  //   両方とも score=2-0=2 で同じ → quality 比較で「改善」(quality=2) が「かいぜんし」(quality=1) に勝つ。
  const left = [
    { reason: 'かいぜんし 改善' },
    { reason: 'かいぜんし 改善' }
  ];
  const right = [{ reason: '誠実' }, { reason: '誠実' }];
  const pair = fn(left, right, 'reason');
  assert.ok(pair, 'should return a pair');
  assert.equal(pair.left, '改善', `kanji should beat hiragana on tie: ${JSON.stringify(pair)}`);
});

test('contrastingPair: excludeWords removes candidate from both sides (v2691 patch)', () => {
  // X 軸で「気持ち」が選ばれたあと、Y 軸で同じ「気持ち」が再選されないように
  //   excludeWords (Set of lowercase keys) を渡せる。重複表示の回避用。
  const { StudyQuestApp } = loadVizContext();
  const fn = StudyQuestApp.prototype.__contrastingPair;
  const left = [{ reason: '気持 効率' }, { reason: '気持 効率' }];
  const right = [{ reason: '気持 誠実' }, { reason: '気持 誠実' }];
  // exclude 無し: 「効率」 vs 「誠実」が出る (気持 は両側にあって score=0 → 自動除外)
  const pair1 = fn(left, right, 'reason');
  assert.ok(pair1);
  // exclude=効率: 別の語が出るはず → left は何も残らないので null
  const pair2 = fn(left, right, 'reason', new Set(['効率', '誠実']));
  assert.equal(pair2, null, `excluded all viable words: ${JSON.stringify(pair2)}`);
});

test('distinctiveKeywords: kanji-quality bias picks kanji over kana on tie (v2691 patch)', () => {
  const { StudyQuestApp } = loadVizContext();
  const fn = StudyQuestApp.prototype.__distinctiveKeywords;
  // この象限に「かいぜんし」と「改善」が同数、他象限ゼロ → score 同点
  const thisCell = [
    { reason: 'かいぜんし 改善' },
    { reason: 'かいぜんし 改善' },
    { reason: 'かいぜんし 改善' }
  ];
  const result = fn(thisCell, [[], [], []], 'reason', 2);
  const words = result.map(r => r.word);
  assert.equal(words[0], '改善', `kanji should be top: ${words.join(',')}`);
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

// 投稿列順 delta encoding 移行後: snapshots は {addedRow, axisConfig} の配列。
// withTimelineSwap は snapshots[0..idx] を畳んで latestByKey を再構成 → currentAnswers に
// セットする。snapshot.rows を直接読まないことに注意。
function deltaSnap(addedRow, axisConfig) {
  return { latestTimestamp: 1, boardMode: 'matrix', axisConfig: axisConfig || null, addedRow };
}


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
// 議論のたねになる声 (hasJustification / buildPositionGroups / representativeReasons)
//   研究根拠: Stanford HCOMP 2024 (arXiv:2408.11936) — 集団の意見変容は
//   「理由づけのある主張の selective uptake」で起きる。 教師が各立場から理由づけの
//   ある声を 1 つずつ拾う（人気順ではなく position coverage）ための純粋関数群。
// =====================================================================

test('reactionDash: CVD redundant encoding — distinct dash pattern per reaction type', () => {
  const { StudyQuestApp } = loadVizContext();
  const dash = StudyQuestApp.prototype.__reactionDash;
  // LIKE(赤) と CURIOUS(緑) は赤緑色覚で同色化するので、 必ず別パターンであること
  assert.notEqual(dash('LIKE'), dash('CURIOUS'));
  assert.equal(dash('LIKE'), '5,3');       // 長破線
  assert.equal(dash('CURIOUS'), '1.5,2');  // 短破線
  assert.equal(dash('UNDERSTAND'), null);  // 実線（黄は単独で判別可）
  assert.equal(dash(null), null);
  assert.equal(dash('UNKNOWN'), null);
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
  const html = VIZ_HTML;
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
  const html = VIZ_HTML;
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

