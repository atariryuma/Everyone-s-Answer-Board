/**
 * v2690 自動化リグレッション抑止テスト
 *
 * AdminPanel から手動操作 UI を撤廃した変更が後で壊れないように、HTML/JS の
 * 構造的特徴を静的に検証する。VM での動作テストではなく、ソース文字列の検査。
 *
 * 守るべき不変条件:
 *   1) 「🔍 検証して接続」ボタンは hidden block 以外に存在しない (legacy 残置のみ可)
 *   2) 表示モード/軸ラベル等の補助設定は `<details>` で畳まれている
 *   3) 「💾 設定を保存」ボタンは display:none で hidden 化されている
 *   4) #autosave-status と setupAutoSave の対応関係が成立している
 *   5) autoAnalyzeFormUrl が debounce 経由で validateCompleteUrl を自動呼び出しする
 *   6) saveDraft が opts.silent / opts.onComplete を受け取れる
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const ADMIN_HTML = fs.readFileSync(
  path.resolve(__dirname, '../src/AdminPanel.html'),
  'utf8'
);
const ADMIN_JS = fs.readFileSync(
  path.resolve(__dirname, '../src/AdminPanel.js.html'),
  'utf8'
);

// ─────────────────────────────────────────────────────────────────────────
// HTML 側: UI 簡素化の不変条件
// ─────────────────────────────────────────────────────────────────────────

test('AdminPanel: 「🔍 検証して接続」ボタンは hidden block 内のみ (常時表示なし)', () => {
  // legacy 残置の hidden div には button が存在してよいが、それ以外には無いこと。
  // HTML を hidden div で切り分けて、可視ブロック側にボタン id="url-validate" が
  // 含まれていないことを確認する。
  const visibleSections = ADMIN_HTML.replace(
    /<div\s+style="display:none;"[\s\S]*?<\/div>/g,
    ''
  );
  assert.ok(
    !/id="url-validate"/.test(visibleSections),
    '可視ブロックに #url-validate が残っている (削除されていない)'
  );
  // legacy hidden は残っていてよい
  assert.match(ADMIN_HTML, /id="url-validate"/, 'legacy hidden #url-validate は残置されているはず');
});

test('AdminPanel: 表示モード/軸ラベルは <details> で畳まれている', () => {
  // #visualization-section が <details> に変わっていること
  assert.match(
    ADMIN_HTML,
    /<details[^>]*id="visualization-section"/,
    '#visualization-section が details になっていない'
  );
  // フォーム連携の代替経路 (新規作成 + 既存一覧) が単一の details にまとめられている
  assert.match(
    ADMIN_HTML,
    /他の方法 \(フォームが手元にない \/ 一覧から選ぶ\)/,
    '代替経路の details summary が見当たらない'
  );
  // 代替経路 details の中にテンプレート作成と既存一覧の両方が含まれている
  assert.match(
    ADMIN_HTML,
    /id="template-type-select"/,
    'テンプレート作成 select が見当たらない'
  );
  assert.match(
    ADMIN_HTML,
    /id="form-select"/,
    '既存フォーム一覧 select が見当たらない'
  );
});

test('AdminPanel: 個別ボード authoring に ① ② ③ のステップ番号が付与されている', () => {
  // セットアップ順序を視覚化する step-number badge が 3 つ揃っているか。
  const matches = ADMIN_HTML.match(/class="step-number"/g) || [];
  assert.ok(matches.length >= 3, `step-number badge が 3 つ未満: ${matches.length}`);
});

test('AdminPanel: 「💾 設定を保存」ボタンは hidden 化されている (autosave に置換済)', () => {
  // #save-config-changes は display:none を持っているはず
  const match = ADMIN_HTML.match(
    /<button[^>]*id="save-config-changes"[^>]*>/
  );
  assert.ok(match, '#save-config-changes ボタンが見つからない');
  assert.match(
    match[0],
    /display:\s*none/,
    '#save-config-changes が hidden 化されていない'
  );
});

test('AdminPanel: #autosave-status が存在する (自動保存の視覚 feedback)', () => {
  assert.match(
    ADMIN_HTML,
    /id="autosave-status"/,
    '#autosave-status が見当たらない (autosave 状態表示用)'
  );
  // aria-live で screen reader にも届くこと
  const match = ADMIN_HTML.match(/<[^>]*id="autosave-status"[^>]*>/);
  assert.match(match[0], /aria-live/, '#autosave-status に aria-live が無い');
});

test('AdminPanel: プロファイル保存ボタンは「新規保存」(別名で新規作成) ラベルで、別動線として残る', () => {
  // 旧「複製」表記から「新規保存」に変更。意味は「現在の状態を別名で新しい profile として保存」。
  // 同名 profile への上書きは active-profile-edit-bar の「このプロファイルを更新」ボタンで行う。
  const saveBtnSection = ADMIN_HTML.match(
    /<button[^>]*id="save-current-as-profile-btn"[\s\S]*?<\/button>/
  );
  assert.ok(saveBtnSection, 'プロファイル新規保存ボタンが見当たらない');
  assert.match(
    saveBtnSection[0],
    /新規保存|保存/,
    'プロファイル保存ボタンが「新規保存」用途のラベルになっていない'
  );
});

test('AdminPanel.html: 左パネル末尾に「💾 プロファイル『〇』に保存」ボタン (save-to-profile-btn) がある', () => {
  // 「編集 → 保存」フローを自然にするため active row inline ではなく左パネル末尾に集約。
  // 表示は hidden 既定で、applyConfig が activeProfile を見て表示/dataset を更新する (drift-proof)。
  assert.match(
    ADMIN_HTML,
    /id="save-to-profile-btn"/,
    'save-to-profile-btn が左パネルに存在しない'
  );
  assert.match(
    ADMIN_HTML,
    /id="save-to-profile-name"/,
    'save-to-profile-name span (動的ラベル用) が無い'
  );
});

test('AdminPanel.js.html: applyConfig が save-to-profile-btn の dataset と label を毎回再生成する (drift 防止)', () => {
  // editBar 時代の textContent 読み出しで起きた「別 profile を誤上書き」を再発させない。
  assert.match(
    ADMIN_JS,
    /save-to-profile-btn[\s\S]{0,400}config\.activeProfile/,
    'applyConfig 内で save-to-profile-btn を activeProfile に同期していない'
  );
  assert.match(
    ADMIN_JS,
    /btn\.dataset\.profileName\s*=\s*active/,
    'save-to-profile-btn の dataset.profileName が active から動的セットされていない'
  );
});

test('AdminPanel.js.html: handleActiveProfileUpdate は呼び出し側からの name 引数で動く', () => {
  const fn = ADMIN_JS.match(/function handleActiveProfileUpdate\s*\(([^)]*)\)/);
  assert.ok(fn, 'handleActiveProfileUpdate 関数が見つからない');
  assert.match(fn[1], /\bname\b/, 'handleActiveProfileUpdate は name 引数を受け取る必要がある');
});

// ─────────────────────────────────────────────────────────────────────────
// JS 側: ハンドラと behavior の不変条件
// ─────────────────────────────────────────────────────────────────────────

test('AdminPanel.js.html: applyVisualizationConfig の boardMode allowlist が HTML の <option> と全件一致 (silent downgrade 防止)', () => {
  // Why: allowlist が HTML より狭いと、profile を「読み込んで編集」した時に UI が
  //   silently 'auto' に downgrade し、続く autosave で profile.boardMode が破壊される。
  // HTML 側の option value を抽出
  const selectBlock = ADMIN_HTML.match(/<select[^>]*id="board-mode-select"[\s\S]*?<\/select>/);
  assert.ok(selectBlock, '#board-mode-select が見つからない');
  const htmlValues = (selectBlock[0].match(/value="([a-z]+)"/g) || []).map(s => s.replace(/value="|"/g, ''));
  // JS 側の allowlist を抽出
  const jsAllowed = ADMIN_JS.match(/const allowed = \[([^\]]+)\];/);
  assert.ok(jsAllowed, 'applyVisualizationConfig 内の allowlist が見つからない');
  const jsValues = jsAllowed[1].match(/'([a-z]+)'/g).map(s => s.replace(/'/g, ''));
  // HTML の各 value が JS allowlist に含まれていること
  htmlValues.forEach((v) => {
    assert.ok(jsValues.includes(v),
      `HTML option value="${v}" が JS allowlist に無い → 読み込んで編集で silent downgrade`);
  });
});

test('AdminPanel.html: 必須フィールドが data-autosave 属性でマーキングされている', () => {
  const expectedIds = [
    'answer-column-select',
    'reason-column-select',
    'class-column-select',
    'name-column-select',
    'numeric-x-column-select',
    'numeric-y-column-select',
    'show-names',
    'show-reactions',
    'board-mode-select',
    'x-min-label', 'x-max-label', 'y-min-label', 'y-max-label',
    'quadrant-lh-label', 'quadrant-hh-label', 'quadrant-ll-label', 'quadrant-hl-label',
    'allow-resubmit'
  ];
  expectedIds.forEach((id) => {
    // id="X" を含む要素タグの中に data-autosave が同居していることを確認。
    //   タグ閉じ `>` まで読み、その中に data-autosave があるかを正規表現で検査。
    const tagMatch = ADMIN_HTML.match(new RegExp(`<[^>]*id="${id}"[^>]*>`));
    assert.ok(tagMatch, `id="${id}" の要素が AdminPanel.html に存在しない`);
    assert.match(
      tagMatch[0],
      /\bdata-autosave\b/,
      `id="${id}" に data-autosave 属性が付いていない (自動保存対象から漏れている)`
    );
  });
});

test('AdminPanel.js: setupAutoSave は data-autosave 要素を列挙している', () => {
  assert.match(
    ADMIN_JS,
    /querySelectorAll\s*\(\s*['"]\[data-autosave\]['"]\s*\)/,
    'setupAutoSave 内で querySelectorAll("[data-autosave]") が呼ばれていない'
  );
});

test('AdminPanel.js: setupAutoSave / scheduleAutoSave / autoSaveDraft が定義されている', () => {
  assert.match(ADMIN_JS, /function\s+setupAutoSave\s*\(/, 'setupAutoSave 未定義');
  assert.match(ADMIN_JS, /function\s+scheduleAutoSave\s*\(/, 'scheduleAutoSave 未定義');
  assert.match(ADMIN_JS, /function\s+autoSaveDraft\s*\(/, 'autoSaveDraft 未定義');
});

test('AdminPanel.js: setupAutoSave が初期化バッチから呼ばれている', () => {
  assert.match(
    ADMIN_JS,
    /setupAutoSave\s*\(\s*\)/,
    'setupAutoSave が initializeBatchedAdminPanel から呼ばれていない'
  );
});

test('AdminPanel.js: saveDraft が opts.silent / opts.onComplete を受け取れる', () => {
  // saveDraft(opts) のシグネチャ + silent 分岐 + onComplete コール
  const fnMatch = ADMIN_JS.match(
    /function\s+saveDraft\s*\(\s*opts\s*\)\s*\{[\s\S]*?function\s+executeDraftSave/
  );
  assert.ok(fnMatch, 'saveDraft(opts) のシグネチャが期待通りでない');
  const body = fnMatch[0];
  assert.match(body, /opts\.silent/, 'silent 分岐が無い');
  assert.match(body, /onComplete/, 'onComplete コールバックが無い');
});

test('AdminPanel.js: autoAnalyzeFormUrl が debounce 経由で validateCompleteUrl を呼ぶ', () => {
  const fnMatch = ADMIN_JS.match(
    /function\s+autoAnalyzeFormUrl\s*\([\s\S]*?\n\s{2}\}/
  );
  assert.ok(fnMatch, 'autoAnalyzeFormUrl 関数本体が見つからない');
  const body = fnMatch[0];
  assert.match(
    body,
    /validateCompleteUrl/,
    'autoAnalyzeFormUrl から validateCompleteUrl が呼ばれていない'
  );
  assert.match(
    body,
    /debounce|setTimeout/,
    'autoAnalyzeFormUrl で debounce/setTimeout が使われていない'
  );
});

test('AdminPanel.js: validateCompleteUrl は validateBtn 参照を nil-safe で扱う', () => {
  const fnMatch = ADMIN_JS.match(
    /function\s+validateCompleteUrl\s*\([\s\S]*?\n\s{2}\}/
  );
  assert.ok(fnMatch, 'validateCompleteUrl 関数本体が見つからない');
  const body = fnMatch[0];
  // hidden 化された legacy ボタンが無いとき (null) にも安全に動くこと
  assert.match(
    body,
    /if\s*\(\s*validateBtn\s*\)/,
    'validateBtn の nil-check が無い (button hidden 化後の安全策が必要)'
  );
});

test('AdminPanel.js: getCurrentDataSourceField が URL→dropdown の順で参照する', () => {
  const fnMatch = ADMIN_JS.match(
    /function\s+getCurrentDataSourceField\s*\([\s\S]*?\n\s{2}\}/
  );
  assert.ok(fnMatch, 'getCurrentDataSourceField 関数本体が見つからない');
  const body = fnMatch[0];
  // form-url-input を先にチェック、次に form-select
  const urlIdx = body.indexOf("'form-url-input'");
  const selIdx = body.indexOf("'form-select'");
  assert.ok(urlIdx >= 0 && selIdx >= 0, '両 dataset 参照が無い');
  assert.ok(urlIdx < selIdx, 'URL 方式が先に参照されていない (priority が逆)');
});

test('AdminPanel.js: scheduleAutoSave が AI 列判定の完了点から呼ばれている', () => {
  // connectToDataSource と manualColumnAnalysis の両方で
  // 「fillColumnDropdowns 後の scheduleAutoSave」が必要
  // (programmatic な select.value 更新は change event を発火しないため)
  const occurrences = (ADMIN_JS.match(/scheduleAutoSave\s*\(\s*\)/g) || []).length;
  assert.ok(
    occurrences >= 3,
    `scheduleAutoSave の呼び出しが少ない (${occurrences} 件)。AI 判定完了点で漏れている可能性`
  );
});
