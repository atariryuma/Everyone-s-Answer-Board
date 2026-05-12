# 可視化モード仕様書 — M1 ポジショニング / M2 マトリクス

**版**: 0.1（ドラフト）
**対象スコープ**: M1（1D 数直線）と M2（2×2 散布図）のみ。M3〜M5（ワードクラウド、ランキング、Q&A板）は本書の対象外。
**前提読者**: 本リポジトリのアーキテクチャ（`CLAUDE.md`）を理解している実装者。

---

## 1. 目的・スコープ

道徳科の電子黒板授業で「Forms に集まった意見を即時に視覚的集約して、揺らぎや分布を全員で議論する」体験を実現する。SKYMENU Cloud のポジショニング機能の代替を、本アプリのインフラに最小コストで載せる。

**追加するもの**:
- 既存「掲示板モード」と並ぶ 2 つの新表示モード（M1, M2）
- Forms 「線形尺度」列を numericScale として自動検出
- 同一児童の再投稿を許可した時の「揺らぎ」表示
- 電子黒板向けの専用ビュー（`?display=projector`）

**追加しないもの**（後続フェーズ）:
- ワードクラウド、ランキング、Q&A 板
- 完全な認可システム拡張（既存の `getPublishedSheetData` の認可をそのまま使う）
- データの永続的スナップショット（軌跡記録）— まずは現状分布の表示のみ

---

## 2. ユーザーストーリー

### US-1（教師）
6 年道徳の授業前、Forms で「ポスターをそのまま使う(1)⇔自分で直す(5)」の線形尺度と理由記述を作成。授業中に児童が回答すると、電子黒板にビーズワーム散布図が即時更新される。点をタップすると匿名コメントが表示される。

### US-2（教師）
児童に再投稿を許可すると、議論前後で意見が動いた児童の「揺らぎ」が円の大きさで可視化される。揺らぎの大きい児童を意図的指名する材料にする。

### US-3（教師）
「効率（X軸: 1〜5）× 倫理（Y軸: 1〜5）」の 2 軸 Forms を作成。授業中、児童の立場が 4 象限のどこに集まるか、議論の経過とともに変化するか、電子黒板で全員と共有する。

### US-4（児童）
普段使っている Google Forms に答えるだけ。新しいアプリの操作を覚える必要はない。

---

## 3. 機能要件

### F-1 表示モードの切替
- `config.displaySettings.boardMode` を新規追加。値は `'board' | 'numberline' | 'matrix'`。
- 既定値: `'board'`（既存挙動と完全互換）。
- 管理画面で教師が選択。プレビューは保存後の View ページで確認。

### F-2 数値列の選択（M1: 1 列、M2: 2 列）
- `config.columnMapping` に以下を追加：
  - `numericX`: 列インデックス（必須）
  - `numericY`: 列インデックス（M2 のみ必須、M1 では無視）
  - `xAxisLabels`: `{ min: '左端ラベル', max: '右端ラベル' }`（任意、未指定なら列ヘッダから推定）
  - `yAxisLabels`: 同上（M2 のみ）
- 線形尺度（Forms の「リニアスケール」）列は **自動検出**（§7 参照）。教師は確認・上書きするだけ。

### F-3 ビーズワーム表示（M1）
- 横軸: numericX の値（1〜5、または検出された min/max）。
- 縦方向: 同位置に重ならないよう d3-beeswarm で配置。
- 各点 = 1 投稿。色はクラス別（class 列がある場合）。
- 点ホバー/タップ:
  - `showNames=false` のとき匿名 ID + コメント（reason 列）。
  - `showNames=true` のとき名前 + コメント。
- 軸下に分布ヒストグラム（補助情報）。

### F-4 散布図表示（M2）
- X軸: numericX、Y軸: numericY。
- 4 象限の背景に薄い色分け（高高/高低/低高/低低）。
- 象限ラベル（任意）: `config.matrixQuadrantLabels = { hh, hl, lh, ll }`。
- 重なり回避にジッタ（±0.15）を加える。
- 点ホバー挙動は M1 と同じ。

### F-5 「揺らぎ」表示（M1, M2 共通、再投稿許可時のみ）
- `config.allowResubmit = true` のとき有効。
- 同一児童の最新投稿のみ表示。それ以前の投稿は薄い色でゴースト表示。
- 児童識別子: **emailHash**（メールアドレスの SHA-1 先頭 8 文字。匿名性確保のためメアドそのものは表示しない）。
- 教師ダッシュボード（adminMode）では、変化量の大きい順に上位 5 名を「気になる児童」として表示。

### F-6 電子黒板向け表示
- URL に `?display=projector` を付けるとプロジェクタ最適化レイアウト。
- 既定: 暗背景、フォント大、操作 UI 非表示、ジョインコード+QR を画面端常時表示。
- 教師端末で `view` ページ通常表示、電子黒板で `projector` 表示を別タブで開き、同じ Forms に向かう。

### F-7 リアルタイム更新
- 既存のポーリング基盤を流用。`POLLING_INTERVAL_MS` 既定 5 秒。
- 新規データ到着時、点が**新規=フェードイン**、消失=フェードアウト、移動=トランジション。
- 電子黒板でも自動更新（ユーザー操作不要）。

### F-8 表示モード切替の即時性
- 同一ページ内で `board` / `numberline` / `matrix` を切替可能（教師権限のみ）。
- 切替は再ロードなし、JavaScript で DOM 入れ替え。

### F-9 可視化モード連動テンプレート
- 管理パネル「📝 Googleフォーム選択」セクションに種別 select を配置:
  - `掲示板` (board, 既定): クラス/名前/回答(MC)/理由 — 既存挙動
  - `数直線 (M1)` (numberline): クラス/名前/**立場(ScaleItem 1-5)**/理由
  - `マトリクス (M2)` (matrix): クラス/名前/**X軸 ScaleItem 1-5**/**Y軸 ScaleItem 1-5**/理由
- 「✨ テンプレート作成」クリックで `createTemplateForm(templateType)` を呼び出し、
  Forms / Spreadsheet / 専用フォルダ「みんなの回答ボード」を一気に生成。
- numberline/matrix を選んだ場合は、テンプレート作成成功時にフロントが
  `board-mode-select` を対応モード（'numberline' / 'matrix'）に**自動セット**して
  change イベントを dispatch、列マッピング UI を即時展開する。
- 教師の作業は: テンプレート作成 → クラス名と軸ラベル編集 → 検証して接続 → 公開 の 4 ステップ。
- Why: §7 の `detectNumericScaleColumns` は実データから検出するため、空の新規フォームでは
  検出に失敗する（応答 0 件）。テンプレート時点で ScaleItem を仕込んでおくことで、
  教師が手動で線形尺度フィールドを追加する手間を省く。

---

## 4. 非機能要件

| 項目 | 要件 |
|---|---|
| 描画性能 | 200 点まで 60fps、500 点まで実用速度（d3 SVG） |
| 同期遅延 | Forms 投稿から表示反映まで最大 10 秒（既存ポーリングの 2 サイクル以内） |
| データ互換性 | 既存ボード（boardMode=board）に影響ゼロ。config に新フィールドが無くても動く |
| アクセシビリティ | 色だけで識別しない（形・ラベル併用）。WCAG AA コントラスト |
| プライバシー | `showNames=false` のときコメントの内容に名前が含まれていても**自動マスクはしない**（教師が事前に注意喚起する運用） |
| ブラウザ | Chromium 系（GIGA スクール標準）。Safari, Firefox は努力目標 |

---

## 5. アーキテクチャ概要

```
[Forms] ──submit──> [Spreadsheet]
                          │
                          ▼
                  ┌───────────────────┐
                  │ DataService       │ ← 既存のまま
                  │ fetchSpreadsheetData
                  └─────────┬─────────┘
                            │ items[]
                            ▼
              ┌─────────────────────────┐
              │ DataApis.               │ ← 既存 + 軽微な拡張
              │ getPublishedSheetData() │
              │  → numericX/Y 列も値を含めて返す
              └─────────┬───────────────┘
                        │
                  google.script.run
                        │
                        ▼
              ┌─────────────────────────┐
              │ page.js.html            │ ← ★主要な変更点
              │  loadSheetData          │
              │   └─ dispatchRender ★新規
              │       ├─ renderBoard (既存)
              │       ├─ renderNumberLine ★新規 (M1)
              │       └─ renderMatrix    ★新規 (M2)
              └─────────────────────────┘
```

### 5.1 設計原則
- **既存 API は変えない**。`getPublishedSheetData` のレスポンスにフィールド追加するだけ（後方互換）。
- **フロントの分岐は 1 箇所**。`renderBoard()` のエントリで `boardMode` を見て dispatcher に渡す。
- **d3 は CDN ロード**。GAS のファイル増を最小化、必要なモードでのみロード（dynamic import パターン）。

---

## 6. データモデル変更

### 6.1 `config` JSON への追加フィールド

```javascript
{
  // 既存フィールドは変更なし
  displaySettings: {
    showNames: false,
    showReactions: true,
    theme: 'default',
    pageSize: 20,
    boardMode: 'numberline'   // ★新規 'board' | 'numberline' | 'matrix'
  },
  columnMapping: {
    answer: 4,
    reason: 5,
    class: 2,
    name: 3,
    numericX: 6,              // ★新規（M1/M2 必須）
    numericY: 7               // ★新規（M2 のみ）
  },
  xAxisLabels: {              // ★新規（任意）
    min: 'そのまま使う',
    max: '自分で直す'
  },
  yAxisLabels: { min: '...', max: '...' },  // ★M2 のみ
  matrixQuadrantLabels: {     // ★新規（任意、M2 のみ）
    hh: '理想', hl: '時短重視',
    lh: 'じっくり', ll: '停滞'
  },
  allowResubmit: false        // ★新規（任意、揺らぎ表示の有効化）
}
```

### 6.2 既存ユーザーへの影響
- 新フィールドはすべて optional。`config` JSON にこれらが無い場合、`boardMode = 'board'` として既存動作。
- マイグレーション不要（initスクリプト走らせる必要なし）。

### 6.3 サニタイズ拡張
- `src/ConfigService.js:sanitizeMapping()` に `numericX`, `numericY` を allowlist 追加。
- `src/ConfigService.js:sanitizeDisplaySettings()` に `boardMode` を allowlist 追加（値は enum チェック）。

---

## 7. 線形尺度列の自動検出

### 7.1 検出ロジック（`ColumnMappingService.js` に追加）

```javascript
// 新規 helper
function detectNumericScaleColumns(headers, sampleData) {
  const candidates = [];
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];
    const samples = sampleData.map(row => row[i]).filter(v => v !== '' && v != null);
    if (samples.length < 3) continue;
    // 全部数字かつ 1〜10 程度の整数なら候補
    const allInt = samples.every(v => Number.isInteger(Number(v)));
    if (!allInt) continue;
    const nums = samples.map(Number);
    const min = Math.min(...nums);
    const max = Math.max(...nums);
    if (min < 0 || max > 10) continue;
    if (max - min < 1) continue;  // 全部同じは除外
    candidates.push({
      index: i,
      header,
      min,
      max,
      confidence: samples.length >= 10 ? 95 : 80
    });
  }
  return candidates;
}
```

### 7.2 軸ラベルの抽出
Forms の「リニアスケール」では、ヘッダ列の値に**ラベル**が入らない（数字だけ）。Forms 設定の min/max ラベルはスプレッドシートに反映されない。
**現実解**: 教師が管理画面で手動入力。デフォルトは空文字。

### 7.3 検出結果の UI 提示
管理画面の「列マッピング」セクションに自動検出結果を表示し、教師が確認・上書き：

```
線形尺度の列を検出しました:
  □ 列F「ポスターをどう扱う？」 (1〜5)  [X軸として使う]
  □ 列G「効率を重視するか」 (1〜5)      [Y軸として使う（M2のみ）]
左端ラベル: [そのまま使う]
右端ラベル: [自分で直す]
```

### 7.4 テンプレート経由フォームの取り扱い
§F-9 のテンプレート（numberline / matrix）で作成された新規フォームは、応答が 0 件の段階では `detectNumericScaleColumns` の候補が空になる（実データに依存するため）。
これを補うため、フロント側でテンプレート作成成功直後に `board-mode-select` を `templateType` に合わせて自動セットする。教師はその後の「検証して接続」で列マッピング UI を開き、ヘッダ名（"立場" / "X軸" / "Y軸"）から手動で `numericX/Y` を確定させる。

応答が 1 件でも届けば `detectNumericScaleColumns` が候補を提示し始める（`confidence` は応答数に比例して上昇）。設計上、テンプレート定数 `TEMPLATE_SCALE_MIN/MAX`（`src/DataApis.js`）と検出ロジックのレンジ判定（0-10）は**意図的に重ねている** — テンプレートで作った 1-5 スケールが auto モードで必ず検出されることを保証するため。

---

## 8. バックエンド変更

### 8.1 `src/DataService.js`

#### 変更箇所 1: `processRawDataBatch()` (lines 401–473)
**現状**: 各 item に answer/reason/class/name のみセット。
**変更**: numericX/Y 列の値も `item.numericX`, `item.numericY` として含める。

```javascript
// fieldIndices 計算箇所に追加
fieldIndices.numericX = resolveColumnIndex(headers, 'numericX', config.columnMapping);
fieldIndices.numericY = resolveColumnIndex(headers, 'numericY', config.columnMapping);

// item 生成箇所に追加
const numericX = fieldIndices.numericX >= 0 ? Number(row[fieldIndices.numericX]) : null;
const numericY = fieldIndices.numericY >= 0 ? Number(row[fieldIndices.numericY]) : null;

// item に追加
item.numericX = Number.isFinite(numericX) ? numericX : null;
item.numericY = Number.isFinite(numericY) ? numericY : null;
```

#### 変更箇所 2: `shouldIncludeRow()` (line 500)
**変更**: 数直線/マトリクスモードのとき、numericX が null の行は除外する判定を追加。
ただし、これは表示側でフィルタしてもよいので、サーバ側ではフィルタしない（柔軟性のため）。

### 8.2 `src/DataApis.js`

#### 変更箇所 1: `buildSafePublishedDataResult()` (lines 442–466)
**変更**: numericX/Y はそのまま wire に乗せる（個人情報ではない）。
匿名性確保のため、`showNames=false` のときも emailHash を送る（個人を識別不能だが同一児童の追跡は可能）。

```javascript
// 追加: emailHash 生成
function emailHash(email) {
  if (!email) return null;
  return Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_1,
    email
  ).slice(0, 4).map(b => (b & 0xff).toString(16).padStart(2, '0')).join('');
}

// items map 処理に追加
if (!showNames) {
  delete item.email;
  delete item.name;
  // ただし、再投稿追跡のために emailHash は残す
  item.emailHash = emailHash(originalEmail);
}
```

#### 変更箇所 2: `getPublishedSheetData()` の戻り値 (lines 555〜)
**変更**: `displaySettings` に `boardMode` を含め、列マッピングのラベル情報も追加。

```javascript
return {
  success: true,
  data: safeItems,
  header: questionText,
  sheetName: ...,
  displaySettings: {
    ...existing,
    boardMode: config.displaySettings.boardMode || 'board'
  },
  axisConfig: {                              // ★新規
    xAxisLabels: config.xAxisLabels || null,
    yAxisLabels: config.yAxisLabels || null,
    matrixQuadrantLabels: config.matrixQuadrantLabels || null,
    allowResubmit: config.allowResubmit || false
  }
};
```

### 8.3 `src/ConfigService.js`

#### 変更箇所: `sanitizeDisplaySettings()` (line 266) / `sanitizeMapping()` (line 280)
**変更**: `boardMode`, `numericX`, `numericY` の allowlist 追加。

```javascript
const VALID_BOARD_MODES = ['board', 'numberline', 'matrix'];

function sanitizeDisplaySettings(s) {
  const out = { /* 既存処理 */ };
  if (VALID_BOARD_MODES.includes(s.boardMode)) {
    out.boardMode = s.boardMode;
  }
  return out;
}

function sanitizeMapping(m) {
  const fields = ['answer', 'reason', 'class', 'name', 'timestamp', 'email',
                  'numericX', 'numericY'];  // ★追加
  const out = {};
  for (const f of fields) {
    if (typeof m[f] === 'number' && m[f] >= 0 && m[f] < 100) out[f] = m[f];
  }
  return out;
}
```

### 8.4 `src/ColumnMappingService.js`
- `detectNumericScaleColumns()` を新規追加（§7.1）。
- `getColumnAnalysis()` (DataApis.js:821) の戻り値に `numericScaleCandidates: [...]` を追加。

### 8.5 `src/main.js`
- 変更不要（既存の `view` モードの routing でそのまま動く）。
- `?display=projector` の判定はフロントで行う（テンプレートは同じ Page.html）。

---

## 9. フロントエンド変更

### 9.1 `src/page.js.html`

#### 変更箇所 1: `renderBoard()` (line 2091)
**現状**: 単一の card grid 描画。
**変更**: dispatcher 化。

```javascript
async renderBoard(isLayoutChange = false, isInitialLoad = false, oldRows = []) {
  const mode = this.state.displaySettings?.boardMode || 'board';
  switch (mode) {
    case 'numberline': return this.renderNumberLine(isInitialLoad, oldRows);
    case 'matrix':     return this.renderMatrix(isInitialLoad, oldRows);
    case 'board':
    default:           return this.renderBoardCards(isLayoutChange, isInitialLoad, oldRows);
  }
}
// 既存 renderBoard 本体を renderBoardCards にリネーム
```

#### 変更箇所 2: 新規 `renderNumberLine()` / `renderMatrix()`
別ファイル（`src/page.viz.js.html`）に切り出し、Page.html の include に追加。

```javascript
renderNumberLine(isInitialLoad, oldRows) {
  this.ensureD3Loaded();                     // CDN から動的ロード
  const container = document.getElementById('answers');
  container.innerHTML = '<svg id="vizSvg"></svg>';
  // d3 で beeswarm レンダリング
  // ...
}
```

#### 変更箇所 3: `loadSheetData()` (line 1872)
**変更**: 戻り値の `axisConfig` を `this.state.axisConfig` に保存。

#### 変更箇所 4: ポーリングと差分更新
**現状**: `renderDirectly()` で existingMap を使って DOM を差分更新。
**変更**: viz モードでは「全件再描画」を採用。d3 のデータ join パターンで差分は自然に処理される（enter / update / exit）。

#### 変更箇所 5: モード切替 UI（教師用）
- `#main-header` に「表示モード」セレクトボックスを追加。
- 教師権限（`isOwnBoard || isAdmin`）のみ表示。
- 選択時、`saveConfig` で `boardMode` を保存し、ローカル状態も即時更新して再描画。

### 9.2 `src/Page.html`
**変更**: include 追加（順序固定なので注意）。

```html
<?!= include('UnifiedStyles.css'); ?>
<?!= include('page.css'); ?>
<?!= include('page.viz.css'); ?>    <!-- ★新規 -->
...
<?!= include('page.js'); ?>
<?!= include('page.viz.js'); ?>     <!-- ★新規 -->
```

**変更**: `#answers` 要素は viz モードでも流用（中身を SVG に差し替える）。

### 9.3 新規ファイル: `src/page.viz.js.html`
- d3 CDN ロード（`https://d3js.org/d3.v7.min.js` + `https://unpkg.com/d3-beeswarm@1`）
- `StudyQuestApp.prototype.renderNumberLine = function(...) { ... }`
- `StudyQuestApp.prototype.renderMatrix = function(...) { ... }`
- ホバー時のツールチップ DOM 生成
- 揺らぎ表示の集計ロジック（emailHash でグルーピング → 最新のみ表示、軌跡は薄色）

### 9.4 新規ファイル: `src/page.viz.css.html`
- 軸、点、象限背景の CSS
- プロジェクタモード時のオーバーライド（`body.projector-mode` をルートに付与）

### 9.5 プロジェクタモード
**実装**: page.js.html 初期化時に `?display=projector` を見て `body.classList.add('projector-mode')`。
CSS で：
- 背景 #1a1a2e、テキスト #fff
- フォントサイズ 1.5× スケール
- ヘッダ非表示
- 画面端に常設で「ボード URL（QR）」表示

---

## 10. UI/UX モック

### 10.1 M1 数直線（教師端末）
```
┌──────────────────────────────────────────────────────────────┐
│ ⓘ ポスターを明日までにどう扱う？             [掲示板|数直線|散布]│
├──────────────────────────────────────────────────────────────┤
│                                                              │
│   そのまま使う                                  自分で直す      │
│       1        2        3        4        5                  │
│   ●●●  ●●●●●●  ●●●●●●●●●●●  ●●●●●●●●  ●●●●●●●●●●         │
│       ●●●●●  ●●●●●●●●  ●●●●●●●●●●●●●  ●●●●●●●  ●●●●●         │
│         ●●●●     ●●●●●●●     ●●●●●●●●     ●●●●●     ●●●●●●●  │
│                                                              │
│  分布:  3   5     12    8     7                              │
│  ────────────────────────────────────                        │
│  ホバーで匿名コメント表示                                       │
│  最終更新: 14:32 (5 秒ごとに自動更新)  接続中: 28 人           │
└──────────────────────────────────────────────────────────────┘
```

### 10.2 M2 マトリクス
```
┌────────────────────────────────────────────────────────┐
│ 倫理 ↑                                                  │
│  高 │  じっくり        │     理想                       │
│     │   ⚫ ⚫            │   ⚫ ⚫ ⚫ ⚫                      │
│     │     ⚫            │   ⚫ ⚫                          │
│     │                  │                                │
│     ├──────────────────┼─────────────────► 効率         │
│     │   停滞           │    時短重視                    │
│     │    ⚫             │   ⚫ ⚫ ⚫                        │
│     │                  │       ⚫                        │
│  低 │                  │                                │
└────────────────────────────────────────────────────────┘
```

### 10.3 プロジェクタモード（暗背景 + QR）
- 上部にタイトル大表示
- 中央に散布図/数直線
- 右下に QR コード + 「同じ Forms で回答 →」テキスト
- 右上に接続人数バッジ

---

## 11. 実装フェーズと工数見積

### フェーズ 1: M1 基本実装（**3 日**）
1. config スキーマ拡張（`ConfigService.js` sanitize 追加）— 2h
2. `processRawDataBatch()` に numericX 取込 — 2h
3. `buildSafePublishedDataResult()` で emailHash + axisConfig 返却 — 3h
4. `renderBoard` dispatcher 化 — 2h
5. `page.viz.js.html` で d3 ビーズワーム描画 — 6h
6. ホバーツールチップ + 匿名コメント表示 — 3h
7. プロジェクタモード CSS — 3h
8. テスト追加 — 3h

### フェーズ 2: M2 マトリクス（**1.5 日**）
1. numericY 追加（M1 拡張で大半流用） — 2h
2. 散布図描画 + 象限背景 + ラベル — 4h
3. 4 象限ラベル設定 UI（管理画面） — 3h
4. テスト追加 — 2h

### フェーズ 3: 揺らぎ表示（**1 日**）
1. emailHash グルーピング + 軌跡表示 — 4h
2. 教師ダッシュボードの「気になる児童」リスト — 3h
3. 再投稿許可フラグの UI（管理画面） — 1h

### フェーズ 4: 自動検出と管理画面 UX（**1 日**）
1. `detectNumericScaleColumns()` 実装 — 2h
2. 管理画面の検出結果表示 + 上書き UI — 4h
3. 軸ラベル入力 UI — 2h

**合計: 6.5 日**（テストとレビュー含めると 8 日見ておくのが現実的）

---

## 12. 既存資産の流用一覧

| 既存 | 流用先 | 修正の有無 |
|---|---|---|
| `getPublishedSheetData` 認可フロー | viz モードも同じエンドポイント | 無 |
| `fetchSpreadsheetData` のバッチ/リトライ | そのまま | 無 |
| `columnMapping` の resolveColumnIndex | numericX/Y 追加だけ | 微 |
| ポーリング基盤（5 秒間隔） | viz モードでも同じ | 無 |
| `classFilter` / `sortOrder` | viz モードでも適用可（クラス別色分けに使える） | 無 |
| `displaySettings.showNames` | 匿名性制御として流用 | 無 |
| Page.html / include 機構 | viz 用 include を追加 | 微 |
| `getColumnAnalysis` の管理画面導線 | numericScale 候補を追加 | 微 |

---

## 13. テスト計画

### 13.1 単体テスト（`tests/` に追加）
- `tests/column.numericScale.test.cjs` — `detectNumericScaleColumns()` の検出精度
- `tests/data.service.numeric.test.cjs` — `processRawDataBatch` が numericX/Y を含めるか
- `tests/config.boardMode.test.cjs` — sanitize の allowlist
- `tests/data.apis.axisConfig.test.cjs` — `buildSafePublishedDataResult` が emailHash を返すか

### 13.2 統合確認（手動）
| ケース | 期待動作 |
|---|---|
| 既存ボードを開く | M1/M2 追加前と完全同じ表示 |
| 新規 Forms（線形尺度 1 問）→ 管理画面で M1 選択 | ビーズワームが表示される |
| 線形尺度 2 問 → M2 選択 | 散布図が表示される |
| 再投稿許可 ON で同一児童が 2 回投稿 | 最新のみ目立つ、前回は薄色 |
| `?display=projector` で URL を開く | 暗背景・大文字・QR表示 |
| 30 件投稿してリアルタイム更新 | 5 秒以内に点が追加表示される |

### 13.3 パフォーマンス確認
- 200 件で表示遅延 1 秒以下を目標
- Chrome DevTools のパフォーマンスタブで FPS 計測

---

## 14. ロールアウト戦略

1. **feature flag**: 当面は `config.displaySettings.boardMode` が `'numberline' | 'matrix'` の場合のみ有効（既存ユーザーには無影響）。
2. **対象校への先行展開**: まず那覇市の中先生のクラス（道徳 5/14 実施予定の指導案）で実証。
3. **本番デプロイ**: `npm run deploy:prod`（既存の URL 維持デプロイ）。再認可は不要（新スコープなし）。
4. **問題発生時のロールバック**: 教師が管理画面で `boardMode='board'` に戻すだけ。バックエンド変更も後方互換なので緊急時の revert は不要。

---

## 15. オープン課題

1. **数値列のラベル抽出**: Forms 「リニアスケール」の min/max ラベルはスプレッドシートに自動転記されない。当面は教師の手入力とするが、Forms API で取得できないか調査の余地あり。
2. **同一児童の識別精度**: メアドが Forms で収集できていない場合、emailHash は使えない。Forms 設定で「メールアドレスを収集」を ON にする運用ルールを明文化する。
3. **`showNames=false` での揺らぎ表示**: emailHash でも同一児童の動きが見えるのは「準匿名」。クラス内なら推測可能性あり。教師に運用注意を伝える。
4. **モバイル端末での表示**: 児童端末（タブレット）で viz モードを見たときのレイアウト。プロジェクタモードと教師ビューの中間が必要かは要検討。
5. **クラスフィルタとの併用**: 学年共有スプレッドシートで「6 年 1 組のみ」フィルタした散布図の有効性は授業実践で確認する。
6. **既存テスト**: `tests/data.service.test.cjs` 等の既存テストが numericX 追加で壊れないか、CI で確認。

---

## 16. 参考文献

- [SKYMENU ポジショニング機能](https://www.sky-school-ict.net/class/report/200930/)
- [d3-beeswarm](https://github.com/Kcnarf/d3-beeswarm)
- [Mentimeter vs Slido 比較](https://www.classpoint.io/blog/slido-vs-mentimeter-vs-classpoint)
- 本プロジェクトの `CLAUDE.md` および `src/` 配下の既存実装
- `docs/AGENT_explorer_2026-05-12.md`（実装把握の生ログ。本書の §5–§9 はこの調査結果に基づく）
