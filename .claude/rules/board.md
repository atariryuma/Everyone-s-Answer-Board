---
paths:
  - "src/validators.js"
  - "src/ColumnMappingService.js"
  - "src/ConfigService.js"
  - "src/DataService.js"
  - "src/DataApis.js"
  - "src/SystemController.js"
  - "src/AdminApis.js"
  - "src/ReactionService.js"
  - "src/page.viz.*"
  - "tests/admin.*"
  - "tests/column.*"
  - "tests/config.*"
  - "tests/data.*"
  - "tests/system.*"
  - "tests/reaction*"
---

# ボード (データの置き場・表示モード・公開)

## Data Store

先生のスプレッドシートが DB。`users` シート (先生 1 人分) の `configJson` 列にボードの設定、`lessons` / `lesson_responses` に授業とアーカイブ、授業モードの回答シートは「授業名｜番号 フェーズ名」で同じスプレッドシートに入る (see `src/SystemController.js` `USERS_SHEET_HEADERS`)。

アプリが作ったフォーム (`createTemplateForm`) は回答シートの列の並びが決まっているので、作った時点の割り当て
(`templateColumnMapping_`) を使い、推定しない (回答 0 件では尺度の列を数値と見分けられないため)。
回答シートには種類を developer metadata で書き添え、`getColumnAnalysis` が見出しの並びを確かめて使う。

## Visualization Modes

`config.displaySettings.boardMode` で表示モードを切替。**許可値・ラベル・必須列の唯一の定義は
`src/validators.js` の `BOARD_MODES` 配列**（select の option / クライアント allowlist / ラベル表は
すべてここから導出される。増やすときは配列に 1 行足す）:

| key | ラベル | 必須マッピング | テンプレ生成 |
| --- | ------ | -------------- | ------------ |
| `auto` | 自動判定（推奨・既定） | — (サーバが解決) | — |
| `board` | 掲示板（カード一覧） | `answer` | ✓ |
| `numberline` | 数直線（d3 ビーズワーム） | `numericX` | ✓ |
| `matrix` | マトリクス（2軸散布図） | `numericX` + `numericY` | ✓ |
| `wordcloud` | ワードランキング | `answer` | — |
| `pie` | 円グラフ | `answer` | ✓ |

管理パネル「📝 Googleフォーム選択」の種別 select から、テンプレ生成対応モード
(`TEMPLATE_BOARD_MODES`) の Forms 構造を 1 クリックで生成可能（`createTemplateForm(templateType)`）。
生成時は `boardMode` も自動セットされる。

M1 (numberline) / M2 (matrix) の設計背景: [docs/SPEC_visualization_modes.md](../../docs/SPEC_visualization_modes.md)
（2026-05 の初期設計書。wordcloud / pie はその後の追加なので同書には載っていない）。

## Board Publish Lifecycle (single source of truth)

`config.isPublished` / `config.publishedAt` を書き換えてよいのは以下 **3 関数のみ**。直接 saveUserConfig_ で書き換えてはいけない（DRY 違反 + テスト不可）：

| 関数 | 場所 | 役割 |
| ---- | ---- | ---- |
| `publishApp(publishConfig)` | `SystemController.js` | 初回公開 / 設定差し替えを伴う公開（allowlist + etag 検証） |
| `republishMyBoard(options?)` | `AdminApis.js` | 所有者が自分のボードを再び公開する（state のみ） |
| `unpublishBoard(targetUserId?, options?)` | `AdminApis.js` | 所有者 or 管理者がボードの公開を終了する |

- `publishApp` accepts only allowlisted fields with `etag` conflict detection - maintain both

**統一不変条件**（`__applyPublishStateChange_` で集約）：

- `publishedAt` は「現在の状態の発生時刻」（公開中 = now / 非公開 = null）
- `lastAccessedAt` は変更時に touch
- `options.sourceEtag` を渡せば etag conflict 検出が有効化される（楽観ロック）
- 認可: 自分のボード or 管理者（`requireAdmin: true` 指定時は管理者のみ）
- 戻り値: `{ success, message, userId, isPublished, boardPublished, publishedAt, etag, redirectUrl }` で統一

**UI 用語**：

- アクション: 「公開する」「再び公開する」「公開を終了する」
- 状態: 「公開中」「非公開」

詳細仕様とテスト: `tests/admin.publishLifecycle.test.cjs`
