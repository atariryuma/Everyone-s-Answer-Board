# 開発・保守ツール一覧（CLI ベース）

このプロジェクトの開発・保守は全て CLI から完結するよう構成されています。
このドキュメントは「何ができて、いつ使うか」のマップです。

---

## 1. 観測：配置の状態とログ

### 1-1. 実行ログ

先生ごとの配置は既定の GCP project で動くので、Cloud Logging を CLI から読めない。
Apps Script エディタ (スプレッドシート → 拡張機能 → Apps Script) の「実行数」で見る。
フロントエンドのエラー (`window.error` / `unhandledrejection` / `ErrorHandler.logError`) も
`reportClientError` doPost アクション経由で同じ場所に出る (`[client/error]` プレフィックス)。

### 1-2. Admin API（配置の状態確認・操作）

送り先は `scripts/single-tenants.json` の配置。`--tenant <名前>` で選ぶ (省略時は検証用を除く最初の配置)。
API キーは配置ごとに `single-tenants.json` の `apiKey` (初回に `setupApiKey` で決める。
`scripts/migrate-to-single.js` / `scripts/verify/e2e-single.js` が初回に自動で決める)。

#### 基本（システム状態）

```bash
npm run api -- systemDiagnosis        # Core props / DB / Deploy の健康診断
npm run api -- getUsers               # 登録 (先生 1 人)
npm run api -- getAppStatus           # アプリ有効/無効状態
npm run api -- enableApp              # アプリを有効化
npm run api -- disableApp             # 緊急停止（アプリ無効化）
npm run api -- getLogs --limit 20     # セキュリティログ
npm run api -- perfMetrics            # パフォーマンス指標
npm run api -- listProperties         # Script Properties（認証情報マスク済み）
npm run api -- cacheReset             # 全キャッシュクリア
```

#### ボードの設定・授業

```bash
npm run api -- getUserConfig --userId <uuid>
npm run api -- exportConfigs --output backup.json          # 設定の dump（ファイル）
npm run api -- setUserConfig --userId <uuid> \
  --patch '{"displaySettings":{"boardMode":"numberline"}}' # 部分 patch、deep merge、保護フィールド自動除外
npm run api -- previewBoard --userId <uuid>                # viewer 視点のサンプル
npm run api -- unpublishBoard / republishMyBoard           # 公開を終了する / 再び公開する

npm run api -- listMyForms                                  # 自分の Drive 内 Form を最大 30 件
npm run api -- connectForm --userId <uuid> --formUrl <URL>  # 既存 Form をボードにつなぐ
npm run api -- createForm --userId <uuid> --templateType board|numberline|matrix|pie

npm run api -- lesson.list --userId <uuid>
npm run api -- lesson.liveAnswers --userId <uuid> --lessonId <id>
npm run api -- lesson.advance --userId <uuid> --lessonId <id> --direction next
npm run api -- lesson.backfillReactions --userId <uuid> --lessonId <id> --dryRun  # 終わった授業の振り返りに
                                                            # 元の回答シートのリアクションの数・ハイライトを足す (--dryRun で数えるだけ)
```

JSON 引数: `--patch` / `--filter` / `--json` は自動パース、`--file path.json` で params を読み込み、
`--output path` で結果を保存。保護フィールド (patch から自動除外): `userId`, `userEmail`, `googleId`, `createdAt`, `etag`

#### 引っ越し (共有の本体 → 先生ごとの配置)

`scripts/migrate-to-single.js --name <配置> --userId <元の userId> [--dry-run]`。元の本体の admin API
(`scripts/config.json`) から授業とボードの設定を読み、`migrate.importLesson` / `migrate.importConfig` で取り込む。
別のスプレッドシートにある授業モードの回答は `migrate.importLessonAnswers` (学級の付け替え・時刻のずらし可)。

---

## 2. 品質ゲート：コミット前後の検証

### 2-1. ユニットテスト

```bash
npm test                              # 全 580+ テスト（VM sandbox）
npm test -- tests/main.doPost.test.cjs   # 個別ファイル
npm run test:coverage                 # カバレッジ表示
```

**仕組み**：各テストは `vm.createContext()` で GAS グローバル（SpreadsheetApp 等）を
stub し、ソースファイルを直接読み込んで検証する。フレームワーク非依存（node:test 標準）。

### 2-2. ESLint（バックエンド JS のみ）

```bash
npm run lint:eslint                   # src/*.js を ESLint で検査
npx eslint --fix src/*.js             # 自動修正可能なら修正
```

設定は `eslint.config.js`。既存負債を吸収するため大半 warn 降格、
バグ予防系（no-dupe-keys, no-unreachable, use-isnan 等）のみ error。

### 2-3. カスタム静的解析（CLAUDE.md ルール）

```bash
npm run lint                          # CLAUDE.md ルール違反検出（全 36 ファイル）
npm run lint:errors                   # error severity のみ（CI 用）
node scripts/lint.js src/main.js      # 特定ファイル
```

検出するアンチパターン：

- `setTimeout` / `setInterval` （GAS V8 では動作しない）
- `Session.getEffectiveUser()` （権限昇格リスク）
- `innerHTML = <変数>` （XSS リスク）
- `for {...getValue()...}` （N+1 性能アンチパターン）
- `PropertiesService...getProperty()` 直接呼び出し（キャッシュ不使用）
- HTML テンプレート JS の top-level 副作用

### 2-4. 統合チェック

```bash
npm run check                         # test + systemDiagnosis（配置への疎通含む）
```

### 2-5. 画面の機能を端から端まで (検証用の配置)

```bash
npm run verify:e2e -- <検証用の名前>  # 掲示板モード・授業モード・児童の操作・後片付けの 64 項目
```

検証用の配置 (`single-tenants.json` に `extraFiles: ["scripts/verify/DevInvoke.js"]`) でだけ流せる。
admin API の `dev.invoke` で画面が `google.script.run` で呼ぶ処理を名前で呼び、児童の操作は `as` で
架空の児童として呼ぶ。フォームや授業を作るので本番の先生には流さない (作ったフォームは Drive に残る)。

---

## 3. 配置

```bash
npm run deploy                                    # 全員の配置を更新 (URL はそのまま)
npm run deploy:single -- deploy --name <名前>     # 1 人分だけ更新
npm run deploy:single -- create --name <名前> --adminEmail <先生のメール> [--user organization] [--title <SS 名>]
npm run deploy:single -- list
```

- `create` は先生の Drive にスプレッドシート + 紐づけスクリプトを作り、配置する。先生がその URL を 1 度
  開いて許可すると、初回の準備 (`ensureDeploymentBootstrap_`) が済み、管理パネルが出る
- 組み立ては `.single-stage/<名前>/` (git に入らない)。src/ を写し、`Deployment.js` だけを差し替える
- 配置の前に生成物の鮮度を確かめる (`gen-pagehead.js --check` / `gen-utilities.js --check`)

---

## 4. pre-commit hook（自動品質ゲート）

`.githooks/pre-commit` がコミット前に以下を実行：

1. `node --check` で構文エラー
2. `scripts/lint.js --severity error` で CLAUDE.md ルール違反（ERROR）
3. `eslint --quiet src/*.js` でバグ予防レベル違反
4. `npm test` で全テスト

**有効化**（リポジトリごとに 1 回）:

```bash
git config core.hooksPath .githooks
```

無効化（緊急時）：`git commit --no-verify` または `core.hooksPath` を unset。

---

## 5. 典型的なワークフロー

### 5-1. バグ調査（ユーザー報告）

```bash
npm run api -- getLogs --limit 50              # セキュリティログ
npm run api -- getUserConfig --userId <uuid>   # 設定が壊れていないか確認
# 実行ログは Apps Script エディタの「実行数」で見る (1-1)
```

### 5-2. コード変更後の配置

```bash
npm test
npm run deploy:single -- deploy --name <検証用の名前>
npm run verify:e2e -- <検証用の名前>          # 64 項目
npm run deploy                                 # 全員の配置へ
npm run api -- systemDiagnosis
git commit -m "feat: ..." && git push
```

### 5-3. 大規模リファクタ前

```bash
npm run check                         # baseline 確認
npm run lint                          # WARN 状態を記録
npm run test:coverage                 # 既存カバレッジを記録
# 変更後に同じコマンドを走らせて差分が悪化していないか確認
```

---

## 6. CI（GitHub Actions）

`.github/workflows/ci.yml` が PR で自動実行：

- node 構文チェック
- `npm run lint:errors`
- `npm run check:utilities`（utility CSS 生成物の鮮度）
- `npm run check:pagehead`（共通 head 生成物の鮮度）
- `npm run check:gas`（GAS 参照整合性）
- `npm test`
- `npm run theme:perfect`

配置は CI からは**行わない**（ローカルの `npm run deploy` のみ）。

**注意: CI は本番のゲートではない。** 配置 → commit → push の順である以上、CI が回るのは配置の後になる。
本番を守るのは pre-commit と、検証用の配置での `verify:e2e`、`deploy-single.js` の preflight (生成物の鮮度)。

---

## 7. 移行専用ツールについて

移行が終わったツールは残さず削除する。残すと「使うべき現役ツール」と見分けがつかず、
しかも削除時に「そのツールを呼んでいたゲート」が無言で壊れる (実際 theme-perfect の
軸 2 が theme-pair-tailwind に判定を外注しており、削除で落ちた)。

**判定を外部ツールに外注しない。** ゲートは自己完結させ、移行ツールは変換だけを担う。

削除済み (2026-08-15、いずれも残作業 0 を確認してから削除):

| スクリプト | 役割 | 完了時点 |
| ---------- | ---- | -------- |
| `theme-tokenize.js` | slate/gray rgba → semantic token | v2925 |
| `theme-pair-tailwind.js` | 生パレット文字色 → theme 対応 | v2925 |
| `theme-normalize-typography.js` | 中途半端な寸法 → Tailwind scale | v2925 |
| `export-board-data.js` | 道徳 5/14 の 3 SS を CSV 書き出し (一回限り) | — |

現役の変換ツールは `tokenize-dimensions.js` (`npm run tokenize:dim`) のみ。
生値の再混入は `npm run check:tokens` が CI で止める。
