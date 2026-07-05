# 開発・保守ツール一覧（CLI ベース）

このプロジェクトの開発・保守は全て CLI から完結するよう構成されています。
このドキュメントは「何ができて、いつ使うか」のマップです。

---

## 1. 観測：本番状態とログの取得

### 1-1. Cloud Logging（GAS 実行ログ + フロントエンドエラー）

```bash
npm run logs:cloud                              # 直近 6h, WARNING 以上 (自分の prodDeployId のみ)
npm run logs:cloud -- --severity ERROR --hours 24
npm run logs:cloud -- --user naha-okinawa      # メアド断片でフィルタ
npm run logs:cloud -- --function doPost        # 関数名でフィルタ
npm run logs:cloud -- --all-deployments        # 共有 GCP project の全アプリログ
npm run logs:cloud -- --deployment <id>        # 特定 deployment にフィルタ
npm run logs:cloud -- --json                   # JSON 形式で出力
npm run logs:cloud -- --brief                  # 1 行/エントリ簡潔モード
npm run logs:summary                           # シグネチャ別集計
npm run logs:errors                            # 直近 24h ERROR 集計
npm run logs:tail                              # 直近 10 分 ERROR
npm run logs:open                              # open 環境用
```

**デフォルトで自分の deployment に絞る**：那覇市版の GCP project には別 GAS アプリ
（students/events_202604）が同居しており、その ERROR が紛れ込んでデバッグを阻害する。
`config.prodDeployId` フィルタが既定でかかるので、自分のアプリのログだけが返る。
共有プロジェクト全体を見たいときは `--all-deployments` を付与。

**フロントエンドエラーも含む**：`window.error` / `unhandledrejection` / `ErrorHandler.logError`
は `reportClientError` doPost アクション経由で Cloud Logging に流れる。
`[client/error]` / `[client/warn]` プレフィックスで識別。

#### `?debug=1` でフロントの console.log/warn も収集

児童・教師が問題を再現するとき、URL に `?debug=1` を付けて開いてもらうと、
ブラウザの console.log/warn も Cloud Logging に流れる（最大 20 件 / セッション）。

### 1-2. Admin API（本番状態確認・操作）

#### 基本（システム状態）

```bash
npm run api -- systemDiagnosis        # Core props / DB / Deploy の健康診断
npm run api -- getUsers               # 全ユーザー一覧
npm run api -- getAppStatus           # アプリ有効/無効状態
npm run api -- enableApp              # アプリを有効化
npm run api -- disableApp             # 緊急停止（アプリ無効化）
npm run api -- getLogs --limit 20     # セキュリティログ
npm run api -- perfMetrics            # パフォーマンス指標
npm run api -- perfDiagnosis          # パフォーマンス診断 + 推奨事項
npm run api -- listProperties         # Script Properties（認証情報マスク済み）
npm run api -- getProperty --k NAME           # 個別 property 取得
npm run api -- setProperty --k NAME --v VALUE # property 更新
npm run api -- cacheReset             # 全キャッシュクリア
npm run api -- autoRepair             # 自動修復
```

#### ユーザー config 操作（CLI から本番設定を読み書き）

```bash
# 読み取り
npm run api -- findUser --email t260781p@naha-okinawa.ed.jp
npm run api -- getUserConfig --userId adb24b94-...
npm run api -- exportConfigs --output backup.json          # 全ユーザー dump（ファイル）

# 書き込み（部分 patch、deep merge、保護フィールド自動除外）
npm run api -- setUserConfig --userId <uuid> \
  --patch '{"displaySettings":{"boardMode":"numberline"}}'

npm run api -- setUserConfig --userId <uuid> \
  --patch '{"xAxisLabels":{"min":"そのまま使う","max":"自分で直す"},"allowResubmit":true}'

# bulk: 条件に合うユーザー全員に同じ patch を適用
npm run api -- bulkSetUserConfig \
  --filter '{"isPublished":false}' \
  --patch '{"displaySettings":{"showReactions":true}}' \
  --dryRun                              # 影響範囲を先に確認

npm run api -- bulkSetUserConfig --filter '{"emailContains":"naha"}' --patch '...'

# アクティブ / 公開状態の切替
npm run api -- toggleUserActive --userId <uuid>            # アクティブ状態切替
npm run api -- toggleUserBoard --userId <uuid>             # ボード公開状態切替

# 分析・プレビュー
npm run api -- runColumnAnalysis --userId <uuid>           # 列マッピング自動検出
npm run api -- previewBoard --userId <uuid>                # viewer 視点のサンプル

# Form 操作 (v3)
npm run api -- listMyForms                                  # 自分の Drive 内 Form を最大 30 件
npm run api -- validateFormUrl --formUrl <URL>              # URL 形式 + 到達性チェック
npm run api -- connectForm --userId <uuid> --formUrl <URL>  # 既存 Form を user に紐付け
npm run api -- createForm --userId <uuid> --templateType board|numberline|matrix
# templateType=numberline/matrix のときは boardMode も自動セット

# JSON 引数の渡し方
#   --patch '<json>'   : JSON-aware（自動パース）
#   --filter '<json>'  : 同上
#   --json '<json>'    : 任意 JSON を json param に格納
#   --file path.json   : ファイルからまるごと params 読み込み
#   --output path      : 結果を JSON ファイルに保存
```

**API キーは `scripts/config.json` に保存**（git ignore 済み、template から作成）。
初回設定は `setupApiKey` アクションで手動投入が必要。

**保護されているフィールド**（patch から自動除外）：
`userId`, `userEmail`, `googleId`, `createdAt`, `etag`

### 1-3. SA pool 管理

700 人スケールの同時アクセスを service account pool で捌く。設計詳細は [ARCHITECTURE.md](ARCHITECTURE.md)。

```bash
npm run api -- listServiceAccountPool         # pool 一覧 (slot/email/projectId)
npm run api -- getServiceAccountUsage         # SA 別 5分窓使用回数 + cooling 状態
npm run api -- addServiceAccountToPool --json '<SA JSON>'     # 1 個追加 (DB editor + verify 自動)
npm run api -- addServiceAccountsToPoolBatch --inputs '<...>' # 複数一括 (改行 or `}\n{` 区切り)
npm run api -- reverifyServiceAccountInPool --slot SERVICE_ACCOUNT_CREDS_2  # 共有反映後の再検証
npm run api -- removeServiceAccountFromPool --slot SERVICE_ACCOUNT_CREDS_2  # secondary 削除
npm run api -- repairSpreadsheetSharing --spreadsheetId <ID>  # 1 SS に SA pool 全員 editor 追加
npm run api -- migrateBoardSharing --dryRun true   # 既存ボード一括 cleanup の dry run
npm run api -- migrateBoardSharing                 # 全ボードに SA pool editor 追加 + domain 共有 revoke
```

**負荷検証**（CI 対象外、手動）:

```bash
node scripts/load-test-concurrent.js --n 30 --op previewBoard  # N 並列 API call で SA pool / 429 を検証
```

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
npm run check                         # test + systemDiagnosis（本番疎通含む）
```

---

## 3. デプロイと環境

```bash
npm run push                          # ローカル → GAS（デプロイは別）
npm run pull                          # GAS → ローカル
npm run deploy:prod                   # URL 維持で本番再デプロイ（push 込み）
npm run deploy                        # 新 URL でテスト用デプロイ
npm run deploy:list                   # 既存デプロイ一覧
```

`open` ドメイン用：`npm run push:open` / `npm run deploy:open` / `npm run api:open` / `npm run logs:open`

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
npm run logs:cloud -- --severity ERROR --hours 24 --user <メアド断片>
npm run api -- getUsers                        # 設定が壊れていないか確認
npm run api -- perfMetrics                     # クォータ枯渇していないか
# 必要なら: ユーザーに ?debug=1 を付けたURLでアクセスを依頼 → ログ再取得
```

### 5-2. コード変更後の本番デプロイ

```bash
# 1. ローカル変更
# 2. pre-commit が走るのでテスト通過を確認
git commit -m "feat: ..."

# 3. デプロイ
npm run deploy:prod

# 4. 本番疎通確認
npm run api -- systemDiagnosis

# 5. 30 分後にエラーが増えていないか
npm run logs:cloud -- --severity ERROR --hours 0.5
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
- `npm test`
- （将来）`npm run lint:errors` と `npm run lint:eslint` も組み込み可能

デプロイは CI からは**行わない**（ローカル `deploy:prod` のみ）。
