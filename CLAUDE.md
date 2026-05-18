# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

**Project**: Google Apps Script Web Application for organization-internal use (collaborative answer board for Google Forms responses)
**Stack**: Zero-dependency, direct GAS API calls, V8 runtime, Tailwind CSS frontend
**Language**: Comments and UI strings are in Japanese; code identifiers are in English

---

## Critical: GAS Single Global Scope

**All script files execute in a single global scope** - no module system, no import/export.

- Every function is globally accessible from any file, including via `google.script.run.funcName()` from HTML
- Watch for namespace collisions - function names must be unique across all files
- `/* global ... */` comments at file tops declare cross-file dependencies for linting only
- File loading order is irrelevant

---

## Commands

```bash
npm run push              # Push code to GAS
npm run pull              # Pull from GAS
npm run logs              # View execution logs
npm test                  # Run all tests
npm run deploy:prod       # URL維持で本番デプロイ（pushも含む）
npm run deploy            # 新しいURLでデプロイ（テスト用）
node --test tests/main.doPost.test.cjs   # Run a single test file
```

**Dev workflow**: Edit locally → `npm test` → `npm run deploy:prod` → Git commit → `git push`

**CI** (.github/workflows/ci.yml): Syntax check + `npm test`（品質ゲートのみ。デプロイはローカルの `deploy:prod` で行う）

---

## Admin API (本番アプリの状態取得・操作)

コード変更後、本番で正しく動いているか確認するために使う。

```bash
# システム / 監視
npm run api -- systemDiagnosis      # システム診断（DB接続、設定、デプロイ状態）
npm run api -- getAppStatus         # アプリ有効/無効状態
npm run api -- enableApp            # アプリを有効化
npm run api -- disableApp           # 緊急停止（アプリ無効化）
npm run api -- getLogs --limit 20   # セキュリティログ
npm run api -- perfMetrics          # パフォーマンスメトリクス
npm run api -- perfDiagnosis        # パフォーマンス診断と推奨事項
npm run api -- cacheReset           # キャッシュ全クリア
npm run api -- autoRepair           # 自動修復

# ユーザー / ボード CRUD
npm run api -- getUsers             # 全ユーザー一覧と設定状態
npm run api -- findUser --email X   # メールでユーザー検索
npm run api -- getUserConfig --userId X       # 個別 config の取得
npm run api -- setUserConfig --userId X --patch '{...}'  # config パッチ適用
npm run api -- bulkSetUserConfig --filter '{...}' --patch '{...}'  # 一括更新
npm run api -- toggleUserActive --userId X    # アクティブ状態切替
npm run api -- toggleUserBoard --userId X     # ボード公開状態切替
npm run api -- exportConfigs        # 全 config のエクスポート
npm run api -- runColumnAnalysis --userId X   # 列マッピング自動分析
npm run api -- previewBoard --userId X        # 公開ボードのプレビュー

# Script Properties
npm run api -- listProperties       # Script Properties一覧（認証情報はマスク）
npm run api -- getProperty --k NAME           # 個別 property 取得
npm run api -- setProperty --k NAME --v VALUE # property 更新

# SA pool 管理 (v2782+)
npm run api -- listServiceAccountPool         # pool 一覧 (slot/email/projectId)
npm run api -- getServiceAccountUsage         # SA 別 5分窓使用回数 + cooling 状態
npm run api -- addServiceAccountToPool --json '<SA JSON>'    # 1 個追加 (DB editor + verify 自動)
npm run api -- addServiceAccountsToPoolBatch --inputs '<...>' # 複数一括 (改行 or `}\n{` 区切り)
npm run api -- reverifyServiceAccountInPool --slot SERVICE_ACCOUNT_CREDS_2  # 共有反映後の再検証
npm run api -- removeServiceAccountFromPool --slot SERVICE_ACCOUNT_CREDS_2  # secondary 削除
npm run api -- repairSpreadsheetSharing --spreadsheetId <ID>  # 1 SS に SA pool 全員 editor 追加
npm run api -- migrateBoardSharing --dryRun true   # 既存ボード一括 cleanup の dry run
npm run api -- migrateBoardSharing                 # 全ボードに SA pool editor 追加 + domain 共有 revoke

# 負荷検証 (CI 対象外、 手動)
node scripts/load-test-concurrent.js --n 30 --op previewBoard  # N 並列 API call で SA pool / 429 を検証
```

詳細・全オプション: [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md) §1-2

**いつ使うか:**

- コード変更後のデプロイ後 → `systemDiagnosis` で本番が壊れていないか確認
- エラー調査時 → `getLogs` でセキュリティログを確認、`perfMetrics` でボトルネック特定
- ユーザー問題の調査 → `getUsers` で設定状態を確認
- 設定変更後 → `listProperties` で反映を確認

**仕組み:** doPostの`adminApi`アクションにAPIキー認証でアクセス。ディスパッチャーは`src/AdminApis.js:dispatchAdminOperation()`。

---

## Cloud Logging (実行ログの取得・分析)

GCPのCloud Logging APIからGASの実行ログを直接取得する。ブラウザ不要。

```bash
npm run logs:cloud                            # 直近6h, WARNING以上 (自分の prodDeployId のみ)
npm run logs:cloud -- --severity ERROR         # ERRORのみ
npm run logs:cloud -- --severity INFO          # INFO以上（全ログ）
npm run logs:cloud -- --hours 24 --limit 50    # 過去24時間、50件
npm run logs:cloud -- --function doPost        # 特定関数のログ
npm run logs:cloud -- --user 35t22             # 特定ユーザーのログ
npm run logs:cloud -- --json                   # JSON形式で出力
npm run logs:cloud -- --all-deployments        # 共有 GCP project の全アプリログ
npm run logs:cloud -- --deployment <id>        # 特定 deployment_id にフィルタ
npm run logs:cloud -- --summary                # メッセージシグネチャで集計
npm run logs:cloud -- --brief                  # 1行/エントリ簡潔モード
npm run logs:cloud -- --tail                   # 直近10分 ERROR のみ（デプロイ直後の確認）
```

**デフォルトで自分の deployment に絞る**: 那覇市版の GCP project は別 GAS アプリ
(students/events_202604) と同居しており、そちらの ERROR が紛れ込む。`config.prodDeployId` で
事前フィルタするため、デフォルトでは自アプリのログのみ返る。全アプリ見たい場合は `--all-deployments`。

**いつ使うか:**

- エラー調査 → `--severity ERROR --hours 72` で直近のエラーを確認
- 特定ユーザーの問題 → `--user メールアドレスの一部` でフィルタ
- デプロイ後の監視 → `--severity WARNING --hours 1` で直近の異常を確認

**フロントエンドエラーも収集される**：`window.error` / `unhandledrejection` / `ErrorHandler.logError`
は `reportClientError` doPost アクション経由で Cloud Logging に流れる（`[client/error]` プレフィックス）。
`?debug=1` を付けると `console.log/warn` も収集（20件/セッション上限）。

詳細・全 CLI ツール一覧：[docs/DEVELOPMENT.md](docs/DEVELOPMENT.md)

---

## Code Quality (lint / pre-commit)

```bash
npm run lint                # CLAUDE.md ルール違反検出（setTimeout/innerHTML+var/getValue ループ等）
npm run lint:errors         # error severity のみ（CI 用）
npm run lint:eslint         # ESLint（バグ予防レベル）
npm run test:coverage       # node --test のカバレッジ
```

pre-commit hook 有効化（1 回のみ実行）：

```bash
git config core.hooksPath .githooks
```

これでコミット前に syntax + lint + test が自動実行されます。緊急時は `--no-verify`。

---

## Architecture

### Request Flow

- **`doGet(e)`** (main.js): Routes `?mode=` parameter to HTML templates via `HtmlService`. Modes: `login`, `setup`, `appSetup`, `manual`, `view`, `admin`
- **`doPost(e)`** (main.js): Routes JSON `action` field to handler functions. Actions are allowlist-managed; each must have input validation.
- Both entry points authenticate via `Session.getActiveUser().getEmail()`

### Service Layer Pattern

```text
main.js (doGet/doPost routing, auth)
  → *Apis.js (API endpoints, parameter extraction, response formatting)
    → *Service.js (business logic, data transformation)
      → DatabaseCore.js (spreadsheet I/O, circuit breaker, retries)
```

- **Apis files** group endpoints by domain (AdminApis, UserApis, DataApis)
- **Service files** contain pure business logic (ConfigService, DataService, UserService, ReactionService)
- **DatabaseCore.js**: All spreadsheet I/O, SA pool (round-robin + cooldown + 2-tier token cache), circuit breaker, `executeWithRetry`
- **SharingHelper.js**: SS 共有設定 (SA pool 全員を editor 追加; v2782+ 旧 domain-wide sharing は廃止)
- **SystemController.js**: System diagnostics, performance monitoring, deployment management (`publishApp` with etag conflict detection)
- **helpers.js**: Response builders (`createSuccessResponse`, `createErrorResponse`), `getCachedProperty` (30s TTL cache for PropertiesService)
- **validators.js**: All input validation functions

### SA Pool アーキテクチャ (v2782+)

「通常 Google Form 同等」 セキュリティモデル (viewer は SS 直接権限なし, Drive 非表示)。
ボード SS への cross-user アクセス (生徒の閲覧 / 管理者の管理) を **SA pool** が担う。

**アクセスモード自動振り分け** (`validateServiceAccountUsage` → `accessMode`):

| Caller | DB | own board | 他人公開 | 他人非公開 |
| ------ | -- | --------- | -------- | ---------- |
| **owner (editor)** | sa | **own (openById)** | sa | denied |
| **admin** | sa | **own** | sa | sa |
| **viewer (生徒)** | sa | — | sa | denied |

owner は own OAuth で SA quota 節約。 viewer / admin の cross-user のみ SA pool 経由。

**SA pool 設計** ([DatabaseCore.js](src/DatabaseCore.js)):

- **round-robin** (`pickServiceAccount_`): ScriptCache 共有 counter `sa_rr_counter` で完全均等分散
- **30s cooldown**: 429 喰った SA は除外 → auto failover
- **2 段 token cache** (`getServiceAccountAccessToken_`): in-memory + ScriptCache 50min
- **per-SA per-sheet verify cache** (`verifyServiceAccountAccess_`): 10min ok / 2min no (transient は焼かない)
- **authResolver closure** (`makeProxyAuthResolver_`): proxy hot path で SA 動的切替

**per-row CAS lock** (v2785, [ReactionService.js](src/ReactionService.js)):

- 旧: `LockService.getScriptLock()` を process() 全体で保持 → 全 board が serialize する 700人 burst の bottleneck
- 新: ScriptLock は ~5-10ms の critical section (cache check+put) のみ。 異 row 同士は完全並列
- throughput ~5 → ~200 req/sec, 40x 向上

**board data cache** (v2783, [DataApis.js](src/DataApis.js)):

- viewer の `getPublishedSheetData` 結果を 10s ScriptCache
- reaction/highlight write 時に `bumpBoardDataVersion_` で即時 stale
- 700 viewer × 8s polling = 5,250 req/min → ~700 req/min (~8x 削減)

**sa_validation cache 即時 invalidate** (v2785):

- `__applyPublishStateChange` で `invalidateSaValidationCache_` を呼び、 該当 SS の cache version を bump
- unpublish 直後の 60秒 access leak を解消

### Cache アーキテクチャ (3 層、 意図的分離)

3 つの層はそれぞれ異なる用途。 統合 facade を作らない (semantic clarity を失う + cost 非対称性が見えなくなる)。

| 層 | TTL | 用途 | 注意 |
| -- | --- | ---- | ---- |
| `getCachedProperty(key)` | 30s in-memory LRU (50 件) | **頻繁に読む config** (ADMIN_EMAIL, DATABASE_SPREADSHEET_ID, SA creds 等) | write は必ず `setCachedProperty` 経由 (in-memory 即時 invalidate)。 PropertiesService 直読みより 50-100ms 速い |
| `CacheService.getScriptCache()` | 最大 6h、 100KB/value | **作業データ** (board data, user objects, SA tokens, rate-limit counter, 429 cooldown) | size overrun は silent fail。 必ず `saveToCacheWithSizeCheck()` 経由 |
| `PropertiesService` | 永続 (no TTL) | **永続的システム設定 / 秘密情報** (SA creds JSON, ADMIN_API_KEY, DATABASE_SPREADSHEET_ID, ADMIN_EMAIL) | hot path で直読みすると 50-100ms/call。 `getCachedProperty` 経由を推奨。 例外: `handleSetupApiKeyAction_` 初回ゲートは 30s stale race 回避のため直読み (`main.js:660` 参照) |

**判断フロー (新規 cache を書く時):**

1. 永続 secret / 設定 → `PropertiesService.setProperty` (direct write) + 読みは `getCachedProperty`
2. ボード data / 計算結果 / token 等 (30s〜6h で再計算可) → `CacheService` + `saveToCacheWithSizeCheck`
3. すでに `getCachedProperty` で 30s memoize される config への参照 → 何もしない (透過的)

### Frontend

HTML templates use `<?!= include('filename') ?>` for composition. Key patterns:

- `SharedUtilities.html`: Common JS utilities including `escapeHtml()`
- `page.js.html` / `AdminPanel.js.html`: Single-file JS (no physical splitting - intentional)
- Top-level side effects forbidden - `google.script.run`, DOM manipulation, and auto-timers must be inside `init()` only
- Include order is fixed; changes require a standalone commit

### Canonical Frontend APIs (Use these, don't reinvent)

新規コードを書くときは以下の集約 API を使う。ad-hoc DOM 操作や native `confirm/alert` を新規に書かない。

| 用途 | API | 場所 |
| ---- | --- | ---- |
| 右上トースト (success/error/warning/info) | `notifications.success/error/warning/info(msg, opts?)` | [SharedUtilities.html](src/SharedUtilities.html) |
| 中央上バナー (ボード切替予告等) | `notifications.banner(msg, { html?, duration? })` | 同上 |
| 確認 / アラートモーダル | `await modals.confirm(msg, opts?)` / `await modals.alert(msg)` | 同上 |
| リダイレクトモーダル (公開遷移用) | `showRedirectModal({ title, message, redirectUrl, variant? })` | 同上 |
| サーバ呼び出し (Promise) | `await runServer('funcName', arg1, arg2)` | 同上 |
| 安全な JSON parse (fallback 返却) | `safeJsonParse(text, fallback)` | 同上 |
| HTML escape | `sharedUtilities.security.escapeHtml(s)` | 同上 |
| ローディングオーバーレイ | `setLoading(true, msg)` / `setLoading(false)` | 同上 |
| デバウンス (key 共有可) | `sharedUtilities.debounce.debounce(fn, key, delayMs)` | 同上 |
| ボタンの処理中 UX | `sharedUtilities.buttons.withBusy(btn, asyncFn, { busyText })` | 同上 |

**禁止 / 非推奨パターン:**

- 新規の右上トースト DOM、中央バナー DOM、`#notification-container` を手書きしない (上記 API に集約)
- 新規の native `confirm()` / `alert()` (緊急 fallback 用途を除く) → `modals.confirm/alert` を使う
- `withSuccessHandler/withFailureHandler` の手書きチェーン → `runServer(...)` を使う
- `JSON.parse(...)` を try/catch なしで呼ぶ → `safeJsonParse(...)` を使う

### Data Store

Google Sheets as database via service account. `users` sheet stores user records; each user's board config is in a JSON `configJson` column (see `src/SystemController.js` `USERS_SHEET_HEADERS`).

### Visualization Modes (M1 / M2)

`config.displaySettings.boardMode` で表示モードを切替:

- `auto` (既定): numericX/Y の有無から自動判定
- `board`: 従来の掲示板（カード一覧）— columnMapping に `answer` / `reason` 必須
- `numberline` (M1): d3 ビーズワーム数直線 — `numericX` 必須
- `matrix` (M2): 2軸散布図 — `numericX` + `numericY` 必須

管理パネル「📝 Googleフォーム選択」の種別 select から、各モード用の Forms 構造を1クリックで生成可能（`createTemplateForm(templateType)`）。`numberline`/`matrix` 選択時は `boardMode` も自動セット。詳細仕様: [docs/SPEC_visualization_modes.md](docs/SPEC_visualization_modes.md)。

---

## Must-Follow Rules

### Performance

- **Batch operations always** - `getDataRange().getValues()` not `getRange(i,j).getValue()` in loops (70x difference)
- **Cache PropertiesService** - use `getCachedProperty()` with 30s TTL (see helpers.js)
- **Minimize external service calls** - in-script JS operations are faster

### Security

- **`Session.getActiveUser()` only** - never `getEffectiveUser()` (privilege escalation risk)
- **Validate all inputs** - use validators.js functions before processing
- **Sanitize HTML** - `textContent` not `innerHTML` for user content; use `escapeHtml()` if innerHTML required

### V8 Runtime Constraints

- **`Utilities.sleep()`** - `setTimeout`/`setInterval` do not exist in GAS
- Modern JS OK (arrow functions, optional chaining, nullish coalescing, destructuring)
- No Node.js APIs - use GAS services (DriveApp, SpreadsheetApp, etc.)

### Maintenance Rules (from README)

- `publishApp` accepts only allowlisted fields with `etag` conflict detection - maintain both
- `doPost` actions are allowlist-managed; adding an action requires adding its input validation simultaneously
- Include order in HTML templates is fixed; changes to it require a standalone commit

### Board Publish Lifecycle (single source of truth)

`config.isPublished` / `config.publishedAt` を書き換えてよいのは以下 **4 関数のみ**。直接 saveUserConfig で書き換えてはいけない（DRY 違反 + テスト不可）：

| 関数 | 場所 | 役割 |
| ---- | ---- | ---- |
| `publishApp(publishConfig)` | `SystemController.js` | 初回公開 / 設定差し替えを伴う公開（allowlist + etag 検証） |
| `republishMyBoard(options?)` | `AdminApis.js` | 所有者が自分のボードを再び公開する（state のみ） |
| `unpublishBoard(targetUserId?, options?)` | `AdminApis.js` | 所有者 or 管理者がボードの公開を終了する |
| `toggleUserBoardStatus(targetUserId, options?)` | `AdminApis.js` | 管理者が他ユーザーのボード公開状態を toggle する |

**統一不変条件**（`__applyPublishStateChange` で集約）：

- `publishedAt` は「現在の状態の発生時刻」（公開中 = now / 非公開 = null）
- `lastAccessedAt` は変更時に touch
- `options.sourceEtag` を渡せば etag conflict 検出が有効化される（楽観ロック）
- 認可: 自分のボード or 管理者（`requireAdmin: true` 指定時は管理者のみ）
- 戻り値: `{ success, message, userId, isPublished, boardPublished, publishedAt, etag, redirectUrl }` で統一

**UI 用語**：

- アクション: 「公開する」「再び公開する」「公開を終了する」
- 状態: 「公開中」「非公開」

詳細仕様とテスト: `tests/admin.publishLifecycle.test.cjs`

---

## Testing

Tests run in Node.js using `node:test` (no frameworks). GAS globals are stubbed via `vm.createContext()` - see `tests/main.doPost.test.cjs` for the pattern. Each test file loads source files into a sandboxed VM context with mock GAS services (`ContentService`, `Session`, `SpreadsheetApp`, etc.).

---

## clasp + Git

- **Never commit**: `.clasp.json`, `.clasprc.json` (contain sensitive scriptId/credentials)
- **Do commit**: `src/**/*.js`, `src/**/*.html`, `src/appsscript.json`, `.clasp.json.template`
- File extension: `.js` locally (clasp convention), shows as `.gs` in GAS editor

### Required Script Properties

Set via GAS editor or SetupPage:

| Property | Description |
| -------- | ----------- |
| `ADMIN_EMAIL` | Administrator email address |
| `DATABASE_SPREADSHEET_ID` | Main database spreadsheet ID |
| `SERVICE_ACCOUNT_CREDS` | Primary SA JSON credentials |
| `SERVICE_ACCOUNT_CREDS_2` 〜 `_10` | (Optional) SA pool secondaries (v2782+, 700人スケールで 3-5 個推奨) |
| `GOOGLE_CLIENT_ID` | (Optional) OAuth client ID |
| `ADMIN_API_KEY` | Admin API authentication key |

### SA pool セキュリティモデル (v2782+)

ボード SS は **通常 Google Form 同等**のセキュリティ:

- viewer (生徒) は SS への直接権限を持たず、 Web App 経由のみアクセス可能 (SA pool 経由)
- viewer の Drive にボード SS は表示されない、 直接編集も不可
- owner (教師) は自分の SS のオーナー権限を持つので `openById` 直接 (SA quota 節約)
- 旧 `DOMAIN_WITH_LINK + EDIT` 共有は廃止 — `migrateBoardSharing` admin API で既存ボードを cleanup

`openSpreadsheet` の accessMode 自動判定:

| Caller | 経路 |
| ------ | ---- |
| owner (= 自分のボード) | `openById` (own OAuth) |
| viewer / admin (= cross-user) | SA pool (round-robin + 30s cooldown) |
| DB sheet | SA pool (常に) |
