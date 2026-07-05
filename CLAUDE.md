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

**CI** (.github/workflows/ci.yml): Syntax check + `npm run lint:errors` + `npm test`（品質ゲートのみ。デプロイはローカルの `deploy:prod` で行う）

**2 テナント構成**: 本番は GAS 2 つ（那覇市 = default / 沖縄県 = open）。`deploy:prod` は那覇市のみを更新する。両方まとめて反映するには `npm run deploy:all`（那覇→沖縄を順にデプロイ）を使う。沖縄県のみは `npm run deploy:open`（`push:open` / `api:open` / `logs:open` も同様）。沖縄県の deploy 忘れによる旧コードドリフトを防ぐため、src 変更時は `deploy:all` 推奨。

---

## Admin API (本番アプリの状態取得・操作)

doPost の `adminApi` アクションに API キー認証でアクセス。ディスパッチャーは `src/AdminApis.js:dispatchAdminOperation()`。

```bash
npm run api -- systemDiagnosis                # DB接続・設定・デプロイ状態の健康診断
npm run api -- getLogs --limit 20             # セキュリティログ
npm run api -- getUsers                       # 全ユーザー一覧と設定状態
npm run api -- setUserConfig --userId X --patch '{...}'  # config パッチ適用
npm run api -- previewBoard --userId X        # viewer 視点の公開ボードプレビュー
```

**いつ使うか:**

- デプロイ後 → `systemDiagnosis` で本番が壊れていないか確認
- エラー調査時 → `getLogs` でセキュリティログ、`perfMetrics` でボトルネック特定
- ユーザー問題の調査 → `getUsers` で設定状態を確認
- 設定変更後 → `listProperties` で反映を確認

**全コマンド一覧**（ユーザー/ボード CRUD、Script Properties、SA pool 管理、Form 操作 v3、負荷検証）: [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md) §1-2。

---

## Cloud Logging (実行ログの取得・分析)

GCP の Cloud Logging API から GAS 実行ログを直接取得する。ブラウザ不要。

```bash
npm run logs:cloud -- --severity ERROR --hours 24   # 直近24hのエラー
npm run logs:cloud -- --user 35t22                  # 特定ユーザーでフィルタ
npm run logs:tail                                   # 直近10分 ERROR（デプロイ直後の確認）
```

**デフォルトで自分の deployment に絞る**: 那覇市版の GCP project は別 GAS アプリ (students/events_202604) と同居しており、そちらの ERROR が紛れ込む。`config.prodDeployId` で事前フィルタするため、デフォルトでは自アプリのログのみ返る。全アプリ見たい場合は `--all-deployments`。

**フロントエンドエラーも収集される**: `window.error` / `unhandledrejection` / `ErrorHandler.logError` は `reportClientError` doPost アクション経由で Cloud Logging に流れる（`[client/error]` プレフィックス）。`?debug=1` を付けると `console.log/warn` も収集（20件/セッション上限）。

**全 CLI オプション一覧**: [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md) §1-1。

---

## Code Quality (lint / pre-commit)

```bash
npm run lint                # CLAUDE.md ルール違反検出（setTimeout/innerHTML+var/getValue ループ等）
npm run lint:errors         # error severity のみ（CI 用）
npm run lint:eslint         # ESLint（バグ予防レベル）
npm run test:coverage       # node --test のカバレッジ
```

pre-commit hook は `npm install` 時に自動有効化される（`postinstall` で `git config core.hooksPath .githooks`）。コミット前に syntax + lint + test が自動実行される。緊急時は `git commit --no-verify`。

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
- **SharingHelper.js**: SS 共有設定 (SA pool 全員を editor 追加; 旧 domain-wide sharing は廃止)
- **SystemController.js**: System diagnostics, performance monitoring, deployment management (`publishApp` with etag conflict detection)
- **helpers.js**: Response builders (`createSuccessResponse`, `createErrorResponse`), `getCachedProperty` (30s TTL cache for PropertiesService)
- **validators.js**: All input validation functions

### SA Pool アーキテクチャ

「通常 Google Form 同等」セキュリティモデル: viewer (生徒) は SS 直接権限を持たず Web App 経由のみ、owner (教師) は自分の SS を `openById` 直接。cross-user アクセス (viewer 閲覧 / admin 管理) を **SA pool** (round-robin + 30s cooldown) が担う。

**アクセスモード自動振り分け** (`validateServiceAccountUsage` → `accessMode`, `openSpreadsheet` が経路判定):

| Caller | DB | own board | 他人公開 | 他人非公開 |
| ------ | -- | --------- | -------- | ---------- |
| **owner (editor)** | sa | **own (openById)** | sa | denied |
| **admin** | sa | **own** | sa | sa |
| **viewer (生徒)** | sa | — | sa | denied |

**設計詳細** (round-robin / cooldown / 2 段 token cache / per-row CAS lock / board data cache / sa_validation invalidate / 負荷検証): [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)。

### Theme / Color アーキテクチャ

ダークモード既定 + ライトモード対応。全 CSS は **semantic theme token** 経由で記述する。

**新規 CSS の Do/Don't**:

- ✅ **Do**: `var(--theme-text-primary)` / `var(--theme-bg-surface)` / `var(--theme-border-subtle)` (文脈依存の色)
- ✅ **Do**: brand identity / status は `var(--brand-like)` / `var(--status-success)` (theme 非依存)
- ✅ **Do**: 新規 UI は semantic primitives (`.card` / `.modal-card` / `.box`) を使う。inline Tailwind chain を新規追加しない
- ❌ **Don't**: ハードコード hex/rgba (`color: #94a3b8` / `background: rgba(255,255,255,0.1)`)
- ❌ **Don't**: Tailwind `dark:` prefix (CI で機械検証・禁止)。light/dark は `--theme-*` token + `body.theme-light` で表現

**色を変更したいとき**: [UnifiedStyles.css.html](src/UnifiedStyles.css.html) の `:root` (dark) と `body.theme-light` (light) の token を書き換えるだけで全体反映。

**検証**: `npm run theme:contrast` (WCAG AA) / `theme:matrix` (全 token + contrast) / `theme:perfect` (統合ゲート)。

**パレット設計哲学・token 一覧・primitives 表・status box token・themeManager API・保守 CLI 詳細**: [docs/THEME.md](docs/THEME.md)。

### Cache アーキテクチャ (3 層、 意図的分離)

3 つの層はそれぞれ異なる用途。統合 facade を作らない (semantic clarity を失う + cost 非対称性が見えなくなる)。

| 層 | TTL | 用途 | 注意 |
| -- | --- | ---- | ---- |
| `getCachedProperty(key)` | 30s in-memory LRU (50 件) | **頻繁に読む config** (ADMIN_EMAIL, DATABASE_SPREADSHEET_ID, SA creds 等) | write は必ず `setCachedProperty` 経由 (in-memory 即時 invalidate)。 PropertiesService 直読みより 50-100ms 速い |
| `CacheService.getScriptCache()` | 最大 6h、 100KB/value | **作業データ** (board data, user objects, SA tokens, rate-limit counter, 429 cooldown) | size overrun は silent fail。 必ず `saveToCacheWithSizeCheck()` 経由 |
| `PropertiesService` | 永続 (no TTL) | **永続的システム設定 / 秘密情報** (SA creds JSON, ADMIN_API_KEY, DATABASE_SPREADSHEET_ID, ADMIN_EMAIL) | hot path で直読みすると 50-100ms/call。 `getCachedProperty` 経由を推奨。 例外: `handleSetupApiKeyAction_` 初回ゲートは 30s stale race 回避のため直読み (`main.js` 参照) |

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
| `SERVICE_ACCOUNT_CREDS_2` 〜 `_10` | (Optional) SA pool secondaries (700人スケールで 3-5 個推奨) |
| `GOOGLE_CLIENT_ID` | (Optional) OAuth client ID |
| `ADMIN_API_KEY` | Admin API authentication key |

SA pool のセキュリティモデル・アクセス経路の詳細は [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)。

---

## 兄弟プロジェクト (StudyQuest との連携)

「みんなの回答ボード」 = 集合可視化の場、 一方「[StudyQuest](file:///Users/ryuma/slide_generator)」 (= slide_generator) は個別マイプラン学習 (3 touch 構造) の場。 両者は **データだけで疎結合に連携**:

- 連携責任: **StudyQuest 側に Answer Board Bridge 機能を実装** (回答ボード側はゼロ変更)
- 設計詳細: [StudyQuest docs/answerboard-bridge.md](file:///Users/ryuma/slide_generator/docs/answerboard-bridge.md)
- 動作: StudyQuest タッチ 3 (ふりかえり) → prefilled Google Form URL 生成 → 回答ボード Form に submit → matrix/wordcloud で集合可視化
- 文献的根拠: 個別最適な学びの継続困難 5 大要因 (教師負担集中・実践知の継承困難・進度差対応・学力低下・標準化テンション) → アプリは「1 単元 8 時間中 7-10% だけ介入」 の minimum viable 個別最適を堅持
- 回答ボード側で新規 API / OAuth scope 追加なし。 Google Form の prefilled URL (`?usp=pp_url&entry.XXX=...`) が連携プロトコル
