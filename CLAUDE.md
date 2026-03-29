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
npm run api -- systemDiagnosis      # システム診断（DB接続、設定、デプロイ状態）
npm run api -- getAppStatus         # アプリ有効/無効状態
npm run api -- getUsers             # 全ユーザー一覧と設定状態
npm run api -- getLogs --limit 20   # セキュリティログ
npm run api -- perfMetrics          # パフォーマンスメトリクス
npm run api -- perfDiagnosis        # パフォーマンス診断と推奨事項
npm run api -- listProperties       # Script Properties一覧（認証情報はマスク）
npm run api -- cacheReset           # キャッシュ全クリア
npm run api -- autoRepair           # 自動修復
```

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
npm run logs:cloud                            # 直近6hのWARNING以上
npm run logs:cloud -- --severity ERROR         # ERRORのみ
npm run logs:cloud -- --severity INFO          # INFO以上（全ログ）
npm run logs:cloud -- --hours 24 --limit 50    # 過去24時間、50件
npm run logs:cloud -- --function doPost        # 特定関数のログ
npm run logs:cloud -- --user 35t22             # 特定ユーザーのログ
npm run logs:cloud -- --json                   # JSON形式で出力
```

**いつ使うか:**

- エラー調査 → `--severity ERROR --hours 72` で直近のエラーを確認
- 特定ユーザーの問題 → `--user メールアドレスの一部` でフィルタ
- デプロイ後の監視 → `--severity WARNING --hours 1` で直近の異常を確認

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
- **DatabaseCore.js**: All spreadsheet reads/writes, includes circuit breaker pattern and `executeWithRetry` for transient failures
- **SystemController.js**: System diagnostics, performance monitoring, deployment management (`publishApp` with etag conflict detection)
- **helpers.js**: Response builders (`createSuccessResponse`, `createErrorResponse`), `getCachedProperty` (30s TTL cache for PropertiesService)
- **validators.js**: All input validation functions

### Frontend

HTML templates use `<?!= include('filename') ?>` for composition. Key patterns:

- `SharedUtilities.html`: Common JS utilities including `escapeHtml()`
- `page.js.html` / `AdminPanel.js.html`: Single-file JS (no physical splitting - intentional)
- Top-level side effects forbidden - `google.script.run`, DOM manipulation, and auto-timers must be inside `init()` only
- Include order is fixed; changes require a standalone commit

### Data Store

Google Sheets as database via service account. `users` sheet stores user records; each user's board config is in a JSON `config` column.

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
| `SERVICE_ACCOUNT_CREDS` | Service account JSON credentials |
| `GOOGLE_CLIENT_ID` | (Optional) OAuth client ID |
| `ADMIN_API_KEY` | Admin API authentication key |
