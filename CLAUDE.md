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
npm test                                 # Run all tests
node --test tests/main.doPost.test.cjs   # Run a single test file
npm run deploy                           # 全員の配置を更新 (URL はそのまま。deploy-single.js deploy --all)
npm run deploy:single -- deploy --name <名前>   # 1 人分だけ更新
npm run deploy:single -- create --name <名前> --adminEmail <先生のメール>   # 先生を足す
npm run verify:e2e -- t260781p-verify    # 検証用の配置で画面の機能を端から端まで動かす (54 項目)
```

**配置の形**: 先生ごとに 1 つ、先生の Drive のスプレッドシートに紐づいたアプリを「先生として実行」で配置する
(`src/appsscript.json` の `webapp.executeAs = USER_DEPLOYING`)。データ (users / lessons / 授業の回答シート) は
そのスプレッドシート 1 つに入る。配置ごとに違うのは `src/Deployment.js` (先生のメールと DB の SS ID) だけで、
`scripts/deploy-single.js` が焼き込む。先生ごとの値は `scripts/single-tenants.json` (git に入れない)。
Why: 2026-10-29 から児童アカウントは「開いた人として実行」の Apps Script を一切開けない (管理側の設定)。

**Dev workflow**: Edit locally → `npm test` → `npm run deploy:single -- deploy --name t260781p-verify` →
`npm run verify:e2e -- t260781p-verify` → `npm run deploy` → Git commit → `git push`

**検証用の配置** (`t260781p-verify`): `extraFiles: ["scripts/verify/DevInvoke.js"]` 付き。admin API の `dev.invoke` で
画面が `google.script.run` で呼ぶ処理を名前で呼べる (児童のふりは `as`)。本番の先生には付けない。

**CI** (.github/workflows/ci.yml): Syntax check + `npm run lint:errors` + `npm test`（品質ゲートのみ。配置はローカルで行う）

---

## Admin API (配置の状態取得・操作)

doPost の `adminApi` アクションに API キー認証でアクセス。ディスパッチャーは `src/AdminApis.js:dispatchAdminOperation()`。
`npm run api` の送り先は `single-tenants.json` の配置 (`--tenant <名前>`、省略時は検証用を除く最初の配置)。

```bash
npm run api -- systemDiagnosis                # DB接続・設定・デプロイ状態の健康診断
npm run api -- getLogs --limit 20             # セキュリティログ
npm run api -- getUserConfig --userId X       # ボードの設定
npm run api -- setUserConfig --userId X --patch '{...}'  # config パッチ適用
npm run api -- previewBoard --userId X        # viewer 視点の公開ボードプレビュー
```

**いつ使うか:**

- 配置の後 → `systemDiagnosis` で壊れていないか確認
- エラー調査時 → `getLogs` でセキュリティログ、`perfMetrics` でボトルネック特定
- 設定変更後 → `listProperties` で反映を確認

**実行ログ**: 先生ごとの配置は既定の GCP project で動くので、Cloud Logging を CLI から読めない。
Apps Script エディタの「実行数」で見る。フロントエンドのエラーも `reportClientError` で同じ場所に出る。

**全コマンド一覧**: [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md) §1-2。

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

- **`doGet(e)`** (main.js): Routes `?mode=` parameter to HTML templates via `HtmlService`. Handler table `DO_GET_HANDLERS` (main.js) is the single source of truth: `admin` / `view` / `review` / `main` (既定・未知の mode も `main` にフォールバック)。ログイン・初期設定・アプリ設定の画面は無い: 先生専用の配置では先生がトップを開くと `processLoginAction` で登録して管理パネルを直接返し、初回のプロパティと DB シートは `ensureSingleTenantBootstrap_` がそろえる。停止/再開などの運用操作は Admin API で行う
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
- **DatabaseCore.js**: users シートの CRUD と、スプレッドシートを開くときのアクセス判定 (`checkSpreadsheetAccess_` → `SpreadsheetApp.openById`)
- **SystemController.js**: System diagnostics, performance monitoring, deployment management (`publishApp` with etag conflict detection)
- **helpers.js**: Response builders (`createSuccessResponse`, `createErrorResponse`), `getCachedProperty_` (30s TTL cache for PropertiesService)
- **validators.js**: All input validation functions

### アクセス (誰が何を開けるか)

先生として実行するので、読み書きはすべて `SpreadsheetApp.openById` で直接行う (実行者 = 先生 = すべてのシートの持ち主)。
サービスアカウント・Sheets API の直叩き・共有設定は使わない。「誰に見せてよいか」だけを `checkSpreadsheetAccess_`
(DatabaseCore.js) が決める:

| 呼び出し元 | 公開中のボードが指すシート | 非公開のボードのシート | どのボードにも無いシート |
| ---------- | -------------------------- | ---------------------- | ------------------------ |
| **先生 (管理者)** | 可 | 可 | 可 (作ったばかりのフォームの回答先など) |
| **児童・ほかの先生** | 可 | 不可 | 不可 |

- DB (先生のスプレッドシート) も同じ扱い。授業モードの回答シートが全部入るので、児童が開けるのは
  授業モードの授業がボードにつながって公開されている間だけ。users / lessons は `openDatabase` で直接読む

- 判定は 60 秒 cache (不許可は 5 秒)。公開状態が変わったら `invalidateAccessDecisionCache_` で版を上げて古くする
- ほかの先生 (共同編集者) はボード SS の編集者。`isBoardCollaborator_` が DriveApp で編集者を読む (AccessControl.js)
- 同時に動かせる処理は先生 1 人あたり 30 本 (Apps Script の上限)。2026-10-29 に 60 本同時でもエラーなし

**設計詳細**: [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)。

### Theme / Color アーキテクチャ

ダークモード既定 + ライトモード対応。全 CSS は **semantic theme token** 経由で記述する。

**新規 CSS の Do/Don't**:

- ✅ **Do**: `var(--theme-text-primary)` / `var(--theme-bg-surface)` / `var(--theme-border-subtle)` (文脈依存の色)
- ✅ **Do**: brand identity / status は `var(--brand-like)` / `var(--status-success)` (theme 非依存)
- ✅ **Do**: 新規 UI は semantic primitives (`.card` / `.modal-card` / `.box`) を使う。inline Tailwind chain を新規追加しない
- ❌ **Don't**: ハードコード hex/rgba (`color: #94a3b8` / `background: rgba(255,255,255,0.1)`)
- ❌ **Don't**: Tailwind `dark:` prefix (CI で機械検証・禁止)。light/dark は `--theme-*` token + `body.theme-light` で表現

**色を変更したいとき**: [UnifiedStyles.css.html](src/UnifiedStyles.css.html) の `:root` (dark) と `body.theme-light` (light) の token を書き換えるだけで全体反映。

**寸法も token 経由で書く**: 角丸・文字サイズ・余白は生値を書かず `var(--radius-md)` /
`var(--font-size-sm)` / `var(--space-4)` を使う。スケールと同値の生値は `check:tokens` が検出する
(`npm run tokenize:dim` で一括変換)。値を散らすとスケールを一括で変えられなくなる。

**検証**: `npm run theme:contrast` (WCAG AA) / `theme:matrix` (全 token + contrast) / `theme:perfect` (統合ゲート) / `check:tokens` (寸法 token)。

**パレット設計哲学・token 一覧・primitives 表・status box token・themeManager API・保守 CLI 詳細**: [docs/THEME.md](docs/THEME.md)。

### Cache アーキテクチャ (3 層、 意図的分離)

3 つの層はそれぞれ異なる用途。統合 facade を作らない (semantic clarity を失う + cost 非対称性が見えなくなる)。

| 層 | TTL | 用途 | 注意 |
| -- | --- | ---- | ---- |
| `getCachedProperty_(key)` | **1 回の実行の中だけ** (メモリ、 50 件) | **config** (ADMIN_EMAIL, DATABASE_SPREADSHEET_ID 等) | ADMIN_EMAIL / DATABASE_SPREADSHEET_ID は `Deployment.js` の焼き込みを返し、プロパティを読まない。 ほかの鍵は実行ごとに 1 回プロパティを読む (google.script.run の 1 呼び出し = 1 実行で、メモリは持ち越さない)。 write は `setCachedProperty_` 経由 |
| `CacheService.getScriptCache()` | 最大 6h、 100KB/value | **作業データ** (board data, user objects, アクセス判定, rate-limit counter) | size overrun は silent fail。 必ず `saveToCacheWithSizeCheck_()` 経由 |
| `PropertiesService` | 永続 (no TTL) | **永続的システム設定 / 秘密情報** (ADMIN_API_KEY, DATABASE_SPREADSHEET_ID, ADMIN_EMAIL) | hot path で直読みすると 50-100ms/call。 `getCachedProperty_` 経由を推奨。 例外: `handleSetupApiKeyAction_` 初回ゲートは 30s stale race 回避のため直読み (`main.js` 参照) |

**先生として実行なので、児童 30 人の処理がすべて先生 1 人の 1 日の上限に数えられる** (プロパティの読み書き
50 万回 / 日、同時実行 30)。5 秒おきに呼ばれる処理 (`getNotificationUpdate` など) でプロパティを読まない。
実行をまたいで覚えたい値は CacheService に置く (例: users の cache の版 `user_cache_ver`、ボードの版)。
版は番号を足さず毎回新しい値にする (重なった書き込みで取りこぼさない)。

**判断フロー (新規 cache を書く時):**

1. 永続 secret / 設定 → `PropertiesService.setProperty` (direct write) + 読みは `getCachedProperty_`
2. ボード data / 計算結果 / token 等 (30s〜6h で再計算可) → `CacheService` + `saveToCacheWithSizeCheck_`
3. すでに `getCachedProperty_` で 30s memoize される config への参照 → 何もしない (透過的)

### Frontend

HTML templates use `<?!= include('filename') ?>` for composition. Key patterns:

- `SharedUtilities.html`: Common JS utilities including `escapeHtml()`
- `page.js.html` / `AdminPanel.js.html`: Single-file JS (no physical splitting - intentional)
- Top-level side effects forbidden - `google.script.run`, DOM manipulation, and auto-timers must be inside `init()` only
- **include 対象に scriptlet を書かない** - `include()` は `createHtmlOutputFromFile().getContent()` (単なるファイル連結) であり、テンプレート評価をしない。`<? ?>` を書いても実行されず生テキストとして出る。`check:gas` 項目 11 が機械検証する
- Include order は `scripts/gen-pagehead.js` の `ORDER` が唯一の定義。変更は独立コミットで

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
| サーバ応答のエラー文言取り出し | `serverMessage(res, fallback)` | 同上 |
| HTML escape | `sharedUtilities.security.escapeHtml(s)` | 同上 |
| ローディングオーバーレイ | `setLoading(true, msg)` / `setLoading(false)` | 同上 |
| デバウンス (key 共有可) | `sharedUtilities.debounce.debounce(fn, key, delayMs)` | 同上 |
| ボタンの処理中 UX | `sharedUtilities.buttons.withBusy(btn, asyncFn, { busyText })` | 同上 |

**禁止 / 非推奨パターン:**

- 新規の右上トースト DOM、中央バナー DOM、`#notification-container` を手書きしない (上記 API に集約)
- 新規の native `confirm()` / `alert()` (緊急 fallback 用途を除く) → `modals.confirm/alert` を使う
- `withSuccessHandler/withFailureHandler` の手書きチェーン → `runServer(...)` を使う
- `JSON.parse(...)` を try/catch なしで呼ぶ → `safeJsonParse(...)` を使う
- `res.message || res.error` と手書きする → `serverMessage(res, fallback)` を使う
  (サーバの失敗レスポンスは `message` のみ / `error` のみの 2 系統が混在する。
   書き忘れた 1 か所だけが「エラーなのに何も出ない」になるので読み手を 1 本にする)

### Data Store

先生のスプレッドシートが DB。`users` シート (先生 1 人分) の `configJson` 列にボードの設定、`lessons` / `lesson_responses` に授業とアーカイブ、授業モードの回答シートは「授業名｜番号 フェーズ名」で同じスプレッドシートに入る (see `src/SystemController.js` `USERS_SHEET_HEADERS`)。

アプリが作ったフォーム (`createTemplateForm`) は回答シートの列の並びが決まっているので、作った時点の割り当て
(`templateColumnMapping_`) を使い、推定しない (回答 0 件では尺度の列を数値と見分けられないため)。
回答シートには種類を developer metadata で書き添え、`getColumnAnalysis` が見出しの並びを確かめて使う。

### Visualization Modes

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

M1 (numberline) / M2 (matrix) の設計背景: [docs/SPEC_visualization_modes.md](docs/SPEC_visualization_modes.md)
（2026-05 の初期設計書。wordcloud / pie はその後の追加なので同書には載っていない）。

### 授業モード (native 入力) — 掲示板モードとは独立した縦の経路

テンプレ `dialogue-reconsider-5phase`（考える→出会う→議論する→もう一度考える→ふりかえる）は
`lessonJson.inputMode='native'` で動く。**Form も列推定も通らない**。掲示板モード
（Form + リアクション + StudyQuest 連携）は無変更で並走する。混ぜないこと。

- **フェーズが児童画面の権能を決める**。`phase.screenRole` = `input`/`browse`/`discuss`/`reinput`/`reflect`。
  「議論中は投稿できない」は UI ではなく `submitLessonAnswer` のサーバ検証で保証する。
- **「考える」では他者の行を返さない**（`__maskOthersDuringInputPhase_`）。viewer cache は
  全件を保持し、mask は cache 取得後に viewer ごとに適用する（順序を逆にすると他人の分が cache される）。
- **児童同士に名前を見せない**（native phase は `showNames:false` が既定）。教師は `isOwnBoard` で見える。
- **`answer` と `reason` を同じ列にしない**。`validateMapping` は `numericX/Y` だけ重複を免除する。
  違反すると config 保存が落ち、教師が「開始」を押した瞬間に失敗する（`__nativeColumnMapping_`）。
- 投稿を受け付けないフェーズは**直前の入力フェーズのシート**を見る（自分のシートは空）。
  軸ラベルは授業を通して 1 つの座標系に揃える（● → ★ の比較が成立する条件）。

---

## Must-Follow Rules

### Performance

- **Batch operations always** - `getDataRange().getValues()` not `getRange(i,j).getValue()` in loops (70x difference)
- **Cache PropertiesService** - use `getCachedProperty_()` with 30s TTL (see helpers.js)
- **Minimize external service calls** - in-script JS operations are faster

### レスポンスの形 (サーバ → クライアント)

- **失敗は必ず `createErrorResponse(message, data?, extraFields?)` で作る**。手書きの
  `{ success: false, ... }` を新規に書かない。ビルダーは `message` と `error` の両方を
  立てるので、受け手が「どちらに入っているか」を気にしなくてよくなる。
  - `data` / `sheetName` / `header` 等の追加フィールドは第 2・第 3 引数で保てる
  - 機械コードは `extraFields` に入れる: `createErrorResponse('...', null, { error: 'ETAG_MISMATCH' })`
- **`error` は「このレスポンスは失敗」の意味だけに使う**。`success: true` の応答に
  診断コードを載せたいときは `reason` を使う (`if (res.error)` の誤判定を防ぐ)。
- 唯一の例外は `DatabaseCore.js` の `DB_NOT_CONFIGURED` (理由をコード側に明記済)。
- 読み取り側は `serverMessage(res, fallback)` を使う (上の canonical API 表を参照)。

### 診断可能性 (不具合の原因を追える状態を保つ)

- **握りつぶし catch は「なぜ黙ってよいか」を必ず書く**。`catch (_) {}` だけだと、
  安全な沈黙なのか書き忘れなのかを後から誰も区別できない。
  例: `catch (_) { /* cache 不通は機能的に無害 (次回再計算) */ }`
- **利用者に見える症状が出るのに痕跡が残らない失敗は、必ず log する**。
  例: `bumpBoardDataVersion_` の失敗は他の児童のボードが最大 12 秒古いままになるので
  `console.warn` を残す (処理は続行してよい)。
- **関数名は実体を表す**。`showLoadingOverlay` はオーバーレイを出さず描画最適化を
  止めるだけだったため `suspendRenderOptimization` に改名した。名前が嘘をつくと、
  症状から探した人が必ず遠回りする。

### Security

- **名前が `_` で終わらない関数は、どれも児童の画面から `google.script.run` で呼べる。しかも先生の権限で動く**
  (先生として実行するため)。画面から呼ばせない処理は名前を `_` で終える。画面から呼ぶ管理パネルの処理は
  入口で `teacherOnlyError_()` を確かめる。呼び出し元を確かめない公開関数を足さない
  (2026-10-29: `getCachedProperty` / `setCachedProperty` / 引っ越し / フォーム操作 / 利用者の登録 / cache の書き込み
  (`saveToCacheWithSizeCheck`) / 他の児童の名前を伏せずに読む処理が、児童から呼べた)
- 公開関数は `scripts/public-functions.json` に載っているものだけ (`check:gas` 項目 15)。足すときは、画面から
  呼ぶ入口か、児童が任意の引数で呼んでも書き込み・秘密の読み出し・重い処理が起きないことを確かめてから載せる
- シートに書く児童の文は `__plainTextCell_` を通す (`=` で始まる文を数式にしない。他の児童のメールを外へ送れた)
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
- Include order は `scripts/gen-pagehead.js` の `ORDER` で定義。変更は独立コミットで
- **生成物を手で編集しない** - `src/UtilityStyles.css.html` (`gen:utilities`) と `src/SharedPageHead.html` (`gen:pagehead`) は生成物。元ファイルを直したら再生成する。`deploy-single.js` の preflight と CI が鮮度を検証する

### Board Publish Lifecycle (single source of truth)

`config.isPublished` / `config.publishedAt` を書き換えてよいのは以下 **3 関数のみ**。直接 saveUserConfig_ で書き換えてはいけない（DRY 違反 + テスト不可）：

| 関数 | 場所 | 役割 |
| ---- | ---- | ---- |
| `publishApp(publishConfig)` | `SystemController.js` | 初回公開 / 設定差し替えを伴う公開（allowlist + etag 検証） |
| `republishMyBoard(options?)` | `AdminApis.js` | 所有者が自分のボードを再び公開する（state のみ） |
| `unpublishBoard(targetUserId?, options?)` | `AdminApis.js` | 所有者 or 管理者がボードの公開を終了する |

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

---

## Testing

Tests run in Node.js using `node:test` (no frameworks). GAS globals are stubbed via `vm.createContext()` - see `tests/main.doPost.test.cjs` for the pattern. Each test file loads source files into a sandboxed VM context with mock GAS services (`ContentService`, `Session`, `SpreadsheetApp`, etc.).

---

## clasp + Git

- **Never commit**: `.clasp.json`, `.clasprc.json`, `scripts/single-tenants.json` (scriptId / API キー / 先生ごとの値)
- **Do commit**: `src/**/*.js`, `src/**/*.html`, `src/appsscript.json`
- 配置は `.single-stage/<名前>/` で組み立てて clasp に渡す (`-P . -I .` 必須。付けないと親の `.clasp.json` を拾う)
- File extension: `.js` locally (clasp convention), shows as `.gs` in GAS editor

### Required Script Properties

配置の初回に `ensureDeploymentBootstrap_` が `Deployment.js` の値から入れる (手で入れる必要はない):

| Property | Description |
| -------- | ----------- |
| `ADMIN_EMAIL` | この配置の先生 (管理者) |
| `DATABASE_SPREADSHEET_ID` | 先生のスプレッドシート (DB) |
| `ADMIN_API_KEY` | Admin API の鍵 (初回に `setupApiKey` で決める。`single-tenants.json` に控える) |

---

## 兄弟プロジェクト (StudyQuest との連携)

「みんなの回答ボード」 = 集合可視化の場、 一方「[StudyQuest](file:///Users/ryuma/slide_generator)」 (= slide_generator) は個別マイプラン学習 (3 touch 構造) の場。 両者は **データだけで疎結合に連携**:

- 連携責任: **StudyQuest 側に Answer Board Bridge 機能を実装** (回答ボード側はゼロ変更)
- 状態: **構想のみ（未実装）**。StudyQuest 側に bridge 実装も設計書もまだ無い（`docs/answerboard-bridge.md` は不在）。着手するときは [StudyQuest docs/](file:///Users/ryuma/slide_generator/docs/) に設計書を置いてここからリンクし直す
- 動作: StudyQuest タッチ 3 (ふりかえり) → prefilled Google Form URL 生成 → 回答ボード Form に submit → matrix/wordcloud で集合可視化
- 文献的根拠: 個別最適な学びの継続困難 5 大要因 (教師負担集中・実践知の継承困難・進度差対応・学力低下・標準化テンション) → アプリは「1 単元 8 時間中 7-10% だけ介入」 の minimum viable 個別最適を堅持
- 回答ボード側で新規 API / OAuth scope 追加なし。 Google Form の prefilled URL (`?usp=pp_url&entry.XXX=...`) が連携プロトコル
