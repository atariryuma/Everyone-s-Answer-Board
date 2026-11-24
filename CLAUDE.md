# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

**Project**: Google Apps Script Web Application for organization-internal use (collaborative answer board for Google Forms responses)
**Stack**: Zero-dependency, direct GAS API calls, V8 runtime, Tailwind CSS frontend
**Language**: Comments and UI strings are in Japanese; code identifiers are in English

このファイルにあるのは**どの仕事にも要る決まり**だけ。分野ごとの規則は `.claude/rules/` にある (下の「分野の規則」)。

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
npm run lint                             # CLAUDE.md ルール違反検出（setTimeout/innerHTML+var/getValue ループ等）
npm run lint:errors                      # error severity のみ（CI 用）
npm run lint:eslint                      # ESLint（バグ予防レベル）
npm run test:coverage                    # node --test のカバレッジ
npm run deploy                           # 全員の配置を更新 (URL はそのまま。deploy-single.js deploy --all)
npm run deploy:single -- deploy --name <名前>   # 1 人分だけ更新
npm run deploy:single -- create --name <名前> --adminEmail <先生のメール>   # 先生を足す
npm run verify:e2e -- <検証用の名前>    # 検証用の配置で画面の機能を端から端まで動かす (75 項目)
npm run api -- systemDiagnosis           # 配置の後に壊れていないか確認 (Admin API)
```

**配置**: 先生ごとに 1 つ、「先生として実行」(`USER_DEPLOYING`) で配置する。**配置はいつも `--no-ops-report` を付ける**。
詳しくは `.claude/rules/deploy.md`。

**Dev workflow**: Edit locally → `npm test` → `npm run deploy:single -- deploy --name <検証用の名前>` →
`npm run verify:e2e -- <検証用の名前>` → `npm run deploy` → Git commit → `git push`

**CI** (.github/workflows/ci.yml): Syntax check + `npm run lint:errors` + `npm test`（品質ゲートのみ。配置はローカルで行う）

pre-commit hook は `npm install` 時に自動有効化される（`postinstall` で `git config core.hooksPath .githooks`）。コミット前に syntax + lint + test が自動実行される。緊急時は `git commit --no-verify`。

---

## 進めかた ―― 少ない人数で、安い確かめから (2026-11-24)

WILD LEAGUE (`wild-league`) と exomemory で 2026-11-24 に決めた進めかたを、同じ日にこのリポジトリにも入れた
(元の文は `wild-league` の `CLAUDE.md` の同じ名前の節。外部記憶にも要点がある)。

1. **同時に動かすのは 3 つまで。** 分野ごとに常設の担当を置かない。仕事が済んだセッションはアーカイブし、
   次の仕事は新しいセッションに短い頼み (何を・どのファイル・終わりの条件) で渡す。
2. **人にしか判断できないもの (画面の見た目・言葉・授業での使い勝手) は、いちばん安い見本を先に出す。**
   画面の写し 1 枚・文の案への返事をもらう前に、仕組み・テストを作らない。
3. **頼む前に方向を決める。** 分からない所は、動かす前に 1〜2 問だけ聞く。
4. **係は中継しない。** 数回のツール呼び出しで済むことは自分でやる。別のセッションに渡すのは、
   1 つの仕事として独り立ちするもの (ブランチや配置で終わるもの) だけ。
5. **AI にできない評価に容量を使わない。** 測れること (テスト・lint・`systemDiagnosis`・`verify:e2e`) は機械で測る。
   「100 点を」のような際限のない基準には、客観的な指標を示して了承を得る。
6. **機械的な仕事は Codex へ** (広く読む調べもの・仕様のはっきりした実装・テスト書き。目安は 30 行を超えるか複数のファイル)。
   手順は `codex-task` スキル (無い PC では Codex に渡さず進める)。Codex はコミットしないので、差分を読んで取り込む。
   書いた側と確かめる側は別の会社にする (Claude が書いたら `codex-review`)。
   **Codex に渡さないもの**: 配置 (`npm run deploy*`)・Admin API (`npm run api`)・`verify:e2e`・
   `scripts/single-tenants.json` / `.single-stage/` / `.clasp*.json` (先生のメール・API キー・scriptId が入る)・
   クラスルームへの投稿・本番のスプレッドシートの中身。Codex に渡すのは手元のコードとテスト (`npm test`) だけ。
7. **まとまった変更を配置したら、掃除を 1 回。** 使われなくなったコード・設定・テスト・文書を洗い出して消す。
   消してよいか迷うもの (テストからしか呼ばれない関数・データの欄) は、消さずに一覧にしてオーナーに聞く。消したら `npm test` と `npm run lint:errors` を通す。

**直す前に確かめること:**

- 「統一しよう」「重複を消そう」と思ったら、まず実装を読んで本当に同じ仕事かを確かめる。名前が似ているだけの
  別物を統合すると壊れる (2026-08-30 の全体の統一で、重複に見えた 3 組のうち 2 組が別物だった)。
  使用回数 0 も、使う先生が少ないうちは削除の根拠にしない。
- 「全体をきれいに」と頼まれたら、孤立した関数の grep だけで済ませず、クローン検出で重複を実測する。
- レビューで最優先は correctness / scope (未宣言変数、関数の外の参照、catch の変数名の誤り)。単一のグローバル
  スコープで型も bundler も無く、エラーの経路でだけ落ちるバグが残りやすい。ファイルの行数では減点しない。
- **入口・許可の判定・`appsscript.json` を変える配置は、授業のある日に入れない** (理由と確かめ方は `.claude/rules/deploy.md`)。

---

## 分野の規則 ―― 触る前に、その分野のファイルを読む

- Claude Code は、下の「読みこまれるとき」に当たるファイルを開く・書くと、その分野のファイルを自動で読みこむ。
- ⚠ **Grep や Glob で見つけただけでは、読みこまれるとは限らない**。直す前に、対象のファイルを 1 度 Read で開くこと。
- Codex などほかの AI は自動では読まない。触る分野のファイルを、自分で開いてから始める (Codex に渡す指示にも、読むファイルを書く)。
- 分野をまたぐ仕事は両方読む。迷ったら読む。
- 規則を足す・直すときは、その分野のファイルへ。このファイルには「どの仕事にも要ること」だけを足す。

| 分野 | ファイル | 読みこまれるとき (主なもの) |
| --- | --- | --- |
| 配置・Admin API・入口・clasp | `.claude/rules/deploy.md` | `scripts/deploy-single.js` `scripts/admin-api.js` `scripts/verify/**` `src/appsscript.json` `src/main.js` `src/AdminApis.js` `src/SystemController.js` |
| 画面 (テーマ・include・共通 API) | `.claude/rules/frontend.md` | `src/**/*.html` `scripts/gen-*.js` `scripts/theme-*.js` |
| ボード (データ・表示モード・公開) | `.claude/rules/board.md` | `src/validators.js` `src/ConfigService.js` `src/DataService.js` `src/DataApis.js` `src/ColumnMappingService.js` |
| 授業モード (native 入力) | `.claude/rules/lesson.md` | `src/LessonService.js` `src/LessonWorkspace.html` `src/DataApis.js` |

詳しい説明は `docs/` (仕組み [ARCHITECTURE.md](docs/ARCHITECTURE.md)・テーマ [THEME.md](docs/THEME.md)・コマンド [DEVELOPMENT.md](docs/DEVELOPMENT.md))。
StudyQuest との連携 (構想のみ) も ARCHITECTURE.md の末尾。

---

## Architecture

### Request Flow

- **`doGet(e)`** (main.js): Routes `?mode=` parameter to HTML templates via `HtmlService`. Handler table `DO_GET_HANDLERS` (main.js) is the single source of truth: `admin` / `view` / `review` / `main` (既定・未知の mode も `main` にフォールバック)。ログイン・初期設定・アプリ設定の画面は無い: 先生専用の配置では先生がトップを開くと `processLoginAction` で登録して管理パネルを直接返し、初回のプロパティと DB シートは `ensureDeploymentBootstrap_` がそろえる。停止/再開などの運用操作は Admin API で行う
- **`doPost(e)`** (main.js): Routes JSON `action` field to handler functions. Actions are allowlist-managed; each must have input validation.
- Both entry points authenticate via `Session.getActiveUser().getEmail()`

### Service Layer Pattern

```text
main.js (doGet/doPost routing, auth)
  → *Apis.js (API endpoints, parameter extraction, response formatting)
    → *Service.js (business logic, data transformation)
      → DatabaseCore.js (users の読み書き, シートを開くときのアクセス判定)
```

- **Apis files** group endpoints by domain (AdminApis, UserApis, DataApis)
- **Service files** contain pure business logic (ConfigService, DataService, UserService, ReactionService)
- **DatabaseCore.js**: users シートの CRUD と、スプレッドシートを開くときのアクセス判定 (`checkSpreadsheetAccess_` → `SpreadsheetApp.openById`)
- **SystemController.js**: System diagnostics, performance monitoring, deployment management (`publishApp` with etag conflict detection)
- **helpers.js**: Response builders (`createSuccessResponse`, `createErrorResponse`), `getCachedProperty_` (PropertiesService の値を 1 回の実行の中だけ覚える)
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

### Cache アーキテクチャ (3 層、 意図的分離)

3 つの層はそれぞれ異なる用途。統合 facade を作らない (semantic clarity を失う + cost 非対称性が見えなくなる)。

| 層 | TTL | 用途 | 注意 |
| -- | --- | ---- | ---- |
| `getCachedProperty_(key)` | **1 回の実行の中だけ** (メモリ、 50 件) | **config** (ADMIN_EMAIL, DATABASE_SPREADSHEET_ID 等) | ADMIN_EMAIL / DATABASE_SPREADSHEET_ID は `Deployment.js` の焼き込みを返し、プロパティを読まない。 ほかの鍵は実行ごとに 1 回プロパティを読む (google.script.run の 1 呼び出し = 1 実行で、メモリは持ち越さない)。 write は `setCachedProperty_` 経由 |
| `CacheService.getScriptCache()` | 最大 6h、 100KB/value | **作業データ** (board data, user objects, アクセス判定, rate-limit counter) | size overrun は silent fail。 必ず `saveToCacheWithSizeCheck_()` 経由 |
| `PropertiesService` | 永続 (no TTL) | **永続的システム設定 / 秘密情報** (ADMIN_API_KEY, DATABASE_SPREADSHEET_ID, ADMIN_EMAIL) | hot path で直読みすると 50-100ms/call。 `getCachedProperty_` 経由を推奨。 例外: `handleSetupApiKeyAction_` 初回ゲートは古い値で二重設定しないよう直読み (`main.js` 参照) |

**先生として実行なので、児童 30 人の処理がすべて先生 1 人の 1 日の上限に数えられる** (プロパティの読み書き
50 万回 / 日、同時実行 30)。5 秒おきに呼ばれる処理 (`getNotificationUpdate` など) でプロパティを読まない。
実行をまたいで覚えたい値は CacheService に置く (例: users の cache の版 `user_cache_ver`、ボードの版)。
版は番号を足さず毎回新しい値にする (重なった書き込みで取りこぼさない)。

**判断フロー (新規 cache を書く時):**

1. 永続 secret / 設定 → `PropertiesService.setProperty` (direct write) + 読みは `getCachedProperty_`
2. ボード data / 計算結果 / token 等 (30s〜6h で再計算可) → `CacheService` + `saveToCacheWithSizeCheck_`
3. すでに `getCachedProperty_` で読む config への参照 → 何もしない (透過的)

---

## Must-Follow Rules

### Performance

- **Batch operations always** - `getDataRange().getValues()` not `getRange(i,j).getValue()` in loops (70x difference)
- **Cache PropertiesService** - `getCachedProperty_()` を通す (1 回の実行の中で覚える。ADMIN_EMAIL / DB は焼き込み。see helpers.js)
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
- 読み取り側は `serverMessage(res, fallback)` を使う (`.claude/rules/frontend.md` の canonical API 表を参照)。

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

- `doPost` actions are allowlist-managed; adding an action requires adding its input validation simultaneously
- `config.isPublished` / `config.publishedAt` を書き換えてよいのは `publishApp` / `republishMyBoard` / `unpublishBoard` の 3 関数だけ (`.claude/rules/board.md`)
- Include order と生成物 (`UtilityStyles.css.html` / `SharedPageHead.html`) の決まりは `.claude/rules/frontend.md`

---

## Testing

Tests run in Node.js using `node:test` (no frameworks). GAS globals are stubbed via `vm.createContext()` - see `tests/main.doPost.test.cjs` for the pattern. Each test file loads source files into a sandboxed VM context with mock GAS services (`ContentService`, `Session`, `SpreadsheetApp`, etc.).

---

## clasp + Git

- **Never commit**: `.clasp.json`, `.clasprc.json`, `scripts/single-tenants.json` (scriptId / API キー / 先生ごとの値)、`scripts/ops-report.local.json` (報告先と鍵)
- **Do commit**: `src/**/*.js`, `src/**/*.html`, `src/appsscript.json`, `.claude/rules/*.md`
- clasp の使い方と Script Properties は `.claude/rules/deploy.md`
