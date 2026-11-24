---
paths:
  - "src/**/*.html"
  - "scripts/gen-*.js"
  - "scripts/theme-*.js"
  - "scripts/theme-colors.json"
  - "scripts/tokenize-dimensions.js"
  - "scripts/lint.js"
  - "docs/THEME.md"
  - "tests/page.*"
  - "tests/shared.*"
  - "tests/theme.*"
  - "tests/lib.theme-core.*"
  - "tests/viz.*"
---

# 画面 (HTML / CSS / フロントの JS)

## Theme / Color アーキテクチャ

ダークモード既定 + ライトモード対応。全 CSS は **semantic theme token** 経由で記述する。

**新規 CSS の Do/Don't**:

- ✅ **Do**: `var(--theme-text-primary)` / `var(--theme-bg-surface)` / `var(--theme-border-subtle)` (文脈依存の色)
- ✅ **Do**: brand identity / status は `var(--brand-like)` / `var(--status-success)` (theme 非依存)
- ✅ **Do**: 新規 UI は semantic primitives (`.card` / `.modal-card` / `.box`) を使う。inline Tailwind chain を新規追加しない
- ❌ **Don't**: ハードコード hex/rgba (`color: #94a3b8` / `background: rgba(255,255,255,0.1)`)
- ❌ **Don't**: Tailwind `dark:` prefix (CI で機械検証・禁止)。light/dark は `--theme-*` token + `body.theme-light` で表現

**色を変更したいとき**: [UnifiedStyles.css.html](../../src/UnifiedStyles.css.html) の `:root` (dark) と `body.theme-light` (light) の token を書き換えるだけで全体反映。

**寸法も token 経由で書く**: 角丸・文字サイズ・余白は生値を書かず `var(--radius-md)` /
`var(--font-size-sm)` / `var(--space-4)` を使う。スケールと同値の生値は `check:tokens` が検出する
(`npm run tokenize:dim` で一括変換)。値を散らすとスケールを一括で変えられなくなる。

**検証**: `npm run theme:contrast` (WCAG AA) / `theme:matrix` (全 token + contrast) / `theme:perfect` (統合ゲート) / `check:tokens` (寸法 token)。

**パレット設計哲学・token 一覧・primitives 表・status box token・themeManager API・保守 CLI 詳細**: [docs/THEME.md](../../docs/THEME.md)。

## Frontend

HTML templates use `<?!= include('filename') ?>` for composition. Key patterns:

- `SharedUtilities.html`: Common JS utilities including `escapeHtml()`
- `page.js.html` / `AdminPanel.js.html`: Single-file JS (no physical splitting - intentional)
- Top-level side effects forbidden - `google.script.run`, DOM manipulation, and auto-timers must be inside `init()` only
- **include 対象に scriptlet を書かない** - `include()` は `createHtmlOutputFromFile().getContent()` (単なるファイル連結) であり、テンプレート評価をしない。`<? ?>` を書いても実行されず生テキストとして出る。`check:gas` 項目 11 が機械検証する
- Include order は `scripts/gen-pagehead.js` の `ORDER` が唯一の定義。変更は独立コミットで
- **生成物を手で編集しない** - `src/UtilityStyles.css.html` (`gen:utilities`) と `src/SharedPageHead.html` (`gen:pagehead`) は生成物。元ファイルを直したら再生成する。`deploy-single.js` の preflight と CI が鮮度を検証する

## Canonical Frontend APIs (Use these, don't reinvent)

新規コードを書くときは以下の集約 API を使う。ad-hoc DOM 操作や native `confirm/alert` を新規に書かない。

| 用途 | API | 場所 |
| ---- | --- | ---- |
| 右上トースト (success/error/warning/info) | `notifications.success/error/warning/info(msg, opts?)` | [SharedUtilities.html](../../src/SharedUtilities.html) |
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
