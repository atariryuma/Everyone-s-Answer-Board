# Theme / Color アーキテクチャ

ダークモード既定 + ライトモード移行準備済。全 CSS は **semantic theme token** 経由で記述する。

CLAUDE.md には要点 (Do/Don't + 検証コマンド) のみを置き、詳細はこのドキュメントに集約している。

---

## パレット設計哲学

Tokyo Night atmospheric dark + 業界定番 light の hybrid。

| 層 | Dark | Light | 用途 |
| ---- | ---- | ----- | ---- |
| `--theme-bg-base` | `#1a1b26` Tokyo Night | `#f1f5f9` slate-100 | ページ最奥 |
| `--theme-bg-surface` | `#252738` (+1 step) | `#ffffff` 純白 | カード・panel (page bg から明確に浮く) |
| `--theme-bg-elevated` | `#2e3046` (+2 step) | `#ffffff` 純白 + shadow | modal・popover |
| `--theme-text-primary` | `#e2e8f0` slate-200 (12:1 AAA) | `#0f172a` slate-900 (16:1 AAA) | 本文 |

**設計判断**: dark mode は **Tokyo Night の atmospheric warmth** を base に維持しつつ、surface / elevated を **solid colors で明示的に lighter step** (旧 alpha 値廃止)。light mode は **GitHub/Linear/Stripe 方式**の off-white page + 純白カード。

---

## 色を変更したいとき

例: 「もう少し warm な印象」「アクセントを purple に」

1. [UnifiedStyles.css.html](../src/UnifiedStyles.css.html) の `:root` (dark) と `body.theme-light` (light) の `--theme-bg-base` / `--theme-text-primary` 等を書き換えるだけ
2. 他の CSS / HTML は全部 `var(--theme-*)` 参照なので **token を 1 箇所変えるだけで全体反映**
3. `npm run theme:contrast` で WCAG AA を機械検証
4. `npm run theme:perfect` で全軸ゲート確認

例: warm な light mode に変更する場合

```css
body.theme-light {
  --theme-bg-base: #fafaf9;  /* stone-50 (warm beige) */
  --theme-bg-surface: #fffdf8; /* warm off-white */
}
```

これだけで全カード・全テキスト・全 modal が warm tone になる。

---

## 3 種類のカラートークン

| カテゴリ | 例 | テーマ切替 | 用途 |
| -------- | -- | --------- | ---- |
| `--theme-*` | `--theme-bg-base`, `--theme-text-primary`, `--theme-border-subtle` | **書き換わる** | 文脈依存の色 (背景・テキスト・境界線・オーバーレイ) |
| `--brand-*` | `--brand-primary`, `--brand-like`, `--brand-curious` | 不変 | ブランドアイデンティティ (アプリ色・リアクション色) |
| `--status-*` | `--status-success`, `--status-error`, `--status-warning` | 不変 | システムフィードバック色 |

---

## UI primitive token — color 以外も Tailwind scale 統一

- Typography: `--font-size-{xs,sm,base,lg,xl,2xl,3xl,4xl}` = Tailwind text-xs..4xl と完全一致
- Font weight: `--font-weight-{normal,medium,semibold,bold}` = 400/500/600/700
- Line height: `--leading-{tight,snug,normal,relaxed,loose}` = 1.25/1.375/1.5/1.625/1.75
- Spacing: `--space-{0,1,2,3,4,5,6,8,10,12,16,20}` = Tailwind p-/m- スケール
- Radius: `--radius-{none,sm,md,lg,xl,full}` = 0/0.25/0.5/0.75/1rem/9999px
- Shadow: `--shadow-{sm,md,lg}` = 影レイヤー 3 段
- Transition: `--transition-{fast,normal,slow}` = 0.1/0.2/0.3s
- Z-index: `--z-{base,content,elevated,overlay,dropdown,modal,notification,tooltip,critical}` = 0..9999 9 段

非標準値は段階的に最近接 token へ移行 (CLI: [scripts/theme-normalize-typography.js](../scripts/theme-normalize-typography.js))。

---

## Semantic theme tokens (`--theme-*`)

詳細は [UnifiedStyles.css.html](../src/UnifiedStyles.css.html) `:root` ブロック:

- 背景: `--theme-bg-base` / `--theme-bg-surface` / `--theme-bg-elevated` / `--theme-bg-overlay`
- カード階層: `--theme-card-1` (軽い elevation) / `--theme-card-2` (deeper) / `--theme-card-3` (solid panel)
- テキスト: `--theme-text-primary` / `--theme-text-secondary` / `--theme-text-muted` / `--theme-text-inverse`
- 境界線: `--theme-border-subtle` / `--theme-border-normal` / `--theme-border-strong`
- 半透明: `--theme-overlay-white-05/10/20/30` / `--theme-overlay-black-20/60/90` / `--theme-overlay-surface-glass`
- インジケーター: `--theme-spinner` / `--theme-accent-cyan` / `--theme-accent-cyan-soft`

**ライトモード適用** ([UnifiedStyles.css.html](../src/UnifiedStyles.css.html)):

```css
body.theme-light, :root[data-theme="light"] {
  --theme-bg-base: #f8fafc;
  --theme-text-primary: #1e293b;
  /* ... 全 token を上書き */
}
```

---

## Tailwind 統合

[SharedTailwindConfig.html](../src/SharedTailwindConfig.html):

- `darkMode: 'class'` 設定済 (body の `theme-dark`/`theme-light` クラスに応答)
- `theme.extend.colors` で CSS 変数を Tailwind utility class に公開済
- 例: `class="bg-theme-base text-theme-primary border-theme-subtle"` で書ける
- 既存 `bg-gray-800` 等のハードコード Tailwind class は **段階的に** `bg-theme-surface` 等へ移行

---

## FOUC 防止 + color-scheme + transition

業界標準ベストプラクティスを導入済:

- [SharedThemeBoot.html](../src/SharedThemeBoot.html): `<head>` 内で paint 前に `<html>` へ theme class + `color-scheme` を同期付与 → ライト保存ユーザがダーク背景フラッシュを見ない
- `color-scheme: dark` / `color-scheme: light`: ネイティブ scrollbar / form 部品 / canvas を OS dark/light モード追従させる (Tailwind/shadcn 標準)
- body の `transition: background-color 200ms, color 200ms`: テーマ切替時に主要要素 (`.glass-panel`, `.answer-card`, `.modern-select`, `.eab-*` 等) がスムーズに色変化 + `prefers-reduced-motion` で OFF
- 参考: [natebal.com Best Practices](https://natebal.com/best-practices-for-dark-mode/), [shadcn dark mode](https://ui.shadcn.com/docs/dark-mode), [astro-fouc-killer](https://github.com/AVGVSTVS96/astro-fouc-killer)

---

## themeManager API

[SharedUtilities.html](../src/SharedUtilities.html):

```js
themeManager.get()          // 'dark' | 'light' | 'auto' (現在 localStorage 値)
themeManager.resolved()     // 'dark' | 'light' (auto なら OS 反映後)
themeManager.set('light')   // 切替 + 永続化 (localStorage 'app-theme')
themeManager.toggle()       // dark ↔ light
themeManager.subscribe(fn)  // 切替時 callback (unsub 関数を返す)
```

- SharedUtilities ロード時に `init()` が自動実行 (FOUC 防止)。body に `theme-dark`/`theme-light` クラスを付与。
- `auto` モードでは `@media (prefers-color-scheme)` の変化を監視し追従。

---

## 新規 CSS を書く時の Do/Don't

- ✅ **Do**: `color: var(--theme-text-primary)` / `background: var(--theme-bg-surface)` / `border: 1px solid var(--theme-border-subtle)`
- ✅ **Do**: brand identity / status は `var(--brand-like)` / `var(--status-success)` (theme 非依存)
- ❌ **Don't**: ハードコード hex/rgba (`color: #94a3b8` / `background: rgba(255,255,255,0.1)`)
- ❌ **Don't**: 旧 `--brand-text-muted` (定義不在、fallback `#94a3b8` で動いていた) → 必ず `--theme-text-muted` を使う

---

## 保守用 CLI (色トラブル / リファクタの起点)

```bash
npm run theme:matrix      # 全 token (dark/light 値) + 主要 contrast マトリクス + hardcoded 集計
npm run theme:uncovered   # body.theme-light 上書きの無い hardcoded を一覧 (修正 worklist)
npm run theme:tokenize:dry # slate/gray rgba を token に置換するプレビュー
npm run theme:tokenize    # 実行 (テスト + theme:contrast で検証推奨)
npm run theme:contrast    # WCAG AA 全 pair 検証 (≥4.5 本文 / ≥3 アイコン)
npm run theme:verify      # 統合ゲート: 12 軸 / 120 点満点 / CI 用
```

`theme:matrix` の出力ロジック:

- token 抽出: `:root { }` と `body.theme-light { }` を AST 風に解析、var(--x) 参照は再帰解決
- contrast: 12 種類の text/accent/status × 3 種類の bg = 72 pair を alpha-blend 後計算
- hardcoded スキャン: brand-identity (cyan/amber/purple/etc.) と `theme:exempt` コメントは exempt

### 移行ツールの状況 (2026-07-05 時点)

移行の apply モードは概ね完了しているが、**多くは dry-run が regression gate の現役依存**なので
安易に削除しないこと。`theme:perfect` は内部で `theme-matrix.js --uncovered` (axis 1) /
`theme-pair-tailwind.js --dry-run` (axis 2) / `theme-contrast.js` (axis 6) / `theme-verify.js`
(axis 7) を spawn する ([scripts/theme-perfect.js](../scripts/theme-perfect.js))。

- **`theme:pair`** (theme-pair-tailwind.js): apply 移行は完了 (残 0 件) **だが `--dry-run` は
  `theme:perfect` axis 2「Tailwind unpaired = 0」の regression gate が実行する現役依存**。削除不可。
- **`theme:tokenize`** (theme-tokenize.js): 残 1 件 (page.viz.css.html の `rgba(148,163,184,0.5)`
  → `--theme-border-strong`)。gate 非依存だが、将来の hardcoded→token 移行に再利用できるので残置。
- **`theme-normalize-typography.js`**: npm script も gate 依存も無い純粋な一度きりツール。
  唯一の削除候補だが、削除前に移行残がゼロか確認すること。
- **`theme:uncovered`**: 43 件残るが **actionable な theme-bleed は 0 件** (残りは brand-identity
  色/shadow で exempt)。light/dark 移行は機能的に完了。

共通ロジック (色パース / token 抽出 / contrast 計算) を `scripts/lib/theme-core.js` へ抽出する
重複排除は未実施 (出力完全一致の検証コストが高いため別タスク)。

---

## 色を追加 / 変更したいとき

1. 既存 token で済むなら使う (基本ルート)
2. 新しい semantic 階層が必要なら `:root` + `body.theme-light` の両方に追加 (片方だけだと bleed)
3. brand identity (アプリ色・リアクション色) なら theme 非依存で OK (両モード同色のまま)
4. 完了したら `npm run theme:matrix` で全 pair の contrast を確認 → AA fail があれば調整

---

## Semantic Primitives — 新規 UI はこれを使う

旧来は `class="glass-panel rounded-xl p-6 max-w-md w-full mx-4"` のような Tailwind utility chain を毎回繰り返していたが、semantic primitive に統合した。新規 UI は以下を使い、既存の inline chain は新規追加しない。

| primitive | 用途 | 旧パターン |
| --------- | ---- | --------- |
| `.card` | 縦積みコンテンツ単位 (設定パネル等) | `class="card glass-panel rounded-xl p-6"` (3 重 alias) |
| `.card-compact` | padding を縮めた card | `class="info-card glass-panel rounded-xl p-4"` |
| `.card-hero` | primary entry point (cyan gradient + accent border) | `class="card glass-panel rounded-xl p-6 lesson-hero-card"` |
| `.modal-card` | 確認 dialog (max-w-md デフォルト) | `class="glass-panel rounded-xl p-6 max-w-md w-full mx-4"` |
| `.modal-card-md` | form modal (max-w-2xl) | `+ max-w-2xl` |
| `.modal-card-lg` | content modal (max-w-5xl) | `+ max-w-5xl` |
| `.modal-card-prominent` | 強調モーダル (cyan border + shadow-2xl) | `+ shadow-2xl border-2 border-cyan-400/80` |
| `.box` / `.box-row` | 親 card 内の step-up inner box (single source of truth) | `.config-item` 旧 alias |

これらは [UnifiedStyles.css.html](../src/UnifiedStyles.css.html) 内の "Card Primitives" / "Modal Primitives" / "INNER BOX" セクションに定義。token (`--theme-bg-surface`, `--radius-xl`, `--space-6` 等) ベースで書かれているので色/spacing 変更は token を変えれば全 primitive に伝播する。

**Theme bug 防止 token**:

- `--theme-bg-elevated-hover`: `bg-theme-elevated` 系 button の hover 用 (旧 `hover:bg-slate-700` hardcode → light mode で破綻していたバグを解消)。Tailwind utility: `hover:bg-theme-elevated-hover`

---

## Status feedback box token — 旧 `dark:` prefix を全廃

旧来 `bg-yellow-100 dark:bg-yellow-900/20 border-yellow-300 dark:border-yellow-500/30` のような Tailwind `dark:` paired shade を使っていたが、`darkMode: 'class'` 動作の脆弱性 / light/dark 色のペアリング不一致 / theme tokens (`--theme-*`) との二重系統による構造的バグ (light mode で `bg-yellow-100` が出続ける等) が複数発生したため全廃した。

| status | bg (Tailwind utility) | border (Tailwind utility) | 用途 |
| ------ | --------------------- | -------------------------- | ---- |
| success | `bg-theme-status-success-soft` | `border-theme-status-success-strong` | 成功通知 box |
| error | `bg-theme-status-error-soft` | `border-theme-status-error-strong` | エラー / 削除確認 box |
| warning | `bg-theme-status-warning-soft` | `border-theme-status-warning-strong` | 警告 / 心理的安全性 box |
| info | `bg-theme-status-info-soft` | `border-theme-status-info-strong` | 情報通知 box |
| accent (cyan) | `bg-theme-status-accent-soft` | `border-theme-status-accent-strong` | active 状態 (公開中 row 等) |
| detection (violet) | `bg-theme-status-detection-soft` | `border-theme-status-detection-strong` | フォーム自動検出通知 |

対応 text token: `text-theme-status-{success,error,warning,info,detection,highlight}`

**Rule (CI で機械検証)**: HTML / JS で **Tailwind `dark:` prefix は使用禁止**。`theme:perfect` axis 27b が 0 件であることを確認。light/dark variation は必ず `--theme-*` token + `body.theme-light` で表現する。
