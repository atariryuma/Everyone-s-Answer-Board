# 指示書: アプリ本体の品質最適化 (フロント修正・API整合・theme CLI 共通化) 2026-07

**目的**: src/ 約 40,000 行 + scripts/ を 4 並列監査 + 本番実測した結果に基づき、実害のある箇所だけを直す。
**前提**: DX 最適化 (docs/PLAN_dx_optimization_2026-07.md, v2867-v2876 実施済み) の続編。
**最重要の結論**: **バックエンド性能は健全で、性能改善作業は不要**。やるべきは (A) フロントの表示/安全修正、(B) API response の整合性、(C) theme CLI の重複排除の 3 つだけ。

この指示書は 2026-07-05 の監査に基づく。行番号は当時のもの — ズレていたら内容 (コード断片) で再特定すること。

---

## 0. 検証済みの事実 (再調査不要)

### 本番実測 (2026-07-05)

| 項目 | 実測値 |
| ---- | ------ |
| `npm run api -- perfDiagnosis` | overallStatus: **healthy**、detectedIssues: **0** |
| Cloud Logging 直近 72h ERROR (Naha) | **0 件** |
| hot path (getData) の Sheets API 呼び出し | 3-5 回/リクエスト (batch 読み済み)。SA token cache 50min / verify cache 10min が効いている |
| N+1 / batch 違反 | **検出ゼロ** (processBatchData / resolveReactionColumns_ とも適切) |
| PropertiesService 直読み | hot path になし (main.js の初回ゲート直読みは意図的・コメント済み) |

→ **性能リファクタは行わない**。監査で見つかった軽微候補 (board data cache の size 実測、getSheetInfo の headers 再利用) は「様子見」であり、この指示書の作業対象ではない。

### コード監査で確定した修正対象 (手動裏取り済み)

| # | 事実 | 根拠 |
| - | ---- | ---- |
| A1 | `src/page.js.html` にグローバル非推奨 `escape()` が **5 箇所** (L2312, L4462, L4463, L4467, L4487)。ローカル定義は無い (grep 確認済み) | `grep "[^a-zA-Z.]escape(" src/page.js.html` = 5 |
| A2 | native `confirm()` fallback が page.js.html:4236、native `alert()` fallback が page.js.html:4276 に残存 (Canonical API 違反) | grep 確認済み |
| A3 | debug `console.log('✅ Recovery succeeded:...')` が page.js.html:3733 に残存 | grep 確認済み |
| B1 | `createSuccessResponse`/`createErrorResponse` (helpers.js) の使用数が **ConfigService.js = 0 / SystemController.js = 0 / UserApis.js = 0** (main.js は 21、DataApis は 25) | `grep -c` 確認済み |
| C1 | scripts/theme-*.js 8 本 (計 ~2,435 行) に色パース/コントラスト計算/token 抽出/exempt 判定が重複。parseColor・blendOnBg・relativeLuminance・contrastRatio は theme-contrast.js と theme-matrix.js で**完全同一実装**。extractBlockBody は 3 ファイルで同一。exempt 判定 regex は 9 箇所 | 監査エージェントの行番号付きマップ (下記 §C) |
| C2 | brand exempt hex リストが theme-audit.js と theme-matrix.js で**異なる** (同期ずれ = 判定不一致リスク) | hex 出現数 16 vs 24 で相違確認 |

### 監査指摘のうち裏取りで**棄却/修正**されたもの (作業しない)

- ~~「DataApis.js:698 に ad-hoc URL 検証 `url.includes('forms.gle')`」~~ → 実在しない。実体は `DataApis.js:2098 isValidFormUrl()` で、https 強制 + host allowlist + path 検証の**構造化された検証**。validators.js への移設は任意の整理であり必須でない。
- ~~「escape() は単引用符をエスケープしないので XSS」~~ → 過大。グローバル `escape()` は `'` を `%27` にエンコードする。**実害は別**: 非 ASCII (日本語) を `%uXXXX` にエンコードするため、**日本語のクラス名/軸ラベルが文字化け表示される可能性**が高い (L4463/L4467 は軸ラベル `そのまま使う` 等を span に出す箇所)。修正必須なのは変わらないが、理由は「文字化け + 意味論の誤り」。
- ~~バックエンドのデッドコード~~ → **ゼロ** (HTML の google.script.run / runServer / doPost dispatch 全参照を突合済み)。ColumnMappingService の thin wrapper は意図的残置 (再指摘しない)。
- フロントの重複実装 → **なし** (escapeHtml/debounce/runServer は SharedUtilities に集約済み)。polling / イベントリスナー管理 / top-level 副作用も監査済みで健全。

## 0.5 ガードレール — 絶対にやらないこと

1. **性能リファクタをしない** (§0 の実測根拠。healthy なものを触らない)。
2. **page.js.html / AdminPanel.js.html の物理分割をしない**。バックエンド .js の分割・キャッシュ 3 層統合もしない (確立済み方針)。
3. **login.js.html の native alert/confirm (L409/L424/L682) は触らない** — 成功画面 HTML injection 失敗時の last-resort fallback として意図的。
4. **response の shape (フィールド名/構造) を変えない**。B フェーズは「同一 shape を builder 呼び出しに置き換える」作業であり、shape 変更は禁止 (フロント/テストとの契約破壊)。`getUserConfig` の envelope `{success, config, corrupted, userId, message}` のような**内部サービス契約は builder 対象外** — builder の shape と一致しないものは手書きのまま残す。v2865 の「save-shape skip は意図的判断」も尊重。
5. **theme-perfect.js の spawnSync プロセス境界を維持する** (子スクリプトを require に変えない)。
6. scripts/ は **CommonJS** (`module.exports` / `require`)。監査レポートの `export function` 表記は ESM なのでそのまま使わない。
7. コミット規約 `type(scope): 日本語 (vNNNN)`、vNNNN は直前コミット +1。git push は各フェーズ完了時。**デプロイ**: src/ を変更する A・B 完了後に `npm run deploy:all` (両テナント)。permission denied になったらユーザーに実行を依頼して待つ。C は scripts のみなのでデプロイ不要。

---

## フェーズ A: フロントエンド修正 (小・実害あり) 🔴 最優先 — 見積 30-60 分

対象: `src/page.js.html` のみ。

### A-1. 非推奨 `escape()` → `sharedUtilities.security.escapeHtml()` (5 箇所)

- L2312 (クラスフィルタ option)、L4462/L4463/L4467 (回答詳細モーダルの軸ラベル)、L4487 (象限名)。
- 置換前に**現象を確認**: 日本語ラベル (例: 軸ラベル「そのまま使う」) で該当 UI を通ると `%uXXXX` 化けが出るはず。tests での再現が難しければ、置換の正しさをもって代替 (escapeHtml は HTML entity エスケープで日本語を素通しする)。
- 同ファイル内の他の `escapeHtml`/`escapeForTooltip` 使用箇所と同じ呼び方 (`window.sharedUtilities.security.escapeHtml` or 既存の局所 alias) に合わせる。

### A-2. Canonical API fallback の除去 (2 箇所)

- L4236: `: confirm('公開を終了しますか？...')` → 三項 fallback を廃止し `await window.modals.confirm(...)` を無条件使用 (SharedUtilities は常にロード済み)。
- L4276: `else alert(errorMsg)` → 削除し、`console.error` へのログのみに (notifications 喪失時に silent degrade)。

### A-3. debug console.log 削除 (1 箇所)

- L3733 `console.log('✅ Recovery succeeded:...')` を削除 (console.warn/error の error-handling ログは対象外・残す)。

### A 受け入れ基準

- [ ] `grep -c "[^a-zA-Z.]escape(" src/page.js.html` → **0**
- [ ] page.js.html に native `confirm(`/`alert(` が残っていない (`grep -nE "[^.a-zA-Z](confirm|alert)\(" src/page.js.html` で確認、コメント内は除外可)
- [ ] `npm test` 全パス + `npm run lint:errors` 0 exit
- [ ] コミット 1 個: `fix(frontend): 非推奨 escape() を escapeHtml に置換 + native dialog fallback 除去 (vNNNN)`

---

## フェーズ B: バックエンド response builder 整合 (中・慎重に) 🟡 — 見積 2-3 時間

**目的**: API 境界の response が helpers.js builder (`{success, message, ...}` + error 時は `error` フィールド併設) と手書き `{success: false, error: ...}` で混在し、フロントが `response.error || response.message` の二重チェックを強いられている状態を段階的に解消する。

**進め方 (shape-preservation が絶対条件)**:

1. まず **UserApis.js** (使用 0・ファイル小) から。全手書き response を列挙し、各々を分類:
   - **(a) API 境界 response** で、builder の出力 shape が現行 shape の**上位集合** (フィールド追加のみ・削除/改名なし) になる → builder に置換
   - **(b) 内部サービス envelope** (呼び出し側が特定フィールドを直接読む) → **触らない**
2. 置換ごとに、その関数の呼び出し側 (フロントの runServer 受け側 + テスト) を grep し、読んでいるフィールドが置換後も存在することを確認する。
3. `npm test` が通ったら次のファイルへ: **ConfigService.js → SystemController.js** の順 (それぞれ別コミット)。ConfigService は `getUserConfig`/`saveUserConfig` の envelope が (b) に該当するものが多い — 無理に置換せず、(a) に該当するものだけでよい。**置換件数ゼロでも「調査の結果 (b) ばかりだった」なら、それを最終レポートに書いて正当に skip する**。
4. DataApis.js の `error` フィールド名不一致 (例: L154 `{success:false, error: ..., forms: []}`) は、builder の `createErrorResponse(message, data, extraFields)` が message と error を両方立てるので置換候補。同様に (a)/(b) 分類してから。

**任意 (時間があれば)**: `isValidFormUrl` (DataApis.js:2098) を validators.js へ移設し `/* global */` を更新。挙動は 1 文字も変えない。

### B 受け入れ基準

- [ ] 置換した各関数について「呼び出し側が読むフィールドが保持されている」grep 証跡を最終レポートに記載
- [ ] `npm test` 全パス (壊れたら shape 変更を疑い、その置換を revert)
- [ ] ファイルごとに 1 コミット (UserApis / ConfigService / SystemController)
- [ ] **A・B 完了後に `npm run deploy:all`** → 両テナント `systemDiagnosis` PASS + `npm run logs:cloud -- --tail` でエラーなし

---

## フェーズ C: scripts/lib/theme-core.js 抽出 (保守性・デプロイ不要) 🟡 — 見積 1 日

**目的**: theme-*.js 8 本の重複 (~485 行) を `scripts/lib/theme-core.js` に共通化。**全ツールの出力は 1 byte も変えない** (theme:perfect が子ツールの stdout を parse しているため)。

### C-0. 事前スナップショット (必須・最初に実行)

```bash
node scripts/theme-contrast.js --json > /tmp/before-contrast.json
node scripts/theme-matrix.js --uncovered > /tmp/before-uncovered.txt
node scripts/theme-verify.js > /tmp/before-verify.txt
node scripts/theme-perfect.js > /tmp/before-perfect.txt 2>&1
node scripts/theme-audit.js > /tmp/before-audit.txt 2>&1
```

### C-1. 色ユーティリティ (完全同一実装の統一)

`scripts/lib/theme-core.js` を新規作成 (CommonJS)。theme-contrast.js から以下を移植し、theme-contrast.js と theme-matrix.js の両方が require するように:

- `parseColor(value)` (contrast L28-53 / matrix L34-61 — matrix 側の null チェック入りを正とする)
- `blendOnBg(fg, bg)` (contrast L56-66 / matrix L63-73)
- `relativeLuminance({r,g,b})` (contrast L72-78 / matrix `relLum` L75-81 — 実装同一・名前だけ違う)
- `contrastRatio(fg, bg)` (contrast L80-86 / matrix `contrast` L83-93)

### C-2. CSS token ユーティリティ

- `extractBlockBody(css, startRegex)` (audit L35-47 / contrast L92-104 / matrix `extractBlock` L99-110 — 3 実装同一)
- `extractTokensFromBody(body)` (regex `/--([a-z0-9-]+)\s*:\s*([^;]+);/gi` — 3 ファイル同一)
- `resolveToken(tokens, name, depth = 0)` — **depth ガードは matrix 版 (depth=8) を正とする** (contrast 版はガード無し。ガード追加は無限再帰時のみ挙動が変わる = 現状の正常入力では出力不変)

### C-3. exempt 判定

- `isLineExempt(line)` (`/theme:exempt/i` — 9 箇所に散在)
- `findExemptBlockRanges(text)` (exempt-block-start/end パース — perfect L68-75 / matrix L259-270 / tokenize `findExcludedRanges` L83-102 / normalize L92-112 の 4 実装を統一)
- `isInExemptBlock(charIdx, ranges)`

### C-4. ファイル走査

- `listSrcFiles(srcDir)` (readdir + .html/.js filter — 4 パターン混在を統一)

### C-5. exempt hex リスト — **挙動が変わり得るので別コミット・別扱い**

theme-audit.js (16 hex 前後) と theme-matrix.js (24 hex 前後) のリストが異なる。統一すると**判定件数が変わる可能性がある** (= C-1〜C-4 の「出力不変」ゲートとは別物)。手順:

1. 2 リストを diff し、差分の各色がどちらのツールで何件マッチしているか実測
2. **和集合**に統一した場合の theme:matrix `--uncovered` 件数 / theme:audit 出力 / theme:perfect スコアの変化を記録
3. 変化がゼロなら統一してコミット。**変化があるならユーザーに差分を提示して判断を仰ぐ** (勝手に exempt を広げない — 検出漏れ側に倒れるため)

### C 受け入れ基準 (C-1〜C-4 の各ステップ後に毎回)

```bash
node scripts/theme-contrast.js --json | diff /tmp/before-contrast.json -   # 差分ゼロ
node scripts/theme-matrix.js --uncovered | diff /tmp/before-uncovered.txt -
node scripts/theme-verify.js | diff /tmp/before-verify.txt -
node scripts/theme-perfect.js 2>&1 | diff /tmp/before-perfect.txt -
node scripts/theme-audit.js 2>&1 | diff /tmp/before-audit.txt -
npm test
```

- [ ] 上記 diff が全て空 (C-5 のみ例外 — 手順内の判断フロー従う)
- [ ] `tests/lib.theme-core.test.cjs` を新設: parseColor (hex3/hex6/rgba/不正値) / blendOnBg (α=0/1) / contrastRatio (白黒=21) / resolveToken (循環参照で null) / findExemptBlockRanges の最低 10 ケース
- [ ] コミット分割: C-1+C-2 / C-3+C-4 / テスト / (C-5 は条件付き)
- [ ] デプロイ不要 (scripts のみ)。docs/THEME.md の「重複排除は未実施」の一文を「実施済み (lib/theme-core.js)」に更新

---

## フェーズ D: やらないことの再確認

- 性能改善 (実測 healthy・ERROR 0)。board data cache の size 計測などの観測強化も今回はしない (課題が観測されたら別途)
- login.js の fallback / ColumnMappingService thin wrapper / VALID_BOARD_MODES の fallback 定義 — すべて意図的設計
- フロントの polling・イベント管理・XSS 対策の追加変更 (監査で健全と確認済み)

---

## 実行順序と最終レポート

1. **A → B → C** (A・B は src 変更なのでまとめて 1 回の `deploy:all`、C はデプロイ不要)
2. 各フェーズ後: `npm test` + `npm run lint:errors`
3. 最終レポート必須項目: A の置換箇所と文字化け再現の有無 / B の (a)/(b) 分類表と置換件数 (ゼロ skip 含む) / C の削減行数 (before/after wc) と diff 全空の証跡 / C-5 の判断結果 / デプロイと両テナント systemDiagnosis 結果 / 見送った項目と理由

---

## 実行プロンプト (他モデル用・コピペ可)

```text
/Users/ryuma/Everyone-s-Answer-Board で作業してください。

docs/PLAN_app_optimization_2026-07.md を最初に全文読み、その指示書に従って
アプリ本体の品質最適化を実行してください。要点:

- フェーズ A (page.js.html の escape()→escapeHtml ×5 / native confirm・alert fallback 除去 /
  console.log 削除) → フェーズ B (response builder の shape 保存統一) → フェーズ C
  (scripts/lib/theme-core.js 抽出) の順。
- §0.5 ガードレール厳守。特に: 性能リファクタ禁止 (本番実測 healthy が根拠)、
  response の shape (フィールド名) 変更禁止、内部サービス envelope は builder 対象外、
  theme-perfect の spawnSync 境界維持、scripts は CommonJS、login.js の fallback は触らない。
- フェーズ B は各置換の前に呼び出し側 (フロント + テスト) を grep し、読まれている
  フィールドが保持されることを確認。該当ゼロなら正当に skip してレポートに書く。
- フェーズ C は着手前に /tmp へ全ツールの出力スナップショットを取り、各ステップ後に
  diff が空であることを確認 (指示書 C-0/C 受け入れ基準のコマンドをそのまま使う)。
  C-5 (exempt hex 統一) だけは出力が変わり得るので、変わる場合はユーザーに判断を仰ぐ。
- 行番号は 2026-07-05 時点。ズレていたらコード断片で再特定。
- 各フェーズの受け入れ基準チェックリストを全て満たしてから次へ。毎フェーズ
  npm test と npm run lint:errors を実行。
- コミットは指示書の分割単位で「type(scope): 日本語 (vNNNN)」(vNNNN = 直前+1)。
  A・B 完了後に git push と npm run deploy:all を実行し、両テナントの
  systemDiagnosis PASS と logs:cloud --tail のエラーなしを確認 (deploy:all が
  permission denied ならユーザーに実行を依頼して待つ)。C はデプロイ不要。
- 完了後の最終レポート: A の置換箇所と日本語文字化け再現の有無 / B の (a)/(b) 分類表と
  置換・skip 件数 / C の削減行数と diff 空の証跡 / C-5 の判断 / デプロイ検証結果 /
  見送った項目と理由。
```
