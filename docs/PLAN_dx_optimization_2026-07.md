# 指示書: 開発体験(DX)最適化リファクタ 2026-07

**目的**: 今後の開発をしやすくするため、(1) 毎セッション読み込まれる CLAUDE.md の肥大・重複を解消し、(2) Claude の永続メモリ/グローバル設定の鮮度を回復し、(3) テスト・CI のガードレール空白を埋める。
**対象外**: ランタイムコード (src/) のロジック変更は一切行わない。デプロイ (`npm run deploy:prod` / `deploy:open`) も不要。

この指示書は 2026-07-05 時点の実測監査 (3 並列エージェント + 手動裏取り) に基づく。行番号は当時のもの。実行時にズレていたらセクション見出しで再特定すること。

---

## 0. 検証済みの事実 (前提)

| 事実 | 根拠 (実測) |
| ---- | ---- |
| CLAUDE.md は 573 行。参照ファイル 25 個は全て実在、npm scripts 記載も全て実在 | `wc -l` / grep 照合済み |
| CLAUDE.md の Admin API 一覧 (40-100 行) は docs/DEVELOPMENT.md と約 95% 重複 | 両ファイル比較済み |
| CLAUDE.md のテーマ節は 222-400 行 (約 180 行)。SA pool 節は 182-221 行 + 546-564 行 | `grep -n "^#" CLAUDE.md` |
| バージョン記述は v2847 で停止、実体は v2866+ (約 19 commit 遅延) | git log 照合済み |
| `src/AccessControl.js` を参照するテストはゼロ (`grep -rl AccessControl tests/` → 0 件) | 実測 |
| CI (.github/workflows/ci.yml) は syntax check + `npm test` のみ。lint は pre-commit hook でしか走らない | ci.yml 確認済み |
| package.json に `postinstall` なし。hooks 有効化は手動 (`git config core.hooksPath .githooks`) | 確認済み |
| .gitignore / .clasp テンプレート分離 / .DS_Store 対策は **既に健全** — 触らない | 確認済み |
| npm scripts は 33 個だが namespace (theme:/logs:/deploy:) が整っており削減不要 | 確認済み |
| Claude メモリ: `~/.claude/projects/-Users-ryuma-Everyone-s-Answer-Board/memory/` に 27 ファイル 828 行。2026-05 末で更新停滞 | 実測 |
| グローバル `~/.claude/settings.local.json` は 2025-08-22 から未更新 (git/python の allow 3 件のみ) | 実測 |

## 0.5 ガードレール — 絶対にやらないこと

1. **page.js.html / AdminPanel.js.html の物理分割をしない** — CLAUDE.md に "Single-file JS (no physical splitting - intentional)" と明記済み。ユーザーのレビュー方針も「ファイルサイズ自体は減点しない」。
2. **DataApis.js 等バックエンド .js の分割をしない** — サイズでなく重複が問題になった時に別途判断。
3. **3 層キャッシュの統合 facade を作らない** — CLAUDE.md に「意図的分離」と明記済み。
4. **.clasp*.json 群の再配置・リネームをしない** — scripts/env-switch.js が依存。
5. **CLAUDE.md から「ルール」を削らない** — 移譲するのは「詳細・経緯・網羅表」。Must-Follow Rules / 禁止パターン / Do-Don't / 公開ライフサイクル 4 関数の不変条件は CLAUDE.md に残す。
6. **src/ のロジック・HTML テンプレート include 順を変更しない**。
7. コミットは論理単位ごとに分け、メッセージは既存規約 `type(scope): 日本語 (vNNNN)` に従う。バージョン番号は直前コミットの vNNNN+1。

---

## フェーズ1: CLAUDE.md ダイエット (573 → 約 280 行) 🔴 最優先

### 1-1. テーマ/カラー節を docs/THEME.md へ移譲

- CLAUDE.md 222-400 行「Theme / Color アーキテクチャ」全体を新規 `docs/THEME.md` に移動 (見出し階層を `#`/`##` に昇格)。
- CLAUDE.md 側には **15 行以内**の要約を残す:
  - 全 CSS は `--theme-*` semantic token 経由。ハードコード hex/rgba 禁止
  - Tailwind `dark:` prefix 使用禁止 (theme:perfect axis 27b で機械検証)
  - 新規 UI は semantic primitives (`.card` / `.modal-card` / `.box`) を使う
  - brand/status 色は theme 非依存 (`--brand-*` / `--status-*`)
  - 検証: `npm run theme:contrast` / `theme:matrix` / `theme:perfect`
  - 詳細: [docs/THEME.md](THEME.md) へのリンク

### 1-2. Admin API / Cloud Logging 節の縮小

- CLAUDE.md 40-100 行の Admin API コマンド全一覧を削除し、以下だけ残す (約 20 行):
  - 「いつ使うか」4 項目 (既存のまま)
  - 代表 5 コマンド: `systemDiagnosis` / `getLogs` / `getUsers` / `setUserConfig` / `previewBoard`
  - 仕組み 1 行 (`dispatchAdminOperation`) + 全コマンドは docs/DEVELOPMENT.md 参照
- **移譲前に必ず diff 照合**: CLAUDE.md 側にだけ存在するコマンド (SA pool 系 v2782+ 等) があれば docs/DEVELOPMENT.md に先に追記してから削る。情報を消失させない。
- Cloud Logging 節 (101-137 行) も同様: 代表 3 コマンド (`--severity ERROR` / `--tail` / `--user`) + deployment フィルタの説明 1 段落 + DEVELOPMENT.md リンクに縮小。全オプション表は DEVELOPMENT.md へ。

### 1-3. SA Pool アーキテクチャを docs/ARCHITECTURE.md へ移譲

- 新規 `docs/ARCHITECTURE.md` を作成し、CLAUDE.md 182-221 行 (SA Pool アーキテクチャ) + 546-564 行 (SA pool セキュリティモデル) を**統合して 1 節に** (現在 2 箇所に分散し内容が一部重複)。
- CLAUDE.md 側に残すもの (約 15 行): アクセスモード自動振り分け表 (owner/admin/viewer × DB/own/他人公開/他人非公開) 1 個 + 「viewer は SS 直接権限なし (通常 Google Form 同等)」の 1 行 + ARCHITECTURE.md リンク。
- per-row CAS lock / board data cache / sa_validation invalidate の詳細段落も ARCHITECTURE.md へ。

### 1-4. バージョン履歴記述の現行仕様化

- CLAUDE.md 全体から `(v2782+)` `(v2800+, v2816 hybrid palette)` 等の**見出し・本文中のバージョン注記を削除**し、現行仕様として書き直す。
  - 例: 「board data cache (v2783, DataApis.js)」→「board data cache (DataApis.js)」
  - 経緯が重要な箇所 (旧方式が廃止された理由等) は docs/ARCHITECTURE.md / THEME.md 側に「経緯」として残してよい。
- 理由: バージョン番号を CLAUDE.md に書き続ける限り恒常的に陳腐化する。経緯は git log と Claude メモリが担う。

### 1-5. README.md 更新

- 2 テナント構成 (那覇市 = default / 沖縄県 = open) と `npm run deploy:open` / `push:open` / `api:open` / `logs:open` の存在を README のデプロイ節に追記。
- README 記載のセットアップ手順・npm scripts が現 package.json と一致するか照合し、乖離のみ修正。

### フェーズ1 受け入れ基準

- [ ] `wc -l CLAUDE.md` が 320 行以下
- [ ] docs/THEME.md / docs/ARCHITECTURE.md が存在し、削った情報が全て収容されている (grep で spot check: `pickServiceAccount_`, `bumpBoardDataVersion_`, `--theme-bg-elevated-hover`, `theme:uncovered` 等が docs 側にヒットする)
- [ ] CLAUDE.md 内に `v2[0-9]{3}` にマッチする文字列が残っていない (`grep -cE "v2[0-9]{3}" CLAUDE.md` → 0)
- [ ] Must-Follow Rules / Board Publish Lifecycle / Canonical Frontend APIs / Cache 3 層の判断フロー / 禁止パターンは CLAUDE.md に残存
- [ ] `npm test` パス (ドキュメントのみの変更なので当然通るはずだが確認)
- [ ] コミット分割: (a) docs 新設+CLAUDE.md 縮小、(b) README 更新

---

## フェーズ2: Claude メモリ / グローバル設定の鮮度回復 🟡

対象はリポジトリ外: `~/.claude/` 配下。

### 2-1. メモリ統合 (consolidate)

`~/.claude/projects/-Users-ryuma-Everyone-s-Answer-Board/memory/` (27 ファイル) について:

- **陳腐化照合**: 以下を git log / 現コードと照合し、解決済みなら本文を「解決済み (vNNNN)」に更新するか削除:
  - `project_viz_toolbar_pending_deploy.md` — 既に deploy 済みと本文にあるがスラッグが "pending"。リネームか統合
  - `project_option_b_remaining_edge_cases.md` — L2/L3/L4 edge case がその後修正されたか照合
  - `project_v2690_admin_autosave.md` / `project_option_b_past_phase_navigation.md` — 「deploy 待ち」記述が残っていれば更新
- **統合**: viz 系 7 ファイルのうち、完了済みの実装記録は 1 つの `project_viz_design_decisions.md` に統合し「なぜその設計か (多数派バイアス回避・CVD 対応等)」の判断だけ残す。実装詳細は git log にある。
- MEMORY.md インデックスを再生成 (1 行 1 メモリ、リンク切れゼロ)。
- **フェーズ1 との連動**: CLAUDE.md から docs へ移譲した内容と重複するメモリがあれば削除 (メモリは「repo に書いていないこと」だけを持つ原則)。

### 2-2. 未記録の設計判断の追記

- v2848〜v2866 の git log を確認し、既存メモリでカバーされていない設計判断があれば 1-2 ファイル追加 (候補: `__applyPublishStateChange` への公開不変条件集約、/simplify 品質リファクタの方針)。
- 判断基準: 「コードや git log から読み取れない why」だけをメモリにする。

### 2-3. グローバル settings.local.json の統合

- `~/.claude/settings.local.json` (2025-08 から未更新、allow: `git --version` / `python3:*` / `git clone:*`) の 3 エントリを `~/.claude/settings.json` の permissions.allow に統合し、settings.local.json は空の permissions にするか削除。
- **注意**: プロジェクト側 `.claude/settings.json` の自己改変は classifier deny の対象 (過去の経緯)。グローバル側の編集で拒否された場合は無理にリトライせず、ユーザーに手動編集用の diff を提示して終了すること。

### フェーズ2 受け入れ基準

- [ ] MEMORY.md の行数がメモリファイル数と一致し、リンク切れゼロ
- [ ] "pending" / "deploy 待ち" を含む陳腐化メモリがゼロ
- [ ] メモリ総ファイル数が 27 → 20 前後に統合されている

---

## フェーズ3: テスト・CI ガードレールの空白埋め 🟡

### 3-1. AccessControl.js のユニットテスト新規作成

- 新規 `tests/access.control.test.cjs` (既存命名規則 `{module}.{scope}.test.cjs` に従う)。
- パターンは `tests/main.doPost.test.cjs` と同じ: `vm.createContext()` に GAS globals (Session / DriveApp / SpreadsheetApp / PropertiesService / CacheService) をスタブして src/AccessControl.js をロード。
- カバーすべき分岐 (src/AccessControl.js を読んで実際の関数名に合わせる):
  - owner 本人 → 許可 / viewer が公開ボード → 許可 / viewer が非公開ボード → 拒否
  - admin が他人の非公開ボード → 許可
  - Drive API エラー時に fail-open していないこと (fail-closed の検証)
  - collaborator (共同編集者) 判定の cache が効くこと (可能なら)
- 最低 8 ケース。src/ のコードは一切変更しない (テストが通らない実挙動を見つけたら**修正せず報告**)。

### 3-2. CI に lint ゲート追加

- `.github/workflows/ci.yml` の "Check JavaScript syntax" と "Run unit tests" の間に 1 ステップ追加:

```yaml
      - name: Lint (CLAUDE.md rules, error severity)
        run: npm run lint:errors
```

- 目的: `--no-verify` コミットや hooks 未設定環境からの push を捕捉。

### 3-3. hooks 自動有効化

- package.json の scripts に追加: `"postinstall": "git config core.hooksPath .githooks || true"`
- `|| true` 必須 (CI の `npm ci` は git repo 外や shallow でも失敗させない)。
- README のセットアップ手順から「`git config core.hooksPath .githooks` を手動実行」の記述を「npm install で自動設定される」に更新。

### 3-4. デプロイ時 git tag (任意・時間があれば)

- scripts/deploy-prod.js はデプロイ時にバージョン番号を扱っている。デプロイ成功後に `git tag v<N>` を打つ処理を追加 (既存 tag と衝突したら skip して警告のみ)。失敗してもデプロイ自体は成功扱いにする (tag は best-effort)。

### フェーズ3 受け入れ基準

- [ ] `node --test tests/access.control.test.cjs` 単体パス + `npm test` 全体パス
- [ ] `npm run lint:errors` がローカルで 0 exit (CI に入れる前に確認)
- [ ] `npm ci` がクリーン環境で成功する (postinstall が壊していない)
- [ ] コミット分割: (a) テスト追加、(b) CI + postinstall、(c) tag 機能 (任意)

---

## フェーズ4: テーマ CLI のスリム化 🟢 低優先 (時間が余れば)

- scripts/theme-*.js 8 本 (計約 2,300 行) のうち、色パース / CSS token 抽出 / contrast 計算の重複実装を `scripts/lib/theme-core.js` に抽出して各ツールから require。**出力フォーマットと npm scripts のインターフェースは変えない**。
- `theme-tokenize.js` / `theme-normalize-typography.js` / `theme-pair-tailwind.js` は移行専用ツール。リポジトリ内に残作業があるか確認 (`npm run theme:uncovered` 等) し、ゼロなら「移行完了につき削除候補」と docs/DEVELOPMENT.md に注記 (削除自体はユーザー判断待ち)。
- 受け入れ基準: 全 `npm run theme:*` が変更前後で同一出力 (代表 3 コマンドで diff 確認)。

---

## 実行順序と検証

1. フェーズ1 → 2 → 3 → (4)。1 と 2 は独立なので順不同可。
2. 各フェーズ完了ごとに `npm test` + `npm run lint:errors`。
3. **デプロイは不要** (ランタイムコード非変更)。フェーズ3-4 の deploy-prod.js 変更のみ、次回ユーザーがデプロイする際に動作確認される旨をコミットメッセージに明記。
4. 全完了後、最終レポート: 削減行数 (CLAUDE.md before/after)、新設ファイル、メモリ統合数、追加テストケース数、未完了項目とその理由。

---

## 実行プロンプト (他モデル用・コピペ可)

```text
/Users/ryuma/Everyone-s-Answer-Board で作業してください。

docs/PLAN_dx_optimization_2026-07.md を最初に全文読み、その指示書に従って
DX 最適化リファクタを実行してください。要点:

- フェーズ1 (CLAUDE.md ダイエット) → フェーズ2 (~/.claude のメモリ/設定整理)
  → フェーズ3 (AccessControl テスト + CI lint + postinstall) の順。
  フェーズ4 は時間が余った場合のみ。
- 指示書の「ガードレール — 絶対にやらないこと」を厳守。特に:
  page.js.html / AdminPanel.js.html の分割禁止、src/ ロジック変更禁止、
  キャッシュ 3 層の統合禁止、CLAUDE.md からルール(Must-Follow/禁止パターン)を
  削ることの禁止。移譲するのは「詳細・網羅表・経緯」のみ。
- 情報は移動であって削除ではない。CLAUDE.md から削る前に、docs 側に
  同内容が存在することを grep で確認してから削ること。
- 行番号は 2026-07-05 時点。ズレていたらセクション見出しで再特定。
- 各フェーズの「受け入れ基準」チェックリストを全て満たしてから次へ。
  各フェーズ完了ごとに npm test と npm run lint:errors を実行。
- コミットは指示書記載の分割単位で、既存規約
  「type(scope): 日本語メッセージ (vNNNN)」に従う (vNNNN は直前コミット+1)。
  git push とデプロイ (deploy:prod / deploy:open) は実行しない。
- テスト作成中に src/ の実挙動バグを見つけた場合は修正せず、最終レポートで報告。
- 完了後の最終レポート: CLAUDE.md の before/after 行数、新設・変更ファイル一覧、
  メモリ統合結果 (27→N)、追加テストケース数、スキップした項目とその理由。
```
