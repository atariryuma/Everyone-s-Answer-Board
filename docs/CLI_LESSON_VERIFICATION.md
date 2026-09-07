# Lesson Workspace — CLI 検証手順

授業ワークスペース（lesson の作成・進行・アーカイブ）を、ブラウザ操作せずに `npm run api`
経由で end-to-end に検証する手順。

## 前提

- `scripts/config.json` に `apiKey` と `prodUrl` が設定済み
- 本番デプロイ済み (`npm run deploy:prod`)
- 検証する `--userId` を `npm run api -- getUsers` で控える

## 全 lesson 操作一覧

唯一の定義は `src/AdminApis.js` の `dispatchAdminOperation()` (`case 'lesson.*'`)。
`scripts/admin-api.js` の `OPERATIONS` は `--help` 表示用のミラー。

```bash
# 作成・編集 (draft のみ)
npm run api -- lesson.templates                            # 利用可能テンプレ一覧 (userId 不要)
npm run api -- lesson.create        --userId <uuid> --name '<名前>' --template <templateKey>
npm run api -- lesson.updateDraft   --userId <uuid> --lessonId <id> --fieldPath <path> --value <json|string>
npm run api -- lesson.duplicate     --userId <uuid> --lessonId <id> [--options '{"copyClasses":true}']
npm run api -- lesson.knownClasses  --userId <uuid>        # 過去授業から使用クラス候補
npm run api -- lesson.reorderPhases --userId <uuid> --lessonId <id> --order '[0,2,1]'

# 進行
npm run api -- lesson.start         --userId <uuid> --lessonId <id>
npm run api -- lesson.advance       --userId <uuid> --lessonId <id> --direction next|previous [--targetIndex <n>]
npm run api -- lesson.end           --userId <uuid> --lessonId <id>
npm run api -- lesson.reopen        --userId <uuid> --lessonId <id>   # completed → active の逆遷移

# 参照・後処理
npm run api -- lesson.list          --userId <uuid>        # snapshots は除外された軽量一覧
npm run api -- lesson.review        --userId <uuid> --lessonId <id>
npm run api -- lesson.reviewGrid    --userId <uuid> --lessonId <id>   # 児童ごとの移動を授業後に読む
npm run api -- lesson.closeForms    --userId <uuid> --lessonId <id>   # 状態を問わず全 Form の受付を締切
npm run api -- lesson.recaptureArchive --userId <uuid> --lessonId <id> --phaseIndex <n>
npm run api -- lesson.delete        --userId <uuid> --lessonId <id>
```

`lesson.closeForms` は **教師本人のブラウザ経由でしか成功しない**（FormApp は Form 所有者の
権限で動くため、API キー = 管理者経路では他人の Form を開けない）。CLI からは
「権限で落ちること」の確認にしかならない。

## テンプレート (`LESSON_TEMPLATES`, src/LessonService.js)

| templateKey | ラベル | フェーズ | 入力経路 |
| ----------- | ------ | -------- | -------- |
| `dialogue-reconsider-5phase` | 考え、議論する道徳（5段階） | 考える → 出会う → 議論する → もう一度考える → ふりかえる | **native** |
| `dialogue-3phase` | 考え、議論する道徳（3段階・短時間） | 考える → 出会う → もう一度考える | **native** |
| `dialogue-2phase` | 考えの変化（2段階） | 考える → もう一度考える | **native** |
| `before-after-2phase` | 立場の変化（数直線・2段階） | 議論のまえ → 議論のあと | Google Form |
| `doutoku-3phase` | 立場の変化（数直線・3段階） | はじめの考え → 話し合いのあと → ふりかえり | Google Form |
| `survey-pie` | アンケート（円グラフ） | アンケート | Google Form |
| `survey-board` | 意見を集める（掲示板） | 意見 | Google Form |

`doutoku-3phase` は旧名の残置 (tests と既存 lessonJson の `template` 参照との互換のため key は変えていない)。
実体は 2026-09-07 に「数直線で 3 回」に揃えた (旧: 数直線 → 4 象限 → 数直線 で、フェーズ間の比較が
成立しなかった)。`kid-3phase` / `inquiry-3phase` は同日に廃止 (可視化がフェーズごとに変わり、変化を
追えない)。native の既定の縦軸は「迷いあり ↔ 迷いなし」(横軸は教材依存なので空)。

**native テンプレ (`inputMode: 'native'`) は経路が違う**: Google Form を作らず、児童が
アプリ内で直接入力する（`submitLessonAnswer`）。したがって `lesson.start` しても Form は
生成されず、以下の smoke の「Form が 3 つできる」前提は当てはまらない。CLI で Form 生成まで
含めて検証したいときは Form 系テンプレ（例 `doutoku-3phase`）を使う。

## end-to-end smoke (約 5 分)

### 1. 一覧 (初期状態確認)

```bash
USER_ID="<your-uuid>"
node scripts/admin-api.js lesson.list --userId $USER_ID
# → {"success":true,"data":{"lessons":[...]}}
```

### 2. draft 作成

```bash
node scripts/admin-api.js lesson.create \
  --userId $USER_ID \
  --name "CLI smoke $(date +%Y-%m-%d)" \
  --template doutoku-3phase   # Form 経路。native 検証なら dialogue-reconsider-5phase
# → lessonId を控える
LESSON_ID="lesson_xxx"
```

### 3. 対象クラスを設定 (lesson.start の前提条件)

```bash
node scripts/admin-api.js lesson.updateDraft \
  --userId $USER_ID --lessonId $LESSON_ID \
  --fieldPath classes --value '["5-1","5-2"]'
```

`--value` は JSON parse 試行 → 失敗時は raw string fallback。
- 配列/オブジェクト: `--value '["a","b"]'` `--value '{"x":1}'`
- 文字列: `--value 単純な文字列` または `--value '"明示的な文字列"'`
- 数値: `--value 42`

### 4. フェーズの質問を変更 (任意)

```bash
node scripts/admin-api.js lesson.updateDraft \
  --userId $USER_ID --lessonId $LESSON_ID \
  --fieldPath 'phases[0].question' \
  --value '「自分の立場は？」(CLI から書き換え)'
```

### 5. lesson 開始 (Form 自動生成 + state=active)

```bash
node scripts/admin-api.js lesson.start --userId $USER_ID --lessonId $LESSON_ID
# 注意: 実際の Google Form が 3 つ作成され、教師の Drive に残る
```

### 6. フェーズ進行

```bash
node scripts/admin-api.js lesson.advance --userId $USER_ID --lessonId $LESSON_ID --direction next
node scripts/admin-api.js lesson.advance --userId $USER_ID --lessonId $LESSON_ID --direction previous
```

### 7. 授業終了 (snapshot freeze + state=completed)

```bash
node scripts/admin-api.js lesson.end --userId $USER_ID --lessonId $LESSON_ID
# レスポンスに reviewUrl が返る
```

### 8. 振り返り取得 (snapshot 詳細)

```bash
node scripts/admin-api.js lesson.review --userId $USER_ID --lessonId $LESSON_ID
# snapshots[] には各 phase の {sheet, startRow, rowCount} ポインタが入り、
# 回答本文は DB の lesson_responses シートに 1 回答 = 1 行で積まれている。
# lesson.review はそのポインタから rows を読み戻して返す (hydrate)。
# → lessons シートの lessonJson に rows 本体を書き戻してはいけない (セル上限に当たる)。
```

### 9. 削除

```bash
node scripts/admin-api.js lesson.delete --userId $USER_ID --lessonId $LESSON_ID
# 注意: 生成済みの Google Form / Spreadsheet は Drive に残る (意図的。原本データを消さない)
```

## トラブルシュート

### `"Cannot read properties of undefined (reading 'forEach')"`

→ lessons シートが DB に未作成。`lesson.create` を最初に 1 回叩くと lazy bootstrap される
(`lesson_responses` シートも初回アーカイブ時に同様に lazy 作成される)。

### `"LESSON_BUSY"`

→ 別の `lesson.start` が並行実行中 (LockService が tryLock(5000ms) で待機)。少し待って再実行。

### `"FORBIDDEN_STATE: draft 状態でのみ編集できます"`

→ active や completed の lesson に `lesson.updateDraft` を呼んだ。draft の lesson にのみ編集可能。

### Auto-archive の検証

`unpublishBoard` 時に lesson が auto-archive されるかを試したい場合:

```bash
# 1. lesson.start でアクティブにする
node scripts/admin-api.js lesson.start --userId $USER_ID --lessonId $LESSON_ID

# 2. unpublish して archive が走るか確認
node scripts/admin-api.js unpublishBoard --userId $USER_ID

# 3. lesson.list で state が completed になっているか
node scripts/admin-api.js lesson.list --userId $USER_ID
```

5 分未満 / 回答 0 件 のときは skip される (auto-archive 条件)。

## Cloud Logging で実行確認

```bash
npm run logs:cloud -- --severity ERROR --hours 1
npm run logs:cloud -- --hours 1 --limit 30  # WARN 以上 (デフォルト)
```

## 関連ファイル

- `src/LessonService.js` — backend implementation (テンプレ定義 `LESSON_TEMPLATES` もここ)
- `src/LessonWorkspace.html` — 教師の授業ワークスペース UI
- `src/AdminApis.js` — dispatchAdminOperation cases (`lesson.*`)
- `scripts/admin-api.js` — CLI wrapper
- `tests/lesson.service.test.cjs` — lesson サービス本体のユニットテスト
- `tests/lesson.nativeMode.test.cjs` — native 入力経路 (フェーズ権能・匿名性)
- `tests/data.apis.lessonMask.test.cjs` — 「考える」フェーズで他者を返さない mask
- CLAUDE.md 「授業モード (native 入力)」 — 設計上の不変条件
