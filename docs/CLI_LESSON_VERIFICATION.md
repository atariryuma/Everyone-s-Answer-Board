# Lesson Workspace — CLI 検証手順

Phase 1+2 で実装した lesson archive 機能を、ブラウザ操作せずに `npm run api` 経由で
end-to-end に検証する手順。

## 前提

- `scripts/config.json` に `apiKey` と `prodUrl` が設定済み
- 本番デプロイ済み (`npm run deploy:prod`)
- 検証する `--userId` を `npm run api -- getUsers` で控える

## 全 lesson 操作一覧

```bash
npm run api -- lesson.list          --userId <uuid>
npm run api -- lesson.create        --userId <uuid> --name '<名前>' --template doutoku-3phase
npm run api -- lesson.updateDraft   --userId <uuid> --lessonId <id> --fieldPath <path> --value <json|string>
npm run api -- lesson.start         --userId <uuid> --lessonId <id>
npm run api -- lesson.advance       --userId <uuid> --lessonId <id> --direction next|previous
npm run api -- lesson.end           --userId <uuid> --lessonId <id>
npm run api -- lesson.review        --userId <uuid> --lessonId <id>
npm run api -- lesson.delete        --userId <uuid> --lessonId <id>
```

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
  --template doutoku-3phase
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
# lessonJson.snapshots[] に各 phase の rows が freeze されているはず
```

### 9. 削除

```bash
node scripts/admin-api.js lesson.delete --userId $USER_ID --lessonId $LESSON_ID
# 注意: 生成済みの Google Form / Spreadsheet は Drive に残る (Phase 1 仕様)
```

## トラブルシュート

### `"Cannot read properties of undefined (reading 'forEach')"`

→ lessons シートが DB に未作成。`lesson.create` を最初に 1 回叩くと lazy bootstrap される。

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

- `src/LessonService.js` — backend implementation
- `src/AdminApis.js` — dispatchAdminOperation cases (`lesson.*`)
- `scripts/admin-api.js` — CLI wrapper
- `tests/lesson.service.test.cjs` — 24 unit tests
