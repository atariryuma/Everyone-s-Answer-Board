# アーキテクチャ詳細: SA Pool / キャッシュ / 負荷対策

CLAUDE.md にはアクセスモード振り分け表と要点のみを置き、実装詳細・設計経緯はこのドキュメントに集約している。700 人スケールの同時アクセスを service account (SA) pool で捌く設計。

---

## セキュリティモデル: 「通常 Google Form 同等」

ボード SS (スプレッドシート) は通常の Google Form と同等のセキュリティを持つ:

- **viewer (生徒)** は SS への直接権限を持たず、Web App 経由のみアクセス可能 (SA pool 経由)。viewer の Drive にボード SS は表示されない、直接編集も不可
- **owner (教師)** は自分の SS のオーナー権限を持つので `openById` 直接 (SA quota 節約)
- 旧 `DOMAIN_WITH_LINK + EDIT` (domain-wide) 共有は廃止 — `migrateBoardSharing` admin API で既存ボードを cleanup 済

ボード SS への cross-user アクセス (生徒の閲覧 / 管理者の管理) を **SA pool** が担う。

---

## アクセスモード自動振り分け

`validateServiceAccountUsage` → `accessMode`、`openSpreadsheet` が経路を自動判定する:

| Caller | DB | own board | 他人公開 | 他人非公開 |
| ------ | -- | --------- | -------- | ---------- |
| **owner (editor)** | sa | **own (openById)** | sa | denied |
| **admin** | sa | **own** | sa | sa |
| **viewer (生徒)** | sa | — | sa | denied |

owner は own OAuth で SA quota 節約。viewer / admin の cross-user のみ SA pool 経由。DB sheet は常に SA pool。

---

## SA pool 設計 ([DatabaseCore.js](../src/DatabaseCore.js))

- **round-robin** (`pickServiceAccount_`): ScriptCache 共有 counter `sa_rr_counter` で完全均等分散
- **30s cooldown**: 429 を喰った SA は除外 → auto failover
- **2 段 token cache** (`getServiceAccountAccessToken_`): in-memory + ScriptCache 50min
- **per-SA per-sheet verify cache** (`verifyServiceAccountAccess_`): 10min ok / 2min no (transient は焼かない)
- **authResolver closure** (`makeProxyAuthResolver_`): proxy hot path で SA 動的切替

SA pool の shared 設定 (SA pool 全員を editor 追加) は [SharingHelper.js](../src/SharingHelper.js) が担う。

---

## per-row CAS lock ([ReactionService.js](../src/ReactionService.js))

- **旧**: `LockService.getScriptLock()` を process() 全体で保持 → 全 board が serialize する。700 人 burst の bottleneck だった
- **新**: ScriptLock は ~5-10ms の critical section (cache check+put) のみ。異 row 同士は完全並列
- 効果: throughput ~5 → ~200 req/sec、40x 向上

---

## board data cache ([DataApis.js](../src/DataApis.js))

- viewer の `getPublishedSheetData` 結果を 12s ScriptCache
- reaction/highlight write 時に `bumpBoardDataVersion_` で即時 stale
- 効果: 700 viewer × 8s polling = 5,250 req/min → ~700 req/min (~8x 削減)
- **cache に書く形は圧縮表現** (`packBoardDataForCache_` / `unpackBoardDataFromCache_`): 行から
  派生 3 項目 (`opinion` = answer / `isEmpty` = false / `formattedTimestamp` = timestamp から復元) を
  落とし、`Utilities.gzip` + base64 で書く。ScriptCache の 1 value 100KB 制限に 150〜250 行で当たり
  「cache に入らない → 毎 poll が Sheets を読む」になっていたため (2026-10)。読むときに元の形へ戻す
  ので wire (受け手) は変わらない。`withStampedeGuard_` の `pack` / `unpack` が差し込み口

---

## circuit breaker ([DatabaseCore.js](../src/DatabaseCore.js) `fetchSheetsAPIWithRetry`)

- 429 ごとに共有 counter (`circuit_breaker_state`) を +1、成功で -1 (net pressure)
- **開くのは pool 全体が焼けたときだけ**: 429 を踏んだ SA を cooldown に入れたあと、cooling でない
  SA が他に残っていれば failover に任せて開かない。残っていなければ、積んだ圧力に応じて
  30s / 60s / 120s 止める。1 つの SA の quota 切れで他の 4 本まで止めていた (2026-10)

---

## users シートの書き込み省略 ([DatabaseCore.js](../src/DatabaseCore.js) `updateUser`)

- 書き込みのたびに `USER_CACHE_VERSION` が上がり、全 viewer の user cache が一斉に落ちる
  (stampede guard で 1 read には抑えてある)
- いま読んだ行と内容が同じ (configJson は `lastAccessedAt` / `etag` を除いて比較) なら書かず、
  `{ success, unchanged: true, currentEtag }` を返す。`saveUserConfig` は永続済の etag を client に返す
- per-user の cache version は見送り: `all_users_v<ver>` 一覧 (findUserByEmail / findUserBySpreadsheetId
  が configJson を読む) の鮮度が崩れ、profile 切替直後の児童が古い SS を見る

---

## sa_validation cache 即時 invalidate

- `__applyPublishStateChange` で `invalidateSaValidationCache_` を呼び、該当 SS の cache version を bump
- unpublish 直後の 60 秒 access leak を解消

---

## データ保存則 (v2931)

共有 DB (テナントごとに別スプレッドシート) は 1 つのルールで統一する:

> **有界で書き換わる状態 = JSON セル / 無界で追記一回きりのアーカイブ = 1 件 1 行**

| シート | 種別 | 中身 |
| ------ | ---- | ---- |
| `users` | 状態 | `configJson` (1 ユーザー数 KB、有界) |
| `lessons` | 状態 | `lessonJson` = 授業の定義 + 遷移履歴 + **範囲ポインタ** (~4KB) |
| `lesson_responses` | アーカイブ | 1 回答 1 行 × 9 列、追記のみ |

- snapshot は `{sheet, startRow, rowCount}` のポインタを持ち、読み出しは範囲読み
  (規模に依らず 1 フェーズ分のセルのみ)。行は `lessonId + phaseIndex` を照合し、
  ポインタずれで他授業の回答が混入しない。
- 書込は SA proxy の `appendRows` (values:append)。複数行が連続範囲で原子的に入り、
  応答の updatedRange から開始行を得る (lock 不要)。
- 授業削除は `lessons` 行のみ (アーカイブ行は孤児として残る = 照合ガードで無害)。
- **アーカイブ行を lessonJson に戻さない**こと。Sheets の 1 セル 50,000 字上限に対する
  本文切り詰め (shrink) サブシステムが復活する。v2931 で 1 授業 44,698 字 → 4,198 字。
- ポインタは sheet 名を持つので、将来 `lesson_responses_2027` のような年次分割へ
  移行しても過去ポインタはそのまま読める。
- 保守: `lesson.recaptureArchive`
  (元 SS が読める phase を全文で焼き直す)。

## 負荷検証 (CI 対象外、手動)

```bash
node scripts/load-test-concurrent.js --n 30 --op previewBoard  # N 並列 API call で SA pool / 429 を検証
```

SA pool の追加・管理は Admin API 経由 (詳細は [DEVELOPMENT.md](DEVELOPMENT.md) の SA pool 管理コマンド参照)。運用上は 700 人スケールで secondary SA を 3-5 個推奨 (`SERVICE_ACCOUNT_CREDS_2` 〜 `_10`)。
