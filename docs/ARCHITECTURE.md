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

- viewer の `getPublishedSheetData` 結果を 10s ScriptCache
- reaction/highlight write 時に `bumpBoardDataVersion_` で即時 stale
- 効果: 700 viewer × 8s polling = 5,250 req/min → ~700 req/min (~8x 削減)

---

## sa_validation cache 即時 invalidate

- `__applyPublishStateChange` で `invalidateSaValidationCache_` を呼び、該当 SS の cache version を bump
- unpublish 直後の 60 秒 access leak を解消

---

## 負荷検証 (CI 対象外、手動)

```bash
node scripts/load-test-concurrent.js --n 30 --op previewBoard  # N 並列 API call で SA pool / 429 を検証
```

SA pool の追加・管理は Admin API 経由 (詳細は [DEVELOPMENT.md](DEVELOPMENT.md) の SA pool 管理コマンド参照)。運用上は 700 人スケールで secondary SA を 3-5 個推奨 (`SERVICE_ACCOUNT_CREDS_2` 〜 `_10`)。
