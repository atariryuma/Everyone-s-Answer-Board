# アーキテクチャ詳細: 配置 / アクセス / キャッシュ

CLAUDE.md にはアクセスの表と要点のみを置き、実装詳細・設計経緯はこのドキュメントに集約している。

---

## 配置の形: 先生ごとに 1 つ、先生として実行

- 先生ごとに、先生の Drive のスプレッドシートに紐づいたアプリを配置する (`scripts/deploy-single.js`)。
  `src/appsscript.json` の `webapp` は `executeAs: USER_DEPLOYING` (先生として実行) / `access: DOMAIN`
- データはそのスプレッドシート 1 つに入る: `users` (先生 1 人分) / `lessons` / `lesson_responses` /
  授業モードの回答シート (「授業名｜番号 フェーズ名」)。フォームの回答先は、フォームが掴んでいる別の SS
- 配置ごとに違うのは `src/Deployment.js` だけ (先生のメール + DB の SS ID を焼き込む)。初回の doGet で
  `ensureDeploymentBootstrap_` がプロパティとシートをそろえる。管理者を「最初に開いた人」で決めないため
- 先生がトップを開くと `processLoginAction` で登録して管理パネルを返す (ログイン・アプリ設定の画面は無い)

**Why**: 2026-10-29 10:20 ごろから、児童アカウントは「開いた人として実行」の Apps Script を中身に関係なく
開けなくなった (管理側の設定)。先生として実行すれば児童は自分の権限を何も使わない。それまでの
「みんなで 1 つのアプリ + 共有 DB + サービスアカウント (SA) pool」は廃止した (同日、ユーザー判断)。

---

## アクセス判定 ([DatabaseCore.js](../src/DatabaseCore.js) `checkSpreadsheetAccess_`)

実行者 = 先生 = すべてのシートの持ち主なので、開くのは常に `SpreadsheetApp.openById`。判定するのは
「いまの呼び出し元に見せてよいか」だけ:

| 呼び出し元 | 公開中のボードが指すシート | 非公開のボードのシート | どのボードにも無いシート |
| ---------- | -------------------------- | ---------------------- | ------------------------ |
| **先生 (管理者)** | 可 | 可 | 可 |
| **児童・ほかの先生** | 可 | 不可 | 不可 |

- 児童に開かせてよいのは「公開中のボードの設定 (いまのボード + 保存したプロファイル) に入っているシート」
  だけ (`findUserBySpreadsheetId_` → `isUserBoardPublished`)
- DB も同じ扱い (授業モードの回答シートが全部入るので、授業がボードにつながって公開されている間だけ)。
  users / lessons を読むサーバの処理は `openDatabase` で直接開く (判定を通らない)

### 画面から呼べる関数

- Apps Script は名前が `_` で終わらない関数をすべて `google.script.run` に出す。先生として実行するので、
  そのどれもが先生の権限で動く。画面から呼ばせない処理 (プロパティの読み書き、users の書き込み、
  引っ越し、リアクションの直書き、フォームの作り替え) は名前を `_` で終える
- 画面から呼ぶ管理パネルの処理 (フォームの作成・接続・変更、列の推定、画像、Drive のフォーム一覧) は
  入口で `teacherOnlyError_()` を確かめる。`processLoginAction` も先生だけを登録する
- 判定は 60 秒 cache (不許可は 5 秒)。key に版を含め、公開状態が変わったら
  `invalidateAccessDecisionCache_` で版を上げて全 viewer 分を古くする (公開終了直後の 60 秒の漏れを防ぐ)
- 共同編集者 (ほかの先生) = ボード SS の編集者。`isBoardCollaborator_` が DriveApp で編集者・持ち主を読む
  ([AccessControl.js](../src/AccessControl.js))。許可 2 分 / 不許可 10 分 cache

---

## 作ったフォームの列 ([DataApis.js](../src/DataApis.js) `templateColumnMapping_`)

- `createTemplateForm` は回答シートの列の並びを決めている: タイムスタンプ / メールアドレス / クラス / 名前 /
  種類ごとの問い / 理由。なので作った時点の割り当てを返し、回答シートに種類を developer metadata で書き添える
- `getColumnAnalysis` は書き添えがあり、見出しがその並びのままなら推定より優先する。先生が質問を足して
  列がずれていたら今までどおり推定 (`inferColumnRoles`)
- Why: 推定は列の値の形も手がかりにするので、回答 0〜数件のうちは尺度の列を数値と見分けられず、
  マトリクス・数直線のボードを公開できなかった / フォームを使う授業を始められなかった (2026-10-29 の検証)
- フォームを使う授業は、フェーズのフォームの割り当てを `phase.columnMapping` に持ち、切り替えるときは
  `applyConfigPatch_(..., { replaceKeys: ['columnMapping'] })` で丸ごと差し替える (前のシートの
  numericY などを残さない)

---

## per-row CAS lock ([ReactionService.js](../src/ReactionService.js))

- **旧**: `LockService.getScriptLock()` を process() 全体で保持 → 全 board が serialize する
- **新**: ScriptLock は ~5-10ms の critical section (cache check+put) のみ。異 row 同士は完全並列

---

## board data cache ([DataApis.js](../src/DataApis.js))

- viewer の `getPublishedSheetData` 結果を 12s ScriptCache
- reaction/highlight write 時に `bumpBoardDataVersion_` で即時 stale
- 同時に cache が切れても `withStampedeGuard_` で 1 人だけが読む (他は直前の値か BUSY)
- **cache に書く形は圧縮表現** (`packBoardDataForCache_` / `unpackBoardDataFromCache_`): 行から
  派生 3 項目 (`opinion` = answer / `isEmpty` = false / `formattedTimestamp` = timestamp から復元) を
  落とし、`Utilities.gzip` + base64 で書く。ScriptCache の 1 value 100KB 制限に 150〜250 行で当たって
  いたため (2026-10)。読むときに元の形へ戻すので wire (受け手) は変わらない

---

## users シートの書き込み省略 ([DatabaseCore.js](../src/DatabaseCore.js) `updateUser_`)

- 書き込みのたびに `USER_CACHE_VERSION` が上がり、全 viewer の user cache が一斉に落ちる
  (stampede guard で 1 read には抑えてある)
- いま読んだ行と内容が同じ (configJson は `lastAccessedAt` / `etag` を除いて比較) なら書かず、
  `{ success, unchanged: true, currentEtag }` を返す。`saveUserConfig_` は永続済の etag を client に返す

---

## データ保存則 (v2931)

先生のスプレッドシートは 1 つのルールで統一する:

> **有界で書き換わる状態 = JSON セル / 無界で追記一回きりのアーカイブ = 1 件 1 行**

| シート | 種別 | 中身 |
| ------ | ---- | ---- |
| `users` | 状態 | `configJson` (数 KB、有界) |
| `lessons` | 状態 | `lessonJson` = 授業の定義 + 遷移履歴 + **範囲ポインタ** (~4KB) |
| `lesson_responses` | アーカイブ | 1 回答 1 行 × 9 列、追記のみ |

- snapshot は `{sheet, startRow, rowCount}` のポインタを持ち、読み出しは範囲読み
  (規模に依らず 1 フェーズ分のセルのみ)。行は `lessonId + phaseIndex` を照合し、
  ポインタずれで他授業の回答が混入しない。
- 書込は `getLastRow` の次の行から 1 回の `setValues`。先生の操作 (フェーズ送り・終了) が重なっても
  範囲がかぶらないよう、その間だけ key の lock (`lesson_responses_append_lock`) を取る。
- 授業削除は `lessons` 行のみ (アーカイブ行は孤児として残る = 照合ガードで無害)。
- **アーカイブ行を lessonJson に戻さない**こと。Sheets の 1 セル 50,000 字上限に対する
  本文切り詰め (shrink) サブシステムが復活する。v2931 で 1 授業 44,698 字 → 4,198 字。
- ポインタは sheet 名を持つので、将来 `lesson_responses_2027` のような年次分割へ
  移行しても過去ポインタはそのまま読める。
- 保守: `lesson.recaptureArchive` (ボードがそのフェーズを映している間に焼き直す)。

---

## 同時実行の上限

- 全員の処理が先生 1 人の実行として数えられる。Apps Script の上限は「同時 30 本 / ユーザー」
  (公式の quota 表)。ただし 2026-10-29 の実測 (先生専用アプリ、放課後) では、ボードの読み込みを
  同時 60 本でも全部成功 (90% 値 3.5 秒)、児童用 URL を同時 30 本開いても全部開けた
  (Google 側の処理は中央値 2.8 秒)。送ったのは先生のトークンで、児童の端末・学校の Wi-Fi では未確認
- 児童の画面は授業中 5 秒 ±20% おきに更新を確かめ、議論・ふりかえりは 15 秒に落とす。失敗が続くと
  8 → 16 → 20 秒に伸ばす (止めない)
- 授業での工夫: URL は一斉に開かせず、班ごとに数秒ずらす
