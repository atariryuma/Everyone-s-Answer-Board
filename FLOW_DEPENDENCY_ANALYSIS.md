# StudyQuest システムフロー依存関係分析

## 主要フロー分類と依存関係

### 1. エントリーポイントフロー (doGet)
**入力:**
- URLパラメータ (`e`)
- ユーザーセッション (`Session.getActiveUser().getEmail()`)
- ユーザープロパティ (`PropertiesService.getUserProperties()`)

**出力:**
- HTML出力 (SetupPage, AdminPanel, LoginPage, AppSetupPage, ErrorPage)

**副作用:**
- `clearAllExecutionCache()` → 実行レベルキャッシュ全削除
- `PropertiesService.getUserProperties()` → セッション状態保存・削除
- `verifyAdminAccess()` → 認証チェック

**依存フロー:**
- `isSystemSetup()` → システム初期化状態チェック
- `checkApplicationAccess()` → アプリケーション有効性チェック
- `parseRequestParams()` → URLパラメータ解析

---

### 2. 認証・ユーザー管理フロー
**フロー:** `verifyUserAccess()` → `getUserId()` → `getOrFetchUserInfo()` → `findUserById()`

**入力:**
- `requestUserId` (null許可)
- アクティブユーザーセッション

**出力:**
- ユーザー情報オブジェクト
- 認証成功/失敗

**副作用:**
- `clearExecutionUserInfoCache()` → ユーザー情報キャッシュクリア
- 統一キャッシュマネージャー更新 (`cacheManager.remove()`)

**キャッシュ階層:**
1. 実行レベルキャッシュ (`executionUserInfoCache`)
2. 統一キャッシュマネージャー (`cacheManager`)
3. Google Apps Script キャッシュ (`CacheService`)

---

### 3. データ取得・表示フロー (getInitialData)
**フロー:** `getInitialData()` → `fixUserDataConsistency()` → DB取得 → フロントエンド表示

**入力:**
- `requestUserId`
- `targetSheetName`

**出力:**
- 初期化データオブジェクト (ボード設定、シートデータ、統計情報)

**副作用:**
- `clearExecutionUserInfoCache()` → 古いキャッシュクリア
- `fixUserDataConsistency()` → データ整合性自動修復
- 診断・修復フロー実行

**競合ポイント:**
- データベースアクセス中の競合
- キャッシュ更新競合

---

### 4. 設定保存・公開フロー (saveAndPublish)
**フロー:** `saveAndPublish()` → `createExecutionContext()` → DB更新 → キャッシュ管理

**入力:**
- `requestUserId`
- `sheetName`
- `config` (設定オブジェクト)

**出力:**
- 保存結果オブジェクト

**副作用:**
- **排他制御:** `LockService.getScriptLock()` (30秒タイムアウト)
- **段階的処理:**
  1. コンテキスト作成
  2. インメモリ更新
  3. DB書き込み前キャッシュクリア
  4. DB一括更新
  5. キャッシュウォーミング

**重要な設計:**
- トランザクション的処理
- キャッシュ無効化 → DB更新 → キャッシュ再構築の順序

---

### 5. リアクション処理フロー
**フロー:** `addReaction()` → `processReaction()` → DB更新 → リアクション状態返却

**入力:**
- `requestUserId`
- `rowIndex`
- `reactionKey`
- `sheetName`

**出力:**
- リアクション更新結果

**副作用:**
- `clearExecutionUserInfoCache()` → ユーザー情報キャッシュクリア
- **排他制御:** `LockService.getScriptLock()` (10秒タイムアウト)
- 重複リアクションチェック・削除

**フロントエンド連携:**
- リアクションキュー (`processReactionQueue()`)
- 楽観的UI更新 → サーバー確認 → 状態同期

---

### 6. クイックスタートフロー
**フロー:** `quickStartSetup()` → `initializeQuickStartContext()` → `createQuickStartFiles()` → `updateQuickStartDatabase()`

**入力:**
- `requestUserId`

**出力:**
- セットアップ完了情報

**副作用:**
- Google Drive ファイル作成
- スプレッドシート・フォーム作成
- データベース更新
- 複数回の `clearExecutionUserInfoCache()`

**段階的処理:**
1. コンテキスト初期化
2. ファイル作成とフォルダ管理
3. データベース更新とキャッシュ管理

---

## 🚨 発見された潜在的競合ポイント

### 1. キャッシュ競合
- **頻繁なキャッシュクリア:** 30箇所以上で `clearExecutionUserInfoCache()` 呼び出し
- **複数キャッシュ層:** 実行レベル、統一、GAS キャッシュの非同期更新
- **競合シナリオ:** 同時リクエストでのキャッシュ無効化競合

### 2. ロック競合
- **異なるタイムアウト値:**
  - `saveAndPublish`: 30秒
  - `processReaction`: 10秒
  - その他DB操作: 15秒, 10秒
- **デッドロック可能性:** 複数ロック取得順序が不定

### 3. セッション状態競合
- **PropertiesService依存:** 複数フローで同時アクセス
- **セッション切り替え:** アカウント変更時の状態不整合

### 4. データベースアクセス競合
- **排他制御なしの操作:** 一部の読み取り専用操作
- **トランザクション境界不明確:** 部分更新時のデータ不整合リスク