# みんなの回答ボード - サイロ型マルチテナントシステム要件定義

## システムアーキテクチャ

**基盤**: Google Apps Script V8 Runtime
**データストア**: Google Sheets (マルチテナント対応)
**認証**: Google OAuth2
**フロントエンド**: HTML/CSS/JavaScript (外部ライブラリ不使用)

## コアクラス構造

```javascript
// エラー処理
UnifiedErrorHandler
SystemIntegrationManager
MultiTenantSecurityManager

// キャッシュ管理
CacheManager
UnifiedExecutionCache
ManagedExecutionContext

// データ処理
UnifiedBatchProcessor
ResilientExecutor
UnifiedSecretManager

// ユーティリティ
UnifiedUserManager
UnifiedURLManager
UnifiedAPIClient
UnifiedValidation
```

## グローバル定数・変数

```javascript
// エラー定義
ERROR_SEVERITY: {CRITICAL,HIGH,MEDIUM,LOW,INFO}
ERROR_CATEGORIES: {SYSTEM,DATA,USER,EXTERNAL,CONFIG,SECURITY}

// 表示モード
DISPLAY_MODES: {ANONYMOUS,NAMED,EMAIL}

// キャッシュ設定
USER_CACHE_TTL: 300
EXECUTION_MAX_LIFETIME: 300000
LOCK_TIMEOUTS: {standard,critical,batch}

// カラム定義
COLUMN_HEADERS: {TIMESTAMP,CLASS,NAME,EMAIL,OPINION,REASON}

// スコア設定
SCORING_CONFIG: {weights,bonuses}
REACTION_KEYS: ['UNDERSTAND','LIKE','CURIOUS']
```

## 主要関数マップ

### エントリーポイント
- `doGet(e)` → `routeRequestByMode(params)` → 各ハンドラー

### ルーティング
- `handleDefaultRoute()` - デフォルト
- `handleLoginMode()` - ログイン画面
- `handleAppSetupMode()` - セットアップ画面
- `handleAdminMode(params)` - 管理画面
- `handleViewMode(params)` - 閲覧画面

### ユーザー管理
- `findUserById(userId)` - ID検索
- `findUserByEmail(email)` - メール検索
- `createUser(userData)` - 新規作成
- `updateUser(userId,data)` - 更新
- `createOrUpdateUser(email,additionalData)` - 作成/更新
- `verifyUserAccess(requestUserId)` - アクセス検証

### データ取得
- `getPublishedSheetData(requestUserId,classFilter,sortOrder,adminMode,bypassCache)`
- `executeGetPublishedSheetData(requestUserId,classFilter,sortOrder,adminMode)`
- `getSheetData(requestUserId,sheetName,classFilter,sortMode,adminMode)`
- `executeGetSheetData(userId,sheetName,classFilter,sortMode,adminMode)`

### セットアップ
- `initializeQuickStartContext(requestUserId)` - クイックスタート初期化
- `createFormAndSheet(email,userId,config)` - フォーム/シート作成
- `executeCustomSetup(requestUserId,config)` - カスタムセットアップ
- `autoMapHeaders(headers,sheetName)` - ヘッダー自動マッピング

### 設定管理
- `saveSheetConfig(userId,spreadsheetId,sheetName,config,options)`
- `switchToSheet(userId,spreadsheetId,sheetName,options)`
- `setDisplayOptions(requestUserId,displayOptions,options)`
- `saveAndActivateSheet(requestUserId,spreadsheetId,sheetName,config)`

### フォーム管理
- `createGoogleForm(email,config)` - フォーム作成
- `updateFormTitle(requestUserId,title,description)` - タイトル更新
- `detectFormUrlFromSpreadsheet(spreadsheetId)` - フォームURL検出

### キャッシュ管理
- `cacheManager.get(key,generator,options)` - 取得
- `cacheManager.set(key,value,options)` - 設定
- `cacheManager.invalidate(patterns)` - 無効化
- `clearExecutionUserInfoCache()` - 実行キャッシュクリア

### セキュリティ
- `getServiceAccountTokenCached()` - SAトークン取得
- `validateConfigSecurity(config)` - 設定検証
- `performComprehensiveSecurityHealthCheck()` - セキュリティ監査

## データ構造

### userInfo (データベースレコード)
```javascript
{
  userId: string,           // 一意識別子
  adminEmail: string,       // 管理者メール
  spreadsheetId: string,    // スプレッドシートID
  configJson: string,       // JSON設定文字列
  createdAt: ISO8601,
  lastAccessedAt: ISO8601,
  isActive: boolean,
  deletedAt: ISO8601|null
}
```

### configJson (解析後)
```javascript
{
  // 基本設定
  setupStatus: 'pending'|'completed',
  ownerId: string,
  
  // 公開設定
  appPublished: boolean,
  publishedSpreadsheetId: string,
  publishedSheetName: string,
  
  // フォーム設定
  formCreated: boolean,
  formUrl: string,
  editFormUrl: string,
  
  // 表示設定
  displayMode: 'ANONYMOUS'|'NAMED'|'EMAIL',
  showReactionCounts: boolean,
  savedClassChoices: string[],
  
  // シート固有設定 (sheet_${sheetName})
  'sheet_xxx': {
    timestampHeader: string,
    classHeader: string,
    nameHeader: string,
    emailHeader: string,
    opinionHeader: string,
    reasonHeader: string,
    guessedConfig: object,
    lastModified: ISO8601
  }
}
```

## マルチテナントフロー

### 1. 初回セットアップ
```
ユーザーアクセス → doGet → handleDefaultRoute
→ 新規ユーザー作成 → handleAppSetupMode
→ initializeQuickStartContext → createFormAndSheet
→ データベース初期化 → 管理画面リダイレクト
```

### 2. 管理者アクセス
```
doGet(mode=admin,uid=xxx) → handleAdminMode
→ verifyUserAccess → findUserById
→ renderAdminPanel → 管理画面表示
```

### 3. 閲覧者アクセス
```
doGet(mode=view,uid=xxx) → handleViewMode
→ verifyUserAccess → getPublishedSheetData
→ renderAnswerBoard → 回答ボード表示
```

### 4. データ更新
```
google.script.run.saveSheetConfig → verifyUserAccess
→ updateUser → キャッシュ無効化
→ 成功レスポンス
```

## キャッシュ戦略

### 階層構造
1. **実行キャッシュ** (メモリ): 単一実行内
2. **スクリプトキャッシュ** (6時間): CacheService.getScriptCache()
3. **ユーザーキャッシュ** (無期限): PropertiesService.getUserProperties()
4. **永続ストレージ**: Google Sheets

### キャッシュキー命名
- ユーザー情報: `user_${userId}`
- シートデータ: `publishedData_${userId}_${ssId}_${sheet}_${filter}_${sort}`
- サービストークン: `SA_TOKEN_CACHE`
- 設定: `config_${userId}_${sheetName}`

## エラー処理パターン

### UnifiedErrorHandler
```javascript
try {
  // 処理
} catch (error) {
  logError(error, functionName, severity, category, metadata);
  throw createStructuredError(message, code, details);
}
```

### ResilientExecutor
```javascript
resilientExecutor.execute(operation, {
  maxRetries: 3,
  retryDelay: 1000,
  exponentialBackoff: true
});
```

## セキュリティ要件

### アクセス制御
- `verifyUserAccess(requestUserId)` - 全API必須
- `isOwner` チェック - 管理機能
- `appPublished` チェック - 公開状態

### データ分離
- userId基準の完全分離
- configJson内で独立管理
- クロステナントアクセス禁止

## パフォーマンス最適化

### バッチ処理
- `batchGetSheetsData()` - 複数範囲同時取得
- `batchUpdateSheetsData()` - 複数更新同時実行
- `UnifiedBatchProcessor` - 統合バッチ管理

### 実行時間管理
- 5分制限内での処理完了
- `EXECUTION_MAX_LIFETIME` チェック
- 長時間処理の分割実行

## デプロイメント

### 必須設定 (PropertiesService.getScriptProperties())
```javascript
SERVICE_ACCOUNT_CREDS  // サービスアカウント認証情報
DATABASE_SPREADSHEET_ID  // データベースID
CLIENT_ID  // OAuth クライアントID (オプション)
DEBUG_MODE  // デバッグモード (true/false)
```

### トリガー設定
- `performPeriodicMaintenance()` - 1時間ごと
- `cleanupOldLogs()` - 1日ごと

## 監視・診断

### ヘルスチェック
- `performHealthCheck()` - システム健全性
- `performPerformanceCheck()` - パフォーマンス
- `performSecurityCheck()` - セキュリティ
- `diagnoseDatabase(targetUserId)` - データベース診断

### ログ管理
- `infoLog()`, `warnLog()`, `errorLog()`, `debugLog()`
- 診断ログ: DIAGNOSTIC_LOGS シート
- 削除ログ: DELETE_LOGS シート

## クライアント連携

### google.script.run 主要API
```javascript
// データ取得
getPublishedSheetData(userId, classFilter, sortOrder, adminMode)
getAppConfig(userId)
getActiveFormInfo(userId)

// 設定管理
saveSheetConfig(userId, spreadsheetId, sheetName, config)
setDisplayOptions(userId, displayOptions)
switchToSheet(userId, spreadsheetId, sheetName)

// セットアップ
executeQuickStartSetup(userId)
executeCustomSetup(userId, config)

// ユーザー管理
updateUserActiveStatus(userId, isActive)
getCurrentUserStatus(userId)
```

## 開発規約

### 命名規則
- 関数: `camelCase`
- 定数: `UPPER_SNAKE_CASE`
- クラス: `PascalCase`
- プライベート変数: `_prefixedCamelCase`

### コード構造
- 早期リターン推奨
- エラーは構造化して投げる
- JSDoc必須
- 日本語コメント可

### テスト要件
- 単体テスト: Jest
- 統合テスト: 実環境検証
- セキュリティ監査: 定期実行