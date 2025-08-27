# みんなの回答ボード - Claude Code AI開発対応システム

🤖 **Claude Code AI開発環境完全対応プロジェクト**

## 🚀 クイックスタート

### 新機能開発（推奨ワークフロー）
```bash
# 1. 新機能開発開始
./scripts/new-feature.sh "機能名"

# 2. 別ターミナルでテスト監視
npm run test:watch

# 3. Claude Codeで「feature/機能名 ブランチで開発します」
#    → TDD→実装→修正のサイクル

# 4. 開発完了後
./scripts/merge-feature.sh "機能名"
```

### 緊急修正
```bash
./scripts/quick-fix.sh      # 修正作業
./scripts/hotfix-deploy.sh  # 即座にデプロイ
```

## 🏗️ システムアーキテクチャ

**基盤**: Google Apps Script V8 Runtime + Claude Code AI開発環境
**テスト**: Jest + GAS API完全モック
**品質管理**: ESLint + Prettier + 自動化CI/CD
**開発フロー**: TDD + feature/main ブランチ戦略
**データストア**: Google Sheets（テナント分離型）
**認証**: Google OAuth2（email-based ownership）
**フロントエンド**: HTML/CSS/JavaScript
**マルチテナンシー**: サイロ型（テナントごとに独立スプレッドシート）

## 🧪 Claude Code AI開発環境

### テスト環境
```javascript
// Jest + GAS API完全モック環境
SpreadsheetApp     // モック実装済み
PropertiesService  // モック実装済み  
CacheService       // モック実装済み
UrlFetchApp       // モック実装済み
HtmlService       // モック実装済み
Utilities         // モック実装済み
+ 全GAS API対応
```

### 品質管理ツール
```bash
npm run test        # Jest単体テスト
npm run test:watch  # テスト監視モード（開発時推奨）
npm run lint        # ESLint品質チェック
npm run format      # Prettier自動整形
npm run check       # 統合品質チェック
npm run deploy      # テスト→GASデプロイ
```

### 自動化スクリプト
```bash
./scripts/new-feature.sh <名前>     # feature開発開始
./scripts/merge-feature.sh <名前>   # マージ・デプロイ
./scripts/quick-fix.sh              # 緊急修正モード
./scripts/hotfix-deploy.sh          # ホットフィックスデプロイ
./scripts/cleanup-branches.sh       # ブランチ整理
```

## 🏗️ 実装済み中核システム

### エラー処理 
```javascript
UnifiedErrorHandler        // 実装済み - 構造化ログ管理
ManagedExecutionContext   // 実装済み - 実行コンテキスト管理
```

### キャッシュ管理
```javascript
cacheManager              // 実装済み - シンプルなキャッシュ管理
globalContextManager     // 実装済み - グローバルコンテキスト
```

### データ処理
```javascript
unifiedBatchProcessor     // 実装済み - 基本バッチ処理
```

## 実際のデータベース構造

### Users シート（データベース）
```javascript
{
  userId: string,           // 一意識別子（MD5ハッシュ）
  adminEmail: string,       // 管理者メールアドレス
  spreadsheetId: string,    // テナント専用スプレッドシートID
  spreadsheetUrl: string,   // 直接URLリンク
  createdAt: ISO8601,       // 作成タイムスタンプ
  configJson: string,       // JSON設定文字列
  lastAccessedAt: ISO8601,  // 最終アクセス時刻
  isActive: boolean         // アクティブ状態
}
```

### configJson（実際の構造）
```javascript
{
  // セットアップ状態
  setupStatus: 'pending'|'completed',
  ownerId: string,
  
  // 公開設定
  appPublished: boolean,
  publishedSpreadsheetId: string,
  publishedSheetName: string,
  
  // フォーム連携
  formCreated: boolean,
  formUrl: string,
  editFormUrl: string,
  
  // 表示設定
  displayMode: 'ANONYMOUS'|'NAMED'|'EMAIL',
  showReactionCounts: boolean,
  savedClassChoices: string[],
  
  // 動的シート設定（シートごと）
  'sheet_${sheetName}': {
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

## 実装済みルーティングシステム

### エントリーポイント
```
doGet(e) → parseRequestParams(e) → routeRequestByMode(params)
├── handleDefaultRoute() - 自動ログイン・セットアップリダイレクト
├── handleLoginMode() - ログイン画面表示
├── handleAppSetupMode() - 初期設定画面
├── handleAdminMode(params) - 管理インターフェース
└── handleViewMode(params) - 公開回答ボード
```

### 実際のアクセス制御
```javascript
verifyUserAccess(requestUserId)  // 中核アクセス制御
verifyAdminAccess(userId)        // 管理者権限チェック
// メール所有権検証: activeUserEmail === userInfo.adminEmail
```

## 中核機能マップ（実装済み）

### ユーザー管理
- `findUserById(userId)` - プライマリ検索
- `findUserByEmail(email)` - メールベース検索
- `createUser(userData)` - 新規テナント作成
- `updateUser(userId, updateData)` - レコード更新
- `getUserInfoCached(requestUserId)` - キャッシュ付き取得

### データ取得・処理
- `getPublishedSheetData(requestUserId, classFilter, sortOrder, adminMode, bypassCache)` - メインデータ取得
- `executeGetPublishedSheetData()` - 生データ処理
- `getSheetData(userId, sheetName, classFilter, sortMode, adminMode)` - シート別データ

### リアクション機能
- `REACTION_KEYS: ['UNDERSTAND', 'LIKE', 'CURIOUS']` - 定義済みリアクション
- `addReaction(requestUserId, rowIndex, reactionKey, sheetName)` - リアクション追加
- `addReactionBatch(requestUserId, reactionList)` - 一括リアクション処理

### セットアップ・設定
- `executeQuickStartSetup(requestUserId)` - 自動セットアップ
- `initializeQuickStartContext()` - セットアップ初期化
- `saveSheetConfig()` - 設定保存
- `switchToSheet()` - シート切り替え
- `autoMapHeaders()` - ヘッダー自動検出

### セキュリティ・認証
- `getServiceAccountTokenCached()` - サービスアカウントトークン取得
- `generateNewServiceAccountToken()` - 新規トークン生成
- `getServiceAccountEmail()` - サービスアカウントメール取得
- `setupServiceAccount(serviceAccountJson)` - サービスアカウント設定
- `performServiceAccountHealthCheck()` - SAヘルスチェック

## 実際のキャッシュ戦略

### キャッシュ階層（実装済み）
1. **実行キャッシュ**（メモリ）- 単一実行内
2. **スクリプトキャッシュ**（6時間）- CacheService.getScriptCache()
3. **ユーザープロパティ**（永続）- PropertiesService

### 実際のキャッシュキー
```javascript
'user_${userId}'                                    // ユーザー情報
'publishedData_${userId}_${ssId}_${sheet}_${filter}_${sort}_${adminMode}' // データキャッシュ
'service_account_token'                             // 認証トークン
'CURRENT_USER_ID_${hash}'                          // セッション情報
```

### キャッシュAPI（実装済み）
```javascript
cacheManager.get(key, valueFn, options)  // 取得・生成
cacheManager.remove(key)                 // 削除
cacheManager.clearByPattern(pattern)     // パターン削除（簡易実装）
```

## サイロ型マルチテナント実装

### テナント分離メカニズム
- **データ分離**: userId（MD5）による完全分離
- **スプレッドシート分離**: テナントごとに独立Google Spreadsheet
- **アクセス制御**: `verifyUserAccess()`による全API保護
- **セッション管理**: Google Account単位のプロパティ管理

### セキュリティ実装
- **メール所有権**: `activeUserEmail === userInfo.adminEmail`
- **リクエスト検証**: 全API関数で`requestUserId`検証必須
- **キャッシュセキュリティ**: ユーザースコープキーによる分離
- **データベース分離**: 全クエリで`userId`フィルタ

## Google API連携（実装済み）

### 統合サービス
- **Google Sheets API v4**: 完全CRUD操作
- **Google Forms API**: フォーム作成・リンク
- **Google Drive API**: フォルダ作成・権限管理
- **Google OAuth2**: ScriptApp経由セッション管理

### 主要API関数
```javascript
// データ操作
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

## エラー処理実装

### UnifiedErrorHandler（実装済み）
```javascript
try {
  // 業務処理
} catch (error) {
  UnifiedErrorHandler.logError(error, functionName, severity, category, metadata);
  throw UnifiedErrorHandler.createStructuredError(message, code, details);
}
```

### ログレベル
```javascript
ERROR_SEVERITY: {CRITICAL: 0, HIGH: 1, MEDIUM: 2, LOW: 3, INFO: 4}
ERROR_CATEGORIES: {SYSTEM, DATA, USER, EXTERNAL, CONFIG, SECURITY}
```

## リアクション機能詳細

### 実装済みリアクション
```javascript
REACTION_KEYS: ['UNDERSTAND', 'LIKE', 'CURIOUS']
```

### データ構造
- **カラムベース**: Google Sheetsの専用列に反応数保存
- **リアルタイム更新**: `addReactionBatch()`による一括更新
- **表示制御**: `showReactionCounts`設定による表示/非表示

### リアクションフロー
```
ユーザーリアクション → addReaction() → 
スプレッドシート更新 → キャッシュ無効化 → 
UI更新 → リアルタイム反映
```

## パフォーマンス最適化

### 実装済み最適化
- **基本バッチ処理**: `unifiedBatchProcessor`による効率化
- **実行時間管理**: 5分制限内での処理完了保証
- **キャッシュ活用**: 頻繁データへの高速アクセス
- **遅延ロード**: 必要時のみデータ取得

## 必須環境設定

### PropertiesService.getScriptProperties()
```javascript
SERVICE_ACCOUNT_CREDS   // サービスアカウント認証情報（オプション）
DATABASE_SPREADSHEET_ID // Users データベースID
CLIENT_ID               // OAuth クライアントID（オプション）
DEBUG_MODE              // デバッグモード（true/false）
```

## 監視・診断（実装済み）

### ログ管理
- `infoLog()`, `warnLog()`, `errorLog()`, `debugLog()` - 構造化ログ
- Users シートでのアクティビティログ
- エラー情報の永続化

## 開発・運用規約

### 命名規則（実装準拠）
- **関数**: `camelCase`
- **定数**: `UPPER_SNAKE_CASE`
- **オブジェクト**: `camelCase`（クラス風でも小文字開始）

### アーキテクチャ原則
- **関数ベース設計**: クラスよりも関数中心
- **早期リターン**: ネストを避ける
- **エラーハンドリング**: UnifiedErrorHandler活用
- **メール所有権**: 全アクセス制御の基盤

## 実際の制約・限界

### 現在の制約
- **クロステナント検索不可**: 完全なサイロ分離
- **グローバルキャッシュクリア**: テナントスコープ限定なし
- **シンプルな権限モデル**: オーナーのみアクセス可能
- **基本的なバッチ処理**: 高度な最適化なし

### スケーラビリティ
- **テナント数**: Google Apps Script実行制限内
- **データ量**: Google Sheets行数制限（200万行）
- **同時実行**: GAS並行実行制限
- **API制限**: Google API呼び出し制限

## サマリー

このシステムは、**サイロ型マルチテナント方式のGoogle連携回答ボード**として、以下を実現：

- **完全テナント分離**: ユーザーごと独立スプレッドシート
- **リアクション機能**: UNDERSTAND/LIKE/CURIOUS反応システム
- **管理機能**: テナント管理者による完全制御
- **Google統合**: Sheets/Forms/Drive/OAuth2完全連携
- **キャッシュ最適化**: 高速レスポンス保証
- **自動セットアップ**: ワンクリック環境構築

実装は理論より実用性重視で、**動作する最小viable システム**として設計されています。