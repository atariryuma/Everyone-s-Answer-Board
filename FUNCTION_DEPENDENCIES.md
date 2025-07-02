# StudyQuest 関数依存関係ドキュメント

## ファイル構造概要

最適化されたファイル構造と主要な関数の依存関係を整理しています。

### 主要ファイルと役割

#### エントリーポイント
- **UltraOptimizedCore.gs** - メインエントリーポイント（doGet関数）
- **Core.gs** - 主要なビジネスロジックとAPI関数

#### 機能別モジュール
- **DatabaseManager.gs** - データベース操作（ユーザー管理、データ取得）
- **AuthManager.gs** - 認証とトークン管理
- **UrlManager.gs** - URL生成と管理
- **CacheManager.gs** - キャッシュ管理
- **PerformanceOptimizer.gs** - パフォーマンス最適化
- **StabilityEnhancer.gs** - エラーハンドリングと安定性

#### レガシー・設定ファイル
- **Code.gs** - レガシーコード（廃止予定関数を含む）
- **config.gs** - 設定管理
- **SetupCode.gs** - 初期セットアップ

### HTML関数依存関係

#### Registration.html が依存している関数
```javascript
// Code.gs で定義、使用継続必須
registerNewUser(adminEmail)
quickStartSetup(userId)
verifyUserAuthentication()
getExistingBoard()
```

#### Page.html が依存している関数
```javascript
// Code.gs で定義、Core.gs で最適化版も提供
addReaction(rowIndex, reactionKey, sheetName)
getPublishedSheetData(classFilter, sortMode)
toggleHighlight(rowIndex)
```

#### AdminPanel.html が依存している関数
```javascript
// Code.gs で定義、Core.gs で最適化版も提供
getActiveFormInfo(userId)
getStatus() // deprecated but still used
publishApp(sheetName)
unpublishApp()
saveDisplayMode(mode)
getAppSettings()
```

### 関数移行状況

#### ✅ 完全に移行済み（最適化版使用推奨）
- **データベース操作**: `findUserByIdOptimized()`, `createUserOptimized()`, `updateUserOptimized()`
- **認証**: `getServiceAccountTokenCached()`
- **URL生成**: `generateAppUrlsOptimized()`, `getWebAppUrlCached()`

#### ⚠️ 移行中（両バージョン併存）
- **シート操作**: `getSheetsService()` → `getOptimizedSheetsService()`
- **キャッシュ**: `clearAllCaches()` → 最適化版
- **データ取得**: `getSheetData()` → `getSheetDataOptimized()`

#### 🔄 継続使用必須（HTML依存）
- **ユーザー登録**: `registerNewUser()`
- **リアクション**: `addReaction()`
- **データ表示**: `getPublishedSheetData()`
- **管理機能**: `publishApp()`, `unpublishApp()`, `saveDisplayMode()`

### 削除済み関数

#### 完全削除済み
- `getServiceAccountToken()` → `getServiceAccountTokenCached()`
- `doGetObsolete()` → `doGet()` in UltraOptimizedCore.gs
- `getWebAppUrlEnhanced()` → `getWebAppUrlCached()`
- `addFormQuestions()` → `addUnifiedQuestions()`
- `addDefaultQuestions()` → `addUnifiedQuestions()`
- `addSimpleQuestions()` → `addUnifiedQuestions()`
- `getAdminSettings()` → `getAppConfig()`
- `safeSpreadsheetOperation()` - 未使用

### 互換性維持

#### Core.gs の互換性関数
```javascript
// 後方互換性のため以下の関数は Core.gs で維持
function findUserById(userId) { return findUserByIdOptimized(userId); }
function findUserByEmail(email) { return findUserByEmailOptimized(email); }
function updateUserInDb(userId, updateData) { return updateUserOptimized(userId, updateData); }
function getSheetsService() { return getOptimizedSheetsService(); }
function getWebAppUrl() { return getWebAppUrlCached(); }
```

### 開発者向け注意事項

1. **新機能開発時**：最適化版関数（Optimized接尾辞）を使用してください
2. **HTML修正時**：Code.gs の関数名を変更しないでください（フロントエンドとの結合があります）
3. **パフォーマンス**：キャッシュを活用する最適化版関数を推奨します
4. **エラーハンドリング**：StabilityEnhancer.gs の機能を活用してください

### 次のステップ

1. HTML側の関数呼び出しを段階的に最適化版に移行
2. Code.gs の残存関数を最適化版で完全置換
3. レガシーコードの完全削除（HTML依存解決後）

最終更新日: 2025年7月2日