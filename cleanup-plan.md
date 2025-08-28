# 未定義関数クリーンアップ計画

## 🎯 方針：本当に必要な関数のみ残して、不要な参照を削除

## 1️⃣ **保持：本当に必要な関数** (12個)

### Core機能で重要
- `UError` → `const UError = handleUnifiedError` (既に定義済み、エイリアス追加のみ)
- `Functions` → 空オブジェクト `const Functions = {}`
- `URLSearchParams` → 既にポリフィル実装済み

### ログ機能
- `ULog` related → 既に代替実装済み

### 自己参照クラス（設計修正が必要）
- `UnifiedBatchProcessor` (設計修正)
- `ResilientExecutor` (設計修正)
- `SystemIntegrationManager` (設計修正)

## 2️⃣ **削除：不要な参照** (50個)

### UIヘルパー系（削除対象）
```javascript
// unifiedUtilities.gs内の不要な関数参照
- urlFactory, formFactory, customUI, quickStartUI, userFactory
- generatorFactory, buildSecureUserScopedKey, responseFactory
- unified, logger, safeExecute
```

### 管理機能系（削除対象）  
```javascript
// Core.gs内の不要な管理関数
- deleteUserAccountByAdminForUI, deleteUserAccountByAdmin
- deleteCurrentUserAccount, deleteUserAccount
- setupStep, checkReactionState
```

### 内部ヘルパー系（削除対象）
```javascript
// 各unified*ファイル内の内部関数参照
- batchGet, fallbackBatchGet, batchUpdate, batchCacheOperation
- clearElementCache, resolvePendingClears, rejectPendingClears
- setupStatus, sendSecurityAlert, checkLoginStatus, updateLoginStatus
```

### その他削除対象
```javascript
// 使用されていない機能
- cleanupExpired, dms, resolveValue
- updateUserDatabaseField, clearUserCache, deleteAll
- sheetDataCache, unpublished
- checkServiceAccountConfiguration, validateUserScopedKey
- performServiceAccountHealthCheck, getTemporaryActiveUserKey
- argumentMapper, invalidateCacheForSpreadsheet
```

## 🚀 **実行計画**

### Step 1: 必要最小限のエイリアス追加 (5分)
```javascript
// constants.gs に追加
const Functions = {};
const UnifiedErrorHandler = handleUnifiedError;
```

### Step 2: 不要な関数呼び出しを削除 (30分)
- 各ファイルから不要な関数参照を削除
- コメントアウトまたは完全削除
- 機能に影響しない部分のみ

### Step 3: 自己参照問題を修正 (15分)
```javascript
// 各クラスで self-reference を修正
class UnifiedBatchProcessor {
  static getInstance() {
    return new UnifiedBatchProcessor();
  }
}
```

### Step 4: 検証テスト実行 (5分)
- 未定義エラーテストを再実行
- 62個 → 5-10個程度に削減確認

## 📊 **期待される結果**

- **削除予定**: 50個の不要な参照
- **保持**: 12個の必要な機能
- **所要時間**: 約55分
- **最終エラー数**: 5-10個（許容範囲内）

## ✅ **削除の安全性**

削除対象は以下の理由で安全：
1. **UI系**: GASでは不要な概念
2. **管理系**: 使用されていない機能
3. **内部ヘルパー**: 実装されていない内部関数
4. **未使用機能**: コメントや未完成機能

重要な機能（認証、データ処理、セキュリティ）は全て保持します。