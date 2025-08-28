# GAS重複宣言・参照エラー修正プラン

## Phase 1: 重複関数削除（14個の重複を解決）

### 削除対象ファイル別

**config.gs から削除:**
- `performGlobalMemoryCleanup` (重複)
- `clearExecutionUserInfoCache` → unifiedCacheManager.gsの実装を使用

**database.gs から削除:**
- `batchUpdateSpreadsheet` → unifiedBatchProcessor.gsの実装を使用
- `fetchUserFromDatabase` → unifiedUserManager.gsの実装を使用
- `findUserByEmail` → unifiedUserManager.gsの実装を使用
- `getUserWithFallback` → unifiedUserManager.gsの実装を使用

**Core.gs から削除:**
- `getCachedSheetsService` → unifiedCacheManager.gsの実装を使用
- `addServiceAccountToSpreadsheet` → unifiedSecurityManager.gsの実装を使用
- `shareSpreadsheetWithServiceAccount` → unifiedSecurityManager.gsの実装を使用

**secretManager.gs から削除:**
- `getSecureServiceAccountCreds` → unifiedSecurityManager.gsの実装を使用
- `getSecureDatabaseId` → unifiedSecurityManager.gsの実装を使用

**unifiedSecurityManager.gs から削除:**
- `getSecureUserInfo` → unifiedUserManager.gsの実装を使用
- `getSecureUserConfig` → unifiedUserManager.gsの実装を使用

**unifiedCacheManager.gs から削除:**
- `invalidateUserCache` → unifiedUserManager.gsの実装を使用

## Phase 2: 参照エラー修正（296個の潜在的エラー）

### 主要な未定義参照:
1. `debugLog`, `infoLog`, `warnLog` → ULog統一ログ関数に統一
2. `createSuccessResponse`, `createErrorResponse` → constants.gsの実装使用
3. `buildUserScopedKey`, `buildSecureUserScopedKey` → 実装確認・統合
4. `generatorFactory`, `responseFactory` → UUtilitiesクラス内プロパティとして整理
5. GAS組み込み関数の適切な参照確認

## Phase 3: 段階的修正アプローチ

### Step 1: 重複関数削除（優先度：高）
- 統合済みファイルから重複を削除
- deprecated関数はコメントアウト化

### Step 2: 参照エラー修正（優先度：中）
- undefined関数への参照をunified実装に置換
- ログ関数の統一

### Step 3: テスト・検証（優先度：高）
- 各修正後にGASプッシュテスト
- 機能テスト実行