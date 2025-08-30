# 未使用コード解析レポート

## 📊 概要

- **解析日時**: 2025/8/30 10:35:32
- **対象ディレクトリ**: `/Users/ryuma/Everyone-s-Answer-Board/src`
- **総ファイル数**: 43
- **未使用ファイル数**: 29 (67.4%)
- **総関数数**: 773
- **未使用関数数**: 30 (3.9%)

## 🗑️ 削除可能な未使用ファイル

- 🟡 `auth.gs` (4280 bytes, gas-backend)
- 🟡 `backendProgressSync.js.html` (8152 bytes, html-frontend)
- 🟡 `ClientOptimizer.html` (9695 bytes, html-frontend)
- 🟡 `Code.gs` (62356 bytes, gas-backend)
- 🟡 `config.gs` (6414 bytes, gas-backend)
- 🟡 `constants.gs` (39635 bytes, gas-backend)
- 🟡 `constants.js.html` (6347 bytes, html-frontend)
- 🟡 `Core.gs` (166084 bytes, gas-backend)
- 🟡 `database.gs` (103568 bytes, gas-backend)
- 🟡 `debugConfig.gs` (4803 bytes, gas-backend)
- 🟡 `ErrorBoundary.html` (22189 bytes, html-frontend)
- 🟡 `errorHandler.gs` (8170 bytes, gas-backend)
- 🟡 `errorMessages.js.html` (11455 bytes, html-frontend)
- 🟡 `lockManager.gs` (2935 bytes, gas-backend)
- 🟡 `monitoring.gs` (12039 bytes, gas-backend)
- 🟡 `page.css.html` (15658 bytes, html-frontend)
- 🟡 `Page.html` (118610 bytes, html-frontend)
- 🟡 `Registration.html` (41839 bytes, html-frontend)
- 🟡 `resilientExecutor.gs` (11570 bytes, gas-backend)
- 🟡 `secretManager.gs` (22690 bytes, gas-backend)
- 🟡 `session-utils.gs` (13106 bytes, gas-backend)
- 🟡 `setup.gs` (3529 bytes, gas-backend)
- 🟡 `SharedModals.html` (76666 bytes, html-frontend)
- 🟡 `ulog.gs` (6864 bytes, gas-backend)
- 🟡 `UnifiedCache.js.html` (942 bytes, html-frontend)
- 🟡 `unifiedCacheManager.gs` (78966 bytes, gas-backend)
- 🟡 `unifiedUtilities.gs` (9993 bytes, gas-backend)
- 🟡 `Unpublished.html` (11340 bytes, html-frontend)
- 🟡 `url.gs` (15578 bytes, gas-backend)

## 🔧 削除可能な未使用関数

- 🟡 `setupStep` (定義: constants.gs)
- 🟢 `cleanupExpired` (定義: constants.gs)
- 🟢 `dms` (定義: constants.gs)
- 🟢 `resolveValue` (定義: constants.gs)
- 🟢 `invalidateCacheForSpreadsheet` (定義: constants.gs)
- 🟢 `updateUserDatabaseField` (定義: constants.gs)
- 🟢 `clearUserCache` (定義: constants.gs)
- 🟢 `deleteAll` (定義: constants.gs)
- 🟡 `initializeSystem` (定義: constants.gs)
- 🟡 `initializeComponent` (定義: constants.gs)
- 🟡 `performInitialHealthCheck` (定義: constants.gs)
- 🟢 `performBasicHealthCheck` (定義: constants.gs)
- 🟢 `shutdown` (定義: constants.gs)
- 🟢 `fallbackBatchGet` (定義: constants.gs)
- 🟢 `batchUpdate` (定義: constants.gs)
- 🟢 `batchCache` (定義: constants.gs)
- 🟡 `setupStatus` (定義: constants.gs)
- 🟢 `sendSecurityAlert` (定義: constants.gs)
- 🟢 `checkLoginStatus` (定義: constants.gs)
- 🟢 `updateLoginStatus` (定義: constants.gs)
- 🟢 `sheetDataCache` (定義: constants.gs)
- 🟡 `checkServiceAccountConfiguration` (定義: constants.gs)
- 🟢 `validateUserScopedKey` (定義: constants.gs)
- 🟢 `performServiceAccountHealthCheck` (定義: constants.gs)
- 🟢 `argumentMapper` (定義: constants.gs)
- 🟢 `getClientId` (定義: login.js.html)
- 🟢 `handleLoginError` (定義: login.js.html)
- 🟡 `measurePerformance` (定義: Page.html, page.js.html)
- 🟢 `handleKeydown` (定義: page.js.html)
- 🟢 `testDebounce` (定義: SharedUtilities.html)

## ⚠️ 削除前の注意事項

1. **バックアップ作成**: 削除前に必ずバックアップを作成してください
2. **テスト実行**: 削除後にすべてのテストが通ることを確認してください
3. **段階的削除**: まず低リスクの項目から削除を開始してください

## 🚀 推奨削除手順

1. 低リスクファイルの削除
2. 低リスク関数の削除
3. テスト実行・動作確認
4. 中リスクファイル・関数の検討
5. 高リスク項目の慎重な検討

---

*このレポートは自動生成されました。削除前に必ず手動確認を行ってください。*
