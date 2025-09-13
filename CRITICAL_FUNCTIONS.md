# 🔒 CRITICAL FUNCTIONS - 削除禁止機能リスト

> **⚠️ IMPORTANT: リファクタリング時に絶対に削除してはいけない機能**
> 
> このファイルは改善実装中に重要機能を誤削除から保護するためのものです。

## 📋 削除禁止 - 46個の復旧済み重要関数

### 🔑 認証・ログイン機能 (6個)
```javascript
// main.gs - 絶対削除禁止
function getUser(kind = 'email') {
  // 戻り値: string(email) | Object(userInfo)
  // 呼び出し: login.js.html, SetupPage.html, AdminPanel.js.html
}

function processLoginAction() {
  // 統合ログイン処理（登録/検証/遷移）
  // 戻り値: {success, userId, email, needsSetup, redirectUrl}
}

function forceUrlSystemReset() {
  // URL内部状態リセット（ログ記録用）
  // 戻り値: {success, message, redirectUrl}
}

function createRedirect(url) {
  // X-Frame対応のJSリダイレクト
  // 戻り値: HtmlOutput with redirect script
}

function confirmUserRegistration() {
  // ユーザー認証確認（シンプル版）
}

function verifyUserAuthentication() {
  // SharedUtilities.html から呼び出し
}
```

### ⚙️ セットアップ・設定機能 (15個)
```javascript
// main.gs - 絶対削除禁止
function setupApplication(serviceAccountJson, databaseId, adminEmail, googleClientId) {
  // 初期セットアップ処理
  // 重要: Service Account設定、DB ID設定、管理者設定
}

function testSetup() {
  // セットアップテスト実行
  // 戻り値: {success, results{serviceAccount, database, adminEmail, webAppUrl}}
}

function getApplicationStatusForUI() {
  // アプリケーション状態取得（AppSetupPage.html用）
  // 戻り値: {systemStatus, userStatus, completionRate, nextSteps}
}

function publishApplication(publishConfig) {
  // アプリケーション公開
  // 重要: appPublished=true, publishedAt設定
}

function getCurrentBoardInfoAndUrls() {
  // 現在のボード情報とURL取得（管理パネル用）
  // 戻り値: {boardInfo, urls{viewUrl, adminUrl, debugUrl}}
}

function setApplicationStatusForUI(enabled) {
  // アプリケーション状態切り替え（有効/無効）
}

function saveDraftConfiguration(config) {
  // 設定の下書き保存
}
```

### 📊 データ操作機能 (12個)
```javascript
// main.gs - 絶対削除禁止
function getSpreadsheetList() {
  // スプレッドシート一覧取得（Drive API使用）
  // 戻り値: {success, spreadsheets[{id, name, url, lastUpdated}]}
}

function getSheetList(spreadsheetId) {
  // シート一覧取得
  // 戻り値: {success, sheets[{name, index, rowCount, columnCount}]}
}

function analyzeColumns(spreadsheetId, sheetName) {
  // 列分析（ヘッダー推測、サンプル取得）
  // 戻り値: {success, columns[{index, header, type, samples}]}
}

function getPublishedSheetData(userId, options = {}) {
  // 公開シートデータ取得
  // 重要: メインデータ取得API
}

function getHeaderIndices(spreadsheetId, sheetName) {
  // ヘッダーインデックス取得（列マッピング支援）
}

function validateAccess(spreadsheetId) {
  // スプレッドシートアクセス検証
}

function addSpreadsheetUrl(url) {
  // スプレッドシートURL追加（Unpublished.html用）
}
```

### 👥 システム管理機能 (8個)
```javascript
// main.gs - 絶対削除禁止（システム管理者専用）
function getAllUsersForAdminForUI(options = {}) {
  // システム管理者専用：全ユーザー取得
  // 重要: 権限チェック必須
}

function deleteUserAccountByAdminForUI(targetUserId) {
  // システム管理者専用：ユーザーアカウント削除
  // 重要: セキュリティログ記録、自己削除防止
}

function forceLogoutAndRedirectToLogin() {
  // 強制ログアウト・リダイレクト
}

function getDeletionLogsForUI(options = {}) {
  // 削除ログ取得（監査用）
}

function checkIsSystemAdmin() {
  // システム管理者判定
  // 戻り値: boolean
}
```

### 🛠️ ユーティリティ・診断機能 (5個)
```javascript
// main.gs - 絶対削除禁止
function getLoginStatus() {
  // ユーザーログインステータス取得
  // 戻り値: {isLoggedIn, userEmail, userId, isActive, isSystemAdmin}
}

function getSystemStatus() {
  // システムステータス取得
  // 戻り値: {system, user, statistics, timestamp}
}

function getSystemDomainInfo() {
  // システムドメイン情報取得
}

function reportClientError(errorInfo) {
  // クライアントエラー報告（ErrorBoundary.html用）
}

function testForceLogoutRedirect() {
  // 強制ログアウトテスト
}
```

## 🎯 POST/GET ハンドラー - 削除禁止

### HTTP リクエストハンドラー
```javascript
// main.gs - 絶対削除禁止
function doGet(e) {
  // GET リクエストハンドラー
  // 重要: ルーティング、認証、セキュリティチェック
}

function doPost(e) {
  // POST リクエストハンドラー  
  // 重要: データ操作API (getData, addReaction, toggleHighlight, refreshData)
}

function handleGetData(request) {
  // データ取得ハンドラー
}

function handleAddReaction(request) {
  // リアクション追加ハンドラー
}

function handleToggleHighlight(request) {
  // ハイライト切り替えハンドラー
}

function handleRefreshData(request) {
  // データ更新ハンドラー
}
```

## 📋 Services Layer - 保護対象機能

### DataService.gs
```javascript
// src/services/DataService.gs - 削除禁止機能
const DataService = {
  getSheetData(userId, options),           // メインデータ取得
  fetchSpreadsheetData(config, options),   // スプレッドシート取得
  processRawData(dataRows, headers),       // データ変換・フィルタリング
  addReaction(userId, rowId, reactionType), // リアクション追加
  toggleHighlight(userId, rowId),          // ハイライト切り替え
  getBulkData(userId, options),            // バルクデータ取得
  processReaction(...),                    // レガシー対応
  processHighlightToggle(...),             // レガシー対応
  deleteAnswer(userId, rowIndex, sheetName) // 回答削除
}
```

### UserService.gs
```javascript
// src/services/UserService.gs - 削除禁止機能
const UserService = {
  getCurrentEmail(),                       // 現在ユーザーメール取得
  getCurrentUserInfo(),                    // ユーザー情報取得（キャッシュ対応）
  isSystemAdmin(),                         // システム管理者判定
  createUser(email),                       // ユーザー作成
  diagnose()                              // サービス診断
}
```

### ConfigService.gs
```javascript
// src/services/ConfigService.gs - 削除禁止機能
const ConfigService = {
  getUserConfig(userId),                   // ユーザー設定取得
  saveUserConfig(userId, config),         // ユーザー設定保存
  hasCoreSystemProps(),                   // システムプロパティ確認
  isSystemSetup(),                        // セットアップ状態確認
  diagnose()                              // サービス診断
}
```

### SecurityService.gs  
```javascript
// src/services/SecurityService.gs - 削除禁止機能
const SecurityService = {
  checkUserPermission(userId, permissionType), // 権限チェック
  persistSecurityLog(logEntry),            // セキュリティログ記録
  getSecurityLogs(options),               // セキュリティログ取得
  diagnose()                              // サービス診断
}
```

## 🎨 フロントエンド - 保護対象UI

### 重要なHTML/JSファイル
```
src/Page.html              - メイン表示画面
src/page.js.html           - メインJavaScript機能
src/AdminPanel.html        - 管理画面
src/AdminPanel.js.html     - 管理画面JavaScript
src/AppSetupPage.html      - セットアップ画面
src/login.js.html          - ログイン機能
src/SetupPage.html         - 初期設定画面
```

### CSS・デザイン要素
```
src/page.css.html          - メインページスタイル
src/AdminPanel.css.html    - 管理画面スタイル
src/SharedTailwindConfig.html - Tailwind設定
src/UnifiedStyles.html     - 統一スタイル
src/SharedSecurityHeaders.html - セキュリティヘッダー
```

## ⚠️ リファクタリング時の注意事項

1. **API戻り値の型を変更禁止**: フロントエンドが期待する形式を維持
2. **関数名変更禁止**: google.script.run.関数名() で呼ばれている
3. **主要パラメータ変更禁止**: 既存呼び出しとの互換性確保
4. **エラーハンドリング削除禁止**: try-catch, ログ出力を維持
5. **権限チェック削除禁止**: セキュリティ機能を必ず保持

## 🔧 改善時の推奨パターン

### ✅ 安全な改善方法
```javascript
// GOOD: 内部実装の最適化（外部インターフェース維持）
function getUser(kind = 'email') {
  try {
    // ✅ 内部ロジックは改善可能
    const userEmail = UserService.getCurrentEmail(); // 新Service使用
    
    if (kind === 'email') {
      return userEmail;  // ✅ 既存の戻り値型を維持
    }
    
    // ✅ 内部処理は最適化、戻り値型は維持
    const userInfo = UserService.getCurrentUserInfo();
    return {
      email: userEmail,
      userId: userInfo?.userId,
      isActive: userInfo?.isActive,
      hasConfig: !!userInfo?.config
    };
  } catch (error) {
    // ✅ エラーハンドリング維持
    console.error('getUser エラー:', error.message);
    return null;
  }
}
```

### ❌ 危険な変更例
```javascript
// BAD: 戻り値型の変更（フロントエンドがエラーになる）
function getUser(kind = 'email') {
  // ❌ 常にオブジェクトを返すように変更 → フロントエンドで型エラー
  return {
    email: userEmail,
    type: kind
  };
}

// BAD: 関数名の変更（呼び出し元でエラー）
function getCurrentUser(kind = 'email') { // ❌ 関数名変更
  // google.script.run.getUser() が動作しなくなる
}

// BAD: パラメータの変更（互換性失失）
function getUser(options = {}) { // ❌ パラメータ形式変更
  // 既存の getUser('email') 呼び出しが動作しなくなる
}
```

---

**🛡️ このファイルは改善作業中の重要機能保護のためのものです。作業完了後も参考資料として保持してください。**