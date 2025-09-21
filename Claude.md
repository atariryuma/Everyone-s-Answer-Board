# Everyone's Answer Board - Claude Code Development Guide

> **🎯 Project**: Google Apps Script Web Application
> **🔧 Stack**: Zero-Dependency Architecture, Direct GAS API Calls
> **🤖 Claude Code**: 2025 Best Practices Compliant
> **⚡ Updated**: 2025-09-20 (Unified Naming Conventions Implemented)

## 🧠 Claude Code Core Principles

### 📋 Primary Workflow: **Explore → Plan → Code → Commit**

1. **Explore**: Read relevant files, analyze codebase - **NO CODING YET**
2. **Plan**: Create detailed implementation plan with TodoWrite
3. **Code**: Implement incrementally with TDD-first approach
4. **Commit**: Structured git workflow with proper documentation

## 🏗️ GAS-Optimized Architecture

**Core Pattern**: Direct GAS API calls with natural global scope utilization

```
🌟 GAS-Native Architecture
├── main.gs                    # Entry Point (doGet/doPost only)
├── auth.gs                    # Authentication (unified logic)
├── database.gs                # Database Operations (direct SpreadsheetApp)
├── permissions.gs             # Permission Management (simple role-based)
├── reactions.gs               # Reaction System (specialized feature)
├── utils.gs                   # Utility Functions (shared operations)
└── *.html                     # Frontend Templates
```

### **GAS-Native Implementation Pattern**
```javascript
// ✅ Direct GAS API usage - natural global scope
function getCurrentEmail() {
  return Session.getActiveUser().getEmail();
}

function isAdministrator(email) {
  const adminEmail = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL');
  return email?.toLowerCase() === adminEmail?.toLowerCase();
}

function getUserData(email) {
  // Direct SpreadsheetApp usage for owner data
  const dbId = PropertiesService.getScriptProperties().getProperty('DATABASE_SPREADSHEET_ID');
  const spreadsheet = SpreadsheetApp.openById(dbId);
  const sheet = spreadsheet.getSheetByName('users');
  // ... direct data operations
}

// ✅ Service Account usage - ONLY for cross-user access
function getViewerBoardData(targetUserId, viewerEmail) {
  const targetUser = findUserById(targetUserId);
  if (targetUser.userEmail === viewerEmail) {
    // Own data: use normal permissions
    return getUserSpreadsheetData(targetUser);
  } else {
    // Other's data: use service account for viewer access
    return getDataWithServiceAccount(targetUser);
  }
}
```

## 🛠️ Development Commands

### **Essential Commands**
```bash
# Quality Gate (MUST pass before commit)
npm run check                   # Combined linting, testing, validation

# GAS Development Cycle
clasp push                      # Deploy to Google Apps Script
clasp logs                      # View execution logs
npm run deploy:safe             # Production deployment with validation
```

### **Claude Code Workflow**
```markdown
1. Ask Claude to read CLAUDE.md - Context loading
2. **Explore**: "Read [files] and analyze [problem] - DO NOT CODE YET"
3. **Plan**: "Create detailed implementation plan with TodoWrite"
4. **Code**: TDD-first implementation with task tracking
5. **Test**: `npm run check` - Quality gate (MUST pass)
6. **Deploy**: `npm run deploy:safe`
```

## 📝 Google Apps Script V8 Best Practices

### 🚨 **Critical Migration (2025)**
- **Rhino Deprecated**: Feb 20, 2025
- **Rhino Shutdown**: Jan 31, 2026 (scripts will stop working)
- **Action Required**: Migrate all scripts to V8 runtime

### 🎯 **Performance Optimization (70x Improvement)**

#### **Batch Operations**
```javascript
// ✅ BEST: Batch operations (1 second vs 70 seconds)
function efficientDataProcessing() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const processedData = data.map(row => row.map(cell => processTransformation(cell)));
  sheet.getDataRange().setValues(processedData);
}

// ❌ AVOID: Individual operations
function inefficientDataProcessing() {
  const sheet = SpreadsheetApp.getActiveSheet();
  for (let i = 1; i <= sheet.getLastRow(); i++) {
    const value = sheet.getRange(i, 1).getValue();
    sheet.getRange(i, 1).setValue(processTransformation(value));
  }
}
```

#### **Modern JavaScript (V8 Runtime)**
```javascript
// ✅ V8 Features: const/let, destructuring, template literals
const { users, configs } = data;
const message = `Processing ${data.length} records`;

// ✅ Error Handling with Exponential Backoff
function executeWithRetry(operation, maxRetries = 3) {
  let retryCount = 0;
  while (retryCount < maxRetries) {
    try {
      return operation();
    } catch (error) {
      retryCount++;
      console.warn(`Attempt ${retryCount} failed:`, error.message);
      if (retryCount >= maxRetries) throw error;
      Utilities.sleep(Math.pow(2, retryCount) * 1000);
    }
  }
}
```

### 🛡️ **V8 Runtime Critical Differences**

#### **Variable Declaration**
```javascript
// ✅ RECOMMENDED: Use const/let, avoid var
const CONFIG_VALUES = { timeout: 5000, retries: 3 };
let processedValue = null;

// ❌ AVOID: var causes scope/hoisting issues
var globalVar = 'avoid this';
```

#### **Undefined Parameter Handling**
```javascript
// ✅ V8-safe undefined checks
function safeFunction(param) {
  if (typeof param !== 'undefined' && param !== null) {
    return param.toString();
  }
  return 'default value';
}

// ✅ Safe object access
const safeValue = (myObject || {}).propertyName;
const safeArrayValue = (myArray || [])[index];
```

#### **Template Literals with Validation**
問題はテンプレートリテラルではなく、**変数の値が実行時に存在しない可能性**です。`+`連結に書き換えても根本解決になりません。正しい修正は**変数の存在チェック**です。

```javascript
// ❌ 誤った修正（対症療法）
const errorInfo = 'Error: ' + error.message; // errorがnullなら同じエラー

// ✅ 正しい修正（根本治療）
if (error && error.message) {
  const errorInfo = `Error: ${error.message}\nStack: ${error.stack || 'N/A'}`;
} else {
  const errorInfo = 'An unknown error occurred.';
}

// ✅ 推奨パターン：事前チェック後のテンプレートリテラル使用
if (spreadsheetId && spreadsheetId.trim()) {
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`;
}
```

#### **Async/Sync Limitations**
```javascript
// ❌ NOT AVAILABLE: setTimeout/setInterval
// setTimeout(() => {}, 1000); // ReferenceError

// ✅ ALTERNATIVE: Utilities.sleep (synchronous, blocks execution)
Utilities.sleep(1000);

// ✅ RECOMMENDED: Batch processing for efficiency
function efficientProcessing(data) {
  const batchSize = 100;
  const results = [];
  for (let i = 0; i < data.length; i += batchSize) {
    const batch = data.slice(i, i + batchSize);
    results.push(...processBatch(batch));
  }
  return results;
}
```

## 📁 File Structure

```
src/
├── main.gs                    # API Gateway (Zero-Dependency)
├── UserService.gs             # User management
├── ConfigService.gs           # Configuration management
├── DataService.gs             # Spreadsheet operations
├── SecurityService.gs         # Security & validation
├── DatabaseCore.gs            # Database operations
├── SystemController.gs        # System management
├── DataController.gs          # Data operations
└── *.html                     # UI templates
```

## 🎯 Main API Functions

```javascript
// User Management
getCurrentEmail()              // Get current user email
getUser(infoType)             // Get user information
processLoginAction(action)     // Handle login

// Data Operations
addReaction(userId, rowId, type)     // Add reaction
toggleHighlight(userId, rowId)       // Toggle highlight
getBulkAdminPanelData()             // Admin data

// Configuration
getConfig()                   // Get configuration
getUserConfig(userId)         // Get user config
```

## 🧪 Testing & Quality

### **Current Status**
- **113/113 tests passing** (100% success rate)
- **Zero ESLint errors**
- **Complete coverage** of critical paths

### **Quality Gate**
```bash
npm run check                 # MUST pass before any commit
npm run deploy:safe           # Safe deployment with validation
```

## 🛡️ Security & Anti-Patterns

### **Security Best Practices**
- **Input Validation**: All user inputs validated
- **XSS/CSRF Protection**: HTML sanitization, token-based validation
- **Access Control**: Role-based permissions
- **Service Account**: **CROSS-USER ACCESS ONLY** - Security-critical restrictions
  - ✅ **Viewer→Editor**: Accessing published board data and reactions
  - ✅ **Editor→Admin**: Accessing shared user database (DATABASE_SPREADSHEET_ID)
  - ✅ **Cross-tenant operations**: Any access across user boundaries
  - ❌ **Self-access**: User accessing own spreadsheets (use direct GAS APIs)
  - ❌ **Same-tenant**: Operations within user's own permissions scope

### **Anti-Patterns to Avoid**
```javascript
// ❌ AVOID: Individual API calls in loops (use batch operations)
// ❌ AVOID: Unnecessary service account usage
function getUserOwnData(email) {
  const auth = Auth.serviceAccount();        // ❌ Unnecessary privilege escalation
  return Data.findUserByEmail(email, auth); // ❌ User accessing own data
}

// ✅ CORRECT: Appropriate permission usage
function getUserOwnData(email) {
  const dbId = PropertiesService.getScriptProperties().getProperty('DATABASE_SPREADSHEET_ID');
  const spreadsheet = SpreadsheetApp.openById(dbId); // ✅ Normal user permissions
  return findUserInSheet(spreadsheet, email);
}

function getViewerCrossUserData(targetUserId, viewerEmail) {
  const targetUser = findUserById(targetUserId);
  if (targetUser.userEmail === viewerEmail) {
    return getUserOwnData(viewerEmail);        // ✅ Own data: normal permissions
  } else {
    return getDataWithServiceAccount(targetUser); // ✅ Cross-user: service account
  }
}

// ❌ AVOID: Synchronous UI blocking operations
// ❌ AVOID: typeof undefined checks (unnecessary in GAS)
```

## 📝 Naming Conventions & Code Standards

### **🎯 Pragmatic Naming System**

**Core Principle**: 自然で読みやすい関数名を優先し、GAS-Native Architectureの可読性と実用性を最大化

#### **Function Naming - Natural English Pattern**
```javascript
// ✅ 推奨: 自然な英語表現パターン
getCurrentEmail()        // 直感的で分かりやすい
getUser()               // シンプルで明確
validateEmail()         // 目的が明確
createErrorResponse()   // 自然な動詞+名詞構造
isAdmin()              // 簡潔な状態確認
checkUserAccess()      // 分かりやすいアクション

// ✅ 特殊機能のみプレフィックス使用
getUserSheetData()   // DataService特有の複雑なデータ処理
addReaction()        // DataService特有の機能
sysLog()              // システムレベルの統一ログ機能

// ❌ 避けるべき: 不要なプレフィックス強制
authGetCurrentEmail()  // → getCurrentEmail() の方が自然
configGetUserConfig()  // → getUserConfig() の方が読みやすい
authIsAdministrator()  // → isAdministrator() の方がシンプル
```

#### **Variable & Property Naming**
```javascript
// ✅ camelCase + 意味的プレフィックス
const isPublished = Boolean(config.isPublished);    // ✅ boolean値: is/has/can
const isEditor = isAdministrator || isOwnBoard;     // ✅ 統一された権限表現
const hasValidForm = validateUrl(formUrl).isValid; // ✅ 存在確認: has

// ✅ オブジェクトプロパティの統一
{
  isPublished: true,        // ✅ appPublished → isPublished
  isEditor: false,          // ✅ showAdminFeatures → isEditor
  spreadsheetId: 'abc123',  // ✅ camelCase統一
  sheetName: 'Sheet1'       // ✅ 標準化されたパラメータ名
}

// ❌ 非推奨パターン
const appPublished = true;           // ❌ 曖昧な名前
const isAdminUser = false;          // ❌ 重複する概念
const showAdminFeatures = true;     // ❌ 表示ロジックと権限の混在
```

#### **Constants - System Standards**
```javascript
// ✅ UPPER_SNAKE_CASE + カテゴリ別構造
const CACHE_DURATION = {
  SHORT: 10,           // 認証ロック
  MEDIUM: 30,          // リアクション・ハイライト
  LONG: 300,           // ユーザー情報キャッシュ
  EXTRA_LONG: 3600     // 設定キャッシュ
};

const TIMEOUT_MS = {
  QUICK: 100,          // UI応答性
  DEFAULT: 5000,       // デフォルトタイムアウト
  EXTENDED: 30000      // 拡張タイムアウト
};

// ✅ 使用例
cache.put(cacheKey, data, CACHE_DURATION.LONG);
Utilities.sleep(SLEEP_MS.SHORT);

// ❌ マジックナンバー（非推奨）
cache.put(cacheKey, data, 300);     // ❌ 意味が不明
Utilities.sleep(100);               // ❌ なぜ100ms？
```

#### **Function Categories & Naming Patterns**
```javascript
// ✅ カテゴリ別命名パターン

// 1. データ取得・操作
getCurrentEmail()       // get + 対象名
getUser()              // シンプルで明確
getUserConfig()        // 対象を明確に
saveUserConfig()       // 動作を明確に

// 2. 検証・確認
validateEmail()        // validate + 対象
isAdmin()             // 状態確認（boolean返却）
checkUserAccess()     // check + 確認内容

// 3. 作成・生成
createErrorResponse()  // create + 作成対象
generateDynamicUrls()  // generate + 生成内容

// 4. 処理・変換
processLoginAction()   // process + 処理対象
handleGetData()       // handle + ハンドル対象
formatTimestamp()     // format + 変換対象

// 5. 特殊プレフィックス（必要時のみ）
getUserSheetData()  // 複雑なデータ処理
sysLog()             // システムレベル機能
```

#### **Parameter Naming Standards**
```javascript
// ✅ 統一されたパラメータ名
function processUserData(userId, spreadsheetId, sheetName, options = {}) {
  // userId: 常にcamelCase、一意識別子
  // spreadsheetId: Google Sheets ID（統一形式）
  // sheetName: シート名（camelCase）
  // options: オプション引数は常にオブジェクト
}

// ✅ 一貫性のあるAPI設計
getCurrentEmail()                           // Auth layer
getUserConfig(userId)                       // Config layer
getUserSheetData(userId, options)        // Data layer (特殊処理)

// ❌ 非一貫的なパラメータ（非推奨）
function badFunction(user_id, spreadsheet-id, sheet_name) { } // ❌ 命名規則混在
function anotherBad(userID, spreadSheetId, SheetName) { }     // ❌ 大文字小文字不統一
```

### **🔧 Implementation Guidelines**

#### **GAS-Native Pattern Compliance**
```javascript
// ✅ GAS-Native: Direct API calls with natural global scope
function getCurrentEmail() {
  return Session.getActiveUser().getEmail();
}

function getUserConfig(userId) {
  const dbId = PropertiesService.getScriptProperties().getProperty('DATABASE_SPREADSHEET_ID');
  const spreadsheet = SpreadsheetApp.openById(dbId);
  // Direct SpreadsheetApp usage - no abstraction layers
}

// ✅ Specialized functions with clear prefixes
function getUserSheetData(userId, options = {}) {
  // Complex data operations warrant specific naming
  // Clear functional responsibility
}

// ✅ GAS-Native constants (no typeof checks needed)
const CACHE_DURATION = {
  SHORT: 10,
  MEDIUM: 30,
  LONG: 300
};

const TIMEOUT_MS = {
  QUICK: 100,
  DEFAULT: 5000,
  EXTENDED: 30000
};
```

#### **Naming Philosophy & Migration**
```javascript
// ✅ CLAUDE.md準拠：実用性重視のシンプル命名
// 自然な英語 > 強制的なプレフィックス
// 読みやすさ > 厳格な分類

// ✅ 保持すべき関数名パターン
getCurrentEmail()         // 直感的
getUserConfig()          // 明確
validateEmail()          // 目的が分かりやすい
createErrorResponse()    // 自然な動詞+名詞
isAdmin()               // シンプルな状態確認

// ✅ 必要なプレフィックス（機能的理由）
getUserSheetData()    // DataService特有の複雑処理
addReaction()         // DataService特有の機能
sysLog()               // システムレベル統一機能

// ❌ 不要なプレフィックス強制は避ける
// 読みやすさを損なうプレフィックスは使用しない
```

### **📊 Benefits of Pragmatic Naming**

- **🎯 Natural Readability**: 自然な英語表現で直感的に理解可能
- **⚡ Development Speed**: シンプルな関数名で開発効率向上
- **📖 Self-Documenting Code**: 機能が名前から即座に理解できる
- **🛠️ Maintenance Efficiency**: 読みやすい関数名で保守性向上
- **🔄 Zero-Dependency Compliance**: ファイル読み込み順序に依存しない設計
- **🎨 Code Aesthetics**: 美しく読みやすいコードベース実現

## 📋 Important Application Notes

### **Web App Entry Flow**
```
/exec access → AccessRestricted (safe landing)
            → (explicit) ?mode=login → Setup → Admin Panel
            → (viewer)  ?mode=view&userId=... → Public Board View
```

**Policy**: `/exec`のデフォルトはログインページではなく、安全な着地点（AccessRestricted）を表示。これは閲覧ユーザーが不用意にアカウント生成することを防ぐ安全策。

### **API Compatibility Guidance**
- フロントエンドは既存API名に合わせる（例: `isAdmin`、`getSheets`）
- リアクション系はDataServiceの公開関数を直接利用（`addReaction`, `toggleHighlight`）
- 新たな中間API（Gatewayラッパ）は原則追加しない

### **OAuth Scopes Policy**
- 必要最小限のスコープのみ採用
- 既定: `spreadsheets`, `drive`, `script.external_request`, `userinfo.email`
- 未使用のAdvanced Servicesは無効化

## 🏆 Architecture Benefits

- **70x Performance Improvement**: Batch operations (1s vs 70s)
- **Zero Dependencies**: Direct GAS API calls for maximum reliability
- **Loading Order Independence**: No file dependency chains
- **100% Test Coverage**: 113/113 tests passing
- **Production Stability**: Enterprise-grade error handling

---

*🤖 Claude Code 2025 Best Practices Compliant*
*📈 Optimized for Google Apps Script Performance and Reliability*
*⚡ Zero-Dependency Architecture Pattern*