# Everyone's Answer Board - Claude Code Development Guide

> **🎯 Project**: Google Apps Script Web Application
> **🔧 Stack**: Zero-Dependency Architecture, Direct GAS API Calls
> **🤖 Claude Code**: 2025 Best Practices Compliant
> **⚡ Updated**: 2025-09-19

## 🧠 Claude Code Core Principles

### 📋 Primary Workflow: **Explore → Plan → Code → Commit**

1. **Explore**: Read relevant files, analyze codebase - **NO CODING YET**
2. **Plan**: Create detailed implementation plan with TodoWrite
3. **Code**: Implement incrementally with TDD-first approach
4. **Commit**: Structured git workflow with proper documentation

## 🏗️ Zero-Dependency Architecture

**Core Pattern**: Direct GAS API calls eliminate file loading order issues

```
🌟 Architecture
├── main.gs                    # API Gateway (Zero-Dependency)
├── *Service.gs                # Services Layer (Direct GAS APIs)
├── *Controller.gs             # Business Logic Controllers
└── *.html                     # Frontend (Unified API calls)
```

### **Implementation Pattern**
```javascript
// ✅ Direct GAS API usage with fallbacks
function getConfig() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email) return { success: false, message: 'Not authenticated' };

    const user = typeof DatabaseOperations !== 'undefined' ?
      DatabaseOperations.findUserByEmail(email) : null;

    return user ?
      { success: true, config: JSON.parse(user.configJson || '{}') } :
      { success: false, message: 'User not found' };
  } catch (error) {
    console.error('getConfig error:', error.message);
    return { success: false, message: error.message };
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
- **Service Account**: Use `Data.open()` instead of `SpreadsheetApp.openById()`

### **Anti-Patterns to Avoid**
```javascript
// ❌ AVOID: Individual API calls in loops
// ❌ AVOID: Library dependencies in production
// ❌ AVOID: User permission-dependent data access
SpreadsheetApp.openById(id);              // ❌ User permission dependent
Data.open(id);                           // ✅ Service account

// ❌ AVOID: Direct service dependencies at file level
// ❌ AVOID: Synchronous UI blocking operations
```

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
- リアクション系はDataServiceの公開関数を直接利用（`dsAddReaction`, `dsToggleHighlight`）
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