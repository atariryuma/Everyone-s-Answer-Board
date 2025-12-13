# Everyone's Answer Board - Claude Code Development Guide

> **🎯 Project**: Google Apps Script Web Application
> **🔧 Stack**: Zero-Dependency Architecture, Direct GAS API Calls
> **🤖 Claude Code**: 2025 Best Practices Compliant
> **⚡ Updated**: 2025-12-13 (GAS + clasp + GitHub Best Practices: .js extension, proper .gitignore/.claspignore)

## 🧠 Claude Code Workflow

**Explore → Plan → Code → Deploy**

1. **Explore**: Read files, analyze (NO coding)
2. **Plan**: TodoWrite for tracking
3. **Code**: Incremental implementation
4. **Deploy**: `clasp push` to GAS
5. **Commit**: Structured git workflow (source code only)

## 🏗️ GAS-Optimized Architecture

**Core Pattern**: Direct GAS API calls with natural global scope utilization

```
🌟 GAS-Native Architecture (.js extension - clasp standard)
├── main.js                    # Entry Point (doGet/doPost only)
├── helpers.js                 # Utility Functions (shared operations)
├── validators.js              # Input Validation
├── formatters.js              # Data Formatting
├── DatabaseCore.js            # Database Operations (direct SpreadsheetApp)
├── SecurityService.js         # Security & Access Control
├── UserService.js             # User Management
├── ConfigService.js           # Configuration Management
├── DataService.js             # Data Operations
├── SystemController.js        # System Management
└── *.html                     # Frontend Templates
```

### **GAS-Native Implementation Pattern**

```javascript
// ✅ Direct GAS API usage - natural global scope
function getCurrentEmail() {
  return Session.getActiveUser().getEmail();
}

function isAdministrator(email) {
  const adminEmail = getCachedProperty('ADMIN_EMAIL');  // ✅ With 30s TTL cache
  return email?.toLowerCase() === adminEmail?.toLowerCase();
}

// ✅ PropertiesService caching with TTL (80-90% API call reduction)
const RUNTIME_PROPERTIES_CACHE = {};
const PROPERTY_CACHE_TTL = 30000; // 30 seconds

function getCachedProperty(key) {
  const now = Date.now();
  const cached = RUNTIME_PROPERTIES_CACHE[key];

  // ✅ TTL check: re-fetch if expired
  if (cached && cached.timestamp && (now - cached.timestamp < PROPERTY_CACHE_TTL)) {
    return cached.value;
  }

  // Fetch from PropertiesService and cache
  const value = PropertiesService.getScriptProperties().getProperty(key);
  RUNTIME_PROPERTIES_CACHE[key] = { value, timestamp: now };
  return value;
}

// ✅ Explicit cache clearing (for system setup/config updates)
function clearPropertyCache(key = null) {
  if (key) {
    delete RUNTIME_PROPERTIES_CACHE[key];
  } else {
    Object.keys(RUNTIME_PROPERTIES_CACHE).forEach(k => delete RUNTIME_PROPERTIES_CACHE[k]);
  }
}

function getUserData(email) {
  // Direct SpreadsheetApp usage for owner data
  const dbId = getCachedProperty('DATABASE_SPREADSHEET_ID');  // ✅ Cached access
  const spreadsheet = SpreadsheetApp.openById(dbId);
  const sheet = spreadsheet.getSheetByName('users');
  // ... direct data operations
}

// ✅ Service Account usage - ONLY for DATABASE_SPREADSHEET
// ユーザーの回答ボードは同一ドメイン共有設定（DOMAIN_WITH_LINK + EDIT）で対応
function getViewerBoardData(targetUserId, viewerEmail) {
  // DATABASE_SPREADSHEETからユーザー情報を取得（サービスアカウント使用）
  const targetUser = findUserById(targetUserId);

  // ユーザーの回答ボードは同一ドメイン共有設定により、全員が通常権限でアクセス可能
  // サービスアカウント不要（API quota問題を回避）
  return getUserSheetData(targetUser.userId, {
    includeTimestamp: true,
    requestingUser: viewerEmail
  });
}
```

## 🛠️ Development Commands

### **Quick Start (Most Used)**

```bash
npm run pull          # Pull code from GAS
npm run push          # Push code to GAS
npm run open          # Open GAS editor
npm run logs          # View execution logs
```

### **Claude Code Workflow**

1. **Explore** → Read files, analyze (NO coding yet)
2. **Plan** → TodoWrite for task tracking
3. **Code** → Incremental implementation
4. **Push** → `clasp push` to deploy
5. **Commit** → Git commit source code only

## 📝 Google Apps Script Critical Rules

### 🚨 **MUST: Performance & V8 Runtime**

```javascript
// ✅ MUST: Use batch operations (70x faster)
const data = sheet.getDataRange().getValues();  // ✅ One API call
const processed = data.map(row => transform(row));
sheet.getDataRange().setValues(processed);

// ❌ NEVER: Individual cell operations in loops
for (let i = 1; i <= sheet.getLastRow(); i++) {  // ❌ Hundreds of API calls
  const value = sheet.getRange(i, 1).getValue();
}

// ✅ MUST: Validate variables before template literals
if (error?.message) {
  const msg = `Error: ${error.message}`;  // ✅ Safe
}

// ❌ NEVER: Use setTimeout/setInterval (not available in GAS)
// ✅ USE: Utilities.sleep(1000) for delays
```

## 📁 File Structure

```
src/
├── main.js                    # API Gateway (frontend-callable functions only)
├── helpers.js                 # Utility functions (cache, response helpers)
├── validators.js              # Input validation functions
├── formatters.js              # Data formatting functions
├── DatabaseCore.js            # Database operations
├── SecurityService.js         # Security & validation
├── UserService.js             # User management
├── ConfigService.js           # Configuration management
├── ColumnMappingService.js    # Column mapping logic
├── ReactionService.js         # Reaction system
├── DataService.js             # Data operations
├── SystemController.js        # System management
├── SharingHelper.js           # Sharing utilities
└── *.html                     # UI templates
```

### **Architecture Rationale: main.js as API Gateway**

**Why main.js must contain all frontend-callable functions:**

- GAS requirement: Frontend uses `google.script.run[funcName]()` which requires global scope functions
- Only functions in main.js (or globally loaded files) can be called from frontend
- Helper functions NOT called by frontend (e.g., `getCurrentEmail`, `isAdministrator`) should be in separate files

**Design principle:**

```javascript
// ✅ main.js: Frontend-callable API functions only
function getUser(infoType) { /* ... */ }           // ✅ Called by frontend
function addReaction(userId, rowId, type) { }      // ✅ Called by frontend

// ✅ helpers.js: Shared utilities (not called by frontend)
function getCurrentEmail() { /* ... */ }           // ✅ Helper function
function isAdministrator(email) { /* ... */ }      // ✅ Helper function
function getCachedProperty(key) { /* ... */ }      // ✅ Utility function
function createErrorResponse(msg) { /* ... */ }    // ✅ Response helper
```

## 🎯 Main API Functions (Frontend-Callable)

**Important**: Functions in this section MUST be in `main.js` to be callable via `google.script.run[funcName]()`

```javascript
// User Management (main.js - frontend-callable)
getUser(infoType)                    // Get user information
processLoginAction(action)            // Handle login
getBatchedUserConfig()                // Get batched user config

// Data Operations (main.js - frontend-callable)
addReaction(userId, rowId, type)      // Add reaction
toggleHighlight(userId, rowId)        // Toggle highlight
getBulkAdminPanelData()              // Admin data

// Configuration (main.js - frontend-callable)
getConfig()                          // Get configuration
getUserConfig(userId)                // Get user config

// Internal Helpers (helpers.js - NOT frontend-callable)
getCurrentEmail()                    // Get current user email (internal use)
isAdministrator(email)               // Check admin privileges (internal use)
getCachedProperty(key)               // Cached property access with 30s TTL
clearPropertyCache(key)              // Explicit cache clearing
createErrorResponse(msg, data)       // Standard error response
createSuccessResponse(msg, data)     // Standard success response
```

### **Why this separation?**

**GAS Constraint**: Frontend can only call global scope functions. Therefore:

- ✅ **main.js**: Contains ALL functions called by frontend (API Gateway pattern)
- ✅ **helpers.js**: Contains internal helpers NOT called by frontend
- ❌ **Anti-pattern**: Moving frontend-callable functions out of main.js breaks frontend calls

## 🛡️ Security & Critical Rules

### **🚨 MUST (Enforced by Design)**

```javascript
// ✅ MUST: Service Account ONLY for cross-user access
function getViewerBoardData(targetUserId, viewerEmail) {
  const target = findUserById(targetUserId);
  if (target.userEmail === viewerEmail) {
    return getUserData(target);  // ✅ Own data: normal permissions
  } else {
    // ✅ Cross-user ONLY: use service account
    const access = openSpreadsheet(target.spreadsheetId, { useServiceAccount: true });
    return getUserData(target, { dataAccess: access });
  }
}

// ❌ NEVER: Service account for own data
function getUserOwnData(email) {
  const auth = Auth.serviceAccount();  // ❌ Privilege escalation
  return Data.findUserByEmail(email, auth);
}

// ✅ MUST: Validate all user inputs
// ✅ MUST: Sanitize HTML before rendering
// ✅ MUST: Use role-based access control
```

### **⚠️ SHOULD (Best Practices)**

- **Input Validation**: Validate email format, spreadsheet IDs, URLs
- **Error Handling**: Use try-catch with exponential backoff
- **Cache Strategy**: Use getCachedProperty for PropertiesService (80-90% API reduction)
- **Batch Operations**: Always prefer batch over individual operations

## 📝 Naming Conventions

### **Core Principle**: Natural English > Forced Prefixes

```javascript
// ✅ RECOMMENDED: Natural, readable names
getCurrentEmail()        // Clear and intuitive
getUserConfig(userId)    // Simple and direct
isAdmin()               // Boolean check
createErrorResponse()   // Verb + noun pattern

// ✅ Constants: UPPER_SNAKE_CASE with categories
const CACHE_DURATION = { SHORT: 10, MEDIUM: 30, LONG: 300 };
const TIMEOUT_MS = { QUICK: 100, DEFAULT: 5000 };

// ✅ Variables: camelCase with semantic prefixes
const isPublished = Boolean(config.isPublished);  // Boolean: is/has/can
const hasValidForm = validateUrl(url).isValid;

// ❌ AVOID: Unnecessary prefixes, magic numbers
authGetCurrentEmail()   // ❌ → getCurrentEmail()
cache.put(key, data, 300);  // ❌ → CACHE_DURATION.LONG
```

## 📋 Important Notes

### **Web App Entry Flow**

```
/exec → AccessRestricted (default safe landing)
     → ?mode=login → Setup → Admin Panel
     → ?mode=view&userId=... → Public Board View
```

### **API Guidelines**

- ✅ Frontend uses existing API names (no wrapper additions)
- ✅ Reactions: Direct DataService calls (`addReaction`, `toggleHighlight`)
- ✅ OAuth: Minimal scopes only (`spreadsheets`, `drive`, `userinfo.email`)

## 🏆 Architecture Benefits

- **70x Performance Improvement**: Batch operations (1s vs 70s)
- **Zero Dependencies**: Direct GAS API calls for maximum reliability
- **Loading Order Independence**: No file dependency chains
- **Production Stability**: Enterprise-grade error handling
- **Optimized Caching**: 80-90% PropertiesService API call reduction with 30s TTL
- **Simple Deployment**: Direct push to GAS with clasp (no build step)

## 🔧 clasp + GitHub Best Practices

### **File Extensions**

- ✅ **Use .js extension**: clasp's default format (not .gs)
- ✅ **Push with .js**: GAS editor displays them as .gs files
- ✅ **Pull gets .js**: clasp pull downloads files as .js

### **Git Workflow**

**Files to .gitignore:**
```
# Credentials (MUST ignore - contains scriptId)
.clasp.json
.clasprc.json

# Build artifacts
node_modules/
coverage/
dist/

# IDE files
.vscode/
.DS_Store
```

**Files to commit:**
```
# Source code
src/**/*.js
src/**/*.html
src/appsscript.json

# Config templates
.clasp.json.template    # Reference for team setup
.claspignore            # What to push to GAS
.gitignore              # What to ignore in git

# Dev environment
package.json
eslint.config.js
jest.config.js
```

### **Setup Instructions**

1. **Clone repository**:
   ```bash
   git clone <repo-url>
   cd Everyone-s-Answer-Board
   npm install
   ```

2. **Create .clasp.json** (copy from template):
   ```bash
   cp .clasp.json.template .clasp.json
   # Edit .clasp.json and add your scriptId
   ```

3. **Login to clasp**:
   ```bash
   npx clasp login
   ```

4. **Pull/Push code**:
   ```bash
   npm run pull    # Download from GAS
   npm run push    # Upload to GAS
   npm run open    # Open GAS editor
   npm run logs    # View execution logs
   ```

### **.claspignore Pattern**

```gitignore
# Ignore everything, then explicitly include
**/**
!appsscript.json
!**/*.js
!**/*.html

# Exclude from push
node_modules/**
.git/**
```

This ensures only production code is pushed to GAS, keeping the project clean.

---

*🤖 Claude Code 2025 Best Practices Compliant*
*📈 Optimized for Google Apps Script Performance and Reliability*
*⚡ Zero-Dependency Architecture Pattern*
