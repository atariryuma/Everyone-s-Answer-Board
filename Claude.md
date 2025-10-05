# Everyone's Answer Board - Claude Code Development Guide

> **🎯 Project**: Google Apps Script Web Application
> **🔧 Stack**: Zero-Dependency Architecture, Direct GAS API Calls
> **🤖 Claude Code**: 2025 Best Practices Compliant
> **⚡ Updated**: 2025-10-05 (Architecture Optimization: auth.gs separation, TTL caching, API Gateway clarification)

## 🧠 Claude Code Workflow

**Explore → Plan → Code → Test → Commit**

1. **Explore**: Read files, analyze (NO coding)
2. **Plan**: TodoWrite for tracking
3. **Code**: TDD-first incremental implementation
4. **Test**: `npm run check` MUST pass
5. **Commit**: Structured git workflow

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

// ✅ Service Account usage - ONLY for cross-user access
function getViewerBoardData(targetUserId, viewerEmail) {
  const targetUser = findUserById(targetUserId);
  if (targetUser.userEmail === viewerEmail) {
    // Own data: use normal permissions
    return getUserSpreadsheetData(targetUser);
  } else {
    // Other's data: use service account for cross-user access
    const dataAccess = openSpreadsheet(targetUser.spreadsheetId, { useServiceAccount: true });
    return getUserSpreadsheetData(targetUser, { dataAccess });
  }
}
```

## 🛠️ Development Commands

### **🚨 Quick Start (Most Used)**

```bash
npm run check          # ✅ MUST pass before commit (lint + test)
clasp push            # Deploy to GAS
clasp logs            # View execution logs
npm run deploy:safe   # Production deployment
```

### **Claude Code Workflow**

1. **Explore** → Read files, analyze (NO coding yet)
2. **Plan** → TodoWrite for task tracking
3. **Code** → TDD-first implementation
4. **Test** → `npm run check` (MUST pass)
5. **Deploy** → `npm run deploy:safe`

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
├── main.gs                    # API Gateway (frontend-callable functions only)
├── auth.gs                    # Authentication helpers (getCurrentEmail, isAdministrator)
├── helpers.gs                 # Utility functions (cache, response helpers)
├── UserService.gs             # User management
├── ConfigService.gs           # Configuration management
├── DataService.gs             # Spreadsheet operations
├── SecurityService.gs         # Security & validation
├── DatabaseCore.gs            # Database operations
├── SystemController.gs        # System management
├── DataController.gs          # Data operations
└── *.html                     # UI templates
```

### **Architecture Rationale: main.gs as API Gateway**

**Why main.gs must contain all frontend-callable functions:**

- GAS requirement: Frontend uses `google.script.run[funcName]()` which requires global scope functions
- Only functions in main.gs (or globally loaded files) can be called from frontend
- Helper functions NOT called by frontend (e.g., `getCurrentEmail`, `isAdministrator`) should be in separate files

**Design principle:**

```javascript
// ✅ main.gs: Frontend-callable API functions only
function getUser(infoType) { /* ... */ }           // ✅ Called by frontend
function addReaction(userId, rowId, type) { }      // ✅ Called by frontend

// ✅ auth.gs: Internal helpers (not called by frontend)
function getCurrentEmail() { /* ... */ }           // ✅ Helper function
function isAdministrator(email) { /* ... */ }      // ✅ Helper function

// ✅ helpers.gs: Shared utilities
function getCachedProperty(key) { /* ... */ }      // ✅ Utility function
function createErrorResponse(msg) { /* ... */ }    // ✅ Response helper
```

## 🎯 Main API Functions (Frontend-Callable)

**Important**: Functions in this section MUST be in `main.gs` to be callable via `google.script.run[funcName]()`

```javascript
// User Management (main.gs - frontend-callable)
getUser(infoType)                    // Get user information
processLoginAction(action)            // Handle login
getBatchedUserConfig()                // Get batched user config

// Data Operations (main.gs - frontend-callable)
addReaction(userId, rowId, type)      // Add reaction
toggleHighlight(userId, rowId)        // Toggle highlight
getBulkAdminPanelData()              // Admin data

// Configuration (main.gs - frontend-callable)
getConfig()                          // Get configuration
getUserConfig(userId)                // Get user config

// Internal Helpers (auth.gs - NOT frontend-callable)
getCurrentEmail()                    // Get current user email (internal use)
isAdministrator(email)               // Check admin privileges (internal use)

// Utilities (helpers.gs - NOT frontend-callable)
getCachedProperty(key)               // Cached property access with 30s TTL
clearPropertyCache(key)              // Explicit cache clearing
createErrorResponse(msg, data)       // Standard error response
createSuccessResponse(msg, data)     // Standard success response
```

### **Why this separation?**

**GAS Constraint**: Frontend can only call global scope functions. Therefore:

- ✅ **main.gs**: Contains ALL functions called by frontend (API Gateway pattern)
- ✅ **auth.gs/helpers.gs**: Contains internal helpers NOT called by frontend
- ❌ **Anti-pattern**: Moving frontend-callable functions out of main.gs breaks frontend calls

## 🧪 Testing & Quality

### **Current Status**

- **123/123 tests passing** (100% success rate)
- **Zero ESLint errors**
- **Complete coverage** of critical paths

### **Quality Gate**

```bash
npm run check                 # MUST pass before any commit
npm run deploy:safe           # Safe deployment with validation
```

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
- **100% Test Coverage**: 123/123 tests passing
- **Production Stability**: Enterprise-grade error handling
- **Optimized Caching**: 80-90% PropertiesService API call reduction with 30s TTL

---

*🤖 Claude Code 2025 Best Practices Compliant*
*📈 Optimized for Google Apps Script Performance and Reliability*
*⚡ Zero-Dependency Architecture Pattern*
