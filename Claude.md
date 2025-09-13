# Everyone's Answer Board - Claude Code Development Guide

> **🎯 Project**: Google Apps Script Web Application
> **🔧 Stack**: GAS, Services Architecture, Spreadsheet Integration
> **⚡ Updated**: 2025-09-13

## Essential Commands

```bash
# Development
npm run test                    # Run tests
npm run lint                    # Code linting
npm run check                   # Full quality check
./scripts/safe-deploy.sh        # Safe deployment to GAS

# GAS Management
clasp push                      # Deploy to Google Apps Script
clasp open                      # Open GAS editor
clasp logs                      # View execution logs
```

## Code Style Guidelines

- **ES2020+ syntax** (GAS V8 runtime)
- **Services architecture**: All business logic in `src/services/`
- **Single responsibility**: One concern per file/function
- **Error handling**: Always use try-catch with proper logging
- **No global variables**: Use service pattern for state management

## Key Architecture

```javascript
// ✅ Current Recommended APIs
const user = UserService.getCurrentUserInfo();
const config = ConfigService.getUserConfig(userId);
const data = DataService.getSheetData(userId, options);

// ⚠️ Legacy (being phased out)
const dbData = DB.findUserByEmail(email);

// ❌ Deleted - Do Not Use
// ConfigurationManager, SimpleCacheManager
```

## Important Flow: Web App Entry Points

```
/exec access → Login Page → Setup (if needed) → Main Board
```

**Critical**: `/exec` always starts with login page, never direct main board access.

## Testing Requirements

- **Before deploy**: Run `npm run check` and ensure all tests pass
- **After changes**: Test actual web app functionality
- **GAS testing**: Check execution logs for runtime errors

## File Structure

```
src/
├── services/           # Business logic (UserService, ConfigService, etc.)
├── infrastructure/     # Data layer (DatabaseService, CacheService)
├── core/              # Constants, errors, service registry
├── utils/             # Utilities (formatters, validators, helpers)
└── *.html             # UI templates
```

## Common Issues & Solutions

- **Duplicate declarations**: Check for existing const/function before creating
- **Authentication flow**: Ensure proper user flow from login → setup → main
- **GAS limitations**: Use service pattern to avoid global scope conflicts

## Safety Rules

1. **Always test locally** before deploying
2. **Use safe-deploy script** for production changes
3. **Check GAS logs** after deployment
4. **Keep backup** of working versions

## Recent Major Refactoring (2025-09-13)

### Architecture Cleanup: main.gs → Controllers Pattern

**Problem Solved**: main.gs had grown to 2,881 lines, violating single responsibility principle

**Solution**: Implemented proper separation of concerns with specialized controllers

### Results:
- **84% Size Reduction**: main.gs reduced from 2,881 → 400 lines
- **4 New Controllers**: Organized business logic into focused modules
- **All Tests Pass**: 113/113 tests continue working after refactoring
- **Successful Deployment**: 39 files deployed with safe-deploy script

### New Controller Structure:

```
src/controllers/
├── AdminController.gs      # Admin panel APIs (getConfig, getSpreadsheetList, etc.)
├── DataController.gs       # Data operations (handleGetData, reactions, highlights)
├── FrontendController.gs   # Frontend APIs (getUser, login, authentication)
└── SystemController.gs     # System management (setup, diagnostics, monitoring)
```

### main.gs Now Contains Only:
- HTTP request routing (doGet/doPost)
- Template inclusion utility (`include` function)
- Mode-based handler delegation
- Error handling and response formatting

### Enhanced Services:
- **DataService**: Added admin panel functions (getSpreadsheetList, analyzeColumns)
- **ConfigService**: Added configuration management (saveDraftConfiguration, publishApplication)

### Key Benefits:
1. **Maintainability**: Each controller has single responsibility
2. **Testability**: Isolated business logic easier to test
3. **Readability**: main.gs is now pure entry point
4. **Scalability**: Adding new features won't bloat main.gs
5. **Debugging**: Errors easier to locate in focused modules

### Backward Compatibility:
All existing HTML files continue to work through global function exports:
```javascript
// Global functions still available for google.script.run calls
function getUser(kind) { return FrontendController.getUser(kind); }
function setupApplication(...args) { return SystemController.setupApplication(...args); }
```

## Function Mapping Guide - Avoid Duplicates

### 📁 File Responsibility Matrix

| **Layer** | **File** | **Primary Purpose** | **Should Contain** | **Should NOT Contain** |
|-----------|----------|-------------------|-------------------|----------------------|
| **Entry** | `main.gs` | HTTP routing & templates | `doGet()`, `doPost()`, `include()`, route handlers | Business logic, database ops |
| **Controllers** | `FrontendController.gs` | Frontend HTML APIs | `getUser()`, `processLoginAction()`, auth functions | Database operations, complex logic |
| | `AdminController.gs` | Admin panel APIs | `getConfig()`, `validateAccess()`, admin functions | User creation, data processing |
| | `DataController.gs` | Data operations | `handleGetData()`, reactions, user management | System setup, authentication |
| | `SystemController.gs` | System management | `setupApplication()`, diagnostics, maintenance | User data, frontend APIs |
| **Services** | `UserService.gs` | User management | `getCurrentUserInfo()`, `createUser()`, permissions | HTTP handling, templates |
| | `ConfigService.gs` | Configuration | `getUserConfig()`, `saveConfig()`, validation | User creation, data processing |
| | `DataService.gs` | Spreadsheet data | `getSheetData()`, `addReaction()`, sheet ops | User authentication, system setup |
| | `SecurityService.gs` | Security & auth | Token management, input sanitization | Data formatting, UI logic |
| **Infrastructure** | `DatabaseService.gs` | Database operations | `DB.*` CRUD functions, Service Account API | Business logic, UI concerns |
| | `CacheService.gs` | Caching strategy | Cache management, TTL, invalidation | Data processing, user logic |
| **Core** | `constants.gs` | App constants | `CORE`, `CONSTANTS`, `PROPS_KEYS`, enums | Functions, business logic |
| | `errors.gs` | Error handling | `ErrorHandler`, error categorization | Data operations, user management |
| **Utils** | `helpers.gs` | Common utilities | `ColumnHelpers`, `FormatHelpers`, calculations | Domain-specific logic |
| | `formatters.gs` | Data formatting | `ResponseFormatter`, `DataFormatter` | Business operations |
| | `validators.gs` | Input validation | `InputValidator`, security validation | Data persistence |

### 🔍 Function Placement Decision Tree

```
New Function Needed?
         │
         ▼
    What does it do?
         │
    ┌────┴────────────┬──────────────┐
    ▼                 ▼              ▼
HTTP/Routing?    Business Logic?   Utility?
    │                 │              │
    ▼                 ▼              ▼
  main.gs        Which domain?    utils/*.gs
                      │
               ┌──────┼──────┬──────────┐
               ▼      ▼      ▼          ▼
             User  Config  Data     System
               │      │      │          │
               ▼      ▼      ▼          ▼
          UserService ConfigService DataService SystemController
```

### ⚠️ Common Duplication Patterns to Avoid

#### 1. **User Information Functions**
- ✅ **Use:** `UserService.getCurrentUserInfo()`
- ❌ **Don't create:** `getUser()`, `getCurrentUser()`, `fetchUserData()`
- **Location:** All user-related functions belong in `UserService.gs`

#### 2. **Configuration Functions**
- ✅ **Use:** `ConfigService.getUserConfig(userId)`
- ❌ **Don't create:** `getConfig()`, `loadConfiguration()`, `fetchSettings()`
- **Location:** All config operations belong in `ConfigService.gs`

#### 3. **Data Retrieval Functions**
- ✅ **Use:** `DataService.getSheetData(userId, options)`
- ❌ **Don't create:** `getData()`, `loadSpreadsheetData()`, `fetchSheetInfo()`
- **Location:** All spreadsheet operations belong in `DataService.gs`

#### 4. **Authentication Functions**
- ✅ **Use:** `UserService.isSystemAdmin(email)`, `SecurityService.*`
- ❌ **Don't create:** `checkAuth()`, `validateUser()`, `isAdmin()`
- **Location:** Auth in `UserService.gs`, security in `SecurityService.gs`

### 📋 Before Adding New Functions Checklist

1. **Search existing code:** Use `Grep` to find similar functionality
```bash
# Search for existing functions
rg "function.*[Gg]et.*[Uu]ser" src/
rg "function.*[Cc]onfig" src/
rg "function.*[Dd]ata" src/
```

2. **Check appropriate service:**
   - User-related? → `UserService.gs`
   - Config-related? → `ConfigService.gs`
   - Data-related? → `DataService.gs`
   - System-related? → `SystemController.gs`

3. **Verify layer placement:**
   - HTTP handling? → `main.gs` or `controllers/*.gs`
   - Business logic? → `services/*.gs`
   - Data access? → `infrastructure/*.gs`
   - Utilities? → `utils/*.gs`

4. **Check naming consistency:**
   - Use domain prefixes: `getUserConfig()`, `saveUserConfig()`
   - Avoid generic names: `getData()`, `save()`, `load()`

### 🎯 Golden Rules

1. **One function, one location** - Never duplicate functionality
2. **Search before create** - Always check if it already exists
3. **Respect boundaries** - Services don't call Controllers
4. **Use clear names** - Include context in function names
5. **Follow the layers** - Maintain architectural separation