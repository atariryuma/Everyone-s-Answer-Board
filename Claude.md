# Everyone's Answer Board - Development Guide

**Project**: Google Apps Script Web Application for organization-internal use
**Stack**: Zero-dependency, direct GAS API calls, V8 runtime
**Updated**: 2025-12-14

---

## 🌟 Google Apps Script - Critical Context

### Single Global Scope Execution (MOST IMPORTANT)

**All script files execute in a single global scope** - no module system, no import/export.

- File loading order is irrelevant - all functions accessible from any file
- `google.script.run.funcName()` calls global scope functions (regardless of which file)
- No dependency management needed (unlike Node.js)
- Watch for namespace collisions - avoid duplicate function names

**References**: [Apps Script Best Practices](https://developers.google.com/apps-script/guides/support/best-practices), [HTML Service Best Practices](https://developers.google.com/apps-script/guides/html/best-practices)

---

## 🚨 MUST Rules (Enforced - CI/Review)

### Performance

- **Batch operations always** - `getDataRange().getValues()` not `getRange(i,j).getValue()` in loops
  - Impact: 70x performance (1s vs 70s)
  - See: `src/DatabaseCore.js:125-128` for reference implementation
- **Cache PropertiesService** - 30s TTL, 80-90% API call reduction
  - See: `src/helpers.js:30-57` for getCachedProperty implementation
- **Minimize external service calls** - JavaScript operations within script are faster

### Security

- **Use Session.getActiveUser() only** - never `getEffectiveUser()` (privilege escalation risk)
  - See: `src/main.js:25`
- **Validate all inputs** - email, IDs, URLs before processing
  - See: `src/validators.js` for validation functions
- **Sanitize HTML** - use `textContent` not `innerHTML` for user content
  - Use `escapeHtml()` helper if innerHTML required
  - See: `src/SharedUtilities.html` for escapeHtml implementation

### V8 Runtime

- **Use Utilities.sleep()** - never `setTimeout/setInterval` (not available)
- **Modern JavaScript OK** - arrow functions, destructuring, optional chaining, nullish coalescing
- **No Node.js APIs** - use GAS services (DriveApp, SpreadsheetApp, etc.)

---

## ⚠️ SHOULD Rules (Strongly Recommended)

### File Organization

- **Target 500-1500 lines per file** - for readability/maintainability
- **Current status**: Files are well-organized with clear separation of concerns
- **Largest file**: `SystemController.js` (1,980 lines) - contains system management and performance monitoring

### Architecture Patterns

- **main.js as API Gateway** - Entry points (doGet/doPost) and core authentication
- **Service-based architecture**:
  - `*Service.js` - Business logic (UserService, ConfigService, DataService, etc.)
  - `*Apis.js` - API endpoints grouped by domain (AdminApis, UserApis, DataApis)
  - `helpers.js` - Shared utilities and response helpers
  - `validators.js` - Input validation functions
- **Error handling** - try-catch with exponential backoff for external calls
  - See: `src/main.js:782-837` for executeWithRetry pattern
- **Naming** - natural English > forced prefixes (`getCurrentEmail` not `authGetCurrentEmail`)

---

## 📁 File Structure

```text
src/
├── main.js (939 lines)           # Entry points, auth, batched data retrieval
├── SystemController.js (1980)    # System management, diagnostics, performance
├── DatabaseCore.js (1324)        # Database operations, circuit breaker
├── DataApis.js (1210)            # Data operation APIs
├── ConfigService.js (839)        # User configuration management
├── DataService.js (676)          # Data retrieval and processing
├── ReactionService.js (553)      # Reactions and highlights
├── ColumnMappingService.js (490) # Column mapping for form data
├── validators.js (415)           # Input validation
├── AdminApis.js (405)            # Admin panel APIs
├── UserService.js (402)          # User management
├── SecurityService.js (355)      # Security and auth helpers
├── UserApis.js (235)             # User-related APIs
├── helpers.js (189)              # Utility functions, response helpers
├── formatters.js (55)            # Data formatting utilities
├── SharingHelper.js (42)         # Sharing and permission helpers
└── *.html (19 files)             # Frontend templates and components
```

### HTML Templates

```text
src/
├── Page.html                 # Main board view
├── AdminPanel.html           # Admin dashboard
├── AdminPanel.js.html         # Admin panel JavaScript
├── LoginPage.html            # Login page
├── SetupPage.html            # Initial setup wizard
├── AppSetupPage.html         # App configuration
├── AccessRestricted.html     # Access denied page
├── ErrorBoundary.html        # Error display
├── Unpublished.html          # Unpublished board notice
├── TeacherManual.html        # User manual
├── SharedUtilities.html      # Shared JS utilities
├── SharedModals.html         # Modal components
├── SharedErrorHandling.html  # Error handling
├── SharedSecurityHeaders.html # Security headers
├── SharedTailwindConfig.html # Tailwind CSS config
├── UnifiedStyles.css.html    # Unified styles
├── page.css.html             # Page-specific styles
├── page.js.html              # Page JavaScript
└── login.js.html             # Login JavaScript
```

---

## 🛠️ Common Commands

```bash
npm run pull    # Pull from GAS
npm run push    # Push to GAS
npm run open    # Open GAS editor
npm run logs    # View execution logs
npm run deploy  # Deploy new version
```

**Workflow**: Explore → Plan → Code → `npm run push` → Test in browser → Git commit

---

## 🔧 clasp + Git

### Git Workflow

- **Ignore**: `.clasp.json`, `.clasprc.json` (contains scriptId - sensitive)
- **Commit**: `src/**/*.js`, `src/**/*.html`, `src/appsscript.json`, `.clasp.json.template`
- **File extension**: Use `.js` (clasp standard), GAS editor shows as `.gs`

### Required Script Properties

System requires these properties (set via GAS editor or SetupPage):

| Property | Description |
|----------|-------------|
| `ADMIN_EMAIL` | Administrator email address |
| `DATABASE_SPREADSHEET_ID` | Main database spreadsheet ID |
| `SERVICE_ACCOUNT_CREDS` | Service account JSON credentials |
| `GOOGLE_CLIENT_ID` | (Optional) OAuth client ID |

---

## 🔗 Web App Entry Points

```text
/exec                          → AccessRestricted (default landing)
/exec?mode=login               → LoginPage (user authentication)
/exec?mode=setup               → SetupPage (initial system setup)
/exec?mode=appSetup            → AppSetupPage (admin only)
/exec?mode=manual              → TeacherManual (user guide)
/exec?mode=view&userId=X       → Page.html (public board view)
/exec?mode=admin&userId=X      → AdminPanel (board management)
```

---

## 📚 References

**Official Google Documentation**:

- [Apps Script Best Practices](https://developers.google.com/apps-script/guides/support/best-practices)
- [HTML Service Best Practices](https://developers.google.com/apps-script/guides/html/best-practices)
- [clasp CLI](https://github.com/google/clasp)
- [Web Apps Guide](https://developers.google.com/apps-script/guides/web)

---

*This file is optimized for AI consumption following 2025 best practices.*
