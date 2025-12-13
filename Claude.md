# Everyone's Answer Board - Development Guide

**Project**: Google Apps Script Web Application for organization-internal use
**Stack**: Zero-dependency, direct GAS API calls, V8 runtime
**Updated**: 2025-12-13

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
  - See: `src/helpers.js:50-67` for getCachedProperty implementation
- **Minimize external service calls** - JavaScript operations within script are faster

### Security

- **Use Session.getActiveUser() only** - never `getEffectiveUser()` (privilege escalation risk)
  - See: `src/main.js:26`
- **Validate all inputs** - email, IDs, URLs before processing
  - See: `src/validators.js` for validation functions
- **Sanitize HTML** - use `textContent` not `innerHTML` for user content
  - Use `escapeHtml()` helper if innerHTML required
  - See: `src/SharedUtilities.html:731-747`
- **XSS prevention critical** - 100+ innerHTML usage in codebase needs review

### V8 Runtime

- **Use Utilities.sleep()** - never `setTimeout/setInterval` (not available)
- **Modern JavaScript OK** - arrow functions, destructuring, optional chaining, nullish coalescing
- **No Node.js APIs** - use GAS services (DriveApp, SpreadsheetApp, etc.)

---

## ⚠️ SHOULD Rules (Strongly Recommended)

### File Organization

- **Target 500-1500 lines per file** - for readability/maintainability
- **Current status**: `main.js` = 2,841 lines → split recommended for readability
- **Suggested split** (optional, not required):
  - `main.js` (500 lines) - doGet/doPost + core APIs
  - `AdminApis.js` (800 lines) - Admin panel APIs
  - `UserApis.js` (600 lines) - User management APIs
  - `DataApis.js` (700 lines) - Data operation APIs

**Note**: This is a readability guideline, not a GAS constraint. Functions work from any file due to global scope.

### Architecture Patterns

- **main.js as API Gateway** (recommended pattern for this project)
  - Frontend-callable functions centralized for clarity
  - Internal helpers in separate files (helpers.js, SecurityService.js, etc.)
  - Benefit: Clear API surface, easier onboarding
- **Error handling** - try-catch with exponential backoff for external calls
  - See: `src/DatabaseCore.js:25-86` for circuit breaker pattern
- **Naming** - natural English > forced prefixes (`getCurrentEmail` not `authGetCurrentEmail`)

### Environment Detection

**Current issue**: `localhost` detection doesn't work in GAS (always script.google.com domain)

- Fix: Use PropertiesService flag (`ENVIRONMENT=development|production`)
- Or: Use deployment URL pattern (`.includes('/dev')` vs `.includes('/exec')`)

---

## 📁 File Structure

```text
src/
├── main.js (2841 lines - split recommended)
├── helpers.js - Utility functions
├── validators.js - Input validation
├── DatabaseCore.js - DB operations
├── SecurityService.js - Security/auth
├── UserService.js - User management
├── ConfigService.js - Configuration
├── DataService.js - Data operations
├── SystemController.js - System management
└── *.html - Frontend templates
```

---

## 🛠️ Common Commands

```bash
npm run pull    # Pull from GAS
npm run push    # Push to GAS (use before git commit)
npm run open    # Open GAS editor
npm run logs    # View execution logs
```

**Workflow**: Explore → Plan (TodoWrite) → Code → `clasp push` → Test → Git commit (source only)

---

## 🔧 clasp + Git

### Git Workflow

- **Ignore**: `.clasp.json`, `.clasprc.json` (contains scriptId - sensitive)
- **Commit**: `src/**/*.js`, `src/**/*.html`, `src/appsscript.json`, `.clasp.json.template`
- **File extension**: Use `.js` (clasp standard), GAS editor shows as `.gs`

### .claspignore

```gitignore
**/**
!appsscript.json
!**/*.js
!**/*.html
node_modules/**
.git/**
```

---

## 📝 Known Issues & Lessons Learned

### XSS Vulnerability (2025-12-13)

- **Issue**: 100+ `innerHTML` usages without sanitization
- **Risk**: User-generated content injection (even internal users)
- **Fix priority**: High - review all innerHTML usages, use textContent or escapeHtml()

### Environment Detection Not Working (2025-12-13)

- **Issue**: `window.location.hostname === 'localhost'` never true (GAS always script.google.com)
- **Impact**: Debug logs in production, performance metrics always on
- **Fix**: Use PropertiesService flag or deployment URL pattern

### main.js Size (2025-12-13)

- **Issue**: 2,841 lines in single file
- **Impact**: Reduced readability, harder code review
- **Status**: Functional but recommended to split for maintainability
- **Note**: Not a violation - GAS global scope allows any organization

---

## 🔗 Web App Entry Points

```text
/exec → AccessRestricted (default)
     → ?mode=login → LoginPage → Admin Panel
     → ?mode=view&userId=X → Public Board View
     → ?mode=admin&userId=X → Admin Panel
```

---

## 📚 References

**Official Google Documentation**:

- [Apps Script Best Practices](https://developers.google.com/apps-script/guides/support/best-practices)
- [HTML Service Best Practices](https://developers.google.com/apps-script/guides/html/best-practices)
- [clasp CLI](https://github.com/google/clasp)
- [Web Apps Guide](https://developers.google.com/apps-script/guides/web)

**Community Resources**:

- [Mastering Code Organization in GAS](https://geekjob.medium.com/mastering-code-organization-in-google-apps-script-c3fdeedc3115)
- [Building Reliable GAS Architecture](https://medium.com/google-developer-experts/tips-on-building-a-reliable-secure-scalable-architecture-using-google-apps-script-615afd4d4066)

**CLAUDE.md Best Practices**:

- [Anthropic: Claude Code Best Practices](https://www.anthropic.com/engineering/claude-code-best-practices)
- [Writing a Good CLAUDE.md](https://www.humanlayer.dev/blog/writing-a-good-claude-md)
- [Using CLAUDE.md Files](https://claude.com/blog/using-claude-md-files)

---

*This file is optimized for AI consumption following 2025 best practices: concise, universal instructions, file references over code snippets, living documentation.*
