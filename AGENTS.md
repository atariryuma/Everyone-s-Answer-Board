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