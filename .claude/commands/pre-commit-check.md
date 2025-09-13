# Pre-commit Quality Check

Automated quality gate checks before committing changes to ensure code integrity.

## Execution Steps

1. **Code Quality Validation**
   - Run `npm run lint` and ensure < 5 errors
   - Execute `npm run typecheck` for type safety
   - Run `npm run test` and verify 113/113 tests pass
   - Check for undefined function references

2. **Security Validation**
   - Scan for hardcoded secrets or API keys
   - Verify input validation in user-facing functions
   - Check authentication flow integrity
   - Validate CORS and XSS protection patterns

3. **Architecture Compliance**
   - Verify no business logic in main.gs
   - Check controller/service layer separation
   - Validate proper error handling patterns
   - Ensure consistent naming conventions

4. **GAS-Specific Checks**
   - Verify all functions exported for HTML access
   - Check quota usage patterns (execution time, API calls)
   - Validate Service Account vs User Auth usage
   - Ensure proper caching implementation

## Automated Validations

### Critical Blockers (Must Fix)
- Test failures
- TypeScript/ESLint errors > 5
- Security vulnerabilities
- Missing function exports for HTML

### Warnings (Should Fix)
- Performance anti-patterns
- Inconsistent error handling
- Missing documentation
- Unused imports or variables

### Quality Gates
```bash
# Run all checks in parallel
npm run check                    # Full quality suite
./scripts/validate-functions.js  # Function completeness
npm run test -- --coverage      # Test coverage report
```

## Pre-commit Workflow

1. **Stage Changes**
   ```bash
   git add .
   git status  # Review staged files
   ```

2. **Quality Validation**
   - Execute all automated checks
   - Fix any critical blockers
   - Address high-priority warnings

3. **Final Verification**
   ```bash
   git diff --staged  # Review exact changes
   npm run check      # Final quality gate
   ```

4. **Commit with Standards**
   - Use conventional commit format
   - Include Co-Authored-By: Claude attribution
   - Reference issue numbers if applicable

## Success Criteria

- [ ] All tests passing (113/113)
- [ ] ESLint errors < 5
- [ ] No TypeScript type errors
- [ ] No security vulnerabilities detected
- [ ] All HTML-callable functions properly exported
- [ ] Changes reviewed and staged correctly

## Rollback Plan

If pre-commit checks fail:
1. `git reset HEAD~1` to undo commit
2. Address failing checks individually
3. Re-run quality validation
4. Retry commit process