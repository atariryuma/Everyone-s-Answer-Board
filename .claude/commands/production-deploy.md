# Production Deployment Workflow

Safe production deployment process for Google Apps Script with comprehensive validation and rollback capabilities.

## Pre-deployment Requirements

### Branch Safety
- Create deployment branch: `git checkout -b deploy/$(date +%Y%m%d-%H%M)`
- Ensure main branch is clean and up-to-date
- All changes committed with proper attribution

### Quality Gates (ALL MUST PASS)
```bash
npm run check           # Full quality suite (tests, lint, typecheck)
./scripts/safe-deploy.sh --dry-run  # Deployment simulation
clasp status           # Verify GAS project connection
```

## Deployment Process

### Phase 1: Pre-flight Validation
1. **System Health Check**
   - Verify all 113 tests pass
   - Confirm ESLint errors < 5
   - Check TypeScript compilation
   - Validate function exports for HTML

2. **GAS Environment Preparation**
   ```bash
   clasp login           # Ensure authenticated
   clasp status          # Verify project connection
   clasp pull            # Get latest remote state
   ```

3. **Backup Current State**
   ```bash
   git tag backup/pre-deploy-$(date +%Y%m%d-%H%M)
   clasp versions        # Document current GAS version
   ```

### Phase 2: Safe Deployment
1. **Execute Safe Deploy Script**
   ```bash
   ./scripts/safe-deploy.sh --production
   ```

2. **Monitor Deployment**
   - Watch script output for errors
   - Verify file upload completion
   - Check for GAS execution errors

3. **Initial Smoke Tests**
   ```bash
   clasp run setupApplication  # Test core function
   clasp logs --watch         # Monitor runtime logs
   ```

### Phase 3: Production Validation
1. **Web App Functionality**
   - Test login flow: `/exec` → login → main board
   - Verify admin panel access
   - Test data operations (add reaction, toggle highlight)
   - Validate user creation and authentication

2. **API Endpoint Testing**
   - Test all google.script.run functions
   - Verify error handling responses
   - Check data persistence
   - Validate security restrictions

3. **Performance Monitoring**
   ```bash
   clasp logs --json | grep "execution_time"
   ```
   - Monitor execution times
   - Check memory usage patterns
   - Verify caching effectiveness

## Post-deployment Actions

### Success Path
1. **Tag Successful Deployment**
   ```bash
   git tag production/$(date +%Y%m%d-%H%M)
   git push origin --tags
   ```

2. **Update Documentation**
   - Record deployment notes
   - Update version numbers
   - Document any configuration changes

3. **Monitor for 24 hours**
   - Check `clasp logs` regularly
   - Monitor user feedback
   - Verify system stability

### Rollback Procedures

#### Immediate Rollback (< 10 minutes)
```bash
git checkout backup/pre-deploy-YYYYMMDD-HHMM
clasp push --force
clasp logs --watch
```

#### Extended Rollback (> 10 minutes)
1. Identify specific failing component
2. Selective revert of problematic changes
3. Incremental deployment of fixes
4. Full system validation

## Deployment Checklist

### Pre-deployment ✅
- [ ] All tests pass (113/113)
- [ ] ESLint errors < 5
- [ ] TypeScript compilation clean
- [ ] Backup created
- [ ] GAS project authenticated
- [ ] Safe-deploy script tested

### During Deployment ✅
- [ ] Script execution completed without errors
- [ ] All files uploaded successfully
- [ ] No GAS execution errors in logs
- [ ] Initial smoke tests pass

### Post-deployment ✅
- [ ] Web app login flow works
- [ ] Admin panel accessible
- [ ] Data operations functional
- [ ] User creation/authentication working
- [ ] Performance metrics within acceptable range
- [ ] No critical errors in 24-hour monitoring period

## Emergency Contacts & Procedures

### Critical Error Response
1. Immediate rollback to last known good state
2. Document error details and reproduction steps
3. Create hotfix branch for urgent repairs
4. Test hotfix in isolation before deployment

### Monitoring Commands
```bash
# Real-time log monitoring
clasp logs --watch --json

# Error pattern detection
clasp logs | grep -E "(ERROR|EXCEPTION|undefined)"

# Performance analysis
clasp logs --json | jq '.executionTime'
```

## Success Criteria

- ✅ Zero critical errors in production logs
- ✅ All user workflows functional
- ✅ Performance within established benchmarks
- ✅ 24-hour stability period completed
- ✅ Rollback procedures validated and ready