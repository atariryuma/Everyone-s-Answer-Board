# Test and Deploy Workflow

Execute complete testing and deployment cycle for Google Apps Script project.

## Execution Steps

1. **Quality Check**: Run `npm run check`
   - Verify all 113 tests pass
   - Check ESLint errors are minimal (< 5)
   - Confirm TypeScript types are valid

2. **Pre-deployment Validation**
   - Ensure no undefined functions
   - Verify API Gateway pattern compliance
   - Check architecture boundaries

3. **Safe Deployment**
   - Execute `./scripts/safe-deploy.sh`
   - Monitor deployment logs
   - Confirm GAS project update

4. **Post-deployment Verification**
   - Run `clasp logs` for error monitoring
   - Test web app functionality
   - Verify all HTML forms work correctly

## Success Criteria

- [ ] All tests passing (113/113)
- [ ] ESLint errors < 5
- [ ] Successful clasp push
- [ ] Web app responds correctly
- [ ] No runtime errors in logs

## Rollback Plan

If deployment fails:
1. `git checkout previous-working-branch`
2. `clasp push` to restore previous version
3. Monitor logs for stability confirmation