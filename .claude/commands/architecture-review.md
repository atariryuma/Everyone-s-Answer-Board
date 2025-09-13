# Architecture Review Command

Comprehensive architectural analysis and recommendations for Google Apps Script project.

## Execution Steps

1. **System Architecture Analysis**
   - Analyze service layer dependencies
   - Verify separation of concerns
   - Check controller/service boundaries
   - Validate data flow patterns

2. **Code Quality Assessment**
   - Run `npm run lint` and analyze patterns
   - Check for duplicate functions across services
   - Verify naming consistency
   - Assess error handling patterns

3. **Performance & Security Review**
   - Analyze caching strategies
   - Review authentication flows
   - Check input validation patterns
   - Assess GAS API usage efficiency

4. **Technical Debt Identification**
   - Identify refactoring opportunities
   - Check for deprecated patterns
   - Analyze test coverage gaps
   - Review documentation completeness

## Analysis Focus Areas

### Service Architecture
- **Controllers**: HTTP routing and request handling only
- **Services**: Business logic and domain operations
- **Infrastructure**: Database and external API integration
- **Utils**: Pure functions and shared utilities

### Common Anti-patterns to Check
- Controllers containing business logic
- Services making HTTP responses
- Direct database access from controllers
- Duplicate functionality across layers

### GAS-Specific Concerns
- Script execution time limits
- Memory usage optimization
- Service account vs user authentication
- HTML Service security patterns

## Output Requirements

Generate detailed report covering:
- Architecture compliance score (0-100)
- Critical issues requiring immediate attention
- Recommended refactoring priorities
- Performance optimization opportunities
- Security enhancement suggestions

## Success Criteria

- [ ] All layers properly separated
- [ ] No duplicate functionality detected
- [ ] Security patterns implemented correctly
- [ ] Performance bottlenecks identified
- [ ] Technical debt documented with priorities