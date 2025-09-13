# Refactor Service Command

## Description
Guided service layer refactoring and architecture improvement workflow.

## Usage
```
/refactor-service [service-name] [operation]
```

Options:
- `service-name`: UserService, ConfigService, DataService, SecurityService
- `operation`: analyze, split, optimize, test, migrate

## Implementation

This command provides guided refactoring workflows:

### 1. Service Analysis
```bash
/refactor-service UserService analyze
```
- Analyze service size and complexity
- Identify single responsibility violations
- Check dependency relationships
- Measure performance metrics

### 2. Service Splitting
```bash
/refactor-service DataService split
```
- Identify logical boundaries for splitting
- Suggest new service organization
- Generate migration plan
- Validate architectural consistency

### 3. Service Optimization
```bash
/refactor-service ConfigService optimize
```
- Performance bottleneck identification
- Cache utilization analysis
- Memory usage optimization
- Query optimization suggestions

### 4. Service Testing
```bash
/refactor-service SecurityService test
```
- Generate comprehensive test cases
- Identify missing test coverage
- Performance benchmark creation
- Integration test validation

### 5. Legacy Migration
```bash
/refactor-service migrate Core.gs
```
- Analyze legacy code for service extraction
- Generate migration roadmap
- Identify breaking changes
- Create backward compatibility layer

## Refactoring Principles

### SOLID Compliance Check
- **S**ingle Responsibility: Each service has one clear purpose
- **O**pen/Closed: Services are open for extension, closed for modification
- **L**iskov Substitution: Service interfaces are consistently implementable
- **I**nterface Segregation: No unnecessary dependencies
- **D**ependency Inversion: Depend on abstractions, not concretions

### Code Quality Metrics
- Cyclomatic complexity < 10 per function
- Service file size < 500 lines
- Function length < 50 lines
- Test coverage > 90%
- Zero circular dependencies

## Expected Output

```
🔧 Service Refactoring Analysis
├── 📊 Current Service Metrics
│   ├── Lines of Code: 245
│   ├── Function Count: 12
│   ├── Complexity Score: 8.2
│   └── Test Coverage: 85%
├── 🎯 Refactoring Recommendations
│   ├── Split authentication logic → AuthService
│   ├── Extract validation → ValidationService
│   └── Optimize cache usage → +15% performance
├── 📋 Migration Plan
│   ├── Phase 1: Extract AuthService (2 hours)
│   ├── Phase 2: Update dependencies (1 hour)
│   └── Phase 3: Test & validate (1 hour)
└── ✅ Quality Gates: All checks passed
```

## Safety Features

- **Impact Analysis**: Shows affected components before refactoring
- **Rollback Plan**: Automatic backup and restoration capabilities
- **Test Coverage**: Ensures no functionality is lost during refactoring
- **Dependency Tracking**: Validates all dependencies remain functional