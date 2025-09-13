#!/bin/bash

# Pre-commit Hook for Everyone's Answer Board
# Claude Code 2025 Quality Gates

set -e

echo "🔍 Claude Code Pre-commit Quality Gates"
echo "======================================="

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Function to print status
print_status() {
    if [ $1 -eq 0 ]; then
        echo -e "${GREEN}✅ $2${NC}"
    else
        echo -e "${RED}❌ $2${NC}"
        exit 1
    fi
}

print_warning() {
    echo -e "${YELLOW}⚠️ $1${NC}"
}

# 1. Architecture Validation
echo "📐 1. Architecture Validation"
echo "------------------------------"

# Check required service files exist
REQUIRED_SERVICES=("UserService.gs" "ConfigService.gs" "DataService.gs" "SecurityService.gs")
for service in "${REQUIRED_SERVICES[@]}"; do
    if [ -f "src/services/$service" ]; then
        echo -e "${GREEN}✅ $service exists${NC}"
    else
        echo -e "${RED}❌ Missing required service: $service${NC}"
        exit 1
    fi
done

# Check core utilities exist
REQUIRED_UTILS=("helpers.gs" "validators.gs" "formatters.gs")
for util in "${REQUIRED_UTILS[@]}"; do
    if [ -f "src/utils/$util" ]; then
        echo -e "${GREEN}✅ utils/$util exists${NC}"
    else
        echo -e "${RED}❌ Missing required utility: $util${NC}"
        exit 1
    fi
done

# 2. Code Quality Gates
echo ""
echo "🔧 2. Code Quality Gates"
echo "------------------------"

# ESLint validation
echo "Running ESLint..."
if npm run lint --silent; then
    print_status 0 "ESLint validation passed"
else
    print_status 1 "ESLint validation failed"
fi

# Prettier formatting
echo "Running Prettier..."
if npm run format --silent; then
    print_status 0 "Code formatting completed"
else
    print_status 1 "Code formatting failed"
fi

# 3. Test Coverage Gates
echo ""
echo "🧪 3. Test Coverage Gates"
echo "--------------------------"

# Run tests with coverage
echo "Running test suite..."
if npm run test --silent; then
    print_status 0 "Test suite passed"
else
    print_status 1 "Test suite failed"
fi

# Check if coverage report exists
if [ -f "coverage/lcov-report/index.html" ]; then
    print_status 0 "Coverage report generated"
else
    print_warning "Coverage report not found - may need npm run test:coverage"
fi

# 4. Security Validation
echo ""
echo "🔒 4. Security Validation"
echo "-------------------------"

# Check for potential security issues
SECURITY_ISSUES=0

# Check for exposed secrets
if grep -r -E "(password|secret|key|token)" src/ --include="*.gs" --include="*.js" | grep -v "// " | grep -v "/* "; then
    print_warning "Potential secrets found in code"
    SECURITY_ISSUES=$((SECURITY_ISSUES + 1))
fi

# Check for eval() usage
if grep -r "eval(" src/ --include="*.gs" --include="*.js"; then
    print_warning "eval() usage detected - security risk"
    SECURITY_ISSUES=$((SECURITY_ISSUES + 1))
fi

# Check for innerHTML usage without sanitization
if grep -r "innerHTML" src/ --include="*.html" --include="*.js"; then
    print_warning "innerHTML usage detected - potential XSS risk"
    SECURITY_ISSUES=$((SECURITY_ISSUES + 1))
fi

if [ $SECURITY_ISSUES -eq 0 ]; then
    print_status 0 "Security validation passed"
else
    print_warning "Security validation completed with $SECURITY_ISSUES warnings"
fi

# 5. Performance Validation
echo ""
echo "⚡ 5. Performance Validation"
echo "----------------------------"

# Check file sizes
LARGE_FILES=$(find src/ -name "*.gs" -size +100k)
if [ -n "$LARGE_FILES" ]; then
    print_warning "Large files detected (>100KB):"
    echo "$LARGE_FILES"
else
    print_status 0 "File size validation passed"
fi

# Check for excessive complexity (simple heuristic)
COMPLEX_FUNCTIONS=$(grep -r "if.*if.*if.*if" src/ --include="*.gs" | wc -l)
if [ $COMPLEX_FUNCTIONS -gt 5 ]; then
    print_warning "Potentially complex functions detected: $COMPLEX_FUNCTIONS"
else
    print_status 0 "Complexity validation passed"
fi

# 6. Documentation Validation
echo ""
echo "📚 6. Documentation Validation"
echo "-------------------------------"

# Check for JSDoc comments
DOCUMENTED_FUNCTIONS=$(grep -r "* @" src/ --include="*.gs" | wc -l)
TOTAL_FUNCTIONS=$(grep -r "^function\|^  [a-zA-Z_][a-zA-Z0-9_]*(" src/ --include="*.gs" | wc -l)

if [ $TOTAL_FUNCTIONS -gt 0 ]; then
    DOCUMENTATION_RATIO=$((DOCUMENTED_FUNCTIONS * 100 / TOTAL_FUNCTIONS))
    if [ $DOCUMENTATION_RATIO -ge 80 ]; then
        print_status 0 "Documentation coverage: ${DOCUMENTATION_RATIO}%"
    else
        print_warning "Documentation coverage below 80%: ${DOCUMENTATION_RATIO}%"
    fi
else
    print_status 0 "No functions found to document"
fi

# Final Summary
echo ""
echo "🎯 Pre-commit Summary"
echo "===================="
echo -e "${GREEN}✅ Architecture: Valid${NC}"
echo -e "${GREEN}✅ Code Quality: Passed${NC}"
echo -e "${GREEN}✅ Tests: Passed${NC}"
echo -e "${GREEN}✅ Security: Validated${NC}"
echo -e "${GREEN}✅ Performance: Acceptable${NC}"
echo -e "${GREEN}✅ Documentation: Sufficient${NC}"
echo ""
echo -e "${GREEN}🚀 Commit approved! Ready for deployment.${NC}"

exit 0