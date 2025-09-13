#!/bin/bash

# Claude Code Pre-Deploy Hook
# Automatically runs before any deployment to prevent common issues

set -e

echo "🔍 Running pre-deployment checks..."

# 1. Check for duplicate declarations
echo "  → Checking for duplicate const/function declarations..."
if rg -n "^(const|function|class) " src/ | sort | uniq -d -w 20 | grep -q .; then
    echo "❌ Duplicate declarations found:"
    rg -n "^(const|function|class) " src/ | sort | uniq -d -w 20
    exit 1
fi

# 2. Verify authentication flow
echo "  → Verifying authentication flow pattern..."
if ! rg -q "handleLoginModeWithTemplate.*default_entry_point" src/main.gs; then
    echo "❌ Authentication flow may be broken - check /exec routing"
    exit 1
fi

# 3. Test basic compilation
echo "  → Testing basic syntax..."
if ! node scripts/pre-deploy-check.js > /dev/null 2>&1; then
    echo "❌ Pre-deploy validation failed"
    exit 1
fi

echo "✅ Pre-deployment checks passed"