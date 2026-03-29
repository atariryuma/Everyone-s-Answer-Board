#!/bin/bash
set -e

# 構文チェック
errors=0
for file in src/*.js; do
  if ! node --check "$file" 2>/dev/null; then
    echo "Syntax error: $file"
    errors=$((errors + 1))
  fi
done
if [ "$errors" -gt 0 ]; then exit 1; fi

# テスト
npm test --silent
