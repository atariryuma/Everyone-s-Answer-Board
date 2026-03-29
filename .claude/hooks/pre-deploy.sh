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

# GAS関数名の重複チェック（GASはグローバルスコープなので名前衝突は致命的）
dupes=$(grep -rh "^function [a-zA-Z_]" src/*.js | sort | uniq -d)
if [ -n "$dupes" ]; then
  echo "Duplicate function names found:"
  echo "$dupes"
  exit 1
fi
