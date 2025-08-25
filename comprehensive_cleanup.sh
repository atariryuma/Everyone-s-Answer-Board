#!/bin/bash

FILE="src/Core.gs"

# 1. 孤立したオブジェクトプロパティのパターンを削除（インデント後にプロパティ名:値の形式）
sed -i '' '/^\s\{6,\}[a-zA-Z][a-zA-Z0-9_]*:\s.*,$/d' "$FILE"

# 2. 孤立したオブジェクトプロパティの最終行（カンマなし）を削除
sed -i '' '/^\s\{6,\}[a-zA-Z][a-zA-Z0-9_]*:\s.*[^,]$/d' "$FILE" 

# 3. 孤立した関数呼び出しを削除
sed -i '' '/^\s\{6,\}[a-zA-Z][a-zA-Z0-9_]*,.*\.[a-zA-Z]/d' "$FILE"

# 4. 孤立した });を削除
sed -i '' '/^\s\{6,\}});$/d' "$FILE"

# 5. 空行の重複を削除
awk '!NF {if (!blank++) print; next} {blank=0; print}' "$FILE" > "$FILE.tmp" && mv "$FILE.tmp" "$FILE"

echo "Comprehensive cleanup completed"
