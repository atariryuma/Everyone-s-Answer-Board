#!/bin/bash
# 孤立したプロパティとデバッグログの痕跡を削除

# 孤立した関数呼び出しのパラメータを削除
sed -i '' '/^\s*[a-zA-Z][a-zA-Z0-9_]*,.*[a-zA-Z][a-zA-Z0-9_]*(/d' src/Core.gs

# 孤立した })呼び出しを削除
sed -i '' '/^\s*});$/d' src/Core.gs

# 空行の連続を1行に削減
sed -i '' '/^$/N;/^\n$/d' src/Core.gs

echo "Fixed orphaned structures in Core.gs"
