#!/bin/bash
# ブランチクリーンアップ用スクリプト

echo "🧹 プロジェクトブランチクリーンアップ"
echo "=================================="

# 現在のブランチを確認
CURRENT_BRANCH=$(git branch --show-current)
echo "📍 現在のブランチ: $CURRENT_BRANCH"

if [ "$CURRENT_BRANCH" != "main" ]; then
  echo "❌ mainブランチに切り替えてから実行してください"
  exit 1
fi

echo ""
echo "🔍 マージ済みブランチの確認中..."

# マージ済みのローカルブランチを表示（mainを除く）
MERGED_BRANCHES=$(git branch --merged main | grep -v "main" | grep -v "\*" | sed 's/^[[:space:]]*//')

if [ -z "$MERGED_BRANCHES" ]; then
  echo "✅ クリーンアップ対象のマージ済みブランチはありません"
else
  echo "📋 マージ済みブランチ（削除候補）:"
  echo "$MERGED_BRANCHES" | nl
  echo ""
  
  echo "⚠️  これらのブランチを削除しますか？ (y/N)"
  read -r response
  
  if [[ "$response" == "y" ]]; then
    echo "🗑️  マージ済みブランチを削除中..."
    echo "$MERGED_BRANCHES" | while IFS= read -r branch; do
      if [ -n "$branch" ]; then
        echo "  削除: $branch"
        git branch -d "$branch"
      fi
    done
    echo "✅ マージ済みブランチの削除完了"
  else
    echo "🛑 削除をキャンセルしました"
  fi
fi

echo ""
echo "🔍 未マージブランチの確認中..."

# 未マージのローカルブランチを表示
UNMERGED_BRANCHES=$(git branch --no-merged main | grep -v "main" | grep -v "\*" | sed 's/^[[:space:]]*//')

if [ -z "$UNMERGED_BRANCHES" ]; then
  echo "✅ 未マージブランチはありません"
else
  echo "📋 未マージブランチ（要確認）:"
  echo "$UNMERGED_BRANCHES" | nl
  echo ""
  echo "💡 これらのブランチは手動で確認してください"
  echo "   - 必要なものは feature/ ブランチにリネーム"
  echo "   - 不要なものは git branch -D <ブランチ名> で強制削除"
fi

echo ""
echo "📊 現在のブランチ状況:"
LOCAL_COUNT=$(git branch | wc -l | xargs)
REMOTE_COUNT=$(git branch -r | wc -l | xargs)
echo "  ローカルブランチ: $LOCAL_COUNT 個"
echo "  リモートブランチ: $REMOTE_COUNT 個"

echo ""
echo "💡 推奨事項:"
echo "  - 新機能は feature/ プレフィックスを使用"
echo "  - 実験は experiment/ プレフィックスを使用" 
echo "  - 修正は fix/ プレフィックスを使用"
echo "  - 定期的にこのスクリプトでクリーンアップ"

echo ""
echo "🎉 ブランチクリーンアップ完了！"