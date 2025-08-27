#!/bin/bash
# 緊急修正（ホットフィックス）用スクリプト

echo "🔥 ホットフィックス モード"
echo "⚠️  このモードはmainブランチに直接変更を加えます"
echo "💡 小さなバグ修正や緊急対応のみに使用してください"
echo ""

# 現在のブランチ確認
CURRENT_BRANCH=$(git branch --show-current)
if [ "$CURRENT_BRANCH" != "main" ]; then
  echo "❌ mainブランチに切り替えてください"
  echo "現在のブランチ: $CURRENT_BRANCH"
  exit 1
fi

# 最新取得
echo "📥 mainブランチを最新に更新中..."
git pull

echo "🚀 開発環境準備完了！"
echo ""
echo "📋 ホットフィックス手順："
echo "1. 別ターミナルで: npm run test:watch"
echo "2. Claude Codeで修正を実施"
echo "3. テストが通ることを確認"
echo "4. このスクリプト終了後: ./scripts/hotfix-deploy.sh で即座にデプロイ"
echo ""
echo "⚠️  大きな変更の場合は Ctrl+C で中断し、feature ブランチを使用してください"