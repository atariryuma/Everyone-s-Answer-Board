#!/bin/bash
# ホットフィックスのコミット・デプロイ自動化

echo "🔥 ホットフィックス デプロイ"

# mainブランチ確認
CURRENT_BRANCH=$(git branch --show-current)
if [ "$CURRENT_BRANCH" != "main" ]; then
  echo "❌ mainブランチから実行してください"
  exit 1
fi

# 変更があるかチェック
if git diff-index --quiet HEAD --; then
  echo "❌ コミットする変更がありません"
  exit 1
fi

# 修正内容の確認
echo "📝 修正内容:"
git diff --name-only
echo ""

# テスト実行
echo "🧪 最終テスト実行中..."
npm run check
if [ $? -ne 0 ]; then
  echo "❌ テストが失敗しました。修正してください"
  exit 1
fi

# コミットメッセージ入力
echo "📋 修正内容を簡潔に説明してください:"
read -r fix_description

if [ -z "$fix_description" ]; then
  echo "❌ 修正内容の説明が必要です"
  exit 1
fi

# コミット実行
git add .
git commit -m "fix: $fix_description

🔥 Hotfix deployment

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"

# デプロイ
echo "🚀 緊急デプロイ実行中..."
npm run deploy

if [ $? -eq 0 ]; then
  echo "✅ ホットフィックス デプロイ成功！"
  echo "📝 修正内容: $fix_description"
  
  # リモートにプッシュするか確認
  echo ""
  echo "🌐 リモートリポジトリにプッシュしますか？ (y/N)"
  read -r response
  if [[ "$response" == "y" ]]; then
    git push
    echo "📤 リモートにプッシュしました"
  fi
  
  echo "🎉 ホットフィックス完了！"
else
  echo "❌ デプロイに失敗しました"
  echo "💡 手動確認してください: npm run deploy"
  exit 1
fi