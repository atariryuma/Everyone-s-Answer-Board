#!/bin/bash
# 機能完成後のマージ・デプロイ自動化

if [ -z "$1" ]; then
  echo "使用法: ./scripts/merge-feature.sh <機能名>"
  exit 1
fi

FEATURE_NAME="feature/$1"
CURRENT_BRANCH=$(git branch --show-current)

# 現在のブランチ確認
if [ "$CURRENT_BRANCH" != "$FEATURE_NAME" ]; then
  echo "❌ エラー: $FEATURE_NAME ブランチに切り替えてください"
  echo "現在のブランチ: $CURRENT_BRANCH"
  exit 1
fi

echo "🔍 フィーチャーブランチ '$1' のマージ準備中..."

# 最終チェック
echo "📋 最終品質チェック中..."
npm run check
if [ $? -ne 0 ]; then
  echo "❌ テストが失敗しました。修正してから再実行してください"
  exit 1
fi

# コミット（未コミットがあれば）
if ! git diff-index --quiet HEAD --; then
  echo "📝 未コミットの変更をコミットしています..."
  git add .
  git commit -m "feat: $1 の実装完了

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"
fi

# mainブランチに切り替え＆最新取得
echo "🔄 mainブランチに切り替え中..."
git checkout main
git pull  # リモートの最新を取得

# mainでも念のためテスト実行
echo "🧪 mainブランチでの品質確認中..."
npm run check
if [ $? -ne 0 ]; then
  echo "❌ mainブランチでテストが失敗しました"
  git checkout "$FEATURE_NAME"
  echo "🔄 フィーチャーブランチに戻りました"
  exit 1
fi

# マージ実行
echo "🔀 '$FEATURE_NAME' をmainにマージ中..."
git merge "$FEATURE_NAME"

if [ $? -ne 0 ]; then
  echo "❌ マージに失敗しました。コンフリクトを解決してください"
  exit 1
fi

# デプロイ
echo "🚀 GASへデプロイ中..."
npm run deploy

if [ $? -eq 0 ]; then
  echo "✅ デプロイ成功！"
  echo ""
  echo "🧹 フィーチャーブランチ '$FEATURE_NAME' をクリーンアップしますか？ (y/N)"
  read -r response
  if [[ "$response" == "y" ]]; then
    git branch -d "$FEATURE_NAME"
    echo "🗑️ $FEATURE_NAME ブランチを削除しました"
  else
    echo "📁 $FEATURE_NAME ブランチは保持されました"
  fi
  
  echo ""
  echo "🎉 機能 '$1' のリリース完了！"
  echo "📊 現在のブランチ状況:"
  git branch -a
else
  echo "❌ デプロイに失敗しました"
  echo "💡 手動でデプロイを確認してください: npm run deploy"
  exit 1
fi