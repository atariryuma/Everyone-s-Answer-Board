#!/bin/bash
# 新機能開発開始の自動化

if [ -z "$1" ]; then
  echo "使用法: ./scripts/new-feature.sh <機能名>"
  exit 1
fi

FEATURE_NAME="feature/$1"

echo "🔍 現在のブランチ状況を確認中..."
git status --porcelain
if [ $? -ne 0 ]; then
  echo "❌ Gitリポジトリではありません"
  exit 1
fi

# 未コミットの変更がある場合は警告
if ! git diff-index --quiet HEAD --; then
  echo "⚠️  未コミットの変更があります。続行しますか？ (y/N)"
  read -r response
  if [[ "$response" != "y" ]]; then
    echo "🛑 中断しました。変更をコミットしてから再実行してください"
    exit 1
  fi
fi

# mainブランチを最新に更新
echo "📥 mainブランチを最新に更新中..."
git checkout main
git pull

# ブランチ作成・切り替え
echo "🌲 フィーチャーブランチ '$FEATURE_NAME' を作成中..."
git checkout -b "$FEATURE_NAME"

echo "🚀 新機能 '$1' の開発環境が準備完了！"
echo ""
echo "📋 次の手順で開発を開始してください："
echo "1. 別ターミナルで: npm run test:watch"
echo "2. Claude Codeでテスト→実装→修正のサイクルを開始"
echo "3. 完了後: ./scripts/merge-feature.sh $1"
echo ""
echo "💡 ヒント: 'feature/ユーザー認証 ブランチで開発します' とClaudeに伝えると効果的です"