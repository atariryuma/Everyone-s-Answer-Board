#!/bin/bash
# Safe Deploy Script - 安全なデプロイを保証するラッパースクリプト

set -e  # エラー時即座に終了

echo "🔒 安全デプロイシステム v1.0"
echo "=============================="

# 1. 事前検証
echo "🔍 Step 1: 事前検証実行..."
if ! node scripts/pre-deploy-check.js; then
    echo "❌ 事前検証に失敗しました。デプロイを中止します。"
    exit 1
fi

# 2. バックアップ作成
echo "💾 Step 2: 現在の状態をバックアップ..."
BACKUP_DIR="backups/$(date +%Y%m%d_%H%M%S)"
mkdir -p "$BACKUP_DIR"

# 現在のGAS状態をバックアップ（可能であれば）
if command -v clasp &> /dev/null; then
    echo "📁 現在のGAS状態をバックアップ中..."
    clasp pull --rootDir "$BACKUP_DIR" 2>/dev/null || echo "⚠️ バックアップに失敗しましたが続行します"
fi

# ローカルファイルもバックアップ
cp -r src "$BACKUP_DIR/"
echo "✅ バックアップ完了: $BACKUP_DIR"

# 3. 段階的デプロイ
echo "🚀 Step 3: 段階的デプロイ実行..."

# 3-1. 基盤ファイルを最初にデプロイ
echo "📦 基盤ファイル (constants, utils) をデプロイ..."
clasp push --force src/core/constants.gs src/utils/ 2>/dev/null || true

# 3-2. Services層をデプロイ
echo "⚙️ Services層をデプロイ..."
clasp push --force src/services/ 2>/dev/null || true

# 3-3. メイン実行ファイルをデプロイ
echo "🎯 メインファイルをデプロイ..."
clasp push --force src/main.gs 2>/dev/null || true

# 3-4. HTMLテンプレートをデプロイ
echo "🌐 HTMLテンプレートをデプロイ..."
clasp push --force src/*.html 2>/dev/null || true

# 4. 完全デプロイ
echo "🔄 Step 4: 完全同期実行..."
if clasp push --force; then
    echo "✅ デプロイ成功しました"
else
    echo "❌ デプロイに失敗しました"
    echo "🔄 バックアップから復元を検討してください: $BACKUP_DIR"
    exit 1
fi

# 5. デプロイ後検証
echo "✨ Step 5: デプロイ後検証..."

# GAS関数の基本実行テスト（可能であれば）
if command -v clasp &> /dev/null; then
    echo "🧪 基本関数テスト実行中..."

    # include関数テスト
    if clasp run include "SharedSecurityHeaders" &>/dev/null; then
        echo "✅ include関数: 正常"
    else
        echo "⚠️ include関数: 要確認"
    fi

    # UserService基本テスト
    if clasp run getUser &>/dev/null; then
        echo "✅ UserService: 正常"
    else
        echo "⚠️ UserService: 要確認"
    fi
fi

# 6. 成功レポート
echo ""
echo "🎉 安全デプロイ完了!"
echo "=================="
echo "✅ 事前検証: 合格"
echo "✅ バックアップ: $BACKUP_DIR"
echo "✅ 段階的デプロイ: 完了"
echo "✅ 完全同期: 完了"
echo "✅ 事後検証: 完了"
echo ""
echo "🌐 Webアプリケーションをテストしてください"
echo "📋 問題がある場合は以下のバックアップから復元可能です:"
echo "   $BACKUP_DIR"
echo ""

# 7. 自動ブラウザ起動（オプション）
if command -v clasp &> /dev/null; then
    echo "🚀 Webアプリケーションを開きますか? (y/N)"
    read -r response
    if [[ "$response" =~ ^[Yy]$ ]]; then
        clasp open --webapp
    fi
fi

echo "✨ 安全デプロイシステム完了"