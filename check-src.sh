#!/bin/bash

echo "🔍 /src ディレクトリ簡易チェック"
echo "================================"

SRC_DIR="./src"
TOTAL_FILES=0
EXISTING_FILES=0

# ファイル存在チェック関数
check_file() {
    local file="$1"
    local description="$2"
    TOTAL_FILES=$((TOTAL_FILES + 1))
    
    if [ -f "$SRC_DIR/$file" ]; then
        local size=$(ls -lh "$SRC_DIR/$file" | awk '{print $5}')
        echo "✅ $file ($size) - $description"
        EXISTING_FILES=$((EXISTING_FILES + 1))
    else
        echo "❌ $file - ファイル不存在"
    fi
}

# HTMLファイルチェック
echo ""
echo "🗂️ HTMLファイル"
echo "----------------"
check_file "AdminPanel.html" "管理画面"
check_file "AppSetupPage.html" "セットアップ画面" 
check_file "LoginPage.html" "ログイン画面"
check_file "Page.html" "メイン画面"

# JavaScriptファイルチェック  
echo ""
echo "📜 JavaScriptファイル"
echo "--------------------"
check_file "adminPanel-core.js.html" "コア機能"
check_file "adminPanel-api.js.html" "API連携"
check_file "adminPanel-events.js.html" "イベント処理"
check_file "page.js.html" "ページ機能"
check_file "UnifiedStyles.html" "統合スタイル"

# GASファイルチェック
echo ""
echo "⚙️ GASファイル" 
echo "-------------"
check_file "main.gs" "メインエントリーポイント"
check_file "Core.gs" "コア業務ロジック"
check_file "database.gs" "データベース操作"
check_file "config.gs" "設定管理"

# その他重要ファイル
echo ""
echo "🔧 その他重要ファイル"
echo "-------------------"
check_file "appsscript.json" "GAS設定"
check_file "SharedSecurityHeaders.html" "セキュリティヘッダー"
check_file "SharedTailwindConfig.html" "Tailwind設定"

# 関数チェック（簡易版）
echo ""
echo "🎯 重要関数チェック"
echo "------------------"

if [ -f "$SRC_DIR/main.gs" ]; then
    if grep -q "function doGet" "$SRC_DIR/main.gs"; then
        echo "✅ doGet関数 - HTTPリクエストハンドラー"
    else
        echo "❌ doGet関数が見つかりません"
    fi
    
    if grep -q "function include" "$SRC_DIR/main.gs"; then
        echo "✅ include関数 - HTMLインクルード"
    else
        echo "❌ include関数が見つかりません"
    fi
fi

if [ -f "$SRC_DIR/Core.gs" ]; then
    if grep -q "function verifyUserAccess" "$SRC_DIR/Core.gs"; then
        echo "✅ verifyUserAccess関数 - ユーザー認証"
    else
        echo "❌ verifyUserAccess関数が見つかりません"
    fi
fi

if [ -f "$SRC_DIR/database.gs" ]; then
    if grep -q "function createUser" "$SRC_DIR/database.gs"; then
        echo "✅ createUser関数 - ユーザー作成"
    else
        echo "❌ createUser関数が見つかりません"
    fi
fi

# サマリー
echo ""
echo "📊 チェック結果サマリー"
echo "====================="
echo "ファイル存在率: $EXISTING_FILES/$TOTAL_FILES ($(( EXISTING_FILES * 100 / TOTAL_FILES ))%)"

if [ $EXISTING_FILES -eq $TOTAL_FILES ]; then
    echo "🎉 すべてのファイルが存在しています"
elif [ $EXISTING_FILES -ge $((TOTAL_FILES * 8 / 10)) ]; then
    echo "✅ ほとんどのファイルが存在しています"  
elif [ $EXISTING_FILES -ge $((TOTAL_FILES / 2)) ]; then
    echo "⚠️ 一部のファイルが不足しています"
else
    echo "❌ 多くのファイルが不足しています"
fi

echo ""
echo "💡 詳細チェックが必要な場合は以下を実行:"
echo "   node check-src.js    (詳細解析)"
echo "   open check-src.html  (ブラウザ検証)"