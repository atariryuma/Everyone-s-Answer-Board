#!/bin/bash

# 自動生成された未使用コード削除スクリプト
# 生成日時: 2025/8/30 10:35:32
# 実行前に必ずバックアップを作成してください！

set -e  # エラー時に停止

echo "🚀 未使用コード削除スクリプトを開始します"
echo "⚠️  実行前にバックアップが作成されていることを確認してください"

# バックアップ確認
read -p "バックアップは作成済みですか？ (y/N): " backup_confirm
if [[ ! "$backup_confirm" =~ ^[Yy]$ ]]; then
    echo "❌ バックアップを作成してから再実行してください"
    exit 1
fi

# 低リスクファイルの削除
echo "🗑️ 低リスクファイルの削除中..."
echo "低リスクの削除対象ファイルはありません"

# テスト実行
echo "🧪 テスト実行中..."
if command -v npm &> /dev/null && [ -f "package.json" ]; then
    npm run test || {
        echo "❌ テストが失敗しました。削除を中止します"
        exit 1
    }
else
    echo "⚠️ npmテストが利用できません。手動でテストを実行してください"
fi

echo "✅ 低リスクファイルの削除が完了しました"
echo "📝 次に中・高リスクファイルを慎重に検討してください："
echo "  🟡 auth.gs (mediumリスク)"
echo "  🟡 backendProgressSync.js.html (mediumリスク)"
echo "  🟡 ClientOptimizer.html (mediumリスク)"
echo "  🟡 Code.gs (mediumリスク)"
echo "  🟡 config.gs (mediumリスク)"
echo "  🟡 constants.gs (mediumリスク)"
echo "  🟡 constants.js.html (mediumリスク)"
echo "  🟡 Core.gs (mediumリスク)"
echo "  🟡 database.gs (mediumリスク)"
echo "  🟡 debugConfig.gs (mediumリスク)"
echo "  🟡 ErrorBoundary.html (mediumリスク)"
echo "  🟡 errorHandler.gs (mediumリスク)"
echo "  🟡 errorMessages.js.html (mediumリスク)"
echo "  🟡 lockManager.gs (mediumリスク)"
echo "  🟡 monitoring.gs (mediumリスク)"
echo "  🟡 page.css.html (mediumリスク)"
echo "  🟡 Page.html (mediumリスク)"
echo "  🟡 Registration.html (mediumリスク)"
echo "  🟡 resilientExecutor.gs (mediumリスク)"
echo "  🟡 secretManager.gs (mediumリスク)"
echo "  🟡 session-utils.gs (mediumリスク)"
echo "  🟡 setup.gs (mediumリスク)"
echo "  🟡 SharedModals.html (mediumリスク)"
echo "  🟡 ulog.gs (mediumリスク)"
echo "  🟡 UnifiedCache.js.html (mediumリスク)"
echo "  🟡 unifiedCacheManager.gs (mediumリスク)"
echo "  🟡 unifiedUtilities.gs (mediumリスク)"
echo "  🟡 Unpublished.html (mediumリスク)"
echo "  🟡 url.gs (mediumリスク)"

echo "完了！"
