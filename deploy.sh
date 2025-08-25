#!/bin/bash

# 新アーキテクチャのデプロイスクリプト

echo "🚀 みんなの回答ボード - 新アーキテクチャ デプロイスクリプト"
echo "================================================"

# カラー定義
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# デプロイモードの選択
echo ""
echo "デプロイモードを選択してください:"
echo "1) テスト環境（新規プロジェクト作成）"
echo "2) ステージング環境（既存プロジェクトにプッシュ）"
echo "3) 本番環境（移行フラグを有効化）"
echo "4) ロールバック（レガシーモードに戻す）"
echo ""
read -p "選択 (1-4): " mode

# バックアップ作成
backup_dir="backup/$(date +%Y%m%d_%H%M%S)"
echo -e "${YELLOW}📦 バックアップを作成中...${NC}"
mkdir -p $backup_dir
cp -r src/* $backup_dir/
echo -e "${GREEN}✅ バックアップ完了: $backup_dir${NC}"

case $mode in
  1)
    echo -e "${GREEN}テスト環境へのデプロイを開始...${NC}"
    
    # 新規プロジェクト作成
    echo "新しいGASプロジェクトを作成します..."
    clasp create --title "みんなの回答ボード-新アーキテクチャ-TEST" --rootDir ./src
    
    # 新アーキテクチャのファイルのみをプッシュ
    echo "新サービスをアップロード中..."
    clasp push --force
    
    echo -e "${GREEN}✅ テスト環境へのデプロイ完了${NC}"
    echo "Webアプリをデプロイするには以下を実行:"
    echo "clasp deploy --description 'New Architecture Test'"
    ;;
    
  2)
    echo -e "${YELLOW}ステージング環境へのデプロイを開始...${NC}"
    
    # appsscript.jsonを更新
    echo "設定ファイルを更新中..."
    cp src/appsscript-new.json src/appsscript.json
    
    # プッシュ
    echo "ファイルをアップロード中..."
    clasp push
    
    echo -e "${GREEN}✅ ステージング環境へのデプロイ完了${NC}"
    echo "移行フラグを有効にするには、GASエディタで以下を実行:"
    echo "PropertiesService.getScriptProperties().setProperty('USE_NEW_ARCHITECTURE', 'true');"
    ;;
    
  3)
    echo -e "${RED}本番環境への移行を開始...${NC}"
    
    # 確認
    read -p "本番環境に移行しますか？ (yes/no): " confirm
    if [ "$confirm" != "yes" ]; then
      echo "移行をキャンセルしました"
      exit 1
    fi
    
    # テスト実行
    echo "テストを実行中..."
    npm test tests/criticalFunctions.test.js
    if [ $? -ne 0 ]; then
      echo -e "${RED}❌ テストが失敗しました。移行を中止します。${NC}"
      exit 1
    fi
    
    # デプロイ
    echo "本番環境にデプロイ中..."
    cp src/appsscript-new.json src/appsscript.json
    clasp push
    
    echo -e "${GREEN}✅ 本番環境へのデプロイ完了${NC}"
    echo ""
    echo "次のステップ:"
    echo "1. GASエディタで移行フラグを有効化"
    echo "2. 動作確認を実施"
    echo "3. 問題がなければ旧ファイルを削除"
    ;;
    
  4)
    echo -e "${YELLOW}ロールバックを実行...${NC}"
    
    # バックアップから復元
    echo "最新のバックアップを検索中..."
    latest_backup=$(ls -t backup/ | head -1)
    
    if [ -z "$latest_backup" ]; then
      echo -e "${RED}バックアップが見つかりません${NC}"
      exit 1
    fi
    
    echo "バックアップから復元: backup/$latest_backup"
    cp -r backup/$latest_backup/* src/
    
    # プッシュ
    clasp push
    
    echo -e "${GREEN}✅ ロールバック完了${NC}"
    ;;
    
  *)
    echo -e "${RED}無効な選択です${NC}"
    exit 1
    ;;
esac

echo ""
echo "================================================"
echo "デプロイ作業完了"
echo ""
echo "📊 移行状況を確認するには:"
echo "  GASエディタで debugArchitectureStatus() を実行"
echo ""
echo "📝 ログを確認するには:"
echo "  表示 > ログ"
echo ""
echo "🔄 問題が発生した場合:"
echo "  ./deploy.sh で 4 を選択してロールバック"