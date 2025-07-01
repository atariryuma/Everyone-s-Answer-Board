/**
 * 新アーキテクチャのすべてのテストを実行するテストランナー
 * Service Account Database Architecture Test Suite
 */

const { runTests: runBasicTests } = require('./serviceAccountDatabase.test.js');
const { runAdvancedTests } = require('./advancedDatabase.test.js');

async function runAllTests() {
  console.log('🚀 新アーキテクチャ総合テストスイートを開始します...\n');
  console.log('================================================================');
  console.log('                Service Account Database Tests                  ');
  console.log('================================================================\n');
  
  let totalPassed = 0;
  let totalTests = 0;
  
  try {
    // 基本テストの実行
    console.log('📋 Phase 1: 基本機能テスト');
    console.log('─'.repeat(60));
    
    const basicResult = runBasicTests();
    if (basicResult) {
      console.log('✅ 基本機能テスト: 全て成功');
      totalPassed += 7; // 基本テスト数
    } else {
      console.log('❌ 基本機能テスト: 一部失敗');
    }
    totalTests += 7;
    
    console.log('\n' + '─'.repeat(60));
    
    // 高度なテストの実行
    console.log('📋 Phase 2: 高度な統合テスト');
    console.log('─'.repeat(60));
    
    const advancedResult = runAdvancedTests();
    if (advancedResult) {
      console.log('✅ 高度な統合テスト: 全て成功');
      totalPassed += 6; // 高度なテスト数
    } else {
      console.log('❌ 高度な統合テスト: 一部失敗');
    }
    totalTests += 6;
    
    // 最終結果の表示
    console.log('\n' + '='.repeat(60));
    console.log('🏁 新アーキテクチャ総合テスト結果');
    console.log('='.repeat(60));
    console.log(`📊 成功率: ${totalPassed}/${totalTests} (${Math.round((totalPassed/totalTests)*100)}%)`);
    
    if (totalPassed === totalTests) {
      console.log('🎉 新アーキテクチャは完全に動作しています！');
      console.log('');
      console.log('✅ テスト済み機能:');
      console.log('   • サービスアカウント認証');
      console.log('   • Google Sheets API v4 統合');
      console.log('   • 中央データベース操作 (CRUD)');
      console.log('   • キャッシュ管理システム');
      console.log('   • データ処理とソート機能');
      console.log('   • 管理機能 (公開/非公開、表示モード)');
      console.log('   • リアクション・ハイライト機能');
      console.log('   • エラーハンドリングと復旧');
      console.log('   • パフォーマンス最適化');
      console.log('');
      console.log('🚀 本番環境での展開準備が完了しました！');
      console.log('');
      console.log('次のステップ:');
      console.log('1. サービスアカウントのJSONキーを作成');
      console.log('2. 中央データベースのスプレッドシートを作成');
      console.log('3. setupApplication(credsJson, dbId) を実行');
      console.log('4. アプリケーションをデプロイ');
      
      return true;
    } else {
      console.log('⚠️  一部のテストが失敗しました。');
      console.log('実装を確認してから本番環境に展開してください。');
      return false;
    }
    
  } catch (error) {
    console.error('❌ テスト実行中にエラーが発生しました:', error.message);
    return false;
  }
}

// テスト概要とアーキテクチャ情報の表示
function showArchitectureInfo() {
  console.log('📋 新アーキテクチャの概要');
  console.log('='.repeat(60));
  console.log('🏗️  アーキテクチャ: サービスアカウントモデル');
  console.log('🔐 認証方式: JWT + Google OAuth2');
  console.log('📊 データベース: Google Sheets API v4');
  console.log('👤 ユーザーファイル所有権: ユーザー自身');
  console.log('🌐 実行権限: USER_ACCESSING');
  console.log('');
  console.log('主な改善点:');
  console.log('• 403 Forbidden エラーの完全解決');
  console.log('• Google Workspace管理者設定不要');
  console.log('• 高速なキャッシュシステム');
  console.log('• 堅牢なエラーハンドリング');
  console.log('• スケーラブルなデータベース設計');
  console.log('');
}

// メイン実行
if (require.main === module) {
  showArchitectureInfo();
  runAllTests().then(success => {
    process.exit(success ? 0 : 1);
  }).catch(error => {
    console.error('予期しないエラー:', error);
    process.exit(1);
  });
}

module.exports = { runAllTests, showArchitectureInfo };