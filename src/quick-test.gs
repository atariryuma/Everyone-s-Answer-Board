/**
 * 迅速な機能検証テスト
 * 最適化されたアーキテクチャの基本動作を確認
 */

function runQuickArchitectureTest() {
  console.log('🚀 迅速なアーキテクチャテストを開始...');
  console.log('='.repeat(50));
  
  var results = {
    passed: 0,
    failed: 0,
    tests: []
  };

  // Test 1: Constants check
  try {
    if (typeof SCRIPT_PROPS_KEYS !== 'undefined' && SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS) {
      console.log('✅ グローバル定数: 正常');
      results.passed++;
    } else {
      console.log('❌ グローバル定数: 定義されていません');
      results.failed++;
    }
    results.tests.push('GlobalConstants');
  } catch (e) {
    console.log('❌ グローバル定数テストエラー:', e.message);
    results.failed++;
    results.tests.push('GlobalConstants');
  }

  // Test 2: Cache Manager
  try {
    if (typeof cacheManager !== 'undefined' && typeof cacheManager.get === 'function') {
      console.log('✅ CacheManager: 正常');
      results.passed++;
    } else {
      console.log('❌ CacheManager: 利用できません');
      results.failed++;
    }
    results.tests.push('CacheManager');
  } catch (e) {
    console.log('❌ CacheManagerテストエラー:', e.message);
    results.failed++;
    results.tests.push('CacheManager');
  }

  // Test 3: Main entry point
  try {
    if (typeof doGet === 'function') {
      console.log('✅ doGet関数: 正常');
      results.passed++;
    } else {
      console.log('❌ doGet関数: 見つかりません');
      results.failed++;
    }
    results.tests.push('DoGetFunction');
  } catch (e) {
    console.log('❌ doGetテストエラー:', e.message);
    results.failed++;
    results.tests.push('DoGetFunction');
  }

  // Test 4: Database functions
  try {
    if (typeof getSheetsService === 'function' && 
        typeof findUserById === 'function' && 
        typeof updateUser === 'function' &&
        typeof createUser === 'function') {
      console.log('✅ データベース関数: 正常');
      results.passed++;
    } else {
      console.log('❌ データベース関数: 一部見つかりません');
      results.failed++;
    }
    results.tests.push('DatabaseFunctions');
  } catch (e) {
    console.log('❌ データベース関数テストエラー:', e.message);
    results.failed++;
    results.tests.push('DatabaseFunctions');
  }

  // Test 5: Auth functions
  try {
    if (typeof getServiceAccountTokenCached === 'function' && 
        typeof clearServiceAccountTokenCache === 'function') {
      console.log('✅ 認証関数: 正常');
      results.passed++;
    } else {
      console.log('❌ 認証関数: 見つかりません');
      results.failed++;
    }
    results.tests.push('AuthFunctions');
  } catch (e) {
    console.log('❌ 認証関数テストエラー:', e.message);
    results.failed++;
    results.tests.push('AuthFunctions');
  }

  // Test 6: URL functions
  try {
    if (typeof getWebAppUrlCached === 'function' && 
        typeof generateAppUrls === 'function') {
      console.log('✅ URL関数: 正常');
      results.passed++;
    } else {
      console.log('❌ URL関数: 見つかりません');
      results.failed++;
    }
    results.tests.push('UrlFunctions');
  } catch (e) {
    console.log('❌ URL関数テストエラー:', e.message);
    results.failed++;
    results.tests.push('UrlFunctions');
  }

  // Test 7: Core functions
  try {
    if (typeof onOpen === 'function' && 
        typeof auditLog === 'function' && 
        typeof refreshBoardData === 'function') {
      console.log('✅ コア関数: 正常');
      results.passed++;
    } else {
      console.log('❌ コア関数: 一部見つかりません');
      results.failed++;
    }
    results.tests.push('CoreFunctions');
  } catch (e) {
    console.log('❌ コア関数テストエラー:', e.message);
    results.failed++;
    results.tests.push('CoreFunctions');
  }

  // Test 8: Utility functions
  try {
    if (typeof isValidEmail === 'function' && 
        typeof getEmailDomain === 'function' && 
        typeof parseReactionString === 'function') {
      console.log('✅ ユーティリティ関数: 正常');
      results.passed++;
    } else {
      console.log('❌ ユーティリティ関数: 一部見つかりません');
      results.failed++;
    }
    results.tests.push('UtilityFunctions');
  } catch (e) {
    console.log('❌ ユーティリティ関数テストエラー:', e.message);
    results.failed++;
    results.tests.push('UtilityFunctions');
  }

  // Test 9: Config functions  
  try {
    if (typeof getConfig === 'function') {
      console.log('✅ 設定関数: 正常');
      results.passed++;
    } else {
      console.log('❌ 設定関数: 見つかりません');
      results.failed++;
    }
    results.tests.push('ConfigFunctions');
  } catch (e) {
    console.log('❌ 設定関数テストエラー:', e.message);
    results.failed++;
    results.tests.push('ConfigFunctions');
  }

  // Results
  var totalTests = results.passed + results.failed;
  var successRate = totalTests > 0 ? Math.round((results.passed / totalTests) * 100) : 100;
  
  console.log('\n' + '='.repeat(50));
  console.log('🏁 迅速テスト結果');
  console.log('='.repeat(50));
  console.log('📊 成功: ' + results.passed + '/' + totalTests + ' (' + successRate + '%)');
  
  if (successRate >= 90) {
    console.log('🎉 素晴らしい！アーキテクチャは正常に動作しています。');
  } else if (successRate >= 75) {
    console.log('✅ 良好！ほとんどの機能が動作していますが、一部要確認。');
  } else if (successRate >= 50) {
    console.log('⚠️ 注意！重要な問題があります。');
  } else {
    console.log('❌ 深刻！大幅な修正が必要です。');
  }

  console.log('\n📁 現在のファイル構造 (標準命名):');
  console.log('• main.gs - エントリーポイント & グローバル定数');
  console.log('• core.gs - ビジネスロジック');
  console.log('• database.gs - データベース操作');
  console.log('• auth.gs - 認証管理');
  console.log('• url.gs - URL管理');
  console.log('• cache.gs - キャッシュ管理');
  console.log('• config.gs - 設定管理');
  console.log('• monitor.gs - パフォーマンス監視');
  console.log('• optimizer.gs - 最適化機能');
  console.log('• test.gs - テストスイート');
  console.log('• stability.gs - 安定性向上');

  return {
    passed: results.passed,
    failed: results.failed,
    total: totalTests,
    successRate: successRate
  };
}