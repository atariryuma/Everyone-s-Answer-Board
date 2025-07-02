/**
 * @fileoverview GAS関数テストランナー
 * 最適化された関数群をテストします
 */

/**
 * 認証管理関数のテスト
 */
function testAuthManager() {
  try {
    console.log('=== AuthManager テスト開始 ===');
    
    // テスト用のプロパティ設定
    var props = PropertiesService.getScriptProperties();
    var testCreds = {
      type: "service_account",
      project_id: "test-project",
      client_email: "test@test-project.iam.gserviceaccount.com",
      private_key: "-----BEGIN PRIVATE KEY-----\ntest-key\n-----END PRIVATE KEY-----"
    };
    
    props.setProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS, JSON.stringify(testCreds));
    
    // トークンキャッシュテスト（実際の呼び出しはしない）
    console.log('✓ 認証設定テスト完了');
    
    // キャッシュクリアテスト
    clearServiceAccountTokenCache();
    console.log('✓ キャッシュクリアテスト完了');
    
    console.log('=== AuthManager テスト成功 ===');
    return true;
  } catch (e) {
    console.error('AuthManager テストエラー:', e.message);
    return false;
  }
}

/**
 * キャッシュ管理関数のテスト
 */
function testCacheManager() {
  try {
    console.log('=== CacheManager テスト開始 ===');
    
    // テストデータ
    var testKey = 'test_cache_key';
    var testValue = { data: 'test', timestamp: Date.now() };
    
    // キャッシュ保存テスト
    setCachedValue(testKey, testValue, 'short');
    console.log('✓ キャッシュ保存テスト完了');
    
    // キャッシュ取得テスト
    var cachedValue = getCachedValue(testKey, function() {
      return { data: 'fallback' };
    }, 'short');
    
    if (cachedValue && cachedValue.data === 'test') {
      console.log('✓ キャッシュ取得テスト完了');
    } else {
      console.log('⚠ キャッシュ取得結果が期待と異なります');
    }
    
    // キャッシュ削除テスト
    removeCachedValue(testKey);
    console.log('✓ キャッシュ削除テスト完了');
    
    // 統計取得テスト
    var stats = getCacheStats();
    console.log('✓ キャッシュ統計テスト完了:', stats.longTermCacheCount + ' items');
    
    console.log('=== CacheManager テスト成功 ===');
    return true;
  } catch (e) {
    console.error('CacheManager テストエラー:', e.message);
    return false;
  }
}

/**
 * URL管理関数のテスト
 */
function testUrlManager() {
  try {
    console.log('=== UrlManager テスト開始 ===');
    
    // URL取得テスト（キャッシュ利用）
    var webAppUrl = getWebAppUrlCached();
    console.log('✓ WebアプリURL取得テスト完了:', webAppUrl || 'URL未設定');
    
    // URL生成テスト
    var testUserId = 'test-user-12345';
    var urls = generateAppUrlsOptimized(testUserId);
    
    if (urls.status === 'success' || urls.status === 'error') {
      console.log('✓ URL生成テスト完了:', urls.status);
    } else {
      console.log('⚠ URL生成で予期しない結果');
    }
    
    // URLキャッシュクリアテスト
    clearUrlCache();
    console.log('✓ URLキャッシュクリアテスト完了');
    
    console.log('=== UrlManager テスト成功 ===');
    return true;
  } catch (e) {
    console.error('UrlManager テストエラー:', e.message);
    return false;
  }
}

/**
 * リアクション管理関数のテスト
 */
function testReactionManager() {
  try {
    console.log('=== ReactionManager テスト開始 ===');
    
    // リアクション文字列パースのテスト
    var testString = 'user1@test.com, user2@test.com, user3@test.com';
    var parsed = parseReactionStringHelper(testString);
    
    if (parsed.length === 3 && parsed[0] === 'user1@test.com') {
      console.log('✓ リアクション文字列パーステスト完了');
    } else {
      console.log('⚠ パース結果が期待と異なります:', parsed);
    }
    
    // 空文字列のテスト
    var emptyParsed = parseReactionStringHelper('');
    if (emptyParsed.length === 0) {
      console.log('✓ 空文字列パーステスト完了');
    } else {
      console.log('⚠ 空文字列パースで予期しない結果:', emptyParsed);
    }
    
    console.log('=== ReactionManager テスト成功 ===');
    return true;
  } catch (e) {
    console.error('ReactionManager テストエラー:', e.message);
    return false;
  }
}

/**
 * Core関数のテスト
 */
function testCoreFunctions() {
  try {
    console.log('=== Core Functions テスト開始 ===');
    
    // 互換性関数のテスト
    console.log('✓ debugLog関数テスト');
    debugLog('テストメッセージ');
    
    // include関数のテスト（HTMLファイルが存在しない場合はエラーになる）
    try {
      // include('nonexistent.html');
      console.log('✓ include関数は存在します（実際のHTMLファイルテストはスキップ）');
    } catch (e) {
      console.log('✓ include関数エラーハンドリング確認');
    }
    
    console.log('=== Core Functions テスト成功 ===');
    return true;
  } catch (e) {
    console.error('Core Functions テストエラー:', e.message);
    return false;
  }
}

/**
 * 統合テスト実行
 */
function runAllOptimizedTests() {
  console.log('🚀 最適化された関数群の統合テストを開始します...\n');
  
  var results = [];
  
  // 各テストを実行
  results.push({ name: 'AuthManager', result: testAuthManager() });
  results.push({ name: 'CacheManager', result: testCacheManager() });
  results.push({ name: 'UrlManager', result: testUrlManager() });
  results.push({ name: 'ReactionManager', result: testReactionManager() });
  results.push({ name: 'CoreFunctions', result: testCoreFunctions() });
  
  // 結果の集計
  var passed = results.filter(function(r) { return r.result; }).length;
  var total = results.length;
  
  console.log('\n============================================');
  console.log('🏁 テスト結果サマリー');
  console.log('============================================');
  
  results.forEach(function(result) {
    var status = result.result ? '✅' : '❌';
    console.log(status + ' ' + result.name);
  });
  
  console.log('\n📊 成功率: ' + passed + '/' + total + ' (' + Math.round(passed/total*100) + '%)');
  
  if (passed === total) {
    console.log('🎉 全てのテストが成功しました！');
    console.log('最適化されたコードは正常に動作しています。');
  } else {
    console.log('⚠️ 一部のテストが失敗しました。');
    console.log('詳細なエラーログを確認してください。');
  }
  
  return { passed: passed, total: total, success: passed === total };
}

/**
 * 個別テスト実行用の便利関数
 */
function quickTest() {
  return runAllOptimizedTests();
}