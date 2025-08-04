/**
 * URL修正システムのテスト関数
 * 管理パネルから実行してURL生成をテスト
 */
function testUrlFix() {
  try {
    console.log('=== URL修正システムテスト開始 ===');
    
    // 1. 現在のキャッシュ状況を確認
    console.log('\n1. 現在のキャッシュ状況:');
    const cache = CacheService.getScriptCache();
    const cachedUrl = cache.get('WEB_APP_URL');
    console.log('キャッシュされたURL:', cachedUrl);
    
    // 2. ScriptApp APIの直接確認
    console.log('\n2. ScriptApp API確認:');
    try {
      const primaryUrl = ScriptApp.getService().getUrl();
      console.log('ScriptApp.getService().getUrl():', primaryUrl);
    } catch (e) {
      console.log('ScriptApp.getService().getUrl() エラー:', e.message);
    }
    
    try {
      const scriptId = ScriptApp.getScriptId();
      console.log('ScriptApp.getScriptId():', scriptId);
    } catch (e) {
      console.log('ScriptApp.getScriptId() エラー:', e.message);
    }
    
    // 3. キャッシュクリアと新規生成
    console.log('\n3. キャッシュクリア実行:');
    const newUrl = clearUrlCache();
    console.log('クリア後の新規URL:', newUrl);
    
    // 4. URL生成テスト
    console.log('\n4. URL生成テスト:');
    const testUserId = 'test-user-' + Date.now();
    const userUrls = generateUserUrls(testUserId);
    console.log('生成されたURL群:', userUrls);
    
    // 5. URL検証
    console.log('\n5. URL検証:');
    if (userUrls.webAppUrl) {
      const isGoogleWorkspace = /^https:\/\/script\.google\.com\/a\/macros\/[^\/]+\/s\/[A-Za-z0-9_-]+\/exec$/.test(userUrls.webAppUrl);
      const isStandard = /^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(userUrls.webAppUrl);
      
      console.log('Google WorkspaceURL形式:', isGoogleWorkspace);
      console.log('標準URL形式:', isStandard);
      console.log('URLパターン適合:', isGoogleWorkspace || isStandard);
      
      // URLに含まれるScript IDを抽出
      const scriptIdMatch = userUrls.webAppUrl.match(/\/s\/([A-Za-z0-9_-]+)\/exec/);
      if (scriptIdMatch) {
        console.log('URLから抽出したScript ID:', scriptIdMatch[1]);
      }
    }
    
    console.log('\n=== テスト完了 ===');
    
    return {
      status: 'success',
      cachedUrl: cachedUrl,
      newUrl: newUrl,
      userUrls: userUrls,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('テスト実行エラー:', error.message);
    return {
      status: 'error',
      error: error.message,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * URL修正の強制実行
 * 問題が発生している場合の緊急対応
 */
function forceUrlSystemFix() {
  try {
    console.log('=== URL修正システム強制実行 ===');
    
    // 全キャッシュを強制クリア
    const cache = CacheService.getScriptCache();
    cache.removeAll(['WEB_APP_URL']);
    
    if (typeof cacheManager !== 'undefined') {
      cacheManager.remove('WEB_APP_URL');
    }
    
    console.log('✅ 全キャッシュ強制クリア完了');
    
    // URL生成を複数回試行
    let finalUrl = null;
    for (let i = 0; i < 3; i++) {
      console.log(`\n試行 ${i + 1}/3:`);
      const testUrl = computeWebAppUrl();
      console.log('生成URL:', testUrl);
      
      if (testUrl && (testUrl.includes('naha-okinawa.ed.jp') || /^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(testUrl))) {
        finalUrl = testUrl;
        console.log('✅ 有効なURL生成成功');
        break;
      }
      
      // 少し待機してから再試行
      Utilities.sleep(1000);
    }
    
    if (finalUrl) {
      console.log('最終確定URL:', finalUrl);
      return {
        status: 'success',
        fixedUrl: finalUrl,
        message: 'URL修正完了'
      };
    } else {
      return {
        status: 'error',
        message: '有効なURL生成に失敗しました'
      };
    }
    
  } catch (error) {
    console.error('強制修正エラー:', error.message);
    return {
      status: 'error',
      error: error.message
    };
  }
}