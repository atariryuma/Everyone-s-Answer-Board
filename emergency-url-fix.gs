/**
 * 緊急URL修正システム
 * 管理パネルから直接実行してURL問題を解決
 * Web公開API対応
 */
function emergencyUrlFix() {
  try {
    console.log('🚨 緊急URL修正システム開始');
    
    // 1. 全キャッシュを完全削除
    console.log('\n1️⃣ キャッシュ完全削除:');
    const cache = CacheService.getScriptCache();
    const allKeys = [
      'WEB_APP_URL',
      'USER_URLS_',
      'CONFIG_JSON',
      'NORMALIZED_CONFIG'
    ];
    
    allKeys.forEach(key => {
      try {
        cache.remove(key);
        console.log(`✅ ${key} 削除完了`);
      } catch (e) {
        console.log(`⚠️ ${key} 削除失敗: ${e.message}`);
      }
    });
    
    // 統合キャッシュからも削除
    if (typeof cacheManager !== 'undefined') {
      allKeys.forEach(key => {
        try {
          cacheManager.remove(key);
          console.log(`✅ 統合キャッシュ ${key} 削除完了`);
        } catch (e) {
          console.log(`⚠️ 統合キャッシュ ${key} 削除失敗`);
        }
      });
    }
    
    // 2. 現在のScript ID情報を確認
    console.log('\n2️⃣ Script ID情報確認:');
    let primaryUrl = null;
    let scriptId = null;
    
    try {
      primaryUrl = ScriptApp.getService().getUrl();
      console.log('ScriptApp.getService().getUrl():', primaryUrl);
    } catch (e) {
      console.log('ScriptApp.getService().getUrl() エラー:', e.message);
    }
    
    try {
      scriptId = ScriptApp.getScriptId();
      console.log('ScriptApp.getScriptId():', scriptId);
    } catch (e) {
      console.log('ScriptApp.getScriptId() エラー:', e.message);
    }
    
    // 3. デプロイメント情報の確認
    console.log('\n3️⃣ デプロイメント情報:');
    try {
      const deployments = ScriptApp.getDeployments();
      console.log('デプロイメント数:', deployments.length);
      
      deployments.forEach((deployment, index) => {
        const config = deployment.getDeploymentConfig();
        console.log(`デプロイ ${index + 1}:`, {
          deploymentId: deployment.getDeploymentId(),
          description: config.description || 'No description',
          version: config.versionNumber || 'HEAD'
        });
      });
    } catch (e) {
      console.log('デプロイメント情報取得エラー:', e.message);
    }
    
    // 4. URL生成の直接テスト
    console.log('\n4️⃣ URL生成直接テスト:');
    
    // 新しいURL生成システムをテスト
    const newUrl = computeWebAppUrl();
    console.log('computeWebAppUrl() 結果:', newUrl);
    
    const testUserId = 'emergency-test-' + Date.now();
    const userUrls = generateUserUrls(testUserId);
    console.log('generateUserUrls() 結果:', userUrls);
    
    // 5. 強制的に正しいURLを設定（一時的）
    console.log('\n5️⃣ 強制URL設定:');
    
    // もし primaryUrl が正しいGoogle Workspace形式なら使用
    let correctUrl = null;
    if (primaryUrl && primaryUrl.includes('naha-okinawa.ed.jp')) {
      correctUrl = primaryUrl;
      console.log('✅ 正しいGoogle WorkspaceURL検出:', correctUrl);
    } else if (scriptId && scriptId.startsWith('AKfycby')) {
      // scriptIdが正しい形式の場合はフォールバック
      correctUrl = `https://script.google.com/macros/s/${scriptId}/exec`;
      console.log('✅ ScriptIDベースURL生成:', correctUrl);
    }
    
    if (correctUrl) {
      // 正しいURLをキャッシュに保存
      cache.put('WEB_APP_URL', correctUrl, 3600);
      if (typeof cacheManager !== 'undefined') {
        cacheManager.put('WEB_APP_URL', correctUrl, 3600);
      }
      console.log('✅ 正しいURLをキャッシュに保存');
      
      // テスト生成
      const correctedUrls = generateUserUrls(testUserId);
      console.log('修正後URL生成結果:', correctedUrls);
    }
    
    return {
      status: 'success',
      primaryUrl: primaryUrl,
      scriptId: scriptId,
      correctUrl: correctUrl,
      testUrls: userUrls,
      correctedUrls: correctUrl ? generateUserUrls(testUserId) : null,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('🚨 緊急修正エラー:', error.message);
    return {
      status: 'error',
      error: error.message,
      stack: error.stack,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * URL問題の診断専用関数
 */
function diagnoseUrlProblem() {
  try {
    console.log('🔍 URL問題診断開始');
    
    const diagnosis = {
      scriptAppService: null,
      scriptId: null,
      cachedUrl: null,
      generatedUrls: null,
      deploymentInfo: [],
      issues: [],
      recommendations: []
    };
    
    // ScriptApp API確認
    try {
      diagnosis.scriptAppService = ScriptApp.getService().getUrl();
    } catch (e) {
      diagnosis.issues.push('ScriptApp.getService().getUrl() 失敗: ' + e.message);
    }
    
    try {
      diagnosis.scriptId = ScriptApp.getScriptId();
    } catch (e) {
      diagnosis.issues.push('ScriptApp.getScriptId() 失敗: ' + e.message);
    }
    
    // キャッシュ確認
    try {
      const cache = CacheService.getScriptCache();
      diagnosis.cachedUrl = cache.get('WEB_APP_URL');
    } catch (e) {
      diagnosis.issues.push('キャッシュ確認失敗: ' + e.message);
    }
    
    // URL生成テスト
    try {
      diagnosis.generatedUrls = generateUserUrls('diagnosis-test');
    } catch (e) {
      diagnosis.issues.push('URL生成失敗: ' + e.message);
    }
    
    // 診断結果の分析
    if (diagnosis.scriptAppService && !diagnosis.scriptAppService.includes('naha-okinawa.ed.jp')) {
      diagnosis.issues.push('ScriptApp.getService().getUrl()が間違ったドメインを返している');
      diagnosis.recommendations.push('デプロイメント設定を確認する');
    }
    
    if (diagnosis.scriptId && !diagnosis.scriptId.startsWith('AKfycby')) {
      diagnosis.issues.push('ScriptIdが期待される形式でない');
    }
    
    if (diagnosis.cachedUrl && diagnosis.cachedUrl.includes('1SAAWR_m9TNMPsx')) {
      diagnosis.issues.push('間違ったURLがキャッシュされている');
      diagnosis.recommendations.push('キャッシュを完全にクリアする');
    }
    
    console.log('診断結果:', diagnosis);
    return diagnosis;
    
  } catch (error) {
    console.error('診断エラー:', error.message);
    return { status: 'error', error: error.message };
  }
}