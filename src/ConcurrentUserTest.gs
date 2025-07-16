/**
 * アカウント作成競合テスト用関数
 * 複数ユーザーの同時作成をシミュレーション
 */
function testConcurrentUserCreation() {
  try {
    const testEmail = 'test+' + Date.now() + '@example.com';
    console.log('=== 競合テスト開始 ===', { testEmail });
    
    // 軽量ユーザー作成テスト
    const result1 = createLightweightUser(testEmail);
    console.log('軽量ユーザー作成結果:', result1);
    
    // 同じメールで再作成（重複チェック）
    const result2 = createLightweightUser(testEmail);
    console.log('重複作成結果:', result2);
    
    // 結果検証
    if (result1.userId === result2.userId && !result2.isNewUser) {
      console.log('✅ 競合テスト成功: 重複防止が正常に動作');
      return { success: true, message: '競合テスト成功' };
    } else {
      console.log('❌ 競合テスト失敗: 重複ユーザーが作成された');
      return { success: false, message: '重複ユーザー作成エラー' };
    }
    
  } catch (error) {
    console.error('競合テストエラー:', error);
    return { success: false, message: error.message };
  }
}

/**
 * パフォーマンステスト関数
 * 軽量ユーザー作成の速度を測定
 */
function testUserCreationPerformance() {
  try {
    const startTime = Date.now();
    const testEmail = 'perf+' + Date.now() + '@example.com';
    
    console.log('=== パフォーマンステスト開始 ===', { testEmail });
    
    // 軽量ユーザー作成
    const result = createLightweightUser(testEmail);
    
    const endTime = Date.now();
    const duration = endTime - startTime;
    
    console.log('✅ パフォーマンステスト完了', {
      duration: duration + 'ms',
      userId: result.userId,
      isNewUser: result.isNewUser
    });
    
    return {
      success: true,
      duration: duration,
      message: `軽量ユーザー作成時間: ${duration}ms`
    };
    
  } catch (error) {
    console.error('パフォーマンステストエラー:', error);
    return { success: false, message: error.message };
  }
}

/**
 * 段階的セットアップのパフォーマンステスト
 * 軽量作成 vs フル作成の時間比較
 */
function testStagedSetupPerformance() {
  try {
    const testEmail = 'staged+' + Date.now() + '@example.com';
    console.log('=== 段階的セットアップテスト開始 ===', { testEmail });
    
    // Phase 1: 軽量ユーザー作成
    const phase1Start = Date.now();
    const userResult = createLightweightUser(testEmail);
    const phase1End = Date.now();
    const phase1Duration = phase1End - phase1Start;
    
    console.log('Phase 1 完了 (軽量ユーザー作成):', phase1Duration + 'ms');
    
    // Phase 2: リソース作成
    const phase2Start = Date.now();
    const resourceResult = createUserResourcesAsync(userResult.userId);
    const phase2End = Date.now();
    const phase2Duration = phase2End - phase2Start;
    
    console.log('Phase 2 完了 (リソース作成):', phase2Duration + 'ms');
    
    const totalDuration = phase1Duration + phase2Duration;
    
    console.log('✅ 段階的セットアップテスト完了', {
      phase1: phase1Duration + 'ms',
      phase2: phase2Duration + 'ms',
      total: totalDuration + 'ms',
      userId: userResult.userId
    });
    
    return {
      success: true,
      phase1Duration: phase1Duration,
      phase2Duration: phase2Duration,
      totalDuration: totalDuration,
      message: `軽量作成: ${phase1Duration}ms, リソース作成: ${phase2Duration}ms, 合計: ${totalDuration}ms`
    };
    
  } catch (error) {
    console.error('段階的セットアップテストエラー:', error);
    return { success: false, message: error.message };
  }
}

/**
 * ロック競合回避の検証テスト
 * 複数の軽量ユーザー作成を並列実行（シミュレーション）
 */
function testLockContentionAvoidance() {
  try {
    console.log('=== ロック競合回避テスト開始 ===');
    
    const results = [];
    const testEmails = [];
    
    // 5つの異なるメールアドレスで軽量ユーザー作成
    for (let i = 0; i < 5; i++) {
      const testEmail = `locktest${i}+${Date.now()}@example.com`;
      testEmails.push(testEmail);
    }
    
    const startTime = Date.now();
    
    // 順次実行（並列実行はGASの制限で困難）
    for (const email of testEmails) {
      const userStart = Date.now();
      const userResult = createLightweightUser(email);
      const userEnd = Date.now();
      
      results.push({
        email: email,
        duration: userEnd - userStart,
        userId: userResult.userId,
        isNewUser: userResult.isNewUser
      });
    }
    
    const endTime = Date.now();
    const totalDuration = endTime - startTime;
    const avgDuration = totalDuration / 5;
    
    console.log('✅ ロック競合回避テスト完了', {
      totalDuration: totalDuration + 'ms',
      avgDuration: avgDuration + 'ms',
      results: results
    });
    
    return {
      success: true,
      totalDuration: totalDuration,
      avgDuration: avgDuration,
      results: results,
      message: `5ユーザー作成: 合計${totalDuration}ms, 平均${avgDuration}ms`
    };
    
  } catch (error) {
    console.error('ロック競合回避テストエラー:', error);
    return { success: false, message: error.message };
  }
}