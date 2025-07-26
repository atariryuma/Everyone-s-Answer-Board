/**
 * リアクションシステムテスト実行スクリプト
 * ブラウザコンソールで実行するか、テストページから呼び出し
 */

// テスト結果を格納するオブジェクト
const testResults = {
  timestamp: new Date().toISOString(),
  environment: {
    userAgent: navigator.userAgent,
    url: window.location.href,
    studyQuestAppAvailable: !!window.studyQuestApp
  },
  tests: {},
  summary: {
    total: 0,
    passed: 0,
    failed: 0,
    errors: []
  }
};

// ログ関数
function logTest(category, name, status, data = null, error = null) {
  if (!testResults.tests[category]) {
    testResults.tests[category] = [];
  }
  
  const result = {
    name,
    status,
    timestamp: new Date().toISOString(),
    data,
    error: error ? error.message : null
  };
  
  testResults.tests[category].push(result);
  testResults.summary.total++;
  
  if (status === 'passed') {
    testResults.summary.passed++;
    console.log(`✅ [${category}] ${name}`, data);
  } else {
    testResults.summary.failed++;
    console.error(`❌ [${category}] ${name}`, error || data);
    if (error) testResults.summary.errors.push(error.message);
  }
}

// 基本環境チェック
async function checkEnvironment() {
  console.log('🔍 環境チェック開始...');
  
  try {
    // StudyQuestApp存在確認
    if (window.studyQuestApp) {
      logTest('environment', 'StudyQuestApp存在確認', 'passed', { available: true });
    } else {
      logTest('environment', 'StudyQuestApp存在確認', 'failed', { available: false });
      return false;
    }
    
    // 回答コンテナ確認
    const container = window.studyQuestApp.elements?.answersContainer;
    if (container) {
      logTest('environment', '回答コンテナ確認', 'passed', { element: !!container });
    } else {
      logTest('environment', '回答コンテナ確認', 'failed', { element: false });
      return false;
    }
    
    // 回答カード数確認
    const cards = container.querySelectorAll('.answer-card');
    logTest('environment', '回答カード数確認', cards.length > 0 ? 'passed' : 'failed', { count: cards.length });
    
    // テスト関数存在確認
    const testFunctions = {
      testReactionSystem: typeof window.testReactionSystem === 'function',
      testAllReactions: typeof window.testAllReactions === 'function',
      debugGetAnswerCardInfo: typeof window.debugGetAnswerCardInfo === 'function'
    };
    
    const allFunctionsExist = Object.values(testFunctions).every(exists => exists);
    logTest('environment', 'テスト関数存在確認', allFunctionsExist ? 'passed' : 'failed', testFunctions);
    
    return cards.length > 0 && allFunctionsExist;
    
  } catch (error) {
    logTest('environment', '環境チェック', 'failed', null, error);
    return false;
  }
}

// 基本機能テスト
async function runBasicTests() {
  console.log('🧪 基本機能テスト開始...');
  
  try {
    // 単一リアクションテスト
    const singleResult = await window.testReactionSystem({ 
      cardIndex: 0, 
      reaction: 'LIKE',
      verbose: false 
    });
    
    logTest('basic', '単一リアクションテスト', singleResult.success ? 'passed' : 'failed', {
      testCount: singleResult.tests?.length,
      errorCount: singleResult.errors?.length,
      details: singleResult
    });
    
    // 全リアクションタイプテスト
    const allReactionsResult = await window.testAllReactions(0);
    const allPassed = allReactionsResult.every(r => r.success);
    
    logTest('basic', '全リアクションタイプテスト', allPassed ? 'passed' : 'failed', {
      reactions: allReactionsResult.map(r => ({ reaction: r.reaction, success: r.success }))
    });
    
  } catch (error) {
    logTest('basic', '基本機能テスト', 'failed', null, error);
  }
}

// UI整合性テスト
async function runUITests() {
  console.log('🎨 UI整合性テスト開始...');
  
  try {
    const cardInfo = window.debugGetAnswerCardInfo(0);
    if (!cardInfo.success) {
      logTest('ui', 'カード情報取得', 'failed', cardInfo);
      return;
    }
    
    logTest('ui', 'カード情報取得', 'passed', {
      rowIndex: cardInfo.cardInfo.rowIndex,
      visible: cardInfo.cardInfo.computedStyle
    });
    
    // リアクションボタン状態確認
    const container = window.studyQuestApp.elements.answersContainer;
    const card = container.querySelectorAll('.answer-card')[0];
    const reactionButtons = card.querySelectorAll('.reaction-btn');
    
    const buttonStates = Array.from(reactionButtons).map(btn => ({
      reaction: btn.dataset.reaction,
      pressed: btn.getAttribute('aria-pressed'),
      disabled: btn.disabled,
      count: btn.querySelector('.count')?.textContent || '0',
      hasIcon: !!btn.querySelector('svg')
    }));
    
    const allButtonsValid = buttonStates.every(state => 
      state.reaction && 
      (state.pressed === 'true' || state.pressed === 'false') &&
      !isNaN(parseInt(state.count)) &&
      state.hasIcon
    );
    
    logTest('ui', 'リアクションボタン状態確認', allButtonsValid ? 'passed' : 'failed', buttonStates);
    
  } catch (error) {
    logTest('ui', 'UI整合性テスト', 'failed', null, error);
  }
}

// パフォーマンステスト
async function runPerformanceTests() {
  console.log('⚡ パフォーマンステスト開始...');
  
  try {
    // 連続クリックテスト
    const startTime = performance.now();
    const results = [];
    
    for (let i = 0; i < 3; i++) {
      const result = await window.testReactionSystem({ 
        cardIndex: 0, 
        reaction: 'LIKE', 
        verbose: false 
      });
      results.push(result.success);
      await new Promise(resolve => setTimeout(resolve, 200));
    }
    
    const endTime = performance.now();
    const allSucceeded = results.every(r => r);
    
    logTest('performance', '連続クリックテスト', allSucceeded ? 'passed' : 'failed', {
      iterations: results.length,
      successCount: results.filter(r => r).length,
      totalTime: endTime - startTime,
      averageTime: (endTime - startTime) / results.length
    });
    
  } catch (error) {
    logTest('performance', 'パフォーマンステスト', 'failed', null, error);
  }
}

// ローカルストレージテスト
async function runStorageTests() {
  console.log('💾 ローカルストレージテスト開始...');
  
  try {
    // ユーザーIDを取得
    const userId = window.studyQuestApp.state?.userId;
    if (!userId) {
      logTest('storage', 'ユーザーID取得', 'failed', { userId: null });
      return;
    }
    
    logTest('storage', 'ユーザーID取得', 'passed', { userId: userId.substring(0, 8) + '...' });
    
    // ローカルストレージ確認
    const storageKey = `reactions_${userId}`;
    const storedReactions = localStorage.getItem(storageKey);
    
    logTest('storage', 'ローカルストレージ確認', storedReactions ? 'passed' : 'failed', {
      hasData: !!storedReactions,
      dataLength: storedReactions?.length || 0
    });
    
    if (storedReactions) {
      try {
        const parsedData = JSON.parse(storedReactions);
        logTest('storage', 'ストレージデータ解析', 'passed', {
          entryCount: Object.keys(parsedData).length,
          sample: Object.keys(parsedData).slice(0, 3)
        });
      } catch (parseError) {
        logTest('storage', 'ストレージデータ解析', 'failed', null, parseError);
      }
    }
    
  } catch (error) {
    logTest('storage', 'ローカルストレージテスト', 'failed', null, error);
  }
}

// エラーハンドリングテスト
async function runErrorHandlingTests() {
  console.log('🔥 エラーハンドリングテスト開始...');
  
  try {
    // 存在しないカードでテスト
    const invalidCardResult = await window.testReactionSystem({ 
      cardIndex: 999, 
      reaction: 'LIKE',
      verbose: false 
    });
    
    logTest('error', '無効カードテスト', !invalidCardResult.success ? 'passed' : 'failed', {
      expectedFailure: true,
      actualSuccess: invalidCardResult.success,
      errors: invalidCardResult.errors
    });
    
  } catch (error) {
    logTest('error', 'エラーハンドリングテスト', 'passed', { caughtError: error.message });
  }
}

// メインテスト実行関数
async function runAllTests() {
  console.log('🚀 リアクションシステム総合テスト開始');
  console.log('=====================================');
  
  testResults.startTime = performance.now();
  
  // 環境チェック
  const envOk = await checkEnvironment();
  if (!envOk) {
    console.error('❌ 環境チェック失敗。テストを中止します。');
    return testResults;
  }
  
  // 各テストカテゴリを実行
  await runBasicTests();
  await runUITests();
  await runPerformanceTests();
  await runStorageTests();
  await runErrorHandlingTests();
  
  testResults.endTime = performance.now();
  testResults.duration = testResults.endTime - testResults.startTime;
  
  // 結果サマリー
  console.log('\n📊 テスト結果サマリー');
  console.log('=====================================');
  console.log(`🎯 総テスト数: ${testResults.summary.total}`);
  console.log(`✅ 成功: ${testResults.summary.passed}`);
  console.log(`❌ 失敗: ${testResults.summary.failed}`);
  console.log(`⏱️ 実行時間: ${testResults.duration.toFixed(2)}ms`);
  console.log(`📈 成功率: ${((testResults.summary.passed / testResults.summary.total) * 100).toFixed(1)}%`);
  
  if (testResults.summary.errors.length > 0) {
    console.log('\n🔍 エラー詳細:');
    testResults.summary.errors.forEach((error, index) => {
      console.log(`${index + 1}. ${error}`);
    });
  }
  
  // カテゴリ別結果
  console.log('\n📋 カテゴリ別結果:');
  Object.entries(testResults.tests).forEach(([category, tests]) => {
    const passed = tests.filter(t => t.status === 'passed').length;
    const total = tests.length;
    console.log(`${category}: ${passed}/${total} (${((passed/total)*100).toFixed(1)}%)`);
  });
  
  console.table(
    Object.entries(testResults.tests).flatMap(([category, tests]) =>
      tests.map(test => ({
        カテゴリ: category,
        テスト名: test.name,
        結果: test.status === 'passed' ? '✅' : '❌',
        エラー: test.error || '-'
      }))
    )
  );
  
  return testResults;
}

// グローバルに公開
window.runAllTests = runAllTests;
window.testResults = testResults;

// 自動実行オプション
if (typeof window.AUTO_RUN_TESTS !== 'undefined' && window.AUTO_RUN_TESTS) {
  setTimeout(() => {
    runAllTests();
  }, 2000);
}