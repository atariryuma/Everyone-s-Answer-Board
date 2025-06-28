/**
 * @fileoverview StudyQuest - 統合セットアップスクリプト
 * デプロイユーザーが1回実行するだけで環境が整う統合セットアップ機能
 * 
 * =================================================================
 * 📚 使用方法（デプロイ管理者向け）
 * =================================================================
 * 
 * 1. 【初回セットアップ】
 *    - Google Apps Scriptプロジェクトを開く
 *    - 「デプロイ」→「新しいデプロイ」でウェブアプリとして公開
 *    - 公開後、Apps Scriptエディタで以下を実行：
 *      studyQuestSetup()
 * 
 * 2. 【DEPLOY_ID手動指定が必要な場合】
 *    - 自動抽出に失敗した場合、ウェブアプリURLからDEPLOY_IDを取得
 *    - 例: https://script.google.com/macros/s/AKfycbxExample123/exec
 *    - 以下を実行：
 *      studyQuestSetup("AKfycbxExample123")
 * 
 * 3. 【セットアップ確認】
 *    - 実行後、ログに表示される次のステップに従う
 *    - 管理パネル、新規登録のテストを実行
 * 
 * 🔄 既存のsetup()関数は非推奨になりました。
 *    このstudyQuestSetup()が全機能を統合しています。
 * 
 * 📖 詳細な使用方法は SETUP_README.md を参照してください。
 * 
 * =================================================================
 */

// =================================================================
// 定数定義（Code.gsの定数を使用）
// =================================================================

// USER_DB_CONFIG は Code.gs で定義済み

// =================================================================
// メインセットアップ関数
// =================================================================

/**
 * StudyQuest統合セットアップ - これ1つを実行するだけで環境が整います
 * 
 * 使用方法:
 * 1. ウェブアプリとして公開
 * 2. この関数を実行（deployIdは自動抽出を試行、失敗時は手動指定）
 * 
 * @param {string} deployId - オプション：手動でのデプロイID指定
 * @return {Object} セットアップ結果とユーザー向けの詳細情報
 */
function studyQuestSetup(deployId = null) {
  console.log('=== StudyQuest統合セットアップ開始 ===');
  
  const results = {
    timestamp: new Date().toISOString(),
    status: 'success',
    steps: [],
    errors: [],
    warnings: [],
    urls: {},
    nextSteps: []
  };

  try {
    // ステップ1: 現在の環境情報取得
    addStep(results, '環境情報取得', '現在のスクリプト情報を取得中...');
    const envInfo = getEnvironmentInfo();
    addStep(results, '環境情報取得', '完了', envInfo);

    // ステップ2: DEPLOY_ID設定
    addStep(results, 'DEPLOY_ID設定', 'デプロイIDを設定中...');
    const deployResult = setupDeployId(deployId, envInfo.currentUrl);
    if (deployResult.success) {
      addStep(results, 'DEPLOY_ID設定', '完了', deployResult);
      results.urls.deployId = deployResult.deployId;
    } else {
      addError(results, 'DEPLOY_ID設定', deployResult.message);
      results.nextSteps.push({
        action: 'DEPLOY_ID手動設定',
        instruction: 'studyQuestSetup("YOUR_DEPLOY_ID_HERE")を実行してください',
        reference: deployResult.extractionHelp
      });
    }

    // ステップ3: ユーザーデータベース作成
    addStep(results, 'データベース作成', 'ユーザーデータベースを作成中...');
    const dbResult = setupUserDatabase();
    if (dbResult.success) {
      addStep(results, 'データベース作成', '完了', dbResult);
      results.urls.userDatabase = dbResult.spreadsheetUrl;
    } else {
      addError(results, 'データベース作成', dbResult.message);
    }

    // ステップ4: URL設定初期化
    addStep(results, 'URL設定', 'ウェブアプリURL設定を初期化中...');
    const urlResult = initializeAppUrls();
    if (urlResult.success) {
      addStep(results, 'URL設定', '完了', urlResult);
      results.urls.webApp = urlResult.webAppUrl;
      results.urls.production = urlResult.productionUrl;
      results.urls.development = urlResult.developmentUrl;
    } else {
      addWarning(results, 'URL設定', urlResult.message);
    }

    // ステップ5: 設定テスト
    addStep(results, '設定テスト', '全体設定をテスト中...');
    const testResult = testConfiguration();
    if (testResult.success) {
      addStep(results, '設定テスト', '完了', testResult);
    } else {
      addWarning(results, '設定テスト', testResult.message);
    }

    // 成功時の次のステップ
    if (results.errors.length === 0) {
      results.nextSteps.push(
        {
          action: '管理パネルアクセス',
          instruction: `${results.urls.webApp}?mode=admin にアクセス`,
          description: '管理画面でスプレッドシートを追加してください'
        },
        {
          action: '新規登録テスト',
          instruction: `${results.urls.webApp} にアクセス`,
          description: '「新規登録」をクリックしてテストしてください'
        },
        {
          action: 'ユーザーマニュアル',
          instruction: 'README.mdまたはドキュメントを確認',
          description: 'ユーザー向けの使用方法を確認してください'
        }
      );
    }

  } catch (error) {
    console.error('セットアップ中に予期しないエラー:', error);
    addError(results, '全体', `予期しないエラー: ${error.message}`);
    results.status = 'error';
  }

  // 最終結果判定
  if (results.errors.length > 0) {
    results.status = 'error';
  } else if (results.warnings.length > 0) {
    results.status = 'warning';
  }

  console.log('=== StudyQuest統合セットアップ完了 ===');
  console.log('結果:', JSON.stringify(results, null, 2));
  
  // ユーザー向けの見やすい結果表示
  displaySetupResults(results);
  
  return results;
}

// =================================================================
// セットアップサブ関数
// =================================================================

/**
 * 現在の環境情報を取得
 */
function getEnvironmentInfo() {
  const info = {
    scriptId: Session.getTemporaryActiveUserKey() || 'unknown',
    userEmail: Session.getActiveUser().getEmail(),
    timezone: Session.getScriptTimeZone(),
    currentUrl: '',
    isDevelopment: false,
    hasDeployId: false
  };

  try {
    info.currentUrl = ScriptApp.getService().getUrl();
    info.isDevelopment = /\/dev(\?|$)/.test(info.currentUrl);
  } catch (e) {
    console.warn('現在のURL取得に失敗:', e.message);
  }

  const props = PropertiesService.getScriptProperties();
  info.hasDeployId = !!props.getProperty('DEPLOY_ID');

  return info;
}

/**
 * DEPLOY_ID設定
 */
function setupDeployId(manualDeployId, currentUrl) {
  const props = PropertiesService.getScriptProperties();
  
  // 手動指定された場合
  if (manualDeployId) {
    if (!/^[a-zA-Z0-9_-]{10,}$/.test(manualDeployId)) {
      return {
        success: false,
        message: '無効なDEPLOY_ID形式です',
        deployId: manualDeployId
      };
    }
    
    props.setProperties({ DEPLOY_ID: manualDeployId });
    return {
      success: true,
      message: '手動でDEPLOY_IDを設定しました',
      deployId: manualDeployId,
      method: 'manual'
    };
  }

  // 自動抽出を試行
  if (currentUrl) {
    const extractedId = extractDeployIdFromUrl(currentUrl);
    if (extractedId) {
      props.setProperties({ DEPLOY_ID: extractedId });
      return {
        success: true,
        message: 'URLからDEPLOY_IDを自動抽出しました',
        deployId: extractedId,
        method: 'auto',
        sourceUrl: currentUrl
      };
    }
  }

  // 既存のDEPLOY_IDをチェック
  const existing = props.getProperty('DEPLOY_ID');
  if (existing) {
    return {
      success: true,
      message: '既存のDEPLOY_IDを使用',
      deployId: existing,
      method: 'existing'
    };
  }

  return {
    success: false,
    message: 'DEPLOY_IDの自動抽出に失敗しました',
    extractionHelp: {
      instruction: 'ウェブアプリのURLからDEPLOY_IDを取得してください',
      example: 'https://script.google.com/macros/s/YOUR_DEPLOY_ID/exec',
      currentUrl: currentUrl || 'URL取得不可'
    }
  };
}

/**
 * ユーザーデータベースセットアップ
 */
function setupUserDatabase() {
  try {
    const props = PropertiesService.getScriptProperties();
    // 既存のプロパティ名に合わせる（DATABASE_ID）
    const existingDbId = props.getProperty('DATABASE_ID') || props.getProperty('USER_DATABASE_ID');
    
    if (existingDbId) {
      try {
        const existingDb = SpreadsheetApp.openById(existingDbId);
        return {
          success: true,
          message: '既存のユーザーデータベースを使用',
          spreadsheetId: existingDbId,
          spreadsheetUrl: existingDb.getUrl(),
          method: 'existing'
        };
      } catch (e) {
        console.warn('既存データベースにアクセス不可、新規作成します:', e.message);
      }
    }

    // 新規データベース作成
    const spreadsheet = SpreadsheetApp.create('StudyQuest_UserDatabase');
    const sheet = spreadsheet.getActiveSheet();
    sheet.setName(USER_DB_CONFIG.SHEET_NAME);
    sheet.appendRow(USER_DB_CONFIG.HEADERS);
    
    const spreadsheetId = spreadsheet.getId();
    // 既存のプロパティ名に合わせる
    props.setProperties({ 
      DATABASE_ID: spreadsheetId,
      USER_DATABASE_ID: spreadsheetId  // 後方互換性
    });
    
    return {
      success: true,
      message: '新しいユーザーデータベースを作成しました',
      spreadsheetId: spreadsheetId,
      spreadsheetUrl: spreadsheet.getUrl(),
      method: 'created'
    };

  } catch (error) {
    return {
      success: false,
      message: `ユーザーデータベース作成エラー: ${error.message}`,
      error: error
    };
  }
}

/**
 * アプリURL初期化
 */
function initializeAppUrls() {
  try {
    const props = PropertiesService.getScriptProperties();
    const deployId = props.getProperty('DEPLOY_ID');
    
    if (!deployId) {
      return {
        success: false,
        message: 'DEPLOY_IDが設定されていません'
      };
    }

    let currentUrl = '';
    try {
      currentUrl = ScriptApp.getService().getUrl();
    } catch (e) {
      console.warn('現在のURL取得に失敗:', e.message);
    }

    const productionUrl = `https://script.google.com/macros/s/${deployId}/exec`;
    const developmentUrl = currentUrl && /\/dev(\?|$)/.test(currentUrl) ? currentUrl : `https://script.google.com/macros/s/${deployId}/dev`;
    
    // プロダクションURLを保存
    props.setProperties({ WEB_APP_URL: productionUrl });

    return {
      success: true,
      message: 'アプリURLを設定しました',
      webAppUrl: productionUrl,
      productionUrl: productionUrl,
      developmentUrl: developmentUrl,
      currentUrl: currentUrl
    };

  } catch (error) {
    return {
      success: false,
      message: `URL初期化エラー: ${error.message}`,
      error: error
    };
  }
}

/**
 * 設定全体のテスト
 */
function testConfiguration() {
  try {
    const props = PropertiesService.getScriptProperties();
    const tests = [];

    // DEPLOY_IDテスト
    const deployId = props.getProperty('DEPLOY_ID');
    tests.push({
      name: 'DEPLOY_ID',
      status: deployId ? 'OK' : 'NG',
      value: deployId || 'なし'
    });

    // データベースIDテスト
    const dbId = props.getProperty('DATABASE_ID') || props.getProperty('USER_DATABASE_ID');
    tests.push({
      name: 'ユーザーデータベース',
      status: dbId ? 'OK' : 'NG',
      value: dbId || 'なし'
    });

    // ウェブアプリURLテスト
    const webAppUrl = props.getProperty('WEB_APP_URL');
    tests.push({
      name: 'ウェブアプリURL',
      status: webAppUrl ? 'OK' : 'NG',
      value: webAppUrl || 'なし'
    });

    // 権限テスト
    let permissionTest = 'OK';
    try {
      Session.getActiveUser().getEmail();
    } catch (e) {
      permissionTest = 'NG - 権限不足';
    }
    tests.push({
      name: '実行権限',
      status: permissionTest,
      value: permissionTest === 'OK' ? '正常' : 'エラー'
    });

    const allPassed = tests.every(test => test.status === 'OK');

    return {
      success: allPassed,
      message: allPassed ? '全テスト通過' : '一部テスト失敗',
      tests: tests
    };

  } catch (error) {
    return {
      success: false,
      message: `設定テストエラー: ${error.message}`,
      error: error
    };
  }
}

// =================================================================
// 結果表示・ヘルパー関数
// =================================================================

/**
 * セットアップ結果をユーザーフレンドリーに表示
 */
function displaySetupResults(results) {
  console.log('\n🚀 === StudyQuest セットアップ結果 ===');
  
  if (results.status === 'success') {
    console.log('✅ セットアップが正常に完了しました！');
  } else if (results.status === 'warning') {
    console.log('⚠️ セットアップは完了しましたが、警告があります');
  } else {
    console.log('❌ セットアップ中にエラーが発生しました');
  }

  console.log(`\n📊 実行時刻: ${results.timestamp}`);
  
  if (Object.keys(results.urls).length > 0) {
    console.log('\n🔗 生成されたURL:');
    for (const [key, url] of Object.entries(results.urls)) {
      console.log(`   ${key}: ${url}`);
    }
  }

  if (results.errors.length > 0) {
    console.log('\n❌ エラー:');
    results.errors.forEach(error => {
      console.log(`   • ${error.step}: ${error.message}`);
    });
  }

  if (results.warnings.length > 0) {
    console.log('\n⚠️ 警告:');
    results.warnings.forEach(warning => {
      console.log(`   • ${warning.step}: ${warning.message}`);
    });
  }

  if (results.nextSteps.length > 0) {
    console.log('\n📋 次のステップ:');
    results.nextSteps.forEach((step, index) => {
      console.log(`   ${index + 1}. ${step.action}`);
      console.log(`      ${step.instruction}`);
      if (step.description) {
        console.log(`      ${step.description}`);
      }
    });
  }

  console.log('\n🎉 StudyQuestへようこそ！');
}

/**
 * ステップ追加ヘルパー
 */
function addStep(results, stepName, status, details = null) {
  results.steps.push({
    step: stepName,
    status: status,
    details: details,
    timestamp: new Date().toISOString()
  });
}

/**
 * エラー追加ヘルパー
 */
function addError(results, stepName, message) {
  results.errors.push({
    step: stepName,
    message: message,
    timestamp: new Date().toISOString()
  });
}

/**
 * 警告追加ヘルパー
 */
function addWarning(results, stepName, message) {
  results.warnings.push({
    step: stepName,
    message: message,
    timestamp: new Date().toISOString()
  });
}

// =================================================================
// 個別セットアップ関数（既存関数の再利用）
// =================================================================

// DEPLOY_IDを抽出する関数はCode.gsのextractDeployIdFromUrlを使用

// =================================================================
// エクスポート（他のファイルから使用する場合）
// =================================================================

if (typeof module !== 'undefined') {
  module.exports = {
    studyQuestSetup,
    getEnvironmentInfo,
    setupDeployId,
    setupUserDatabase,
    initializeAppUrls,
    testConfiguration
  };
}