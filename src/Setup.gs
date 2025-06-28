/**
 * @fileoverview StudyQuest -みんなの回答ボード- - 統合セットアップスクリプト
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
 * 💾 データベースID統合: USER_DATABASE_IDは廃止され、DATABASE_IDに統合されました。
 *    既存環境では後方互換性のためフォールバック読み込みを維持しています。
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
 * StudyQuestの統合セットアップ関数
 * フォルダ作成、データベース作成、ファイル移動、プロパティ設定を自動で行う
 * @param {string} [manualDeployId] - 手動で設定するデプロイID（任意）
 */
function studyQuestSetup(manualDeployId) {
  const FOLDER_NAME = "StudyQuest - みんなの回答ボード"; // アプリ専用フォルダ名
  const DB_FILENAME = "StudyQuest_UserDatabase";     // データベースファイル名
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // 30秒待機

  try {
    console.log("🚀 StudyQuest 統合セットアップを開始します...");

    // ステップ1: 専用フォルダの確認と作成
    let folder;
    const folders = DriveApp.getFoldersByName(FOLDER_NAME);
    if (folders.hasNext()) {
      folder = folders.next();
      console.log(`✅ 専用フォルダ「${FOLDER_NAME}」が既に存在します。`);
    } else {
      folder = DriveApp.createFolder(FOLDER_NAME);
      console.log(`✅ 専用フォルダ「${FOLDER_NAME}」を新規作成しました。`);
    }

    // ステップ2: データベース（スプレッドシート）の確認と作成
    let dbFile;
    const files = folder.getFilesByName(DB_FILENAME);
    if (files.hasNext()) {
      dbFile = files.next();
      console.log(`✅ データベースファイルがフォルダ内に既に存在します。`);
    } else {
      // マイドライブ直下も検索して、あれば移動させる
      const rootFiles = DriveApp.getRootFolder().getFilesByName(DB_FILENAME);
      if (rootFiles.hasNext()) {
        dbFile = rootFiles.next();
        dbFile.moveTo(folder);
        console.log(`✅ 既存のデータベースファイルを専用フォルダに移動しました。`);
      } else {
        // どこにもない場合のみ新規作成
        const newDb = SpreadsheetApp.create(DB_FILENAME);
        dbFile = DriveApp.getFileById(newDb.getId());
        dbFile.moveTo(folder);
        const sheet = newDb.getSheets()[0];
        sheet.setName("users");
        sheet.appendRow(["userId", "adminEmail", "spreadsheetId", "spreadsheetUrl", "createdAt", "accessToken", "configJson", "lastAccessedAt", "isActive"]);
        console.log(`✅ データベースファイルを新規作成し、専用フォルダに配置しました。`);
      }
    }
    const dbId = dbFile.getId();
    PropertiesService.getScriptProperties().setProperty("DATABASE_ID", dbId);
    console.log(`✅ データベースIDを設定しました: ${dbId}`);


    // ステップ3: デプロイIDとウェブアプリURLの設定
    const deployId = manualDeployId || ScriptApp.getDeploymentId();
    if (!deployId) {
      console.error("❌ DEPLOY_IDの取得に失敗しました。ウェブアプリとしてデプロイされているか確認してください。");
      throw new Error("DEPLOY_IDの取得に失敗しました。");
    }
    PropertiesService.getScriptProperties().setProperty("DEPLOY_ID", deployId);
    console.log(`✅ DEPLOY_IDを設定しました: ${deployId}`);

    const webAppUrl = `https://script.google.com/macros/s/${deployId}/exec`;
    PropertiesService.getScriptProperties().setProperty("WEB_APP_URL", webAppUrl);
    console.log(`✅ ウェブアプリURLを設定しました: ${webAppUrl}`);

    console.log("🎉 すべてのセットアップが正常に完了しました！");
    console.log("---");
    console.log("次のステップ:");
    console.log(`1. 管理画面にアクセスして動作を確認してください: ${webAppUrl}?mode=admin`);
    console.log(`2. 新規ユーザーとして登録テストを行ってください: ${webAppUrl}`);

  } catch (e) {
    console.error(`❌ セットアップ中にエラーが発生しました: ${e.message}`);
    console.error(e.stack);
    // 失敗した場合でも、取得できた情報はログに出力
    const props = PropertiesService.getScriptProperties().getProperties();
    console.log("現在のプロパティ設定:", props);
  } finally {
    lock.releaseLock();
  }
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
    hasDeployId: false
  };

  try {
    info.currentUrl = ScriptApp.getService().getUrl();
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
    // DATABASE_IDを優先、USER_DATABASE_IDは後方互換性のためのフォールバック
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
    // DATABASE_IDを標準として設定（USER_DATABASE_IDは廃止）
    props.setProperties({ 
      DATABASE_ID: spreadsheetId
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
    
    // プロダクションURLを保存
    props.setProperties({ WEB_APP_URL: productionUrl });

    return {
      success: true,
      message: 'アプリURLを設定しました',
      webAppUrl: productionUrl,
      productionUrl: productionUrl,
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

    // データベースIDテスト（DATABASE_ID優先、USER_DATABASE_IDは後方互換性）
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
  debugLog('\n🚀 === StudyQuest セットアップ結果 ===');
  
  if (results.status === 'success') {
    debugLog('✅ セットアップが正常に完了しました！');
  } else if (results.status === 'warning') {
    debugLog('⚠️ セットアップは完了しましたが、警告があります');
  } else {
    debugLog('❌ セットアップ中にエラーが発生しました');
  }

  debugLog(`\n📊 実行時刻: ${results.timestamp}`);
  
  if (Object.keys(results.urls).length > 0) {
    debugLog('\n🔗 生成されたURL:');
    for (const [key, url] of Object.entries(results.urls)) {
      debugLog(`   ${key}: ${url}`);
    }
  }

  if (results.errors.length > 0) {
    debugLog('\n❌ エラー:');
    results.errors.forEach(error => {
      debugLog(`   • ${error.step}: ${error.message}`);
    });
  }

  if (results.warnings.length > 0) {
    debugLog('\n⚠️ 警告:');
    results.warnings.forEach(warning => {
      debugLog(`   • ${warning.step}: ${warning.message}`);
    });
  }

  if (results.nextSteps.length > 0) {
    debugLog('\n📋 次のステップ:');
    results.nextSteps.forEach((step, index) => {
      debugLog(`   ${index + 1}. ${step.action}`);
      debugLog(`      ${step.instruction}`);
      if (step.description) {
        debugLog(`      ${step.description}`);
      }
    });
  }

  debugLog('\n🎉 StudyQuestへようこそ！');
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

/**
 * USER_DATABASE_IDからDATABASE_IDへの移行関数（将来のクリーンアップ用）
 * 既存環境でUSER_DATABASE_IDのみが設定されている場合、DATABASE_IDに移行する
 */
function migrateDatabaseIdProperty() {
  const props = PropertiesService.getScriptProperties();
  const databaseId = props.getProperty('DATABASE_ID');
  const userDatabaseId = props.getProperty('USER_DATABASE_ID');
  
  // DATABASE_IDが未設定でUSER_DATABASE_IDが設定されている場合のみ移行
  if (!databaseId && userDatabaseId) {
    props.setProperties({ DATABASE_ID: userDatabaseId });
    debugLog('Migrated USER_DATABASE_ID to DATABASE_ID:', userDatabaseId);
    
    // 移行後、古いプロパティを削除することも可能（コメントアウト）
    // props.deleteProperty('USER_DATABASE_ID');
    
    return {
      success: true,
      message: 'USER_DATABASE_IDからDATABASE_IDへ移行しました',
      migratedId: userDatabaseId
    };
  } else if (databaseId) {
    return {
      success: true,
      message: 'DATABASE_IDは既に設定済みです',
      currentId: databaseId
    };
  } else {
    return {
      success: false,
      message: '移行可能なデータベースIDが見つかりません'
    };
  }
}

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
    testConfiguration,
    migrateDatabaseIdProperty
  };
}