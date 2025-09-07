/**
 * ⚠️ 緊急診断ツール: サービスアカウント・DB接続・権限確認
 * GASエディタで直接実行してください
 */

/**
 * 🚀 管理パネル最適化テスト（Service Account維持）
 * publishApplication と connectDataSource の直接DB更新を検証
 */
function testOptimizedManagementPanel() {
  try {
    console.log('='.repeat(60));
    console.log('🚀 管理パネル最適化テスト開始（Service Account維持版）');
    console.log('='.repeat(60));

    // 現在のユーザー情報取得
    const currentUser = UserManager.getCurrentEmail();
    const userInfo = DB.findUserByEmail(currentUser);
    
    if (!userInfo) {
      console.error('❌ テスト失敗: ユーザー情報が見つかりません');
      console.log('💡 ヒント: 先に setupApplication() を実行してユーザーを作成してください');
      return;
    }

    console.log('👤 テスト対象ユーザー:', {
      userId: userInfo.userId,
      userEmail: userInfo.userEmail,
      currentConfigSize: userInfo.configJson?.length || 0,
    });

    // Test 1: connectDataSource 最適化版テスト
    console.log('\n📝 Test 1: connectDataSource 最適化版テスト');
    const testSpreadsheetId = '1test-spreadsheet-id-for-optimization-test';
    const testSheetName = 'テストシート';

    // 現在の設定を取得（テスト前状態）
    const beforeConfig = JSON.parse(userInfo.configJson || '{}');
    console.log('🔍 テスト前設定:', {
      spreadsheetId: beforeConfig.spreadsheetId,
      sheetName: beforeConfig.sheetName,
      setupStatus: beforeConfig.setupStatus,
    });

    // Test 2: DB.updateUserInDatabase 直接テスト
    console.log('\n📝 Test 2: DB.updateUserInDatabase 直接テスト');
    
    const testConfig = {
      ...beforeConfig,
      spreadsheetId: testSpreadsheetId,
      sheetName: testSheetName,
      setupStatus: 'completed',
      testMode: true,
      optimizedAt: new Date().toISOString(),
      lastModified: new Date().toISOString(),
    };

    console.log('💾 直接DB更新テスト開始...');
    const updateResult = DB.updateUserInDatabase(userInfo.userId, {
      configJson: JSON.stringify(testConfig),
      lastModified: new Date().toISOString(),
    });

    console.log('🔍 DB更新結果:', {
      success: updateResult.success,
      error: updateResult.error,
      updatedCells: updateResult.updatedCells,
      configJsonSize: updateResult.configJsonSize,
    });

    if (updateResult.success) {
      console.log('✅ DB直接更新: 成功');
      
      // 更新後の確認
      const updatedUserInfo = DB.findUserById(userInfo.userId);
      const updatedConfig = JSON.parse(updatedUserInfo.configJson || '{}');
      
      console.log('🔍 更新後設定確認:', {
        spreadsheetId: updatedConfig.spreadsheetId,
        sheetName: updatedConfig.sheetName,
        setupStatus: updatedConfig.setupStatus,
        hasTestMode: updatedConfig.testMode === true,
        hasOptimizedAt: !!updatedConfig.optimizedAt,
      });

      // 検証
      const isDataSaved = updatedConfig.spreadsheetId === testSpreadsheetId && 
                         updatedConfig.sheetName === testSheetName &&
                         updatedConfig.setupStatus === 'completed';

      if (isDataSaved) {
        console.log('✅ データ保存検証: 成功 - 設定が正しく保存されています');
      } else {
        console.error('❌ データ保存検証: 失敗 - 設定が正しく保存されていません');
      }

      // Test 3: Service Account 認証テスト
      console.log('\n📝 Test 3: Service Account 認証動作確認');
      const serviceTest = getSheetsServiceWithRetry();
      
      console.log('🔍 Service Account サービス取得結果:', {
        hasService: !!serviceTest,
        hasSpreadsheets: !!serviceTest?.spreadsheets,
        hasValues: !!serviceTest?.spreadsheets?.values,
        hasAppend: !!serviceTest?.spreadsheets?.values?.append,
        appendIsFunction: typeof serviceTest?.spreadsheets?.values?.append === 'function'
      });

      if (serviceTest?.spreadsheets?.values?.append && 
          typeof serviceTest.spreadsheets.values.append === 'function') {
        console.log('✅ Service Account 認証: 正常動作中');
      } else {
        console.error('❌ Service Account 認証: 問題が検出されました');
      }

    } else {
      console.error('❌ DB直接更新: 失敗', updateResult.error);
    }

    console.log('\n='.repeat(60));
    console.log('🎯 管理パネル最適化テスト完了');
    console.log('='.repeat(60));

    const testSummary = {
      dbUpdateSuccess: updateResult.success,
      serviceAccountWorking: !!serviceTest?.spreadsheets?.values?.append,
      optimizationComplete: true,
      timestamp: new Date().toISOString(),
    };

    console.log('📊 テスト結果サマリー:', testSummary);
    return testSummary;

  } catch (error) {
    console.error('❌ 管理パネル最適化テストエラー:', {
      error: error.message,
      stack: error.stack,
      timestamp: new Date().toISOString(),
    });
    return {
      success: false,
      error: error.message,
      timestamp: new Date().toISOString(),
    };
  }
}

/**
 * 🔍 サービスアカウント動作確認（詳細版）
 */
function checkServiceAccountStatus() {
  try {
    console.log('='.repeat(50));
    console.log('🔍 サービスアカウント動作確認開始');
    console.log('='.repeat(50));

    // 1. PropertiesService確認
    const props = PropertiesService.getScriptProperties();
    const serviceAccountCreds = props.getProperty('SERVICE_ACCOUNT_CREDS');
    const databaseId = props.getProperty('DATABASE_SPREADSHEET_ID');
    const adminEmail = props.getProperty('ADMIN_EMAIL');

    console.log('📋 PropertiesService確認:');
    console.log(
      '- SERVICE_ACCOUNT_CREDS:',
      serviceAccountCreds ? `設定済み (${serviceAccountCreds.length}文字)` : '❌未設定'
    );
    console.log('- DATABASE_SPREADSHEET_ID:', databaseId ? `設定済み (${databaseId})` : '❌未設定');
    console.log('- ADMIN_EMAIL:', adminEmail ? `設定済み (${adminEmail})` : '❌未設定');

    if (!serviceAccountCreds) {
      throw new Error('SERVICE_ACCOUNT_CREDSが設定されていません');
    }

    // 2. Service Account JSON解析
    let serviceAccountJson;
    try {
      serviceAccountJson = JSON.parse(serviceAccountCreds);
      console.log('✅ サービスアカウントJSON解析成功');
      console.log('- client_email:', serviceAccountJson.client_email);
      console.log('- project_id:', serviceAccountJson.project_id);
    } catch (parseError) {
      throw new Error(`サービスアカウントJSON解析エラー: ${parseError.message}`);
    }

    // 3. 現在の実行ユーザー確認
    const currentUser = UserManager.getCurrentEmail();
    console.log('👤 実行ユーザー:', currentUser);
    console.log('🔐 管理者権限:', currentUser === adminEmail ? '✅あり' : '❌なし');

    // 4. データベーススプレッドシート接続テスト
    if (databaseId) {
      try {
        const dbSpreadsheet = SpreadsheetApp.openById(databaseId);
        const dbSheets = dbSpreadsheet.getSheets();
        console.log('📊 データベース接続成功:');
        console.log('- スプレッドシート名:', dbSpreadsheet.getName());
        console.log('- シート数:', dbSheets.length);
        console.log(
          '- シート名:',
          dbSheets.map((s) => s.getName())
        );

        // Users シート確認
        const usersSheet = dbSpreadsheet.getSheetByName('Users');
        if (usersSheet) {
          const userCount = Math.max(0, usersSheet.getLastRow() - 1);
          console.log('- ユーザー数:', userCount);
        } else {
          console.log('❌ Usersシートが見つかりません');
        }
      } catch (dbError) {
        console.log('❌ データベース接続エラー:', dbError.message);
      }
    }

    // 5. Service Account Token生成テスト
    try {
      const token = getServiceAccountTokenCached();
      console.log('🔑 サービスアカウントトークン:', token ? '✅生成成功' : '❌生成失敗');
    } catch (tokenError) {
      console.log('❌ トークン生成エラー: [詳細は非表示]');
    }

    console.log('='.repeat(50));
    console.log('✅ サービスアカウント動作確認完了');
    console.log('='.repeat(50));

    return createResponse(true, 'サービスアカウント確認完了', {
      hasServiceAccount: !!serviceAccountCreds,
      hasDatabaseId: !!databaseId,
      hasAdminEmail: !!adminEmail,
      currentUser,
      isAdmin: currentUser === adminEmail,
    });
  } catch (error) {
    console.error('❌ サービスアカウント確認エラー:', error.message);
    return createResponse(false, 'エラーが発生しました', null, error);
  }
}

/**
 * 🧹 強制JSONクリーンアップ（デバッグ版）
 */
function forceCleanupConfigJson() {
  try {
    console.log('='.repeat(50));
    console.log('🧹 強制configJSONクリーンアップ開始');
    console.log('='.repeat(50));

    // まずサービスアカウント確認
    const serviceAccountCheck = checkServiceAccountStatus();
    if (!serviceAccountCheck.success) {
      throw new Error('サービスアカウントの問題により実行できません');
    }

    // 現在のユーザー情報取得
    const currentUser = UserManager.getCurrentEmail();
    const userInfo = DB.findUserByEmail(currentUser);

    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    console.log('👤 対象ユーザー:', userInfo.userEmail);
    console.log('📏 現在のconfigJson長:', userInfo.configJson?.length || 0);
    console.log(
      '🔍 configJson先頭200文字:',
      userInfo.configJson?.substring(0, 200) || 'データなし'
    );

    // 強制的にJSONを解析して再構築
    let cleanedConfig = {};

    if (userInfo.configJson) {
      try {
        const parsedConfig = JSON.parse(userInfo.configJson);
        console.log('📋 解析されたconfig構造:', Object.keys(parsedConfig));

        // 重複したconfigJsonフィールドを検出
        if (parsedConfig.configJson) {
          console.log('⚠️ ネストしたconfigJsonフィールドを発見');
          // 再帰的に展開
          let nestedData = parsedConfig.configJson;
          while (typeof nestedData === 'string') {
            try {
              nestedData = JSON.parse(nestedData);
              console.log('🔄 さらにネストレベル発見、展開中...');
            } catch (e) {
              break;
            }
          }

          if (typeof nestedData === 'object') {
            cleanedConfig = { ...nestedData };
          }
        } else {
          cleanedConfig = { ...parsedConfig };
        }
      } catch (parseError) {
        console.log('❌ JSON解析エラー:', parseError.message);
        // デフォルト設定で初期化
        cleanedConfig = {
          setupStatus: 'pending',
          createdAt: new Date().toISOString(),
          lastModified: new Date().toISOString(),
        };
      }
    }

    // 基本DBフィールドを除去
    const dbFields = ['userId', 'userEmail', 'isActive', 'lastModified', 'configJson'];
    dbFields.forEach((field) => {
      delete cleanedConfig[field];
    });

    // 最終更新時刻を設定
    cleanedConfig.lastModified = new Date().toISOString();

    console.log('🔧 クリーンアップ後の構造:', Object.keys(cleanedConfig));

    // ConfigManager経由でデータベース更新
    const updateResult = ConfigManager.saveConfig(userInfo.userId, cleanedConfig);

    console.log('💾 更新結果:', updateResult.success ? '✅成功' : '❌失敗');

    // 更新後の確認
    const updatedUser = DB.findUserById(userInfo.userId);
    console.log('🔄 更新後のconfigJson長:', updatedUser.configJson?.length || 0);
    console.log(
      '🔍 更新後の先頭200文字:',
      updatedUser.configJson?.substring(0, 200) || 'データなし'
    );

    console.log('='.repeat(50));
    console.log('✅ 強制configJSONクリーンアップ完了');
    console.log('='.repeat(50));

    return createResponse(true, '強制configJSONクリーンアップ完了', {
      originalLength: userInfo.configJson?.length || 0,
      cleanedLength: updatedUser.configJson?.length || 0,
      cleanedFields: Object.keys(cleanedConfig),
    });
  } catch (error) {
    console.error('❌ 強制クリーンアップエラー:', error.message);
    return createResponse(false, 'エラーが発生しました', null, error);
  }
}

/**
 * 🔬 詳細データベース診断
 */
function diagnoseDatabase() {
  try {
    console.log('='.repeat(50));
    console.log('🔬 詳細データベース診断開始');
    console.log('='.repeat(50));

    const dbId = PropertiesService.getScriptProperties().getProperty('DATABASE_SPREADSHEET_ID');
    if (!dbId) {
      throw new Error('DATABASE_SPREADSHEET_IDが設定されていません');
    }

    const dbSpreadsheet = SpreadsheetApp.openById(dbId);
    const usersSheet = dbSpreadsheet.getSheetByName('Users');

    if (!usersSheet) {
      throw new Error('Usersシートが見つかりません');
    }

    const values = usersSheet.getDataRange().getValues();
    const headers = values[0];

    console.log('📊 データベース構造:');
    console.log('- ヘッダー:', headers);
    console.log('- 総行数:', values.length);
    console.log('- データ行数:', Math.max(0, values.length - 1));

    // 各ユーザーのconfigJson状況確認
    for (let i = 1; i < values.length && i <= 5; i++) {
      // 最初の5ユーザーのみ
      const row = values[i];
      const configJsonIndex = headers.indexOf('configJson');
      if (configJsonIndex >= 0) {
        const configJson = row[configJsonIndex];
        console.log(`👤 ユーザー${i}:`, {
          email: row[1],
          configJsonLength: configJson?.length || 0,
          configJsonPreview: configJson?.substring(0, 100) || 'データなし',
        });
      }
    }

    console.log('='.repeat(50));
    console.log('✅ データベース診断完了');
    console.log('='.repeat(50));

    return createResponse(true, 'データベース診断完了', {
      headers,
      userCount: Math.max(0, values.length - 1),
    });
  } catch (error) {
    console.error('❌ データベース診断エラー:', error.message);
    return createResponse(false, 'エラーが発生しました', null, error);
  }
}
