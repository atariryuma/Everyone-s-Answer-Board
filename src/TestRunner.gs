/**
 * ⚠️ 緊急診断ツール: サービスアカウント・DB接続・権限確認
 * GASエディタで直接実行してください
 */

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
      console.log('❌ トークン生成エラー:', tokenError.message);
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
