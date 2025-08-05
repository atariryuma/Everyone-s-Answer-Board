/**
 * @fileoverview セットアップ関連のユーティリティ関数
 */

/**
 * アプリケーションの初期セットアップ（管理者が手動で実行） (マルチテナント対応版)
 * @param {string} credsJson サービスアカウント認証情報のJSON文字列
 * @param {string} dbId データベーススプレッドシートID（44文字）
 * @returns {{status:string, message:string}} セットアップ結果
 */
function setupApplication(credsJson, dbId) {
  try {
    JSON.parse(credsJson);
    if (typeof dbId !== 'string' || dbId.length !== 44) {
      throw new Error('無効なスプレッドシートIDです。IDは44文字の文字列である必要があります。');
    }

    var props = PropertiesService.getScriptProperties();
    props.setProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS, credsJson);
    props.setProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID, dbId);

    var adminEmail = Session.getEffectiveUser().getEmail();
    if (adminEmail) {
      props.setProperty(SCRIPT_PROPS_KEYS.ADMIN_EMAIL, adminEmail);
    }

    // データベースシートの初期化
    initializeDatabaseSheet(dbId);

    infoLog('✅ セットアップが正常に完了しました。');
    return { status: 'success', message: 'セットアップが正常に完了しました。' };
  } catch (e) {
    logError(e, 'customSetup', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM, { userId });
    throw new Error('セットアップに失敗しました: ' + e.message);
  }
}

/**
 * セットアップ状態をテストする (マルチテナント対応版)
 * @returns {{status:string, message:string, details?:Object}} テスト結果
 */
function testSetup() {
  try {
    var props = PropertiesService.getScriptProperties();
    var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    var creds = props.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS);

    if (!dbId) {
      return { status: 'error', message: 'データベーススプレッドシートIDが設定されていません。' };
    }

    if (!creds) {
      return { status: 'error', message: 'サービスアカウント認証情報が設定されていません。' };
    }

    // データベースへの接続テスト
    try {
      var userInfo = findUserByEmail(Session.getActiveUser().getEmail());
      return {
        status: 'success',
        message: 'セットアップは正常に完了しています。システムは使用準備が整いました。',
        details: {
          databaseConnected: true,
          userCount: userInfo ? 'ユーザー登録済み' : '未登録',
          serviceAccountConfigured: true,
        },
      };
    } catch (dbError) {
      return {
        status: 'warning',
        message: '設定は保存されていますが、データベースアクセスに問題があります。',
        details: { error: dbError.message },
      };
    }
  } catch (e) {
    logError(e, 'testCustomSetup', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM, { userId });
    return { status: 'error', message: 'セットアップテストに失敗しました: ' + e.message };
  }
}
