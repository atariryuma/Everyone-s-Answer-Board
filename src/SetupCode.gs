/**
 * ===================================================================================
 * アプリケーション初期設定用スクリプト (サービスアカウントモデル対応)
 * ===================================================================================
 * このファイルには、管理者が最初の一度だけ実行するセットアップ関連の関数を格納します。
 * 新アーキテクチャでは、サービスアカウント認証情報と中央データベースIDの設定が必要です。
 */

/**
 * 【管理者が最初に一度だけ実行する関数】
 * セットアップの状況を確認し、必要に応じて初期設定の手順を表示します。
 */
function initialSetup() {
  var properties = PropertiesService.getScriptProperties();
  var serviceAccountCreds = properties.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS);
  var databaseId = properties.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
  
  if (serviceAccountCreds && databaseId) {
    Logger.log('✅ セットアップは既に完了しています。');
    try {
      var url = ScriptApp.getService().getUrl();
      if (url) {
        Logger.log('アプリケーションURL: ' + url);
        Logger.log('新規登録ページ: ' + url);
      } else {
        Logger.log('プロジェクトをデプロイすると、ここにアプリケーションURLが表示されます。');
      }
    } catch(e) {
      Logger.log('プロジェクトをデプロイすると、アプリケーションURLが利用可能になります。');
    }
    return;
  }
  
  Logger.log('🚀 新アーキテクチャのセットアップが必要です。');
  Logger.log('');
  Logger.log('セットアップ手順:');
  Logger.log('1. Google Cloud Platform (GCP) でサービスアカウントを作成');
  Logger.log('2. サービスアカウントのJSONキーをダウンロード');
  Logger.log('3. 中央データベース用のスプレッドシートを作成');
  Logger.log('4. サービスアカウントを中央データベースの編集者として追加');
  Logger.log('5. setupApplication(credsJson, dbId) 関数を実行');
  Logger.log('');
  Logger.log('詳細な手順については、要件定義書を参照してください。');
}

/**
 * セットアップ完了後のテスト関数
 * サービスアカウント認証とデータベースアクセスをテストします。
 */
function testSetup() {
  try {
    var properties = PropertiesService.getScriptProperties();
    var serviceAccountCreds = properties.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS);
    var databaseId = properties.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    
    if (!serviceAccountCreds || !databaseId) {
      throw new Error('セットアップが完了していません。setupApplication() を実行してください。');
    }
    
    Logger.log('=== セットアップテスト ===');
    
    // サービスアカウント認証テスト
    Logger.log('1. サービスアカウント認証テスト...');
    var token = getServiceAccountToken();
    if (token) {
      Logger.log('✅ サービスアカウント認証: 成功');
    } else {
      throw new Error('サービスアカウント認証に失敗');
    }
    
    // データベースアクセステスト
    Logger.log('2. データベースアクセステスト...');
    var service = getSheetsService();
    var spreadsheet = service.spreadsheets.get(databaseId);
    if (spreadsheet) {
      Logger.log('✅ データベースアクセス: 成功');
      Logger.log('データベース名: ' + spreadsheet.properties.title);
    } else {
      throw new Error('データベースアクセスに失敗');
    }
    
    // テストユーザー検索
    Logger.log('3. テストクエリ実行...');
    var testUser = findUserByEmail('test@example.com'); // 存在しないユーザー
    Logger.log('✅ テストクエリ: 成功 (結果: ' + (testUser ? 'あり' : 'なし') + ')');
    
    Logger.log('');
    Logger.log('🎉 すべてのテストが成功しました！アプリケーションは正常に動作する準備ができています。');
    
    return {
      status: 'success',
      message: 'セットアップテストが成功しました',
      databaseName: spreadsheet.properties.title
    };
    
  } catch (e) {
    Logger.log('❌ テスト失敗: ' + e.message);
    return {
      status: 'error',
      message: 'セットアップテストに失敗しました: ' + e.message
    };
  }
}

/**
 * セットアップ情報を表示する
 */
function showSetupInfo() {
  var properties = PropertiesService.getScriptProperties();
  var serviceAccountCreds = properties.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS);
  var databaseId = properties.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
  
  Logger.log('=== 現在のセットアップ状況 ===');
  Logger.log('サービスアカウント設定: ' + (serviceAccountCreds ? '✅ あり' : '❌ なし'));
  Logger.log('データベースID設定: ' + (databaseId ? '✅ あり' : '❌ なし'));
  
  if (serviceAccountCreds) {
    try {
      var creds = JSON.parse(serviceAccountCreds);
      Logger.log('サービスアカウントメール: ' + creds.client_email);
      Logger.log('プロジェクトID: ' + creds.project_id);
    } catch (e) {
      Logger.log('⚠️ サービスアカウント情報の解析に失敗: ' + e.message);
    }
  }
  
  if (databaseId) {
    Logger.log('データベースID: ' + databaseId);
    try {
      var service = getSheetsService();
      var spreadsheet = service.spreadsheets.get(databaseId);
      Logger.log('データベース名: ' + spreadsheet.properties.title);
      Logger.log('データベースURL: https://docs.google.com/spreadsheets/d/' + databaseId);
    } catch (e) {
      Logger.log('⚠️ データベース情報の取得に失敗: ' + e.message);
    }
  }
  
  var url = ScriptApp.getService().getUrl();
  if (url) {
    Logger.log('アプリケーションURL: ' + url);
  }
}

/**
 * データベースの状況を確認する
 */
function checkDatabaseStatus() {
  try {
    var properties = PropertiesService.getScriptProperties();
    var databaseId = properties.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    
    if (!databaseId) {
      Logger.log('❌ データベースIDが設定されていません');
      return { status: 'error', message: 'データベースIDが設定されていません' };
    }
    
    var service = getSheetsService();
    var sheetName = DB_SHEET_CONFIG.SHEET_NAME;
    
    // データベースの基本情報
    var spreadsheet = service.spreadsheets.get(databaseId);
    Logger.log('=== データベース状況 ===');
    Logger.log('名前: ' + spreadsheet.properties.title);
    Logger.log('シート数: ' + spreadsheet.sheets.length);
    
    // Usersシートの存在確認
    var usersSheet = spreadsheet.sheets.find(function(s) { 
      return s.properties.title === sheetName; 
    });
    
    if (!usersSheet) {
      Logger.log('⚠️ Usersシートが見つかりません');
      return { status: 'warning', message: 'Usersシートがありません' };
    }
    
    Logger.log('✅ Usersシート: 存在確認');
    
    // データの確認
    var data = service.spreadsheets.values.get(databaseId, sheetName + '!A:H').values || [];
    Logger.log('データ行数: ' + data.length + ' (ヘッダー含む)');
    
    if (data.length > 0) {
      Logger.log('ヘッダー: ' + JSON.stringify(data[0]));
      Logger.log('ユーザー数: ' + (data.length - 1));
    }
    
    return {
      status: 'success',
      message: 'データベースは正常です',
      userCount: data.length > 0 ? data.length - 1 : 0
    };
    
  } catch (e) {
    Logger.log('❌ データベース確認エラー: ' + e.message);
    return { status: 'error', message: 'データベース確認に失敗: ' + e.message };
  }
}

/**
 * 新規ユーザー登録のテスト（実際には登録しない）
 */
function testUserRegistration() {
  try {
    var testEmail = Session.getActiveUser().getEmail();
    Logger.log('=== ユーザー登録テスト ===');
    Logger.log('テスト対象メール: ' + testEmail);
    
    // 既存ユーザーチェックのテスト
    var existingUser = findUserByEmail(testEmail);
    if (existingUser) {
      Logger.log('✅ 既存ユーザー検索: 成功 (ユーザーが見つかりました)');
      Logger.log('既存ユーザーID: ' + existingUser.userId);
      Logger.log('作成日時: ' + existingUser.createdAt);
    } else {
      Logger.log('✅ 既存ユーザー検索: 成功 (新規ユーザーです)');
    }
    
    // サービスアカウント情報の確認
    var properties = PropertiesService.getScriptProperties();
    var serviceAccountCreds = JSON.parse(properties.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS));
    Logger.log('サービスアカウントメール: ' + serviceAccountCreds.client_email);
    
    Logger.log('🎉 ユーザー登録テストが成功しました（実際の登録は行われていません）');
    
    return {
      status: 'success',
      message: 'ユーザー登録テストが成功しました',
      isExistingUser: !!existingUser
    };
    
  } catch (e) {
    Logger.log('❌ ユーザー登録テストエラー: ' + e.message);
    return { status: 'error', message: 'ユーザー登録テストに失敗: ' + e.message };
  }
}

/**
 * 全体的な診断を実行する
 */
function runDiagnostics() {
  Logger.log('🔍 アプリケーション診断を開始します...');
  Logger.log('');
  
  var results = {
    setup: testSetup(),
    database: checkDatabaseStatus(),
    userRegistration: testUserRegistration()
  };
  
  Logger.log('');
  Logger.log('=== 診断結果サマリー ===');
  Logger.log('セットアップ: ' + (results.setup.status === 'success' ? '✅' : '❌'));
  Logger.log('データベース: ' + (results.database.status === 'success' ? '✅' : (results.database.status === 'warning' ? '⚠️' : '❌')));
  Logger.log('ユーザー登録: ' + (results.userRegistration.status === 'success' ? '✅' : '❌'));
  
  var allGood = results.setup.status === 'success' && 
                results.database.status === 'success' && 
                results.userRegistration.status === 'success';
  
  if (allGood) {
    Logger.log('');
    Logger.log('🎉 すべての診断項目が正常です！アプリケーションは本番環境で使用できます。');
  } else {
    Logger.log('');
    Logger.log('⚠️ 一部の診断項目で問題が検出されました。詳細を確認してください。');
  }
  
  return results;
}