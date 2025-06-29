/**
 * ===================================================================================
 * アプリケーション初期設定用スクリプト
 * ===================================================================================
 * このファイルには、管理者が最初の一度だけ実行するセットアップ関連の関数を格納します。
 */

/**
 * 【管理者が最初に一度だけ実行する関数】
 * 初期設定用のWebページへのURLをログに出力します。
 * この関数を実行後、ログに表示される指示に従ってください。
 */
function initialSetup() {
  const properties = PropertiesService.getScriptProperties();
  if (properties.getProperty('MAIN_DB_ID') && properties.getProperty('LOGGER_API_URL')) {
    Logger.log('✅ セットアップは既に完了しています。再度行う必要はありません。');
    try {
        const url = ScriptApp.getService().getUrl();
        if (url) {
            Logger.log('管理画面URL (デプロイ後): ' + url + '?mode=admin');
        } else {
            Logger.log('プロジェクトをデプロイすると、ここに管理画面URLが表示されます。');
        }
    } catch(e) {
        Logger.log('プロジェクトをデプロイすると、管理画面URLが利用可能になります。');
    }
    return;
  }
  
  Logger.log('セットアップを開始します。');
  Logger.log('まず、このプロジェクトを一度「ウェブアプリ」としてデプロイしてください。');
  Logger.log('デプロイ後、発行されたURLの末尾に「?setup=true」を付けてブラウザで開き、設定を続けてください。');
  Logger.log('例: https://script.google.com/macros/s/xxxx/exec?setup=true');
}

/**
 * セットアップ用のHTMLページを返す関数。doGetから呼び出される。
 * @returns {HtmlService.HtmlOutput} セットアップページのHTML
 */
function handleSetupRequest() {
  return HtmlService.createTemplateFromFile('SetupPage').evaluate()
    .setTitle('みんなの回答ボード - 初回セットアップ');
}

/**
 * API接続テスト用の関数
 */
function testApiConnection() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const apiUrl = properties.getProperty('LOGGER_API_URL');
    
    Logger.log('=== API接続テスト ===');
    Logger.log(`API URL: ${apiUrl}`);
    
    if (!apiUrl) {
      Logger.log('❌ API URLが設定されていません');
      return 'API URLが設定されていません';
    }
    
    // 簡単なping テストを実行
    const testResult = callDatabaseApi('ping', { test: true });
    
    Logger.log(`✅ API接続成功: ${JSON.stringify(testResult)}`);
    return `API接続成功: ${JSON.stringify(testResult)}`;
    
  } catch (e) {
    Logger.log(`❌ API接続失敗: ${e.message}`);
    return `API接続失敗: ${e.message}`;
  }
}

/**
 * デバッグ用: 現在の設定状況を確認する関数
 */
function debugCurrentSetup() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const currentUser = Session.getActiveUser().getEmail();
    
    Logger.log('=== デバッグ情報 ===');
    Logger.log(`現在のユーザー: ${currentUser}`);
    Logger.log(`MAIN_DB_ID: ${properties.getProperty('MAIN_DB_ID') || 'なし'}`);
    Logger.log(`LOGGER_API_URL: ${properties.getProperty('LOGGER_API_URL') || 'なし'}`);
    
    // データベースファイルの存在確認
    const dbId = properties.getProperty('MAIN_DB_ID');
    if (dbId) {
      try {
        const dbFile = DriveApp.getFileById(dbId);
        Logger.log(`データベースファイル名: ${dbFile.getName()}`);
        Logger.log(`データベースファイルURL: ${dbFile.getUrl()}`);
        
        // アクセス権限確認
        const editors = dbFile.getEditors();
        Logger.log(`編集者数: ${editors.length}`);
        editors.forEach((editor, index) => {
          Logger.log(`編集者${index + 1}: ${editor.getEmail()}`);
        });
        
      } catch (e) {
        Logger.log(`データベースファイルアクセスエラー: ${e.message}`);
      }
    }
    
    // 【ログデータベース】みんなの回答ボードの検索
    try {
      const files = DriveApp.getFilesByName('【ログデータベース】みんなの回答ボード');
      let fileCount = 0;
      while (files.hasNext()) {
        const file = files.next();
        fileCount++;
        Logger.log(`【ログデータベース】${fileCount}: ${file.getId()}`);
        Logger.log(`URL: ${file.getUrl()}`);
      }
      Logger.log(`【ログデータベース】みんなの回答ボード の総数: ${fileCount}`);
    } catch (e) {
      Logger.log(`【ログデータベース】検索エラー: ${e.message}`);
    }
    
    return 'デバッグ情報をログに出力しました';
  } catch (e) {
    Logger.log(`デバッグ関数エラー: ${e.message}`);
    return `デバッグエラー: ${e.message}`;
  }
}

/**
 * セットアップ用HTMLから呼び出され、設定を保存しDBを作成する関数
 * @param {string} apiUrl - 管理者向けログ記録APIのURL
 * @returns {string} 処理結果のメッセージ
 */
function saveSettingsAndCreateDb(apiUrl) {
  try {
    const properties = PropertiesService.getScriptProperties();
    
    // URL検証（Google Apps Script の複数の形式に対応）
    if (!apiUrl || typeof apiUrl !== 'string') {
      throw new Error('入力されたURLが無効です。正しい「管理者向けログ記録API」のウェブアプリURLを入力してください。');
    }
    
    // Google Apps Script の有効なURL形式を確認
    const validUrlPatterns = [
      /^https:\/\/script\.google\.com\/macros\/s\/[a-zA-Z0-9_-]+\/exec$/,  // 標準形式
      /^https:\/\/script\.google\.com\/a\/macros\/[^\/]+\/s\/[a-zA-Z0-9_-]+\/exec$/  // G Suite/Workspace カスタムドメイン形式
    ];
    
    const isValidUrl = validUrlPatterns.some(pattern => pattern.test(apiUrl));
    
    if (!isValidUrl) {
      throw new Error('入力されたURLが無効です。正しい「管理者向けログ記録API」のウェブアプリURLを入力してください。\n\n有効な形式:\n- https://script.google.com/macros/s/xxxxx/exec\n- https://script.google.com/a/macros/domain.com/s/xxxxx/exec');
    }
    properties.setProperty('LOGGER_API_URL', apiUrl);

    // データベースがなければ作成する（統合されたデータベースを使用）
    const mainDb = getOrCreateMainDatabase();
    
    // 重複する古いデータベースをクリーンアップ
    cleanupDuplicateDatabases();
    
    // セットアップ完了の確認とユーザーへの詳細情報提供
    let setupResult = '✅ 設定が正常に保存され、【ログデータベース】みんなの回答ボードに統合されました。\n\nセットアップが完了しました。\n\n';
    
    try {
      const properties = PropertiesService.getScriptProperties();
      const dbId = properties.getProperty('MAIN_DB_ID');
      
      if (dbId && mainDb) {
        const dbFile = DriveApp.getFileById(dbId);
        const dbUrl = dbFile.getUrl();
        
        setupResult += `📊 データベース情報:\n`;
        setupResult += `• ファイル名: ${dbFile.getName()}\n`;
        setupResult += `• URL: ${dbUrl}\n\n`;
        
        // 権限確認
        try {
          const editors = dbFile.getEditors();
          const currentUser = Session.getActiveUser().getEmail();
          const hasPermission = editors.some(editor => editor.getEmail() === currentUser);
          
          if (hasPermission) {
            setupResult += `✅ データベースアクセス権限: 確認済み\n`;
          } else {
            setupResult += `⚠️ データベースアクセス権限: 要確認\n`;
            setupResult += `上記URLから直接データベースにアクセスして編集権限を確認してください。\n`;
          }
        } catch (permError) {
          setupResult += `⚠️ 権限確認エラー: ${permError.message}\n`;
        }
        
        setupResult += `\n次のステップ:\n`;
        setupResult += `1. メインアプリで新規ユーザー登録をテストしてください\n`;
        setupResult += `2. 問題がある場合は上記データベースURLで権限を確認してください`;
        
      }
    } catch (infoError) {
      setupResult += `\n⚠️ 詳細情報取得エラー: ${infoError.message}`;
    }
    
    return setupResult;
  } catch(e) {
    Logger.log(e);
    return `❌ エラーが発生しました: ${e.message}`;
  }
}