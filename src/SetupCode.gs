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
 * 強化版API接続テスト - 詳細な診断機能付き
 */
function testApiConnection() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const apiUrl = properties.getProperty('LOGGER_API_URL');
    
    Logger.log('=== 詳細API接続テスト ===');
    Logger.log(`API URL: ${apiUrl}`);
    Logger.log(`現在のユーザー: ${Session.getActiveUser().getEmail()}`);
    Logger.log(`実効ユーザー: ${Session.getEffectiveUser().getEmail()}`);
    
    if (!apiUrl) {
      Logger.log('❌ API URLが設定されていません');
      return 'API URLが設定されていません';
    }
    
    // URL形式の検証
    const urlValidation = validateApiUrl(apiUrl);
    if (!urlValidation.isValid) {
      Logger.log(`❌ URL形式エラー: ${urlValidation.error}`);
      return `URL形式エラー: ${urlValidation.error}`;
    }
    
    // 基本的な接続テスト
    Logger.log('--- 基本接続テスト ---');
    try {
      const basicTestOptions = {
        method: 'get',
        muteHttpExceptions: true,
        headers: {
          'User-Agent': 'StudyQuest-MainApp-Test/1.0'
        }
      };
      
      const basicResponse = UrlFetchApp.fetch(apiUrl, basicTestOptions);
      const basicResponseCode = basicResponse.getResponseCode();
      const basicResponseText = basicResponse.getContentText();
      
      Logger.log(`基本テスト応答コード: ${basicResponseCode}`);
      Logger.log(`基本テスト応答内容: ${basicResponseText.substring(0, 200)}${basicResponseText.length > 200 ? '...' : ''}`);
      
      if (basicResponseCode === 200) {
        Logger.log('✅ 基本接続: 成功');
      } else {
        Logger.log(`⚠️ 基本接続: 問題あり (${basicResponseCode})`);
      }
      
    } catch (basicError) {
      Logger.log(`❌ 基本接続テスト失敗: ${basicError.message}`);
    }
    
    // API呼び出しテスト
    Logger.log('--- API呼び出しテスト ---');
    const testResult = callDatabaseApi('ping', { test: true });
    
    Logger.log(`✅ API呼び出し成功: ${JSON.stringify(testResult)}`);
    return `API接続テスト完了 - 基本接続とAPI呼び出しの両方が成功しました`;
    
  } catch (e) {
    Logger.log(`❌ API接続テスト失敗: ${e.message}`);
    Logger.log(`エラー詳細: ${e.stack || 'スタックトレースなし'}`);
    return `API接続テスト失敗: ${e.message}`;
  }
}

/**
 * Logger API URLの形式を検証する
 * @param {string} url - 検証するURL
 * @returns {Object} {isValid: boolean, error?: string}
 */
function validateApiUrl(url) {
  if (!url || typeof url !== 'string') {
    return { isValid: false, error: 'URLが空または無効です' };
  }
  
  // Google Apps Script の有効なURL形式をチェック
  const validPatterns = [
    /^https:\/\/script\.google\.com\/macros\/s\/[a-zA-Z0-9_-]+\/exec$/,  // 標準形式
    /^https:\/\/script\.google\.com\/a\/macros\/[^\/]+\/s\/[a-zA-Z0-9_-]+\/exec$/  // Workspace カスタムドメイン形式
  ];
  
  const isValidFormat = validPatterns.some(pattern => pattern.test(url));
  
  if (!isValidFormat) {
    return { 
      isValid: false, 
      error: 'Google Apps Script Web AppのURL形式ではありません。正しい形式: https://script.google.com/macros/s/xxxxx/exec または https://script.google.com/a/macros/domain.com/s/xxxxx/exec' 
    };
  }
  
  return { isValid: true };
}

/**
 * 新しいAdmin Logger API URLを自動設定（DOMAIN対応版）
 */
function setNewLoggerApiUrl() {
  const newApiUrl = 'https://script.google.com/macros/s/AKfycbwH55G3O92Gqqj5VNLbmiBBKl7cbZ8DtKh4g2IhFt-iw4lXMEyr5um2q9SvP61kU2XZ/exec';
  return updateLoggerApiUrl(newApiUrl);
}

/**
 * Logger APIの詳細診断 - デプロイ状況とアクセス権限をチェック
 */
function diagnoseLoggerApi() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const apiUrl = properties.getProperty('LOGGER_API_URL');
    
    Logger.log('=== Logger API 詳細診断 ===');
    
    if (!apiUrl) {
      Logger.log('❌ Logger API URLが未設定');
      return 'Logger API URLが未設定';
    }
    
    Logger.log(`API URL: ${apiUrl}`);
    
    // URL形式チェック
    const urlCheck = validateApiUrl(apiUrl);
    if (!urlCheck.isValid) {
      Logger.log(`❌ URL形式問題: ${urlCheck.error}`);
      return `URL形式問題: ${urlCheck.error}`;
    }
    Logger.log('✅ URL形式: 正常');
    
    // 段階的接続テスト
    const testResults = [];
    
    // 1. HTTP GETテスト
    try {
      const getResponse = UrlFetchApp.fetch(apiUrl, { 
        method: 'get', 
        muteHttpExceptions: true 
      });
      const getCode = getResponse.getResponseCode();
      const getText = getResponse.getContentText();
      
      testResults.push(`GET テスト: ${getCode} - ${getText.substring(0, 100)}`);
      Logger.log(`GET テスト結果: ${getCode}`);
      
      if (getText.includes('Page Not Found') || getText.includes('Not Found')) {
        Logger.log('⚠️ APIが「Page Not Found」を返しています - デプロイ状況を確認してください');
      }
      
    } catch (getError) {
      testResults.push(`GET テスト失敗: ${getError.message}`);
      Logger.log(`GET テスト失敗: ${getError.message}`);
    }
    
    // 2. HTTP POSTテスト（最小構成）
    try {
      const postResponse = UrlFetchApp.fetch(apiUrl, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ action: 'test', data: {} }),
        muteHttpExceptions: true
      });
      const postCode = postResponse.getResponseCode();
      const postText = postResponse.getContentText();
      
      testResults.push(`POST テスト: ${postCode} - ${postText.substring(0, 100)}`);
      Logger.log(`POST テスト結果: ${postCode}`);
      
    } catch (postError) {
      testResults.push(`POST テスト失敗: ${postError.message}`);
      Logger.log(`POST テスト失敗: ${postError.message}`);
    }
    
    // 診断結果のまとめ
    Logger.log('--- 診断結果まとめ ---');
    testResults.forEach((result, index) => {
      Logger.log(`${index + 1}. ${result}`);
    });
    
    // 推奨アクション
    Logger.log('--- 推奨アクション ---');
    Logger.log('1. Logger APIプロジェクトが正しくデプロイされているか確認');
    Logger.log('2. Logger APIが「ウェブアプリ」として公開されているか確認');
    Logger.log('3. Logger APIの実行設定が「USER_DEPLOYING」になっているか確認');
    Logger.log('4. Logger APIのアクセス設定が「DOMAIN」になっているか確認');
    
    return `診断完了 - 詳細はログを確認してください`;
    
  } catch (e) {
    Logger.log(`診断エラー: ${e.message}`);
    return `診断エラー: ${e.message}`;
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
    
    // 強化されたURL検証を使用
    const urlValidation = validateApiUrl(apiUrl);
    if (!urlValidation.isValid) {
      throw new Error(`URL検証エラー: ${urlValidation.error}`);
    }
    
    // URLを保存
    properties.setProperty('LOGGER_API_URL', apiUrl);

    
    // Logger APIの接続テストを実行
    Logger.log('=== Logger API 接続確認 ===');
    try {
      // 基本的な接続テスト
      const testResponse = UrlFetchApp.fetch(apiUrl, {
        method: 'get',
        muteHttpExceptions: true,
        headers: {
          'User-Agent': 'StudyQuest-Setup/1.0'
        }
      });
      
      const testCode = testResponse.getResponseCode();
      const testText = testResponse.getContentText();
      
      Logger.log(`接続テスト結果: ${testCode}`);
      Logger.log(`レスポンス: ${testText.substring(0, 200)}`);
      
      if (testCode === 404 || testText.includes('Page Not Found')) {
        Logger.log('⚠️ Logger APIが404を返しています - デプロイ状況を確認してください');
        // 404でも設定は保存するが、警告を表示
      } else if (testCode === 200) {
        Logger.log('✅ Logger APIへの基本接続が成功しました');
      } else {
        Logger.log(`⚠️ Logger APIが予期しないレスポンスを返しました: ${testCode}`);
      }
      
    } catch (connectError) {
      Logger.log(`接続テストエラー: ${connectError.message}`);
      // 接続エラーでも設定は保存するが、ログに記録
    }

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
        
        
        // 権限確認
        try {
          const editors = dbFile.getEditors();
          const currentUser = Session.getActiveUser().getEmail();
          const hasPermission = editors.some(editor => editor.getEmail() === currentUser);
          
          if (hasPermission) {
            setupResult += `✅ データベースアクセス権限: 確認済み\n`;
          } else {
            setupResult += `⚠️ データベースアクセス権限: 不足している可能性があります。
`;
          }
        } catch (permError) {
          setupResult += `⚠️ 権限確認エラー: ${permError.message}\n`;
        }
        
        setupResult += `\n次のステップ:\n`;
        setupResult += `1. メインアプリで新規ユーザー登録をテストしてください\n`;
        setupResult += `2. 問題がある場合は、Apps Scriptプロジェクトの「プロジェクトの設定」からデータベースのアクセス権限を確認してください。`;
        
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