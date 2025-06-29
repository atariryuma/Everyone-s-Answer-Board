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
 * セットアップ用HTMLから呼び出され、設定を保存しDBを作成する関数
 * @param {string} apiUrl - 管理者向けログ記録APIのURL
 * @returns {string} 処理結果のメッセージ
 */
function saveSettingsAndCreateDb(apiUrl) {
  try {
    const properties = PropertiesService.getScriptProperties();
    
    if (!apiUrl || !apiUrl.startsWith('https://script.google.com/macros/s/')) {
      throw new Error('入力されたURLが無効です。正しい「管理者向けログ記録API」のウェブアプリURLを入力してください。');
    }
    properties.setProperty('LOGGER_API_URL', apiUrl);

    // データベースがなければ作成する
    getOrCreateMainDatabase();
    
    return '✅ 設定が正常に保存されました。このプロジェクトを再度デプロイし直し、発行された新しいURLをご利用ください。';
  } catch(e) {
    Logger.log(e);
    return `❌ エラーが発生しました: ${e.message}`;
  }
}