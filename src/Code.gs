/**
 * @fileoverview StudyQuest - みんなの回答ボード (サービスアカウントモデル)
 * 最小構成版 - claspの構文エラーを回避するためのクリーンなファイル。
 */

// =================================================================
// グローバル設定
// =================================================================

var SCRIPT_PROPS_KEYS = {
  SERVICE_ACCOUNT_CREDS: 'SERVICE_ACCOUNT_CREDS',
  DATABASE_SPREADSHEET_ID: 'DATABASE_SPREADSHEET_ID'
};

// =================================================================
// セットアップ＆管理用関数
// =================================================================

/**
 * アプリケーションの初期セットアップ（管理者が手動で実行）
 */
function setupApplication() {
  var ui = SpreadsheetApp.getUi();
  
  var credsPrompt = ui.prompt(
    'サービスアカウント認証情報の設定',
    'ダウンロードしたサービスアカウントのJSONキーファイルの内容を貼り付けてください。',
    ui.ButtonSet.OK_CANCEL
  );
  if (credsPrompt.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  var credsJson = credsPrompt.getResponseText();
  
  var dbIdPrompt = ui.prompt(
    'データベースのスプレッドシートIDの設定',
    '中央データベースとして使用するスプレッドシートのIDを入力してください。',
    ui.ButtonSet.OK_CANCEL
  );
  if (dbIdPrompt.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  var dbId = dbIdPrompt.getResponseText();

  try {
    JSON.parse(credsJson);
    var sheetIdRegex = new RegExp("^[a-zA-Z0-9-_]{44}$");
    if (!sheetIdRegex.test(dbId)) {
      throw new Error('無効なスプレッドシートIDです。');
    }

    var props = PropertiesService.getScriptProperties();
    props.setProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS, credsJson);
    props.setProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID, dbId);
    
    ui.alert('セットアップが正常に完了しました。');
  } catch (e) {
    console.error('セットアップエラー:', e);
    ui.alert('セットアップに失敗しました: ' + e.message);
  }
}

/**
 * ウェブアプリのエントリーポイント（仮実装）
 */
function doGet(e) {
  return HtmlService.createHtmlOutput("セットアップが進行中です。この画面は後でアプリケーションのUIに置き換わります。");
}
