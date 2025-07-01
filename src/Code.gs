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
 * サービスアカウントの認証情報とデータベースのスプレッドシートIDを設定する。
 * @param {string} credsJson - ダウンロードしたサービスアカウントのJSONキーファイルの内容
 * @param {string} dbId - 中央データベースとして使用するスプレッドシートのID
 */
function setupApplication(credsJson, dbId) {
  try {
    JSON.parse(credsJson);
    // ★診断のため、正規表現による検証を一時的に削除
    // var sheetIdRegex = new RegExp("^[a-zA-Z0-9-_]{44}$");
    // if (!sheetIdRegex.test(dbId)) {
    //   throw new Error('無効なスプレッドシートIDです。');
    // }

    var props = PropertiesService.getScriptProperties();
    props.setProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS, credsJson);
    props.setProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID, dbId);

    console.log('✅ セットアップが正常に完了しました。');
  } catch (e) {
    console.error('セットアップエラー:', e);
    throw new Error('セットアップに失敗しました: ' + e.message);
  }
}

/**
 * ウェブアプリのエントリーポイント（仮実装）
 */
function doGet(e) {
  return HtmlService.createHtmlOutput("セットアップが進行中です。この画面は後でアプリケーションのUIに置き換わります。");
}
