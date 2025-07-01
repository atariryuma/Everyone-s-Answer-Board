/**
 * config.gs - 軽量化版
 * 新サービスアカウントアーキテクチャで必要最小限の関数のみ
 */

/**
 * 現在のユーザーのスプレッドシートを取得
 * AdminPanel.htmlから呼び出される
 */
function getCurrentSpreadsheet() {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーIDが設定されていません');
    }
    
    var userInfo = findUserById(currentUserId);
    if (!userInfo || !userInfo.spreadsheetId) {
      throw new Error('ユーザー情報またはスプレッドシートIDが見つかりません');
    }
    
    return SpreadsheetApp.openById(userInfo.spreadsheetId);
  } catch (e) {
    console.error('getCurrentSpreadsheet エラー: ' + e.message);
    throw new Error('スプレッドシートの取得に失敗しました: ' + e.message);
  }
}

/**
 * 簡易設定取得関数（AdminPanel.htmlとの互換性のため）
 * 新アーキテクチャでは基本的にデフォルト設定を使用
 */
function getConfig(sheetName) {
  try {
    return {
      questionHeader: '問題',
      answerHeader: '回答',
      reasonHeader: '理由',
      nameHeader: '名前',
      classHeader: 'クラス'
    };
  } catch (error) {
    console.error('getConfig error for sheet:', sheetName, error.message);
    return {
      questionHeader: '',
      answerHeader: '回答',
      reasonHeader: '',
      nameHeader: '',
      classHeader: ''
    };
  }
}

/**
 * 簡易設定保存関数（AdminPanel.htmlとの互換性のため）
 * 新アーキテクチャでは実際の保存は行わない
 */
function saveSheetConfig(sheetName, cfg) {
  try {
    console.log('設定保存（新アーキテクチャでは無操作）:', sheetName, cfg);
    return `シート「${sheetName}」の設定を確認しました`;
  } catch (error) {
    console.error('Error in saveSheetConfig:', error);
    return `設定の確認中にエラーが発生しました: ${error.message}`;
  }
}