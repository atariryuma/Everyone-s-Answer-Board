/**
 * getSpreadsheetList テスト用スクリプト
 * GAS環境で直接実行してnull問題を診断
 */

function testGetSpreadsheetList() {
  console.log('=== getSpreadsheetList テスト開始 ===');

  try {
    // 1. DataService が定義されているかチェック
    console.log('DataService定義チェック:', typeof DataService);

    if (typeof DataService === 'undefined') {
      console.error('DataService が定義されていません');
      return;
    }

    // 2. getSpreadsheetList メソッドが存在するかチェック
    console.log('getSpreadsheetList メソッドチェック:', typeof DataService.getSpreadsheetList);

    if (typeof DataService.getSpreadsheetList !== 'function') {
      console.error('DataService.getSpreadsheetList が関数ではありません');
      return;
    }

    // 3. 実際に実行してみる
    console.log('DataService.getSpreadsheetList() 実行開始...');
    const result = DataService.getSpreadsheetList();

    console.log('実行結果:', {
      resultType: typeof result,
      isNull: result === null,
      isUndefined: result === undefined,
      result: result
    });

    return result;

  } catch (error) {
    console.error('テスト実行エラー:', {
      message: error.message,
      stack: error.stack,
      name: error.name
    });
    return null;
  }
}

// main.gs の getSpreadsheetList を直接テスト
function testMainGetSpreadsheetList() {
  console.log('=== main.gs getSpreadsheetList テスト開始 ===');

  try {
    const result = getSpreadsheetList();
    console.log('main.gs getSpreadsheetList 結果:', {
      resultType: typeof result,
      isNull: result === null,
      result: result
    });

    return result;
  } catch (error) {
    console.error('main.gs getSpreadsheetList エラー:', error);
    return null;
  }
}