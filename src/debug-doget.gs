/**
 * @fileoverview デバッグ用の最小限doGet関数
 * 問題を特定するための簡単なテスト
 */

/**
 * デバッグ用のシンプルなdoGet関数
 * この関数が動作するかで基本的なアクセス可能性をテスト
 */
function debugDoGet(e) {
  try {
    console.log('DEBUG: debugDoGet called');
    console.log('DEBUG: Event object:', JSON.stringify(e));
    
    // 基本的な情報を取得
    const userEmail = Session.getActiveUser().getEmail();
    console.log('DEBUG: User email:', userEmail);
    
    const webAppUrl = ScriptApp.getService().getUrl();
    console.log('DEBUG: Web App URL:', webAppUrl);
    
    // 簡単なHTMLレスポンスを返す
    const html = HtmlService.createHtmlOutput(`
      <h1>デバッグテスト成功</h1>
      <p>doGet関数は正常に動作しています</p>
      <p>ユーザー: ${userEmail}</p>
      <p>WebアプリURL: ${webAppUrl}</p>
      <p>時刻: ${new Date().toISOString()}</p>
      <p>パラメータ: ${JSON.stringify(e.parameter)}</p>
      
      <h2>次のステップ</h2>
      <p>このページが表示されている場合、基本的なアクセスは正常です</p>
      <a href="${webAppUrl}">通常のアプリに戻る</a>
    `);
    
    return html;
    
  } catch (error) {
    console.error('DEBUG: debugDoGet error:', error.message);
    console.error('DEBUG: Error stack:', error.stack);
    
    // エラー時でもHTMLを返す
    const errorHtml = HtmlService.createHtmlOutput(`
      <h1>デバッグエラー</h1>
      <p>エラーが発生しました: ${error.message}</p>
      <p>時刻: ${new Date().toISOString()}</p>
    `);
    
    return errorHtml;
  }
}

/**
 * メインのdoGet関数の先頭にデバッグログを追加
 */
function addDebugToMainDoGet() {
  // このコメントは、main.gsのdoGet関数の先頭に以下を追加することを示します：
  /*
  console.log('MAIN_DEBUG: doGet called at', new Date().toISOString());
  console.log('MAIN_DEBUG: Event object:', JSON.stringify(e, null, 2));
  
  try {
    const userEmail = Session.getActiveUser().getEmail();
    console.log('MAIN_DEBUG: User email obtained:', userEmail);
  } catch (emailError) {
    console.error('MAIN_DEBUG: Failed to get user email:', emailError.message);
  }
  */
}

/**
 * 権限テスト関数
 */
function testPermissions() {
  try {
    console.log('PERMISSION_TEST: Starting permission tests');
    
    // 1. Session.getActiveUser()テスト
    try {
      const userEmail = Session.getActiveUser().getEmail();
      console.log('PERMISSION_TEST: User email OK:', userEmail);
    } catch (sessionError) {
      console.error('PERMISSION_TEST: Session error:', sessionError.message);
    }
    
    // 2. PropertiesServiceテスト
    try {
      const props = PropertiesService.getUserProperties();
      const testProp = props.getProperty('TEST_PROPERTY');
      console.log('PERMISSION_TEST: PropertiesService OK, test property:', testProp);
    } catch (propsError) {
      console.error('PERMISSION_TEST: PropertiesService error:', propsError.message);
    }
    
    // 3. ScriptApp.getService().getUrl()テスト
    try {
      const webAppUrl = ScriptApp.getService().getUrl();
      console.log('PERMISSION_TEST: ScriptApp URL OK:', webAppUrl);
    } catch (urlError) {
      console.error('PERMISSION_TEST: ScriptApp URL error:', urlError.message);
    }
    
    console.log('PERMISSION_TEST: All tests completed');
    
  } catch (overallError) {
    console.error('PERMISSION_TEST: Overall error:', overallError.message);
  }
}