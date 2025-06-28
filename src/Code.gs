function openActiveSpreadsheet() {
  const props = PropertiesService.getUserProperties();
  const userId = props.getProperty('CURRENT_USER_ID');
  if (!userId) {
    throw new Error('ユーザー情報が見つかりません。');
  }
  const userInfo = getUserInfo(userId);
  if (!userInfo || !userInfo.spreadsheetId) {
    throw new Error('アクティブなスプレッドシートが見つかりません。');
  }
  const spreadsheet = SpreadsheetApp.openById(userInfo.spreadsheetId);
  return spreadsheet.getUrl();
}