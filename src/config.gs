function getConfig(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Config');
  if (!sheet) {
    throw new Error('Config シートが見つかりません');
  }
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    throw new Error('Config シートに設定がありません');
  }
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (row[0] === sheetName) {
      return {
        questionHeader: row[1],
        answerHeader: row[2],
        reasonHeader: row[3] || null,
        nameMode: row[4],
        nameHeader: row[5] || null,
        classHeader: row[6] || null
      };
    }
  }
  throw new Error('Config に設定がありません');
}

if (typeof module !== 'undefined') {
  module.exports = { getConfig };
}
