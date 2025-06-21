function getConfig(sheetName) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Config');
  if (!sheet) throw new Error('Configシートが見つかりません。');
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) throw new Error('Configシートが空です。');
  const headers = values[0];
  const idx = {};
  headers.forEach((h,i) => { if(h) idx[h] = i; });
  const target = values.find((row, rIdx) => rIdx>0 && row[idx['表示シート名']]===sheetName);
  if(!target) throw new Error(`Configに${sheetName}の設定がありません。`);
  return {
    questionHeader: target[idx['問題文ヘッダー']],
    answerHeader: target[idx['回答ヘッダー']],
    reasonHeader: idx['理由ヘッダー']!==undefined ? target[idx['理由ヘッダー']] : null,
    nameMode: target[idx['名前取得モード']] || '同一シート',
    nameHeader: idx['名前列ヘッダー']!==undefined ? target[idx['名前列ヘッダー']] : null,
    classHeader: idx['クラス列ヘッダー']!==undefined ? target[idx['クラス列ヘッダー']] : null
  };
}

if (typeof module !== 'undefined') {
  module.exports = { getConfig };
}
