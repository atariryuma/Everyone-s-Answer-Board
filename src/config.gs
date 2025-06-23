function getConfig(sheetName) {
  const sheet = getCurrentSpreadsheet().getSheetByName('Config');
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
    nameHeader: idx['名前列ヘッダー']!==undefined ? target[idx['名前列ヘッダー']] : null,
    classHeader: idx['クラス列ヘッダー']!==undefined ? target[idx['クラス列ヘッダー']] : null
  };
}

function saveSheetConfig(sheetName, cfg) {
  const ss = getCurrentSpreadsheet();
  let sheet = ss.getSheetByName('Config');
  if (!sheet) {
    sheet = ss.insertSheet('Config');
  }
  const headers = ['表示シート名','問題文ヘッダー','回答ヘッダー','理由ヘッダー','名前列ヘッダー','クラス列ヘッダー'];
  let values = sheet.getDataRange().getValues();
  if (values.length === 0) {
    sheet.appendRow(headers);
    values = [headers];
  }
  const idx = {};
  headers.forEach((h,i)=>{ idx[h]=i; });
  let rowIndex = -1;
  for (let i=1; i<values.length; i++) {
    if (values[i][idx['表示シート名']] === sheetName) { rowIndex = i; break; }
  }
  const row = new Array(headers.length).fill('');
  row[idx['表示シート名']] = sheetName;
  row[idx['問題文ヘッダー']] = cfg.questionHeader || '';
  row[idx['回答ヘッダー']] = cfg.answerHeader || '';
  row[idx['理由ヘッダー']] = cfg.reasonHeader || '';
  row[idx['名前列ヘッダー']] = cfg.nameHeader || '';
  row[idx['クラス列ヘッダー']] = cfg.classHeader || '';
  if (rowIndex === -1) {
    sheet.appendRow(row);
  } else {
    sheet.getRange(rowIndex+1,1,1,row.length).setValues([row]);
  }
  return 'Config saved';
}

function createConfigSheet() {
  const ss = getCurrentSpreadsheet();
  const created = [];
  let cfgSheet = ss.getSheetByName('Config');
  if (!cfgSheet) {
    cfgSheet = ss.insertSheet('Config');
    const headers = ['表示シート名','問題文ヘッダー','回答ヘッダー','理由ヘッダー','名前列ヘッダー','クラス列ヘッダー'];
    cfgSheet.appendRow(headers);
    created.push('Config');
  }
  const rosterName = typeof ROSTER_CONFIG !== 'undefined' ? ROSTER_CONFIG.SHEET_NAME : 'roster';
  let rosterSheet = ss.getSheetByName(rosterName);
  if (!rosterSheet) {
    rosterSheet = ss.insertSheet(rosterName);
    const rosterHeaders = ['学年','組','番号','姓','名','Googleアカウント','ニックネーム'];
    rosterSheet.appendRow(rosterHeaders);
    created.push(rosterName);
  }
  if (created.length) {
    SpreadsheetApp.getUi().alert(`${created.join('と')}シートを作成しました。設定を入力してください。`);
  } else {
    SpreadsheetApp.getUi().alert('ConfigシートとRosterシートは既に存在します。');
  }
}

if (typeof module !== 'undefined') {
  module.exports = { getConfig, saveSheetConfig, createConfigSheet };
}
