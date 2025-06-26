function getConfig(sheetName) {
  try {
    const sheet = getCurrentSpreadsheet().getSheetByName('Config');
    if (!sheet) {
      throw new Error('Configシートが見つかりません。新しいスプレッドシートの場合、設定を自動作成します。');
    }
    
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) {
      throw new Error('Configシートにデータがありません。新しいシートの設定を作成してください。');
    }
    
    const headers = values[0];
    const idx = {};
    headers.forEach((h,i) => { if(h) idx[h] = i; });
    
    const target = values.find((row, rIdx) => rIdx>0 && row[idx['表示シート名']]===sheetName);
    if(!target) {
      throw new Error(`シート「${sheetName}」の設定がConfigシートに見つかりません。自動で基本設定を作成します。`);
    }
    
    return {
      questionHeader: target[idx['問題文ヘッダー']] || '',
      answerHeader: target[idx['回答ヘッダー']] || '',
      reasonHeader: idx['理由ヘッダー']!==undefined ? target[idx['理由ヘッダー']] || '' : '',
      nameHeader: idx['名前列ヘッダー']!==undefined ? target[idx['名前列ヘッダー']] || '' : '',
      classHeader: idx['クラス列ヘッダー']!==undefined ? target[idx['クラス列ヘッダー']] || '' : ''
    };
  } catch (error) {
    console.error('getConfig error for sheet:', sheetName, error.message);
    throw error;
  }
}

function saveSheetConfig(sheetName, cfg) {
  try {
    const ss = getCurrentSpreadsheet();
    let sheet = ss.getSheetByName('Config');
    if (!sheet) {
      console.log('Creating Config sheet for the first time');
      sheet = ss.insertSheet('Config');
    }
    
    const headers = ['表示シート名','問題文ヘッダー','回答ヘッダー','理由ヘッダー','名前列ヘッダー','クラス列ヘッダー'];
    let values = [];
    
    // Get existing data or create header row
    try {
      values = sheet.getDataRange().getValues();
    } catch (error) {
      console.log('Config sheet appears to be empty, creating header row');
      values = [];
    }
    
    if (values.length === 0) {
      sheet.appendRow(headers);
      values = [headers];
    }
    
    const idx = {};
    headers.forEach((h,i)=>{ idx[h]=i; });
    
    // Find existing row or determine it needs to be added
    let rowIndex = -1;
    for (let i=1; i<values.length; i++) {
      if (values[i][idx['表示シート名']] === sheetName) { 
        rowIndex = i; 
        break; 
      }
    }
    
    // Prepare row data
    const row = new Array(headers.length).fill('');
    row[idx['表示シート名']] = sheetName;
    row[idx['問題文ヘッダー']] = cfg.questionHeader || '';
    row[idx['回答ヘッダー']] = cfg.answerHeader || '';
    row[idx['理由ヘッダー']] = cfg.reasonHeader || '';
    row[idx['名前列ヘッダー']] = cfg.nameHeader || '';
    row[idx['クラス列ヘッダー']] = cfg.classHeader || '';
    
    // Save data
    if (rowIndex === -1) {
      sheet.appendRow(row);
      console.log('Added new config row for sheet:', sheetName);
    } else {
      sheet.getRange(rowIndex+1,1,1,row.length).setValues([row]);
      console.log('Updated existing config row for sheet:', sheetName);
    }
    
    return `シート「${sheetName}」の設定を保存しました`;
  } catch (error) {
    console.error('Error saving sheet config:', error);
    throw new Error(`設定の保存に失敗しました: ${error.message}`);
  }
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
  if (created.length) {
    SpreadsheetApp.getUi().alert(`${created.join('と')}シートを作成しました。設定を入力してください。`);
  } else {
    SpreadsheetApp.getUi().alert('Configシートは既に存在します。');
  }
}

if (typeof module !== 'undefined') {
  module.exports = { getConfig, saveSheetConfig, createConfigSheet };
}
