function getConfig(sheetName) {
  try {
    let sheet = getCurrentSpreadsheet().getSheetByName('Config');
    
    // Configシートが存在しない場合は作成
    if (!sheet) {
      debugLog('Config sheet not found, creating new one');
      sheet = getCurrentSpreadsheet().insertSheet('Config');
      const headers = ['表示シート名','問題文ヘッダー','回答ヘッダー','理由ヘッダー','名前列ヘッダー','クラス列ヘッダー'];
      sheet.appendRow(headers);
      
      // 自動設定を試行
      return createAutoConfig(sheetName);
    }
    
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) {
      debugLog('Config sheet has no data, creating auto config');
      return createAutoConfig(sheetName);
    }
    
    // 正しいヘッダーを定義
    const correctHeaders = ['表示シート名','問題文ヘッダー','回答ヘッダー','理由ヘッダー','名前列ヘッダー','クラス列ヘッダー'];
    const existingHeaders = values[0];
    
    // ヘッダー行が正しいかチェックし、必要に応じて修正
    const isHeaderCorrect = correctHeaders.every((header, index) => 
      existingHeaders[index] && String(existingHeaders[index]).trim() === header
    );
    
    if (!isHeaderCorrect || existingHeaders.length !== correctHeaders.length) {
      debugLog('Config sheet header is incorrect, fixing it');
      // 1行目（ヘッダー）を正しいヘッダーで置き換え
      sheet.getRange(1, 1, 1, correctHeaders.length).setValues([correctHeaders]);
      values[0] = correctHeaders;
    }
    
    const headers = values[0];
    const idx = {};
    headers.forEach((h,i) => { if(h) idx[h] = i; });
    
    const target = values.find((row, rIdx) => rIdx>0 && row[idx['表示シート名']]===sheetName);
    if(!target) {
      debugLog(`No config found for sheet ${sheetName}, creating auto config`);
      return createAutoConfig(sheetName);
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
    
    // エラーが発生した場合も自動設定を試行
    try {
      return createAutoConfig(sheetName);
    } catch (autoError) {
      console.error('Auto config creation failed:', autoError.message);
      throw error; // 元のエラーを投げる
    }
  }
}

/**
 * シートのヘッダーを分析して自動設定を作成
 */
function createAutoConfig(sheetName) {
  try {
    const targetSheet = getCurrentSpreadsheet().getSheetByName(sheetName);
    if (!targetSheet) {
      throw new Error(`シート「${sheetName}」が見つかりません。`);
    }
    
    // ヘッダー行を取得
    const lastColumn = targetSheet.getLastColumn();
    if (lastColumn === 0) {
      throw new Error(`シート「${sheetName}」にデータがありません。`);
    }
    
    const headerRow = targetSheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    const headers = headerRow.map(v => String(v || '').trim()).filter(h => h !== '');
    
    debugLog('Creating auto config for headers:', headers);
    
    // ヘッダーを推測（Code.gsの関数を使用）
    if (typeof guessHeadersFromArray === 'undefined') {
      throw new Error('ヘッダー推測機能が利用できません。');
    }
    const guessedConfig = guessHeadersFromArray(headers);
    
    // 設定を保存
    if (guessedConfig.answerHeader) {
      saveSheetConfig(sheetName, guessedConfig);
      debugLog('Auto-created config saved:', guessedConfig);
      return guessedConfig;
    } else {
      throw new Error('適切なヘッダーを推測できませんでした。手動で設定してください。');
    }
    
  } catch (error) {
    console.error('createAutoConfig error:', error);
    throw new Error(`自動設定の作成に失敗しました: ${error.message}`);
  }
}

function saveSheetConfig(sheetName, cfg) {
  try {
    const ss = getCurrentSpreadsheet();
    return saveSheetConfigForSpreadsheet(ss, sheetName, cfg);
  } catch (error) {
    console.error('Error saving sheet config:', error);
    throw new Error(`設定の保存に失敗しました: ${error.message}`);
  }
}

function saveSheetConfigForSpreadsheet(ss, sheetName, cfg) {
  try {
    let sheet = ss.getSheetByName('Config');
    if (!sheet) {
      debugLog('Creating Config sheet for the first time');
      sheet = ss.insertSheet('Config');
    }
    
    const headers = ['表示シート名','問題文ヘッダー','回答ヘッダー','理由ヘッダー','名前列ヘッダー','クラス列ヘッダー'];
    let values = [];
    
    // Get existing data or create header row
    try {
      values = sheet.getDataRange().getValues();
    } catch (error) {
      debugLog('Config sheet appears to be empty, creating header row');
      values = [];
    }
    
    if (values.length === 0) {
      sheet.appendRow(headers);
      values = [headers];
    } else {
      // ヘッダー行が正しいかチェックし、必要に応じて修正
      const existingHeaders = values[0];
      const isHeaderCorrect = headers.every((header, index) => 
        existingHeaders[index] && String(existingHeaders[index]).trim() === header
      );
      
      if (!isHeaderCorrect || existingHeaders.length !== headers.length) {
        debugLog('Config sheet header is incorrect, fixing it');
        // 1行目（ヘッダー）を正しいヘッダーで置き換え
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        values[0] = headers;
      }
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
      debugLog('Added new config row for sheet:', sheetName);
    } else {
      sheet.getRange(rowIndex+1,1,1,row.length).setValues([row]);
      debugLog('Updated existing config row for sheet:', sheetName);
    }
    
    return `シート「${sheetName}」の設定を保存しました`;
  } catch (error) {
    // Re-throw with more context
    throw new Error(`saveSheetConfigForSpreadsheet failed for sheet ${sheetName}: ${error.message}`);
  }
}

function createConfigSheet() {
  const ss = getCurrentSpreadsheet();
  const created = [];
  const correctHeaders = ['表示シート名','問題文ヘッダー','回答ヘッダー','理由ヘッダー','名前列ヘッダー','クラス列ヘッダー'];
  
  let cfgSheet = ss.getSheetByName('Config');
  if (!cfgSheet) {
    cfgSheet = ss.insertSheet('Config');
    cfgSheet.appendRow(correctHeaders);
    created.push('Config');
  } else {
    // 既存のConfigシートのヘッダーをチェックして修正
    try {
      const values = cfgSheet.getDataRange().getValues();
      if (values.length > 0) {
        const existingHeaders = values[0];
        const isHeaderCorrect = correctHeaders.every((header, index) => 
          existingHeaders[index] && String(existingHeaders[index]).trim() === header
        );
        
        if (!isHeaderCorrect || existingHeaders.length !== correctHeaders.length) {
          debugLog('Fixing incorrect Config sheet header');
          cfgSheet.getRange(1, 1, 1, correctHeaders.length).setValues([correctHeaders]);
          created.push('Config (header fixed)');
        }
      } else {
        // 空のシートの場合、ヘッダーを追加
        cfgSheet.appendRow(correctHeaders);
        created.push('Config (header added)');
      }
    } catch (error) {
      debugLog('Error checking Config sheet header:', error);
      // エラーの場合も安全にヘッダーを設定
      cfgSheet.clear();
      cfgSheet.appendRow(correctHeaders);
      created.push('Config (recreated)');
    }
  }
  
  if (created.length) {
    SpreadsheetApp.getUi().alert(`${created.join('と')}シートを作成/修正しました。設定を入力してください。`);
  } else {
    SpreadsheetApp.getUi().alert('Configシートは既に正しく設定されています。');
  }
}

if (typeof module !== 'undefined') {
  module.exports = { getConfig, saveSheetConfig, createConfigSheet };
}
