/**
 * 診断ツール - columnMapping問題のデバッグ用
 */

/**
 * 現在のユーザーのconfigJSON内容を詳細診断
 */
function diagnoseCurrentUserConfig() {
  try {
    console.log('🔍 現在ユーザーconfigJSON診断開始');
    
    const currentUser = UserManager.getCurrentEmail();
    console.log('現在のユーザーメール:', currentUser);
    
    if (!currentUser) {
      console.error('❌ ユーザーメールが取得できません');
      return { success: false, error: 'ユーザー認証失敗' };
    }
    
    // DBからユーザー情報を取得
    const userInfo = DB.findUserByEmail(currentUser);
    console.log('ユーザー情報:', {
      found: !!userInfo,
      userId: userInfo?.userId,
      userEmail: userInfo?.userEmail,
      isActive: userInfo?.isActive,
      lastModified: userInfo?.lastModified
    });
    
    if (!userInfo) {
      console.error('❌ ユーザー情報がDBに存在しません');
      return { success: false, error: 'ユーザー情報なし' };
    }
    
    // configJsonを解析
    let configJson = null;
    try {
      configJson = JSON.parse(userInfo.configJson || '{}');
    } catch (parseError) {
      console.error('❌ configJSON解析エラー:', parseError.message);
      return { success: false, error: 'configJSON解析失敗' };
    }
    
    console.log('📊 configJSON詳細分析:', {
      configJsonRaw: userInfo.configJson?.substring(0, 200) + '...',
      configJsonSize: userInfo.configJson?.length || 0,
      parsedKeys: Object.keys(configJson),
      hasSpreadsheetId: !!configJson.spreadsheetId,
      hasSheetName: !!configJson.sheetName,
      hasColumnMapping: !!configJson.columnMapping,
      hasReasonHeader: !!configJson.reasonHeader,
      setupStatus: configJson.setupStatus
    });
    
    // columnMapping詳細診断
    if (configJson.columnMapping) {
      console.log('✅ columnMapping存在:', {
        type: typeof configJson.columnMapping,
        keys: Object.keys(configJson.columnMapping),
        mapping: configJson.columnMapping.mapping,
        confidence: configJson.columnMapping.confidence,
        hasReason: !!configJson.columnMapping.mapping?.reason,
        reasonIndex: configJson.columnMapping.mapping?.reason
      });
      
      if (configJson.columnMapping.mapping) {
        console.log('📋 mapping詳細:', configJson.columnMapping.mapping);
      }
    } else {
      console.warn('⚠️ columnMappingが存在しません');
    }
    
    // 理由ヘッダー関連診断
    console.log('🔍 理由ヘッダー情報:', {
      reasonHeader: configJson.reasonHeader,
      opinionHeader: configJson.opinionHeader,
      classHeader: configJson.classHeader,
      nameHeader: configJson.nameHeader
    });
    
    return {
      success: true,
      userInfo: {
        userId: userInfo.userId,
        userEmail: userInfo.userEmail,
        isActive: userInfo.isActive
      },
      configJson: configJson,
      diagnosis: {
        hasColumnMapping: !!configJson.columnMapping,
        hasReasonMapping: !!configJson.columnMapping?.mapping?.reason,
        hasLegacyReasonHeader: !!configJson.reasonHeader,
        spreadsheetConfigured: !!(configJson.spreadsheetId && configJson.sheetName),
        setupComplete: configJson.setupStatus === 'completed'
      }
    };
    
  } catch (error) {
    console.error('❌ diagnoseCurrentUserConfig エラー:', error.message);
    return { success: false, error: error.message };
  }
}

/**
 * スプレッドシートの実際のヘッダー構造を確認
 */
function diagnoseSpreadsheetHeaders() {
  try {
    console.log('📊 スプレッドシートヘッダー診断開始');
    
    const userDiagnosis = diagnoseCurrentUserConfig();
    if (!userDiagnosis.success) {
      throw new Error('ユーザー診断失敗: ' + userDiagnosis.error);
    }
    
    const config = userDiagnosis.configJson;
    if (!config.spreadsheetId || !config.sheetName) {
      console.warn('⚠️ スプレッドシート情報が不足');
      return { success: false, error: 'スプレッドシート未設定' };
    }
    
    console.log('🔗 スプレッドシートアクセス:', {
      spreadsheetId: config.spreadsheetId.substring(0, 10) + '...',
      sheetName: config.sheetName
    });
    
    // スプレッドシートにアクセス
    const spreadsheet = SpreadsheetApp.openById(config.spreadsheetId);
    const sheet = spreadsheet.getSheetByName(config.sheetName);
    
    if (!sheet) {
      throw new Error(`シート "${config.sheetName}" が見つかりません`);
    }
    
    // ヘッダー行を取得
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    console.log('📋 実際のヘッダー構造:', {
      count: headerRow.length,
      headers: headerRow.map((h, i) => `${i}: "${h}"`),
      dataRows: sheet.getLastRow() - 1
    });
    
    // columnMappingと実際のヘッダーの対応確認
    if (config.columnMapping?.mapping) {
      console.log('🔍 columnMapping vs 実際ヘッダー対応確認:');
      
      Object.entries(config.columnMapping.mapping).forEach(([type, index]) => {
        if (typeof index === 'number' && index >= 0 && index < headerRow.length) {
          console.log(`✅ ${type}[${index}]: "${headerRow[index]}"`);
        } else {
          console.warn(`⚠️ ${type}[${index}]: インデックス範囲外`);
        }
      });
    }
    
    return {
      success: true,
      headerRow: headerRow,
      headerCount: headerRow.length,
      dataRowCount: sheet.getLastRow() - 1,
      columnMapping: config.columnMapping
    };
    
  } catch (error) {
    console.error('❌ diagnoseSpreadsheetHeaders エラー:', error.message);
    return { success: false, error: error.message };
  }
}

/**
 * columnMapping再生成・修復
 */
function regenerateColumnMapping() {
  try {
    console.log('🔧 columnMapping再生成開始');
    
    // 現在の設定を取得
    const userDiagnosis = diagnoseCurrentUserConfig();
    if (!userDiagnosis.success) {
      throw new Error('ユーザー診断失敗');
    }
    
    const config = userDiagnosis.configJson;
    if (!config.spreadsheetId || !config.sheetName) {
      throw new Error('スプレッドシート未設定');
    }
    
    // スプレッドシートデータを取得
    const spreadsheet = SpreadsheetApp.openById(config.spreadsheetId);
    const sheet = spreadsheet.getSheetByName(config.sheetName);
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // サンプルデータも取得（分析精度向上）
    const dataRows = Math.min(10, sheet.getLastRow() - 1);
    let allData = [];
    if (dataRows > 0) {
      allData = sheet.getRange(1, 1, dataRows + 1, sheet.getLastColumn()).getValues();
    }
    
    console.log('📊 データ取得完了:', {
      headerCount: headerRow.length,
      dataRowCount: dataRows,
      headers: headerRow.map((h, i) => `${i}: "${h}"`)
    });
    
    // 新しいcolumnMappingを生成
    const newColumnMapping = generateColumnMapping(headerRow, allData);
    console.log('🎯 新しいcolumnMapping生成完了:', newColumnMapping);
    
    // 設定を更新
    const updatedConfig = {
      ...config,
      columnMapping: newColumnMapping,
      lastModified: new Date().toISOString()
    };
    
    // DBに保存
    const currentUser = UserManager.getCurrentEmail();
    const userInfo = DB.findUserByEmail(currentUser);
    DB.updateUser(userInfo.userId, updatedConfig);
    
    console.log('✅ columnMapping再生成・保存完了');
    
    return {
      success: true,
      oldMapping: config.columnMapping,
      newMapping: newColumnMapping,
      message: 'columnMappingを再生成しました'
    };
    
  } catch (error) {
    console.error('❌ regenerateColumnMapping エラー:', error.message);
    return { success: false, error: error.message };
  }
}