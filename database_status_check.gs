/**
 * データベース状況調査スクリプト（軽量版）
 * CLAUDE.md準拠で現在のデータベース状況を調査・報告
 */

/**
 * データベースの基本的な状況を調査する軽量版関数
 * 実際のGAS環境で実行可能
 * @returns {Object} 調査結果レポート
 */
function checkDatabaseStatus() {
  console.log('🔍 データベース状況調査を開始します...');
  
  const startTime = new Date();
  
  try {
    // システム設定の確認
    const dbId = PropertiesService.getScriptProperties().getProperty('DATABASE_SPREADSHEET_ID');
    
    if (!dbId) {
      return {
        success: false,
        error: 'DATABASE_SPREADSHEET_IDが設定されていません',
        timestamp: startTime.toISOString()
      };
    }

    // Service Account Token取得
    const service = getServiceAccountTokenCached();
    const sheetName = DB_CONFIG.SHEET_NAME;
    
    console.log('📊 基本情報:');
    console.log('- データベースID:', dbId.substring(0, 10) + '...');
    console.log('- シート名:', sheetName);
    console.log('- 期待ヘッダー数:', DB_CONFIG.HEADERS.length);
    console.log('- 期待ヘッダー:', DB_CONFIG.HEADERS);
    
    // データベース内容を取得
    const data = batchGetSheetsData(service, dbId, [`'${sheetName}'!A:N`]);
    const values = data.valueRanges[0].values || [];
    
    if (values.length === 0) {
      return {
        success: true,
        isEmpty: true,
        totalRows: 0,
        dataRows: 0,
        message: 'データベースは空です',
        timestamp: startTime.toISOString()
      };
    }
    
    const actualHeaders = values[0] || [];
    const dataRows = values.slice(1);
    
    // ヘッダー比較
    const headerMatch = JSON.stringify(actualHeaders) === JSON.stringify(DB_CONFIG.HEADERS);
    
    // 基本統計の計算
    const stats = calculateBasicStats(actualHeaders, dataRows);
    
    const report = {
      success: true,
      isEmpty: false,
      timestamp: startTime.toISOString(),
      executionTime: Date.now() - startTime.getTime() + 'ms',
      
      // 基本情報
      databaseId: dbId.substring(0, 10) + '...',
      sheetName: sheetName,
      
      // スキーマ情報
      schema: {
        expectedHeaders: DB_CONFIG.HEADERS,
        expectedHeaderCount: DB_CONFIG.HEADERS.length,
        actualHeaders: actualHeaders,
        actualHeaderCount: actualHeaders.length,
        headerMatch: headerMatch
      },
      
      // データ統計
      data: {
        totalRows: values.length,
        dataRows: dataRows.length,
        statistics: stats
      },
      
      // 問題のサマリー
      issues: []
    };
    
    // 問題の特定
    if (!headerMatch) {
      const missing = DB_CONFIG.HEADERS.filter(h => !actualHeaders.includes(h));
      const extra = actualHeaders.filter(h => !DB_CONFIG.HEADERS.includes(h));
      
      if (missing.length > 0) {
        report.issues.push({
          type: 'schema_mismatch',
          severity: 'error',
          description: '不足しているヘッダー',
          details: missing
        });
      }
      
      if (extra.length > 0) {
        report.issues.push({
          type: 'schema_mismatch', 
          severity: 'warning',
          description: '余分なヘッダー',
          details: extra
        });
      }
    }
    
    if (stats.emptyConfigCount > 0) {
      report.issues.push({
        type: 'data_quality',
        severity: 'warning',
        description: 'configJsonが空のユーザー',
        count: stats.emptyConfigCount
      });
    }
    
    if (stats.emptySpreadsheetCount > 0) {
      report.issues.push({
        type: 'data_quality',
        severity: 'warning', 
        description: 'spreadsheetIdが空のユーザー',
        count: stats.emptySpreadsheetCount
      });
    }
    
    console.log('✅ データベース状況調査完了');
    console.log('📊 レポート:', JSON.stringify(report, null, 2));
    
    return report;
    
  } catch (error) {
    console.error('❌ データベース調査エラー:', {
      message: error.message,
      stack: error.stack
    });
    
    return {
      success: false,
      error: error.message,
      timestamp: startTime.toISOString(),
      executionTime: Date.now() - startTime.getTime() + 'ms'
    };
  }
}

/**
 * 基本統計を計算する
 * @param {Array} headers - ヘッダー配列
 * @param {Array} dataRows - データ行配列  
 * @returns {Object} 統計情報
 */
function calculateBasicStats(headers, dataRows) {
  if (dataRows.length === 0) {
    return {
      totalUsers: 0,
      activeUsers: 0,
      inactiveUsers: 0,
      usersWithSpreadsheet: 0,
      usersWithConfig: 0,
      emptySpreadsheetCount: 0,
      emptyConfigCount: 0
    };
  }
  
  // インデックスを取得
  const isActiveIndex = headers.indexOf('isActive');
  const spreadsheetIdIndex = headers.indexOf('spreadsheetId');
  const configJsonIndex = headers.indexOf('configJson');
  const userEmailIndex = headers.indexOf('userEmail');
  
  let activeUsers = 0;
  let usersWithSpreadsheet = 0;
  let usersWithConfig = 0;
  let validEmails = 0;
  let emptySpreadsheetCount = 0;
  let emptyConfigCount = 0;
  
  dataRows.forEach(row => {
    // アクティブユーザー統計
    const isActive = row[isActiveIndex];
    if (isActive === true || isActive === 'TRUE' || isActive === 'true') {
      activeUsers++;
    }
    
    // スプレッドシート統計
    const spreadsheetId = row[spreadsheetIdIndex];
    if (spreadsheetId && spreadsheetId !== '') {
      usersWithSpreadsheet++;
    } else {
      emptySpreadsheetCount++;
    }
    
    // 設定統計  
    const configJson = row[configJsonIndex];
    if (configJson && configJson !== '') {
      usersWithConfig++;
    } else {
      emptyConfigCount++;
    }
    
    // メールアドレス統計
    const email = row[userEmailIndex];
    if (email && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      validEmails++;
    }
  });
  
  return {
    totalUsers: dataRows.length,
    activeUsers: activeUsers,
    inactiveUsers: dataRows.length - activeUsers,
    usersWithSpreadsheet: usersWithSpreadsheet,
    usersWithConfig: usersWithConfig,
    validEmails: validEmails,
    emptySpreadsheetCount: emptySpreadsheetCount,
    emptyConfigCount: emptyConfigCount,
    
    // パーセンテージ
    percentages: {
      activeUsers: Math.round((activeUsers / dataRows.length) * 100),
      usersWithSpreadsheet: Math.round((usersWithSpreadsheet / dataRows.length) * 100),
      usersWithConfig: Math.round((usersWithConfig / dataRows.length) * 100),
      validEmails: Math.round((validEmails / dataRows.length) * 100)
    }
  };
}

/**
 * 設定JSON内容のサンプル分析
 * @param {Array} headers - ヘッダー配列
 * @param {Array} dataRows - データ行配列
 * @param {number} sampleCount - 分析するサンプル数
 * @returns {Array} サンプル分析結果
 */
function analyzeConfigSamples(headers, dataRows, sampleCount = 3) {
  const configJsonIndex = headers.indexOf('configJson');
  const userEmailIndex = headers.indexOf('userEmail');
  
  if (configJsonIndex === -1) {
    return [];
  }
  
  const samples = [];
  let analyzed = 0;
  
  for (let i = 0; i < dataRows.length && analyzed < sampleCount; i++) {
    const row = dataRows[i];
    const configJson = row[configJsonIndex];
    const userEmail = row[userEmailIndex];
    
    if (configJson && configJson !== '') {
      try {
        const config = JSON.parse(configJson);
        samples.push({
          userEmail: userEmail || 'unknown',
          rowIndex: i + 2, // スプレッドシートの行番号（1ベース + ヘッダー）
          isValidJson: true,
          configKeys: Object.keys(config),
          hasAppName: !!config.appName,
          hasDisplayMode: !!config.displayMode,
          hasReactionsEnabled: !!config.reactionsEnabled,
          content: config
        });
        analyzed++;
      } catch (error) {
        samples.push({
          userEmail: userEmail || 'unknown',
          rowIndex: i + 2,
          isValidJson: false,
          error: error.message,
          rawContent: configJson.substring(0, 100) + (configJson.length > 100 ? '...' : '')
        });
        analyzed++;
      }
    }
  }
  
  return samples;
}

/**
 * 簡単な実行テスト関数
 * GAS環境でテスト実行する際に利用
 */
function testDatabaseStatusCheck() {
  console.log('🧪 データベース状況調査のテスト実行...');
  
  try {
    const result = checkDatabaseStatus();
    console.log('✅ テスト実行完了');
    console.log('結果:', result);
    return result;
  } catch (error) {
    console.error('❌ テスト実行エラー:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}