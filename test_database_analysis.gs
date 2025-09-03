/**
 * データベース分析のテスト実行スクリプト
 * GAS環境で直接実行可能
 */

function testDatabaseAnalysis() {
  console.log('🧪 データベース分析テストを実行します...');
  
  try {
    // システム設定の確認
    const dbId = PropertiesService.getScriptProperties().getProperty('DATABASE_SPREADSHEET_ID');
    console.log('データベースID:', dbId ? dbId.substring(0, 10) + '...' : 'なし');
    
    if (!dbId) {
      console.error('❌ DATABASE_SPREADSHEET_IDが設定されていません');
      return {
        success: false,
        error: 'DATABASE_SPREADSHEET_IDが設定されていません'
      };
    }
    
    // Service Account Token取得テスト
    let service;
    try {
      service = getServiceAccountTokenCached();
      console.log('✅ Service Account Token取得成功');
    } catch (error) {
      console.error('❌ Service Account Token取得失敗:', error.message);
      return {
        success: false,
        error: `Service Account Token取得失敗: ${error.message}`
      };
    }
    
    // DB_CONFIG確認
    console.log('📋 DB_CONFIG確認:');
    console.log('- シート名:', DB_CONFIG.SHEET_NAME);
    console.log('- ヘッダー数:', DB_CONFIG.HEADERS.length);
    console.log('- ヘッダー:', DB_CONFIG.HEADERS);
    
    // 実際のデータベース内容を取得・分析
    const sheetName = DB_CONFIG.SHEET_NAME;
    
    try {
      console.log('📊 データベース内容を取得中...');
      const data = batchGetSheetsData(service, dbId, [`'${sheetName}'!A:N`]);
      const values = data.valueRanges[0].values || [];
      
      console.log('📈 基本統計:');
      console.log('- 総行数:', values.length);
      console.log('- データ行数:', Math.max(0, values.length - 1));
      
      if (values.length > 0) {
        const actualHeaders = values[0];
        console.log('- 実際のヘッダー数:', actualHeaders.length);
        console.log('- 実際のヘッダー:', actualHeaders);
        
        // ヘッダー比較
        const headerMatch = JSON.stringify(actualHeaders) === JSON.stringify(DB_CONFIG.HEADERS);
        console.log('- ヘッダー一致:', headerMatch ? '✅' : '❌');
        
        if (!headerMatch) {
          console.log('- 期待:', DB_CONFIG.HEADERS);
          console.log('- 実際:', actualHeaders);
          
          const missing = DB_CONFIG.HEADERS.filter(h => !actualHeaders.includes(h));
          const extra = actualHeaders.filter(h => !DB_CONFIG.HEADERS.includes(h));
          
          if (missing.length > 0) {
            console.log('- 不足ヘッダー:', missing);
          }
          if (extra.length > 0) {
            console.log('- 余分ヘッダー:', extra);
          }
        }
        
        // ユーザーデータ分析
        if (values.length > 1) {
          console.log('\n👥 ユーザーデータ分析:');
          const dataRows = values.slice(1);
          
          // 基本統計
          const userIdIndex = actualHeaders.indexOf('userId');
          const emailIndex = actualHeaders.indexOf('userEmail');
          const isActiveIndex = actualHeaders.indexOf('isActive');
          const spreadsheetIdIndex = actualHeaders.indexOf('spreadsheetId');
          const configJsonIndex = actualHeaders.indexOf('configJson');
          
          let activeCount = 0;
          let hasSpreadsheetCount = 0;
          let hasConfigCount = 0;
          let validEmailCount = 0;
          
          // 最初の5ユーザーの詳細表示
          console.log('\n📝 ユーザーサンプル (最初の5ユーザー):');
          
          dataRows.slice(0, 5).forEach((row, index) => {
            console.log(`\n--- ユーザー ${index + 1} (行 ${index + 2}) ---`);
            
            actualHeaders.forEach((header, colIndex) => {
              const value = row[colIndex];
              console.log(`${header}: ${value || '[空]'}`);
            });
          });
          
          // 統計計算
          dataRows.forEach(row => {
            // アクティブユーザー
            const isActive = row[isActiveIndex];
            if (isActive === true || isActive === 'TRUE' || isActive === 'true') {
              activeCount++;
            }
            
            // スプレッドシートID
            const spreadsheetId = row[spreadsheetIdIndex];
            if (spreadsheetId && spreadsheetId !== '') {
              hasSpreadsheetCount++;
            }
            
            // 設定JSON
            const configJson = row[configJsonIndex];
            if (configJson && configJson !== '') {
              hasConfigCount++;
            }
            
            // メールアドレス検証
            const email = row[emailIndex];
            if (email && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
              validEmailCount++;
            }
          });
          
          console.log('\n📊 統計サマリー:');
          console.log('- 総ユーザー数:', dataRows.length);
          console.log('- アクティブユーザー:', activeCount);
          console.log('- スプレッドシート設定済み:', hasSpreadsheetCount);
          console.log('- 設定JSON保持:', hasConfigCount);
          console.log('- 有効メールアドレス:', validEmailCount);
          
          // configJson詳細分析
          if (hasConfigCount > 0) {
            console.log('\n🔧 configJson分析:');
            let configAnalysisCount = 0;
            
            dataRows.forEach((row, index) => {
              const configJson = row[configJsonIndex];
              if (configJson && configJson !== '' && configAnalysisCount < 3) {
                console.log(`\n--- 設定サンプル ${configAnalysisCount + 1} (ユーザー行 ${index + 2}) ---`);
                
                try {
                  const config = JSON.parse(configJson);
                  console.log('有効なJSON:', Object.keys(config));
                  console.log('設定内容:', JSON.stringify(config, null, 2));
                } catch (error) {
                  console.log('❌ 無効なJSON:', configJson.substring(0, 100));
                }
                configAnalysisCount++;
              }
            });
          }
        }
      }
      
      console.log('\n✅ データベース分析完了');
      return {
        success: true,
        totalRows: values.length,
        dataRows: Math.max(0, values.length - 1),
        hasHeaders: values.length > 0,
        headerMatch: values.length > 0 ? JSON.stringify(values[0]) === JSON.stringify(DB_CONFIG.HEADERS) : false
      };
      
    } catch (dataError) {
      console.error('❌ データベースデータ取得エラー:', dataError.message);
      return {
        success: false,
        error: `データ取得エラー: ${dataError.message}`
      };
    }
    
  } catch (error) {
    console.error('❌ テスト実行エラー:', {
      message: error.message,
      stack: error.stack
    });
    return {
      success: false,
      error: error.message
    };
  }
}