/**
 * データベース構造と内容の詳細分析スクリプト
 * CLAUDE.md準拠でデータベースの現状を調査・報告
 */

/**
 * データベースの構造と内容を包括的に調査する
 * @returns {Object} 調査結果レポート
 */
function analyzeDatabaseStructureAndContent() {
  console.log('🔍 データベース構造・内容分析を開始します...');
  
  try {
    const service = getServiceAccountTokenCached();
    const dbId = PropertiesService.getScriptProperties().getProperty('DATABASE_SPREADSHEET_ID');
    
    if (!dbId) {
      throw new Error('DATABASE_SPREADSHEET_IDが設定されていません');
    }

    const report = {
      timestamp: new Date().toISOString(),
      databaseId: dbId.substring(0, 10) + '...',
      schema: analyzeSchema(),
      content: analyzeContent(service, dbId),
      issues: []
    };

    console.log('📊 データベース分析レポート:', JSON.stringify(report, null, 2));
    return report;

  } catch (error) {
    console.error('❌ データベース分析エラー:', {
      error: error.message,
      stack: error.stack
    });
    throw error;
  }
}

/**
 * データベーススキーマの分析
 * @returns {Object} スキーマ分析結果
 */
function analyzeSchema() {
  return {
    configSource: 'database.gs DB_CONFIG',
    sheetName: DB_CONFIG.SHEET_NAME,
    expectedColumns: DB_CONFIG.HEADERS.length,
    headers: DB_CONFIG.HEADERS,
    headerDetails: DB_CONFIG.HEADERS.map((header, index) => ({
      index: index,
      name: header,
      type: getExpectedDataType(header)
    }))
  };
}

/**
 * 期待されるデータ型を取得
 * @param {string} header - ヘッダー名
 * @returns {string} データ型
 */
function getExpectedDataType(header) {
  const typeMap = {
    'userId': 'string (UUID)',
    'userEmail': 'string (email)',
    'createdAt': 'string (ISO datetime)',
    'lastAccessedAt': 'string (ISO datetime)', 
    'isActive': 'boolean',
    'spreadsheetId': 'string (Sheets ID)',
    'spreadsheetUrl': 'string (URL)',
    'configJson': 'string (JSON)',
    'formUrl': 'string (URL)',
    'sheetName': 'string',
    'columnMappingJson': 'string (JSON)',
    'publishedAt': 'string (ISO datetime)',
    'appUrl': 'string (URL)',
    'lastModified': 'string (ISO datetime)'
  };
  
  return typeMap[header] || 'unknown';
}

/**
 * データベース内容の詳細分析
 * @param {Object} service - Service Account Token
 * @param {string} dbId - データベーススプレッドシートID
 * @returns {Object} 内容分析結果
 */
function analyzeContent(service, dbId) {
  const sheetName = DB_CONFIG.SHEET_NAME;
  
  try {
    // データベース全体を取得 (14列対応)
    const data = batchGetSheetsData(service, dbId, [`'${sheetName}'!A:N`]);
    const values = data.valueRanges[0].values || [];
    
    if (values.length === 0) {
      return {
        isEmpty: true,
        totalRows: 0,
        dataRows: 0,
        actualHeaders: [],
        users: []
      };
    }
    
    const actualHeaders = values[0] || [];
    const dataRows = values.slice(1);
    
    const result = {
      isEmpty: false,
      totalRows: values.length,
      dataRows: dataRows.length,
      actualHeaders: actualHeaders,
      schemaComparison: compareHeaders(actualHeaders),
      users: analyzeUserData(actualHeaders, dataRows),
      statistics: calculateStatistics(actualHeaders, dataRows)
    };
    
    return result;
    
  } catch (error) {
    console.error('❌ データベース内容取得エラー:', error.message);
    return {
      error: error.message,
      isEmpty: true,
      totalRows: 0
    };
  }
}

/**
 * ヘッダーの比較分析
 * @param {Array} actualHeaders - 実際のヘッダー
 * @returns {Object} 比較結果
 */
function compareHeaders(actualHeaders) {
  const expected = DB_CONFIG.HEADERS;
  
  return {
    expectedCount: expected.length,
    actualCount: actualHeaders.length,
    isMatch: actualHeaders.length === expected.length && 
             actualHeaders.every((header, index) => header === expected[index]),
    missingHeaders: expected.filter(header => !actualHeaders.includes(header)),
    extraHeaders: actualHeaders.filter(header => !expected.includes(header)),
    orderMismatch: actualHeaders.map((header, index) => ({
      position: index,
      actual: header,
      expected: expected[index],
      match: header === expected[index]
    }))
  };
}

/**
 * ユーザーデータの詳細分析
 * @param {Array} headers - ヘッダー配列
 * @param {Array} dataRows - データ行配列
 * @returns {Array} ユーザー分析結果
 */
function analyzeUserData(headers, dataRows) {
  return dataRows.map((row, index) => {
    const user = {};
    const issues = [];
    
    // 各フィールドを分析
    headers.forEach((header, colIndex) => {
      const value = row[colIndex];
      user[header] = value;
      
      // データ検証
      const validation = validateFieldData(header, value);
      if (!validation.isValid) {
        issues.push({
          field: header,
          value: value,
          issue: validation.issue
        });
      }
    });
    
    // configJsonの詳細分析
    const configAnalysis = analyzeConfigJson(user.configJson);
    
    return {
      rowIndex: index + 2, // スプレッドシートの行番号（1ベース + ヘッダー）
      userId: user.userId,
      userEmail: user.userEmail,
      isActive: user.isActive,
      hasSpreadsheetId: !!user.spreadsheetId,
      hasSheetName: !!user.sheetName,
      configJson: configAnalysis,
      issues: issues,
      completenessScore: calculateCompletenessScore(user, headers)
    };
  });
}

/**
 * フィールドデータの検証
 * @param {string} fieldName - フィールド名
 * @param {*} value - 値
 * @returns {Object} 検証結果
 */
function validateFieldData(fieldName, value) {
  // 空値チェック
  const isEmpty = value === undefined || value === null || value === '';
  
  switch (fieldName) {
    case 'userId':
      if (isEmpty) return { isValid: false, issue: 'userId is empty' };
      if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(value)) {
        return { isValid: false, issue: 'Invalid UUID format' };
      }
      break;
      
    case 'userEmail':
      if (isEmpty) return { isValid: false, issue: 'userEmail is empty' };
      if (!/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/.test(value)) {
        return { isValid: false, issue: 'Invalid email format' };
      }
      break;
      
    case 'isActive':
      if (typeof value !== 'boolean' && value !== 'TRUE' && value !== 'FALSE' && value !== true && value !== false) {
        return { isValid: false, issue: 'isActive should be boolean' };
      }
      break;
      
    case 'configJson':
      if (!isEmpty && typeof value === 'string') {
        try {
          JSON.parse(value);
        } catch {
          return { isValid: false, issue: 'Invalid JSON format' };
        }
      }
      break;
      
    case 'columnMappingJson':
      if (!isEmpty && typeof value === 'string') {
        try {
          JSON.parse(value);
        } catch {
          return { isValid: false, issue: 'Invalid JSON format' };
        }
      }
      break;
  }
  
  return { isValid: true };
}

/**
 * configJsonの詳細分析
 * @param {string} configJsonString - configJson文字列
 * @returns {Object} 分析結果
 */
function analyzeConfigJson(configJsonString) {
  if (!configJsonString || configJsonString.trim() === '') {
    return {
      isEmpty: true,
      isValid: false,
      content: null
    };
  }
  
  try {
    const config = JSON.parse(configJsonString);
    
    return {
      isEmpty: false,
      isValid: true,
      content: config,
      keys: Object.keys(config),
      hasAppName: !!config.appName,
      hasDisplayMode: !!config.displayMode,
      hasReactionsEnabled: !!config.reactionsEnabled,
      keyCount: Object.keys(config).length
    };
    
  } catch (error) {
    return {
      isEmpty: false,
      isValid: false,
      error: error.message,
      rawValue: configJsonString
    };
  }
}

/**
 * ユーザーデータの完全性スコアを計算
 * @param {Object} user - ユーザーオブジェクト
 * @param {Array} headers - ヘッダー配列
 * @returns {number} 完全性スコア (0-1)
 */
function calculateCompletenessScore(user, headers) {
  const requiredFields = ['userId', 'userEmail', 'createdAt'];
  const importantFields = ['spreadsheetId', 'sheetName', 'configJson'];
  
  let score = 0;
  let totalWeight = 0;
  
  // 必須フィールド (重み: 3)
  requiredFields.forEach(field => {
    totalWeight += 3;
    if (user[field] && user[field] !== '') {
      score += 3;
    }
  });
  
  // 重要フィールド (重み: 2)
  importantFields.forEach(field => {
    totalWeight += 2;
    if (user[field] && user[field] !== '') {
      score += 2;
    }
  });
  
  // その他フィールド (重み: 1)
  headers.forEach(header => {
    if (!requiredFields.includes(header) && !importantFields.includes(header)) {
      totalWeight += 1;
      if (user[header] && user[header] !== '') {
        score += 1;
      }
    }
  });
  
  return totalWeight > 0 ? score / totalWeight : 0;
}

/**
 * データベース統計の計算
 * @param {Array} headers - ヘッダー配列
 * @param {Array} dataRows - データ行配列
 * @returns {Object} 統計情報
 */
function calculateStatistics(headers, dataRows) {
  if (dataRows.length === 0) {
    return {
      totalUsers: 0,
      activeUsers: 0,
      inactiveUsers: 0
    };
  }
  
  const isActiveIndex = headers.indexOf('isActive');
  const spreadsheetIdIndex = headers.indexOf('spreadsheetId');
  const configJsonIndex = headers.indexOf('configJson');
  
  const stats = {
    totalUsers: dataRows.length,
    activeUsers: 0,
    inactiveUsers: 0,
    usersWithSpreadsheet: 0,
    usersWithoutSpreadsheet: 0,
    usersWithConfig: 0,
    usersWithoutConfig: 0,
    emptyFields: {}
  };
  
  // フィールド別の空値統計を初期化
  headers.forEach(header => {
    stats.emptyFields[header] = 0;
  });
  
  dataRows.forEach(row => {
    // アクティブユーザー統計
    const isActive = row[isActiveIndex];
    if (isActive === true || isActive === 'TRUE') {
      stats.activeUsers++;
    } else {
      stats.inactiveUsers++;
    }
    
    // スプレッドシート統計
    const spreadsheetId = row[spreadsheetIdIndex];
    if (spreadsheetId && spreadsheetId !== '') {
      stats.usersWithSpreadsheet++;
    } else {
      stats.usersWithoutSpreadsheet++;
    }
    
    // 設定統計
    const configJson = row[configJsonIndex];
    if (configJson && configJson !== '') {
      stats.usersWithConfig++;
    } else {
      stats.usersWithoutConfig++;
    }
    
    // フィールド別空値統計
    row.forEach((value, index) => {
      const header = headers[index];
      if (!value || value === '') {
        stats.emptyFields[header]++;
      }
    });
  });
  
  return stats;
}

// 分析実行
if (typeof analyzeDatabaseStructureAndContent === 'function') {
  console.log('🚀 データベース分析スクリプトが読み込まれました');
}