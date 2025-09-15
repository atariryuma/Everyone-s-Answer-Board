#!/usr/bin/env node

/**
 * Refactor the large analyzeColumns function into smaller, manageable functions
 * Following GAS best practices of keeping functions under 100 lines
 */

const fs = require('fs');
const path = require('path');

function generateRefactoredFunctions() {
  return `
// ===========================================
// 📊 Column Analysis - Refactored Functions
// ===========================================

/**
 * 列分析のメイン関数（リファクタリング版）
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} 列分析結果
 */
function analyzeColumns(spreadsheetId, sheetName) {
  const started = Date.now();
  try {
    console.log('DataService.analyzeColumns: 開始', {
      spreadsheetId: spreadsheetId ? \`\${spreadsheetId.substring(0, 10)}...\` : 'null',
      sheetName: sheetName || 'null'
    });

    // 🎯 GAS Best Practice: パラメータ検証を別関数に分離
    const paramValidation = validateAnalyzeColumnsParams(spreadsheetId, sheetName);
    if (!paramValidation.isValid) {
      return paramValidation.errorResponse;
    }

    // 🎯 GAS Best Practice: スプレッドシート接続を別関数に分離
    const connectionResult = connectToSpreadsheet(spreadsheetId, sheetName);
    if (!connectionResult.success) {
      return connectionResult.errorResponse;
    }

    // 🎯 GAS Best Practice: データ取得を別関数に分離
    const dataResult = extractSheetData(connectionResult.sheet);
    if (!dataResult.success) {
      return dataResult.errorResponse;
    }

    // 🎯 GAS Best Practice: 列分析を別関数に分離
    const analysisResult = performColumnAnalysis(dataResult.headers, dataResult.sampleData);

    return {
      success: true,
      headers: dataResult.headers,
      columns: analysisResult.columns,
      columnMapping: analysisResult.mapping,
      sampleData: dataResult.sampleData.slice(0, 3),
      executionTime: \`\${Date.now() - started}ms\`
    };

  } catch (error) {
    console.error('DataService.analyzeColumns: 予期しないエラー', {
      error: error.message,
      executionTime: \`\${Date.now() - started}ms\`
    });

    return {
      success: false,
      message: \`予期しないエラーが発生しました: \${error.message}\`,
      headers: [],
      columns: [],
      columnMapping: { mapping: {}, confidence: {} },
      executionTime: \`\${Date.now() - started}ms\`
    };
  }
}

/**
 * パラメータ検証（GAS Best Practice: 単一責任）
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} 検証結果
 */
function validateAnalyzeColumnsParams(spreadsheetId, sheetName) {
  if (!spreadsheetId || !sheetName) {
    const errorResponse = {
      success: false,
      message: 'スプレッドシートIDとシート名が必要です',
      headers: [],
      columns: [],
      columnMapping: { mapping: {}, confidence: {} }
    };
    console.error('DataService.analyzeColumns: 必須パラメータ不足', {
      spreadsheetId: !!spreadsheetId,
      sheetName: !!sheetName
    });
    return { isValid: false, errorResponse };
  }

  return { isValid: true };
}

/**
 * スプレッドシート接続（GAS Best Practice: エラーハンドリング分離）
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} 接続結果
 */
function connectToSpreadsheet(spreadsheetId, sheetName) {
  try {
    console.log('DataService.analyzeColumns: スプレッドシート接続開始');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    console.log('DataService.analyzeColumns: スプレッドシート接続成功');

    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      return {
        success: false,
        errorResponse: {
          success: false,
          message: \`シート "\${sheetName}" が見つかりません\`,
          headers: [],
          columns: [],
          columnMapping: { mapping: {}, confidence: {} }
        }
      };
    }

    console.log('DataService.analyzeColumns: シート取得成功');
    return { success: true, sheet };

  } catch (error) {
    console.error('DataService.analyzeColumns: 接続エラー', {
      error: error.message,
      spreadsheetId: \`\${spreadsheetId.substring(0, 10)}...\`
    });

    return {
      success: false,
      errorResponse: {
        success: false,
        message: \`スプレッドシートアクセスエラー: \${error.message}\`,
        headers: [],
        columns: [],
        columnMapping: { mapping: {}, confidence: {} }
      }
    };
  }
}

/**
 * シートデータ抽出（GAS Best Practice: データアクセス最適化）
 * @param {Sheet} sheet - シートオブジェクト
 * @returns {Object} データ抽出結果
 */
function extractSheetData(sheet) {
  try {
    console.log('DataService.analyzeColumns: シートサイズ取得開始');
    const lastColumn = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();

    console.log('DataService.analyzeColumns: シートサイズ', { lastColumn, lastRow });

    if (lastColumn === 0 || lastRow === 0) {
      return {
        success: false,
        errorResponse: {
          success: false,
          message: 'スプレッドシートが空です',
          headers: [],
          columns: [],
          columnMapping: { mapping: {}, confidence: {} }
        }
      };
    }

    // 🎯 GAS Best Practice: バッチでデータ取得
    console.log('DataService.analyzeColumns: ヘッダー取得開始');
    const [headers] = sheet.getRange(1, 1, 1, lastColumn).getValues();
    console.log('DataService.analyzeColumns: ヘッダー取得成功', { headersCount: headers.length });

    // サンプルデータを取得（最大5行）
    let sampleData = [];
    const sampleRowCount = Math.min(5, lastRow - 1);
    if (sampleRowCount > 0) {
      console.log('DataService.analyzeColumns: サンプルデータ取得開始');
      sampleData = sheet.getRange(2, 1, sampleRowCount, lastColumn).getValues();
      console.log('DataService.analyzeColumns: サンプルデータ取得成功', {
        sampleRowCount: sampleData.length
      });
    }

    return { success: true, headers, sampleData };

  } catch (error) {
    console.error('DataService.analyzeColumns: データ取得エラー', {
      error: error.message
    });

    return {
      success: false,
      errorResponse: {
        success: false,
        message: \`データ取得エラー: \${error.message}\`,
        headers: [],
        columns: [],
        columnMapping: { mapping: {}, confidence: {} }
      }
    };
  }
}

/**
 * 列分析実行（GAS Best Practice: ビジネスロジック分離）
 * @param {Array} headers - ヘッダー配列
 * @param {Array} sampleData - サンプルデータ配列
 * @returns {Object} 分析結果
 */
function performColumnAnalysis(headers, sampleData) {
  // 列情報を分析
  const columns = headers.map((header, index) => {
    const samples = sampleData.map(row => row[index]).filter(v => v);

    // 列タイプを推測
    let type = 'text';
    if (header.toLowerCase().includes('timestamp') || header.toLowerCase().includes('日時')) {
      type = 'datetime';
    } else if (header.toLowerCase().includes('email') || header.toLowerCase().includes('メール')) {
      type = 'email';
    } else if (header.toLowerCase().includes('class') || header.toLowerCase().includes('クラス')) {
      type = 'class';
    } else if (header.toLowerCase().includes('name') || header.toLowerCase().includes('名前')) {
      type = 'name';
    } else if (samples.length > 0 && samples.every(s => !isNaN(s))) {
      type = 'number';
    }

    return {
      index,
      header,
      type,
      samples: samples.slice(0, 3) // 最大3つのサンプル
    };
  });

  // AI検出シミュレーション（シンプルなパターンマッチング）
  const mapping = { mapping: {}, confidence: {} };

  headers.forEach((header, index) => {
    const headerLower = header.toLowerCase();

    // 回答列の検出
    if (headerLower.includes('回答') || headerLower.includes('answer') || headerLower.includes('意見')) {
      mapping.mapping.answer = index;
      mapping.confidence.answer = 85;
    }

    // 理由列の検出
    if (headerLower.includes('理由') || headerLower.includes('根拠') || headerLower.includes('reason')) {
      mapping.mapping.reason = index;
      mapping.confidence.reason = 80;
    }

    // クラス列の検出
    if (headerLower.includes('クラス') || headerLower.includes('class') || headerLower.includes('組')) {
      mapping.mapping.class = index;
      mapping.confidence.class = 90;
    }

    // 名前列の検出
    if (headerLower.includes('名前') || headerLower.includes('name') || headerLower.includes('氏名')) {
      mapping.mapping.name = index;
      mapping.confidence.name = 85;
    }
  });

  console.log('DataService.analyzeColumns: 分析完了', {
    headersCount: headers.length,
    columnsCount: columns.length,
    mappingDetected: Object.keys(mapping.mapping).length
  });

  return { columns, mapping };
}
`;
}

function main() {
  const filePath = path.join(__dirname, '../src/DataService.gs');

  try {
    let content = fs.readFileSync(filePath, 'utf8');
    console.log(`📄 Processing: ${filePath}`);

    // Find and replace the large analyzeColumns function
    const functionStartRegex = /function analyzeColumns\(spreadsheetId, sheetName\) \{/;
    const match = content.match(functionStartRegex);

    if (!match) {
      console.error('❌ analyzeColumns function not found');
      process.exit(1);
    }

    const startIndex = match.index;
    let braceCount = 0;
    let endIndex = startIndex;
    let inFunction = false;

    // Find the complete function body
    for (let i = startIndex; i < content.length; i++) {
      if (content[i] === '{') {
        braceCount++;
        inFunction = true;
      } else if (content[i] === '}' && inFunction) {
        braceCount--;
        if (braceCount === 0) {
          endIndex = i + 1;
          break;
        }
      }
    }

    // Replace with refactored functions
    const before = content.substring(0, startIndex);
    const after = content.substring(endIndex);
    const refactoredFunctions = generateRefactoredFunctions();

    const newContent = before + refactoredFunctions + after;

    fs.writeFileSync(filePath, newContent, 'utf8');

    console.log(`✅ Successfully refactored analyzeColumns function`);
    console.log(`📊 Original function: ~232 lines`);
    console.log(`📊 New functions: 5 functions, each under 100 lines`);
    console.log(`🎯 GAS Best Practice: Function length compliance achieved`);

  } catch (error) {
    console.error('❌ Error processing file:', error.message);
    process.exit(1);
  }
}

if (require.main === module) {
  main();
}