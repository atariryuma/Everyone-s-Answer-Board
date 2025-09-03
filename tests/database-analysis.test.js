/**
 * データベース構造・内容分析テスト
 * CLAUDE.md準拠でデータベースの現状を詳細調査
 */

const path = require('path');
const fs = require('fs');

describe('Database Structure and Content Analysis', () => {
  let gasCode;
  let DB_CONFIG;
  let SYSTEM_CONSTANTS;

  beforeEach(() => {
    // GAS環境のモック設定
    global.console = {
      log: jest.fn(),
      error: jest.fn(),
      warn: jest.fn(),
      info: jest.fn()
    };

    // PropertiesService Mock
    global.PropertiesService = {
      getScriptProperties: () => ({
        getProperty: jest.fn((key) => {
          const mockProperties = {
            'DATABASE_SPREADSHEET_ID': '1ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklm',
            'SERVICE_ACCOUNT_CREDS': JSON.stringify({
              type: "service_account",
              project_id: "test-project"
            })
          };
          return mockProperties[key];
        })
      })
    };

    // Sheets API Mock
    global.batchGetSheetsData = jest.fn();
    global.getServiceAccountTokenCached = jest.fn(() => 'mock_token_123');

    // GASコードを読み込み
    try {
      // constants.gsからSYSTEM_CONSTANTSを読み込み
      const constantsPath = path.join(__dirname, '../src/constants.gs');
      const constantsCode = fs.readFileSync(constantsPath, 'utf8');
      
      // database.gsからDB_CONFIGを読み込み
      const databasePath = path.join(__dirname, '../src/database.gs');
      const databaseCode = fs.readFileSync(databasePath, 'utf8');

      // COREも読み込み
      const corePath = path.join(__dirname, '../src/Core.gs');
      const coreCode = fs.readFileSync(corePath, 'utf8');

      gasCode = constantsCode + '\n' + coreCode + '\n' + databaseCode;
      
      // DB_CONFIGを抽出
      eval(gasCode);
      
      console.log('✅ GASコード読み込み完了');
      
    } catch (error) {
      console.error('❌ GASコード読み込みエラー:', error.message);
      throw error;
    }
  });

  test('DB_CONFIG schema validation', () => {
    console.log('🔍 DB_CONFIGスキーマ検証開始...');
    
    expect(DB_CONFIG).toBeDefined();
    expect(DB_CONFIG.SHEET_NAME).toBe('Users');
    expect(DB_CONFIG.HEADERS).toBeDefined();
    expect(Array.isArray(DB_CONFIG.HEADERS)).toBe(true);
    
    // 期待されるヘッダー数（14列）
    expect(DB_CONFIG.HEADERS.length).toBe(14);
    
    // 必須ヘッダーの確認
    const expectedHeaders = [
      'userId', 'userEmail', 'createdAt', 'lastAccessedAt', 'isActive',
      'spreadsheetId', 'spreadsheetUrl', 'configJson', 'formUrl', 
      'sheetName', 'columnMappingJson', 'publishedAt', 'appUrl', 'lastModified'
    ];
    
    expectedHeaders.forEach((header, index) => {
      expect(DB_CONFIG.HEADERS[index]).toBe(header);
    });
    
    console.log('✅ DB_CONFIG schema validation passed');
    console.log('📋 Headers:', DB_CONFIG.HEADERS);
  });

  test('SYSTEM_CONSTANTS database section validation', () => {
    console.log('🔍 SYSTEM_CONSTANTS.DATABASE検証開始...');
    
    expect(SYSTEM_CONSTANTS).toBeDefined();
    expect(SYSTEM_CONSTANTS.DATABASE).toBeDefined();
    
    // データベース定数の確認
    expect(SYSTEM_CONSTANTS.DATABASE.SHEET_NAME).toBe('Users');
    expect(SYSTEM_CONSTANTS.DATABASE.HEADERS).toBeDefined();
    
    // DB_CONFIGとの一致性確認
    expect(SYSTEM_CONSTANTS.DATABASE.SHEET_NAME).toBe(DB_CONFIG.SHEET_NAME);
    expect(JSON.stringify(SYSTEM_CONSTANTS.DATABASE.HEADERS)).toBe(JSON.stringify(DB_CONFIG.HEADERS));
    
    console.log('✅ SYSTEM_CONSTANTS.DATABASE validation passed');
  });

  test('Mock database content structure analysis', () => {
    console.log('🔍 データベース内容構造分析（モック）開始...');
    
    // モックデータの設定
    const mockSheetData = {
      valueRanges: [{
        values: [
          // ヘッダー行
          DB_CONFIG.HEADERS,
          // サンプルデータ行
          [
            'user-123-uuid',
            'test@example.com',
            '2025-01-01T00:00:00Z',
            '2025-01-02T12:00:00Z',
            'TRUE',
            'sheet-abc-123',
            'https://docs.google.com/spreadsheets/d/sheet-abc-123',
            '{"appName":"テストアプリ","displayMode":"anonymous","reactionsEnabled":true}',
            'https://docs.google.com/forms/d/form-123',
            'フォームの回答 1',
            '{"answer":{"header":"回答","index":2},"reason":{"header":"理由","index":3}}',
            '2025-01-01T10:00:00Z',
            'https://script.google.com/macros/s/app-123',
            '2025-01-02T08:00:00Z'
          ],
          // 空のフィールドを含むサンプル
          [
            'user-456-uuid',
            'test2@example.com',
            '2025-01-01T01:00:00Z',
            '',
            'FALSE',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            ''
          ]
        ]
      }]
    };
    
    global.batchGetSheetsData.mockReturnValue(mockSheetData);
    
    // 分析実行
    const service = global.getServiceAccountTokenCached();
    const dbId = PropertiesService.getScriptProperties().getProperty('DATABASE_SPREADSHEET_ID');
    const sheetName = DB_CONFIG.SHEET_NAME;
    
    const data = batchGetSheetsData(service, dbId, [`'${sheetName}'!A:N`]);
    const values = data.valueRanges[0].values || [];
    
    // 基本統計の検証
    expect(values.length).toBe(3); // ヘッダー + 2データ行
    expect(values[0]).toEqual(DB_CONFIG.HEADERS);
    
    // データ行分析
    const dataRows = values.slice(1);
    expect(dataRows.length).toBe(2);
    
    // 1行目のデータ検証（完全なデータ）
    const completeUser = dataRows[0];
    expect(completeUser[0]).toBe('user-123-uuid'); // userId
    expect(completeUser[1]).toBe('test@example.com'); // userEmail
    expect(completeUser[4]).toBe('TRUE'); // isActive
    expect(completeUser[5]).toBe('sheet-abc-123'); // spreadsheetId
    
    // configJson検証
    const configJson = completeUser[8];
    expect(configJson).toBeTruthy();
    expect(() => JSON.parse(configJson)).not.toThrow();
    
    const config = JSON.parse(configJson);
    expect(config.appName).toBe('テストアプリ');
    expect(config.displayMode).toBe('anonymous');
    expect(config.reactionsEnabled).toBe(true);
    
    // 2行目のデータ検証（空フィールド多数）
    const incompleteUser = dataRows[1];
    expect(incompleteUser[0]).toBe('user-456-uuid'); // userId
    expect(incompleteUser[1]).toBe('test2@example.com'); // userEmail
    expect(incompleteUser[3]).toBe(''); // lastAccessedAt - 空
    expect(incompleteUser[5]).toBe(''); // spreadsheetId - 空
    expect(incompleteUser[8]).toBe(''); // configJson - 空
    
    // 統計計算
    let activeCount = 0;
    let hasSpreadsheetCount = 0;
    let hasConfigCount = 0;
    
    dataRows.forEach(row => {
      const isActive = row[4];
      const spreadsheetId = row[5];
      const configJson = row[8];
      
      if (isActive === 'TRUE' || isActive === true) activeCount++;
      if (spreadsheetId && spreadsheetId !== '') hasSpreadsheetCount++;
      if (configJson && configJson !== '') hasConfigCount++;
    });
    
    expect(activeCount).toBe(1);
    expect(hasSpreadsheetCount).toBe(1);
    expect(hasConfigCount).toBe(1);
    
    console.log('✅ データベース内容構造分析完了');
    console.log('📊 統計:');
    console.log(`  - 総ユーザー数: ${dataRows.length}`);
    console.log(`  - アクティブユーザー: ${activeCount}`);
    console.log(`  - スプレッドシート設定済み: ${hasSpreadsheetCount}`);
    console.log(`  - 設定JSON保持: ${hasConfigCount}`);
  });

  test('Data validation patterns', () => {
    console.log('🔍 データ検証パターンテスト開始...');
    
    // SecurityValidator相当の検証ロジックをテスト
    const validationTests = [
      {
        field: 'userId',
        validValues: ['12345678-1234-1234-1234-123456789abc', 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'],
        invalidValues: ['not-a-uuid', '123', '', null]
      },
      {
        field: 'userEmail',
        validValues: ['test@example.com', 'user.name+tag@domain.co.jp'],
        invalidValues: ['invalid-email', '@example.com', 'test@', '']
      },
      {
        field: 'isActive',
        validValues: [true, false, 'TRUE', 'FALSE'],
        invalidValues: ['yes', 'no', '1', '0', '']
      }
    ];
    
    validationTests.forEach(({ field, validValues, invalidValues }) => {
      console.log(`🧪 Testing ${field} validation...`);
      
      validValues.forEach(value => {
        const result = validateFieldValue(field, value);
        expect(result.isValid).toBe(true);
      });
      
      invalidValues.forEach(value => {
        const result = validateFieldValue(field, value);
        expect(result.isValid).toBe(false);
      });
    });
    
    console.log('✅ データ検証パターンテスト完了');
  });

  test('ConfigJson analysis patterns', () => {
    console.log('🔍 configJson分析パターンテスト開始...');
    
    const configSamples = [
      // 完全な設定
      {
        json: '{"appName":"完全なアプリ","displayMode":"anonymous","reactionsEnabled":true,"additionalSettings":{"theme":"dark"}}',
        expected: {
          isValid: true,
          hasAppName: true,
          hasDisplayMode: true,
          hasReactionsEnabled: true,
          keyCount: 4
        }
      },
      // 最小設定
      {
        json: '{"appName":"最小アプリ"}',
        expected: {
          isValid: true,
          hasAppName: true,
          hasDisplayMode: false,
          hasReactionsEnabled: false,
          keyCount: 1
        }
      },
      // 無効なJSON
      {
        json: '{"appName":"無効JSON",}',
        expected: {
          isValid: false
        }
      },
      // 空文字列
      {
        json: '',
        expected: {
          isEmpty: true,
          isValid: false
        }
      }
    ];
    
    configSamples.forEach(({ json, expected }, index) => {
      console.log(`🧪 Testing config sample ${index + 1}...`);
      
      const analysis = analyzeConfigJson(json);
      
      if (expected.isEmpty) {
        expect(analysis.isEmpty).toBe(true);
        expect(analysis.isValid).toBe(false);
      } else if (expected.isValid) {
        expect(analysis.isValid).toBe(true);
        expect(analysis.hasAppName).toBe(expected.hasAppName);
        expect(analysis.hasDisplayMode).toBe(expected.hasDisplayMode);
        expect(analysis.hasReactionsEnabled).toBe(expected.hasReactionsEnabled);
        expect(analysis.keyCount).toBe(expected.keyCount);
      } else {
        expect(analysis.isValid).toBe(false);
      }
    });
    
    console.log('✅ configJson分析パターンテスト完了');
  });
});

// ヘルパー関数: フィールド値検証
function validateFieldValue(fieldName, value) {
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
  }
  
  return { isValid: true };
}

// ヘルパー関数: configJson分析
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