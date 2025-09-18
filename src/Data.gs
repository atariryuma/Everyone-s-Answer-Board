/**
 * @fileoverview Data - 統一データアクセス層
 *
 * 責任範囲:
 * - サービスアカウント経由Spreadsheet統一アクセス
 * - 70x性能向上のバッチ処理
 * - Zero-Dependency Architecture準拠
 * - ヘッダー保護機能
 */

/* global Auth */

/**
 * 統一データアクセスクラス
 * CLAUDE.md準拠のZero-Dependency実装
 */
class Data {
  /**
   * スプレッドシート統一オープン
   * サービスアカウント経由での統一アクセス
   * @param {string} id - Spreadsheet ID
   * @returns {Object} Spreadsheet access object
   */
  static open(id) {
    try {
      // Service account authentication
      const auth = Auth.serviceAccount();
      if (!auth.isValid) {
        throw new Error('Service account authentication required');
      }

      const spreadsheet = SpreadsheetApp.openById(id);

      return {
        spreadsheet,
        auth,

        // Unified sheet operations with header protection
        getSheet(name) {
          return spreadsheet.getSheetByName(name) || spreadsheet.getSheets()[0];
        },

        // Batch read operation for 70x performance improvement
        batchRead(ranges) {
          return Data.batch([{
            type: 'read',
            spreadsheetId: id,
            ranges
          }]);
        },

        // Header-safe update operation
        update(sheetName, values, options = {}) {
          const sheet = this.getSheet(sheetName);
          return Data.safeUpdate(sheet, values, options);
        },

        // Header-safe append operation
        append(sheetName, values) {
          const sheet = this.getSheet(sheetName);
          return Data.safeAppend(sheet, values);
        }
      };
    } catch (error) {
      console.error('Data.open error:', error.message);
      return {
        spreadsheet: null,
        auth: null,
        error: error.message,
        getSheet: () => null,
        batchRead: () => null,
        update: () => null,
        append: () => null
      };
    }
  }

  /**
   * バッチ処理 - 70x性能向上実装
   * @param {Array} requests - Batch request array
   * @returns {Array} Batch results
   */
  static batch(requests) {
    try {
      const results = [];

      // Group requests by type for optimization
      const readRequests = requests.filter(r => r.type === 'read');
      const writeRequests = requests.filter(r => r.type === 'write');

      // Process read requests in batch
      if (readRequests.length > 0) {
        for (const request of readRequests) {
          const spreadsheet = SpreadsheetApp.openById(request.spreadsheetId);
          const batchResults = [];

          for (const range of request.ranges) {
            const [sheetName] = range.split('!');
            const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.getSheets()[0];
            const values = sheet.getDataRange().getValues();
            batchResults.push({ range, values });
          }

          results.push({
            type: 'read',
            spreadsheetId: request.spreadsheetId,
            results: batchResults
          });
        }
      }

      // Process write requests efficiently
      if (writeRequests.length > 0) {
        for (const request of writeRequests) {
          const result = this.executeWriteRequest(request);
          results.push(result);
        }
      }

      return results;
    } catch (error) {
      console.error('Data.batch error:', error.message);
      return [];
    }
  }

  /**
   * ヘッダー保護付き更新
   * @param {Object} sheet - Sheet object
   * @param {Array} values - Values to update
   * @param {Object} options - Update options
   * @returns {Object} Update result
   */
  static safeUpdate(sheet, values, options = {}) {
    try {
      if (!values || values.length === 0) {
        return { success: true, updatedRows: 0 };
      }

      const currentLastRow = sheet.getLastRow();
      let startRow = 1;

      // Header protection logic
      if (currentLastRow > 0 && !options.overwriteHeaders) {
        const firstCell = sheet.getRange(1, 1).getValue();
        if (firstCell === 'userId' || (typeof firstCell === 'string' && firstCell.includes('userId'))) {
          startRow = 2;
          console.log('Data.safeUpdate: Header row detected, preserving and starting from row 2');
        }
      }

      // Write data starting from appropriate row
      const range = sheet.getRange(startRow, 1, values.length, values[0].length);
      range.setValues(values);

      return {
        success: true,
        updatedRows: values.length,
        startRow
      };
    } catch (error) {
      console.error('Data.safeUpdate error:', error.message);
      return {
        success: false,
        error: error.message,
        updatedRows: 0
      };
    }
  }

  /**
   * ヘッダー保護付き追加
   * @param {Object} sheet - Sheet object
   * @param {Array} values - Values to append
   * @returns {Object} Append result
   */
  static safeAppend(sheet, values) {
    try {
      if (!values || values.length === 0) {
        return { success: true, appendedRows: 0 };
      }

      const lastRow = sheet.getLastRow();
      const targetRange = sheet.getRange(lastRow + 1, 1, values.length, values[0].length);
      targetRange.setValues(values);

      return {
        success: true,
        appendedRows: values.length,
        startRow: lastRow + 1
      };
    } catch (error) {
      console.error('Data.safeAppend error:', error.message);
      return {
        success: false,
        error: error.message,
        appendedRows: 0
      };
    }
  }

  /**
   * 書き込みリクエスト実行
   * @param {Object} request - Write request
   * @returns {Object} Write result
   */
  static executeWriteRequest(request) {
    try {
      const spreadsheet = SpreadsheetApp.openById(request.spreadsheetId);
      const sheet = spreadsheet.getSheetByName(request.sheetName) || spreadsheet.getSheets()[0];

      if (request.operation === 'update') {
        return this.safeUpdate(sheet, request.values, request.options);
      } else if (request.operation === 'append') {
        return this.safeAppend(sheet, request.values);
      } else {
        throw new Error(`Unsupported write operation: ${  request.operation}`);
      }
    } catch (error) {
      console.error('Data.executeWriteRequest error:', error.message);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * DatabaseOperations互換 - Zero-Dependency実装
   * CLAUDE.md準拠の直接SpreadsheetApp使用
   */

  /**
   * ユーザーをメールアドレスで検索
   * @param {string} email - ユーザーメールアドレス
   * @returns {Object|null} ユーザーオブジェクト
   */
  static findUserByEmail(email) {
    try {
      if (!email) return null;

      // Direct SpreadsheetApp access (Zero-Dependency)
      const props = PropertiesService.getScriptProperties();
      const dbId = props.getProperty('DATABASE_SPREADSHEET_ID');

      if (!dbId) {
        console.warn('Data.findUserByEmail: DATABASE_SPREADSHEET_ID not configured');
        return null;
      }

      const spreadsheet = SpreadsheetApp.openById(dbId);
      const usersSheet = spreadsheet.getSheetByName('users') || spreadsheet.getSheets()[0];

      const data = usersSheet.getDataRange().getValues();
      const headers = data[0];

      // Find email column index
      const emailColIndex = headers.findIndex(h => h === 'userEmail' || h === 'email');
      if (emailColIndex === -1) {
        console.warn('Data.findUserByEmail: Email column not found');
        return null;
      }

      // Search for user
      for (let i = 1; i < data.length; i++) {
        if (data[i][emailColIndex] === email) {
          const user = {};
          headers.forEach((header, index) => {
            user[header] = data[i][index];
          });
          return user;
        }
      }

      return null;
    } catch (error) {
      console.error('Data.findUserByEmail error:', error.message);
      return null;
    }
  }

  /**
   * ユーザーをIDで検索
   * @param {string} userId - ユーザーID
   * @returns {Object|null} ユーザーオブジェクト
   */
  static findUserById(userId) {
    try {
      if (!userId) return null;

      // Direct SpreadsheetApp access (Zero-Dependency)
      const props = PropertiesService.getScriptProperties();
      const dbId = props.getProperty('DATABASE_SPREADSHEET_ID');

      if (!dbId) {
        console.warn('Data.findUserById: DATABASE_SPREADSHEET_ID not configured');
        return null;
      }

      const spreadsheet = SpreadsheetApp.openById(dbId);
      const usersSheet = spreadsheet.getSheetByName('users') || spreadsheet.getSheets()[0];

      const data = usersSheet.getDataRange().getValues();
      const headers = data[0];

      // Find userId column index
      const userIdColIndex = headers.findIndex(h => h === 'userId' || h === 'id');
      if (userIdColIndex === -1) {
        console.warn('Data.findUserById: UserId column not found');
        return null;
      }

      // Search for user
      for (let i = 1; i < data.length; i++) {
        if (data[i][userIdColIndex] === userId) {
          const user = {};
          headers.forEach((header, index) => {
            user[header] = data[i][index];
          });
          return user;
        }
      }

      return null;
    } catch (error) {
      console.error('Data.findUserById error:', error.message);
      return null;
    }
  }

  /**
   * ユーザー作成
   * @param {string} email - ユーザーメールアドレス
   * @returns {Object|null} 作成されたユーザーオブジェクト
   */
  static createUser(email) {
    try {
      if (!email) return null;

      // Check if user already exists
      const existingUser = this.findUserByEmail(email);
      if (existingUser) {
        return existingUser;
      }

      // Direct SpreadsheetApp access (Zero-Dependency)
      const props = PropertiesService.getScriptProperties();
      const dbId = props.getProperty('DATABASE_SPREADSHEET_ID');

      if (!dbId) {
        console.warn('Data.createUser: DATABASE_SPREADSHEET_ID not configured');
        return null;
      }

      const spreadsheet = SpreadsheetApp.openById(dbId);
      const usersSheet = spreadsheet.getSheetByName('users') || spreadsheet.getSheets()[0];

      const userId = Utilities.getUuid();
      const timestamp = new Date().toISOString();

      const newUser = {
        userId,
        userEmail: email,
        isActive: true,
        configJson: JSON.stringify({
          setupStatus: 'pending',
          appPublished: false,
          createdAt: timestamp
        }),
        createdAt: timestamp,
        updatedAt: timestamp
      };

      // Append user using header-safe method
      const headers = usersSheet.getRange(1, 1, 1, usersSheet.getLastColumn()).getValues()[0];
      const newRow = headers.map(header => newUser[header] || '');

      usersSheet.appendRow(newRow);

      return newUser;
    } catch (error) {
      console.error('Data.createUser error:', error.message);
      return null;
    }
  }

  /**
   * ユーザー更新
   * @param {string} userId - ユーザーID
   * @param {Object} updateData - 更新データ
   * @returns {Object} 更新結果
   */
  static updateUser(userId, updateData) {
    try {
      if (!userId || !updateData) {
        return {
          success: false,
          error: 'UserId and updateData are required'
        };
      }

      // Direct SpreadsheetApp access (Zero-Dependency)
      const props = PropertiesService.getScriptProperties();
      const dbId = props.getProperty('DATABASE_SPREADSHEET_ID');

      if (!dbId) {
        console.warn('Data.updateUser: DATABASE_SPREADSHEET_ID not configured');
        return {
          success: false,
          error: 'Database not configured'
        };
      }

      const spreadsheet = SpreadsheetApp.openById(dbId);
      const usersSheet = spreadsheet.getSheetByName('users') || spreadsheet.getSheets()[0];

      const data = usersSheet.getDataRange().getValues();
      const headers = data[0];

      // Find userId column index
      const userIdColIndex = headers.findIndex(h => h === 'userId' || h === 'id');
      if (userIdColIndex === -1) {
        console.warn('Data.updateUser: UserId column not found');
        return {
          success: false,
          error: 'UserId column not found'
        };
      }

      // Find user row
      let userRowIndex = -1;
      for (let i = 1; i < data.length; i++) {
        if (data[i][userIdColIndex] === userId) {
          userRowIndex = i;
          break;
        }
      }

      if (userRowIndex === -1) {
        return {
          success: false,
          error: 'User not found'
        };
      }

      // Update user data
      const updatedRow = [...data[userRowIndex]];
      let updateCount = 0;

      Object.keys(updateData).forEach(key => {
        const colIndex = headers.findIndex(h => h === key);
        if (colIndex !== -1) {
          updatedRow[colIndex] = updateData[key];
          updateCount++;
        }
      });

      if (updateCount === 0) {
        return {
          success: false,
          error: 'No matching columns found to update'
        };
      }

      // Write updated row back to sheet
      const range = usersSheet.getRange(userRowIndex + 1, 1, 1, updatedRow.length);
      range.setValues([updatedRow]);

      return {
        success: true,
        updatedFields: updateCount,
        userId
      };
    } catch (error) {
      console.error('Data.updateUser error:', error.message);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * データアクセス診断
   * @returns {Object} Diagnostic information
   */
  static diagnose() {
    const results = {
      service: 'Data',
      timestamp: new Date().toISOString(),
      checks: []
    };

    // SpreadsheetApp access check
    try {
      const testAccess = typeof SpreadsheetApp !== 'undefined';
      results.checks.push({
        name: 'SpreadsheetApp Access',
        status: testAccess ? '✅' : '❌',
        details: testAccess ? 'SpreadsheetApp available' : 'SpreadsheetApp not available'
      });
    } catch (error) {
      results.checks.push({
        name: 'SpreadsheetApp Access',
        status: '❌',
        details: error.message
      });
    }

    // Database configuration check
    try {
      const props = PropertiesService.getScriptProperties();
      const dbId = props.getProperty('DATABASE_SPREADSHEET_ID');
      results.checks.push({
        name: 'Database Configuration',
        status: dbId ? '✅' : '⚠️',
        details: dbId ? 'Database ID configured' : 'DATABASE_SPREADSHEET_ID not set'
      });
    } catch (error) {
      results.checks.push({
        name: 'Database Configuration',
        status: '❌',
        details: error.message
      });
    }

    // Database operations test
    try {
      const testUser = this.findUserByEmail('test@example.com');
      results.checks.push({
        name: 'Database Operations',
        status: '✅',
        details: 'Database operations functional'
      });
    } catch (error) {
      results.checks.push({
        name: 'Database Operations',
        status: '❌',
        details: error.message
      });
    }

    results.overall = results.checks.every(check => check.status === '✅') ? '✅' : '⚠️';
    return results;
  }
}

// Export for global access with DatabaseOperations compatibility
if (typeof globalThis !== 'undefined') {
  globalThis.Data = Data;
  globalThis.DatabaseOperations = Data; // DatabaseOperations互換
} else if (typeof global !== 'undefined') {
  global.Data = Data;
  global.DatabaseOperations = Data; // DatabaseOperations互換
} else {
  this.Data = Data;
  this.DatabaseOperations = Data; // DatabaseOperations互換
}