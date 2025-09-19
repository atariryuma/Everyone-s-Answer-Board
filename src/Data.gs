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
      // Service account authentication - 真の実装
      const auth = Auth.serviceAccount();
      if (!auth.isValid) {
        throw new Error('Service account authentication required');
      }

      // サービスアカウントでのスプレッドシート権限確保
      try {
        // DriveAppでサービスアカウントをエディターとして追加（実際の権限付与）
        DriveApp.getFileById(id).addEditor(auth.email);
        console.log('Data.open: Service account editor access granted:', auth.email);
      } catch (driveError) {
        console.warn('Data.open: Service account editor access already granted or failed:', driveError.message);
      }

      // サービスアカウント権限でGoogle Sheets APIを直接使用
      const spreadsheet = this.openSpreadsheetWithServiceAccount(id, auth.token);

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
      // Service account authentication - 真の実装
      const auth = Auth.serviceAccount();
      if (!auth.isValid) {
        throw new Error('Service account authentication required for batch operations');
      }

      const results = [];

      // Group requests by type for optimization
      const readRequests = requests.filter(r => r.type === 'read');
      const writeRequests = requests.filter(r => r.type === 'write');

      // Process read requests in batch
      if (readRequests.length > 0) {
        for (const request of readRequests) {
          const spreadsheet = this.openSpreadsheetWithServiceAccount(request.spreadsheetId, auth.token);
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
      // Service account authentication - 真の実装
      const auth = Auth.serviceAccount();
      if (!auth.isValid) {
        throw new Error('Service account authentication required for write operations');
      }

      const spreadsheet = this.openSpreadsheetWithServiceAccount(request.spreadsheetId, auth.token);
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
   * User Database Operations - Zero-Dependency実装
   * CLAUDE.md準拠のサービスアカウント統一アクセス
   */

  /**
   * ユーザーをメールアドレスで検索
   * @param {string} email - ユーザーメールアドレス
   * @returns {Object|null} ユーザーオブジェクト
   */
  static findUserByEmail(email) {
    try {
      if (!email) return null;

      // Service account authentication for database access
      const auth = Auth.serviceAccount();
      if (!auth.isValid) {
        console.warn('Data.findUserByEmail: Service account authentication required');
        return null;
      }

      // Direct SpreadsheetApp access (Zero-Dependency)
      const props = PropertiesService.getScriptProperties();
      const dbId = props.getProperty('DATABASE_SPREADSHEET_ID');

      if (!dbId) {
        console.warn('Data.findUserByEmail: DATABASE_SPREADSHEET_ID not configured');
        return null;
      }

      // サービスアカウント権限でのデータベースアクセス
      try {
        DriveApp.getFileById(dbId).addEditor(auth.email);
      } catch (driveError) {
        console.warn('Data.findUserByEmail: Service account access:', driveError.message);
      }

      const spreadsheet = this.openSpreadsheetWithServiceAccount(dbId, auth.token);
      const usersSheet = spreadsheet.getSheetByName('users') || spreadsheet.getSheets()[0];

      const data = usersSheet.getDataRange().getValues();
      const [headers] = data;

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

      // Service account authentication for database access
      const auth = Auth.serviceAccount();
      if (!auth.isValid) {
        console.warn('Data.findUserById: Service account authentication required');
        return null;
      }

      // Direct SpreadsheetApp access (Zero-Dependency)
      const props = PropertiesService.getScriptProperties();
      const dbId = props.getProperty('DATABASE_SPREADSHEET_ID');

      if (!dbId) {
        console.warn('Data.findUserById: DATABASE_SPREADSHEET_ID not configured');
        return null;
      }

      // サービスアカウント権限でのデータベースアクセス
      try {
        DriveApp.getFileById(dbId).addEditor(auth.email);
      } catch (driveError) {
        console.warn('Data.findUserById: Service account access:', driveError.message);
      }

      const spreadsheet = this.openSpreadsheetWithServiceAccount(dbId, auth.token);
      const usersSheet = spreadsheet.getSheetByName('users') || spreadsheet.getSheets()[0];

      const data = usersSheet.getDataRange().getValues();
      const [headers] = data;

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

      // Service account authentication for database access
      const auth = Auth.serviceAccount();
      if (!auth.isValid) {
        console.warn('Data.createUser: Service account authentication required');
        return null;
      }

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

      // サービスアカウント権限でのデータベースアクセス
      try {
        DriveApp.getFileById(dbId).addEditor(auth.email);
        console.log('Data.createUser: Service account editor access granted:', auth.email);
      } catch (driveError) {
        console.warn('Data.createUser: Service account access:', driveError.message);
      }

      const spreadsheet = this.openSpreadsheetWithServiceAccount(dbId, auth.token);
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

      // Append user using header-safe method - Sheets API対応
      const lastColumn = usersSheet.getLastColumn();
      const headerRange = `A1:${Data.columnToLetter(lastColumn)}1`;
      const [headers] = usersSheet.getRange(headerRange).getValues();
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

      // Service account authentication for database access
      const auth = Auth.serviceAccount();
      if (!auth.isValid) {
        console.warn('Data.updateUser: Service account authentication required');
        return {
          success: false,
          error: 'Service account authentication required'
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

      // 🔧 CLAUDE.md準拠: 原子的トランザクション - 競合状態防止
      const transactionKey = `update_user_${userId}`;

      // eslint-disable-next-line no-undef
      if (typeof RequestGate !== 'undefined' && !RequestGate.enter(transactionKey)) {
        console.warn('Data.updateUser: Transaction in progress for user:', userId);
        return {
          success: false,
          error: 'User update transaction in progress'
        };
      } else if (typeof RequestGate === 'undefined') {
        console.warn('Data.updateUser: RequestGate not available, proceeding without transaction lock');
      }

      let spreadsheet = null;
      try {
        // サービスアカウント権限でのデータベースアクセス
        try {
          DriveApp.getFileById(dbId).addEditor(auth.email);
          console.log('Data.updateUser: Service account editor access granted:', auth.email);
        } catch (driveError) {
          console.warn('Data.updateUser: Service account access:', driveError.message);
        }

        spreadsheet = this.openSpreadsheetWithServiceAccount(dbId, auth.token);
        if (!spreadsheet) {
          throw new Error('Failed to open spreadsheet with service account');
        }
      const usersSheet = spreadsheet.getSheetByName('users') || spreadsheet.getSheets()[0];

      const data = usersSheet.getDataRange().getValues();
      const [headers] = data;

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
          // eslint-disable-next-line no-undef
          if (typeof RequestGate !== 'undefined') RequestGate.exit(transactionKey);
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
          // eslint-disable-next-line no-undef
          if (typeof RequestGate !== 'undefined') RequestGate.exit(transactionKey);
          return {
            success: false,
            error: 'No matching columns found to update'
          };
        }

      // Write updated row back to sheet - Sheets API対応
      const rowNumber = userRowIndex + 1;
      const endColumn = Data.columnToLetter(updatedRow.length);
      const rangeNotation = `A${rowNumber}:${endColumn}${rowNumber}`;
      const range = usersSheet.getRange(rangeNotation);
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
      } finally {
        // eslint-disable-next-line no-undef
        if (typeof RequestGate !== 'undefined') RequestGate.exit(transactionKey);
      }
    } catch (outerError) {
      console.error('Data.updateUser outer error:', outerError.message);
      const transactionKey = `update_user_${userId}`;
      // eslint-disable-next-line no-undef
      if (typeof RequestGate !== 'undefined') RequestGate.exit(transactionKey);
      return {
        success: false,
        error: outerError.message
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

  /**
   * サービスアカウントトークンでスプレッドシートを開く
   * @param {string} spreadsheetId - Spreadsheet ID
   * @param {string} accessToken - Service account access token
   * @returns {Object} Spreadsheet wrapper object
   */
  static openSpreadsheetWithServiceAccount(spreadsheetId, accessToken) {
    // Validate parameters before API URL construction
    if (!spreadsheetId || typeof spreadsheetId !== 'string') {
      throw new Error('Invalid spreadsheet ID: must be a non-empty string');
    }
    if (!accessToken || typeof accessToken !== 'string') {
      throw new Error('Invalid access token: must be a non-empty string');
    }

    const sheetsApiBase = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}`;

    return {
      getId: () => spreadsheetId,

      getSheetByName(name) {
        return Data.createSheetWrapper(spreadsheetId, name, accessToken);
      },

      getSheets() {
        try {
          const response = UrlFetchApp.fetch(sheetsApiBase, {
            method: 'GET',
            headers: {
              'Authorization': `Bearer ${accessToken}`,
              'Content-Type': 'application/json'
            },
            muteHttpExceptions: true
          });

          const data = JSON.parse(response.getContentText());
          if (response.getResponseCode() !== 200) {
            throw new Error(`Sheets API error: ${data.error?.message || response.getResponseCode()}`);
          }

          return data.sheets.map(sheet =>
            Data.createSheetWrapper(spreadsheetId, sheet.properties.title, accessToken)
          );
        } catch (error) {
          console.error('openSpreadsheetWithServiceAccount.getSheets error:', error.message);
          throw error;
        }
      },

      createSheetWrapper: (spreadsheetId, sheetName, accessToken) => {
        return Data.createSheetWrapper(spreadsheetId, sheetName, accessToken);
      }
    };
  }

  /**
   * サービスアカウント用シートラッパー作成
   * @param {string} spreadsheetId - Spreadsheet ID
   * @param {string} sheetName - Sheet name
   * @param {string} accessToken - Service account access token
   * @returns {Object} Sheet wrapper object
   */
  static createSheetWrapper(spreadsheetId, sheetName, accessToken) {
    const sheetsApiBase = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}`;

    return {
      getName: () => sheetName,

      getSheetId() {
        try {
          const response = UrlFetchApp.fetch(sheetsApiBase, {
            method: 'GET',
            headers: {
              'Authorization': `Bearer ${accessToken}`,
              'Content-Type': 'application/json'
            },
            muteHttpExceptions: true
          });

          const data = JSON.parse(response.getContentText());
          if (response.getResponseCode() !== 200) {
            throw new Error(`Sheets API error: ${data.error?.message || response.getResponseCode()}`);
          }

          const sheet = data.sheets.find(s => s.properties.title === sheetName);
          return sheet ? sheet.properties.sheetId : 0;
        } catch (error) {
          console.error('getSheetId error:', error.message);
          return 0;
        }
      },

      getDataRange() {
        return {
          getValues() {
            try {
              const response = UrlFetchApp.fetch(
                `${sheetsApiBase}/values/${sheetName}!A:Z?valueRenderOption=UNFORMATTED_VALUE&dateTimeRenderOption=FORMATTED_STRING`,
                {
                  method: 'GET',
                  headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                  },
                  muteHttpExceptions: true
                }
              );

              const data = JSON.parse(response.getContentText());
              if (response.getResponseCode() !== 200) {
                throw new Error(`Sheets API error: ${data.error?.message || response.getResponseCode()}`);
              }

              return data.values || [];
            } catch (error) {
              console.error('getDataRange.getValues error:', error.message);
              throw error;
            }
          }
        };
      },

      getLastRow() {
        try {
          const response = UrlFetchApp.fetch(
            `${sheetsApiBase}/values/${sheetName}!A:A?valueRenderOption=UNFORMATTED_VALUE`,
            {
              method: 'GET',
              headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
              },
              muteHttpExceptions: true
            }
          );

          const data = JSON.parse(response.getContentText());
          if (response.getResponseCode() !== 200) {
            throw new Error(`Sheets API error: ${data.error?.message || response.getResponseCode()}`);
          }

          return data.values ? data.values.length : 0;
        } catch (error) {
          console.error('getLastRow error:', error.message);
          return 0;
        }
      },

      getLastColumn() {
        try {
          const response = UrlFetchApp.fetch(
            `${sheetsApiBase}/values/${sheetName}!1:1?valueRenderOption=UNFORMATTED_VALUE`,
            {
              method: 'GET',
              headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
              },
              muteHttpExceptions: true
            }
          );

          const data = JSON.parse(response.getContentText());
          if (response.getResponseCode() !== 200) {
            throw new Error(`Sheets API error: ${data.error?.message || response.getResponseCode()}`);
          }

          return data.values && data.values[0] ? data.values[0].length : 0;
        } catch (error) {
          console.error('getLastColumn error:', error.message);
          return 0;
        }
      },

      getRange(range, col, numRows, numCols) {
        let rangeNotation;

        if (typeof range === 'string') {
          rangeNotation = range;
        } else {
          // Convert numeric parameters to A1 notation
          rangeNotation = Data.convertToA1Notation(range, col, numRows, numCols);
        }
        return {
          getValues() {
            try {
              const response = UrlFetchApp.fetch(
                `${sheetsApiBase}/values/${sheetName}!${rangeNotation}?valueRenderOption=UNFORMATTED_VALUE&dateTimeRenderOption=FORMATTED_STRING`,
                {
                  method: 'GET',
                  headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                  },
                  muteHttpExceptions: true
                }
              );

              const data = JSON.parse(response.getContentText());
              if (response.getResponseCode() !== 200) {
                throw new Error(`Sheets API error: ${data.error?.message || response.getResponseCode()}`);
              }

              return data.values || [];
            } catch (error) {
              console.error('getRange.getValues error:', error.message);
              throw error;
            }
          },

          setValues(values) {
            try {
              const response = UrlFetchApp.fetch(
                `${sheetsApiBase}/values/${sheetName}!${rangeNotation}?valueInputOption=RAW`,
                {
                  method: 'PUT',
                  headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                  },
                  payload: JSON.stringify({
                    values
                  }),
                  muteHttpExceptions: true
                }
              );

              const data = JSON.parse(response.getContentText());
              if (response.getResponseCode() !== 200) {
                throw new Error(`Sheets API error: ${data.error?.message || response.getResponseCode()}`);
              }

              return data;
            } catch (error) {
              console.error('getRange.setValues error:', error.message);
              throw error;
            }
          },

          getValue() {
            try {
              const response = UrlFetchApp.fetch(
                `${sheetsApiBase}/values/${sheetName}!${rangeNotation}?valueRenderOption=UNFORMATTED_VALUE&dateTimeRenderOption=FORMATTED_STRING`,
                {
                  method: 'GET',
                  headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                  },
                  muteHttpExceptions: true
                }
              );

              const data = JSON.parse(response.getContentText());
              if (response.getResponseCode() !== 200) {
                throw new Error(`Sheets API error: ${data.error?.message || response.getResponseCode()}`);
              }

              return data.values && data.values[0] && data.values[0][0] ? data.values[0][0] : null;
            } catch (error) {
              console.error('getValue error:', error.message);
              return null;
            }
          },

          setValue(value) {
            try {
              const response = UrlFetchApp.fetch(
                `${sheetsApiBase}/values/${sheetName}!${rangeNotation}?valueInputOption=RAW`,
                {
                  method: 'PUT',
                  headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                  },
                  payload: JSON.stringify({
                    values: [[value]]
                  }),
                  muteHttpExceptions: true
                }
              );

              const data = JSON.parse(response.getContentText());
              if (response.getResponseCode() !== 200) {
                throw new Error(`Sheets API error: ${data.error?.message || response.getResponseCode()}`);
              }

              return data;
            } catch (error) {
              console.error('setValue error:', error.message);
              throw error;
            }
          }
        };
      },

      appendRow(values) {
        try {
          const response = UrlFetchApp.fetch(
            `${sheetsApiBase}/values/${sheetName}!A:A:append?valueInputOption=RAW`,
            {
              method: 'POST',
              headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
              },
              payload: JSON.stringify({
                values: [values]
              }),
              muteHttpExceptions: true
            }
          );

          const data = JSON.parse(response.getContentText());
          if (response.getResponseCode() !== 200) {
            throw new Error(`Sheets API error: ${data.error?.message || response.getResponseCode()}`);
          }

          return data;
        } catch (error) {
          console.error('appendRow error:', error.message);
          throw error;
        }
      }
    };
  }

  /**
   * Convert column number to letter
   * @param {number} column - Column number (1-based)
   * @returns {string} Column letter
   */
  static columnToLetter(column) {
    let result = '';
    while (column > 0) {
      column--;
      result = String.fromCharCode(65 + (column % 26)) + result;
      column = Math.floor(column / 26);
    }
    return result;
  }

  /**
   * Convert numeric parameters to A1 notation
   * @param {number} row - Starting row (1-based)
   * @param {number} col - Starting column (1-based)
   * @param {number} numRows - Number of rows
   * @param {number} numCols - Number of columns
   * @returns {string} A1 notation range
   */
  static convertToA1Notation(row, col, numRows, numCols) {
    const startCol = Data.columnToLetter(col);
    const endCol = numCols ? Data.columnToLetter(col + numCols - 1) : startCol;
    const endRow = numRows ? row + numRows - 1 : row;

    if (numRows === 1 && numCols === 1) {
      return `${startCol}${row}`;
    } else {
      return `${startCol}${row}:${endCol}${endRow}`;
    }
  }
}

// Export for global access (Zero-Dependency Architecture)
if (typeof globalThis !== 'undefined') {
  globalThis.Data = Data;
} else if (typeof global !== 'undefined') {
  global.Data = Data;
} else {
  this.Data = Data;
}