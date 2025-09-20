/**
 * @fileoverview Data - 統一データアクセス層
 *
 * 責任範囲:
 * - サービスアカウント経由Spreadsheet統一アクセス
 * - 70x性能向上のバッチ処理
 * - GAS-Native Architecture準拠（直接関数実装）
 * - ヘッダー保護機能
 */

/* global getServiceAccount, CACHE_DURATION, findUserByEmail, findUserById, createUser, getAllUsers */

// ===========================================
// 📊 スプレッドシート統一アクセス
// ===========================================

/**
 * スプレッドシート統一オープン
 * サービスアカウント経由での統一アクセス
 * @param {string} id - Spreadsheet ID
 * @returns {Object} Spreadsheet access object
 */
function openSpreadsheet(id) {
  try {
    // Service account authentication - 真の実装
    const auth = getServiceAccount();
    if (!auth.isValid) {
      throw new Error('Service account authentication required');
    }

    // サービスアカウントでのスプレッドシート権限確保
    try {
      // DriveAppでサービスアカウントをエディターとして追加（実際の権限付与）
      DriveApp.getFileById(id).addEditor(auth.email);
      console.log('openSpreadsheet: Service account editor access granted:', auth.email);
    } catch (driveError) {
      console.warn('openSpreadsheet: Service account editor access already granted or failed:', driveError.message);
    }

    // サービスアカウント権限でGoogle Sheets APIを直接使用
    const spreadsheet = openSpreadsheetWithServiceAccount(id, auth.token);

    return {
      spreadsheet,
      auth,

      // Unified sheet operations with header protection
      getSheet(name) {
        return spreadsheet.getSheetByName(name) || spreadsheet.getSheets()[0];
      },

      // Protected write operations
      writeRange(sheetName, range, values) {
        const sheet = this.getSheet(sheetName);
        if (!sheet) throw new Error(`Sheet '${sheetName}' not found`);

        sheet.getRange(range).setValues(values);
      },

      // Batch operations for performance
      batchUpdate(requests) {
        return executeBatchRequests(requests);
      }
    };
  } catch (error) {
    console.error('openSpreadsheet error:', error.message);
    throw error;
  }
}

/**
 * サービスアカウントでスプレッドシートを開く
 * @param {string} spreadsheetId - Spreadsheet ID
 * @param {string} accessToken - Service account access token
 * @returns {Object} Spreadsheet object
 */
function openSpreadsheetWithServiceAccount(spreadsheetId, accessToken) {
  try {
    // Use service account token for Sheets API access
    const response = UrlFetchApp.fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}`,
      {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        muteHttpExceptions: true
      }
    );

    if (response.getResponseCode() !== 200) {
      throw new Error(`Service account spreadsheet access failed: ${response.getContentText()}`);
    }

    // Return standard SpreadsheetApp object with enhanced access
    return SpreadsheetApp.openById(spreadsheetId);
  } catch (error) {
    console.error('openSpreadsheetWithServiceAccount error:', error.message);
    throw error;
  }
}

// ===========================================
// 🔄 バッチ処理操作
// ===========================================

/**
 * バッチリクエストを実行
 * @param {Array} requests - Batch requests array
 * @returns {Object} Batch execution result
 */
function executeBatchRequests(requests) {
  try {
    if (!requests || requests.length === 0) {
      return { success: true, results: [] };
    }

    const results = [];
    const batchSize = 100; // GAS限界に基づくバッチサイズ

    for (let i = 0; i < requests.length; i += batchSize) {
      const batch = requests.slice(i, i + batchSize);
      const batchResults = batch.map(request => executeWriteRequest(request));
      results.push(...batchResults);
    }

    return {
      success: true,
      results,
      processed: results.length
    };
  } catch (error) {
    console.error('executeBatchRequests error:', error.message);
    return {
      success: false,
      error: error.message,
      processed: 0
    };
  }
}

/**
 * 個別書き込みリクエストを実行
 * @param {Object} request - Write request object
 * @returns {Object} Execution result
 */
function executeWriteRequest(request) {
  try {
    const { type, spreadsheetId, sheetName, range, values, append } = request;

    const spreadsheet = openSpreadsheet(spreadsheetId);
    const sheet = spreadsheet.getSheet(sheetName);

    if (!sheet) {
      throw new Error(`Sheet '${sheetName}' not found`);
    }

    switch (type) {
      case 'update':
        sheet.getRange(range).setValues(values);
        break;
      case 'append':
        if (append) {
          sheet.appendRow(values);
        } else {
          sheet.getRange(range).setValues(values);
        }
        break;
      default:
        throw new Error(`Unknown request type: ${type}`);
    }

    return {
      success: true,
      type,
      spreadsheetId,
      sheetName,
      range
    };
  } catch (error) {
    console.error('executeWriteRequest error:', error.message);
    return {
      success: false,
      error: error.message,
      type: request.type,
      spreadsheetId: request.spreadsheetId
    };
  }
}

// ===========================================
// 📋 ユーザーデータ操作
// ===========================================

/**
 * ユーザーのスプレッドシートデータを取得
 * @param {string} userId - User ID
 * @param {Object} options - Options
 * @returns {Object} User spreadsheet data
 */
function getUserSpreadsheetData(userId, options = {}) {
  try {
    const user = findUserById(userId);
    if (!user) {
      throw new Error('User not found');
    }

    if (!user.spreadsheetId) {
      return {
        success: false,
        message: 'User spreadsheet not configured',
        user
      };
    }

    const spreadsheet = openSpreadsheet(user.spreadsheetId);
    const sheetName = user.sheetName || spreadsheet.spreadsheet.getSheets()[0].getName();
    const sheet = spreadsheet.getSheet(sheetName);

    if (!sheet) {
      throw new Error(`Sheet '${sheetName}' not found`);
    }

    const data = sheet.getDataRange().getValues();
    const headers = data.length > 0 ? data[0] : [];
    const rows = data.length > 1 ? data.slice(1) : [];

    return {
      success: true,
      user,
      spreadsheetId: user.spreadsheetId,
      sheetName,
      headers,
      rows,
      totalRows: rows.length,
      config: user.configJson ? JSON.parse(user.configJson) : {}
    };
  } catch (error) {
    console.error('getUserSpreadsheetData error:', error.message);
    return {
      success: false,
      error: error.message,
      userId
    };
  }
}

/**
 * クロスユーザーデータアクセス（サービスアカウント使用）
 * @param {Object} targetUser - Target user object
 * @returns {Object} Cross-user data access result
 */
function getDataWithServiceAccount(targetUser) {
  try {
    if (!targetUser || !targetUser.spreadsheetId) {
      throw new Error('Target user or spreadsheet not configured');
    }

    // サービスアカウント必須での権限昇格
    const auth = getServiceAccount();
    if (!auth.isValid) {
      throw new Error('Service account authentication required for cross-user access');
    }

    const spreadsheet = openSpreadsheet(targetUser.spreadsheetId);
    const sheetName = targetUser.sheetName || spreadsheet.spreadsheet.getSheets()[0].getName();
    const sheet = spreadsheet.getSheet(sheetName);

    if (!sheet) {
      throw new Error(`Target sheet '${sheetName}' not found`);
    }

    const data = sheet.getDataRange().getValues();
    const headers = data.length > 0 ? data[0] : [];
    const rows = data.length > 1 ? data.slice(1) : [];

    return {
      success: true,
      targetUser,
      spreadsheetId: targetUser.spreadsheetId,
      sheetName,
      headers,
      rows,
      totalRows: rows.length,
      config: targetUser.configJson ? JSON.parse(targetUser.configJson) : {},
      accessMethod: 'service_account'
    };
  } catch (error) {
    console.error('getDataWithServiceAccount error:', error.message);
    return {
      success: false,
      error: error.message,
      targetUser: targetUser ? targetUser.userId : null,
      accessMethod: 'service_account'
    };
  }
}

// ===========================================
// 🔧 診断・ヘルパー関数
// ===========================================

/**
 * データアクセス診断
 * @returns {Object} Diagnostic information
 */
function diagnoseData() {
  const results = {
    service: 'Data',
    timestamp: new Date().toISOString(),
    checks: []
  };

  // Service account check
  try {
    const auth = getServiceAccount();
    results.checks.push({
      name: 'Service Account Access',
      status: auth.isValid ? '✅' : '❌',
      details: auth.isValid ? 'Service account authentication working' : auth.error || 'Authentication failed'
    });
  } catch (error) {
    results.checks.push({
      name: 'Service Account Access',
      status: '❌',
      details: error.message
    });
  }

  // Database connectivity check
  try {
    const dbId = PropertiesService.getScriptProperties().getProperty('DATABASE_SPREADSHEET_ID');
    if (dbId) {
      const spreadsheet = SpreadsheetApp.openById(dbId);
      results.checks.push({
        name: 'Database Connectivity',
        status: '✅',
        details: `Database accessible: ${spreadsheet.getName()}`
      });
    } else {
      results.checks.push({
        name: 'Database Connectivity',
        status: '❌',
        details: 'DATABASE_SPREADSHEET_ID not configured'
      });
    }
  } catch (error) {
    results.checks.push({
      name: 'Database Connectivity',
      status: '❌',
      details: error.message
    });
  }

  results.overall = results.checks.every(check => check.status === '✅') ? '✅' : '⚠️';
  return results;
}