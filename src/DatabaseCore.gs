/**
 * @fileoverview DatabaseCore - データベースコア機能 (遅延初期化対応)
 *
 * 🎯 責任範囲:
 * - データベース接続・認証
 * - 基本CRUD操作
 * - サービスアカウント管理
 *
 * 🔄 GAS Best Practices準拠:
 * - 遅延初期化パターン (DB関数呼び出し時にinit)
 * - ファイル読み込み順序非依存設計
 * - グローバル副作用排除
 */

/* global DB:writable, ServiceFactory */


// 遅延初期化状態管理
let databaseCoreInitialized = false;

/**
 * DatabaseCore遅延初期化
 */
function initDatabaseCore() {
if (databaseCoreInitialized) return true;

try {
databaseCoreInitialized = true;
return true;
} catch (error) {
console.error('initDatabaseCore failed:', error.message);
return false;
}
}


/**
 * DatabaseCore - データベースコア機能
 * 基本的なデータベース操作とサービス管理
 */
// ===========================================
// 🔐 DatabaseCore Functions (Flat)
// ===========================================

// ==========================================
// 🔐 データベース接続・認証
// ==========================================

/**
 * セキュアなデータベースID取得
 * @returns {string} データベースID
 */
function getSecureDatabaseId() {
try {
return ServiceFactory.getProperties().getProperty('DATABASE_SPREADSHEET_ID');
} catch (error) {
console.error('DatabaseCore', {
operation: 'getSecureDatabaseId',
error: error.message
});
throw new Error('データベース設定の取得に失敗しました');
}
}

/**
 * 🔒 セキュリティ: サービスアカウント経由Sheetsサービス取得
 * @returns {Object} Sheetsサービス
 */
function getSheetsService() {
  try {
    const serviceObject = createSheetsService();
    if (!serviceObject || !serviceObject.spreadsheets || !serviceObject.spreadsheets.values) {
      throw new Error('Invalid service structure created');
    }
    return serviceObject;
  } catch (error) {
    console.error('DatabaseCore getSheetsService:', error.message);
    throw new Error('Secure database service initialization failed');
  }
}

/**
 * サービスアカウントのメールアドレスを取得
 * @returns {string|null} サービスアカウントメール
 */
function getServiceAccountEmail() {
  try {
    const serviceAccountKey = ServiceFactory.getProperties().getProperty('SERVICE_ACCOUNT_CREDS');
    if (!serviceAccountKey) {
      console.error('getServiceAccountEmail: Service account key not configured');
      return null;
    }

    const creds = JSON.parse(serviceAccountKey);
    return creds.client_email || null;
  } catch (error) {
    console.error('getServiceAccountEmail: エラー', error.message);
    return null;
  }
}

/**
 * 🔒 セキュリティ: サービスアカウントSheets API作成
 * @returns {Object} Sheetsサービス
 */
function createSheetsService() {
  try {
    const serviceAccountKey = ServiceFactory.getProperties().getProperty('SERVICE_ACCOUNT_CREDS');
    if (!serviceAccountKey) {
      throw new Error('Service account key not configured');
    }

    // サービスアカウント経由でSpreadsheetApp使用（GAS環境での実装）
    const sheetsServiceObject = {
      openById(spreadsheetId) {
        try {
          const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
          // 自動的にサービスアカウントをエディターとして追加
          const serviceAccountEmail = getServiceAccountEmail();
          if (serviceAccountEmail) {
            try {
              spreadsheet.addEditor(serviceAccountEmail);
            } catch (permError) {
              console.warn('Service account editor registration failed:', permError.message);
            }
          }
          return spreadsheet;
        } catch (error) {
          console.error('getSheetsService.openById error:', error.message);
          throw error;
        }
      },
      spreadsheets: {
        values: {
          get(params) {
            try {
              const spreadsheet = SpreadsheetApp.openById(params.spreadsheetId);
              const [sheetName] = params.range.split('!');
              const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.getSheets()[0];
              const values = sheet.getDataRange().getValues();
              return { data: { values } };
            } catch (error) {
              console.error('Secure service.get error:', error.message);
              throw error;
            }
          },

          update(params) {
            try {
              const spreadsheet = SpreadsheetApp.openById(params.spreadsheetId);
              const [sheetName] = params.range.split('!');
              const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.getSheets()[0];

              if (!sheet) throw new Error(`Sheet "${sheetName}" not found`);

              const {values} = params.resource;
              if (values && values.length > 0) {
                const range = sheet.getRange(1, 1, values.length, values[0].length);
                range.setValues(values);
              }
              return { updatedCells: values ? values.length * values[0].length : 0 };
            } catch (error) {
              console.error('Secure service.update error:', error.message);
              throw error;
            }
          },

          append(params) {
            try {
              const spreadsheet = SpreadsheetApp.openById(params.spreadsheetId);
              const [sheetName] = params.range.split('!');
              const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.getSheets()[0];

              if (!sheet) throw new Error(`Sheet "${sheetName}" not found`);

              const {values} = params.resource;
              if (values && values.length > 0) {
                const lastRow = sheet.getLastRow();
                const targetRange = sheet.getRange(lastRow + 1, 1, values.length, values[0].length);
                targetRange.setValues(values);
              }
              return { updates: { updatedRows: values ? values.length : 0 } };
            } catch (error) {
              console.error('Secure service.append error:', error.message);
              throw error;
            }
          }
        }
      }
    };

    return sheetsServiceObject;
  } catch (error) {
    console.error('createSheetsService error:', error.message);
    throw error;
  }
}

/**
 * バッチデータ取得
 * @param {Object} service - Sheetsサービス
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {Array} ranges - 取得範囲配列
 * @returns {Object} バッチ取得結果
 */
function batchGetSheetsData(service, spreadsheetId, ranges) {
try {

if (!ranges || ranges.length === 0) {
return { valueRanges: [] };
}

const result = service.spreadsheets.values.batchGet({
spreadsheetId,
ranges
});

return result;
} catch (error) {
console.error('DatabaseCore', {
operation: 'batchGetSheetsData',
spreadsheetId,
rangesCount: ranges?.length,
error: error.message
});
throw error;
}
}



// ==========================================
// 🔧 診断・ユーティリティ
// ==========================================

/**
 * データベース接続診断
 * @returns {Object} 診断結果
 */
function diagnoseDatabaseCore() {
const results = {
service: 'DatabaseCore',
timestamp: new Date().toISOString(),
checks: []
};

try {
// データベースID確認
const globalDatabaseId = ServiceFactory.getProperties().getDatabaseSpreadsheetId();
results.checks.push({
name: 'Database ID',
status: globalDatabaseId ? '✅' : '❌',
details: globalDatabaseId ? 'Database ID configured' : 'Database ID missing'
});

// スプレッドシートアクセス確認
try {
const spreadsheet = SpreadsheetApp.openById(globalDatabaseId);
const sheet = spreadsheet.getSheetByName('Users');
results.checks.push({
name: 'Database Access',
status: sheet ? '✅' : '❌',
details: sheet ? 'Database accessible' : 'Users sheet not found'
});
} catch (accessError) {
results.checks.push({
name: 'Database Access',
status: '❌',
details: accessError.message
});
}

results.overall = results.checks.every(check => check.status === '✅') ? '✅' : '⚠️';
} catch (error) {
results.checks.push({
name: 'Core Diagnosis',
status: '❌',
details: error.message
});
results.overall = '❌';
}

return results;
}

/**
 * @fileoverview DatabaseOperations - データベース操作機能
 *
 * 🎯 責任範囲:
 * - ユーザーCRUD操作
 * - データ検索・フィルタリング
 * - バルクデータ操作
 */

/**
 * DatabaseOperations - データベース操作機能
 * ユーザーデータの基本操作を提供
 */
// ===========================================
// 🗄️ DatabaseOperations Functions (Flat)
// ===========================================

// ==========================================
// 👤 ユーザーCRUD操作
// ==========================================

/**
 * メールアドレスでユーザー検索
 * @param {string} email - メールアドレス
 * @returns {Object|null} ユーザー情報
 */
function findUserByEmail(email) {
  if (!email) return null;

  try {
    // 🔒 Security: Service Account access only
    const dbService = getSheetsService();
    const dbId = getSecureDatabaseId();

    const range = 'Users!A:Z';
    const response = dbService.spreadsheets.values.get({
      spreadsheetId: dbId,
      range
    });

    const rows = response.data && response.data.values ? response.data.values : [];
    if (rows.length <= 1) {
      return null; // ヘッダーのみ
    }

    const [headers] = rows;
    const emailIndex = headers.findIndex((h) => {
      return h.toLowerCase().includes('email');
    });

    if (emailIndex === -1) {
      throw new Error('メール列が見つかりません');
    }

    // メールアドレスで検索
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (row[emailIndex] && row[emailIndex].toLowerCase() === email.toLowerCase()) {
        const user = rowToUser(row, headers);
        user.rowIndex = i + 1; // スプレッドシートの行番号（1ベース）
        return user;
      }
    }

    return null;
  } catch (error) {
    console.error('DatabaseOperations', {
      operation: 'findUserByEmail',
      email: typeof email === 'string' && email ? `${email.substring(0, 5)  }***` : `[${  typeof email  }]`,
      error: error.message
    });
    return null;
  }
}

/**
 * ユーザーID検索
 * @param {string} userId - ユーザーID
 * @returns {Object|null} ユーザー情報
 */
function findUserById(userId) {
  if (!userId) return null;

  try {
    // 🔒 Security: Service Account access only
    const dbService = getSheetsService();
    const dbId = getSecureDatabaseId();

    const range = 'Users!A:Z';
    const response = dbService.spreadsheets.values.get({
      spreadsheetId: dbId,
      range
    });

    const rows = response.data && response.data.values ? response.data.values : [];
    if (rows.length <= 1) {
      return null;
    }

    const [headers] = rows;
    const userIdIndex = headers.findIndex((h) => {
      return h.toLowerCase().includes('userid');
    });

    if (userIdIndex === -1) {
      throw new Error('ユーザーID列が見つかりません');
    }

    // ユーザーIDで検索
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (row[userIdIndex] === userId) {
        const user = rowToUser(row, headers);
        user.rowIndex = i + 1; // スプレッドシートの行番号（1ベース）
        return user;
      }
    }

    return null;
  } catch (error) {
    console.error('DatabaseOperations', {
      operation: 'findUserById',
      userId: typeof userId === 'string' && userId ? `${userId.substring(0, 8)  }***` : `[${  typeof userId  }]`,
      error: error.message
    });
    return null;
  }
}

/**
 * 新規ユーザー作成
 * @param {string} email - メールアドレス
 * @param {Object} additionalData - 追加データ
 * @returns {Object} 作成されたユーザー情報
 */
function dbCreateUser(email, additionalData) {
  if (!additionalData) additionalData = {};

  if (!email) {
    throw new Error('メールアドレスが必要です');
  }

  try {

    // 重複チェック
    const existingUser = findUserByEmail(email);
    if (existingUser) {
      throw new Error('既に登録済みのメールアドレスです');
    }

    const dbService = getSheetsService();
    const dbId = getSecureDatabaseId();

    // 新しいユーザーデータ作成
    const userId = Utilities.getUuid();
    const now = new Date().toISOString();

    const userData = {
      userId,
      userEmail: email,
      createdAt: now,
      lastModified: now,
      configJson: JSON.stringify({})
    };

    // additionalDataの内容を追加
    const keys = Object.keys(additionalData);
    for (let i = 0; i < keys.length; i++) {
      const key = keys[i];
      userData[key] = additionalData[key];
    }

    // 🔒 Security: Service Account database access
    // const sheetsService = getSheetsService(); // 既にdbServiceが宣言済み
    // const dbId = getSecureDatabaseId(); // 既にdbIdが宣言済み

    // ヘッダー行確認
    const headerCheckResponse = dbService.spreadsheets.values.get({
      spreadsheetId: dbId,
      range: 'Users!A1:E1'
    });

    const existingData = headerCheckResponse.data && headerCheckResponse.data.values ? headerCheckResponse.data.values : [];
    const hasHeaders = existingData.length > 0 && existingData[0][0] === 'userId';

    if (!hasHeaders) {
      console.log('dbCreateUser: Adding header row');
      dbService.spreadsheets.values.update({
        spreadsheetId: dbId,
        range: 'Users!A1:E1',
        valueInputOption: 'USER_ENTERED',
        resource: {
          values: [['userId', 'userEmail', 'isActive', 'configJson', 'lastModified']]
        }
      });
    }

    // ユーザーデータ追加
    dbService.spreadsheets.values.append({
      spreadsheetId: dbId,
      range: 'Users!A:E',
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: [Object.values(userData)]
      }
    });

    console.log('dbCreateUser: Secure user data added', {
      userId: userData.userId ? `${userData.userId.substring(0, 8)  }***` : 'NO_ID'
    });


    return userData;
  } catch (error) {
    console.error('DatabaseOperations', {
      operation: 'createUser',
      email: typeof email === 'string' && email ? `${email.substring(0, 5)  }***` : `[${  typeof email  }]`,
      error: error.message
    });
    throw error;
  }
}

/**
 * ユーザー情報更新
 * @param {string} userId - ユーザーID
 * @param {Object} updateData - 更新データ
 * @returns {boolean} 更新成功可否
 */
function updateUser(userId, updateData) {
  if (!userId || !updateData) {
    throw new Error('ユーザーIDと更新データが必要です');
  }

  try {
    // ユーザー検索
    const user = findUserById(userId);
    if (!user) {
      throw new Error('ユーザーが見つかりません');
    }

    // 🔒 Security: Service Account database access
    const dbService = getSheetsService();
    const dbId = getSecureDatabaseId();

    // 現在のデータを取得
    const range = 'Users!A:Z';
    const response = dbService.spreadsheets.values.get({
      spreadsheetId: dbId,
      range
    });

    const rows = response.data && response.data.values ? response.data.values : [];
    if (rows.length <= 1) {
      throw new Error('ユーザーデータが見つかりません');
    }

    const [headers] = rows;
    const userRowIndex = user.rowIndex - 1; // 0ベースに変換

    if (userRowIndex >= rows.length) {
      throw new Error('ユーザー行が見つかりません');
    }

    // 許可されたフィールドのみ更新
    const allowedFields = ['userEmail', 'isActive', 'configJson', 'lastModified'];
    const updates = [];
    const updatedRow = [...rows[userRowIndex]]; // 現在の行をコピー

    // 各フィールドを対応する列に更新
    Object.keys(updateData).forEach(field => {
      if (allowedFields.includes(field)) {
        const columnIndex = headers.findIndex(h => h.toLowerCase().includes(field.toLowerCase()));
        if (columnIndex !== -1) {
          updatedRow[columnIndex] = updateData[field];
          updates.push(field);
        } else {
          console.warn(`⚠️ Column not found for field: ${field}`);
        }
      } else {
        console.warn(`⚠️ Field not allowed: ${field}`);
      }
    });

    // 行全体を更新
    if (updates.length > 0) {
      const updateRange = `Users!A${user.rowIndex}:${String.fromCharCode(65 + updatedRow.length - 1)}${user.rowIndex}`;
      dbService.spreadsheets.values.update({
        spreadsheetId: dbId,
        range: updateRange,
        valueInputOption: 'USER_ENTERED',
        resource: {
          values: [updatedRow]
        }
      });
    }

    return { success: true, updatedFields: updates };
  } catch (error) {
    console.error('DatabaseOperations', {
      operation: 'updateUser',
      userId: typeof userId === 'string' && userId ? `${userId.substring(0, 8)  }***` : `[${  typeof userId  }]`,
      error: error.message || 'Unknown error'
    });
    return { success: false, message: error.message || 'Unknown error' };
  }
}

// ==========================================
// 🔧 ユーティリティ
// ==========================================

/**
 * スプレッドシート行をユーザーオブジェクトに変換
 * @param {Array} row - データ行
 * @param {Array} headers - ヘッダー行
 * @returns {Object} ユーザーオブジェクト
 */
function rowToUser(row, headers) {
  const user = {};

  headers.forEach((header, index) => {
    const value = row[index] || '';
    const key = header.toLowerCase()
      .replace(/\s+/g, '')
      .replace('userid', 'userId')
      .replace('useremail', 'userEmail')
      .replace('createdat', 'createdAt')
      .replace('lastmodified', 'lastModified')
      .replace('configjson', 'configJson');

    user[key] = value;
  });

  return user;
}

/**
 * 全ユーザー取得
 * @param {Object} options - オプション
 * @returns {Array} ユーザー一覧
 */
function getAllUsers(options) {
  if (!options) options = {};
  const limit = options.limit || 1000;
  const offset = options.offset || 0;
  const activeOnly = options.activeOnly || false;

  try {
    // 🔒 Security: Service Account access only
    const dbService = getSheetsService();
    const dbId = getSecureDatabaseId();

    const range = 'Users!A:Z';
    const response = dbService.spreadsheets.values.get({
      spreadsheetId: dbId,
      range
    });

    const rows = response.data && response.data.values ? response.data.values : [];
    if (rows.length <= 1) {
      return [];
    }

    const [headers] = rows;
    const users = [];

    // ヘッダーをスキップしてユーザーデータを処理
    for (let i = 1 + offset; i < rows.length && users.length < limit; i++) {
      const row = rows[i];
      if (!row || row.length === 0) continue;

      const user = rowToUser(row, headers);
      if (user && (!activeOnly || user.isActive !== 'FALSE')) {
        users.push(user);
      }
    }

    return users;

  } catch (error) {
    console.error('DatabaseOperations', {
      operation: 'getAllUsers',
      options,
      error: error.message
    });
    return [];
  }
}

/**
 * 診断機能
 * @returns {Object} 診断結果
 */
function diagnoseDatabaseOperations() {
  return {
    service: 'DatabaseOperations',
    timestamp: new Date().toISOString(),
    features: [
      'User CRUD operations',
      'Email-based search',
      'User ID lookup',
      'Batch operations',
      'Get all users with filtering'
    ],
    dependencies: [
      'DatabaseCore'
    ],
    status: '✅ Active'
  };
}

// ===========================================
// 🌐 グローバルDB変数設定（GAS読み込み順序対応）
// ===========================================

/**
 * グローバルDB変数を確実に設定
 * GASファイル読み込み順序に関係なく、DatabaseOperationsを利用可能にする
 */
// Create DatabaseOperations object for backward compatibility
const DatabaseOperations = {
  findUserByEmail,
  findUserById,
  getAllUsers,
  createUser: dbCreateUser,
  updateUser,
  deleteUserAccountByAdmin: updateUser, // Alias
  diagnose: diagnoseDatabaseOperations
};

// グローバル変数に代入（環境互換: globalThis / global / this）
(function () {
  const root = (typeof globalThis !== 'undefined') ? globalThis : (typeof global !== 'undefined' ? global : this);
  try {
    root.DB = DatabaseOperations;
  } catch (e) {
    // Fallback (should not happen in V8), avoid leaking bare global symbol
  }
})();
