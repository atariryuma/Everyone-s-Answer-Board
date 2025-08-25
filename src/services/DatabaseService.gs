/**
 * @fileoverview データベースサービス - シンプルなデータアクセス層
 */

/**
 * データベース設定
 */
const DB_CONFIG = {
  USERS_SHEET: 'Users',
  LOGS_SHEET: 'Logs',
  USER_HEADERS: ['userId', 'adminEmail', 'spreadsheetId', 'spreadsheetUrl', 'createdAt', 'configJson', 'lastAccessedAt', 'isActive'],
  LOG_HEADERS: ['timestamp', 'userId', 'action', 'details']
};

/**
 * データベーススプレッドシートを取得
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function getDatabaseSpreadsheet() {
  const dbId = Properties.getScriptProperty('DATABASE_SPREADSHEET_ID');
  if (!dbId) {
    throw new Error('データベーススプレッドシートIDが設定されていません');
  }
  
  return getSpreadsheetCached(dbId, (id) => {
    try {
      return Spreadsheet.openById(id);
    } catch (error) {
      logError(error, 'getDatabaseSpreadsheet', ERROR_SEVERITY.CRITICAL, ERROR_CATEGORIES.DATABASE);
      throw new Error('データベーススプレッドシートにアクセスできません');
    }
  });
}

/**
 * ユーザーシートを取得（なければ作成）
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getUsersSheet() {
  const spreadsheet = getDatabaseSpreadsheet();
  let sheet = spreadsheet.getSheetByName(DB_CONFIG.USERS_SHEET);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(DB_CONFIG.USERS_SHEET);
    sheet.getRange(1, 1, 1, DB_CONFIG.USER_HEADERS.length).setValues([DB_CONFIG.USER_HEADERS]);
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

/**
 * ログシートを取得（なければ作成）
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getLogsSheet() {
  const spreadsheet = getDatabaseSpreadsheet();
  let sheet = spreadsheet.getSheetByName(DB_CONFIG.LOGS_SHEET);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(DB_CONFIG.LOGS_SHEET);
    sheet.getRange(1, 1, 1, DB_CONFIG.LOG_HEADERS.length).setValues([DB_CONFIG.LOG_HEADERS]);
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

/**
 * ユーザーIDからユーザー情報を取得
 * @param {string} userId - ユーザーID
 * @returns {Object|null} ユーザー情報
 */
function getUserById(userId) {
  if (!userId) return null;
  
  // キャッシュから取得
  const cached = getUserDataCached(userId, (id) => {
    try {
      const sheet = getUsersSheet();
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const userIdIndex = headers.indexOf('userId');
      
      for (let i = 1; i < data.length; i++) {
        if (data[i][userIdIndex] === id) {
          return rowToObject(headers, data[i]);
        }
      }
      return null;
    } catch (error) {
      logError(error, 'getUserById', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.DATABASE, { userId: id });
      return null;
    }
  });
  
  return cached;
}

/**
 * メールアドレスからユーザー情報を取得
 * @param {string} email - メールアドレス
 * @returns {Object|null} ユーザー情報
 */
function getUserByEmail(email) {
  if (!email) return null;
  
  try {
    const sheet = getUsersSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const emailIndex = headers.indexOf('adminEmail');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][emailIndex] === email) {
        const user = rowToObject(headers, data[i]);
        // キャッシュに保存
        setCacheValue(`user_${user.userId}`, user, CACHE_CONFIG.USER_DATA_TTL);
        return user;
      }
    }
    return null;
  } catch (error) {
    logError(error, 'getUserByEmail', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.DATABASE, { email });
    return null;
  }
}

/**
 * 新しいユーザーを作成
 * @param {Object} userData - ユーザーデータ
 * @returns {Object} 作成されたユーザー情報
 */
function createUser(userData) {
  try {
    const sheet = getUsersSheet();
    const headers = DB_CONFIG.USER_HEADERS;
    
    // デフォルト値を設定
    const newUser = {
      userId: userData.userId || Utils.getUuid(),
      adminEmail: userData.adminEmail,
      spreadsheetId: userData.spreadsheetId || '',
      spreadsheetUrl: userData.spreadsheetUrl || '',
      createdAt: new Date().toISOString(),
      configJson: JSON.stringify(userData.config || {}),
      lastAccessedAt: new Date().toISOString(),
      isActive: 'true'
    };
    
    // 行データを作成
    const row = headers.map(header => newUser[header] || '');
    sheet.appendRow(row);
    
    // キャッシュをクリア
    clearUserCache(newUser.userId);
    
    return newUser;
  } catch (error) {
    logError(error, 'createUser', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.DATABASE, userData);
    throw error;
  }
}

/**
 * ユーザー情報を更新
 * @param {string} userId - ユーザーID
 * @param {Object} updates - 更新内容
 * @returns {boolean} 成功フラグ
 */
function updateUser(userId, updates) {
  try {
    const sheet = getUsersSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const userIdIndex = headers.indexOf('userId');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][userIdIndex] === userId) {
        // 更新する列のみ変更
        Object.keys(updates).forEach(key => {
          const colIndex = headers.indexOf(key);
          if (colIndex !== -1) {
            data[i][colIndex] = updates[key];
          }
        });
        
        // lastAccessedAtを更新
        const lastAccessedIndex = headers.indexOf('lastAccessedAt');
        if (lastAccessedIndex !== -1) {
          data[i][lastAccessedIndex] = new Date().toISOString();
        }
        
        // データを書き戻し
        sheet.getRange(i + 1, 1, 1, headers.length).setValues([data[i]]);
        
        // キャッシュをクリア
        clearUserCache(userId);
        
        return true;
      }
    }
    return false;
  } catch (error) {
    logError(error, 'updateUser', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.DATABASE, { userId, updates });
    throw error;
  }
}

/**
 * ユーザーを削除（非アクティブ化）
 * @param {string} userId - ユーザーID
 * @returns {boolean} 成功フラグ
 */
function deleteUser(userId) {
  return updateUser(userId, { isActive: 'false' });
}

/**
 * アクティブユーザーのリストを取得
 * @returns {Array<Object>} ユーザーリスト
 */
function getActiveUsers() {
  try {
    const sheet = getUsersSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const isActiveIndex = headers.indexOf('isActive');
    
    const users = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][isActiveIndex] === 'true') {
        users.push(rowToObject(headers, data[i]));
      }
    }
    
    return users;
  } catch (error) {
    logError(error, 'getActiveUsers', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.DATABASE);
    return [];
  }
}

/**
 * アクションログを記録
 * @param {string} userId - ユーザーID
 * @param {string} action - アクション
 * @param {Object} details - 詳細
 */
function logAction(userId, action, details = {}) {
  try {
    const sheet = getLogsSheet();
    sheet.appendRow([
      new Date().toISOString(),
      userId,
      action,
      JSON.stringify(details)
    ]);
  } catch (error) {
    // ログの失敗は無視（システムを止めない）
    warnLog('Failed to log action:', error.message);
  }
}

/**
 * 配列をオブジェクトに変換
 * @private
 */
function rowToObject(headers, row) {
  const obj = {};
  headers.forEach((header, index) => {
    let value = row[index];
    // configJsonをパース
    if (header === 'configJson' && value) {
      try {
        value = JSON.parse(value);
      } catch (e) {
        value = {};
      }
    }
    obj[header] = value;
  });
  return obj;
}

/**
 * ユーザー情報の取得（統合版）
 * @param {string} userId - ユーザーID
 * @returns {Object|null} ユーザー情報
 */
function getUserInfo(userId) {
  return getUserById(userId);
}

/**
 * ユーザーIDをメールアドレスから取得
 * @param {string} email - メールアドレス
 * @returns {string} ユーザーID
 */
function getUserIdFromEmail(email) {
  const user = getUserByEmail(email);
  return user ? user.userId : Utils.getUuid();
}

/**
 * 現在のユーザーIDを取得
 * @returns {string} ユーザーID
 */
function getUserId() {
  const email = SessionService.getActiveUserEmail();
  return getUserIdFromEmail(email);
}

/**
 * デプロイユーザーかチェック
 * @returns {boolean}
 */
function isDeployUser() {
  const email = SessionService.getActiveUserEmail();
  const adminEmail = Properties.getScriptProperty('ADMIN_EMAIL');
  return email === adminEmail;
}

/**
 * デバッグモードを有効にするかチェック
 * @returns {boolean}
 */
function shouldEnableDebugMode() {
  try {
    const debugMode = Properties.getScriptProperty('DEBUG_MODE');
    return debugMode === 'true';
  } catch (error) {
    return false;
  }
}