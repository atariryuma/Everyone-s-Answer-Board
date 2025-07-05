/**
 * @fileoverview データベース管理 - バッチ操作とキャッシュ最適化
 * GAS互換の関数ベースの実装
 */

// データベース管理のための定数
var USER_CACHE_TTL = 300; // 5分
var DB_BATCH_SIZE = 100;

/**
 * 最適化されたSheetsサービスを取得
 * @returns {object} Sheets APIサービス
 */
function getSheetsService() {
  var accessToken = getServiceAccountTokenCached();
  return createSheetsService(accessToken);
}

/**
 * ユーザー情報を効率的に検索（キャッシュ優先）
 * @param {string} userId - ユーザーID
 * @returns {object|null} ユーザー情報
 */
function findUserById(userId) {
  var cacheKey = 'user_' + userId;
  var cache = CacheService.getScriptCache();
  
  // キャッシュから取得を試行
  var cachedUser = cache.get(cacheKey);
  if (cachedUser) {
    debugLog('ユーザーキャッシュヒット: ' + userId);
    try {
      return JSON.parse(cachedUser);
    } catch (e) {
      console.error('キャッシュデータ解析エラー:', e);
    }
  }
  
  // データベースから取得
  var user = fetchUserFromDatabase('userId', userId);
  
  if (user) {
    // キャッシュに保存
    cache.put(cacheKey, JSON.stringify(user), USER_CACHE_TTL);
    cache.put('email_' + user.adminEmail, JSON.stringify(user), USER_CACHE_TTL);
  }
  
  return user;
}

/**
 * メールアドレスでユーザー検索
 * @param {string} email - メールアドレス
 * @returns {object|null} ユーザー情報
 */
function findUserByEmail(email) {
  var cacheKey = 'email_' + email;
  var cache = CacheService.getScriptCache();
  
  // キャッシュから取得を試行
  var cachedUser = cache.get(cacheKey);
  if (cachedUser) {
    debugLog('メールキャッシュヒット: ' + email);
    try {
      return JSON.parse(cachedUser);
    } catch (e) {
      console.error('キャッシュデータ解析エラー:', e);
    }
  }
  
  // データベースから取得
  var user = fetchUserFromDatabase('adminEmail', email);
  
  if (user) {
    // 両方のキーでキャッシュ
    cache.put(cacheKey, JSON.stringify(user), USER_CACHE_TTL);
    cache.put('user_' + user.userId, JSON.stringify(user), USER_CACHE_TTL);
  }
  
  return user;
}

/**
 * データベースからユーザーを取得
 * @param {string} field - 検索フィールド
 * @param {string} value - 検索値
 * @returns {object|null} ユーザー情報
 */
function fetchUserFromDatabase(field, value) {
  try {
    var props = PropertiesService.getScriptProperties();
    var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    var service = getSheetsService();
    var sheetName = DB_SHEET_CONFIG.SHEET_NAME;
    
    var data = batchGetSheetsData(service, dbId, [sheetName + '!A:H']);
    var values = data.valueRanges[0].values || [];
    
    if (values.length === 0) return null;
    
    var headers = values[0];
    var fieldIndex = headers.indexOf(field);
    
    if (fieldIndex === -1) return null;
    
    for (var i = 1; i < values.length; i++) {
      if (values[i][fieldIndex] === value) {
        var user = {};
        headers.forEach(function(header, index) {
          user[header] = values[i][index] || '';
        });
        return user;
      }
    }
    
    return null;
  } catch (error) {
    console.error('ユーザー検索エラー (' + field + ':' + value + '):', error);
    return null;
  }
}

/**
 * ユーザー情報をキャッシュ優先で取得し、見つからなければDBから直接取得
 * @param {string} userId - ユーザーID
 * @returns {object|null} ユーザー情報
 */
function getUserWithFallback(userId) {
  var user = findUserById(userId);
  if (user) return user;

  debugLog('[Cache] MISS for key: user_' + userId + '. Fetching from DB.');
  user = fetchUserFromDatabase('userId', userId);
  if (user) {
    var cache = CacheService.getScriptCache();
    cache.put('user_' + userId, JSON.stringify(user), USER_CACHE_TTL);
    cache.put('email_' + user.adminEmail, JSON.stringify(user), USER_CACHE_TTL);
  } else {
    handleMissingUser(userId);
  }
  return user;
}

/**
 * ユーザー情報を一括更新
 * @param {string} userId - ユーザーID
 * @param {object} updateData - 更新データ
 * @returns {object} 更新結果
 */
function updateUser(userId, updateData) {
  try {
    var props = PropertiesService.getScriptProperties();
    var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    var service = getSheetsService();
    var sheetName = DB_SHEET_CONFIG.SHEET_NAME;
    
    // 現在のデータを取得
    var data = batchGetSheetsData(service, dbId, [sheetName + '!A:H']);
    var values = data.valueRanges[0].values || [];
    
    if (values.length === 0) {
      throw new Error('データベースが空です');
    }
    
    var headers = values[0];
    var userIdIndex = headers.indexOf('userId');
    var rowIndex = -1;
    
    // ユーザーの行を特定
    for (var i = 1; i < values.length; i++) {
      if (values[i][userIdIndex] === userId) {
        rowIndex = i + 1; // 1-based index
        break;
      }
    }
    
    if (rowIndex === -1) {
      throw new Error('更新対象のユーザーが見つかりません');
    }
    
    // バッチ更新リクエストを作成
    var requests = Object.keys(updateData).map(function(key) {
      var colIndex = headers.indexOf(key);
      if (colIndex === -1) return null;
      
      return {
        range: sheetName + '!' + String.fromCharCode(65 + colIndex) + rowIndex,
        values: [[updateData[key]]]
      };
    }).filter(function(item) { return item !== null; });
    
    if (requests.length > 0) {
      batchUpdateSheetsData(service, dbId, requests);
    }
    
    // キャッシュを無効化
    invalidateUserCache(userId, updateData.adminEmail);
    
    return { success: true };
  } catch (error) {
    console.error('ユーザー更新エラー:', error);
    throw error;
  }
}

/**
 * 新規ユーザーをデータベースに作成
 * @param {object} userData - 作成するユーザーデータ
 * @returns {object} 作成されたユーザーデータ
 */
function createUser(userData) {
  // メールアドレスの重複チェック
  var existingUser = findUserByEmail(userData.adminEmail);
  if (existingUser) {
    throw new Error('このメールアドレスは既に登録されています。');
  }

  var props = PropertiesService.getScriptProperties();
  var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
  var service = getSheetsService();
  var sheetName = DB_SHEET_CONFIG.SHEET_NAME;

  var newRow = DB_SHEET_CONFIG.HEADERS.map(function(header) { 
    return userData[header] || ''; 
  });
  
  appendSheetsData(service, dbId, sheetName + '!A1', [newRow]);
  
  // キャッシュを無効化
  invalidateUserCache(userData.userId, userData.adminEmail);
  
  return userData;
}

/**
 * データベースシートを初期化
 * @param {string} spreadsheetId - データベースのスプレッドシートID
 */
function initializeDatabaseSheet(spreadsheetId) {
  var service = getSheetsService();
  var sheetName = DB_SHEET_CONFIG.SHEET_NAME;

  try {
    // シートが存在するか確認
    var spreadsheet = getSpreadsheetsData(service, spreadsheetId);
    var sheetExists = spreadsheet.sheets.some(function(s) { 
      return s.properties.title === sheetName; 
    });

    if (!sheetExists) {
      // シートが存在しない場合は作成
      batchUpdateSpreadsheet(service, spreadsheetId, {
        requests: [{ addSheet: { properties: { title: sheetName } } }]
      });
    }
    
    // ヘッダーを書き込み
    var headerRange = sheetName + '!A1:' + String.fromCharCode(65 + DB_SHEET_CONFIG.HEADERS.length - 1) + '1';
    updateSheetsData(service, spreadsheetId, headerRange, [DB_SHEET_CONFIG.HEADERS]);

    debugLog('データベースシート「' + sheetName + '」の初期化が完了しました。');
  } catch (e) {
    console.error('データベースシートの初期化に失敗: ' + e.message);
    throw new Error('データベースシートの初期化に失敗しました。サービスアカウントに編集者権限があるか確認してください。詳細: ' + e.message);
  }
}

/**
 * ユーザーが見つからない場合のキャッシュ処理
 * @param {string} userId - キャッシュ削除対象のユーザーID
 */
function handleMissingUser(userId) {
  invalidateUserCache(userId);
  clearDatabaseCache();
}

/**
 * 全キャッシュをクリア
 */
function clearDatabaseCache() {
  cacheManager.clearExpired();
  PropertiesService.getUserProperties().deleteProperty('CURRENT_USER_ID');
  debugLog('データベース関連キャッシュをクリアしました');
}

// =================================================================
// 最適化されたSheetsサービス関数群
// =================================================================

/**
 * 最適化されたSheetsサービスを作成
 * @param {string} accessToken - アクセストークン
 * @returns {object} Sheetsサービスオブジェクト
 */
function createSheetsService(accessToken) {
  return {
    accessToken: accessToken,
    baseUrl: 'https://sheets.googleapis.com/v4/spreadsheets'
  };
}

/**
 * バッチ取得
 * @param {object} service - Sheetsサービス
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string[]} ranges - 取得範囲の配列
 * @returns {object} レスポンス
 */
function batchGetSheetsData(service, spreadsheetId, ranges) {
  var url = service.baseUrl + '/' + spreadsheetId + '/values:batchGet?' + 
    ranges.map(function(range) { return 'ranges=' + encodeURIComponent(range); }).join('&');
  
  var response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + service.accessToken }
  });
  
  return JSON.parse(response.getContentText());
}

/**
 * バッチ更新
 * @param {object} service - Sheetsサービス
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {object[]} requests - 更新リクエストの配列
 * @returns {object} レスポンス
 */
function batchUpdateSheetsData(service, spreadsheetId, requests) {
  var url = service.baseUrl + '/' + spreadsheetId + '/values:batchUpdate';
  
  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + service.accessToken },
    payload: JSON.stringify({
      data: requests,
      valueInputOption: 'RAW'
    })
  });
  
  return JSON.parse(response.getContentText());
}

/**
 * データ追加
 * @param {object} service - Sheetsサービス
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} range - 範囲
 * @param {array} values - 値の配列
 * @returns {object} レスポンス
 */
function appendSheetsData(service, spreadsheetId, range, values) {
  var url = service.baseUrl + '/' + spreadsheetId + '/values/' + encodeURIComponent(range) + 
    ':append?valueInputOption=RAW&insertDataOption=INSERT_ROWS';
  
  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + service.accessToken },
    payload: JSON.stringify({ values: values })
  });
  
  return JSON.parse(response.getContentText());
}

/**
 * スプレッドシート情報取得
 * @param {object} service - Sheetsサービス
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {object} スプレッドシート情報
 */
function getSpreadsheetsData(service, spreadsheetId) {
  var url = service.baseUrl + '/' + spreadsheetId;
  var response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + service.accessToken }
  });
  return JSON.parse(response.getContentText());
}

/**
 * データ更新
 * @param {object} service - Sheetsサービス
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} range - 範囲
 * @param {array} values - 値の配列
 * @returns {object} レスポンス
 */
function updateSheetsData(service, spreadsheetId, range, values) {
  var url = service.baseUrl + '/' + spreadsheetId + '/values/' + encodeURIComponent(range) + 
    '?valueInputOption=RAW';
  
  var response = UrlFetchApp.fetch(url, {
    method: 'put',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + service.accessToken },
    payload: JSON.stringify({ values: values })
  });
  
  return JSON.parse(response.getContentText());
}

/**
 * スプレッドシート構造更新
 * @param {object} service - Sheetsサービス
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {object} requestBody - リクエストボディ
 * @returns {object} レスポンス
 */
function batchUpdateSpreadsheet(service, spreadsheetId, requestBody) {
  var url = service.baseUrl + '/' + spreadsheetId + ':batchUpdate';
  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + service.accessToken },
    payload: JSON.stringify(requestBody)
  });
  return JSON.parse(response.getContentText());
}

/**
 * 現在のユーザーのアカウント情報をデータベースから完全に削除する
 * @returns {string} 成功メッセージ
 */
function deleteCurrentUserAccount() {
  try {
    const userId = getUserId();
    if (!userId) {
      throw new Error('ユーザーが特定できません。');
    }
    
    // ユーザー情報を取得して、関連情報を得る
    const userInfo = findUserById(userId);
    if (!userInfo) {
      throw new Error('データベースにユーザー情報が見つかりません。');
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(15000); // 15秒待機

    try {
      // データベース（シート）からユーザー行を削除
      const dbSheet = getDbSheet();
      const data = dbSheet.getDataRange().getValues();
      // ユーザーIDに基づいて行を探す（A列がIDと仮定）
      for (let i = data.length - 1; i >= 1; i--) {
        if (data[i][0] === userId) {
          dbSheet.deleteRow(i + 1);
          break;
        }
      }
      
      // 関連するすべてのキャッシュを削除
      invalidateUserCache(userId, userInfo.adminEmail, userInfo.spreadsheetId);
      
      // UserPropertiesからも関連情報を削除
      const props = PropertiesService.getUserProperties();
      props.deleteProperty('CURRENT_USER_ID');

    } finally {
      lock.releaseLock();
    }
    
    const successMessage = 'アカウント「' + userInfo.adminEmail + '」が正常に削除されました。';
    console.log(successMessage);
    return successMessage;

  } catch (error) {
    console.error('アカウント削除エラー:', error.message, error.stack);
    throw new Error('アカウントの削除に失敗しました: ' + error.message);
  }
}