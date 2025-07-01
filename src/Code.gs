/**
 * @fileoverview StudyQuest - みんなの回答ボード (サービスアカウントモデル)
 * ユーザーがファイル所有権を持ち、中央データベースへのアクセスはサービスアカウント経由で行う。
 * これにより、Workspace管理者の設定変更を不要にし、403エラーを回避する。
 */

// =================================================================
// グローバル設定
// =================================================================

// スクリプトプロパティに保存されるキー
var SCRIPT_PROPS_KEYS = {
  SERVICE_ACCOUNT_CREDS: 'SERVICE_ACCOUNT_CREDS',
  DATABASE_SPREADSHEET_ID: 'DATABASE_SPREADSHEET_ID'
};

// ユーザー専用フォルダ管理の設定
var USER_FOLDER_CONFIG = {
  ROOT_FOLDER_NAME: "StudyQuest - ユーザーデータ",
  FOLDER_NAME_PATTERN: "StudyQuest - {email} - ファイル"
};

// 中央データベースのシート設定
var DB_SHEET_CONFIG = {
  SHEET_NAME: 'Users',
  HEADERS: [
    'userId', 'adminEmail', 'spreadsheetId', 'spreadsheetUrl', 
    'createdAt', 'configJson', 'lastAccessedAt', 'isActive'
  ]
};

// 回答ボードのスプレッドシートの列ヘッダー
var COLUMN_HEADERS = {
  TIMESTAMP: 'タイムスタンプ',
  EMAIL: 'メールアドレス',
  CLASS: 'クラス',
  OPINION: '回答',
  REASON: '理由',
  NAME: '名前',
  UNDERSTAND: 'なるほど！',
  LIKE: 'いいね！',
  CURIOUS: 'もっと知りたい！',
  HIGHLIGHT: 'ハイライト'
};

var REACTION_KEYS = ["UNDERSTAND", "LIKE", "CURIOUS"];
var EMAIL_REGEX = new RegExp("^[^
@]+@[^
@]+\.[^
@]+$");
var DEBUG = true;

function debugLog() {
  if (DEBUG && typeof console !== 'undefined' && console.log) {
    console.log.apply(console, arguments);
  }
}

// =================================================================
// セットアップ＆管理用関数
// =================================================================

/**
 * アプリケーションの初期セットアップ（管理者が手動で実行）
 * サービスアカウントの認証情報とデータベースのスプレッドシートIDを設定する。
 * @param {string} credsJson - ダウンロードしたサービスアカウントのJSONキーファイルの内容
 * @param {string} dbId - 中央データベースとして使用するスプレッドシートのID
 */
function setupApplication(credsJson, dbId) {
  try {
    JSON.parse(credsJson);
    // ★正規表現による検証を文字列長チェックに置き換え
    if (typeof dbId !== 'string' || dbId.length !== 44) {
      throw new Error('無効なスプレッドシートIDです。IDは44文字の文字列である必要があります。');
    }

    var props = PropertiesService.getScriptProperties();
    props.setProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS, credsJson);
    props.setProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID, dbId);

    // データベースシートの初期化
    initializeDatabaseSheet(dbId);

    console.log('✅ セットアップが正常に完了しました。');
  } catch (e) {
    console.error('セットアップエラー:', e);
    throw new Error('セットアップに失敗しました: ' + e.message);
  }
}

/**
 * データベースシートに必要なヘッダーを作成する。
 * @param {string} spreadsheetId - データベースのスプレッドシートID
 */
function initializeDatabaseSheet(spreadsheetId) {
  var service = getSheetsService();
  var sheetName = DB_SHEET_CONFIG.SHEET_NAME;

  try {
    // シートが存在するか確認
    var spreadsheet = service.spreadsheets.get(spreadsheetId);
    var sheetExists = spreadsheet.sheets.some(function(s) { return s.properties.title === sheetName; });

    if (!sheetExists) {
      // シートが存在しない場合は作成
      service.spreadsheets.batchUpdate(spreadsheetId, {
        requests: [{ addSheet: { properties: { title: sheetName } } }]
      });
    }
    
    // ヘッダーを書き込み
    var headerRange = sheetName + '!A1:' + String.fromCharCode(65 + DB_SHEET_CONFIG.HEADERS.length - 1) + '1';
    service.spreadsheets.values.update(
      spreadsheetId,
      headerRange,
      { values: [DB_SHEET_CONFIG.HEADERS] },
      { valueInputOption: 'RAW' }
    );

    debugLog('データベースシート「' + sheetName + '」の初期化が完了しました。');
  } catch (e) {
    console.error('データベースシートの初期化に失敗: ' + e.message);
    throw new Error('データベースシートの初期化に失敗しました。サービスアカウントに編集者権限があるか確認してください。詳細: ' + e.message);
  }
}


// =================================================================
// サービスアカウント認証 & Sheets API ラッパー
// =================================================================

/**
 * サービスアカウントの認証トークンを取得する。
 * @returns {string} アクセストークン
 */
function getServiceAccountToken() {
  var props = PropertiesService.getScriptProperties();
  var serviceAccountCreds = JSON.parse(props.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS));

  var privateKey = serviceAccountCreds.private_key;
  var clientEmail = serviceAccountCreds.client_email;
  var tokenUrl = "https://www.googleapis.com/oauth2/v4/token";

  var jwtHeader = {
    alg: "RS256",
    typ: "JWT"
  };

  var now = Math.floor(Date.now() / 1000);
  var jwtClaimSet = {
    iss: clientEmail,
    scope: "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive",
    aud: tokenUrl,
    exp: now + 3600,
    iat: now
  };

  var encodedHeader = Utilities.base64EncodeWebSafe(JSON.stringify(jwtHeader));
  var encodedClaimSet = Utilities.base64EncodeWebSafe(JSON.stringify(jwtClaimSet));
  var signatureInput = encodedHeader + '.' + encodedClaimSet;
  var signature = Utilities.computeRsaSha256Signature(signatureInput, privateKey);
  var encodedSignature = Utilities.base64EncodeWebSafe(signature);
  var jwt = signatureInput + '.' + encodedSignature;

  var response = UrlFetchApp.fetch(tokenUrl, {
    method: "post",
    contentType: "application/x-www-form-urlencoded",
    payload: {
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion: jwt
    }
  });

  return JSON.parse(response.getContentText()).access_token;
}

/**
 * Google Sheets API v4 のための認証済みサービスオブジェクトを返す。
 * @returns {object} Sheets APIサービス
 */
function getSheetsService() {
  var accessToken = getServiceAccountToken();
  
  return {
    spreadsheets: {
      get: function(spreadsheetId) {
        var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + spreadsheetId;
        var response = UrlFetchApp.fetch(url, {
          headers: { 'Authorization': 'Bearer ' + accessToken }
        });
        return JSON.parse(response.getContentText());
      },
      values: {
        get: function(spreadsheetId, range) {
          var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + spreadsheetId + '/values/' + range;
          var response = UrlFetchApp.fetch(url, {
            headers: { 'Authorization': 'Bearer ' + accessToken }
          });
          return JSON.parse(response.getContentText());
        },
        update: function(spreadsheetId, range, body, params) {
          var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + spreadsheetId + '/values/' + range + '?valueInputOption=' + params.valueInputOption;
          var response = UrlFetchApp.fetch(url, {
            method: 'put',
            contentType: 'application/json',
            headers: { 'Authorization': 'Bearer ' + accessToken },
            payload: JSON.stringify(body)
          });
          return JSON.parse(response.getContentText());
        },
        append: function(spreadsheetId, range, body, params) {
          var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + spreadsheetId + '/values/' + range + ':append?valueInputOption=' + params.valueInputOption + '&insertDataOption=INSERT_ROWS';
           var response = UrlFetchApp.fetch(url, {
            method: 'post',
            contentType: 'application/json',
            headers: { 'Authorization': 'Bearer ' + accessToken },
            payload: JSON.stringify(body)
          });
          return JSON.parse(response.getContentText());
        }
      },
      batchUpdate: function(spreadsheetId, body) {
        var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + spreadsheetId + ':batchUpdate';
        var response = UrlFetchApp.fetch(url, {
          method: 'post',
          contentType: 'application/json',
          headers: { 'Authorization': 'Bearer ' + accessToken },
          payload: JSON.stringify(body)
        });
        return JSON.parse(response.getContentText());
      }
    }
  };
}


// =================================================================
// データベース操作関数 (サービスアカウント経由)
// =================================================================

/**
 * ユーザーIDでユーザー情報をデータベースから検索する。
 * @param {string} userId - 検索するユーザーID
 * @returns {object|null} ユーザー情報オブジェクトまたはnull
 */
function findUserById(userId) {
  var props = PropertiesService.getScriptProperties();
  var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
  var service = getSheetsService();
  var sheetName = DB_SHEET_CONFIG.SHEET_NAME;

  try {
    var data = service.spreadsheets.values.get(dbId, sheetName + '!A:H').values || [];
    var headers = data[0];
    var userIdIndex = headers.indexOf('userId');

    for (var i = 1; i < data.length; i++) {
      if (data[i][userIdIndex] === userId) {
        var user = {};
        headers.forEach(function(header, index) { user[header] = data[i][index]; });
        return user;
      }
    }
    return null;
  } catch (e) {
    console.error('findUserByIdエラー: ' + e.message);
    return null;
  }
}

/**
 * メールアドレスでユーザー情報をデータベースから検索する。
 * @param {string} email - 検索するメールアドレス
 * @returns {object|null} ユーザー情報オブジェクトまたはnull
 */
function findUserByEmail(email) {
  var props = PropertiesService.getScriptProperties();
  var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
  var service = getSheetsService();
  var sheetName = DB_SHEET_CONFIG.SHEET_NAME;

  try {
    var data = service.spreadsheets.values.get(dbId, sheetName + '!A:H').values || [];
    var headers = data[0];
    var emailIndex = headers.indexOf('adminEmail');

    for (var i = 1; i < data.length; i++) {
      if (data[i][emailIndex] === email) {
        var user = {};
        headers.forEach(function(header, index) { user[header] = data[i][index]; });
        return user;
      }
    }
    return null;
  } catch (e) {
    console.error('findUserByEmailエラー: ' + e.message);
    return null;
  }
}

/**
 * 新規ユーザーをデータベースに作成する。
 * @param {object} userData - 作成するユーザーデータ
 * @returns {object} 作成されたユーザーデータ
 */
function createUserInDb(userData) {
  var props = PropertiesService.getScriptProperties();
  var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
  var service = getSheetsService();
  var sheetName = DB_SHEET_CONFIG.SHEET_NAME;

  var newRow = DB_SHEET_CONFIG.HEADERS.map(function(header) { return userData[header] || ''; });
  
  service.spreadsheets.values.append(
    dbId,
    sheetName + '!A1',
    { values: [newRow] },
    { valueInputOption: 'RAW' }
  );
  return userData;
}

/**
 * ユーザー情報をデータベースで更新する。
 * @param {string} userId - 更新するユーザーのID
 * @param {object} updateData - 更新するデータ
 * @returns {object} 更新結果
 */
function updateUserInDb(userId, updateData) {
    var props = PropertiesService.getScriptProperties();
    var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    var service = getSheetsService();
    var sheetName = DB_SHEET_CONFIG.SHEET_NAME;

    var data = service.spreadsheets.values.get(dbId, sheetName + '!A:H').values || [];
    var headers = data[0];
    var userIdIndex = headers.indexOf('userId');
    var rowIndex = -1;

    for (var i = 1; i < data.length; i++) {
        if (data[i][userIdIndex] === userId) {
            rowIndex = i + 1;
            break;
        }
    }

    if (rowIndex === -1) {
        throw new Error('更新対象のユーザーが見つかりません。');
    }

    var requests = Object.keys(updateData).map(function(key) {
        var colIndex = headers.indexOf(key);
        if (colIndex === -1) return null;
        return {
            range: sheetName + '!' + String.fromCharCode(65 + colIndex) + rowIndex,
            values: [[updateData[key]]]
        };
    }).filter(Boolean);

    if (requests.length > 0) {
        service.spreadsheets.values.batchUpdate(dbId, {
            data: requests,
            valueInputOption: 'RAW'
        });
    }
    return { success: true };
}


// =================================================================
// メインロジック (旧Code.gsの関数群を新アーキテクチャに適応)
// =================================================================

function doGet(e) {
  var userId = e.parameter.userId;
  if (!userId) {
    return HtmlService.createTemplateFromFile('Registration').evaluate().setTitle('新規登録');
  }

  var userInfo = findUserById(userId);
  if (!userInfo) {
    return HtmlService.createHtmlOutput('無効なユーザーIDです。');
  }
  
  // ユーザーの最終アクセス日時を更新
  updateUserInDb(userId, { lastAccessedAt: new Date().toISOString() });

  // ここから先のロジックは元のdoGetとほぼ同じ
  // AdminPanelやPage.htmlをユーザー情報に基づいて表示する
  var template = HtmlService.createTemplateFromFile('Page');
  // ... テンプレートに変数を渡す処理 ...
  return template.evaluate().setTitle('みんなの回答ボード');
}

/**
 * 新規ユーザーを登録する。
 * 実行者: アクセスしたユーザー本人
 * 処理: により、Workspace管理者の設定変更を不要にし、403エラーを回避する。
 */

// =================================================================
// グローバル設定
// =================================================================

// スクリプトプロパティに保存されるキー
var SCRIPT_PROPS_KEYS = {
  SERVICE_ACCOUNT_CREDS: 'SERVICE_ACCOUNT_CREDS',
  DATABASE_SPREADSHEET_ID: 'DATABASE_SPREADSHEET_ID'
};

// ユーザー専用フォルダ管理の設定
var USER_FOLDER_CONFIG = {
  ROOT_FOLDER_NAME: "StudyQuest - ユーザーデータ",
  FOLDER_NAME_PATTERN: "StudyQuest - {email} - ファイル"
};

// 中央データベースのシート設定
var DB_SHEET_CONFIG = {
  SHEET_NAME: 'Users',
  HEADERS: [
    'userId', 'adminEmail', 'spreadsheetId', 'spreadsheetUrl', 
    'createdAt', 'configJson', 'lastAccessedAt', 'isActive'
  ]
};

// 回答ボードのスプレッドシートの列ヘッダー
var COLUMN_HEADERS = {
  TIMESTAMP: 'タイムスタンプ',
  EMAIL: 'メールアドレス',
  CLASS: 'クラス',
  OPINION: '回答',
  REASON: '理由',
  NAME: '名前',
  UNDERSTAND: 'なるほど！',
  LIKE: 'いいね！',
  CURIOUS: 'もっと知りたい！',
  HIGHLIGHT: 'ハイライト'
};

var REACTION_KEYS = ["UNDERSTAND", "LIKE", "CURIOUS"];
var EMAIL_REGEX = new RegExp("^[^
@]+@[^
@]+\.[^
@]+$");
var DEBUG = true;

function debugLog() {
  if (DEBUG && typeof console !== 'undefined' && console.log) {
    console.log.apply(console, arguments);
  }
}

// =================================================================
// セットアップ＆管理用関数
// =================================================================

/**
 * アプリケーションの初期セットアップ（管理者が手動で実行）
 * サービスアカウントの認証情報とデータベースのスプレッドシートIDを設定する。
 * @param {string} credsJson - ダウンロードしたサービスアカウントのJSONキーファイルの内容
 * @param {string} dbId - 中央データベースとして使用するスプレッドシートのID
 */
function setupApplication(credsJson, dbId) {
  try {
    JSON.parse(credsJson);
    // ★正規表現による検証を文字列長チェックに置き換え
    if (typeof dbId !== 'string' || dbId.length !== 44) {
      throw new Error('無効なスプレッドシートIDです。IDは44文字の文字列である必要があります。');
    }

    var props = PropertiesService.getScriptProperties();
    props.setProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS, credsJson);
    props.setProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID, dbId);

    // データベースシートの初期化
    initializeDatabaseSheet(dbId);

    console.log('✅ セットアップが正常に完了しました。');
  } catch (e) {
    console.error('セットアップエラー:', e);
    throw new Error('セットアップに失敗しました: ' + e.message);
  }
}

/**
 * データベースシートに必要なヘッダーを作成する。
 * @param {string} spreadsheetId - データベースのスプレッドシートID
 */
function initializeDatabaseSheet(spreadsheetId) {
  var service = getSheetsService();
  var sheetName = DB_SHEET_CONFIG.SHEET_NAME;

  try {
    // シートが存在するか確認
    var spreadsheet = service.spreadsheets.get(spreadsheetId);
    var sheetExists = spreadsheet.sheets.some(function(s) { return s.properties.title === sheetName; });

    if (!sheetExists) {
      // シートが存在しない場合は作成
      service.spreadsheets.batchUpdate(spreadsheetId, {
        requests: [{ addSheet: { properties: { title: sheetName } } }]
      });
    }
    
    // ヘッダーを書き込み
    var headerRange = sheetName + '!A1:' + String.fromCharCode(65 + DB_SHEET_CONFIG.HEADERS.length - 1) + '1';
    service.spreadsheets.values.update(
      spreadsheetId,
      headerRange,
      { values: [DB_SHEET_CONFIG.HEADERS] },
      { valueInputOption: 'RAW' }
    );

    debugLog('データベースシート「' + sheetName + '」の初期化が完了しました。');
  } catch (e) {
    console.error('データベースシートの初期化に失敗: ' + e.message);
    throw new Error('データベースシートの初期化に失敗しました。サービスアカウントに編集者権限があるか確認してください。詳細: ' + e.message);
  }
}


// =================================================================
// サービスアカウント認証 & Sheets API ラッパー
// =================================================================

/**
 * サービスアカウントの認証トークンを取得する。
 * @returns {string} アクセストークン
 */
function getServiceAccountToken() {
  var props = PropertiesService.getScriptProperties();
  var serviceAccountCreds = JSON.parse(props.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS));

  var privateKey = serviceAccountCreds.private_key;
  var clientEmail = serviceAccountCreds.client_email;
  var tokenUrl = "https://www.googleapis.com/oauth2/v4/token";

  var jwtHeader = {
    alg: "RS256",
    typ: "JWT"
  };

  var now = Math.floor(Date.now() / 1000);
  var jwtClaimSet = {
    iss: clientEmail,
    scope: "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive",
    aud: tokenUrl,
    exp: now + 3600,
    iat: now
  };

  var encodedHeader = Utilities.base64EncodeWebSafe(JSON.stringify(jwtHeader));
  var encodedClaimSet = Utilities.base64EncodeWebSafe(JSON.stringify(jwtClaimSet));
  var signatureInput = encodedHeader + '.' + encodedClaimSet;
  var signature = Utilities.computeRsaSha256Signature(signatureInput, privateKey);
  var encodedSignature = Utilities.base64EncodeWebSafe(signature);
  var jwt = signatureInput + '.' + encodedSignature;

  var response = UrlFetchApp.fetch(tokenUrl, {
    method: "post",
    contentType: "application/x-www-form-urlencoded",
    payload: {
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion: jwt
    }
  });

  return JSON.parse(response.getContentText()).access_token;
}

/**
 * Google Sheets API v4 のための認証済みサービスオブジェクトを返す。
 * @returns {object} Sheets APIサービス
 */
function getSheetsService() {
  var accessToken = getServiceAccountToken();
  
  return {
    spreadsheets: {
      get: function(spreadsheetId) {
        var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + spreadsheetId;
        var response = UrlFetchApp.fetch(url, {
          headers: { 'Authorization': 'Bearer ' + accessToken }
        });
        return JSON.parse(response.getContentText());
      },
      values: {
        get: function(spreadsheetId, range) {
          var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + spreadsheetId + '/values/' + range;
          var response = UrlFetchApp.fetch(url, {
            headers: { 'Authorization': 'Bearer ' + accessToken }
          });
          return JSON.parse(response.getContentText());
        },
        update: function(spreadsheetId, range, body, params) {
          var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + spreadsheetId + '/values/' + range + '?valueInputOption=' + params.valueInputOption;
          var response = UrlFetchApp.fetch(url, {
            method: 'put',
            contentType: 'application/json',
            headers: { 'Authorization': 'Bearer ' + accessToken },
            payload: JSON.stringify(body)
          });
          return JSON.parse(response.getContentText());
        },
        append: function(spreadsheetId, range, body, params) {
          var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + spreadsheetId + '/values/' + range + ':append?valueInputOption=' + params.valueInputOption + '&insertDataOption=INSERT_ROWS';
           var response = UrlFetchApp.fetch(url, {
            method: 'post',
            contentType: 'application/json',
            headers: { 'Authorization': 'Bearer ' + accessToken },
            payload: JSON.stringify(body)
          });
          return JSON.parse(response.getContentText());
        }
      },
      batchUpdate: function(spreadsheetId, body) {
        var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + spreadsheetId + ':batchUpdate';
        var response = UrlFetchApp.fetch(url, {
          method: 'post',
          contentType: 'application/json',
          headers: { 'Authorization': 'Bearer ' + accessToken },
          payload: JSON.stringify(body)
        });
        return JSON.parse(response.getContentText());
      }
    }
  };
}


// =================================================================
// データベース操作関数 (サービスアカウント経由)
// =================================================================

/**
 * ユーザーIDでユーザー情報をデータベースから検索する。
 * @param {string} userId - 検索するユーザーID
 * @returns {object|null} ユーザー情報オブジェクトまたはnull
 */
function findUserById(userId) {
  var props = PropertiesService.getScriptProperties();
  var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
  var service = getSheetsService();
  var sheetName = DB_SHEET_CONFIG.SHEET_NAME;

  try {
    var data = service.spreadsheets.values.get(dbId, sheetName + '!A:H').values || [];
    var headers = data[0];
    var userIdIndex = headers.indexOf('userId');

    for (var i = 1; i < data.length; i++) {
      if (data[i][userIdIndex] === userId) {
        var user = {};
        headers.forEach(function(header, index) { user[header] = data[i][index]; });
        return user;
      }
    }
    return null;
  } catch (e) {
    console.error('findUserByIdエラー: ' + e.message);
    return null;
  }
}

/**
 * メールアドレスでユーザー情報をデータベースから検索する。
 * @param {string} email - 検索するメールアドレス
 * @returns {object|null} ユーザー情報オブジェクトまたはnull
 */
function findUserByEmail(email) {
  var props = PropertiesService.getScriptProperties();
  var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
  var service = getSheetsService();
  var sheetName = DB_SHEET_CONFIG.SHEET_NAME;

  try {
    var data = service.spreadsheets.values.get(dbId, sheetName + '!A:H').values || [];
    var headers = data[0];
    var emailIndex = headers.indexOf('adminEmail');

    for (var i = 1; i < data.length; i++) {
      if (data[i][emailIndex] === email) {
        var user = {};
        headers.forEach(function(header, index) { user[header] = data[i][index]; });
        return user;
      }
    }
    return null;
  } catch (e) {
    console.error('findUserByEmailエラー: ' + e.message);
    return null;
  }
}

/**
 * 新規ユーザーをデータベースに作成する。
 * @param {object} userData - 作成するユーザーデータ
 * @returns {object} 作成されたユーザーデータ
 */
function createUserInDb(userData) {
  var props = PropertiesService.getScriptProperties();
  var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
  var service = getSheetsService();
  var sheetName = DB_SHEET_CONFIG.SHEET_NAME;

  var newRow = DB_SHEET_CONFIG.HEADERS.map(function(header) { return userData[header] || ''; });
  
  service.spreadsheets.values.append(
    dbId,
    sheetName + '!A1',
    { values: [newRow] },
    { valueInputOption: 'RAW' }
  );
  return userData;
}

/**
 * ユーザー情報をデータベースで更新する。
 * @param {string} userId - 更新するユーザーのID
 * @param {object} updateData - 更新するデータ
 * @returns {object} 更新結果
 */
function updateUserInDb(userId, updateData) {
    var props = PropertiesService.getScriptProperties();
    var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    var service = getSheetsService();
    var sheetName = DB_SHEET_CONFIG.SHEET_NAME;

    var data = service.spreadsheets.values.get(dbId, sheetName + '!A:H').values || [];
    var headers = data[0];
    var userIdIndex = headers.indexOf('userId');
    var rowIndex = -1;

    for (var i = 1; i < data.length; i++) {
        if (data[i][userIdIndex] === userId) {
            rowIndex = i + 1;
            break;
        }
    }

    if (rowIndex === -1) {
        throw new Error('更新対象のユーザーが見つかりません。');
    }

    var requests = Object.keys(updateData).map(function(key) {
        var colIndex = headers.indexOf(key);
        if (colIndex === -1) return null;
        return {
            range: sheetName + '!' + String.fromCharCode(65 + colIndex) + rowIndex,
            values: [[updateData[key]]]
        };
    }).filter(Boolean);

    if (requests.length > 0) {
        service.spreadsheets.values.batchUpdate(dbId, {
            data: requests,
            valueInputOption: 'RAW'
        });
    }
    return { success: true };
}


// =================================================================
// メインロジック (旧Code.gsの関数群を新アーキテクチャに適応)
// =================================================================

function doGet(e) {
  var userId = e.parameter.userId;
  if (!userId) {
    return HtmlService.createTemplateFromFile('Registration').evaluate().setTitle('新規登録');
  }

  var userInfo = findUserById(userId);
  if (!userInfo) {
    return HtmlService.createHtmlOutput('無効なユーザーIDです。');
  }
  
  // ユーザーの最終アクセス日時を更新
  updateUserInDb(userId, { lastAccessedAt: new Date().toISOString() });

  // ここから先のロジックは元のdoGetとほぼ同じ
  // AdminPanelやPage.htmlをユーザー情報に基づいて表示する
  var template = HtmlService.createTemplateFromFile('Page');
  // ... テンプレートに変数を渡す処理 ...
  return template.evaluate().setTitle('みんなの回答ボード');
}

/**
 * 新規ユーザーを登録する。
 * 実行者: アクセスしたユーザー本人
 * 処理:
 * 1. ユーザー自身の権限でフォームとスプレッドシートを作成する。
 * 2. サービスアカウント経由で中央データベースにユーザー情報を登録する。
 */
function registerNewUser(adminEmail) {
  var activeUser = Session.getActiveUser();
  if (adminEmail !== activeUser.getEmail()) {
    throw new Error('認証エラー: 操作を実行しているユーザーとメールアドレスが一致しません。');
  }

  // ステップ1: ユーザー自身の権限でファイル作成
  var userId = Utilities.getUuid();
  var formAndSsInfo = createStudyQuestForm(adminEmail, userId); // この関数は元のままでOK

  // ステップ2: サービスアカウント経由でDBに登録
  var initialConfig = {
    formUrl: formAndSsInfo.viewFormUrl || formAndSsInfo.formUrl,
    editFormUrl: formAndSsInfo.editFormUrl,
    createdAt: new Date().toISOString()
  };
  
  var userData = {
    userId: userId,
    adminEmail: adminEmail,
    spreadsheetId: formAndSsInfo.spreadsheetId,
    spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
    createdAt: new Date().toISOString(),
    configJson: JSON.stringify(initialConfig),
    lastAccessedAt: new Date().toISOString(),
    isActive: true
  };

  try {
    createUserInDb(userData);
    debugLog('✅ データベースに新規ユーザーを登録しました: ' + adminEmail);
  } catch (e) {
    console.error('データベースへのユーザー登録に失敗: ' + e.message);
    // ここで作成したフォームなどを削除するクリーンアップ処理を入れるのが望ましい
    throw new Error('ユーザー登録に失敗しました。システム管理者に連絡してください。');
  }

  // 成功レスポンスを返す (元のコードと同様)
  var webAppUrl = ScriptApp.getService().getUrl();
  return {
    userId: userId,
    spreadsheetId: formAndSsInfo.spreadsheetId,
    adminUrl: webAppUrl + '?userId=' + userId + '&mode=admin',
    viewUrl: webAppUrl + '?userId=' + userId,
    message: '新しいボードが作成されました！'
  };
}

/**
 * リアクションを追加/削除する。
 * 実行者: アクセスしたユーザー本人
 * 処理:
 * 1. リアクションしたユーザーのメールアドレスを取得する。
 * 2. サービスアカウント経由で、対象のユーザーのスプレッドシートを更新する。
 *    (注意: この実装は簡略化しています。実際にはユーザーのスプレッドシートに
 *     サービスアカウントを編集者として追加するフローが必要になります)
 */
function addReaction(rowIndex, reactionKey, sheetName) {
  var reactingUserEmail = Session.getActiveUser().getEmail();
  var props = PropertiesService.getUserProperties(); // doGetで設定されたコンテキストから取得
  var ownerUserId = props.getProperty('CURRENT_USER_ID');

  if (!ownerUserId) {
    throw new Error('ボードのオーナー情報が見つかりません。');
  }

  // ボードオーナーの情報をDBから取得
  var boardOwnerInfo = findUserById(ownerUserId);
  if (!boardOwnerInfo) {
    throw new Error('無効なボードです。');
  }

  // ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
  // 重要：このアーキテクチャの課題点
  // サービスアカウントは、ユーザーが作成した個々のスプレッドシートに
  // アクセスする権限をデフォルトでは持っていません。
  // この問題を解決するには、ユーザーがボードを作成した際に、そのスプレッドシートに
  // サービスアカウントのメールアドレスを「編集者」として追加する処理が必要です。
  // ここでは、その処理が実装済みであると仮定して進めます。
  // ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★

  var targetSpreadsheetId = boardOwnerInfo.spreadsheetId;
  
  // ここに、サービスアカウントを使って targetSpreadsheetId の
  // リアクション列を更新するロジックを実装します。
  // (この部分は非常に複雑になるため、概念的な実装に留めます)

  debugLog('ユーザー ' + reactingUserEmail + ' がシート ' + sheetName + ' の ' + rowIndex + ' 行目に ' + reactionKey + ' リアクションを追加/削除しました。');

  // LockServiceを使って競合を防ぐ
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // サービスアカウント経由でスプレッドシートを開く
    var service = getSheetsService();
    var range = sheetName + '!A' + rowIndex + ':Z' + rowIndex; // 仮の範囲
    var values = service.spreadsheets.values.get(targetSpreadsheetId, range).values;

    // ... (元のaddReactionのロジックをここに移植し、valuesを操作) ...
    // 例: values[0][columnIndex] = updatedReactionString;

    // 更新された値を書き戻す
    service.spreadsheets.values.update(
      targetSpreadsheetId,
      range,
      { values: values },
      { valueInputOption: 'RAW' }
    );

  } finally {
    lock.releaseLock();
  }

  return { status: 'ok', message: 'リアクションを更新しました。' };
}


// createStudyQuestForm, getWebAppUrlEnhanced などのヘルパー関数は元のままで流用可能
// ただし、API呼び出しに依存している部分はすべて修正が必要

// 以下、元のCode.gsから必要な関数を移植・修正する
// (例: createStudyQuestForm, getWebAppUrlEnhanced, etc.)

function getWebAppUrlEnhanced() {
  return ScriptApp.getService().getUrl();
}

function createStudyQuestForm(userEmail, userId) {
  // この関数は、実行ユーザー自身の権限で動作するため、元のままで問題ありません。
  // ... (元の createStudyQuestForm のコードをここに貼り付け) ...
  // ただし、内部で getWebAppUrlEnhanced を呼んでいることを確認
  var now = new Date();
  var dateTimeString = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
  var formTitle = 'StudyQuest - みんなの回答ボード - ' + userEmail.split('@')[0] + ' - ' + dateTimeString;
  var form = FormApp.create(formTitle);
  
  form.setCollectEmail(true);
  form.setRequireLogin(true);
  try {
    if (typeof form.setEmailCollectionType === 'function') {
      form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED);
    }
  } catch (undocumentedError) {
    // ignore
  }
  form.setLimitOneResponsePerUser(true);
  form.setAllowResponseEdits(true);

  var boardUrl = '';
  try {
    var webAppUrl = getWebAppUrlEnhanced();
    if (webAppUrl) {
      boardUrl = webAppUrl + '?userId=' + userId;
    }
  } catch (e) {
    // ignore
  }
  var confirmationMessage = boardUrl 
    ? '🎉 回答ありがとうございます！\n\nあなたの大切な意見が届きました。\nみんなの回答ボードで、お友達の色々な考えも見てみましょう。\n新しい発見があるかもしれませんね！\n\n' + boardUrl
    : '🎉 回答ありがとうございます！\n\nあなたの大切な意見が届きました。';
  form.setConfirmationMessage(confirmationMessage);

  var classItem = form.addTextItem();
  classItem.setTitle('クラス名');
  classItem.setRequired(true);
  var pattern = '^[A-Za-z0-9]+-[A-Za-z0-9]+$';
  var helpText = "【重要】クラス名は決められた形式で入力してください。\n\n✅ 正しい例：\n• 6年1組 → 6-1\n• 5年2組 → 5-2  \n• 中1年A組 → 1-A\n• 中3年B組 → 3-B\n\n❌ 間違いの例：6年1組、6-1組、６－１\n\n※ 半角英数字とハイフン（-）のみ使用可能です";
  var textValidation = FormApp.createTextValidation()
    .setHelpText(helpText)
    .requireTextMatchesPattern(pattern)
    .build();
  classItem.setValidation(textValidation);

  var nameItem = form.addTextItem();
  nameItem.setTitle('名前');
  nameItem.setHelpText('あなたの名前を入力してください。この名前は先生だけが見ることができ、みんなには表示されません。一人ひとりの意見を大切にするために必要です。');
  nameItem.setRequired(true);
  
  var answerItem = form.addParagraphTextItem();
  answerItem.setTitle('回答');
  answerItem.setHelpText('質問に対するあなたの考えを、自分の言葉で表現してください。正解はありません。あなたらしい考えや感じ方を大切にして、思考力を育てましょう。');
  answerItem.setRequired(true);
  
  var reasonItem = form.addParagraphTextItem();
  reasonItem.setTitle('理由');
  reasonItem.setHelpText('なぜそう思ったのか、根拠や理由を説明してください。論理的に考える力を身につけ、自分の意見に責任を持つ習慣を育てましょう。');
  reasonItem.setRequired(false);
  
  var spreadsheetTitle = 'StudyQuest - みんなの回答ボード - 回答データ - ' + userEmail.split('@')[0] + ' - ' + dateTimeString;
  var spreadsheet = SpreadsheetApp.create(spreadsheetTitle);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());

  // ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
  // 【最重要】ここで、作成したスプレッドシートにサービスアカウントを編集者として追加する
  // ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
  try {
    var props = PropertiesService.getScriptProperties();
    var serviceAccountCreds = JSON.parse(props.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS));
    var serviceAccountEmail = serviceAccountCreds.client_email;
    if (serviceAccountEmail) {
      spreadsheet.addEditor(serviceAccountEmail);
      debugLog('サービスアカウント (' + serviceAccountEmail + ') をスプレッドシートの編集者として追加しました。');
    }
  } catch (e) {
    console.error('サービスアカウントの追加に失敗: ' + e.message);
    // このエラーは致命的ではないが、リアクション機能などが動作しなくなるためログに残す
  }

  var sheet = spreadsheet.getSheets()[0];
  var additionalHeaders = [
    COLUMN_HEADERS.UNDERSTAND,
    COLUMN_HEADERS.LIKE,
    COLUMN_HEADERS.CURIOUS,
    COLUMN_HEADERS.HIGHLIGHT
  ];
  var currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var startCol = currentHeaders.length + 1;
  sheet.getRange(1, startCol, 1, additionalHeaders.length).setValues([additionalHeaders]);
  
  var allHeadersRange = sheet.getRange(1, 1, 1, currentHeaders.length + additionalHeaders.length);
  allHeadersRange.setFontWeight('bold').setBackground('#E3F2FD');
  try {
    sheet.autoResizeColumns(1, allHeadersRange.getNumColumns());
  } catch (e) {
    console.warn('Auto-resize failed:', e);
  }

  return {
      formId: form.getId(),
      formUrl: form.getPublishedUrl(),
      spreadsheetId: spreadsheet.getId(),
      spreadsheetUrl: spreadsheet.getUrl(),
      editFormUrl: form.getEditUrl(),
      viewFormUrl: 'https://docs.google.com/forms/d/e/' + form.getId() + '/viewform'
  };
}

// その他のヘルパー関数（元のCode.gsから必要に応じて移植）
function isValidEmail(email) {
  if (!email || typeof email !== 'string') {
    return false;
  }
  return EMAIL_REGEX.test(email);
}

function getEmailDomain(email) {
  return (email || '').toString().split('@').pop().toLowerCase();
}

function safeGetUserEmail() {
  try {
    var email = Session.getActiveUser().getEmail();
    if (!email || typeof email !== 'string' || !isValidEmail(email)) {
      throw new Error('Invalid user email');
    }
    return email;
  } catch (e) {
    console.error('Failed to get user email:', e);
    throw new Error('認証が必要です。Googleアカウントでログインしてください。');
  }
}

function parseReactionString(val) {
  if (!val) return [];
  return val
    .toString()
    .split(',')
    .map(function(s) { return s.trim(); })
    .filter(Boolean);
}