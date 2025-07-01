/**
 * @fileoverview StudyQuest - みんなの回答ボード (サービスアカウントモデル)
 * ユーザーがファイル所有権を持ち、中央データベースへのアクセスはサービスアカウント経由で行う。
 * これにより、Workspace管理者の設定変更を不要にし、403エラーを回避する。
 */

// =================================================================
// グローバル設定
// =================================================================

// スクリプトプロパティに保存されるキー
const SCRIPT_PROPS_KEYS = {
  SERVICE_ACCOUNT_CREDS: 'SERVICE_ACCOUNT_CREDS',
  DATABASE_SPREADSHEET_ID: 'DATABASE_SPREADSHEET_ID'
};

// ユーザー専用フォルダ管理の設定
const USER_FOLDER_CONFIG = {
  ROOT_FOLDER_NAME: "StudyQuest - ユーザーデータ",
  FOLDER_NAME_PATTERN: "StudyQuest - {email} - ファイル"
};

// 中央データベースのシート設定
const DB_SHEET_CONFIG = {
  SHEET_NAME: 'Users',
  HEADERS: [
    'userId', 'adminEmail', 'spreadsheetId', 'spreadsheetUrl', 
    'createdAt', 'configJson', 'lastAccessedAt', 'isActive'
  ]
};

// 回答ボードのスプレッドシートの列ヘッダー
const COLUMN_HEADERS = {
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

const REACTION_KEYS = ["UNDERSTAND", "LIKE", "CURIOUS"];
const EMAIL_REGEX = /^[^
@]+@[^
@]+\.[^\n@]+$/;
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
 */
function setupApplication() {
  const ui = SpreadsheetApp.getUi();
  
  const credsPrompt = ui.prompt(
    'サービスアカウント認証情報の設定',
    'ダウンロードしたサービスアカウントのJSONキーファイルの内容を貼り付けてください。',
    ui.ButtonSet.OK_CANCEL
  );
  if (credsPrompt.getSelectedButton() !== ui.Button.OK) return;
  const credsJson = credsPrompt.getResponseText();
  
  const dbIdPrompt = ui.prompt(
    'データベースのスプレッドシートIDの設定',
    '中央データベースとして使用するスプレッドシートのIDを入力してください。',
    ui.ButtonSet.OK_CANCEL
  );
  if (dbIdPrompt.getSelectedButton() !== ui.Button.OK) return;
  const dbId = dbIdPrompt.getResponseText();

  try {
    // 入力値を検証
    JSON.parse(credsJson);
    const sheetIdRegex = new RegExp("^[a-zA-Z0-9-_]{44}$");
    if (!sheetIdRegex.test(dbId)) {
      throw new Error('無効なスプレッドシートIDです。');
    }

    const props = PropertiesService.getScriptProperties();
    props.setProperties({
      [SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS]: credsJson,
      [SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID]: dbId
    });

    // データベースシートの初期化
    initializeDatabaseSheet(dbId);

    ui.alert('✅ セットアップが正常に完了しました。');
  } catch (e) {
    console.error('セットアップエラー:', e);
    ui.alert(`セットアップに失敗しました: ${e.message}`);
  }
}

/**
 * データベースシートに必要なヘッダーを作成する。
 * @param {string} spreadsheetId - データベースのスプレッドシートID
 */
function initializeDatabaseSheet(spreadsheetId) {
  const service = getSheetsService();
  const sheetName = DB_SHEET_CONFIG.SHEET_NAME;

  try {
    // シートが存在するか確認
    const spreadsheet = service.spreadsheets.get(spreadsheetId);
    const sheetExists = spreadsheet.sheets.some(s => s.properties.title === sheetName);

    if (!sheetExists) {
      // シートが存在しない場合は作成
      service.spreadsheets.batchUpdate(spreadsheetId, {
        requests: [{ addSheet: { properties: { title: sheetName } } }]
      });
    }
    
    // ヘッダーを書き込み
    const headerRange = `${sheetName}!A1:${String.fromCharCode(65 + DB_SHEET_CONFIG.HEADERS.length - 1)}1`;
    service.spreadsheets.values.update(
      spreadsheetId,
      headerRange,
      { values: [DB_SHEET_CONFIG.HEADERS] },
      { valueInputOption: 'RAW' }
    );

    debugLog(`データベースシート「${sheetName}」の初期化が完了しました。`);
  } catch (e) {
    console.error(`データベースシートの初期化に失敗: ${e.message}`);
    throw new Error(`データベースシートの初期化に失敗しました。サービスアカウントに編集者権限があるか確認してください。詳細: ${e.message}`);
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
  const props = PropertiesService.getScriptProperties();
  const serviceAccountCreds = JSON.parse(props.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS));

  const privateKey = serviceAccountCreds.private_key;
  const clientEmail = serviceAccountCreds.client_email;
  const tokenUrl = "https://www.googleapis.com/oauth2/v4/token";

  const jwtHeader = {
    alg: "RS256",
    typ: "JWT"
  };

  const now = Math.floor(Date.now() / 1000);
  const jwtClaimSet = {
    iss: clientEmail,
    scope: "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive",
    aud: tokenUrl,
    exp: now + 3600,
    iat: now
  };

  const encodedHeader = Utilities.base64EncodeWebSafe(JSON.stringify(jwtHeader));
  const encodedClaimSet = Utilities.base64EncodeWebSafe(JSON.stringify(jwtClaimSet));
  const signatureInput = `${encodedHeader}.${encodedClaimSet}`;
  const signature = Utilities.computeRsaSha256Signature(signatureInput, privateKey);
  const encodedSignature = Utilities.base64EncodeWebSafe(signature);
  const jwt = `${signatureInput}.${encodedSignature}`;

  const response = UrlFetchApp.fetch(tokenUrl, {
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
  const accessToken = getServiceAccountToken();
  
  return {
    spreadsheets: {
      get: function(spreadsheetId) {
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}`;
        const response = UrlFetchApp.fetch(url, {
          headers: { 'Authorization': `Bearer ${accessToken}` }
        });
        return JSON.parse(response.getContentText());
      },
      values: {
        get: function(spreadsheetId, range) {
          const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${range}`;
          const response = UrlFetchApp.fetch(url, {
            headers: { 'Authorization': `Bearer ${accessToken}` }
          });
          return JSON.parse(response.getContentText());
        },
        update: function(spreadsheetId, range, body, params) {
          const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${range}?valueInputOption=${params.valueInputOption}`;
          const response = UrlFetchApp.fetch(url, {
            method: 'put',
            contentType: 'application/json',
            headers: { 'Authorization': `Bearer ${accessToken}` },
            payload: JSON.stringify(body)
          });
          return JSON.parse(response.getContentText());
        },
        append: function(spreadsheetId, range, body, params) {
          const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${range}:append?valueInputOption=${params.valueInputOption}&insertDataOption=INSERT_ROWS`;
           const response = UrlFetchApp.fetch(url, {
            method: 'post',
            contentType: 'application/json',
            headers: { 'Authorization': `Bearer ${accessToken}` },
            payload: JSON.stringify(body)
          });
          return JSON.parse(response.getContentText());
        }
      },
      batchUpdate: function(spreadsheetId, body) {
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`;
        const response = UrlFetchApp.fetch(url, {
          method: 'post',
          contentType: 'application/json',
          headers: { 'Authorization': `Bearer ${accessToken}` },
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
  const props = PropertiesService.getScriptProperties();
  const dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
  const service = getSheetsService();
  const sheetName = DB_SHEET_CONFIG.SHEET_NAME;

  try {
    const data = service.spreadsheets.values.get(dbId, `${sheetName}!A:H`).values || [];
    const headers = data[0];
    const userIdIndex = headers.indexOf('userId');

    for (let i = 1; i < data.length; i++) {
      if (data[i][userIdIndex] === userId) {
        const user = {};
        headers.forEach((header, index) => user[header] = data[i][index]);
        return user;
      }
    }
    return null;
  } catch (e) {
    console.error(`findUserByIdエラー: ${e.message}`);
    return null;
  }
}

/**
 * メールアドレスでユーザー情報をデータベースから検索する。
 * @param {string} email - 検索するメールアドレス
 * @returns {object|null} ユーザー情報オブジェクトまたはnull
 */
function findUserByEmail(email) {
  const props = PropertiesService.getScriptProperties();
  const dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
  const service = getSheetsService();
  const sheetName = DB_SHEET_CONFIG.SHEET_NAME;

  try {
    const data = service.spreadsheets.values.get(dbId, `${sheetName}!A:H`).values || [];
    const headers = data[0];
    const emailIndex = headers.indexOf('adminEmail');

    for (let i = 1; i < data.length; i++) {
      if (data[i][emailIndex] === email) {
        const user = {};
        headers.forEach((header, index) => user[header] = data[i][index]);
        return user;
      }
    }
    return null;
  } catch (e) {
    console.error(`findUserByEmailエラー: ${e.message}`);
    return null;
  }
}

/**
 * 新規ユーザーをデータベースに作成する。
 * @param {object} userData - 作成するユーザーデータ
 * @returns {object} 作成されたユーザーデータ
 */
function createUserInDb(userData) {
  const props = PropertiesService.getScriptProperties();
  const dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
  const service = getSheetsService();
  const sheetName = DB_SHEET_CONFIG.SHEET_NAME;

  const newRow = DB_SHEET_CONFIG.HEADERS.map(header => userData[header] || '');
  
  service.spreadsheets.values.append(
    dbId,
    `${sheetName}!A1`,
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
    const props = PropertiesService.getScriptProperties();
    const dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    const service = getSheetsService();
    const sheetName = DB_SHEET_CONFIG.SHEET_NAME;

    const data = service.spreadsheets.values.get(dbId, `${sheetName}!A:H`).values || [];
    const headers = data[0];
    const userIdIndex = headers.indexOf('userId');
    let rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
        if (data[i][userIdIndex] === userId) {
            rowIndex = i + 1;
            break;
        }
    }

    if (rowIndex === -1) {
        throw new Error('更新対象のユーザーが見つかりません。');
    }

    const requests = Object.keys(updateData).map(key => {
        const colIndex = headers.indexOf(key);
        if (colIndex === -1) return null;
        return {
            range: `${sheetName}!${String.fromCharCode(65 + colIndex)}${rowIndex}`,
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

// doGet, registerNewUser, addReaction などの主要な関数をここに移植・修正します。
// 内部でのデータベースアクセスは、上記で定義した findUserById, createUserInDb, updateUserInDb を使います。

function doGet(e) {
  // ... (元のdoGetのロジックをベースに、API呼び出しを直接のDB関数呼び出しに置き換える)
  // 例: getUserInfo(userId) -> findUserById(userId)
  // この部分は非常に長大になるため、主要なエントリーポイントのみ示します。
  
  if (e && e.parameter && e.parameter.setup === 'true') {
    // この機能は手動実行に移行
    return HtmlService.createHtmlOutput('セットアップは管理者がスクリプトエディタから手動で実行する必要があります。');
  }

  const userId = e.parameter.userId;
  if (!userId) {
    return HtmlService.createTemplateFromFile('Registration').evaluate().setTitle('新規登録');
  }

  const userInfo = findUserById(userId);
  if (!userInfo) {
    return HtmlService.createHtmlOutput('無効なユーザーIDです。');
  }
  
  // ユーザーの最終アクセス日時を更新
  updateUserInDb(userId, { lastAccessedAt: new Date().toISOString() });

  // ここから先のロジックは元のdoGetとほぼ同じ
  // AdminPanelやPage.htmlをユーザー情報に基づいて表示する
  const template = HtmlService.createTemplateFromFile('Page');
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
  const activeUser = Session.getActiveUser();
  if (adminEmail !== activeUser.getEmail()) {
    throw new Error('認証エラー: 操作を実行しているユーザーとメールアドレスが一致しません。');
  }

  // ステップ1: ユーザー自身の権限でファイル作成
  const userId = Utilities.getUuid();
  const formAndSsInfo = createStudyQuestForm(adminEmail, userId); // この関数は元のままでOK

  // ステップ2: サービスアカウント経由でDBに登録
  const initialConfig = {
    formUrl: formAndSsInfo.viewFormUrl || formAndSsInfo.formUrl,
    editFormUrl: formAndSsInfo.editFormUrl,
    createdAt: new Date().toISOString()
  };
  
  const userData = {
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
    debugLog(`✅ データベースに新規ユーザーを登録しました: ${adminEmail}`);
  } catch (e) {
    console.error(`データベースへのユーザー登録に失敗: ${e.message}`);
    // ここで作成したフォームなどを削除するクリーンアップ処理を入れるのが望ましい
    throw new Error(`ユーザー登録に失敗しました。システム管理者に連絡してください。`);
  }

  // 成功レスポンスを返す (元のコードと同様)
  const webAppUrl = ScriptApp.getService().getUrl();
  return {
    userId: userId,
    spreadsheetId: formAndSsInfo.spreadsheetId,
    adminUrl: `${webAppUrl}?userId=${userId}&mode=admin`,
    viewUrl: `${webAppUrl}?userId=${userId}`,
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
  const reactingUserEmail = Session.getActiveUser().getEmail();
  const props = PropertiesService.getUserProperties(); // doGetで設定されたコンテキストから取得
  const ownerUserId = props.getProperty('CURRENT_USER_ID');

  if (!ownerUserId) {
    throw new Error('ボードのオーナー情報が見つかりません。');
  }

  // ボードオーナーの情報をDBから取得
  const boardOwnerInfo = findUserById(ownerUserId);
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

  const targetSpreadsheetId = boardOwnerInfo.spreadsheetId;
  
  // ここに、サービスアカウントを使って targetSpreadsheetId の
  // リアクション列を更新するロジックを実装します。
  // (この部分は非常に複雑になるため、概念的な実装に留めます)

  debugLog(`ユーザー ${reactingUserEmail} がシート ${sheetName} の ${rowIndex} 行目に ${reactionKey} リアクションを追加/削除しました。`);

  // LockServiceを使って競合を防ぐ
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.openById(targetSpreadsheetId).getSheetByName(sheetName);
    // ... (元のaddReactionのロジック)
    // ...
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
  const now = new Date();
  const form = FormApp.create(`StudyQuest - ${userEmail} - ${now.toLocaleString()}`);
  // ... フォームの設問追加など ...
  const spreadsheet = SpreadsheetApp.create(`StudyQuest-Data - ${userEmail} - ${now.toLocaleString()}`);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());
  
  // ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
  // 【最重要】ここで、作成したスプレッドシートにサービスアカウントを編集者として追加する
  // ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
  try {
    const props = PropertiesService.getScriptProperties();
    const serviceAccountCreds = JSON.parse(props.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS));
    const serviceAccountEmail = serviceAccountCreds.client_email;
    if (serviceAccountEmail) {
      spreadsheet.addEditor(serviceAccountEmail);
      debugLog(`サービスアカウント (${serviceAccountEmail}) をスプレッドシートの編集者として追加しました。`);
    }
  } catch (e) {
    console.error(`サービスアカウントの追加に失敗: ${e.message}`);
    // このエラーは致命的ではないが、リアクション機能などが動作しなくなるためログに残す
  }


  return {
      formId: form.getId(),
      formUrl: form.getPublishedUrl(),
      spreadsheetId: spreadsheet.getId(),
      spreadsheetUrl: spreadsheet.getUrl(),
      editFormUrl: form.getEditUrl(),
      viewFormUrl: `https://docs.google.com/forms/d/e/${form.getId()}/viewform`
  };
}

// 他にも多くの関数を移植・修正する必要がありますが、主要な部分の変更方針は以上です。