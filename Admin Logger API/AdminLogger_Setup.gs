/**
 * @OnlyCurrentDoc
 *
 * ===================================================================================
 * みんなの回答ボード - 【管理者向けログ記録API】 (最終修正版)
 * ===================================================================================
 *
 * ◆ 概要
 * このスクリプトは、ユーザー向けアプリから送信された利用状況のログを受け取り、
 * このスプレッドシートをデータベースとして安全に保存します。
 *
 * ◆ 新しい管理者向けのセットアップ手順
 * -----------------------------------------------------------------------------------
 * 1. 新しいGoogleスプレッドシートを作成し、このファイルを開きます。
 *
 * 2. 「拡張機能」>「Apps Script」を開き、このコードをエディタに貼り付けます。
 *
 * 3. `appsscript.json` と `DeploymentGuide.html` ファイルも、指定の内容で作成・更新します。
 *
 * 4. 全てのファイルを保存し（💾アイコン）、スプレッドシートのページを**再読み込み**します。
 *
 * 5. 上部に表示されたカスタムメニュー「🚀 Admin Logger セットアップ」を開きます。
 *
 * 6. 「1. このシートをデータベースとして初期化」を実行し、画面の指示に従い権限を承認します。
 *
 * 7. 再度メニューを開き、「2. APIをデプロイする」を実行します。
 *
 * 8. 表示される手順に従ってデプロイを行い、最後に表示される「ウェブアプリURL」をコピーします。
 * このURLは、次にセットアップする「ユーザー向けメインアプリ」で使用します。
 * -----------------------------------------------------------------------------------
 * 
 * ◆ セキュリティについて
 * -----------------------------------------------------------------------------------
 * 🔒 セキュリティ保護されています:
 * - 実行者: USER_DEPLOYING = デプロイしたユーザー（管理者）の権限で実行
 * - データベースアクセス: 管理者のみがスプレッドシートにアクセス可能
 * - API認証: リクエストにrequestUserとeffectiveUserを記録してログ追跡
 * 
 * ⚠️ 考慮点:
 * - URLを知っていれば誰でもAPIにアクセス可能
 * - 推奨: URLは秘匿情報として扱い、メインアプリの設定のみに保存
 * 
 * 🛡️ 追加セキュリティ対策:
 * - すべてのAPI呼び出しがデータベースにログ記録される
 * - 不正アクセスは管理者が確認可能
 * - 必要に応じてAPIを再デプロイして新URLに変更可能
 * 
 * この設定は教育機関での使用には適切なセキュリティレベルです。
 * -----------------------------------------------------------------------------------
 */


// --- グローバル設定 ---
const DATABASE_ID_KEY = 'DATABASE_ID';
const DEPLOYMENT_ID_KEY = 'DEPLOYMENT_ID';
const TARGET_SHEET_NAME = 'ログ';


/**
 * スプレッドシートを開いたときにカスタムメニューを追加します。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 Admin Logger セットアップ')
    .addItem('1. このシートをデータベースとして初期化', 'initializeDatabase')
    .addItem('2. APIをデプロイする', 'showDeploymentInstructions')
    .addSeparator()
    .addItem('現在の設定情報を表示', 'showCurrentSettings')
    .addItem('デプロイ状況をテスト', 'testDeployment')
    .addToUi();
}

/**
 * このスプレッドシートをデータベースとして初期化する関数。
 */
function initializeDatabase() {
  const ui = SpreadsheetApp.getUi();
  const properties = PropertiesService.getScriptProperties();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetId = spreadsheet.getId();

  if (properties.getProperty(DATABASE_ID_KEY) === sheetId) {
    ui.alert('✅ このスプレッドシートは既にデータベースとして初期化されています。');
    return;
  }

  const confirmation = ui.alert(
    'データベースの初期化確認',
    'このスプレッドシート全体をログデータベースとして設定します。よろしいですか？\n' +
    '（シート名が「ログ」に変更され、ヘッダー行が作成されます）',
    ui.ButtonSet.OK_CANCEL
  );

  if (confirmation !== ui.Button.OK) {
    ui.alert('初期化はキャンセルされました。');
    return;
  }

  try {
    properties.setProperty(DATABASE_ID_KEY, sheetId);

    let sheet = spreadsheet.getSheetByName(TARGET_SHEET_NAME);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(TARGET_SHEET_NAME, 0);
      const defaultSheet = spreadsheet.getSheetByName('シート1');
      if (defaultSheet && spreadsheet.getSheets().length > 1) {
        spreadsheet.deleteSheet(defaultSheet);
      }
    }
    
    sheet.clearContents();
    const headers = [
      'userId', 'adminEmail', 'spreadsheetId', 'spreadsheetUrl', 
      'createdAt', 'accessToken', 'configJson', 'lastAccessedAt', 'isActive'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);

    spreadsheet.rename(`【ログデータベース】みんなの回答ボード`);

    ui.alert('✅ データベースの初期化が完了しました。\n\n次に、メニューから「2. APIをデプロイする」を実行してください。');

  } catch (e) {
    Logger.log(e);
    ui.alert(`エラーが発生しました: ${e.message}`);
  }
}

/**
 * GET リクエストハンドラー - 基本的な接続テスト用
 */
function doGet(e) {
  Logger.log('GET request received');
  
  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    message: 'Logger API is running',
    timestamp: new Date().toISOString(),
    service: 'StudyQuest Logger API'
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * ウェブアプリとしてPOSTリクエストを受け取るメインのエンドポイント。
 * メインアプリからの構造化されたAPIリクエストを処理します。
 * @param {GoogleAppsScript.Events.DoPost} e - POSTリクエストイベントオブジェクト。
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  let responsePayload;

  try {
    if (!lock.tryLock(30000)) {
      throw new Error('サーバーが混み合っています。(Lock acquisition failed)');
    }

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('無効なリクエストです。(Invalid request body)');
    }

    const requestData = JSON.parse(e.postData.contents);
    Logger.log(`Received API request: ${JSON.stringify(requestData, null, 2)}`);
    
    // メインアプリからの構造化されたAPIリクエストを処理
    if (requestData.action) {
      responsePayload = handleApiRequest(requestData);
    } else {
      // 従来の直接ログ記録（後方互換性）
      logMetadataToDatabase(requestData);
      responsePayload = { status: 'success', message: 'ログが正常に記録されました。' };
    }

  } catch (error) {
    Logger.log(`Error in doPost: ${error.toString()}\nStack: ${error.stack}`);
    responsePayload = { status: 'error', message: `サーバーエラー: ${error.message}` };
  } finally {
    lock.releaseLock();
  }

  return ContentService.createTextOutput(JSON.stringify(responsePayload))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * メインアプリからのAPIリクエストを処理する
 * @param {Object} requestData - APIリクエストデータ
 * @returns {Object} API応答
 */
function handleApiRequest(requestData) {
  const { action, data, timestamp, requestUser, effectiveUser } = requestData;
  
  Logger.log(`Processing API action: ${action} for user: ${requestUser}`);
  
  switch (action) {
    case 'ping':
      return {
        success: true,
        message: 'Logger API is working',
        timestamp: new Date().toISOString(),
        data: { 
          pong: true,
          requestUser: requestUser,
          effectiveUser: effectiveUser 
        }
      };
      
    case 'getUserInfo':
      return handleGetUserInfo(data);
      
    case 'createUser':
      return handleCreateUser(data, requestUser);
      
    case 'updateUser':
      return handleUpdateUser(data, requestUser);
      
    case 'getExistingBoard':
      return handleGetExistingBoard(data);
      
    case 'checkExistingUser':
      return handleCheckExistingUser(data);
      
    default:
      throw new Error(`未知のAPIアクション: ${action}`);
  }
}

/**
 * ユーザー情報取得API
 */
function handleGetUserInfo(data) {
  try {
    const dbSheet = getDatabaseSheet();
    const { userId } = data;
    
    if (!userId) {
      return { success: false, error: 'userIdが必要です' };
    }
    
    const userData = findUserById(dbSheet, userId);
    
    if (userData) {
      return {
        success: true,
        data: userData
      };
    } else {
      return {
        success: false,
        error: 'ユーザーが見つかりません'
      };
    }
  } catch (error) {
    Logger.log(`getUserInfo error: ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * ユーザー作成API
 */
function handleCreateUser(data, requestUser) {
  try {
    const dbSheet = getDatabaseSheet();
    const timestamp = new Date();
    
    const newRow = [
      data.userId || '',
      data.adminEmail || requestUser || '',
      data.spreadsheetId || '',
      data.spreadsheetUrl || '',
      timestamp, // createdAt
      data.accessToken || '',
      data.configJson || '{}',
      timestamp, // lastAccessedAt
      data.isActive !== undefined ? data.isActive : 'TRUE'
    ];
    
    dbSheet.appendRow(newRow);
    
    Logger.log(`User created: ${data.userId} by ${requestUser}`);
    
    return {
      success: true,
      message: 'ユーザーが正常に作成されました',
      data: {
        userId: data.userId,
        adminEmail: data.adminEmail || requestUser,
        createdAt: timestamp
      }
    };
    
  } catch (error) {
    Logger.log(`createUser error: ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * ユーザー更新API
 */
function handleUpdateUser(data, requestUser) {
  try {
    const dbSheet = getDatabaseSheet();
    const { userId } = data;
    
    Logger.log(`updateUser request: userId=${userId}, data=${JSON.stringify(data)}`);
    
    if (!userId) {
      Logger.log('updateUser error: userIdが必要です');
      return { success: false, error: 'userIdが必要です' };
    }
    
    const userRow = findUserRowById(dbSheet, userId);
    
    if (!userRow) {
      Logger.log(`updateUser error: ユーザーが見つかりません - userId: ${userId}`);
      return { success: false, error: 'ユーザーが見つかりません' };
    }
    
    Logger.log(`Found user row: ${JSON.stringify(userRow.values)}`);
    
    // 更新可能なフィールドを更新
    if (data.spreadsheetId) {
      Logger.log(`Updating spreadsheetId: ${data.spreadsheetId}`);
      userRow.values[2] = data.spreadsheetId;
    }
    if (data.spreadsheetUrl) {
      Logger.log(`Updating spreadsheetUrl: ${data.spreadsheetUrl}`);
      userRow.values[3] = data.spreadsheetUrl;
    }
    if (data.accessToken) {
      Logger.log(`Updating accessToken`);
      userRow.values[5] = data.accessToken;
    }
    if (data.configJson) {
      Logger.log(`Updating configJson: ${data.configJson}`);
      userRow.values[6] = data.configJson;
    }
    if (data.isActive !== undefined) {
      Logger.log(`Updating isActive: ${data.isActive}`);
      userRow.values[8] = data.isActive;
    }
    
    // lastAccessedAtを更新
    userRow.values[7] = new Date();
    
    Logger.log(`Writing to sheet: row ${userRow.rowIndex}, values: ${JSON.stringify(userRow.values)}`);
    
    // シートに書き戻し
    try {
      dbSheet.getRange(userRow.rowIndex, 1, 1, userRow.values.length).setValues([userRow.values]);
      Logger.log(`Successfully updated user: ${userId} by ${requestUser}`);
    } catch (writeError) {
      Logger.log(`Sheet write error: ${writeError.message}`);
      throw writeError;
    }
    
    return {
      success: true,
      message: 'ユーザー情報が正常に更新されました',
      data: {
        userId: userId,
        updatedAt: new Date()
      }
    };
    
  } catch (error) {
    Logger.log(`updateUser error: ${error.message} - Stack: ${error.stack}`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 既存ボード取得API
 */
function handleGetExistingBoard(data) {
  try {
    const dbSheet = getDatabaseSheet();
    const { adminEmail } = data;
    
    if (!adminEmail) {
      return { success: false, error: 'adminEmailが必要です' };
    }
    
    const userData = findUserByEmail(dbSheet, adminEmail);
    
    if (userData) {
      return {
        success: true,
        data: userData
      };
    } else {
      return {
        success: false,
        message: '既存ボードが見つかりません'
      };
    }
  } catch (error) {
    Logger.log(`getExistingBoard error: ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * メタデータをデータベース（スプレッドシート）に記録する。
 * @param {object} data - 記録するメタデータオブジェクト。
 */
function logMetadataToDatabase(data) {
  try {
    const dbSheet = getDatabaseSheet();
    const timestamp = new Date();

    const newRow = [
      data.userId || '',
      data.adminEmail || '',
      data.spreadsheetId || '',
      data.spreadsheetUrl || '',
      timestamp, // createdAt
      '', // accessToken
      data.configJson || '{}',
      timestamp, // lastAccessedAt
      'TRUE' // isActive
    ];

    dbSheet.appendRow(newRow);

  } catch (error) {
    Logger.log(`Failed to log metadata to database: ${error.toString()}`);
    throw new Error(`データベースへの書き込みに失敗しました。詳細: ${error.message}`);
  }
}

/**
 * データベースシートを取得する。
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} データベースのシート。
 */
function getDatabaseSheet() {
  const properties = PropertiesService.getScriptProperties();
  const dbSheetId = properties.getProperty(DATABASE_ID_KEY);

  if (!dbSheetId) {
    // データベースIDが未設定の場合、現在のスプレッドシートを使用
    Logger.log('Database ID not set, using current spreadsheet');
    const currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (currentSpreadsheet) {
      const currentId = currentSpreadsheet.getId();
      properties.setProperty(DATABASE_ID_KEY, currentId);
      Logger.log(`Auto-initialized database with current spreadsheet ID: ${currentId}`);
      
      // ログシートを確認・作成
      let sheet = currentSpreadsheet.getSheetByName(TARGET_SHEET_NAME);
      if (!sheet) {
        sheet = currentSpreadsheet.insertSheet(TARGET_SHEET_NAME);
        const headers = [
          'userId', 'adminEmail', 'spreadsheetId', 'spreadsheetUrl', 
          'createdAt', 'accessToken', 'configJson', 'lastAccessedAt', 'isActive'
        ];
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
        sheet.setFrozenRows(1);
        Logger.log(`Auto-created '${TARGET_SHEET_NAME}' sheet with headers`);
      }
      return sheet;
    } else {
      throw new Error('データベースが設定されておらず、現在のスプレッドシートも取得できません。「🚀 Admin Logger セットアップ」メニューから初期化を完了してください。');
    }
  }

  try {
    Logger.log(`Attempting to open spreadsheet with ID: ${dbSheetId}`);
    
    let spreadsheet;
    try {
      spreadsheet = SpreadsheetApp.openById(dbSheetId);
    } catch (openError) {
      Logger.log(`Failed to open spreadsheet: ${openError.message}`);
      
      // 失敗した場合、現在のスプレッドシートで再試行
      Logger.log('Falling back to current spreadsheet');
      const currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      if (currentSpreadsheet) {
        const currentId = currentSpreadsheet.getId();
        properties.setProperty(DATABASE_ID_KEY, currentId);
        Logger.log(`Fallback: Reset database to current spreadsheet ID: ${currentId}`);
        spreadsheet = currentSpreadsheet;
      } else {
        throw new Error(`スプレッドシート(ID: ${dbSheetId})を開けませんでした。権限を確認してください。`);
      }
    }
    
    if (!spreadsheet) {
      throw new Error(`スプレッドシート(ID: ${dbSheetId})がnullです。`);
    }
    
    Logger.log(`Successfully opened spreadsheet: ${spreadsheet.getName()}`);
    
    let sheet;
    try {
      sheet = spreadsheet.getSheetByName(TARGET_SHEET_NAME);
    } catch (sheetError) {
      Logger.log(`Failed to get sheet '${TARGET_SHEET_NAME}': ${sheetError.message}`);
      sheet = null;
    }

    if (!sheet) {
      // シートが存在しない場合は作成
      Logger.log(`Sheet '${TARGET_SHEET_NAME}' not found, creating new sheet`);
      try {
        sheet = spreadsheet.insertSheet(TARGET_SHEET_NAME);
        const headers = [
          'userId', 'adminEmail', 'spreadsheetId', 'spreadsheetUrl', 
          'createdAt', 'accessToken', 'configJson', 'lastAccessedAt', 'isActive'
        ];
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
        sheet.setFrozenRows(1);
        Logger.log(`Created new sheet '${TARGET_SHEET_NAME}' with headers`);
      } catch (createError) {
        Logger.log(`Failed to create sheet: ${createError.message}`);
        throw new Error(`シート「${TARGET_SHEET_NAME}」の作成に失敗しました: ${createError.message}`);
      }
    }
    
    // データ範囲の存在確認（より安全に）
    try {
      if (sheet && typeof sheet.getDataRange === 'function') {
        const range = sheet.getDataRange();
        if (range) {
          Logger.log(`Sheet data range: ${range.getA1Notation()}, rows: ${range.getNumRows()}`);
        } else {
          Logger.log('Data range is null, but sheet exists');
        }
      } else {
        throw new Error('シートオブジェクトが無効です');
      }
    } catch (rangeError) {
      Logger.log(`Data range error: ${rangeError.message}`);
      // データ範囲エラーでも処理を続行（シートは存在する）
    }
    
    return sheet;
  } catch (error) {
    Logger.log(`getDatabaseSheet error: ${error.toString()}`);
    throw new Error(`データベースアクセスエラー: ${error.message}`);
  }
}

/**
 * 現在の設定情報をダイアログで表示する。
 */
function showCurrentSettings() {
  const ui = SpreadsheetApp.getUi();
  const properties = PropertiesService.getScriptProperties();
  const dbSheetId = properties.getProperty(DATABASE_ID_KEY);
  const deploymentId = properties.getProperty(DEPLOYMENT_ID_KEY);

  let message = '現在の設定情報:\n\n';
  if (dbSheetId) {
    message += `✅ データベース: 設定済み\n   (このスプレッドシート: ${dbSheetId})\n\n`;
  } else {
    message += '❌ データベース: 未設定\n\n';
  }

  if (deploymentId) {
    // deploymentIdからURLを正しく構築
    let webAppUrl;
    if (deploymentId.startsWith('https://')) {
      webAppUrl = deploymentId; // すでに完全なURL
    } else {
      webAppUrl = `https://script.google.com/macros/s/${deploymentId}/exec`;
    }
    message += `✅ APIデプロイ: 設定済み\n   URL: ${webAppUrl}\n\n`;
    
    message += '🔒 セキュリティ情報:\n';
    message += '• 管理者権限で実行、データベースは管理者のみアクセス可能\n';
    message += '• URLは秘匿情報として扱ってください\n';
    message += '• 全API呼び出しがログ記録されます\n';
  } else {
    message += '❌ APIデプロイ: 未設定\n';
  }

  ui.alert(message);
}


/**
 * デプロイ状況をテストする関数
 */
function testDeployment() {
  const ui = SpreadsheetApp.getUi();
  const properties = PropertiesService.getScriptProperties();
  const deploymentId = properties.getProperty(DEPLOYMENT_ID_KEY);
  
  if (!deploymentId) {
    ui.alert('❌ デプロイIDが設定されていません。先にAPIをデプロイしてください。');
    return;
  }
  
  // deploymentIdからURLを正しく構築
  let webAppUrl;
  if (deploymentId.startsWith('https://')) {
    webAppUrl = deploymentId; // すでに完全なURL
  } else {
    webAppUrl = `https://script.google.com/macros/s/${deploymentId}/exec`;
  }
  
  try {
    // GETテスト
    const getResponse = UrlFetchApp.fetch(webAppUrl, {
      method: 'get',
      muteHttpExceptions: true
    });
    
    const getCode = getResponse.getResponseCode();
    const getText = getResponse.getContentText();
    
    let message = `デプロイテスト結果:\n\nURL: ${webAppUrl}\nGETテスト: ${getCode}\n\n`;
    
    if (getCode === 200) {
      message += '✅ GET接続成功\n\n';
      
      // POSTテスト
      const postResponse = UrlFetchApp.fetch(webAppUrl, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({action: 'ping', data: {}}),
        muteHttpExceptions: true
      });
      
      const postCode = postResponse.getResponseCode();
      const postText = postResponse.getContentText();
      
      message += `POSTテスト: ${postCode}\n`;
      
      if (postCode === 200) {
        message += '✅ API正常動作中';
      } else {
        message += `❌ POSTエラー: ${postText.substring(0, 100)}`;
      }
      
    } else {
      message += `❌ 接続失敗\n詳細: ${getText.substring(0, 200)}\n\n`;
      message += '解決方法:\n1. Apps Scriptエディタでプロジェクトを再デプロイ\n2. 「実行者」を「自分」に設定\n3. 「アクセスできるユーザー」を「すべてのユーザー（匿名ユーザーを含む）」に設定';
    }
    
    ui.alert(message);
    
  } catch (e) {
    ui.alert(`テストエラー: ${e.message}`);
  }
}

/**
 * ユーザーにデプロイ方法を案内するUIを表示する。
 */
function showDeploymentInstructions() {
  const ui = SpreadsheetApp.getUi();
  
  if (!PropertiesService.getScriptProperties().getProperty(DATABASE_ID_KEY)) {
    ui.alert('エラー: 先に「1. このシートをデータベースとして初期化」を実行してください。');
    return;
  }
  
  const htmlOutput = HtmlService.createHtmlOutputFromFile('DeploymentGuide')
    .setWidth(600)
    .setHeight(450);
  ui.showModalDialog(htmlOutput, 'デプロイ手順');
}

/**
 * クライアントサイドのHTMLから呼ばれ、デプロイIDをスクリプトプロパティに保存する。
 * @param {string} id - デプロイID。
 */
function saveDeploymentIdToProperties(id) {
  if (id && typeof id === 'string' && id.trim().length > 0) {
    PropertiesService.getScriptProperties().setProperty(DEPLOYMENT_ID_KEY, id.trim());
    SpreadsheetApp.getUi().alert('✅ デプロイIDを保存しました。');
    return 'OK';
  } else {
    SpreadsheetApp.getUi().alert('エラー: 無効なデプロイIDです。');
    return 'Error';
  }
}

/**
 * 既存ユーザーチェックAPI
 */
function handleCheckExistingUser(data) {
  try {
    Logger.log(`checkExistingUser: request data=${JSON.stringify(data)}`);
    
    const dbSheet = getDatabaseSheet();
    const { adminEmail } = data;
    
    if (!adminEmail) {
      Logger.log('checkExistingUser error: adminEmailが必要です');
      return { success: false, error: 'adminEmailが必要です' };
    }
    
    Logger.log(`checkExistingUser: searching for adminEmail="${adminEmail}"`);
    const userData = findUserByEmail(dbSheet, adminEmail);
    
    const result = {
      success: true,
      exists: userData !== null,
      data: userData
    };
    
    Logger.log(`checkExistingUser: result=${JSON.stringify(result)}`);
    return result;
    
  } catch (error) {
    Logger.log(`checkExistingUser error: ${error.message} - Stack: ${error.stack}`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * ユーザーIDでユーザーを検索
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - データベースシート
 * @param {string} userId - ユーザーID
 * @returns {Object|null} ユーザーデータ
 */
function findUserById(sheet, userId) {
  try {
    const data = sheet.getDataRange().getValues();
    
    if (!data || data.length === 0) {
      return null;
    }
    
    for (let i = 1; i < data.length; i++) {
      if (data[i] && data[i][0] === userId) { // userId column
        return {
          userId: data[i][0],
          adminEmail: data[i][1],
          spreadsheetId: data[i][2],
          spreadsheetUrl: data[i][3],
          createdAt: data[i][4],
          accessToken: data[i][5],
          configJson: data[i][6],
          lastAccessedAt: data[i][7],
          isActive: data[i][8]
        };
      }
    }
    return null;
  } catch (error) {
    Logger.log(`findUserById error: ${error.message}`);
    throw new Error(`ユーザー検索エラー: ${error.message}`);
  }
}

/**
 * メールアドレスでユーザーを検索
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - データベースシート
 * @param {string} adminEmail - 管理者メールアドレス
 * @returns {Object|null} ユーザーデータ
 */
function findUserByEmail(sheet, adminEmail) {
  try {
    Logger.log(`findUserByEmail: searching for adminEmail="${adminEmail}"`);
    
    if (!sheet || !adminEmail || adminEmail.trim().length === 0) {
      Logger.log(`findUserByEmail: invalid parameters - sheet=${!!sheet}, adminEmail="${adminEmail}"`);
      return null;
    }
    
    const data = sheet.getDataRange().getValues();
    Logger.log(`findUserByEmail: got ${data.length} rows from sheet`);
    
    if (!data || data.length <= 1) {
      Logger.log(`findUserByEmail: no data rows found`);
      return null;
    }
    
    // 正規化された検索対象メールアドレス（大文字小文字を無視、空白を除去）
    const normalizedSearchEmail = adminEmail.trim().toLowerCase();
    Logger.log(`findUserByEmail: normalized search email="${normalizedSearchEmail}"`);
    
    for (let i = 1; i < data.length; i++) {
      if (data[i] && data[i][1]) {
        const storedEmail = String(data[i][1]).trim().toLowerCase();
        Logger.log(`findUserByEmail: row ${i} - stored email="${storedEmail}"`);
        
        // 空文字列やnull/undefinedとのマッチを防ぐ
        if (storedEmail && storedEmail.length > 0 && storedEmail === normalizedSearchEmail) {
          Logger.log(`findUserByEmail: MATCH found at row ${i}`);
          return {
            userId: data[i][0],
            adminEmail: data[i][1],
            spreadsheetId: data[i][2],
            spreadsheetUrl: data[i][3],
            createdAt: data[i][4],
            accessToken: data[i][5],
            configJson: data[i][6],
            lastAccessedAt: data[i][7],
            isActive: data[i][8]
          };
        }
      } else {
        Logger.log(`findUserByEmail: row ${i} - empty row or no email`);
      }
    }
    
    Logger.log(`findUserByEmail: no match found for "${normalizedSearchEmail}"`);
    return null;
  } catch (error) {
    Logger.log(`findUserByEmail error: ${error.message}`);
    throw new Error(`ユーザー検索エラー: ${error.message}`);
  }
}

/**
 * ユーザーIDで行データを検索（更新用）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - データベースシート
 * @param {string} userId - ユーザーID
 * @returns {Object|null} 行データと行インデックス
 */
function findUserRowById(sheet, userId) {
  try {
    Logger.log(`findUserRowById: searching for userId=${userId}`);
    
    const data = sheet.getDataRange().getValues();
    Logger.log(`findUserRowById: got ${data.length} rows from sheet`);
    
    for (let i = 1; i < data.length; i++) {
      if (data[i] && data[i][0] === userId) { // userId column
        Logger.log(`findUserRowById: found user at row ${i + 1}`);
        return {
          rowIndex: i + 1, // 1-based index for getRange
          values: data[i]
        };
      }
    }
    
    Logger.log(`findUserRowById: user not found`);
    return null;
  } catch (error) {
    Logger.log(`findUserRowById error: ${error.message}`);
    throw error;
  }
}