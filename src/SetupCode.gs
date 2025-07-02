/**
 * SetupCode.gs
 * アプリケーションの初期設定とデプロイ関連の関数
 */

// =================================================================
// セットアップ関数
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

/**
 * クイックスタートセットアップ
 * Registration.htmlから呼び出される
 */
function quickStartSetup(userId) {
  try {
    debugLog('クイックスタートセットアップ開始: ' + userId);
    
    // ユーザー情報の取得
    var userInfo = findUserById(userId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    var userEmail = userInfo.adminEmail;
    var spreadsheetId = userInfo.spreadsheetId;

    // 1. Googleフォームの作成（既に作成済みの場合はスキップ）
    var formUrl = configJson.formUrl;
    var editFormUrl = configJson.editFormUrl;
    var sheetName = 'フォームの回答 1'; // Default sheet name for form responses

    if (!formUrl) {
      var formAndSsInfo = createStudyQuestForm(userEmail, userId);
      formUrl = formAndSsInfo.formUrl;
      editFormUrl = formAndSsInfo.editFormUrl;
      spreadsheetId = formAndSsInfo.spreadsheetId;
      sheetName = formAndSsInfo.sheetName;

      // Update user info with new form/spreadsheet details
      updateUserInDb(userId, {
        spreadsheetId: spreadsheetId,
        spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
        configJson: JSON.stringify({
          ...configJson,
          formUrl: formUrl,
          editFormUrl: editFormUrl,
          publishedSheet: sheetName, // Set initial published sheet
          appPublished: true // Publish app on quick start
        })
      });
    }

    // 2. Configシートの作成と初期化
    createAndInitializeConfigSheet(spreadsheetId);

    debugLog('クイックスタートセットアップ完了: ' + userId);
    return { status: 'success', message: 'クイックスタートセットアップが完了しました。' };

  } catch (e) {
    console.error('クイックスタートセットアップエラー: ' + e.message);
    return { status: 'error', message: 'クイックスタートセットアップに失敗しました: ' + e.message };
  }
}

/**
 * Configシートを作成し、デフォルト設定で初期化する。
 * @param {string} spreadsheetId - 対象のスプレッドシートID
 */
function createAndInitializeConfigSheet(spreadsheetId) {
  try {
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var configSheet = spreadsheet.getSheetByName(CONFIG_SHEET_NAME);

    if (!configSheet) {
      configSheet = spreadsheet.insertSheet(CONFIG_SHEET_NAME);
      debugLog('Configシートを作成しました。');
    }

    // デフォルト設定を書き込む
    var defaultConfigs = [
      ['questionHeader', '問題'],
      ['answerHeader', '回答'],
      ['reasonHeader', '理由'],
      ['nameHeader', '名前'],
      ['classHeader', 'クラス'],
      ['rosterSheetName', '名簿']
    ];

    var existingData = configSheet.getDataRange().getValues();
    var existingKeys = existingData.map(row => row[0]);

    var dataToWrite = [];
    defaultConfigs.forEach(config => {
      if (!existingKeys.includes(config[0])) {
        dataToWrite.push(config);
      }
    });

    if (dataToWrite.length > 0) {
      var startRow = existingData.length + 1;
      configSheet.getRange(startRow, 1, dataToWrite.length, dataToWrite[0].length).setValues(dataToWrite);
      debugLog('Configシートにデフォルト設定を書き込みました。');
    } else {
      debugLog('Configシートは既に最新です。');
    }

    // ヘッダーのスタイル設定
    configSheet.getRange('A1:B1').setFontWeight('bold').setBackground('#E3F2FD');
    configSheet.autoResizeColumns(1, 2);

  } catch (e) {
    console.error('Configシートの作成と初期化に失敗しました: ' + e.message);
    throw new Error('Configシートの作成と初期化に失敗しました: ' + e.message);
  }
}

// =================================================================
// デプロイ関連関数
// =================================================================

/**
 * WebアプリのURLを保存する
 * @param {string} url - WebアプリのURL
 */
function saveWebAppUrl(url) {
  PropertiesService.getScriptProperties().setProperty('WEB_APP_URL', url);
  console.log('WebアプリURLを保存しました: ' + url);
}

/**
 * 保存されたWebアプリのURLを取得する
 * @returns {string} WebアプリのURL
 */
function getSavedWebAppUrl() {
  return PropertiesService.getScriptProperties().getProperty('WEB_APP_URL');
}

/**
 * デプロイIDを保存する
 * @param {string} deployId - デプロイID
 */
function saveDeployId(deployId) {
  PropertiesService.getScriptProperties().setProperty('DEPLOY_ID', deployId);
  console.log('デプロイIDを保存しました: ' + deployId);
}

/**
 * 保存されたデプロイIDを取得する
 * @returns {string} デプロイID
 */
function getSavedDeployId() {
  return PropertiesService.getScriptProperties().getProperty('DEPLOY_ID');
}

/**
 * スクリプトIDを保存する
 * @param {string} scriptId - スクリプトID
 */
function saveScriptId(scriptId) {
  PropertiesService.getScriptProperties().setProperty('SCRIPT_ID', scriptId);
  console.log('スクリプトIDを保存しました: ' + scriptId);
}

/**
 * 保存されたスクリプトIDを取得する
 * @returns {string} スクリプトID
 */
function getSavedScriptId() {
  return PropertiesService.getScriptProperties().getProperty('SCRIPT_ID');
}

/**
 * スクリプトプロパティをクリアする
 */
function clearScriptProperties() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  console.log('スクリプトプロパティをすべてクリアしました。');
}

/**
 * 現在のデプロイ情報を取得
 */
function getDeploymentInfo() {
  try {
    var scriptId = ScriptApp.getScriptId();
    var webAppUrl = ScriptApp.getService().getUrl();
    var deployId = ScriptApp.getCurrentVersion().getDeploymentId(); // This might not work as expected for all deployments

    return {
      scriptId: scriptId,
      webAppUrl: webAppUrl,
      deployId: deployId || 'N/A' // Fallback if not available
    };
  } catch (e) {
    console.error('デプロイ情報取得エラー: ' + e.message);
    return {
      scriptId: 'Error',
      webAppUrl: 'Error',
      deployId: 'Error',
      error: e.message
    };
  }
}
