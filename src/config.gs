/**
 * config.gs - 軽量化版
 * 新サービスアカウントアーキテクチャで必要最小限の関数のみ
 */

const CONFIG_SHEET_NAME = 'Config';

/**
 * 現在のユーザーのスプレッドシートを取得
 * AdminPanel.htmlから呼び出される
 */
function getCurrentSpreadsheet() {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーIDが設定されていません');
    }
    
    var userInfo = findUserById(currentUserId);
    if (!userInfo || !userInfo.spreadsheetId) {
      throw new Error('ユーザー情報またはスプレッドシートIDが見つかりません');
    }
    
    return SpreadsheetApp.openById(userInfo.spreadsheetId);
  } catch (e) {
    console.error('getCurrentSpreadsheet エラー: ' + e.message);
    throw new Error('スプレッドシートの取得に失敗しました: ' + e.message);
  }
}

/**
 * アクティブなスプレッドシートのURLを取得
 * AdminPanel.htmlから呼び出される
 * @returns {string} スプレッドシートURL
 */
function openActiveSpreadsheet() {
  try {
    var ss = getCurrentSpreadsheet();
    return ss.getUrl();
  } catch (e) {
    console.error('openActiveSpreadsheet エラー: ' + e.message);
    throw new Error('スプレッドシートのURL取得に失敗しました: ' + e.message);
  }
}

/**
 * 簡易設定取得関数（AdminPanel.htmlとの互換性のため）
 * 新アーキテクチャでは基本的にデフォルト設定を使用
 */
function getConfig() {
  try {
    var spreadsheet = getCurrentSpreadsheet();
    var configSheet = spreadsheet.getSheetByName(CONFIG_SHEET_NAME);
    
    var config = {
      questionHeader: '問題',
      answerHeader: '回答',
      reasonHeader: '理由',
      nameHeader: '名前',
      classHeader: 'クラス',
      rosterSheetName: '名簿'
    };

    if (configSheet) {
      var data = configSheet.getDataRange().getValues();
      for (var i = 0; i < data.length; i++) {
        var key = data[i][0];
        var value = data[i][1];
        if (key && value !== undefined) {
          config[key] = value;
        }
      }
    }
    return config;
  } catch (error) {
    console.error('getConfig error:', error.message);
    return {
      questionHeader: '問題',
      answerHeader: '回答',
      reasonHeader: '理由',
      nameHeader: '名前',
      classHeader: 'クラス',
      rosterSheetName: '名簿'
    };
  }
}

/**
 * 名簿シート名を取得
 * @returns {string} 名簿シート名
 */
function getRosterSheetName() {
  try {
    var cfg = getConfig();
    return cfg.rosterSheetName || '名簿';
  } catch (e) {
    console.warn('getRosterSheetName error: ' + e.message);
    return '名簿';
  }
}

/**
 * 簡易設定保存関数（AdminPanel.htmlとの互換性のため）
 * 新アーキテクチャでは実際の保存は行わない
 */
function saveSheetConfig(sheetName, cfg) {
  try {
    console.log('設定保存（新アーキテクチャでは無操作）:', sheetName, cfg);
    return `シート「${sheetName}」の設定を確認しました`;
  } catch (error) {
    console.error('Error in saveSheetConfig:', error);
    return `設定の確認中にエラーが発生しました: ${error.message}`;
  }
}

/**
 * シートヘッダーを取得
 * AdminPanel.htmlから呼び出される
 */
function getSheetHeaders(sheetName) {
  try {
    var spreadsheet = getCurrentSpreadsheet();
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error('シートが見つかりません: ' + sheetName);
    }
    var lastColumn = sheet.getLastColumn();
    if (lastColumn === 0) {
      return [];
    }
    var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    return headers.filter(String); // 空のヘッダーを除外
  } catch (e) {
    console.error('getSheetHeaders エラー: ' + e.message);
    throw new Error('シートヘッダーの取得に失敗しました: ' + e.message);
  }
}

/**
 * スプレッドシートURLを追加
 * Unpublished.htmlから呼び出される
 */
function addSpreadsheetUrl(url) {
  try {
    var spreadsheetId = url.match(/\/d\/([a-zA-Z0-9_-]+)/)[1];
    if (!spreadsheetId) {
      throw new Error('無効なスプレッドシートURLです。');
    }
    
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません。');
    }
    
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません。');
    }
    
    // サービスアカウントをスプレッドシートに追加
    addServiceAccountToSpreadsheet(spreadsheetId);
    
    // ユーザー情報にスプレッドシートIDとURLを更新
    updateUserInDb(currentUserId, {
      spreadsheetId: spreadsheetId,
      spreadsheetUrl: url
    });
    
    return { status: 'success', message: 'スプレッドシートが正常に追加されました。' };
  } catch (e) {
    console.error('addSpreadsheetUrl エラー: ' + e.message);
    throw new Error('スプレッドシートの追加に失敗しました: ' + e.message);
  }
}

/**
 * シートを切り替える
 * AdminPanel.htmlから呼び出される
 */
function switchToSheet(sheetName) {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません。');
    }
    
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません。');
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    configJson.publishedSheet = sheetName;
    configJson.appPublished = true; // 公開状態にする
    
    updateUserInDb(currentUserId, {
      configJson: JSON.stringify(configJson)
    });
    
    return 'シートが正常に切り替わり、公開されました。';
  } catch (e) {
    console.error('switchToSheet エラー: ' + e.message);
    throw new Error('シートの切り替えに失敗しました: ' + e.message);
  }
}

/**
 * アクティブシートをクリア（公開停止）
 * AdminPanel.htmlから呼び出される
 */
function clearActiveSheet() {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません。');
    }
    
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません。');
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    configJson.publishedSheet = ''; // アクティブシートをクリア
    configJson.appPublished = false; // 公開停止
    
    updateUserInDb(currentUserId, {
      configJson: JSON.stringify(configJson)
    });
    
    return '回答ボードの公開を停止しました。';
  } catch (e) {
    console.error('clearActiveSheet エラー: ' + e.message);
    throw new Error('回答ボードの公開停止に失敗しました: ' + e.message);
  }
}

/**
 * 表示オプションを設定
 * AdminPanel.htmlから呼び出される
 */
function setDisplayOptions(options) {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません。');
    }
    
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません。');
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    configJson.showNames = options.showNames;
    configJson.showCounts = options.showCounts;
    
    updateUserInDb(currentUserId, {
      configJson: JSON.stringify(configJson)
    });
    
    return '表示オプションを保存しました。';
  } catch (e) {
    console.error('setDisplayOptions エラー: ' + e.message);
    throw new Error('表示オプションの保存に失敗しました: ' + e.message);
  }
}

/**
 * 管理者権限チェック
 * Page.htmlから呼び出される
 */
function checkAdmin() {
  debugLog('checkAdmin function called.');
  try {
    var activeUserEmail = Session.getActiveUser().getEmail();
    var props = PropertiesService.getScriptProperties();
    var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    var service = getSheetsService();
    var sheetName = DB_SHEET_CONFIG.SHEET_NAME;
    
    var response = service.spreadsheets.values.get(dbId, sheetName + '!A:H');
    var data = (response && response.values) ? response.values : [];
    
    if (data.length === 0) {
      debugLog('checkAdmin: No data found in Users sheet or sheet is empty.');
      return false;
    }
    
    var headers = data[0];
    var emailIndex = headers.indexOf('adminEmail');
    var isActiveIndex = headers.indexOf('isActive');
    
    if (emailIndex === -1 || isActiveIndex === -1) {
      console.warn('checkAdmin: Missing required headers in Users sheet.');
      return false;
    }

    for (var i = 1; i < data.length; i++) {
      if (data[i][emailIndex] === activeUserEmail && data[i][isActiveIndex] === 'true') {
        return true;
      }
    }
    return false;
  } catch (e) {
    console.error('checkAdmin エラー: ' + e.message);
    return false;
  }
}

/**
 * 管理者からボードを作成
 * AdminPanel.htmlから呼び出される
 */
function createBoardFromAdmin() {
  try {
    var activeUserEmail = Session.getActiveUser().getEmail();
    var userId = Utilities.getUuid(); // 新しいユーザーIDを生成
    
    // フォームとスプレッドシートを作成
    var formAndSsInfo = createStudyQuestForm(activeUserEmail, userId);
    
    // 中央データベースにユーザー情報を登録
    var initialConfig = {
      formUrl: formAndSsInfo.viewFormUrl || formAndSsInfo.formUrl,
      editFormUrl: formAndSsInfo.editFormUrl,
      createdAt: new Date().toISOString(),
      publishedSheet: formAndSsInfo.sheetName, // 作成時に公開シートを設定
      appPublished: true // 作成時に公開状態にする
    };
    
    var userData = {
      userId: userId,
      adminEmail: activeUserEmail,
      spreadsheetId: formAndSsInfo.spreadsheetId,
      spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
      createdAt: new Date().toISOString(),
      configJson: JSON.stringify(initialConfig),
      lastAccessedAt: new Date().toISOString(),
      isActive: 'true'
    };
    
    createUserInDb(userData);
    
    // 成功レスポンスを返す
    var appUrls = generateAppUrls(userId);
    return {
      status: 'success',
      message: '新しいボードが作成され、公開されました！',
      formUrl: formAndSsInfo.formUrl,
      spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
      adminUrl: appUrls.adminUrl,
      viewUrl: appUrls.viewUrl,
      formTitle: formAndSsInfo.formTitle // フォームタイトルも返す
    };
  } catch (e) {
    console.error('createBoardFromAdmin エラー: ' + e.message);
    throw new Error('ボードの作成に失敗しました: ' + e.message);
  }
}

/**
 * 既存ボード情報を取得
 * Registration.htmlから呼び出される
 */
function getExistingBoard() {
  try {
    var activeUserEmail = Session.getActiveUser().getEmail();
    var userInfo = findUserByEmail(activeUserEmail);
    
    if (userInfo && userInfo.isActive === 'true') {
      var appUrls = generateAppUrls(userInfo.userId);
      return {
        status: 'existing_user',
        userId: userInfo.userId,
        adminUrl: appUrls.adminUrl,
        viewUrl: appUrls.viewUrl
      };
    } else if (userInfo && userInfo.isActive === 'false') {
      return {
        status: 'setup_required',
        userId: userInfo.userId
      };
    } else {
      return {
        status: 'new_user'
      };
    }
  } catch (e) {
    console.error('getExistingBoard エラー: ' + e.message);
    return { status: 'error', message: '既存ボード情報の取得に失敗しました: ' + e.message };
  }
}

/**
 * ユーザー認証を検証
 * Registration.htmlから呼び出される
 */
function verifyUserAuthentication() {
  try {
    var email = Session.getActiveUser().getEmail();
    if (email) {
      return { authenticated: true, email: email };
    } else {
      return { authenticated: false, email: null };
    }
  } catch (e) {
    console.error('verifyUserAuthentication エラー: ' + e.message);
    return { authenticated: false, email: null, error: e.message };
  }
}
