/**
 * @fileoverview StudyQuest - Core Functions (最適化版)
 * 主要な業務ロジックとAPI エンドポイント
 */

// =================================================================
// メインロジック
// =================================================================

/**
 * 旧バージョンのdoGet関数（後方互換性のため保持）
 * 注意: UltraOptimizedCore.gsのdoGet関数が優先されます
 */
function doGetLegacy(e) {
  var userId = e.parameter.userId;
  var mode = e.parameter.mode;
  var setup = e.parameter.setup;
  
  // セットアップページの表示
  if (setup === 'true') {
    return HtmlService.createTemplateFromFile('SetupPage').evaluate().setTitle('StudyQuest - サービスアカウント セットアップ');
  }
  
  if (!userId) {
    return HtmlService.createTemplateFromFile('Registration').evaluate().setTitle('新規登録');
  }

  var userInfo = findUserByIdOptimized(userId);
  if (!userInfo) {
    return HtmlService.createHtmlOutput('無効なユーザーIDです。');
  }
  
  // ユーザーの最終アクセス日時を更新（非同期）
  try {
    updateUserOptimized(userId, { lastAccessedAt: new Date().toISOString() });
  } catch (e) {
    console.error('最終アクセス日時の更新に失敗: ' + e.message);
  }

  // ユーザー情報をプロパティに保存（リアクション機能で使用）
  PropertiesService.getUserProperties().setProperty('CURRENT_USER_ID', userId);

  if (mode === 'admin') {
    var template = HtmlService.createTemplateFromFile('AdminPanel');
    template.userInfo = userInfo;
    template.userId = userId;
    return template.evaluate().setTitle('管理パネル - みんなの回答ボード');
  } else {
    var template = HtmlService.createTemplateFromFile('Page');
    template.userInfo = userInfo;
    template.userId = userId;
    return template.evaluate().setTitle('みんなの回答ボード');
  }
}

/**
 * 新規ユーザーを登録する（最適化版）
 */
function registerNewUser(adminEmail) {
  var activeUser = Session.getActiveUser();
  if (adminEmail !== activeUser.getEmail()) {
    throw new Error('認証エラー: 操作を実行しているユーザーとメールアドレスが一致しません。');
  }

  // 既存ユーザーチェック（キャッシュ利用）
  var existingUser = findUserByEmailOptimized(adminEmail);
  if (existingUser) {
    throw new Error('このメールアドレスは既に登録されています。');
  }

  // ステップ1: ユーザー自身の権限でファイル作成
  var userId = Utilities.getUuid();
  var formAndSsInfo = createStudyQuestFormOptimized(adminEmail, userId);

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
    isActive: 'true'
  };

  try {
    createUserOptimized(userData);
    debugLog('✅ データベースに新規ユーザーを登録しました: ' + adminEmail);
  } catch (e) {
    console.error('データベースへのユーザー登録に失敗: ' + e.message);
    throw new Error('ユーザー登録に失敗しました。システム管理者に連絡してください。');
  }

  // 成功レスポンスを返す
  var appUrls = generateAppUrlsOptimized(userId);
  return {
    userId: userId,
    spreadsheetId: formAndSsInfo.spreadsheetId,
    adminUrl: appUrls.adminUrl,
    viewUrl: appUrls.viewUrl,
    message: '新しいボードが作成されました！'
  };
}

/**
 * リアクションを追加/削除する（最適化版）
 */
function addReaction(rowIndex, reactionKey, sheetName) {
  var reactingUserEmail = Session.getActiveUser().getEmail();
  var props = PropertiesService.getUserProperties();
  var ownerUserId = props.getProperty('CURRENT_USER_ID');

  if (!ownerUserId) {
    throw new Error('ボードのオーナー情報が見つかりません。');
  }

  // ボードオーナーの情報をDBから取得（キャッシュ利用）
  var boardOwnerInfo = findUserByIdOptimized(ownerUserId);
  if (!boardOwnerInfo) {
    throw new Error('無効なボードです。');
  }

  return processReactionOptimized(
    boardOwnerInfo.spreadsheetId,
    sheetName,
    rowIndex,
    reactionKey,
    reactingUserEmail
  );
}

// =================================================================
// データ取得関数
// =================================================================

/**
 * 公開されたシートのデータを取得（最適化版）
 */
function getPublishedSheetData(classFilter, sortMode) {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません');
    }
    
    var userInfo = findUserByIdOptimized(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    // 設定から公開シートを取得
    var configJson = JSON.parse(userInfo.configJson || '{}');
    var publishedSheet = configJson.publishedSheet || 'フォームの回答 1';
    
    return getSheetDataOptimized(currentUserId, publishedSheet, classFilter, sortMode);
  } catch (e) {
    console.error('公開シートデータ取得エラー: ' + e.message);
    return {
      status: 'error',
      message: 'データの取得に失敗しました: ' + e.message,
      data: [],
      headers: []
    };
  }
}

/**
 * アプリ設定を取得（最適化版）
 */
function getAppConfig() {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      // コンテキストが設定されていない場合、現在のユーザーで検索
      var activeUser = Session.getActiveUser().getEmail();
      var userInfo = findUserByEmailOptimized(activeUser);
      if (userInfo) {
        currentUserId = userInfo.userId;
        props.setProperty('CURRENT_USER_ID', currentUserId);
      } else {
        throw new Error('ユーザー情報が見つかりません');
      }
    }
    
    var userInfo = findUserByIdOptimized(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    var sheets = getSheetsListOptimized(currentUserId);
    var appUrls = generateAppUrlsOptimized(currentUserId);
    
    return {
      status: 'success',
      userId: currentUserId,
      adminEmail: userInfo.adminEmail,
      publishedSheet: configJson.publishedSheet || '',
      displayMode: configJson.displayMode || DISPLAY_MODES.ANONYMOUS,
      isPublished: configJson.appPublished || false,
      availableSheets: sheets,
      spreadsheetUrl: userInfo.spreadsheetUrl,
      formUrl: configJson.formUrl || '',
      editFormUrl: configJson.editFormUrl || '',
      webAppUrl: appUrls.webAppUrl,
      adminUrl: appUrls.adminUrl,
      viewUrl: appUrls.viewUrl,
      activeSheetName: configJson.publishedSheet || '',
      appUrls: appUrls
    };
  } catch (e) {
    console.error('アプリ設定取得エラー: ' + e.message);
    return {
      status: 'error',
      message: '設定の取得に失敗しました: ' + e.message
    };
  }
}

// =================================================================
// セットアップ関数
// =================================================================

/**
 * アプリケーションの初期セットアップ（管理者が手動で実行）
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
    initializeDatabaseSheetOptimized(dbId);

    console.log('✅ セットアップが正常に完了しました。');
  } catch (e) {
    console.error('セットアップエラー:', e);
    throw new Error('セットアップに失敗しました: ' + e.message);
  }
}

// =================================================================
// ヘルパー関数
// =================================================================

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getActiveFormInfo(userId) {
  var userInfo = findUserByIdOptimized(userId);
  if (!userInfo) {
    return { status: 'error', message: 'ユーザー情報が見つかりません' };
  }

  try {
    var configJson = JSON.parse(userInfo.configJson || '{}');
    return {
      status: 'success',
      formUrl: configJson.formUrl || '',
      editFormUrl: configJson.editFormUrl || '',
      spreadsheetUrl: userInfo.spreadsheetUrl || ''
    };
  } catch (e) {
    console.error('設定情報の取得に失敗: ' + e.message);
    return { status: 'error', message: '設定情報の取得に失敗しました' };
  }
}

function getResponsesData(userId, sheetName) {
  var userInfo = findUserByIdOptimized(userId);
  if (!userInfo) {
    return { status: 'error', message: 'ユーザー情報が見つかりません' };
  }

  try {
    var service = getOptimizedSheetsService();
    var spreadsheetId = userInfo.spreadsheetId;
    var range = (sheetName || 'フォームの回答 1') + '!A:Z';
    
    var response = batchGetSheetsData(service, spreadsheetId, [range]);
    var values = response.valueRanges[0].values || [];
    
    if (values.length === 0) {
      return { status: 'success', data: [], headers: [] };
    }

    return {
      status: 'success',
      data: values.slice(1),
      headers: values[0]
    };
  } catch (e) {
    console.error('回答データの取得に失敗: ' + e.message);
    return { status: 'error', message: '回答データの取得に失敗しました: ' + e.message };
  }
}

// =================================================================
// 互換性関数（後方互換性のため）
// =================================================================

function getWebAppUrl() {
  return getWebAppUrlCached();
}

function findUserById(userId) {
  return findUserByIdOptimized(userId);
}

function findUserByEmail(email) {
  return findUserByEmailOptimized(email);
}

function updateUserInDb(userId, updateData) {
  return updateUserOptimized(userId, updateData);
}

function getSheetsService() {
  return getOptimizedSheetsService();
}

function clearAllCaches() {
  clearAllCache();
  clearServiceAccountTokenCache();
}

function debugLog() {
  if (DEBUG && typeof console !== 'undefined' && console.log) {
    console.log.apply(console, arguments);
  }
}

// =================================================================
// プレースホルダー関数（実装予定）
// =================================================================

/**
 * フォーム作成（プレースホルダー）
 */
function createStudyQuestFormOptimized(userEmail, userId) {
  // 元のcreateStudyQuestForm関数の最適化版を実装
  throw new Error('createStudyQuestFormOptimized is not implemented yet');
}

/**
 * シートデータ取得（プレースホルダー）
 */
function getSheetDataOptimized(userId, sheetName, classFilter, sortMode) {
  // DataProcessorの機能を関数ベースで実装
  throw new Error('getSheetDataOptimized is not implemented yet');
}

/**
 * シート一覧取得（プレースホルダー）
 */
function getSheetsListOptimized(userId) {
  // シート一覧取得の最適化版を実装
  throw new Error('getSheetsListOptimized is not implemented yet');
}