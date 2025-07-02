/**
 * @fileoverview StudyQuest - Core Functions (最適化版)
 * 主要な業務ロジックとAPI エンドポイント
 */

// =================================================================
// メインロジック
// =================================================================

// doGetLegacy function removed - consolidated into main doGet in UltraOptimizedCore.gs

/**
 * 新規ユーザーを登録する
 * 最適化された実装：キャッシュ利用、エラーハンドリング強化、パフォーマンス改善
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
    spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
    adminUrl: appUrls.adminUrl,
    viewUrl: appUrls.viewUrl,
    message: '新しいボードが作成されました！'
  };
}

/**
 * リアクションを追加/削除する
 * 最適化された実装：パフォーマンス改善、エラーハンドリング強化
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
 * 公開されたシートのデータを取得
 * 最適化された実装：キャッシュ利用、エラーハンドリング強化
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
// HTML依存関数（UI連携）
// =================================================================

/**
 * 管理画面用のステータス情報を取得
 * AdminPanel.htmlから呼び出される
 */
function getStatus() {
  return getAppConfig();
}

/**
 * アクティブなフォーム情報を取得
 * AdminPanel.htmlから呼び出される
 */
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

/**
 * ハイライト状態の切り替え
 * Page.htmlから呼び出される
 */
function toggleHighlight(rowIndex, sheetName) {
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
    
    return processHighlightToggleOptimized(
      userInfo.spreadsheetId,
      sheetName || 'フォームの回答 1',
      rowIndex
    );
  } catch (e) {
    console.error('ハイライト切り替えエラー: ' + e.message);
    return { status: 'error', message: e.message };
  }
}

/**
 * 管理者権限の確認
 * Page.htmlから呼び出される
 */
function checkAdmin() {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      return false;
    }
    
    var userInfo = findUserByIdOptimized(currentUserId);
    if (!userInfo) {
      return false;
    }
    
    var activeUser = Session.getActiveUser().getEmail();
    return activeUser === userInfo.adminEmail;
  } catch (e) {
    console.error('管理者確認エラー: ' + e.message);
    return false;
  }
}

/**
 * アクティブシートのクリア（公開終了）
 * Page.htmlから呼び出される
 */
function clearActiveSheet() {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません');
    }
    
    return updateUserOptimized(currentUserId, {
      configJson: JSON.stringify({
        appPublished: false,
        publishedSheet: ''
      })
    });
  } catch (e) {
    console.error('公開終了エラー: ' + e.message);
    return { status: 'error', message: e.message };
  }
}

/**
 * 利用可能なシート一覧を取得
 * Page.htmlではgetAvailableSheetsとして呼び出される
 */
function getAvailableSheets() {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません');
    }
    
    return getSheetsListOptimized(currentUserId);
  } catch (e) {
    console.error('シート一覧取得エラー: ' + e.message);
    return [];
  }
}

/**
 * ハイライト切り替えの最適化処理
 */
function processHighlightToggleOptimized(spreadsheetId, sheetName, rowIndex) {
  try {
    var service = getOptimizedSheetsService();
    var headerIndices = getHeaderIndicesCached(spreadsheetId, sheetName);
    var highlightColumnIndex = headerIndices[COLUMN_HEADERS.HIGHLIGHT];
    
    if (highlightColumnIndex === undefined) {
      throw new Error('ハイライト列が見つかりません');
    }
    
    // 現在の値を取得
    var range = sheetName + '!' + String.fromCharCode(65 + highlightColumnIndex) + rowIndex;
    var currentValue = service.spreadsheets.values.get(spreadsheetId, range).values;
    var isHighlighted = currentValue && currentValue[0] && currentValue[0][0] === 'true';
    
    // 値を切り替え
    var newValue = isHighlighted ? 'false' : 'true';
    service.spreadsheets.values.update(
      spreadsheetId,
      range,
      { values: [[newValue]] },
      { valueInputOption: 'RAW' }
    );
    
    return {
      status: 'success',
      highlighted: !isHighlighted,
      message: isHighlighted ? 'ハイライトを解除しました' : 'ハイライトしました'
    };
  } catch (e) {
    console.error('ハイライト処理エラー: ' + e.message);
    return { status: 'error', message: e.message };
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
 * フォーム作成（最適化版）
 */
function createStudyQuestFormOptimized(userEmail, userId) {
  try {
    // パフォーマンス測定開始
    var profiler = (typeof globalProfiler !== 'undefined') ? globalProfiler : {
      start: function() {},
      end: function() {}
    };
    profiler.start('createForm');
    
    // 共通ファクトリを使用してフォーム作成
    var formResult = createFormFactory({
      userEmail: userEmail,
      userId: userId,
      questions: 'default',
      formDescription: 'このフォームは「みんなの回答ボード」で表示されます。デジタル・シティズンシップの観点から、オンライン空間での責任ある行動と建設的な対話を育むことを目的としています。回答内容は匿名で表示されます。'
    });
    
    // カスタマイズされた設定を追加
    var form = FormApp.openById(formResult.formId);
    
    // Email収集タイプの設定（可能な場合）
    try {
      if (typeof form.setEmailCollectionType === 'function') {
        form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED);
      }
    } catch (undocumentedError) {
      // 失敗しても続行
      console.warn('Email collection type setting failed:', undocumentedError.message);
    }
    
    // 確認メッセージの設定
    var confirmationMessage = form.getPublishedUrl()
      ? '🎉 回答ありがとうございます！\n\nあなたの大切な意見が届きました。\nみんなの回答ボードで、お友達の色々な考えも見てみましょう。\n新しい発見があるかもしれませんね！\n\n' + form.getPublishedUrl()
      : '🎉 回答ありがとうございます！\n\nあなたの大切な意見が届きました。';
    form.setConfirmationMessage(confirmationMessage);
    
    // サービスアカウントをスプレッドシートに追加（最適化版）
    addServiceAccountToSpreadsheetOptimized(formResult.spreadsheetId);
    
    // リアクション列をスプレッドシートに追加（最適化版）
    addReactionColumnsToSpreadsheetOptimized(formResult.spreadsheetId, formResult.sheetName);
    
    profiler.end('createForm');
    return formResult;
    
  } catch (e) {
    console.error('createStudyQuestFormOptimizedエラー: ' + e.message);
    throw new Error('フォームの作成に失敗しました: ' + e.message);
  }
}

/**
 * サービスアカウントをスプレッドシートに追加（最適化版）
 */
function addServiceAccountToSpreadsheetOptimized(spreadsheetId) {
  try {
    var props = PropertiesService.getScriptProperties();
    var serviceAccountCreds = JSON.parse(props.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS));
    var serviceAccountEmail = serviceAccountCreds.client_email;
    
    if (serviceAccountEmail) {
      var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      spreadsheet.addEditor(serviceAccountEmail);
      console.log('サービスアカウント (' + serviceAccountEmail + ') をスプレッドシートの編集者として追加しました。');
    }
  } catch (e) {
    console.error('サービスアカウントの追加に失敗: ' + e.message);
    // エラーでも処理は継続
  }
}

/**
 * スプレッドシートにリアクション列を追加（最適化版）
 */
function addReactionColumnsToSpreadsheetOptimized(spreadsheetId, sheetName) {
  try {
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.getSheets()[0];
    
    var additionalHeaders = [
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ];
    
    // 効率的にヘッダー情報を取得
    var currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var startCol = currentHeaders.length + 1;
    
    // バッチでヘッダーを追加
    sheet.getRange(1, startCol, 1, additionalHeaders.length).setValues([additionalHeaders]);
    
    // スタイリングを一括適用
    var allHeadersRange = sheet.getRange(1, 1, 1, currentHeaders.length + additionalHeaders.length);
    allHeadersRange.setFontWeight('bold').setBackground('#E3F2FD');
    
    // 自動リサイズ（エラーが出ても続行）
    try {
      sheet.autoResizeColumns(1, allHeadersRange.getNumColumns());
    } catch (resizeError) {
      console.warn('Auto-resize failed:', resizeError.message);
    }
    
    console.log('リアクション列を追加しました: ' + sheetName);
  } catch (e) {
    console.error('リアクション列追加エラー: ' + e.message);
    // エラーでも処理は継続
  }
}

/**
 * シートデータ取得（最適化版）
 */
function getSheetDataOptimized(userId, sheetName, classFilter, sortMode) {
  try {
    var userInfo = findUserByIdOptimized(userId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    var spreadsheetId = userInfo.spreadsheetId;
    var service = getOptimizedSheetsService();
    
    // バッチでデータ、ヘッダー、名簿を取得
    var ranges = [
      sheetName + '!A:Z',
      getConfig().rosterSheetName + '!A:Z'
    ];
    
    var responses = batchGetSheetsData(service, spreadsheetId, ranges);
    var sheetData = responses.valueRanges[0].values || [];
    
    // 「名簿」シートの存在を確認し、存在しない場合は空の配列を返す
    var rosterData = [];
    if (responses.valueRanges[1] && responses.valueRanges[1].values) {
      rosterData = responses.valueRanges[1].values;
    } else {
      debugLog('「名簿」シートが見つからないか、データがありません。');
    }
    
    if (sheetData.length === 0) {
      return {
        status: 'success',
        data: [],
        headers: [],
        totalCount: 0
      };
    }
    
    var headers = sheetData[0];
    var dataRows = sheetData.slice(1);
    
    // ヘッダーインデックスを取得（キャッシュ利用）
    var headerIndices = getHeaderIndicesCached(spreadsheetId, sheetName);
    
    // 名簿マップを作成（キャッシュ利用）
    var rosterMap = buildRosterMap(rosterData);
    
    // 表示モードを取得
    var configJson = JSON.parse(userInfo.configJson || '{}');
    var displayMode = configJson.displayMode || DISPLAY_MODES.ANONYMOUS;
    
    // データを処理
    var processedData = dataRows.map(function(row, index) {
      return processRowDataOptimized(row, headers, headerIndices, rosterMap, displayMode, index + 2);
    });
    
    // フィルタリング
    var filteredData = processedData;
    if (classFilter && classFilter !== 'all') {
      var classIndex = headerIndices[COLUMN_HEADERS.CLASS];
      if (classIndex !== undefined) {
        filteredData = processedData.filter(function(row) {
          return row.originalData[classIndex] === classFilter;
        });
      }
    }
    
    // ソート適用
    var sortedData = applySortModeOptimized(filteredData, sortMode || 'newest');
    
    return {
      status: 'success',
      data: sortedData,
      headers: headers,
      totalCount: sortedData.length,
      displayMode: displayMode
    };
    
  } catch (e) {
    console.error('シートデータ取得エラー: ' + e.message);
    return {
      status: 'error',
      message: 'データの取得に失敗しました: ' + e.message,
      data: [],
      headers: []
    };
  }
}

/**
 * シート一覧取得（最適化版）
 */
function getSheetsListOptimized(userId) {
  try {
    var userInfo = findUserByIdOptimized(userId);
    if (!userInfo) {
      return [];
    }
    
    var service = getOptimizedSheetsService();
    var spreadsheet = getSpreadsheetsData(service, userInfo.spreadsheetId);
    
    return spreadsheet.sheets.map(function(sheet) {
      return {
        name: sheet.properties.title,
        id: sheet.properties.sheetId
      };
    });
  } catch (e) {
    console.error('シート一覧取得エラー: ' + e.message);
    return [];
  }
}

/**
 * 名簿マップを構築
 * @param {array} rosterData - 名簿データ
 * @returns {object} 名簿マップ
 */
function buildRosterMap(rosterData) {
  if (rosterData.length === 0) return {};
  
  var headers = rosterData[0];
  var emailIndex = headers.indexOf(ROSTER_CONFIG.EMAIL_COLUMN);
  var nameIndex = headers.indexOf(ROSTER_CONFIG.NAME_COLUMN);
  var classIndex = headers.indexOf(ROSTER_CONFIG.CLASS_COLUMN);
  
  var rosterMap = {};
  
  for (var i = 1; i < rosterData.length; i++) {
    var row = rosterData[i];
    if (emailIndex !== -1 && row[emailIndex]) {
      rosterMap[row[emailIndex]] = {
        name: nameIndex !== -1 ? row[nameIndex] : '',
        class: classIndex !== -1 ? row[classIndex] : ''
      };
    }
  }
  
  return rosterMap;
}

/**
 * 行データを処理（スコア計算、名前変換など）
 */
function processRowDataOptimized(row, headers, headerIndices, rosterMap, displayMode, rowNumber) {
  var processedRow = {
    rowNumber: rowNumber,
    originalData: row,
    score: 0,
    likeCount: 0,
    understandCount: 0,
    curiousCount: 0,
    isHighlighted: false
  };
  
  // リアクションカウント計算
  REACTION_KEYS.forEach(function(reactionKey) {
    var columnName = COLUMN_HEADERS[reactionKey];
    var columnIndex = headerIndices[columnName];
    
    if (columnIndex !== undefined && row[columnIndex]) {
      var reactions = parseReactionStringOptimized(row[columnIndex]);
      var count = reactions.length;
      
      switch (reactionKey) {
        case 'LIKE':
          processedRow.likeCount = count;
          break;
        case 'UNDERSTAND':
          processedRow.understandCount = count;
          break;
        case 'CURIOUS':
          processedRow.curiousCount = count;
          break;
      }
    }
  });
  
  // ハイライト状態チェック
  var highlightIndex = headerIndices[COLUMN_HEADERS.HIGHLIGHT];
  if (highlightIndex !== undefined && row[highlightIndex]) {
    processedRow.isHighlighted = row[highlightIndex].toString().toLowerCase() === 'true';
  }
  
  // スコア計算
  processedRow.score = calculateRowScoreOptimized(processedRow);
  
  // 名前の表示処理
  var emailIndex = headerIndices[COLUMN_HEADERS.EMAIL];
  if (emailIndex !== undefined && row[emailIndex] && displayMode === DISPLAY_MODES.NAMED) {
    var email = row[emailIndex];
    var rosterInfo = rosterMap[email];
    if (rosterInfo && rosterInfo.name) {
      processedRow.displayName = rosterInfo.name;
    }
  }
  
  return processedRow;
}

/**
 * 行のスコアを計算
 */
function calculateRowScoreOptimized(rowData) {
  var baseScore = 1.0;
  
  // いいね！による加算
  var likeBonus = rowData.likeCount * SCORING_CONFIG.LIKE_MULTIPLIER_FACTOR;
  
  // その他のリアクションも軽微な加算
  var reactionBonus = (rowData.understandCount + rowData.curiousCount) * 0.01;
  
  // ハイライトによる大幅加算
  var highlightBonus = rowData.isHighlighted ? 0.5 : 0;
  
  // ランダム要素（同じスコアの項目をランダムに並べるため）
  var randomFactor = Math.random() * SCORING_CONFIG.RANDOM_SCORE_FACTOR;
  
  return baseScore + likeBonus + reactionBonus + highlightBonus + randomFactor;
}

/**
 * データにソートを適用
 */
function applySortModeOptimized(data, sortMode) {
  switch (sortMode) {
    case 'score':
      return data.sort(function(a, b) { return b.score - a.score; });
    case 'newest':
      return data.reverse();
    case 'oldest':
      return data; // 元の順序（古い順）
    case 'random':
      return shuffleArrayOptimized(data.slice()); // コピーをシャッフル
    case 'likes':
      return data.sort(function(a, b) { return b.likeCount - a.likeCount; });
    default:
      return data;
  }
}

/**
 * 配列をシャッフル（Fisher-Yates shuffle）
 */
function shuffleArrayOptimized(array) {
  for (var i = array.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = array[i];
    array[i] = array[j];
    array[j] = temp;
  }
  return array;
}

/**
 * リアクション文字列をパース
 */
function parseReactionStringOptimized(val) {
  if (!val) return [];
  return val.toString().split(',').map(function(s) { return s.trim(); }).filter(Boolean);
}