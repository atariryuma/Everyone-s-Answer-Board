/**
 * @fileoverview StudyQuest - Core Functions (最適化版)
 * 主要な業務ロジックとAPI エンドポイント
 */

// =================================================================
// メインロジック
// =================================================================

// doGetLegacy function removed - consolidated into main doGet in UltraOptimizedCore.gs

/**
 * 新規ユーザーを登録する（データベース登録のみ）
 * フォーム作成はクイックスタートで実行される
 */
function registerNewUser(adminEmail) {
  var activeUser = Session.getActiveUser();
  if (adminEmail !== activeUser.getEmail()) {
    throw new Error('認証エラー: 操作を実行しているユーザーとメールアドレスが一致しません。');
  }

  // 既存ユーザーチェック（1ユーザー1行の原則）
  var existingUser = findUserByEmail(adminEmail);
  var userId, appUrls;
  
  if (existingUser) {
    // 既存ユーザーの場合は情報を更新
    userId = existingUser.userId;
    var existingConfig = JSON.parse(existingUser.configJson || '{}');
    
    // 設定をリセット（新規登録状態に戻す）
    var updatedConfig = {
      ...existingConfig,
      setupStatus: 'pending',
      lastRegistration: new Date().toISOString(),
      formCreated: false,
      appPublished: false
    };
    
    // 既存ユーザー情報を更新
    updateUser(userId, {
      lastAccessedAt: new Date().toISOString(),
      isActive: 'true',
      configJson: JSON.stringify(updatedConfig)
    });
    
    debugLog('✅ 既存ユーザー情報を更新しました: ' + adminEmail);
    appUrls = generateAppUrls(userId);
    
    return {
      userId: userId,
      adminUrl: appUrls.adminUrl,
      viewUrl: appUrls.viewUrl,
      setupRequired: true,
      message: '既存ユーザーの情報を更新しました。クイックスタートで新しいフォームを作成してください。',
      isExistingUser: true
    };
  }

  // 新規ユーザーの場合
  userId = Utilities.getUuid();
  
  var initialConfig = {
    setupStatus: 'pending',
    createdAt: new Date().toISOString(),
    formCreated: false,
    appPublished: false
  };
  
  var userData = {
    userId: userId,
    adminEmail: adminEmail,
    spreadsheetId: '',
    spreadsheetUrl: '',
    createdAt: new Date().toISOString(),
    configJson: JSON.stringify(initialConfig),
    lastAccessedAt: new Date().toISOString(),
    isActive: 'true'
  };

  try {
    createUser(userData);
    debugLog('✅ データベースに新規ユーザーを登録しました: ' + adminEmail);
  } catch (e) {
    console.error('データベースへのユーザー登録に失敗: ' + e.message);
    throw new Error('ユーザー登録に失敗しました。システム管理者に連絡してください。');
  }

  // 成功レスポンスを返す
  appUrls = generateAppUrls(userId);
  return {
    userId: userId,
    adminUrl: appUrls.adminUrl,
    viewUrl: appUrls.viewUrl,
    setupRequired: true,
    message: 'ユーザー登録が完了しました！次にクイックスタートでフォームを作成してください。',
    isExistingUser: false
  };
}

/**
 * リアクションを追加/削除する
 * Page.htmlから呼び出される - フロントエンド期待形式に対応
 */
function addReaction(rowIndex, reactionKey, sheetName) {
  try {
    var reactingUserEmail = Session.getActiveUser().getEmail();
    var props = PropertiesService.getUserProperties();
    var ownerUserId = props.getProperty('CURRENT_USER_ID');

    if (!ownerUserId) {
      throw new Error('ボードのオーナー情報が見つかりません。');
    }

    // ボードオーナーの情報をDBから取得（キャッシュ利用）
    var boardOwnerInfo = findUserById(ownerUserId);
    if (!boardOwnerInfo) {
      throw new Error('無効なボードです。');
    }

    var result = processReaction(
      boardOwnerInfo.spreadsheetId,
      sheetName,
      rowIndex,
      reactionKey,
      reactingUserEmail
    );
    
    // Page.html期待形式に変換
    if (result && result.status === 'success') {
      // 更新後のリアクション情報を取得
      var updatedReactions = getRowReactions(boardOwnerInfo.spreadsheetId, sheetName, rowIndex, reactingUserEmail);
      
      return {
        status: "ok",
        reactions: updatedReactions
      };
    } else {
      throw new Error(result.message || 'リアクションの処理に失敗しました');
    }
  } catch (e) {
    console.error('addReaction エラー: ' + e.message);
    return {
      status: "error",
      message: e.message
    };
  }
}

// =================================================================
// データ取得関数
// =================================================================

/**
 * 公開されたシートのデータを取得
 * Page.htmlから呼び出される - フロントエンド期待形式に対応
 */
function getPublishedSheetData(sheetName, classFilter, sortOrder) {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません');
    }
    
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    
    // シート名の決定（パラメータまたは設定から）
    var targetSheet = sheetName || configJson.publishedSheet || 'フォームの回答 1';
    
    // データ取得
    var sheetData = getSheetData(currentUserId, targetSheet, classFilter, sortOrder);
    
    if (sheetData.status === 'error') {
      throw new Error(sheetData.message);
    }
    
    // Page.html期待形式に変換
    var formattedData = sheetData.data.map(function(row, index) {
      return {
        rowIndex: row.rowNumber || (index + 2), // 実際の行番号
        name: (sheetData.displayMode === 'named' && row.displayName) ? row.displayName : '',
        class: row.originalData[getHeaderIndex(sheetData.headers, COLUMN_HEADERS.CLASS)] || '',
        opinion: row.originalData[getHeaderIndex(sheetData.headers, COLUMN_HEADERS.OPINION)] || '',
        reason: row.originalData[getHeaderIndex(sheetData.headers, COLUMN_HEADERS.REASON)] || '',
        reactions: {
          UNDERSTAND: { count: row.understandCount || 0, reacted: false },
          LIKE: { count: row.likeCount || 0, reacted: false },
          CURIOUS: { count: row.curiousCount || 0, reacted: false }
        },
        highlight: row.isHighlighted || false
      };
    });
    
    // ヘッダー情報を取得（問題列があれば使用、なければデフォルト）
    var questionHeaderIndex = getHeaderIndex(sheetData.headers, COLUMN_HEADERS.TIMESTAMP); // 問題列は通常ないので、タイムスタンプ列を確認
    var headerTitle = '回答ボード'; // デフォルト
    
    // より適切なヘッダータイトルを設定
    if (sheetData.headers && sheetData.headers.length > 0) {
      // フォームの質問タイトルなどがあれば使用する場合の処理を将来的に追加可能
      headerTitle = '回答ボード';
    }
    
    return {
      header: headerTitle,
      sheetName: targetSheet,
      showCounts: configJson.showCounts !== false,
      displayMode: sheetData.displayMode || 'anonymous',
      data: formattedData,
      rows: formattedData // 後方互換性のため
    };
    
  } catch (e) {
    console.error('公開シートデータ取得エラー: ' + e.message);
    return {
      status: 'error',
      message: 'データの取得に失敗しました: ' + e.message,
      data: [],
      rows: []
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
      var userInfo = findUserByEmail(activeUser);
      if (userInfo) {
        currentUserId = userInfo.userId;
        props.setProperty('CURRENT_USER_ID', currentUserId);
      } else {
        throw new Error('ユーザー情報が見つかりません');
      }
    }
    
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');

    // --- Auto-healing for inconsistent setup states ---
    var needsUpdate = false;
    if (configJson.formUrl && !configJson.formCreated) {
      configJson.formCreated = true;
      needsUpdate = true;
    }
    if (configJson.formCreated && configJson.setupStatus !== 'completed') {
      configJson.setupStatus = 'completed';
      needsUpdate = true;
    }
    if (configJson.publishedSheet && !configJson.appPublished) {
      configJson.appPublished = true;
      needsUpdate = true;
    }
    if (needsUpdate) {
      try {
        updateUser(currentUserId, { configJson: JSON.stringify(configJson) });
      } catch (updateErr) {
        console.warn('Config auto-heal failed: ' + updateErr.message);
      }
    }

    var sheets = getSheetsList(currentUserId);
    var appUrls = generateAppUrls(currentUserId);
    
    // 回答数を取得
    var answerCount = 0;
    var totalReactions = 0;
    try {
      if (userInfo.spreadsheetId && configJson.publishedSheet) {
        var responseData = getResponsesData(currentUserId, configJson.publishedSheet);
        if (responseData.status === 'success') {
          answerCount = responseData.data.length;
          // リアクション数の概算計算（詳細実装は後回し）
          totalReactions = answerCount * 2; // 暫定値
        }
      }
    } catch (countError) {
      console.warn('回答数の取得に失敗: ' + countError.message);
    }
    
    return {
      status: 'success',
      userId: currentUserId,
      adminEmail: userInfo.adminEmail,
      publishedSheet: configJson.publishedSheet || '',
      displayMode: configJson.displayMode || DISPLAY_MODES.ANONYMOUS,
      isPublished: configJson.appPublished || false,
      availableSheets: sheets,
      allSheets: sheets, // AdminPanel.htmlで使用される
      spreadsheetUrl: userInfo.spreadsheetUrl,
      formUrl: configJson.formUrl || '',
      editFormUrl: configJson.editFormUrl || '',
      webAppUrl: appUrls.webAppUrl,
      adminUrl: appUrls.adminUrl,
      viewUrl: appUrls.viewUrl,
      activeSheetName: configJson.publishedSheet || '',
      appUrls: appUrls,
      // データベース詳細情報
      userInfo: {
        userId: currentUserId,
        adminEmail: userInfo.adminEmail,
        spreadsheetId: userInfo.spreadsheetId || '',
        spreadsheetUrl: userInfo.spreadsheetUrl || '',
        createdAt: userInfo.createdAt || '',
        lastAccessedAt: userInfo.lastAccessedAt || '',
        isActive: userInfo.isActive || 'false'
      },
      // 統計情報
      answerCount: answerCount,
      totalReactions: totalReactions,
      // システム状態
      systemStatus: {
        setupStatus: configJson.setupStatus || 'unknown',
        formCreated: configJson.formCreated || false,
        appPublished: configJson.appPublished || false,
        lastUpdated: new Date().toISOString()
      }
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
    initializeDatabaseSheet(dbId);

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


function getResponsesData(userId, sheetName) {
  var userInfo = findUserById(userId);
  if (!userInfo) {
    return { status: 'error', message: 'ユーザー情報が見つかりません' };
  }

  try {
    var service = getSheetsService();
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
 * AdminPanel.htmlから呼び出される（パラメータなし）
 */
function getActiveFormInfo(userId) {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = userId || props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      // コンテキストが設定されていない場合、現在のユーザーで検索
      var activeUser = Session.getActiveUser().getEmail();
      var userInfo = findUserByEmail(activeUser);
      if (userInfo) {
        currentUserId = userInfo.userId;
        props.setProperty('CURRENT_USER_ID', currentUserId);
      } else {
        return { status: 'error', message: 'ユーザー情報が見つかりません' };
      }
    }
    
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      return { status: 'error', message: 'ユーザー情報が見つかりません' };
    }

    var configJson = JSON.parse(userInfo.configJson || '{}');
    
    // フォーム回答数を取得
    var answerCount = 0;
    try {
      if (userInfo.spreadsheetId && configJson.publishedSheet) {
        var responseData = getResponsesData(currentUserId, configJson.publishedSheet);
        if (responseData.status === 'success') {
          answerCount = responseData.data.length;
        }
      }
    } catch (countError) {
      console.warn('回答数の取得に失敗: ' + countError.message);
    }
    
    return {
      status: 'success',
      formUrl: configJson.formUrl || '',
      editFormUrl: configJson.editFormUrl || '',
      editUrl: configJson.editFormUrl || '',  // AdminPanel.htmlが期待するフィールド名
      formId: extractFormIdFromUrl(configJson.formUrl || configJson.editFormUrl || ''),
      spreadsheetUrl: userInfo.spreadsheetUrl || '',
      answerCount: answerCount,
      isFormActive: !!(configJson.formUrl && configJson.formCreated)
    };
  } catch (e) {
    console.error('フォーム情報の取得に失敗: ' + e.message);
    return { status: 'error', message: 'フォーム情報の取得に失敗しました: ' + e.message };
  }
}

/**
 * ハイライト状態の切り替え
 * Page.htmlから呼び出される - フロントエンド期待形式に対応
 */
function toggleHighlight(rowIndex, sheetName) {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません');
    }
    
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    var result = processHighlightToggle(
      userInfo.spreadsheetId,
      sheetName || 'フォームの回答 1',
      rowIndex
    );
    
    // Page.html期待形式に変換
    if (result && result.status === 'success') {
      return {
        status: "ok",
        highlight: result.highlighted || false
      };
    } else {
      throw new Error(result.message || 'ハイライト切り替えに失敗しました');
    }
  } catch (e) {
    console.error('ハイライト切り替えエラー: ' + e.message);
    return { 
      status: "error", 
      message: e.message 
    };
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
    
    var userInfo = findUserById(currentUserId);
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
 * 利用可能なシート一覧を取得
 * Page.htmlから呼び出される - フロントエンド期待形式に対応
 */
function getAvailableSheets() {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません');
    }
    
    var sheets = getSheetsList(currentUserId);
    
    // Page.html期待形式に変換: [{name: string}]
    return sheets.map(function(sheet) {
      return {
        name: typeof sheet === 'string' ? sheet : (sheet.name || sheet.title || sheet)
      };
    });
  } catch (e) {
    console.error('シート一覧取得エラー: ' + e.message);
    return [];
  }
}

/**
 * クイックスタートセットアップ（完全版）
 * フォルダ作成、フォーム作成、スプレッドシート作成、ボード公開まで一括実行
 */
function quickStartSetup(userId) {
  try {
    debugLog('🚀 クイックスタートセットアップ開始: ' + userId);
    
    // ユーザー情報の取得
    var userInfo = findUserById(userId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    var userEmail = userInfo.adminEmail;
    
    // 既にセットアップ済みかチェック
    if (configJson.formCreated && userInfo.spreadsheetId) {
      return {
        status: 'already_completed',
        message: 'クイックスタートは既に完了しています。',
        urls: generateAppUrls(userId)
      };
    }
    
    // ステップ1: ユーザー専用フォルダを作成
    debugLog('📁 ステップ1: フォルダ作成中...');
    var folder = createUserFolder(userEmail);
    
    // ステップ2: Googleフォームとスプレッドシートを作成
    debugLog('📝 ステップ2: フォーム作成中...');
    var formAndSsInfo = createStudyQuestForm(userEmail, userId);
    
    // 作成したファイルをフォルダに移動
    if (folder) {
      try {
        var formFile = DriveApp.getFileById(formAndSsInfo.formId);
        var ssFile = DriveApp.getFileById(formAndSsInfo.spreadsheetId);
        
        folder.addFile(formFile);
        folder.addFile(ssFile);
        
        // 元の場所から削除（Myドライブから移動）
        DriveApp.getRootFolder().removeFile(formFile);
        DriveApp.getRootFolder().removeFile(ssFile);
        
        debugLog('📁 ファイルをフォルダに移動しました: ' + folder.getName());
      } catch (moveError) {
        // フォルダ移動に失敗しても処理は継続
        console.warn('ファイル移動に失敗しましたが、処理を継続します: ' + moveError.message);
      }
    }
    
    // ステップ3: データベースを更新
    debugLog('💾 ステップ3: データベース更新中...');
    var updatedConfig = {
      ...configJson,
      setupStatus: 'completed',
      formCreated: true,
      formUrl: formAndSsInfo.viewFormUrl || formAndSsInfo.formUrl,
      editFormUrl: formAndSsInfo.editFormUrl,
      publishedSheet: formAndSsInfo.sheetName || 'フォームの回答 1',
      appPublished: true,
      folderId: folder ? folder.getId() : '',
      folderUrl: folder ? folder.getUrl() : '',
      completedAt: new Date().toISOString()
    };
    
    updateUser(userId, {
      spreadsheetId: formAndSsInfo.spreadsheetId,
      spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
      configJson: JSON.stringify(updatedConfig)
    });
    
    // ステップ4: 回答ボードを公開状態に設定
    debugLog('🌐 ステップ4: 回答ボード公開中...');
    
    debugLog('✅ クイックスタートセットアップ完了: ' + userId);
    
    var appUrls = generateAppUrls(userId);
    return {
      status: 'success',
      message: 'クイックスタートが完了しました！回答ボードをお楽しみください。',
      urls: appUrls,
      formUrl: updatedConfig.formUrl,
      editFormUrl: updatedConfig.editFormUrl,
      spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
      folderUrl: updatedConfig.folderUrl
    };
    
  } catch (e) {
    console.error('❌ クイックスタートセットアップエラー: ' + e.message);
    
    // エラー時はセットアップ状態をリセット
    try {
      var currentConfig = JSON.parse(userInfo.configJson || '{}');
      currentConfig.setupStatus = 'error';
      currentConfig.lastError = e.message;
      currentConfig.errorAt = new Date().toISOString();
      
      updateUser(userId, {
        configJson: JSON.stringify(currentConfig)
      });
    } catch (updateError) {
      console.error('エラー状態の更新に失敗: ' + updateError.message);
    }
    
    return {
      status: 'error',
      message: 'クイックスタートセットアップに失敗しました: ' + e.message
    };
  }
}

/**
 * ユーザー専用フォルダを作成
 */
function createUserFolder(userEmail) {
  try {
    var rootFolderName = "StudyQuest - ユーザーデータ";
    var userFolderName = "StudyQuest - " + userEmail + " - ファイル";
    
    // ルートフォルダを検索または作成
    var rootFolder;
    var folders = DriveApp.getFoldersByName(rootFolderName);
    if (folders.hasNext()) {
      rootFolder = folders.next();
    } else {
      rootFolder = DriveApp.createFolder(rootFolderName);
    }
    
    // ユーザー専用フォルダを作成
    var userFolders = rootFolder.getFoldersByName(userFolderName);
    if (userFolders.hasNext()) {
      return userFolders.next(); // 既存フォルダを返す
    } else {
      return rootFolder.createFolder(userFolderName);
    }
    
  } catch (e) {
    console.error('フォルダ作成エラー: ' + e.message);
    return null; // フォルダ作成に失敗してもnullを返して処理を継続
  }
}

/**
 * ハイライト切り替え処理
 */
function processHighlightToggle(spreadsheetId, sheetName, rowIndex) {
  try {
    var service = getSheetsService();
    var headerIndices = getHeaderIndices(spreadsheetId, sheetName);
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

function createUserInDb(userData) {
  return createUserOptimized(userData);
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

/**
 * GoogleフォームURLからフォームIDを抽出
 */
function extractFormIdFromUrl(url) {
  if (!url) return '';
  
  try {
    // Regular expression to extract form ID from Google Forms URLs
    var formIdMatch = url.match(/\/forms\/d\/([a-zA-Z0-9-_]+)/);
    if (formIdMatch && formIdMatch[1]) {
      return formIdMatch[1];
    }
    
    // Alternative pattern for e/ URLs  
    var eFormIdMatch = url.match(/\/forms\/d\/e\/([a-zA-Z0-9-_]+)/);
    if (eFormIdMatch && eFormIdMatch[1]) {
      return eFormIdMatch[1];
    }
    
    return '';
  } catch (e) {
    console.warn('フォームID抽出エラー: ' + e.message);
    return '';
  }
}

// =================================================================
// リアクション処理関数
// =================================================================

/**
 * リアクション処理
 */
function processReaction(spreadsheetId, sheetName, rowIndex, reactionKey, reactingUserEmail) {
  try {
    // LockServiceを使って競合を防ぐ
    var lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);
      
      var service = getSheetsService();
      var headerIndices = getHeaderIndices(spreadsheetId, sheetName);
      
      var reactionColumnName = COLUMN_HEADERS[reactionKey];
      var reactionColumnIndex = headerIndices[reactionColumnName];
      
      if (reactionColumnIndex === undefined) {
        throw new Error('リアクション列が見つかりません: ' + reactionColumnName);
      }
      
      // 現在のリアクション文字列を取得
      var cellRange = sheetName + '!' + String.fromCharCode(65 + reactionColumnIndex) + rowIndex;
      var response = service.spreadsheets.values.get(spreadsheetId, cellRange);
      var currentReactionString = (response.values && response.values[0] && response.values[0][0]) || '';
      
      // リアクションの追加/削除処理
      var currentReactions = parseReactionString(currentReactionString);
      var userIndex = currentReactions.indexOf(reactingUserEmail);
      
      if (userIndex >= 0) {
        // 既にリアクション済み → 削除
        currentReactions.splice(userIndex, 1);
      } else {
        // 未リアクション → 追加
        currentReactions.push(reactingUserEmail);
      }
      
      // 更新された値を書き戻す
      var updatedReactionString = currentReactions.join(', ');
      service.spreadsheets.values.update(
        spreadsheetId,
        cellRange,
        { values: [[updatedReactionString]] },
        { valueInputOption: 'RAW' }
      );
      
      debugLog('リアクション更新完了: ' + reactingUserEmail + ' → ' + reactionKey + ' (' + (userIndex >= 0 ? '削除' : '追加') + ')');
      
      return { 
        status: 'success', 
        message: 'リアクションを更新しました。',
        action: userIndex >= 0 ? 'removed' : 'added',
        count: currentReactions.length
      };
      
    } finally {
      lock.releaseLock();
    }
    
  } catch (e) {
    console.error('リアクション処理エラー: ' + e.message);
    return { 
      status: 'error', 
      message: 'リアクションの処理に失敗しました: ' + e.message 
    };
  }
}

// =================================================================
// プレースホルダー関数（実装予定）
// =================================================================

/**
 * フォーム作成
 */
function createStudyQuestForm(userEmail, userId) {
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
    
    // 確認メッセージの設定（回答ボードURLを含む）
    var appUrls = generateAppUrls(userId);
    var boardUrl = appUrls.viewUrl || (appUrls.webAppUrl + '?userId=' + userId);
    
    var confirmationMessage = '🎉 回答ありがとうございます！\n\n' +
      'あなたの大切な意見が届きました。\n' +
      'みんなの回答ボードで、お友達の色々な考えも見てみましょう。\n' +
      '新しい発見があるかもしれませんね！\n\n' +
      '📋 みんなの回答ボード:\n' + boardUrl;
      
    if (form.getPublishedUrl()) {
      confirmationMessage += '\n\n✏️ 回答を編集:\n' + form.getPublishedUrl();
    }
    
    form.setConfirmationMessage(confirmationMessage);
    
    // サービスアカウントをスプレッドシートに追加
    addServiceAccountToSpreadsheet(formResult.spreadsheetId);
    
    // リアクション列をスプレッドシートに追加
    addReactionColumnsToSpreadsheet(formResult.spreadsheetId, formResult.sheetName);
    
    profiler.end('createForm');
    return formResult;
    
  } catch (e) {
    console.error('createStudyQuestFormエラー: ' + e.message);
    throw new Error('フォームの作成に失敗しました: ' + e.message);
  }
}

/**
 * サービスアカウントをスプレッドシートに追加
 */
function addServiceAccountToSpreadsheet(spreadsheetId) {
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
 * スプレッドシートにリアクション列を追加
 */
function addReactionColumnsToSpreadsheet(spreadsheetId, sheetName) {
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
 * シートデータ取得
 */
function getSheetData(userId, sheetName, classFilter, sortMode) {
  try {
    var userInfo = findUserById(userId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    var spreadsheetId = userInfo.spreadsheetId;
    var service = getSheetsService();
    
    // フォーム回答データのみを取得（名簿機能は使用しない）
    var ranges = [sheetName + '!A:Z'];
    
    var responses = batchGetSheetsData(service, spreadsheetId, ranges);
    var sheetData = responses.valueRanges[0].values || [];
    
    // 名簿機能は使用せず、空の配列を設定
    var rosterData = [];
    debugLog('名簿機能は無効化されています。名前はフォーム入力を使用します。');
    
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
    var headerIndices = getHeaderIndices(spreadsheetId, sheetName);
    
    // 名簿マップを作成（キャッシュ利用）
    var rosterMap = buildRosterMap(rosterData);
    
    // 表示モードを取得
    var configJson = JSON.parse(userInfo.configJson || '{}');
    var displayMode = configJson.displayMode || DISPLAY_MODES.ANONYMOUS;
    
    // データを処理
    var processedData = dataRows.map(function(row, index) {
      return processRowData(row, headers, headerIndices, rosterMap, displayMode, index + 2);
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
    var sortedData = applySortMode(filteredData, sortMode || 'newest');
    
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
 * シート一覧取得
 */
function getSheetsList(userId) {
  try {
    var userInfo = findUserById(userId);
    if (!userInfo) {
      return [];
    }
    
    var service = getSheetsService();
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
 * 名簿マップを構築（名簿機能無効化のため空のマップを返す）
 * @param {array} rosterData - 名簿データ（使用されません）
 * @returns {object} 空の名簿マップ
 */
function buildRosterMap(rosterData) {
  // 名簿機能は使用しないため、常に空のマップを返す
  // 名前はフォーム入力から直接取得
  return {};
}

/**
 * 行データを処理（スコア計算、名前変換など）
 */
function processRowData(row, headers, headerIndices, rosterMap, displayMode, rowNumber) {
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
      var reactions = parseReactionString(row[columnIndex]);
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
  processedRow.score = calculateRowScore(processedRow);
  
  // 名前の表示処理（フォーム入力の名前を使用）
  var nameIndex = headerIndices[COLUMN_HEADERS.NAME];
  if (nameIndex !== undefined && row[nameIndex] && displayMode === DISPLAY_MODES.NAMED) {
    processedRow.displayName = row[nameIndex];
  } else if (displayMode === DISPLAY_MODES.NAMED) {
    // 名前入力がない場合のフォールバック
    processedRow.displayName = '匿名';
  }
  
  return processedRow;
}

/**
 * 行のスコアを計算
 */
function calculateRowScore(rowData) {
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
function applySortMode(data, sortMode) {
  switch (sortMode) {
    case 'score':
      return data.sort(function(a, b) { return b.score - a.score; });
    case 'newest':
      return data.reverse();
    case 'oldest':
      return data; // 元の順序（古い順）
    case 'random':
      return shuffleArray(data.slice()); // コピーをシャッフル
    case 'likes':
      return data.sort(function(a, b) { return b.likeCount - a.likeCount; });
    default:
      return data;
  }
}

/**
 * 配列をシャッフル（Fisher-Yates shuffle）
 */
function shuffleArray(array) {
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
function parseReactionString(val) {
  if (!val) return [];
  return val.toString().split(',').map(function(s) { return s.trim(); }).filter(Boolean);
}

/**
 * ヘルパー関数：ヘッダー配列から指定した名前のインデックスを取得
 * COLUMN_HEADERSと統一された方式を使用
 */
function getHeaderIndex(headers, headerName) {
  if (!headers || !headerName) return -1;
  return headers.indexOf(headerName);
}

/**
 * COLUMN_HEADERSキーから適切なヘッダー名を取得
 * @param {string} columnKey - COLUMN_HEADERSのキー（例：'OPINION', 'CLASS'）
 * @returns {string} ヘッダー名
 */
function getColumnHeaderName(columnKey) {
  return COLUMN_HEADERS[columnKey] || '';
}

/**
 * 統一されたヘッダーインデックス取得関数
 * @param {array} headers - ヘッダー配列
 * @param {string} columnKey - COLUMN_HEADERSのキー
 * @returns {number} インデックス（見つからない場合は-1）
 */
function getColumnIndex(headers, columnKey) {
  var headerName = getColumnHeaderName(columnKey);
  return getHeaderIndex(headers, headerName);
}

/**
 * 特定の行のリアクション情報を取得
 */
function getRowReactions(spreadsheetId, sheetName, rowIndex, userEmail) {
  try {
    var service = getSheetsService();
    var headerIndices = getHeaderIndices(spreadsheetId, sheetName);
    
    var reactionData = {
      UNDERSTAND: { count: 0, reacted: false },
      LIKE: { count: 0, reacted: false },
      CURIOUS: { count: 0, reacted: false }
    };
    
    // 各リアクション列からデータを取得
    ['UNDERSTAND', 'LIKE', 'CURIOUS'].forEach(function(reactionKey) {
      var columnName = COLUMN_HEADERS[reactionKey];
      var columnIndex = headerIndices[columnName];
      
      if (columnIndex !== undefined) {
        var range = sheetName + '!' + String.fromCharCode(65 + columnIndex) + rowIndex;
        try {
          var response = service.spreadsheets.values.get(spreadsheetId, range);
          var cellValue = response.values && response.values[0] && response.values[0][0];
          
          if (cellValue) {
            var reactions = parseReactionString(cellValue);
            reactionData[reactionKey].count = reactions.length;
            reactionData[reactionKey].reacted = reactions.indexOf(userEmail) !== -1;
          }
        } catch (cellError) {
          console.warn('リアクション取得エラー(' + reactionKey + '): ' + cellError.message);
        }
      }
    });
    
    return reactionData;
  } catch (e) {
    console.error('getRowReactions エラー: ' + e.message);
    return {
      UNDERSTAND: { count: 0, reacted: false },
      LIKE: { count: 0, reacted: false },
      CURIOUS: { count: 0, reacted: false }
    };
  }
}